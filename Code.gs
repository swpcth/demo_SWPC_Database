/**
 * SWC_Database — Backend API (Google Apps Script Web App)
 * ---------------------------------------------------------
 * วิธีติดตั้ง:
 * 1. สร้าง Google Sheet ชื่อ SWC_Database ตามโครงสร้างใน docs/SHEET_SCHEMA.md
 * 2. เปิด Extensions > Apps Script ในชีตนั้น แล้ววางไฟล์นี้แทนไฟล์ Code.gs เริ่มต้น
 * 3. ตั้งค่า Script Properties: SECRET_KEY (ใช้เซ็น token), ADMIN_SETUP_KEY (ใช้สร้างแอดมินคนแรกครั้งเดียว)
 *    (Project Settings > Script Properties)
 * 4. Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 *    - คัดลอก Web App URL ไปใส่ใน web/js/api.js (API_BASE_URL)
 * 5. เรียก action=setupFirstAdmin ครั้งเดียวเพื่อสร้างบัญชีแอดมินคนแรก (ดูฟังก์ชันด้านล่าง)
 */

// เวอร์ชันของไฟล์ backend นี้ — เปลี่ยนทุกครั้งที่ส่งมอบ Code.gs ชุดใหม่ ใช้ตรวจสอบว่า Deploy เวอร์ชันล่าสุดแล้วหรือยัง
// (หน้าเว็บบางหน้าจะเรียก action=getBackendVersion มาเทียบกับเวอร์ชันที่ฝั่งเว็บคาดหวังไว้ ถ้าไม่ตรงจะเตือนให้ Deploy ใหม่)
const BACKEND_VERSION = '2026-11-15-pdpa-v21';

/**
 * ใส่ ID ของ Google Sheet ตรงนี้ถ้าโปรเจกต์ Apps Script นี้เป็นแบบ "แยกเดี่ยว" (standalone) —
 * คือสร้างจาก script.google.com โดยตรง ไม่ได้เปิดจากเมนู Extensions > Apps Script ในตัวชีตเอง
 * วิธีหา ID: เปิด Google Sheet แล้วดูจาก URL
 *   https://docs.google.com/spreadsheets/d/ ID_ตรงนี้ /edit
 *
 * ถ้าเปิดสคริปต์นี้จากเมนู Extensions > Apps Script ในตัวชีตอยู่แล้ว (แบบ container-bound)
 * ปล่อยเป็นค่าว่าง '' ไว้ได้เลย ระบบจะใช้ getActiveSpreadsheet() ให้อัตโนมัติ
 */
const SPREADSHEET_ID = '1YzS5awhIwTsi2ygWv9i3MXS8s9zl1sVmH_i7_2x1Z-4';

function getSS_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error(
      'ไม่พบสเปรดชีตที่ผูกกับสคริปต์นี้ — ถ้าเปิดสคริปต์จาก script.google.com โดยตรง ' +
      'ให้ใส่ ID ของ Google Sheet ในตัวแปร SPREADSHEET_ID ที่ต้นไฟล์ Code.gs ' +
      '(หรือเปิดสคริปต์จากเมนู Extensions > Apps Script ภายในตัวชีตแทน)'
    );
  }
  return ss;
}

// ---------- Sheet definitions (ต้นทางเดียวกับ docs/SHEET_SCHEMA.md) ----------

// คอลัมน์ที่เป็น "ตัวเลขแต่จริงๆ คือรหัส/เบอร์" ต้องบังคับเป็นข้อความล้วน (Plain text) เสมอ กัน Google Sheets ตัดเลข 0 นำหน้าทิ้งอัตโนมัติ
// (เบอร์โทร, เลขบัตรประชาชน, รหัสไปรษณีย์ เป็นตัวเลขล้วนแท้ๆ เสี่ยงถูกตัด 0 นำหน้า — เลขสมาชิก/เลขใบอนุญาตมีตัวอักษรนำหน้าอยู่แล้วจึงไม่เสี่ยง แต่ใส่กันไว้ด้วยเผื่อกรณีอนาคต)
const DIGIT_STRING_COLUMNS = {
  Members: [
    'member_no', 'national_id', 'phone', 'workplace_zipcode', 'workplace_phone',
    'addr_current_zipcode', 'addr_registered_zipcode', 'mailing_recipient_phone', 'license_no'
  ],
  MembersOrg: ['juristic_no', 'tax_id', 'rep_national_id', 'rep_phone', 'rep_phone2', 'addr_zipcode', 'office_phone', 'reg_no'],
  Licenses: ['license_no'],
  LicenseSanctions: ['license_no', 'decision_date', 'effective_date', 'suspend_until', 'ethics_case_no', 'ethics_meeting_no'], // วันที่เก็บเป็นข้อความ YYYY-MM-DD กัน Sheets แปลงเป็นวันที่
  MembersHonorary: ['honorary_no', 'phone'], // เลขลำดับ 5 หลักมี 0 นำหน้า ต้องเป็นข้อความ
  Beneficiaries: ['phone'],
  InterviewSessions: ['time_text'], // เช่น "09.30 - 14.30 น." กัน Sheets แปลงเป็นเวลา
  InterviewGroups: ['time_range'],
  InterviewCandidates: ['candidate_no', 'time_slot'], // "0001" / "09:30" ต้องเป็นข้อความ
  ExamRounds: ['round_label'] // "2/2569" ถ้าไม่บังคับเป็นข้อความ Google Sheets จะแปลงเป็นวันที่ (ก.พ. 2569) เอง
};

const SHEET_DEFINITIONS = {
  Users: ['user_id', 'username', 'password_hash', 'role', 'member_id', 'thaid_pid', 'status', 'created_at', 'last_login', 'failed_login_attempts', 'locked_until', 'display_name', 'account_type'],

  Members: [
    'member_id', 'member_no', 'national_id', 'prefix_th', 'prefix_en', 'first_name_th', 'last_name_th',
    'first_name_en', 'last_name_en', 'dob', 'gender', 'religion', 'ethnicity', 'nationality', 'phone',
    'email', 'line_id', 'photo_url', 'education_json', 'work_status', 'position', 'agency',
    'affiliation', 'work_province', 'expertise', 'workplace_name',
    'workplace_house_no', 'workplace_building', 'workplace_floor', 'workplace_moo', 'workplace_soi',
    'workplace_road', 'workplace_subdistrict', 'workplace_district', 'workplace_province', 'workplace_zipcode',
    'workplace_email', 'workplace_phone',
    'addr_current_house_no', 'addr_current_building', 'addr_current_moo', 'addr_current_soi',
    'addr_current_road', 'addr_current_subdistrict', 'addr_current_district', 'addr_current_province', 'addr_current_zipcode',
    'addr_registered_house_no', 'addr_registered_building', 'addr_registered_moo', 'addr_registered_soi',
    'addr_registered_road', 'addr_registered_subdistrict', 'addr_registered_district', 'addr_registered_province', 'addr_registered_zipcode',
    'mailing_address_type', 'mailing_recipient_name', 'mailing_recipient_phone',
    'member_type', 'member_status',
    'applied_date', 'membership_expire_date', 'license_no',
    'license_status', 'ce_score_cat1', 'ce_score_cat2', 'ce_verified_via_api', 'ce_synced_at', 'use_professional_prefix', 'photo_public_consent', 'documents_json', 'created_at', 'updated_at'
  ],

  MembersOrg: [
    'org_id', 'org_name_th', 'org_name_en', 'org_type', 'juristic_no', 'tax_id', 'founded_date',
    'purposes_json', 'rep_prefix', 'rep_name', 'rep_dob', 'rep_ethnicity', 'rep_nationality', 'rep_religion',
    'rep_marital_status', 'rep_position', 'rep_national_id', 'rep_phone', 'rep_phone2', 'rep_email', 'rep_line_id',
    'addr_house_no', 'addr_building', 'addr_floor', 'addr_moo', 'addr_soi', 'addr_road',
    'addr_subdistrict', 'addr_district', 'addr_province', 'addr_zipcode',
    'office_phone', 'website', 'logo_url', 'reg_no', 'applied_date', 'approved_date', 'status',
    'end_date', 'note', 'documents_json', 'created_at', 'updated_at'
  ],

  Applications: [
    'application_id', 'applicant_type', 'applicant_user_id', 'ref_data_json', 'documents_json', 'status',
    'reject_reason', 'fee_amount', 'submitted_at', 'reviewed_by', 'reviewed_at', 'approved_member_id'
  ],

  // รอบสอบข้อเขียน/สอบสัมภาษณ์ — แอดมินเปิด/ปิดรอบได้เอง (แต่ละปีมีมากกว่า 1 รอบ) ใบสมัครสอบจริงเก็บในชีต Applications
  // (applicant_type: 'สมัครสอบข้อเขียน' / 'สมัครสอบสัมภาษณ์') โดยอ้างอิง round_id ไว้ใน ref_data_json
  ExamRounds: [
    'round_id', 'exam_type', 'round_label', 'open_date', 'close_date', 'exam_date',
    'format_options_json', 'location_note', 'status', 'note', 'created_by', 'created_at', 'updated_at'
  ],

  // ระบบจัดการสอบสัมภาษณ์ (ดูหมายเหตุที่หัวข้อ 6.5): ตั้งค่ารอบ / กลุ่มสอบ+กรรมการ+รหัสเข้าใช้ / ผู้เข้าสอบที่จัดเข้ากลุ่ม+คะแนน
  InterviewSessions: [
    'session_id', 'round_id', 'time_text', 'pass_score', 'scoring_closed', 'results_applied_at', 'note', 'created_by', 'created_at', 'updated_at'
  ],

  InterviewGroups: [
    'group_id', 'round_id', 'group_no', 'time_range', 'inv1_name', 'inv2_name', 'password_hash', 'created_by', 'created_at', 'updated_at'
  ],

  InterviewCandidates: [
    'ic_id', 'round_id', 'group_id', 'application_id', 'candidate_no', 'candidate_name', 'order_no', 'time_slot', 'absent',
    'scores_json', 'remarks_json', 'created_by', 'created_at', 'updated_at', 'updated_by'
  ],

  Licenses: [
    'license_id', 'member_id', 'license_no', 'issue_date', 'expire_date', 'status',
    'renewal_docs_json', 'renewed_at', 'revoked_reason'
  ],

  // คำสั่งพักใช้/เพิกถอนใบอนุญาต (มติ/หนังสือคณะกรรมการจรรยาบรรณ) — เจ้าหน้าที่บันทึก + แนบหลักฐาน, เจ้าหน้าที่ทุกบทบาทอ่านได้
  LicenseSanctions: [
    'sanction_id', 'member_id', 'license_id', 'license_no', 'member_name', 'action_type', 'status',
    'ethics_case_no', 'ethics_meeting_no', 'decision_date', 'effective_date', 'suspend_until', 'reason', 'remarks',
    'evidence_json', 'prev_license_status',
    'lifted_at', 'lifted_by', 'lift_type', 'lift_reason', 'lift_evidence_json',
    'created_by', 'created_at', 'updated_at', 'updated_by'
  ],

  MembershipRenewals: [
    'renewal_id', 'member_id', 'renewed_at', 'old_expire_date', 'new_expire_date',
    'years', 'payment_id', 'renewed_by', 'note'
  ],

  Notifications: [
    'notification_id', 'ref_type', 'ref_id', 'member_id', 'channel', 'type',
    'email_to', 'subject', 'sent_at', 'status'
  ],

  Cards: ['card_id', 'member_id', 'card_type', 'qr_token', 'issued_at', 'valid_until'],

  QrScanLog: ['scan_id', 'qr_token', 'member_id', 'scanned_at'],

  LicenseDownloadLog: ['log_id', 'license_no', 'member_id', 'downloaded_by', 'downloaded_at'],

  Welfare: [
    'welfare_id', 'member_id', 'welfare_type', 'beneficiary_id', 'description', 'incident_date',
    'documents_json', 'status', 'reject_reason',
    'requested_at', 'reviewed_by', 'reviewed_at', 'approved_amount', 'note',
    'disbursed_at', 'disbursed_by', 'disbursement_doc_url'
  ],

  Beneficiaries: ['beneficiary_id', 'member_id', 'prefix', 'name', 'phone', 'relationship', 'created_at'],

  WelfareDocLog: ['log_id', 'welfare_id', 'action', 'doc_type', 'url', 'actor_user_id', 'actor_role', 'at'],

  Payments: [
    'payment_id', 'ref_type', 'ref_id', 'member_id', 'amount', 'slip_url', 'method', 'channel', 'status',
    'verified_by', 'verified_at', 'paid_at', 'gateway_ref', 'gateway_raw_json',
    'receipt_no', 'receipt_url', 'receipt_issued_at', 'reconciled', 'reconciled_at'
  ],

  Refunds: [
    'refund_id', 'payment_id', 'member_id', 'reason', 'requested_by', 'requested_at',
    'status', 'approved_by', 'approved_at', 'refund_amount', 'refund_note', 'refunded_at'
  ],

  ReconciliationLog: [
    'recon_id', 'recon_date', 'source', 'matched_count', 'unmatched_count',
    'performed_by', 'performed_at', 'notes'
  ],

  WebhookLog: ['log_id', 'received_at', 'raw_payload', 'processed', 'payment_id', 'result'],

  AuditLog: [
    'log_id', 'ref_type', 'ref_id', 'action', 'old_value_json', 'new_value_json',
    'actor_user_id', 'created_at'
  ],

  // PDPA — ความยินยอมเก็บ/ใช้ข้อมูลส่วนบุคคล (append-only: แถวล่าสุดของแต่ละบัญชีคือสถานะปัจจุบัน) และคำขอใช้สิทธิของเจ้าของข้อมูล
  PdpaConsents: ['consent_id', 'user_id', 'username', 'member_id', 'policy_version', 'status', 'channel', 'reason', 'created_at'],
  PdpaRequests: ['request_id', 'user_id', 'username', 'member_id', 'request_type', 'detail', 'status', 'response_note', 'handled_by', 'handled_at', 'created_at'],

  // ผู้ดำรงตำแหน่งนายกสภาฯ/เลขาธิการฯ/เหรัญญิกฯ แต่ละวาระ — ใช้แสดงในบัตรสมาชิก/บัตรใบอนุญาต/ใบอนุญาต/ใบรับรองสมาชิกวิสามัญ/ใบเสร็จ
  // เดิมเป็นข้อมูลตายตัวในโค้ด (PRESIDENT_TERMS/SECRETARY_GENERAL_TERMS/TREASURER_TERMS) ย้ายมาเก็บในชีตนี้แทน ให้แอดมินจัดการเองได้ผ่านหน้าเว็บ
  CouncilOfficers: [
    'officer_id', 'position_type', 'name_th', 'name_en', 'position_label', 'signature_url',
    'term_start', 'term_end', 'note', 'created_at', 'updated_at'
  ],

  // ฐานข้อมูลสมาชิกกิตติมศักดิ์ (ผู้ทรงคุณวุฒิ/ผู้มีผลงานเชิงประจักษ์ด้านสังคมสงเคราะห์ศาสตร์) — ไม่ผูกกับใบอนุญาต/ค่าบำรุง
  MembersHonorary: [
    'honorary_id', 'honorary_no', 'prefix_th', 'first_name_th', 'last_name_th', 'first_name_en', 'last_name_en',
    'qualification_type', 'expertise', 'position', 'organization', 'phone', 'email', 'photo_url',
    'appointed_date', 'term_end', 'status', 'note', 'created_by', 'created_at', 'updated_at'
  ],

  VoidedLicenses: [
    'void_id', 'old_license_no', 'new_license_no', 'member_id', 'member_name',
    'void_date', 'reason', 'actor_user_id'
  ],

  LicenseRenewalHistory: [
    'renewal_id', 'license_id', 'member_id', 'method', 'old_expire_date', 'new_expire_date',
    'renewed_by', 'renewed_at', 'note'
  ],

  RenewalEligibility: [
    'eligibility_id', 'member_id', 'import_type', 'source', 'batch_name',
    'imported_at', 'imported_by', 'status', 'decided_by', 'decided_at'
  ],

  Sessions: ['token', 'user_id', 'role', 'created_at', 'expires_at']
};

// ค่าธรรมเนียมสมาชิกภาพตามประเภทและระยะเวลา (TOR ข้อ 2.2.1) — ✅ อัปเดตให้ตรงกับข้อมูลจริงที่ได้รับมาแล้ว
// สมาชิกเลือกระยะเวลาได้เอง: 1 ปี, 5 ปี, และ "ตลอดชีพ" (เฉพาะสมาชิกสามัญที่มีอายุ 65 ปีขึ้นไป ณ วันที่สมัคร/ต่ออายุเท่านั้น)
const MEMBERSHIP_FEE_TABLE = {
  'สามัญ': { '1': 200, '5': 1000, 'lifetime': 2000 },
  'สมทบ': { '1': 100, '5': 500 },
  'วิสามัญ': { '1': 200, '5': 1000 }
};
const LIFETIME_MIN_AGE = 65; // อายุขั้นต่ำสำหรับสมัคร/ต่ออายุแบบตลอดชีพ (เฉพาะสมาชิกสามัญ)

// ค่าสมาชิกครั้งแรกเข้า (แยกจากค่าบำรุงสมาชิก) — เก็บเฉพาะตอนสมัครครั้งแรกเท่านั้น ไม่เก็บตอนต่ออายุหรือเปลี่ยนประเภท
// ✅ ยืนยันจากแบบฟอร์มสมัครสมาชิกวิสามัญจริง (สวชส.10) แล้วว่าเก็บกับสมาชิกทุกประเภทรวมวิสามัญด้วย (ไม่ได้ยกเว้นวิสามัญตามที่เคยสันนิษฐานไว้ผิด)
const MEMBERSHIP_ENTRANCE_FEE = 100;

// ค่าจัดส่งหนังสือรับรองการเป็นสมาชิก — เก็บเฉพาะตอนสมัครสมาชิกวิสามัญครั้งแรกเท่านั้น (มีระบุชัดในแบบฟอร์ม สวชส.10)
// ยังไม่มีหลักฐานว่าสมาชิกบุคคล (สามัญ/สมทบ) ต้องเสียค่านี้ด้วยหรือไม่ จึงเก็บเฉพาะวิสามัญตามหลักฐานที่มีก่อน
const CERTIFICATE_MAILING_FEE = 20;

// ค่าธรรมเนียมเปลี่ยนประเภทสมาชิก (สมทบ→สามัญ) — ✅ ยืนยันแล้วว่า "ไม่มีค่าธรรมเนียมแยก คิดเป็นค่าบำรุงสมาชิกสามัญตามระยะเวลาที่เลือกเลย (ไม่เก็บค่าแรกเข้าซ้ำ)"
// จึงไม่มีค่าคงที่ตายตัวอีกต่อไป — คำนวณจาก MEMBERSHIP_FEE_TABLE['สามัญ'][ระยะเวลาที่เลือก] โดยตรงตอนยื่นคำขอเปลี่ยนประเภท

// อายุใบอนุญาต/สมาชิกภาพเริ่มนับจากวันที่อนุมัติ ตามระยะเวลาที่เลือก (1/5/ตลอดชีพ) — "ตลอดชีพ" ในระบบนี้บันทึกเป็นวันหมดอายุ 100 ปีข้างหน้า (แทนค่า "ไม่มีวันหมดอายุ" เพื่อให้ใช้ร่วมกับ logic คำนวณวันอื่นๆ ในระบบได้)
const LIFETIME_YEARS_EQUIVALENT = 100;

// จำนวนวันก่อนหมดอายุที่เริ่มถือว่า "ใกล้หมดอายุ" และเริ่มส่งอีเมลแจ้งเตือน (TOR ข้อ 3.2, 3.3)
const EXPIRY_WARNING_DAYS = 30;

// ---------- ระบบใบอนุญาต (TOR ข้อ 4) — ✅ อัปเดตให้ตรงกับข้อมูลจริงที่ได้รับมาแล้ว ----------
const LICENSE_REGISTRATION_FEE = 2000;            // ค่าขึ้นทะเบียนรับใบอนุญาตเป็นผู้ประกอบวิชาชีพฯ (TOR ข้อ 4.1)
const LICENSE_RENEWAL_FEE = 1000;                 // ค่าต่ออายุใบอนุญาตฯ ก่อนหมดอายุด้วยผล CE (TOR ข้อ 4.2)
const LICENSE_RENEWAL_FEE_AFTER_EXPIRY = 2000;    // ค่าต่ออายุหลังใบอนุญาตหมดอายุด้วยผล CE — แพงกว่าเพราะถือเป็นการขึ้นทะเบียนใหม่
const EXAM_SUBJECT_FEE = 500;                     // ค่าธรรมเนียมสอบข้อเขียน ต่อวิชา (มี 2 วิชาบังคับ รวม 1,000 บาท) — ใช้กับการต่ออายุใบอนุญาตด้วยวิธีสอบข้อเขียนใหม่
const EXAM_SUBJECTS = ['หลักคิด ทฤษฎีทางสังคมสงเคราะห์', 'การปฏิบัติงานสังคมสงเคราะห์']; // รายชื่อวิชาสอบข้อเขียนจริงทั้ง 2 วิชา — ใช้ทั้งตอนขึ้นทะเบียนครั้งแรกและต่ออายุด้วยวิธีสอบ
const EXAM_INTERVIEW_LABEL = 'ผลการสอบสัมภาษณ์เพื่อประเมินทักษะและทัศนคติทางสังคมสงเคราะห์'; // ชื่อเอกสารผลสอบสัมภาษณ์ ใช้ร่วมกันทั้งขึ้นทะเบียนครั้งแรกและต่ออายุ

// ---------- ระบบสมัครสอบข้อเขียน/สอบสัมภาษณ์ (TOR ข้อ 4) — ✅ อ้างอิงจากแบบฟอร์ม สวชส.5/สวชส.8 จริง ----------
const EXAM_TYPES = ['เขียน', 'สัมภาษณ์'];
const EXAM_APPLICANT_TYPES = ['สมัครสอบข้อเขียน', 'สมัครสอบสัมภาษณ์']; // applicant_type ใหม่ 2 ค่าในชีต Applications (ใช้ชีตเดิมร่วมกับคำขอขึ้นทะเบียน/ต่ออายุ)
const EXAM_ROUND_STATUSES = ['เปิดรับสมัคร', 'ปิดรับสมัคร', 'เสร็จสิ้น'];
const EXAM_FORMAT_OPTIONS = ['online', 'onsite']; // สอบ Online (Zoom) หรือ Onsite ตามที่สภาฯ ประกาศ — ตามแบบฟอร์มจริง
// เอกสารแสดงทักษะ/ทัศนคติทางสังคมสงเคราะห์ สำหรับสมัครสอบสัมภาษณ์ (สวชส.8 — เลือกอย่างใดอย่างหนึ่ง)
const EXAM_INTERVIEW_ESSAY_CATEGORIES = [
  'เอกสารสรุปบทเรียนการปฏิบัติงานสังคมสงเคราะห์',
  'เอกสารสรุปบทเรียนจากการฝึกภาคปฏิบัติทางสังคมสงเคราะห์',
  'เอกสารสรุปบทเรียนจากการมีส่วนร่วมในกิจกรรมทางสังคม',
  'เอกสารเขียนสรุปบทเรียนจากการปฏิบัติงานในองค์กรสังคมสงเคราะห์',
  'เอกสารสรุปบทเรียนจากการทำงานวิจัยทางสังคมสงเคราะห์'
];
// สอบสัมภาษณ์ไม่มีค่าธรรมเนียม (ไม่พบส่วนแจ้งค่าธรรมเนียมในแบบฟอร์ม สวชส.8) — ค่าสมัครสอบข้อเขียนใช้ EXAM_SUBJECT_FEE ต่อวิชาเดิมที่มีอยู่แล้ว
// ประวัติการสอบข้อเขียนที่ผ่านมา (สวชส.5 — เลือกข้อใดข้อหนึ่ง) ใช้เฉพาะตอนสมัครสอบข้อเขียน — เรียงลำดับ/ถ้อยคำให้ตรงกับแบบฟอร์มจริง
const EXAM_HISTORY_OPTIONS = ['passed_subject1', 'passed_subject2', 'failed_both', 'never'];
const EXAM_HISTORY_LABELS = {
  passed_subject1: 'เคยสอบผ่านวิชาหลักคิด ปรัชญา จริยธรรมทางสังคมสงเคราะห์แล้ว',
  passed_subject2: 'เคยสอบผ่านวิชาการปฏิบัติงานสังคมสงเคราะห์แล้ว',
  failed_both: 'เคยสอบข้อเขียนแล้วแต่ยังไม่ผ่านทั้ง 2 วิชา',
  never: 'ไม่เคยสมัครสอบข้อเขียนมาก่อน'
};
// ตัวเลือกที่ต้องระบุ "วันที่สอบผ่าน" เพิ่มเติม (ตามแบบฟอร์ม สวชส.5 ที่มีช่องวันที่/เดือน/ปี ต่อท้ายตัวเลือกนี้)
const EXAM_HISTORY_DATE_REQUIRED = ['passed_subject1', 'passed_subject2'];
// เอกสารแสดงทักษะและทัศนคติทางสังคมสงเคราะห์ ต้องส่งทั้งไฟล์ PDF และไฟล์ Word (ตามประกาศของสภาฯ ที่ให้ส่งอีเมลทั้ง 2 รูปแบบ)
function examEssayDocTypes_(category) {
  return [category + ' (PDF)', category + ' (Word)'];
}
const LICENSE_DURATION_YEARS = 5;      // อายุใบอนุญาตต่อรอบ (ปี) — ปรับให้ตรงกับรอบสะสมคะแนน CE ที่ระเบียบสภาฯ ระบุไว้ชัดเจนว่า 5 ปี
const RENEWAL_WINDOW_DAYS = 180;       // กดต่ออายุล่วงหน้าได้ไม่เกินกี่วันก่อนหมดอายุ (TOR ข้อ 4.2.1, 4.2.2)

// คำนำหน้าวิชาชีพสำหรับนักสังคมสงเคราะห์รับอนุญาต — คล้ายกับ "นพ./พญ." ของแพทยสภา
// สมาชิกที่มีใบอนุญาตสถานะปกติ เลือกใช้แทนคำนำหน้าส่วนตัว (นาย/นาง/นางสาว) ได้ในการแสดงผลชื่อสาธารณะ (ไม่ระบุเพศ ใช้รูปแบบเดียวกันทุกคน)
const PROFESSIONAL_TITLE_TH = 'สค.ร.';
const PROFESSIONAL_TITLE_EN = 'LSW'; // Licensed Social Worker — ตัวย่อภาษาอังกฤษตามที่สภาฯ กำหนด (LSW) ยังไม่มีคำแปลทางการจากสภาฯ ควรตรวจสอบก่อนใช้งานจริง

/** จัดรูปแบบชื่อสำหรับแสดงผล — ใช้คำนำหน้าวิชาชีพ "สค.ร." แทนคำนำหน้าส่วนตัว ถ้าสมาชิกเลือกใช้และมีใบอนุญาตสถานะปกติเท่านั้น
 * ป้องกันการเลือกใช้คำนำหน้าวิชาชีพทั้งที่ใบอนุญาตหมดอายุ/ถูกพักใช้/เพิกถอนแล้ว (ตรวจสอบสถานะจริงทุกครั้งที่แสดงผล ไม่ใช่แค่ตอนติ๊กเลือก)
 */
function formatMemberDisplayName_(member, licenseStatus) {
  const nameTh = [member.first_name_th, member.last_name_th].filter(Boolean).join(' ');
  const nameEn = [member.first_name_en, member.last_name_en].filter(Boolean).join(' ');
  const useProfessional = isTrue_(member.use_professional_prefix) && licenseStatus === 'ปกติ';
  return {
    th: useProfessional ? (PROFESSIONAL_TITLE_TH + ' ' + nameTh) : ([member.prefix_th, nameTh].filter(Boolean).join(' ')),
    en: useProfessional ? (PROFESSIONAL_TITLE_EN + ' ' + nameEn) : ([member.prefix_en, nameEn].filter(Boolean).join(' ')),
    usedProfessionalTitle: useProfessional
  };
}

/** คำนวณค่าธรรมเนียมสมาชิกภาพตามประเภท+ระยะเวลาที่เลือก พร้อมตรวจสอบสิทธิ์ตลอดชีพ (อายุ 65+ เฉพาะสามัญ) */
function getMembershipFee_(memberType, duration, dob) {
  const table = MEMBERSHIP_FEE_TABLE[memberType];
  if (!table) throw new Error('ประเภทสมาชิกไม่ถูกต้อง');
  if (!table[duration]) throw new Error('ระยะเวลาสมาชิกภาพนี้ไม่รองรับสำหรับประเภท ' + memberType);
  if (duration === 'lifetime') {
    if (memberType !== 'สามัญ') throw new Error('สมาชิกภาพตลอดชีพรองรับเฉพาะสมาชิกสามัญเท่านั้น');
    const age = ageFromDob_(dob);
    if (age === null || age < LIFETIME_MIN_AGE) throw new Error('สมาชิกภาพตลอดชีพสงวนสิทธิ์สำหรับผู้ที่มีอายุ ' + LIFETIME_MIN_AGE + ' ปีขึ้นไป ณ วันที่สมัคร/ต่ออายุเท่านั้น');
  }
  return table[duration];
}

/** คำนวณวันหมดอายุใหม่จากวันเริ่มต้น+ระยะเวลาที่เลือก ("ตลอดชีพ" ใช้ค่าเทียบเท่า 100 ปี แทนการไม่มีวันหมดอายุ) */
function computeMembershipExpiry_(baseDateIso, duration) {
  const years = duration === 'lifetime' ? LIFETIME_YEARS_EQUIVALENT : Number(duration);
  return addYearsIso_(baseDateIso, years);
}

// เกณฑ์คะแนนการศึกษาต่อเนื่อง (CE) ขั้นต่ำ — อ้างอิงระเบียบจริงของสภาฯ (รอบสะสม 5 ปี แบ่ง 2 หมวด ต้องผ่านทั้ง 2 หมวด):
// หมวด 1 = ความรู้/ปฏิบัติการด้านสังคมสงเคราะห์ศาสตร์โดยตรง (ทักษะ จรรยาบรรณ กฎหมาย หรืองานเฉพาะทาง เช่น เด็ก ผู้สูงอายุ คนพิการ ฯลฯ)
// หมวด 2 = ความรู้ด้านสังคมศาสตร์/มนุษยศาสตร์/จิตวิทยา/พฤติกรรมศาสตร์และสาขาที่เกี่ยวข้อง เช่น เทคโนโลยีสารสนเทศ
// ยังผูกกับคะแนนที่กรอกด้วยมือ เพราะยังไม่มีการเชื่อมต่อ API ระบบ CE จริง (ดู LICENSE_SYSTEM.md)
const CE_MINIMUM_CATEGORY1 = 40; // หมวด 1
const CE_MINIMUM_CATEGORY2 = 10; // หมวด 2
const CE_MINIMUM_SCORE = CE_MINIMUM_CATEGORY1 + CE_MINIMUM_CATEGORY2; // รวม 50 หน่วยคะแนน (แสดงผลรวมเท่านั้น เกณฑ์จริงต้องผ่านแยกทั้ง 2 หมวด)

// เว็บไซต์การศึกษาต่อเนื่อง (CE) จริงของสภาฯ — https://cswpeducation.swpc.or.th/
// เป็นเว็บสำหรับสมาชิก/หน่วยงาน/ผู้ดูแลระบบล็อกอินใช้งาน ไม่มี API สาธารณะเปิดให้เชื่อมต่อ
// ต้องขอสิทธิ์เชื่อมต่อ (API key/endpoint) จากผู้ดูแลระบบเว็บไซต์นี้โดยตรงก่อนจึงจะเชื่อมได้จริง (ดู LICENSE_SYSTEM.md)
const CE_WEBSITE_URL = 'https://cswpeducation.swpc.or.th/';

// ---------- ระบบข้อมูลสวัสดิการสมาชิก (TOR ข้อ 7) ----------
// อ้างอิงระเบียบคณะกรรมการสภาวิชาชีพสังคมสงเคราะห์ ว่าด้วยการจัดสวัสดิการหรือบริการอื่นๆ
// ให้แก่สมาชิกประเภทสามัญของสภาวิชาชีพสังคมสงเคราะห์ พ.ศ. 2567 (ราชกิจจานุเบกษา 28 มิ.ย. 2567)
const WELFARE_TYPES = [
  'การช่วยเหลือกรณีถึงแก่กรรม',
  'การช่วยเหลือกรณีประสบภัยพิบัติหรือสาธารณภัย',
  'สวัสดิการหรือบริการอื่นๆ ตามที่คณะอนุกรรมการสวัสดิการอนุมัติ'
];

// วงเงินสูงสุดตามระเบียบ ข้อ 22 (ค่าตั้งต้นตามระเบียบจริง — "อื่นๆ" ไม่มีเพดานตายตัว แล้วแต่คณะกรรมการประกาศกำหนด)
const WELFARE_MAX_AMOUNTS = {
  'การช่วยเหลือกรณีถึงแก่กรรม': 5000,               // (ก) หรีด/เครื่องเคารพศพ ไม่เกิน 2,000 + (ข) ช่วยงานศพ 3,000 (และเงินเหลือจาก ก ถ้ามี)
  'การช่วยเหลือกรณีประสบภัยพิบัติหรือสาธารณภัย': 10000 // ต้องไม่เกิดจากการกระทำของสมาชิกโดยเจตนาหรือประมาทเลินเล่ออย่างร้ายแรง (ข้อ 22(2))
};

const WELFARE_CLAIM_DEADLINE_DAYS = 180; // สิทธิขอรับสวัสดิการสิ้นสุดเมื่อพ้น 180 วันนับแต่วันเกิดเหตุ/ถึงแก่กรรม (ข้อ 7)

// ---------- ระบบ Dashboard และรายงาน (TOR ข้อ 8) ----------
// SLA มาตรฐานพิจารณาอนุมัติคำขอ — ค่าตั้งต้นตามกล่อง "ข้อเสนอแนะ" ใน TOR ข้อ 2.2 (ไม่เกิน 15 วันทำการ) ปรับได้ตามระเบียบจริง
const SLA_APPROVAL_DAYS = 15;

// ---------- ระบบชำระเงิน (Payment Gateway) — TOR ข้อ 11 ----------
// ช่องทางการชำระเงินที่ TOR ต้องการให้รองรับ — ใช้กำกับ dropdown ช่องทางทั้งระบบ
const PAYMENT_CHANNELS = ['บัตรเครดิต/เดบิต', 'พร้อมเพย์ (QR PromptPay)', 'Mobile Banking', 'เคาน์เตอร์เซอร์วิส/Bill Payment', 'โอนเงินผ่านธนาคาร (ตรวจสอบสลิปด้วยมือ)'];

// ผู้ให้บริการ Payment Gateway — **ยังไม่ได้ระบุ/เชื่อมต่อจริง** เพราะ TOR ขอให้ระบุผู้ให้บริการก่อนเริ่มพัฒนา (เช่น K-Payment, 2C2P, GB PromptPay)
// แต่การสนทนานี้ไม่มีการระบุผู้ให้บริการหรือให้ข้อมูล credential ใดๆ มา — ดู PAYMENT_SYSTEM.md เรื่องขั้นตอนเชื่อมต่อจริงเมื่อเลือกผู้ให้บริการแล้ว
const PAYMENT_GATEWAY_PROVIDER = null; // เช่น 'K_PAYMENT' | '2C2P' | 'GB_PROMPTPAY' — ใส่ค่าเมื่อเลือกผู้ให้บริการจริงแล้ว


// ---------- ระบบแสดงบัตรสมาชิก (TOR ข้อ 5) ----------
const MEMBERSHIP_RENEWAL_WINDOW_DAYS = 90; // ต่ออายุสมาชิก (บัตร) ล่วงหน้าได้ไม่เกินกี่วัน (TOR ข้อ 5.1)

// ผู้ดำรงตำแหน่งนายกสภาฯ/เลขาธิการฯ/เหรัญญิกฯ — ย้ายจากข้อมูลตายตัวในโค้ดมาเก็บในชีต CouncilOfficers แทน (แก้ไขได้ผ่านหน้า "จัดการนายกสภาฯ/เลขาธิการฯ")
// ข้อมูลชุดแรก (จากหลักฐานจริงที่เคยตรวจสอบไว้) จะถูกสร้างให้อัตโนมัติครั้งเดียวตอนรัน setupAllSheets ถ้าชีตยังว่างอยู่ — ดู seedCouncilOfficersIfEmpty_()

/** หาผู้ดำรงตำแหน่ง ณ วันที่ที่กำหนด จากชีต CouncilOfficers — ถ้าไม่พบวาระที่ตรงกับวันที่ ใช้วาระล่าสุดแทน (กันเอกสารออกไม่ได้เพราะข้อมูลตกหล่น)
 * คืนค่า name = ชื่อภาษาไทยเท่านั้น (ใช้แสดงใต้ลายเซ็นในเอกสารทางการ ตรงกับรูปแบบเอกสารจริงที่ตรวจสอบไว้ — เอกสารจริงไม่มีชื่ออังกฤษกำกับใต้ลายเซ็น)
 * ส่วน name_en ให้ไว้เผื่อใช้งานในอนาคต (เอกสารภาษาอังกฤษ ฯลฯ)
 */
function getOfficerForDate_(positionType, dateIso) {
  const d = new Date(dateIso || nowIso_());
  const officers = readAll_('CouncilOfficers').filter(o => o.position_type === positionType);
  if (!officers.length) return { name: '(ยังไม่ได้ตั้งค่าในระบบ)', name_en: '', position: positionType, signature_url: '' };
  const sorted = officers.sort((a, b) => new Date(b.term_start) - new Date(a.term_start));
  const match = sorted.find(o => new Date(o.term_start) <= d && (!o.term_end || new Date(o.term_end) >= d));
  const chosen = match || sorted[0];
  return { name: chosen.name_th, name_en: chosen.name_en || '', position: chosen.position_label, signature_url: chosen.signature_url || '' };
}

function getPresidentForDate_(dateIso) { return getOfficerForDate_('นายก', dateIso); }
function getSecretaryGeneralForDate_(dateIso) { return getOfficerForDate_('เลขาธิการ', dateIso); }
function getTreasurerForDate_(dateIso) { return getOfficerForDate_('เหรัญญิก', dateIso); }

/** สร้างข้อมูลผู้ดำรงตำแหน่งชุดเริ่มต้น (จากหลักฐานจริงที่เคยตรวจสอบไว้) — รันครั้งเดียวตอน setupAllSheets ถ้าชีตยังว่างอยู่เท่านั้น ไม่ทับข้อมูลที่แอดมินแก้ไขไปแล้ว */
function seedCouncilOfficersIfEmpty_() {
  if (readAll_('CouncilOfficers').length > 0) return;
  const seed = [
    // นายกสภาวิชาชีพสังคมสงเคราะห์
    { position_type: 'นายก', name_th: 'รองศาสตราจารย์อภิญญา เวชยชัย', name_en: 'Assoc. Prof. Apinya Vejjajiva', position_label: 'นายกสภาวิชาชีพสังคมสงเคราะห์', term_start: '2013-01-01', term_end: '2019-12-31', note: 'ชุดที่ 1-2 (2556-2562)' },
    { position_type: 'นายก', name_th: 'ผู้ช่วยศาสตราจารย์อัญมณี บูรณกานนท์', name_en: 'Asst. Prof. Anyamanee Booranakann', position_label: 'นายกสภาวิชาชีพสังคมสงเคราะห์', term_start: '2020-01-01', term_end: '2022-12-31', note: 'ชุดที่ 3 (2562-2565) — ยืนยันคำนำหน้าจากใบอนุญาตจริงเลขที่ 1848' },
    { position_type: 'นายก', name_th: 'ศาสตราจารย์ระพีพรรณ คำหอม', name_en: 'Prof. Rapeepan Khamhom', position_label: 'นายกสภาวิชาชีพสังคมสงเคราะห์', term_start: '2023-01-01', term_end: '2025-12-31', note: 'ชุดที่ 4 (2565-2568)' },
    { position_type: 'นายก', name_th: 'ผู้ช่วยศาสตราจารย์โสภา อ่อนโอภาส', name_en: 'Asst. Prof. Sopha On-opas', position_label: 'นายกสภาวิชาชีพสังคมสงเคราะห์', term_start: '2026-01-01', term_end: '', note: 'ชุดที่ 5 (2568-2571) — วาระปัจจุบัน' },
    // เลขาธิการสภาวิชาชีพสังคมสงเคราะห์
    { position_type: 'เลขาธิการ', name_th: 'ผู้ช่วยศาสตราจารย์โสภา อ่อนโอภาส', name_en: 'Asst. Prof. Sopha On-opas', position_label: 'เลขาธิการสภาวิชาชีพสังคมสงเคราะห์', term_start: '2013-01-01', term_end: '2019-12-31', note: 'ชุดที่ 1-2 (2556-2562)' },
    { position_type: 'เลขาธิการ', name_th: 'นายเสรีภาพ สมทรัพย์', name_en: 'Mr. Sereepap Somsub', position_label: 'เลขาธิการสภาวิชาชีพสังคมสงเคราะห์', term_start: '2020-01-01', term_end: '2022-12-31', note: 'ชุดที่ 3 (2562-2565)' },
    { position_type: 'เลขาธิการ', name_th: 'นางวรรภา ลำเจียกเทศ', name_en: 'Mrs. Wanpha Lamjiakthet', position_label: 'เลขาธิการสภาวิชาชีพสังคมสงเคราะห์', term_start: '2023-01-01', term_end: '2025-12-31', note: 'ชุดที่ 4 (2565-2568)' },
    { position_type: 'เลขาธิการ', name_th: 'นางณัฐวดี ณ มโนรม', name_en: 'Mrs. Nattawadee Na Manorom', position_label: 'เลขาธิการสภาวิชาชีพสังคมสงเคราะห์', term_start: '2026-01-01', term_end: '', note: 'ชุดที่ 5 (2568-2571) — วาระปัจจุบัน' },
    // เหรัญญิกสภาวิชาชีพสังคมสงเคราะห์
    { position_type: 'เหรัญญิก', name_th: 'นางสาวบุญล้อม กานต์ศักดิ์สราญ', name_en: '', position_label: 'เหรัญญิกสภาวิชาชีพสังคมสงเคราะห์', term_start: '2013-01-01', term_end: '2019-12-31', note: 'ชุดที่ 1-2 (2556-2562)' },
    { position_type: 'เหรัญญิก', name_th: 'นางสาวพรประภา สินธุนาวา', name_en: '', position_label: 'เหรัญญิกสภาวิชาชีพสังคมสงเคราะห์', term_start: '2020-01-01', term_end: '2022-12-31', note: 'ชุดที่ 3 (2562-2565)' },
    { position_type: 'เหรัญญิก', name_th: 'นายฐานวัฒน์ พรนิธิดลวัฒน์', name_en: '', position_label: 'เหรัญญิกสภาวิชาชีพสังคมสงเคราะห์', term_start: '2023-01-01', term_end: '2025-12-31', note: 'ชุดที่ 4 (2565-2568) — ตรงกับใบเสร็จตัวอย่างจริง' },
    { position_type: 'เหรัญญิก', name_th: 'นางสาวบุญล้อม กานต์ศักดิ์สราญ', name_en: '', position_label: 'เหรัญญิกสภาวิชาชีพสังคมสงเคราะห์', term_start: '2026-01-01', term_end: '', note: 'ชุดที่ 5 (2568-2571) — วาระปัจจุบัน' }
  ];
  seed.forEach(o => {
    appendRow_('CouncilOfficers', Object.assign({
      officer_id: newId_('OFC'), signature_url: '', created_at: nowIso_(), updated_at: nowIso_()
    }, o));
  });
}

/** แปลงจำนวนเงินเป็นคำอ่านภาษาไทย เช่น 1000 -> "หนึ่งพันบาทถ้วน" — ใช้ในใบเสร็จรับเงิน */
function thaiBahtText_(amount) {
  const digits = ['ศูนย์', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
  const positions = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน', 'ล้าน'];
  function readChunk(numStr) {
    let result = '';
    const len = numStr.length;
    for (let i = 0; i < len; i++) {
      const digit = parseInt(numStr[i], 10);
      const place = len - i - 1;
      if (digit === 0) continue;
      if (place === 0 && digit === 1 && len > 1) { result += 'เอ็ด'; }
      else if (place === 1 && digit === 1) { result += 'สิบ'; }
      else if (place === 1 && digit === 2) { result += 'ยี่สิบ'; }
      else { result += digits[digit] + positions[place]; }
    }
    return result;
  }
  const intAmount = Math.floor(Math.abs(Number(amount) || 0));
  const satang = Math.round((Math.abs(Number(amount) || 0) - intAmount) * 100);
  let baht = '';
  if (intAmount === 0) { baht = 'ศูนย์บาท'; }
  else {
    let numStr = String(intAmount);
    let millionsPart = '';
    while (numStr.length > 6) {
      const chunk = numStr.slice(0, numStr.length - 6);
      millionsPart += readChunk(chunk) + 'ล้าน';
      numStr = numStr.slice(numStr.length - 6);
    }
    baht = millionsPart + readChunk(numStr) + 'บาท';
  }
  return baht + (satang > 0 ? readChunk(String(satang)) + 'สตางค์' : 'ถ้วน');
}

const THAI_MONTHS_ = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

/** จัดรูปแบบวันที่แบบทางการไทย เช่น "4 พฤษภาคม 2565" (ใช้ในเอกสารใบอนุญาตจริง) */
function formatThaiFormalDate_(dateIso) {
  const d = new Date(dateIso);
  return d.getDate() + ' เดือน ' + THAI_MONTHS_[d.getMonth()] + ' พ.ศ. ' + (d.getFullYear() + 543);
}

/** ดึงไฟล์รูปจาก URL ของ Google Drive (เช่น photo_url ที่ได้จาก action_uploadDocument) มาเป็น Blob เพื่อฝังใน PDF
 * คืนค่า null ถ้าดึงไม่ได้ (ลิงก์ไม่ใช่ไฟล์ Drive ของเรา หรือไม่มีสิทธิ์เข้าถึง) — ให้ผู้เรียกจัดการ fallback เอง
 */
function getDriveBlobFromUrl_(url) {
  if (!url) return null;
  const m = url.match(/[-\w]{25,}/);
  if (!m) return null;
  try {
    return DriveApp.getFileById(m[0]).getBlob();
  } catch (e) {
    return null;
  }
}

/**
 * สร้างแท็บทั้งหมดตาม SHEET_DEFINITIONS พร้อมหัวคอลัมน์ — รันครั้งเดียวตอนติดตั้งระบบ
 * วิธีรัน: เปิดไฟล์นี้ใน Apps Script editor > เลือกฟังก์ชัน "setupAllSheets" จาก dropdown ด้านบน > กด Run
 * (ครั้งแรกจะขอ authorize สิทธิ์เข้าถึงชีต — กด Allow)
 * รันซ้ำได้อย่างปลอดภัย: แท็บที่มีอยู่แล้วจะไม่ถูกลบข้อมูล มีแค่ header/รูปแบบที่จะถูกจัดใหม่
 */
/** ย้ายข้อมูลชีต CouncilOfficers จาก schema เก่า (คอลัมน์ 'name' รวมไทย/อังกฤษไว้ด้วยกัน) มาเป็น schema ใหม่ (name_th + name_en แยกกัน)
 * ทำงานเฉพาะตอนตรวจพบว่าชีตยังเป็น schema เก่าเท่านั้น (มีคอลัมน์ชื่อ 'name' ตรงๆ) — ถ้าเป็น schema ใหม่อยู่แล้วหรือยังไม่มีชีต จะข้ามไปเงียบๆ
 * ต้องรันขั้นนี้ก่อนเขียนหัวคอลัมน์ใหม่ทับเสมอ ไม่งั้นข้อมูลแถวเดิมจะเลื่อนตำแหน่งผิดคอลัมน์ (เพราะ schema ใหม่แทรกคอลัมน์ name_en เพิ่มเข้ามาระหว่างกลาง)
 */
function migrateCouncilOfficersNameColumn_() {
  const sh = getSS_().getSheetByName('CouncilOfficers');
  if (!sh || sh.getLastColumn() === 0) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const nameColIdx = headers.indexOf('name'); // ตำแหน่ง 0-based ของคอลัมน์ 'name' แบบเก่า
  if (nameColIdx === -1) return; // ไม่พบคอลัมน์ชื่อนี้ — เป็น schema ใหม่อยู่แล้ว หรือยังไม่เคยมีข้อมูล ไม่ต้อง migrate
  sh.insertColumnAfter(nameColIdx + 1); // แทรกคอลัมน์ว่างสำหรับ name_en ไว้ทันทีหลังคอลัมน์ name (ข้อมูลแถวอื่นเลื่อนขวาไปเองอัตโนมัติ ไม่กระทบค่าที่มีอยู่)
  sh.getRange(1, nameColIdx + 1).setValue('name_th'); // เปลี่ยนหัวคอลัมน์เดิมชั่วคราว (setupAllSheets จะเขียนทับให้ตรงเป๊ะอีกทีอยู่ดี)
}

function setupAllSheets() {
  const created = [];
  const existed = [];

  // ย้ายข้อมูลคอลัมน์ 'name' เดิมของ CouncilOfficers (schema เก่า ก่อนแยกไทย/อังกฤษ) ให้ตรงกับ schema ใหม่ก่อนเขียนหัวคอลัมน์ทับ
  // ถ้าไม่ทำขั้นนี้ก่อน การเขียนหัวคอลัมน์ทับตรงๆ จะทำให้ข้อมูลเดิมเลื่อนตำแหน่งผิด (เพราะ schema ใหม่แทรกคอลัมน์ name_en เพิ่มเข้ามา)
  migrateCouncilOfficersNameColumn_();

  Object.keys(SHEET_DEFINITIONS).forEach(name => {
    const headers = SHEET_DEFINITIONS[name];
    let sh = getSS_().getSheetByName(name);

    if (!sh) {
      sh = getSS_().insertSheet(name);
      created.push(name);
    } else {
      existed.push(name);
    }

    // เขียน/ซ่อมหัวคอลัมน์แถวที่ 1 โดยไม่แตะข้อมูลแถวอื่น
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);

    // ล้างคอลัมน์เกินที่ไม่ได้อยู่ใน schema (กันแท็บที่เคยแก้มือแล้วเพี้ยน)
    const maxCols = sh.getMaxColumns();
    if (maxCols > headers.length) {
      sh.deleteColumns(headers.length + 1, maxCols - headers.length);
    }

    // จัดรูปแบบหัวตาราง
    const headerRange = sh.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold')
      .setBackground('#164a4f')
      .setFontColor('#ffffff')
      .setVerticalAlignment('middle');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, headers.length);

    // ตั้งค่าคอลัมน์ที่เป็น "ตัวเลขแต่จริงๆ คือรหัส/เบอร์" (เบอร์โทร, เลขบัตรประชาชน, รหัสไปรษณีย์ ฯลฯ) ให้เป็นข้อความล้วน (Plain text)
    // กัน Google Sheets ตัดเลข 0 นำหน้าทิ้งอัตโนมัติเวลาค่าที่เขียนเข้าไปดูเหมือนตัวเลขล้วน (เช่น เบอร์โทร 0891234567 → 891234567)
    const digitStringColumns = DIGIT_STRING_COLUMNS[name];
    if (digitStringColumns && digitStringColumns.length) {
      digitStringColumns.forEach(colName => {
        const colIdx = headers.indexOf(colName);
        if (colIdx === -1) return;
        sh.getRange(2, colIdx + 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@');
      });
    }

    // ป้องกันแถวหัวตารางไม่ให้ถูกแก้ไขโดยไม่ตั้งใจ
    const protection = headerRange.protect().setDescription('header row — ห้ามแก้ไข');
    protection.setWarningOnly(true);
  });

  // ลบแท็บ "Sheet1" เริ่มต้นของ Google ถ้ายังว่างอยู่และมีแท็บอื่นสร้างสำเร็จแล้ว
  const defaultSheet = getSS_().getSheetByName('Sheet1');
  if (defaultSheet && defaultSheet.getLastRow() === 0 && getSS_().getSheets().length > 1) {
    getSS_().deleteSheet(defaultSheet);
  }

  // สร้างข้อมูลนายกสภาฯ/เลขาธิการฯ/เหรัญญิกฯ ชุดเริ่มต้นให้อัตโนมัติ ถ้าชีต CouncilOfficers ยังว่างอยู่ (ไม่ทับข้อมูลที่แอดมินเคยแก้ไขไปแล้ว)
  seedCouncilOfficersIfEmpty_();

  const summary = 'สร้างใหม่: ' + (created.join(', ') || '-') + '\nมีอยู่แล้ว (ซ่อมหัวคอลัมน์ให้): ' + (existed.join(', ') || '-');
  Logger.log(summary);
  return summary;
}

// ---------- ข้อมูลเดโม (สำหรับทดลองใช้ระบบเท่านั้น) ----------

/**
 * สร้างข้อมูลตัวอย่างสำหรับทดลองใช้ทั้งระบบ — รันครั้งเดียวหลัง setupAllSheets สำเร็จแล้ว
 * วิธีรัน: เลือกฟังก์ชัน "seedDemoData" จาก dropdown ใน Apps Script editor > กด Run
 * รันซ้ำได้อย่างปลอดภัย: ถ้าเจอข้อมูลเดโมอยู่แล้วจะไม่สร้างซ้ำ
 * ข้อมูลทั้งหมดเป็นข้อมูลสมมติ (เลขบัตร/อีเมล/เบอร์โทรไม่ใช่ของจริง) ใช้ตัวขึ้นต้น 11000000000xx เพื่อให้แยกออกจากข้อมูลจริงง่าย
 */
function seedDemoData() {
  const existing = readAll_('Members').find(m => m.national_id === '1100000000001');
  if (existing) {
    return 'มีข้อมูลเดโมอยู่แล้ว ไม่สร้างซ้ำ (ถ้าต้องการล้างข้อมูลเดโม ให้รัน clearDemoData แล้วรัน seedDemoData ใหม่)';
  }

  const demoMembers = [
    { national_id: '1100000000001', prefix_th: 'นาย', first_name_th: 'สมชาย', last_name_th: 'ใจดี',
      member_type: 'สามัญ', member_status: 'ปกติ', work_province: 'กรุงเทพมหานคร',
      email: 'somchai.demo@example.com', phone: '0890000001', username: 'demo_member1', password: 'Demo1234',
      expireYears: 3, licenseStatus: 'ปกติ', ceCat1: 42, ceCat2: 12, useProfessionalPrefix: true, photoConsent: true,
      photoUrl: 'https://i.pravatar.cc/200?img=12' }, // ผ่านเกณฑ์ CE ครบทั้ง 2 หมวดแล้ว + เปิดใช้คำนำหน้า สค.ร. + ยินยอมแสดงรูปสาธารณะ + มีรูปถ่ายจริงให้ทดสอบว่ารูปแสดงครบทุกจุด (เดโมทดสอบระบบ)
    { national_id: '1100000000002', prefix_th: 'นาง', first_name_th: 'สมหญิง', last_name_th: 'รักงาน',
      member_type: 'สามัญ', member_status: 'ปกติ', work_province: 'เชียงใหม่',
      email: 'somying.demo@example.com', phone: '0890000002', username: 'demo_member2', password: 'Demo1234',
      expireDays: 15, licenseStatus: 'ใกล้หมดอายุ', ceCat1: 45, ceCat2: 3 }, // ใกล้หมดอายุ + ผ่านหมวด 1 แต่ไม่ผ่านหมวด 2 (แสดงว่าต้องผ่าน "ทั้ง 2 หมวด" ไม่ใช่แค่รวมกันครบ) — ไว้ทดสอบระบบแจ้งเตือน (TOR ข้อ 3.3) และแถบ CE (ข้อ 6.3)
    { national_id: '1100000000003', prefix_th: 'นาย', first_name_th: 'วิชัย', last_name_th: 'หยุดพักงาน',
      member_type: 'สมทบ', member_status: 'พักใช้', work_province: 'ขอนแก่น',
      email: 'wichai.demo@example.com', phone: '0890000003', username: '', password: '',
      expireYears: 1, licenseStatus: 'พักใช้' },
    { national_id: '1100000000004', prefix_th: 'นางสาว', first_name_th: 'อรุณี', last_name_th: 'สิ้นสภาพ',
      member_type: 'สามัญ', member_status: 'เพิกถอน', work_province: 'สงขลา',
      email: 'arunee.demo@example.com', phone: '0890000004', username: '', password: '',
      expireYears: -1, licenseStatus: 'เพิกถอน' },
    { national_id: '1100000000005', prefix_th: 'นาย', first_name_th: 'ประเสริฐ', last_name_th: 'รอขึ้นทะเบียน',
      member_type: 'สามัญ', member_status: 'ปกติ', work_province: 'นครราชสีมา',
      email: 'prasert.demo@example.com', phone: '0890000005', username: 'demo_member3', password: 'Demo1234',
      expireYears: 4, noLicense: true }, // เป็นสมาชิกสามัญปกติ แต่ยังไม่เคยขึ้นทะเบียนรับใบอนุญาต — ไว้ทดสอบ TOR ข้อ 4.1
    { national_id: '1100000000006', prefix_th: 'นางสาว', first_name_th: 'มาลี', last_name_th: 'หมดอายุใบอนุญาต',
      member_type: 'สามัญ', member_status: 'ปกติ', work_province: 'ภูเก็ต',
      email: 'malee.demo@example.com', phone: '0890000006', username: 'demo_member4', password: 'Demo1234',
      expireYears: 4, expireDays: undefined, licenseStatus: 'หมดอายุ', licenseExpireDays: -30 }, // ใบอนุญาตหมดอายุแล้ว — ไว้ทดสอบต่อหลังหมดอายุ + Void เลข (TOR ข้อ 4.2, 4.4.4)

    // ---------- สมาชิกเพิ่มเติม ครอบคลุมทุกขั้นของ "ระบบขึ้นทะเบียนรับใบอนุญาต" (TOR ข้อ 4.1) ----------
    { national_id: '1100000000007', prefix_th: 'นาย', first_name_th: 'กิตติ', last_name_th: 'ก่อนขึ้นทะเบียน',
      member_type: 'สามัญ', member_status: 'ปกติ', work_province: 'อุบลราชธานี',
      email: 'kitti.demo@example.com', phone: '0890000007', username: 'demo_member5', password: 'Demo1234',
      expireYears: 4, noLicense: true, gender: 'LGBTQ+' }, // ยังไม่เคยยื่นคำขอเลย — สถานะ "ก่อนขึ้นทะเบียน" ล้วนๆ (เดโมตัวอย่างเพศ LGBTQ+)
    { national_id: '1100000000008', prefix_th: 'นางสาว', first_name_th: 'ดวงใจ', last_name_th: 'รอชำระค่าขึ้นทะเบียน',
      member_type: 'สามัญ', member_status: 'ปกติ', work_province: 'เชียงราย',
      email: 'duangjai.demo@example.com', phone: '0890000008', username: 'demo_member6', password: 'Demo1234',
      expireYears: 4, noLicense: true }, // ยื่นคำขอ+อนุมัติแล้ว รอชำระเงิน
    { national_id: '1100000000009', prefix_th: 'นาย', first_name_th: 'ธีระ', last_name_th: 'รอตรวจสลิปขึ้นทะเบียน',
      member_type: 'สามัญ', member_status: 'ปกติ', work_province: 'สุราษฎร์ธานี',
      email: 'theera.demo@example.com', phone: '0890000009', username: 'demo_member7', password: 'Demo1234',
      expireYears: 4, noLicense: true }, // แนบสลิปแล้ว รอแอดมินตรวจสอบ
    { national_id: '1100000000010', prefix_th: 'นางสาว', first_name_th: 'ปิยะดา', last_name_th: 'ขึ้นทะเบียนสำเร็จ',
      member_type: 'สามัญ', member_status: 'ปกติ', work_province: 'นครปฐม',
      email: 'piyada.demo@example.com', phone: '0890000010', username: 'demo_member8', password: 'Demo1234',
      expireYears: 4, licenseStatus: 'ปกติ' }, // ขึ้นทะเบียนสำเร็จผ่าน workflow เต็มรูปแบบ มีใบอนุญาตแล้ว

    // ---------- สมาชิกเพิ่มเติม ครอบคลุมทุกขั้นของ "ระบบต่ออายุใบอนุญาต" (TOR ข้อ 4.2) ----------
    { national_id: '1100000000011', prefix_th: 'นาย', first_name_th: 'สุรชัย', last_name_th: 'ก่อนต่ออายุ',
      member_type: 'สามัญ', member_status: 'ปกติ', work_province: 'สงขลา',
      email: 'surachai.demo@example.com', phone: '0890000011', username: 'demo_member9', password: 'Demo1234',
      expireYears: 4, licenseExpireDays: 60, licenseStatus: 'ปกติ' }, // ใบอนุญาตอยู่ในช่วง 180 วัน แต่ยังไม่เคยยื่นต่ออายุ
    { national_id: '1100000000012', prefix_th: 'นางสาว', first_name_th: 'อรวรรณ', last_name_th: 'รอตรวจสลิปต่ออายุ',
      member_type: 'สามัญ', member_status: 'ปกติ', work_province: 'ลำปาง',
      email: 'orawan.demo@example.com', phone: '0890000012', username: 'demo_member10', password: 'Demo1234',
      expireYears: 4, licenseExpireDays: 60, licenseStatus: 'ปกติ' }, // ยื่นต่อก่อนหมดอายุด้วย CE + แนบสลิปแล้ว รอแอดมินตรวจสอบ
    { national_id: '1100000000013', prefix_th: 'นาย', first_name_th: 'ณรงค์', last_name_th: 'รอชำระค่าต่ออายุ',
      member_type: 'สามัญ', member_status: 'ปกติ', work_province: 'พิษณุโลก',
      email: 'narong.demo@example.com', phone: '0890000013', username: 'demo_member11', password: 'Demo1234',
      expireYears: 4, licenseExpireDays: -20, licenseStatus: 'หมดอายุ' }, // ยื่นต่อหลังหมดอายุ(สอบใหม่) + แอดมินอนุมัติแล้ว รอชำระเงิน
    { national_id: '1100000000014', prefix_th: 'นางสาว', first_name_th: 'เบญจมาศ', last_name_th: 'ต่ออายุสำเร็จ',
      member_type: 'สามัญ', member_status: 'ปกติ', work_province: 'ระยอง',
      email: 'benjamas.demo@example.com', phone: '0890000014', username: 'demo_member12', password: 'Demo1234',
      expireYears: 4, licenseExpireDays: 730, licenseStatus: 'ปกติ' }, // ต่ออายุหลังหมดอายุสำเร็จ — ได้เลขใบอนุญาตใหม่แล้ว (เลขเดิมถูก Void)

    // ---------- สมาชิกเพิ่มเติม ครอบคลุมสถานะสมทบปกติ + การเปลี่ยนประเภทสมาชิก (TOR ข้อ 1.1(7)) ----------
    { national_id: '1100000000015', prefix_th: 'นาย', first_name_th: 'ชาญ', last_name_th: 'สมทบปกติ',
      member_type: 'สมทบ', member_status: 'ปกติ', work_province: 'อยุธยา',
      email: 'chan.demo@example.com', phone: '0890000015', username: 'demo_member13', password: 'Demo1234',
      expireYears: 1, noLicense: true }, // สมาชิกสมทบสถานะปกติ (ตัวอย่างสะอาด ไม่มีคำขออื่นค้าง) — ไว้ทดสอบสิทธิ์/ข้อจำกัดของสมาชิกสมทบ
    { national_id: '1100000000016', prefix_th: 'นางสาว', first_name_th: 'พิมพ์ใจ', last_name_th: 'รอเปลี่ยนประเภท',
      member_type: 'สมทบ', member_status: 'ปกติ', work_province: 'นนทบุรี',
      email: 'pimjai.demo@example.com', phone: '0890000016', username: 'demo_member14', password: 'Demo1234',
      expireYears: 1, noLicense: true }, // สมทบปกติ ยื่นคำขอเปลี่ยนประเภทเป็นสามัญแล้ว สถานะ "รอตรวจสอบ" — ไว้ทดสอบ TOR ข้อ 1.1(7)
    { national_id: '1100000000017', prefix_th: 'นาย', first_name_th: 'วิสุทธิ์', last_name_th: 'เปลี่ยนประเภทสำเร็จ',
      member_type: 'สามัญ', member_status: 'ปกติ', work_province: 'สมุทรปราการ',
      email: 'wisut.demo@example.com', phone: '0890000017', username: 'demo_member15', password: 'Demo1234',
      expireYears: 4, noLicense: true }, // เดิมเป็นสมทบ เปลี่ยนเป็นสามัญสำเร็จแล้ว (มีประวัติคำขอเปลี่ยนประเภทที่สำเร็จแล้วผูกไว้) — ไว้ทดสอบว่าประวัติแสดงถูกต้อง
    { national_id: '1100000000018', prefix_th: 'นาง', first_name_th: 'สมศรี', last_name_th: 'รอผูกบัญชี',
      member_type: 'สามัญ', member_status: 'ปกติ', work_province: 'เชียงราย',
      email: 'somsri.demo@example.com', phone: '0890000018', username: '', password: '',
      expireYears: 3, licenseStatus: 'ปกติ' }, // สมาชิกสามัญปกติ มีข้อมูลอยู่แล้วในระบบ (เหมือนนำเข้าจากฐานข้อมูลเดิม) แต่ยังไม่เคยสร้างบัญชีเข้าใช้งานเอง — ไว้ทดสอบฟีเจอร์ "เป็นสมาชิกอยู่แล้ว? คลิกที่นี่" ที่หน้าสมัครใช้งานบัญชี

    // ---------- ระบบขอเปลี่ยนชื่อ-นามสกุล (ภาษาไทย) — ครอบคลุมทุกสถานะ ----------
    { national_id: '1100000000019', prefix_th: 'นางสาว', first_name_th: 'จันทร์เพ็ญ', last_name_th: 'รอเปลี่ยนชื่อ',
      member_type: 'สามัญ', member_status: 'ปกติ', work_province: 'นครสวรรค์',
      email: 'janpen.demo@example.com', phone: '0890000019', username: 'demo_member16', password: 'Demo1234',
      expireYears: 3, licenseStatus: 'ปกติ' }, // ยื่นคำขอเปลี่ยนชื่อ-นามสกุลแล้ว สถานะ "รอตรวจสอบ" พร้อมเอกสารแนบ
    { national_id: '1100000000020', prefix_th: 'นาง', first_name_th: 'อำไพ', last_name_th: 'รอเอกสารเพิ่มเติม',
      member_type: 'สามัญ', member_status: 'ปกติ', work_province: 'อุดรธานี',
      email: 'ampai.demo@example.com', phone: '0890000020', username: 'demo_member17', password: 'Demo1234',
      expireYears: 3, licenseStatus: 'ปกติ' }, // ยื่นคำขอเปลี่ยนชื่อแล้ว เจ้าหน้าที่ตรวจแล้วขอเอกสารเพิ่มเติม สถานะ "รอเอกสารเพิ่มเติม"
    { national_id: '1100000000021', prefix_th: 'นางสาว', first_name_th: 'ศิริพร', last_name_th: 'ถูกไม่อนุมัติ',
      member_type: 'สามัญ', member_status: 'ปกติ', work_province: 'สุรินทร์',
      email: 'siriporn.demo@example.com', phone: '0890000021', username: 'demo_member18', password: 'Demo1234',
      expireYears: 3, licenseStatus: 'ปกติ' }, // ยื่นคำขอเปลี่ยนชื่อแล้ว เจ้าหน้าที่ไม่อนุมัติพร้อมเหตุผล — ไว้ทดสอบว่าสมาชิกเห็นเหตุผลแบบเรียลไทม์
    { national_id: '1100000000022', prefix_th: 'นาง', first_name_th: 'รัตนา', last_name_th: 'เปลี่ยนชื่อสำเร็จแล้ว',
      member_type: 'สามัญ', member_status: 'ปกติ', work_province: 'เพชรบุรี',
      email: 'rattana.demo@example.com', phone: '0890000022', username: 'demo_member19', password: 'Demo1234',
      expireYears: 3, licenseStatus: 'ปกติ' } // ชื่อปัจจุบันคือชื่อ "ใหม่" ที่อนุมัติแล้ว — ไว้ทดสอบว่าประวัติการเปลี่ยนแปลง (History) แสดงชื่อเดิม→ชื่อใหม่+ผู้อนุมัติ ทั้งฝั่งสมาชิกและแอดมิน
  ];

  const demoMemberIds = {};
  const demoUserIds = {};
  demoMembers.forEach(d => {
    const memberId = newId_('MBR');
    demoMemberIds[d.national_id] = memberId;
    const appliedDate = addYearsIso_(nowIso_(), -(d.expireYears && d.expireYears > 0 ? d.expireYears - 1 : 1));
    const expireDate = d.expireDays !== undefined
      ? new Date(Date.now() + d.expireDays * 24 * 3600 * 1000).toISOString()
      : addYearsIso_(nowIso_(), d.expireYears || 1);
    const licenseNo = 'LC' + (new Date().getFullYear() + 543) + '-DEMO' + d.national_id.slice(-2);
    const licenseExpireDate = d.licenseExpireDays !== undefined
      ? new Date(Date.now() + d.licenseExpireDays * 24 * 3600 * 1000).toISOString()
      : expireDate;

    appendRow_('Members', {
      member_id: memberId, member_no: generateMemberNumber_(d.member_type), national_id: d.national_id,
      prefix_th: d.prefix_th, first_name_th: d.first_name_th, last_name_th: d.last_name_th,
      gender: d.gender || (d.prefix_th === 'นาย' ? 'ชาย' : (d.prefix_th === 'นาง' || d.prefix_th === 'นางสาว' ? 'หญิง' : 'ไม่ระบุ')),
      phone: d.phone, email: d.email, work_province: d.work_province,
      workplace_name: 'หน่วยงานเดโม ' + d.first_name_th, workplace_district: 'เขต/อำเภอเดโม',
      workplace_province: d.work_province,
      member_type: d.member_type, member_status: d.member_status,
      applied_date: appliedDate, membership_expire_date: expireDate,
      license_no: d.noLicense ? '' : licenseNo, license_status: d.noLicense ? '' : d.licenseStatus,
      ce_score_cat1: d.ceCat1 !== undefined ? d.ceCat1 : '', ce_score_cat2: d.ceCat2 !== undefined ? d.ceCat2 : '',
      use_professional_prefix: d.useProfessionalPrefix ? 'true' : 'false',
      photo_public_consent: d.photoConsent ? 'true' : 'false', photo_url: d.photoUrl || '', documents_json: '[]',
      created_at: nowIso_(), updated_at: nowIso_()
    });
    if (!d.noLicense) {
      appendRow_('Licenses', {
        license_id: newId_('LIC'), member_id: memberId, license_no: licenseNo,
        issue_date: appliedDate, expire_date: licenseExpireDate, status: d.licenseStatus,
        renewal_docs_json: '[]', renewed_at: '', revoked_reason: d.member_status === 'เพิกถอน' ? 'ข้อมูลเดโม' : ''
      });
    }
    if (d.username) {
      const userId = newId_('USR');
      demoUserIds[d.national_id] = userId;
      appendRow_('Users', {
        user_id: userId, username: d.username, password_hash: hashPassword_(d.password),
        role: 'member', member_id: memberId, thaid_pid: '', status: 'active',
        created_at: nowIso_(), last_login: ''
      });
    }
  });

  // คำสั่งพักใช้/เพิกถอนตัวอย่าง (มติคณะกรรมการจรรยาบรรณ) — ผูกกับสมาชิกเดโมที่สถานะ พักใช้ / เพิกถอน (หลักฐานเป็นตัวอย่าง ไม่มีไฟล์จริง)
  [['1100000000003', 'พักใช้'], ['1100000000004', 'เพิกถอน']].forEach(pair => {
    const mid = demoMemberIds[pair[0]];
    if (!mid) return;
    const lic = readAll_('Licenses').find(l => l.member_id === mid);
    if (!lic) return;
    const mem = findRowById_('Members', 'member_id', mid);
    const todayD = todayYmd_();
    const addDays = n => Utilities.formatDate(new Date(Date.now() + n * 86400000), 'Asia/Bangkok', 'yyyy-MM-dd');
    appendRow_('LicenseSanctions', {
      sanction_id: newId_('SNC'), member_id: mid, license_id: lic.license_id, license_no: lic.license_no,
      member_name: memberDisplayName_(mem), action_type: pair[1], status: 'มีผลบังคับใช้',
      ethics_case_no: pair[1] === 'พักใช้' ? 'จร.12/2569 (ตัวอย่าง)' : 'จร.25/2569 (ตัวอย่าง)',
      ethics_meeting_no: pair[1] === 'พักใช้' ? 'ครั้งที่ 8/2569' : 'ครั้งที่ 9/2569',
      decision_date: addDays(-20), effective_date: addDays(-15), suspend_until: pair[1] === 'พักใช้' ? addDays(75) : '',
      reason: pair[1] === 'พักใช้' ? 'ข้อมูลเดโม: ละเมิดจรรยาบรรณวิชาชีพข้อ 5 — พักใช้ 90 วัน' : 'ข้อมูลเดโม: ประพฤติผิดจรรยาบรรณวิชาชีพร้ายแรง',
      remarks: 'ข้อมูลเดโม', prev_license_status: 'ปกติ',
      evidence_json: JSON.stringify([
        { name: '(ตัวอย่าง) หนังสือแจ้งมติคณะกรรมการจรรยาบรรณ.pdf', url: '', doc_type: 'มติ/หนังสือคณะกรรมการจรรยาบรรณ', note: 'ไฟล์ตัวอย่าง', uploaded_by: 'demo_staff', uploaded_at: nowIso_() },
        { name: '(ตัวอย่าง) รายงานการประชุม.pdf', url: '', doc_type: 'รายงานการประชุม', note: 'ไฟล์ตัวอย่าง', uploaded_by: 'demo_staff', uploaded_at: nowIso_() }
      ]),
      lifted_at: '', lifted_by: '', lift_type: '', lift_reason: '', lift_evidence_json: '[]',
      created_by: 'demo_staff', created_at: nowIso_(), updated_at: nowIso_(), updated_by: 'demo_staff'
    });
  });

  // ประวัติการต่ออายุตัวอย่าง (TOR ข้อ 3.1) — ผูกกับ demo_member1 ที่เคยต่ออายุมาแล้ว 1 ครั้ง
  const member1Id = demoMemberIds['1100000000001'];
  const user1Id = demoUserIds['1100000000001'];
  if (member1Id) {
    appendRow_('MembershipRenewals', {
      renewal_id: newId_('RNW'), member_id: member1Id,
      renewed_at: addYearsIso_(nowIso_(), -2), old_expire_date: addYearsIso_(nowIso_(), -2),
      new_expire_date: addYearsIso_(nowIso_(), -2 + 5), years: 5,
      payment_id: '', renewed_by: 'demo_staff', note: 'ต่ออายุครั้งแรก (ข้อมูลเดโม)'
    });
  }

  // ---------- ข้อมูลเดโมสำหรับ "บัตรสมาชิก" (TOR ข้อ 5) — ตัวอย่างประวัติการต่ออายุผ่านหน้าบัตรสมาชิก ----------
  if (member1Id && user1Id) {
    const cardRenAppId = newId_('MREN');
    appendRow_('Applications', {
      application_id: cardRenAppId, applicant_type: 'ต่ออายุสมาชิก', applicant_user_id: user1Id,
      ref_data_json: JSON.stringify({ member_id: member1Id, duration: '5', email: 'somchai.demo@example.com', name: 'นาย สมชาย ใจดี' }),
      documents_json: '[]', status: 'ต่ออายุสำเร็จ', reject_reason: '', fee_amount: MEMBERSHIP_FEE_TABLE['สามัญ']['5'],
      submitted_at: addYearsIso_(nowIso_(), -2), reviewed_by: 'ระบบ (ต่ออายุสมาชิกโดยตรง)', reviewed_at: addYearsIso_(nowIso_(), -2), approved_member_id: member1Id
    });
    appendRow_('Payments', {
      payment_id: newId_('PAY'), ref_type: 'application', ref_id: cardRenAppId, member_id: member1Id,
      amount: MEMBERSHIP_FEE_TABLE['สามัญ']['5'], slip_url: 'https://example.com/demo-slip-placeholder.jpg',
      method: 'โอนเงิน', channel: 'โอนเงินผ่านธนาคาร (ตรวจสอบสลิปด้วยมือ)', status: 'ชำระแล้ว', verified_by: 'demo_staff', verified_at: addYearsIso_(nowIso_(), -2), paid_at: addYearsIso_(nowIso_(), -2)
    });
  }

  // ---------- ข้อมูลเดโมสำหรับระบบใบอนุญาต (TOR ข้อ 4) ----------
  const member2Id = demoMemberIds['1100000000002'];
  const member5Id = demoMemberIds['1100000000005'];
  const member6Id = demoMemberIds['1100000000006'];
  const user2Id = demoUserIds['1100000000002'];
  const user5Id = demoUserIds['1100000000005'];
  const user6Id = demoUserIds['1100000000006'];

  // 4.1: คำขอขึ้นทะเบียนรับใบอนุญาตใหม่ (demo_member3 / ประเสริฐ) — สถานะ "รอตรวจสอบ" พร้อมแนบผลสอบข้อเขียน 2 วิชา + ผลสอบสัมภาษณ์ครบตามที่กำหนด (TOR ข้อ 4.1.1)
  if (member5Id && user5Id) {
    appendRow_('Applications', {
      application_id: newId_('REG'), applicant_type: 'ขึ้นทะเบียน', applicant_user_id: user5Id,
      ref_data_json: JSON.stringify({ member_id: member5Id, email: 'prasert.demo@example.com', name: 'นาย ประเสริฐ รอขึ้นทะเบียน' }),
      documents_json: JSON.stringify([
        { name: 'ผลสอบวิชา1.pdf', doc_type: 'ผลสอบข้อเขียน ' + EXAM_SUBJECTS[0], url: 'https://example.com/demo-doc-placeholder.pdf' },
        { name: 'ผลสอบวิชา2.pdf', doc_type: 'ผลสอบข้อเขียน ' + EXAM_SUBJECTS[1], url: 'https://example.com/demo-doc-placeholder.pdf' },
        { name: 'ผลสอบสัมภาษณ์.pdf', doc_type: EXAM_INTERVIEW_LABEL, url: 'https://example.com/demo-doc-placeholder.pdf' }
      ]), status: 'รอตรวจสอบ', reject_reason: '', fee_amount: LICENSE_REGISTRATION_FEE,
      submitted_at: nowIso_(), reviewed_by: '', reviewed_at: '', approved_member_id: ''
    });
  }

  // 4.2: ต่ออายุก่อนหมดอายุด้วยผล CE (demo_member2 / สมหญิง) — ผ่านการตรวจสอบแล้ว รอชำระเงินทันที (ตาม TOR ข้อ 4.2.1)
  if (member2Id && user2Id) {
    const renAppId = newId_('REN');
    appendRow_('Applications', {
      application_id: renAppId, applicant_type: 'ต่ออายุใบอนุญาต', applicant_user_id: user2Id,
      ref_data_json: JSON.stringify({ member_id: member2Id, method: 'CE_before', isAfterExpiry: false, email: 'somying.demo@example.com', name: 'นาง สมหญิง รักงาน' }),
      documents_json: '[]', status: 'รอชำระเงิน', reject_reason: '', fee_amount: LICENSE_RENEWAL_FEE,
      submitted_at: nowIso_(), reviewed_by: 'ระบบ (ผ่าน CE อัตโนมัติ)', reviewed_at: nowIso_(), approved_member_id: ''
    });
    appendRow_('Payments', {
      payment_id: newId_('PAY'), ref_type: 'application', ref_id: renAppId, member_id: member2Id,
      amount: LICENSE_RENEWAL_FEE, slip_url: '', method: '', status: 'รอชำระ', verified_by: '', verified_at: '', paid_at: ''
    });
  }

  // 4.2: ต่ออายุหลังหมดอายุด้วยสอบข้อเขียนใหม่ (demo_member4 / มาลี) — สถานะ "รอตรวจสอบ" (ต้องรออนุมัติก่อนชำระเงิน)
  if (member6Id && user6Id) {
    appendRow_('Applications', {
      application_id: newId_('REN'), applicant_type: 'ต่ออายุใบอนุญาต', applicant_user_id: user6Id,
      ref_data_json: JSON.stringify({ member_id: member6Id, method: 'exam', isAfterExpiry: true, email: 'malee.demo@example.com', name: 'นางสาว มาลี หมดอายุใบอนุญาต' }),
      documents_json: '[]', status: 'รอตรวจสอบ', reject_reason: '', fee_amount: LICENSE_REGISTRATION_FEE,
      submitted_at: nowIso_(), reviewed_by: '', reviewed_at: '', approved_member_id: ''
    });
  }

  // ========== ครอบคลุมทุกขั้นของ "ขึ้นทะเบียนรับใบอนุญาต" (TOR ข้อ 4.1) — ตั้งแต่ก่อนยื่นจนมีใบอนุญาต ==========
  const member7Id = demoMemberIds['1100000000007']; // กิตติ — ก่อนขึ้นทะเบียน (ไม่มีคำขอใดๆ เลย ไม่ต้องสร้างอะไรเพิ่ม)
  const member8Id = demoMemberIds['1100000000008']; // ดวงใจ — รอชำระเงิน
  const user8Id = demoUserIds['1100000000008'];
  const member9Id = demoMemberIds['1100000000009']; // ธีระ — แนบสลิปแล้ว รอแอดมินตรวจสอบ
  const user9Id = demoUserIds['1100000000009'];
  const member10Id = demoMemberIds['1100000000010']; // ปิยะดา — ขึ้นทะเบียนสำเร็จ
  const user10Id = demoUserIds['1100000000010'];

  if (member8Id && user8Id) {
    const appId = newId_('REG');
    appendRow_('Applications', {
      application_id: appId, applicant_type: 'ขึ้นทะเบียน', applicant_user_id: user8Id,
      ref_data_json: JSON.stringify({ member_id: member8Id, email: 'duangjai.demo@example.com', name: 'นางสาว ดวงใจ รอชำระค่าขึ้นทะเบียน' }),
      documents_json: JSON.stringify([
        { name: 'ผลสอบวิชา1.pdf', doc_type: 'ผลสอบข้อเขียน ' + EXAM_SUBJECTS[0], url: 'https://example.com/demo-doc-placeholder.pdf' },
        { name: 'ผลสอบวิชา2.pdf', doc_type: 'ผลสอบข้อเขียน ' + EXAM_SUBJECTS[1], url: 'https://example.com/demo-doc-placeholder.pdf' },
        { name: 'ผลสอบสัมภาษณ์.pdf', doc_type: EXAM_INTERVIEW_LABEL, url: 'https://example.com/demo-doc-placeholder.pdf' }
      ]), status: 'รอชำระเงิน', reject_reason: '', fee_amount: LICENSE_REGISTRATION_FEE,
      submitted_at: addYearsIso_(nowIso_(), 0), reviewed_by: 'demo_staff', reviewed_at: nowIso_(), approved_member_id: ''
    });
    appendRow_('Payments', {
      payment_id: newId_('PAY'), ref_type: 'application', ref_id: appId, member_id: member8Id,
      amount: LICENSE_REGISTRATION_FEE, slip_url: '', method: '', status: 'รอชำระ', verified_by: '', verified_at: '', paid_at: ''
    });
  }

  if (member9Id && user9Id) {
    const appId = newId_('REG');
    appendRow_('Applications', {
      application_id: appId, applicant_type: 'ขึ้นทะเบียน', applicant_user_id: user9Id,
      ref_data_json: JSON.stringify({ member_id: member9Id, email: 'theera.demo@example.com', name: 'นาย ธีระ รอตรวจสลิปขึ้นทะเบียน' }),
      documents_json: JSON.stringify([
        { name: 'ผลสอบวิชา1.pdf', doc_type: 'ผลสอบข้อเขียน ' + EXAM_SUBJECTS[0], url: 'https://example.com/demo-doc-placeholder.pdf' },
        { name: 'ผลสอบวิชา2.pdf', doc_type: 'ผลสอบข้อเขียน ' + EXAM_SUBJECTS[1], url: 'https://example.com/demo-doc-placeholder.pdf' },
        { name: 'ผลสอบสัมภาษณ์.pdf', doc_type: EXAM_INTERVIEW_LABEL, url: 'https://example.com/demo-doc-placeholder.pdf' }
      ]), status: 'รอชำระเงิน', reject_reason: '', fee_amount: LICENSE_REGISTRATION_FEE,
      submitted_at: nowIso_(), reviewed_by: 'demo_staff', reviewed_at: nowIso_(), approved_member_id: ''
    });
    appendRow_('Payments', {
      payment_id: newId_('PAY'), ref_type: 'application', ref_id: appId, member_id: member9Id,
      amount: LICENSE_REGISTRATION_FEE, slip_url: 'https://example.com/demo-slip-placeholder.jpg',
      method: 'โอนเงิน', status: 'รอตรวจสอบ', verified_by: '', verified_at: '', paid_at: ''
    });
  }

  if (member10Id && user10Id) {
    const appId = newId_('REG');
    const payId = newId_('PAY');
    appendRow_('Applications', {
      application_id: appId, applicant_type: 'ขึ้นทะเบียน', applicant_user_id: user10Id,
      ref_data_json: JSON.stringify({ member_id: member10Id, email: 'piyada.demo@example.com', name: 'นางสาว ปิยะดา ขึ้นทะเบียนสำเร็จ' }),
      documents_json: JSON.stringify([
        { name: 'ผลสอบวิชา1.pdf', doc_type: 'ผลสอบข้อเขียน ' + EXAM_SUBJECTS[0], url: 'https://example.com/demo-doc-placeholder.pdf' },
        { name: 'ผลสอบวิชา2.pdf', doc_type: 'ผลสอบข้อเขียน ' + EXAM_SUBJECTS[1], url: 'https://example.com/demo-doc-placeholder.pdf' },
        { name: 'ผลสอบสัมภาษณ์.pdf', doc_type: EXAM_INTERVIEW_LABEL, url: 'https://example.com/demo-doc-placeholder.pdf' }
      ]), status: 'ขึ้นทะเบียนสำเร็จ', reject_reason: '', fee_amount: LICENSE_REGISTRATION_FEE,
      submitted_at: addYearsIso_(nowIso_(), 0), reviewed_by: 'demo_staff', reviewed_at: nowIso_(), approved_member_id: member10Id
    });
    appendRow_('Payments', {
      payment_id: payId, ref_type: 'application', ref_id: appId, member_id: member10Id,
      amount: LICENSE_REGISTRATION_FEE, slip_url: 'https://example.com/demo-slip-placeholder.jpg',
      method: 'โอนเงิน', channel: 'พร้อมเพย์ (QR PromptPay)', status: 'ชำระแล้ว', verified_by: 'demo_staff', verified_at: nowIso_(), paid_at: nowIso_()
    });
    try { issueReceiptForPayment_(payId, 'demo_staff'); } catch (e) { /* ข้ามถ้าออกใบเสร็จเดโมไม่สำเร็จ ไม่กระทบการ seed ข้อมูลหลัก */ }
  }

  // ========== ระบบสมัครสอบข้อเขียน/สอบสัมภาษณ์ (TOR ข้อ 4 — เปิดได้หลายรอบต่อปี) ==========
  const user7Id = demoUserIds['1100000000007']; // กิตติ — ใช้สมัครสอบข้อเขียนเดโม (รอตรวจสอบ)

  const examRoundWrittenOpen = { round_id: newId_('EXR'), exam_type: 'เขียน', round_label: '2/2569', open_date: nowIso_(), close_date: addDaysIso_(nowIso_(), 30), exam_date: addDaysIso_(nowIso_(), 55), format_options_json: JSON.stringify(['online', 'onsite']), location_note: 'มหาวิทยาลัยธรรมศาสตร์ ศูนย์รังสิต', status: 'เปิดรับสมัคร', note: '', created_by: 'demo_staff', created_at: nowIso_(), updated_at: nowIso_() };
  appendRow_('ExamRounds', examRoundWrittenOpen);
  const examRoundWrittenClosed = { round_id: newId_('EXR'), exam_type: 'เขียน', round_label: '1/2569', open_date: addDaysIso_(nowIso_(), -110), close_date: addDaysIso_(nowIso_(), -73), exam_date: addDaysIso_(nowIso_(), -55), format_options_json: JSON.stringify(['online', 'onsite']), location_note: 'มหาวิทยาลัยธรรมศาสตร์ ศูนย์รังสิต', status: 'เสร็จสิ้น', note: '', created_by: 'demo_staff', created_at: addDaysIso_(nowIso_(), -110), updated_at: addDaysIso_(nowIso_(), -55) };
  appendRow_('ExamRounds', examRoundWrittenClosed);
  const examRoundInterviewOpen = { round_id: newId_('EXR'), exam_type: 'สัมภาษณ์', round_label: '3/2569', open_date: nowIso_(), close_date: addDaysIso_(nowIso_(), 37), exam_date: addDaysIso_(nowIso_(), 66), format_options_json: JSON.stringify(['online', 'onsite']), location_note: 'สำนักงานสภาวิชาชีพสังคมสงเคราะห์', status: 'เปิดรับสมัคร', note: '', created_by: 'demo_staff', created_at: nowIso_(), updated_at: nowIso_() };
  appendRow_('ExamRounds', examRoundInterviewOpen);
  const examRoundInterviewClosed = { round_id: newId_('EXR'), exam_type: 'สัมภาษณ์', round_label: '2/2569', open_date: addDaysIso_(nowIso_(), -110), close_date: addDaysIso_(nowIso_(), -73), exam_date: addDaysIso_(nowIso_(), -37), format_options_json: JSON.stringify(['online', 'onsite']), location_note: 'สำนักงานสภาวิชาชีพสังคมสงเคราะห์', status: 'เสร็จสิ้น', note: '', created_by: 'demo_staff', created_at: addDaysIso_(nowIso_(), -110), updated_at: addDaysIso_(nowIso_(), -37) };
  appendRow_('ExamRounds', examRoundInterviewClosed);

  if (member7Id && user7Id) { // รอตรวจสอบเอกสาร — สมัครสอบข้อเขียนรอบใหม่ (2 วิชา)
    appendRow_('Applications', {
      application_id: newId_('EXA'), applicant_type: 'สมัครสอบข้อเขียน', applicant_user_id: user7Id,
      ref_data_json: JSON.stringify({ member_id: member7Id, round_id: examRoundWrittenOpen.round_id, round_label: examRoundWrittenOpen.round_label, format: 'onsite', email: 'kitti.demo@example.com', name: 'นาย กิตติ รอตรวจสอบใบสมัครสอบ', subjects: EXAM_SUBJECTS.slice() }),
      documents_json: JSON.stringify([
        { name: 'id.pdf', doc_type: 'สำเนาบัตรประชาชน', url: 'https://example.com/demo-doc-placeholder.pdf' },
        { name: 'photo.jpg', doc_type: 'รูปถ่าย', url: 'https://example.com/demo-doc-placeholder.pdf' },
        { name: 'degree.pdf', doc_type: 'สำเนาวุฒิการศึกษา', url: 'https://example.com/demo-doc-placeholder.pdf' }
      ]), status: 'รอตรวจสอบ', reject_reason: '', fee_amount: EXAM_SUBJECT_FEE * 2,
      submitted_at: nowIso_(), reviewed_by: '', reviewed_at: '', approved_member_id: ''
    });
  }

  if (member8Id && user8Id) { // รอชำระเงิน — อนุมัติแล้ว สมัครสอบข้อเขียน 1 วิชา
    const appId = newId_('EXA');
    appendRow_('Applications', {
      application_id: appId, applicant_type: 'สมัครสอบข้อเขียน', applicant_user_id: user8Id,
      ref_data_json: JSON.stringify({ member_id: member8Id, round_id: examRoundWrittenOpen.round_id, round_label: examRoundWrittenOpen.round_label, format: 'online', email: 'duangjai.demo@example.com', name: 'นางสาว ดวงใจ รอชำระค่าสมัครสอบ', subjects: [EXAM_SUBJECTS[0]] }),
      documents_json: JSON.stringify([
        { name: 'id.pdf', doc_type: 'สำเนาบัตรประชาชน', url: 'https://example.com/demo-doc-placeholder.pdf' },
        { name: 'photo.jpg', doc_type: 'รูปถ่าย', url: 'https://example.com/demo-doc-placeholder.pdf' },
        { name: 'degree.pdf', doc_type: 'สำเนาวุฒิการศึกษา', url: 'https://example.com/demo-doc-placeholder.pdf' }
      ]), status: 'รอชำระเงิน', reject_reason: '', fee_amount: EXAM_SUBJECT_FEE,
      submitted_at: nowIso_(), reviewed_by: 'demo_staff', reviewed_at: nowIso_(), approved_member_id: ''
    });
    appendRow_('Payments', {
      payment_id: newId_('PAY'), ref_type: 'application', ref_id: appId, member_id: member8Id,
      amount: EXAM_SUBJECT_FEE, slip_url: '', method: '', status: 'รอชำระ', verified_by: '', verified_at: '', paid_at: ''
    });
  }

  if (member9Id && user9Id) { // รอผลสอบ — ชำระเงินสมัครสอบข้อเขียนแล้ว รอประกาศผล
    const appId = newId_('EXA');
    const payId = newId_('PAY');
    appendRow_('Applications', {
      application_id: appId, applicant_type: 'สมัครสอบข้อเขียน', applicant_user_id: user9Id,
      ref_data_json: JSON.stringify({ member_id: member9Id, round_id: examRoundWrittenClosed.round_id, round_label: examRoundWrittenClosed.round_label, format: 'onsite', email: 'theera.demo@example.com', name: 'นาย ธีระ รอผลสอบข้อเขียน', subjects: EXAM_SUBJECTS.slice(), candidate_no: '0001' }),
      documents_json: JSON.stringify([
        { name: 'id.pdf', doc_type: 'สำเนาบัตรประชาชน', url: 'https://example.com/demo-doc-placeholder.pdf' },
        { name: 'photo.jpg', doc_type: 'รูปถ่าย', url: 'https://example.com/demo-doc-placeholder.pdf' },
        { name: 'degree.pdf', doc_type: 'สำเนาวุฒิการศึกษา', url: 'https://example.com/demo-doc-placeholder.pdf' }
      ]), status: 'รอผลสอบ', reject_reason: '', fee_amount: EXAM_SUBJECT_FEE * 2,
      submitted_at: addDaysIso_(nowIso_(), -90), reviewed_by: 'demo_staff', reviewed_at: addDaysIso_(nowIso_(), -80), approved_member_id: ''
    });
    appendRow_('Payments', {
      payment_id: payId, ref_type: 'application', ref_id: appId, member_id: member9Id,
      amount: EXAM_SUBJECT_FEE * 2, slip_url: 'https://example.com/demo-slip-placeholder.jpg',
      method: 'โอนเงิน', status: 'ชำระแล้ว', verified_by: 'demo_staff', verified_at: addDaysIso_(nowIso_(), -73), paid_at: addDaysIso_(nowIso_(), -73)
    });
    try { issueReceiptForPayment_(payId, 'demo_staff'); } catch (e) { /* ข้ามถ้าออกใบเสร็จเดโมไม่สำเร็จ ไม่กระทบการ seed ข้อมูลหลัก */ }
  }

  if (member10Id && user10Id) { // สอบผ่าน — ประกาศผลสอบข้อเขียนแล้ว ผ่านทั้ง 2 วิชา
    appendRow_('Applications', {
      application_id: newId_('EXA'), applicant_type: 'สมัครสอบข้อเขียน', applicant_user_id: user10Id,
      ref_data_json: JSON.stringify({
        member_id: member10Id, round_id: examRoundWrittenClosed.round_id, round_label: examRoundWrittenClosed.round_label, format: 'onsite',
        email: 'piyada.demo@example.com', name: 'นางสาว ปิยะดา สอบข้อเขียนผ่าน', subjects: EXAM_SUBJECTS.slice(), candidate_no: '0002',
        subject_results: { [EXAM_SUBJECTS[0]]: 'ผ่าน', [EXAM_SUBJECTS[1]]: 'ผ่าน' },
        exam_result: { kind: 'เขียน', exam_date: examRoundWrittenClosed.exam_date, seat_no: '0002', scores: { [EXAM_SUBJECTS[0]]: 72, [EXAM_SUBJECTS[1]]: 80 },
          subject_results: { [EXAM_SUBJECTS[0]]: 'ผ่าน', [EXAM_SUBJECTS[1]]: 'ผ่าน' }, result: 'ผ่าน', basis: 'ได้คะแนนทั้ง 2 วิชาไม่น้อยกว่าร้อยละ 60', recorded_at: addDaysIso_(nowIso_(), -37), recorded_by: 'demo_staff' }
      }),
      documents_json: JSON.stringify([
        { name: 'id.pdf', doc_type: 'สำเนาบัตรประชาชน', url: 'https://example.com/demo-doc-placeholder.pdf' },
        { name: 'photo.jpg', doc_type: 'รูปถ่าย', url: 'https://example.com/demo-doc-placeholder.pdf' }
      ]), status: 'สอบผ่าน', reject_reason: '', fee_amount: EXAM_SUBJECT_FEE * 2,
      submitted_at: addDaysIso_(nowIso_(), -90), reviewed_by: 'demo_staff', reviewed_at: addDaysIso_(nowIso_(), -37), approved_member_id: ''
    });
  }

  const demoMember16IdForExam = demoMemberIds['1100000000016'];
  if (demoMember16IdForExam) { // ไม่อนุมัติ — สมัครสอบสัมภาษณ์ เอกสารไม่ครบ
    appendRow_('Applications', {
      application_id: newId_('EXA'), applicant_type: 'สมัครสอบสัมภาษณ์', applicant_user_id: demoUserIds['1100000000016'] || '',
      ref_data_json: JSON.stringify({ member_id: demoMember16IdForExam, round_id: examRoundInterviewOpen.round_id, round_label: examRoundInterviewOpen.round_label, format: 'online', email: 'pimjai.demo@example.com', name: 'นางสาว พิมพ์ใจ ใบสมัครสอบสัมภาษณ์ไม่ผ่านตรวจสอบ', essay_category: EXAM_INTERVIEW_ESSAY_CATEGORIES[0] }),
      documents_json: JSON.stringify([
        { name: 'essay.pdf', doc_type: examEssayDocTypes_(EXAM_INTERVIEW_ESSAY_CATEGORIES[0])[0], url: 'https://example.com/demo-doc-placeholder.pdf' },
        { name: 'essay.docx', doc_type: examEssayDocTypes_(EXAM_INTERVIEW_ESSAY_CATEGORIES[0])[1], url: 'https://example.com/demo-doc-placeholder.pdf' }
      ]), status: 'ไม่อนุมัติ', reject_reason: 'ยังไม่แนบหลักฐานการสอบผ่านข้อเขียน 2 วิชา หรือหลักฐานผ่านเกณฑ์ CE', fee_amount: 0,
      submitted_at: addDaysIso_(nowIso_(), -18), reviewed_by: 'demo_staff', reviewed_at: nowIso_(), approved_member_id: ''
    });
  }

  // ========== ระบบจัดการสอบสัมภาษณ์ (ตัวอย่าง) — รอบ 4/2569 ปิดรับสมัครแล้ว มีผู้สอบ 6 คน แบ่ง 2 กลุ่ม กรรมการให้คะแนนไปบางส่วน ==========
  // หน้ากรรมการ interviewer.html: เลือกรอบ 4/2569 → กลุ่ม → กรรมการ 1/2 · กลุ่ม 1 รหัสผ่าน Demo1234 · กลุ่ม 2 ไม่มีรหัสผ่าน
  const examRoundInterviewScoring = { round_id: newId_('EXR'), exam_type: 'สัมภาษณ์', round_label: '4/2569', open_date: addDaysIso_(nowIso_(), -45), close_date: addDaysIso_(nowIso_(), -10), exam_date: addDaysIso_(nowIso_(), 2), format_options_json: JSON.stringify(['onsite']), location_note: 'สำนักงานสภาวิชาชีพสังคมสงเคราะห์ (ห้องสัมภาษณ์ 1-2)', status: 'ปิดรับสมัคร', note: 'รอบตัวอย่างสำหรับทดสอบระบบจัดการสอบสัมภาษณ์', created_by: 'demo_staff', created_at: addDaysIso_(nowIso_(), -45), updated_at: nowIso_() };
  appendRow_('ExamRounds', examRoundInterviewScoring);
  const ivNames = ['นางสาว พิมลวรรณ ทดสอบสัมภาษณ์', 'นาย ธนากร ทดสอบสัมภาษณ์', 'นางสาว สุภาพร ทดสอบสัมภาษณ์', 'นาย วรวุฒิ ทดสอบสัมภาษณ์', 'นางสาว กนกพร ทดสอบสัมภาษณ์', 'นาย ชัยวัฒน์ ทดสอบสัมภาษณ์'];
  const ivAppIds = ivNames.map((nm, i) => {
    const appId = newId_('EXA');
    appendRow_('Applications', {
      application_id: appId, applicant_type: 'สมัครสอบสัมภาษณ์', applicant_user_id: '',
      ref_data_json: JSON.stringify({ national_id: '110000000090' + (i + 1), round_id: examRoundInterviewScoring.round_id, round_label: examRoundInterviewScoring.round_label, format: 'onsite', email: 'interview' + (i + 1) + '.demo@example.com', name: nm, essay_category: EXAM_INTERVIEW_ESSAY_CATEGORIES[0], candidate_no: String(i + 1).padStart(4, '0') }),
      documents_json: '[]', status: 'รอผลสอบ', reject_reason: '', fee_amount: 0,
      submitted_at: addDaysIso_(nowIso_(), -30), reviewed_by: 'demo_staff', reviewed_at: addDaysIso_(nowIso_(), -20), approved_member_id: ''
    });
    return appId;
  });
  appendRow_('InterviewSessions', { session_id: newId_('IVS'), round_id: examRoundInterviewScoring.round_id, time_text: '09.30 - 12.00 น.', pass_score: INTERVIEW_DEFAULT_PASS_SCORE, scoring_closed: false, results_applied_at: '', note: 'ข้อมูลตัวอย่าง', created_by: 'demo_staff', created_at: nowIso_(), updated_at: nowIso_() });
  const ivGroup1 = { group_id: newId_('IVG'), round_id: examRoundInterviewScoring.round_id, group_no: 1, time_range: '09.30 - 11.00 น.', inv1_name: 'ผศ.ดร.สมศรี เมตตาธรรม', inv2_name: 'ดร.ประเสริฐ พัฒนาสังคม', password_hash: hashPassword_('Demo1234'), created_by: 'demo_staff', created_at: nowIso_(), updated_at: nowIso_() };
  const ivGroup2 = { group_id: newId_('IVG'), round_id: examRoundInterviewScoring.round_id, group_no: 2, time_range: '09.30 - 11.00 น.', inv1_name: 'รศ.อรุณี ใจเย็น', inv2_name: 'นายวิชัย ชุมชนสุข', password_hash: '', created_by: 'demo_staff', created_at: nowIso_(), updated_at: nowIso_() };
  appendRow_('InterviewGroups', ivGroup1);
  appendRow_('InterviewGroups', ivGroup2);
  const ivScore = (a, b, c, d) => ({ skill: a, attitude: b, writing: c, presentation: d, total: a + b + c + d, saved_at: nowIso_() });
  const ivRows = [
    { g: ivGroup1, ord: 1, time: '09.30 - 10.00', scores: { inv1: ivScore(25, 24, 17, 16), inv2: ivScore(24, 22, 16, 15) }, remarks: { inv1: 'สื่อสารชัดเจน วิเคราะห์กรณีได้ดี', inv2: '' } }, // เฉลี่ย 79.5 ผ่าน
    { g: ivGroup1, ord: 2, time: '10.00 - 10.30', scores: { inv1: ivScore(12, 13, 10, 9), inv2: ivScore(14, 12, 11, 10) }, remarks: { inv1: 'ควรพัฒนาทักษะการเยี่ยมบ้านและการเขียนรายงาน', inv2: '' } }, // เฉลี่ย 45.5 ไม่ผ่าน
    { g: ivGroup1, ord: 3, time: '10.30 - 11.00', scores: { inv1: ivScore(20, 22, 15, 14) }, remarks: { inv1: '' } }, // กรรมการ 1 ให้คะแนนแล้ว รอกรรมการ 2
    { g: ivGroup2, ord: 1, time: '09.30 - 10.00', scores: {}, remarks: {} },
    { g: ivGroup2, ord: 2, time: '10.00 - 10.30', scores: {}, remarks: {} },
    { g: ivGroup2, ord: 3, time: '10.30 - 11.00', scores: {}, remarks: {}, absent: true } // ไม่เข้าสอบ
  ];
  ivRows.forEach((r, i) => {
    appendRow_('InterviewCandidates', {
      ic_id: newId_('IVC'), round_id: examRoundInterviewScoring.round_id, group_id: r.g.group_id, application_id: ivAppIds[i],
      candidate_no: String(i + 1).padStart(4, '0'), candidate_name: ivNames[i], order_no: r.ord, time_slot: r.time, absent: r.absent === true,
      scores_json: JSON.stringify(r.scores), remarks_json: JSON.stringify(r.remarks),
      created_by: 'demo_staff', created_at: nowIso_(), updated_at: nowIso_(), updated_by: 'demo_staff'
    });
  });

  // ========== ชุดเดโมสอบเต็มกระบวนการ (v17) — ครอบคลุมทุกสถานะของสอบข้อเขียน/สัมภาษณ์ ใช้บัญชีสมาชิกจริงเพื่อให้ล็อกอินดูได้ ==========
  // demo_member1 และ demo_member13 ไม่มีใบสมัครสอบ — เก็บไว้สำหรับสาธิตสมัครสอบสดตั้งแต่ต้น
  (function seedExamFullFlow_() {
    const DOC = 'https://example.com/demo-doc-placeholder.pdf', SLIP = 'https://example.com/demo-slip-placeholder.jpg';
    const stdDocs = JSON.stringify([
      { name: 'id.pdf', doc_type: 'สำเนาบัตรประชาชน', url: DOC }, { name: 'photo.jpg', doc_type: 'รูปถ่าย', url: DOC }, { name: 'degree.pdf', doc_type: 'สำเนาวุฒิการศึกษา', url: DOC }
    ]);
    const essayDocs = JSON.stringify([
      { name: 'id.pdf', doc_type: 'สำเนาบัตรประชาชน', url: DOC }, { name: 'photo.jpg', doc_type: 'รูปถ่าย', url: DOC },
      { name: 'essay.pdf', doc_type: examEssayDocTypes_(EXAM_INTERVIEW_ESSAY_CATEGORIES[0])[0], url: DOC },
      { name: 'essay.docx', doc_type: examEssayDocTypes_(EXAM_INTERVIEW_ESSAY_CATEGORIES[0])[1], url: DOC }
    ]);
    const S0 = EXAM_SUBJECTS[0], S1 = EXAM_SUBJECTS[1];
    const whoIs = nid => {
      const d = demoMembers.find(x => x.national_id === nid);
      return { memberId: demoMemberIds[nid], userId: demoUserIds[nid], email: d ? d.email : '', prefix: d ? d.prefix_th : '', first: d ? d.first_name_th : '' };
    };
    // เพิ่มใบสมัครสอบ 1 ใบ (+ รายการชำระเงิน/คืนเงินถ้าระบุ) — opts: nid, type, round, name, status, candidate_no, pay ('paid'|'slip'|'unpaid'|null), exam_result, reject_reason, extra, submittedDays
    const addExamApp = o => {
      const w = whoIs(o.nid);
      if (!w.memberId || !w.userId) return null;
      const written = o.type === 'สมัครสอบข้อเขียน';
      const data = Object.assign({ member_id: w.memberId, round_id: o.round.round_id, round_label: o.round.round_label, format: o.format || 'onsite', email: w.email,
        name: w.prefix + ' ' + w.first + ' ' + o.tag }, written ? { subjects: o.subjects || EXAM_SUBJECTS.slice() } : { essay_category: EXAM_INTERVIEW_ESSAY_CATEGORIES[0] }, o.extra || {});
      if (o.candidate_no) data.candidate_no = o.candidate_no;
      if (o.exam_result) { data.exam_result = o.exam_result; if (o.exam_result.subject_results) data.subject_results = o.exam_result.subject_results; }
      const fee = written ? EXAM_SUBJECT_FEE * data.subjects.length : 0;
      const appId = newId_('EXA');
      const sd = o.submittedDays !== undefined ? o.submittedDays : -30;
      appendRow_('Applications', {
        application_id: appId, applicant_type: o.type, applicant_user_id: w.userId, ref_data_json: JSON.stringify(data),
        documents_json: written ? stdDocs : essayDocs, status: o.status, reject_reason: o.reject_reason || '', fee_amount: fee,
        submitted_at: addDaysIso_(nowIso_(), sd), reviewed_by: o.status === 'รอตรวจสอบ' ? '' : 'demo_staff', reviewed_at: o.status === 'รอตรวจสอบ' ? '' : addDaysIso_(nowIso_(), sd + 5), approved_member_id: ''
      });
      let payId = null;
      if (written && o.pay) {
        payId = newId_('PAY');
        const paid = o.pay === 'paid';
        appendRow_('Payments', {
          payment_id: payId, ref_type: 'application', ref_id: appId, member_id: w.memberId, amount: fee,
          slip_url: o.pay === 'unpaid' ? '' : SLIP, method: o.pay === 'unpaid' ? '' : 'โอนเงิน', status: paid ? 'ชำระแล้ว' : (o.pay === 'slip' ? 'รอตรวจสอบ' : 'รอชำระ'),
          verified_by: paid ? 'demo_staff' : '', verified_at: paid ? addDaysIso_(nowIso_(), sd + 8) : '', paid_at: paid ? addDaysIso_(nowIso_(), sd + 8) : ''
        });
        if (paid) { try { issueReceiptForPayment_(payId, 'demo_staff'); } catch (e) { /* ข้ามได้ */ } }
      }
      return { appId: appId, payId: payId, memberId: w.memberId };
    };
    const wResult = (tp, pp, result, basis, round) => ({ kind: 'เขียน', exam_date: round.exam_date, seat_no: '', scores: { [S0]: tp, [S1]: pp },
      subject_results: { [S0]: tp >= 60 ? 'ผ่าน' : 'ไม่ผ่าน', [S1]: pp >= 60 ? 'ผ่าน' : 'ไม่ผ่าน' }, result: result, basis: basis, recorded_at: addDaysIso_(nowIso_(), -37), recorded_by: 'demo_staff' });
    const R = examRoundWrittenClosed; // 1/2569 (เสร็จสิ้น/ประกาศผลแล้ว)
    const R2 = examRoundWrittenOpen;  // 2/2569 (เปิดรับสมัคร)

    // ----- สอบข้อเขียน -----
    // (1) สอบไม่ผ่าน: เฉลี่ย 57.5 < 65 — demo_member9 (สุรชัย)
    const f1 = wResult(55, 60, 'ไม่ผ่าน', 'คะแนนเฉลี่ย 2 วิชาต่ำกว่าร้อยละ 65', R); f1.seat_no = '0003';
    addExamApp({ nid: '1100000000011', type: 'สมัครสอบข้อเขียน', round: R, tag: 'สอบข้อเขียนไม่ผ่าน', status: 'สอบไม่ผ่าน', candidate_no: '0003', pay: 'paid', exam_result: f1, submittedDays: -90 });
    // (2) ผ่านด้วยเกณฑ์เฉลี่ย (ทฤษฎี 52 / ปฏิบัติ 80 เฉลี่ย 66) — demo_member10 (อรวรรณ)
    const f2 = wResult(52, 80, 'ผ่าน', 'คะแนนเฉลี่ย 2 วิชาไม่น้อยกว่าร้อยละ 65 และแต่ละวิชาไม่น้อยกว่าร้อยละ 50', R); f2.seat_no = '0004';
    addExamApp({ nid: '1100000000012', type: 'สมัครสอบข้อเขียน', round: R, tag: 'สอบผ่านด้วยเกณฑ์เฉลี่ย', status: 'สอบผ่าน', candidate_no: '0004', pay: 'paid', exam_result: f2, submittedDays: -90 });
    // (3)(4) รอผลสอบ รอบ 1/2569 เพิ่มอีก 2 คน — ใช้สาธิตนำเข้าผลหลายคนพร้อมกับธีระ (0001)
    addExamApp({ nid: '1100000000013', type: 'สมัครสอบข้อเขียน', round: R, tag: 'รอผลสอบข้อเขียน', status: 'รอผลสอบ', candidate_no: '0005', pay: 'paid', submittedDays: -88 });
    addExamApp({ nid: '1100000000014', type: 'สมัครสอบข้อเขียน', round: R, tag: 'รอผลสอบข้อเขียน (วิชาเดียว)', status: 'รอผลสอบ', candidate_no: '0006', pay: 'paid', subjects: [S0], submittedDays: -87 });
    // (5) ไม่อนุมัติ — demo_member15 (วิสุทธิ์) รอบ 2/2569
    addExamApp({ nid: '1100000000017', type: 'สมัครสอบข้อเขียน', round: R2, tag: 'ใบสมัครสอบข้อเขียนไม่อนุมัติ', status: 'ไม่อนุมัติ', reject_reason: 'สำเนาวุฒิการศึกษาไม่ชัดเจน กรุณาแนบใหม่แล้วสมัครอีกครั้ง', submittedDays: -6 });
    // (6) แนบสลิปแล้ว รอแอดมินตรวจสลิป — demo_member17 (อำไพ) รอบ 2/2569
    addExamApp({ nid: '1100000000020', type: 'สมัครสอบข้อเขียน', round: R2, tag: 'รอตรวจสลิปสอบข้อเขียน', status: 'รอชำระเงิน', pay: 'slip', submittedDays: -4 });
    // (7) ขอคืนเงินแล้ว (กรณีพิเศษ รอตรวจสอบ) — demo_member18 (ศิริพร) รอบ 1/2569
    const refundApp = addExamApp({ nid: '1100000000021', type: 'สมัครสอบข้อเขียน', round: R, tag: 'ขอคืนเงินค่าสมัครสอบข้อเขียน', status: 'รอผลสอบ', candidate_no: '0007', pay: 'paid', submittedDays: -85 });
    if (refundApp && refundApp.payId) {
      appendRow_('Refunds', {
        refund_id: newId_('RFD'), payment_id: refundApp.payId, member_id: refundApp.memberId, reason: '[คืนเงินค่าสมัครสอบข้อเขียน ใบสมัคร ' + refundApp.appId + '] ผู้สมัครป่วยหนัก ตามหนังสืออนุมัติสภาฯ เลขที่ 12/2569 (ข้อมูลเดโม)',
        requested_by: 'demo_staff', requested_at: nowIso_(), status: 'รอตรวจสอบ', approved_by: '', approved_at: '', refund_amount: EXAM_SUBJECT_FEE * 2, refund_note: '', refunded_at: ''
      });
    }
    // (8) เปลี่ยนรอบแล้ว (ประวัติเปลี่ยนรอบ 1/2569 → 2/2569) — demo_member19 (รัตนา)
    addExamApp({ nid: '1100000000022', type: 'สมัครสอบข้อเขียน', round: R2, tag: 'เปลี่ยนรอบสอบข้อเขียน', status: 'รอผลสอบ', candidate_no: '0001', pay: 'paid', submittedDays: -60,
      extra: { round_changes: [{ from_round_id: R.round_id, from_round_label: R.round_label, from_candidate_no: '0008', to_round_id: R2.round_id, to_round_label: R2.round_label, reason: 'ติดภารกิจราชการ ตามหนังสืออนุมัติสภาฯ เลขที่ 15/2569 (ข้อมูลเดโม)', changed_by: 'demo_staff', changed_at: addDaysIso_(nowIso_(), -20) }] } });

    // ----- สอบสัมภาษณ์ -----
    const RI = examRoundInterviewClosed, RI3 = examRoundInterviewOpen; // 2/2569 (เสร็จสิ้น) , 3/2569 (เปิดรับสมัคร)
    // (9) ประกาศผลแล้ว: ผ่าน (เฉลี่ย 82.5) — demo_member2 (สมหญิง)
    addExamApp({ nid: '1100000000002', type: 'สมัครสอบสัมภาษณ์', round: RI, tag: 'สอบสัมภาษณ์ผ่าน', status: 'สอบผ่าน', candidate_no: '0001', submittedDays: -90,
      exam_result: { kind: 'สัมภาษณ์', exam_date: RI.exam_date, seat_no: '0001', score: 82.5, result: 'ผ่าน', note: 'คะแนนเฉลี่ยกรรมการ 2 ท่าน (ก.1 83 / ก.2 82)', recorded_at: addDaysIso_(nowIso_(), -36), recorded_by: 'demo_staff' } });
    // (10) ประกาศผลแล้ว: ไม่ผ่าน (เฉลี่ย 48) — demo_member3 (ประเสริฐ)
    addExamApp({ nid: '1100000000005', type: 'สมัครสอบสัมภาษณ์', round: RI, tag: 'สอบสัมภาษณ์ไม่ผ่าน', status: 'สอบไม่ผ่าน', candidate_no: '0002', submittedDays: -90,
      exam_result: { kind: 'สัมภาษณ์', exam_date: RI.exam_date, seat_no: '0002', score: 48, result: 'ไม่ผ่าน', note: 'คะแนนเฉลี่ยกรรมการ 2 ท่าน (ก.1 47 / ก.2 49)', recorded_at: addDaysIso_(nowIso_(), -36), recorded_by: 'demo_staff' } });
    // (11) รอตรวจสอบเอกสาร — demo_member4 (มาลี) รอบ 3/2569 ใช้สาธิตอนุมัติ
    addExamApp({ nid: '1100000000006', type: 'สมัครสอบสัมภาษณ์', round: RI3, tag: 'สมัครสอบสัมภาษณ์ รอตรวจสอบ', status: 'รอตรวจสอบ', submittedDays: -2 });
    // (12) อนุมัติแล้ว รอผลสอบ (ได้รหัสสอบ) — demo_member16 (จันทร์เพ็ญ) รอบ 3/2569 ใช้สาธิตจัดกลุ่ม/ให้คะแนนในหน้าจัดการสอบสัมภาษณ์
    addExamApp({ nid: '1100000000019', type: 'สมัครสอบสัมภาษณ์', round: RI3, tag: 'สอบสัมภาษณ์ รอผลสอบ', status: 'รอผลสอบ', candidate_no: '0001', submittedDays: -8 });
    addExamApp({ nid: '1100000000020', type: 'สมัครสอบสัมภาษณ์', round: RI3, tag: 'สอบสัมภาษณ์ รอผลสอบ', status: 'รอผลสอบ', candidate_no: '0002', submittedDays: -8 });
  })();

  // ========== สมาชิกกิตติมศักดิ์ (ตัวอย่าง) ==========
  [
    { prefix_th: 'ศาสตราจารย์ ดร.', first_name_th: 'สมชาย', last_name_th: 'ใจดีงาม', qualification_type: 'ผู้ทรงคุณวุฒิ ผู้รอบรู้ด้านสังคมสงเคราะห์ศาสตร์', expertise: 'นโยบายสวัสดิการสังคม', position: 'ศาสตราจารย์เกียรติคุณ', organization: 'มหาวิทยาลัยธรรมศาสตร์', status: 'ดำรงสถานะ' },
    { prefix_th: 'นาง', first_name_th: 'สุดา', last_name_th: 'เมตตาธรรม', qualification_type: 'ผู้ที่มีผลงานเชิงประจักษ์ด้านสังคมสงเคราะห์ศาสตร์', expertise: 'งานพัฒนาชุมชนและการช่วยเหลือผู้สูงอายุ', position: 'ประธานมูลนิธิ', organization: 'มูลนิธิเพื่อผู้สูงวัย', status: 'ดำรงสถานะ' },
    { prefix_th: 'นาย', first_name_th: 'วิชัย', last_name_th: 'พัฒนาชน', qualification_type: 'ผู้ที่มีผลงานเชิงประจักษ์ด้านสังคมสงเคราะห์ศาสตร์', expertise: 'กฎหมายคุ้มครองเด็ก', position: 'ที่ปรึกษากฎหมาย', organization: 'สำนักงานกฎหมายเพื่อสังคม', status: 'สิ้นสุดสถานะ' }
  ].forEach((h, hi) => {
    const hRow = Object.assign({
      honorary_id: newId_('HNR'), honorary_no: nextHonoraryNo_(), first_name_en: '', last_name_en: '', phone: '', email: '', photo_url: '',
      appointed_date: addDaysIso_(nowIso_(), -400).slice(0, 10), term_end: '', note: 'ข้อมูลตัวอย่าง',
      created_by: 'demo_staff', created_at: nowIso_(), updated_at: nowIso_()
    }, h);
    if (hi === 2) hRow.term_end = addDaysIso_(nowIso_(), -30).slice(0, 10);
    appendRow_('MembersHonorary', hRow);
    // บัญชีเดโมสมาชิกกิตติมศักดิ์: demo_honorary1/2 ใช้งานได้, demo_honorary3 สิ้นสุดสถานะ (บัญชีถูกปิด)
    appendRow_('Users', {
      user_id: newId_('USR'), username: 'demo_honorary' + (hi + 1), password_hash: hashPassword_('Demo1234'),
      role: 'member', member_id: hRow.honorary_id, thaid_pid: '', status: h.status === 'ดำรงสถานะ' ? 'active' : 'inactive',
      created_at: nowIso_(), last_login: '', failed_login_attempts: 0, locked_until: '',
      display_name: [h.prefix_th, h.first_name_th, h.last_name_th].filter(Boolean).join(' ')
    });
  });

  // ========== ครอบคลุมทุกขั้นของ "ต่ออายุใบอนุญาต" (TOR ข้อ 4.2) — ตั้งแต่ก่อนยื่นจนต่ออายุสำเร็จ ==========
  const member11Id = demoMemberIds['1100000000011']; // สุรชัย — ก่อนต่ออายุ (มีใบอนุญาตในช่วง 180 วัน ไม่มีคำขอเลย ไม่ต้องสร้างอะไรเพิ่ม)
  const member12Id = demoMemberIds['1100000000012']; // อรวรรณ — แนบสลิปแล้ว รอแอดมินตรวจสอบ
  const user12Id = demoUserIds['1100000000012'];
  const member13Id = demoMemberIds['1100000000013']; // ณรงค์ — อนุมัติแล้ว รอชำระเงิน (หลังหมดอายุ)
  const user13Id = demoUserIds['1100000000013'];
  const member14Id = demoMemberIds['1100000000014']; // เบญจมาศ — ต่ออายุสำเร็จ (ได้เลขใหม่)

  if (member12Id && user12Id) {
    const appId = newId_('REN');
    appendRow_('Applications', {
      application_id: appId, applicant_type: 'ต่ออายุใบอนุญาต', applicant_user_id: user12Id,
      ref_data_json: JSON.stringify({ member_id: member12Id, method: 'CE_before', isAfterExpiry: false, email: 'orawan.demo@example.com', name: 'นางสาว อรวรรณ รอตรวจสลิปต่ออายุ' }),
      documents_json: '[]', status: 'รอชำระเงิน', reject_reason: '', fee_amount: LICENSE_RENEWAL_FEE,
      submitted_at: nowIso_(), reviewed_by: 'ระบบ (ผ่าน CE อัตโนมัติ)', reviewed_at: nowIso_(), approved_member_id: ''
    });
    appendRow_('Payments', {
      payment_id: newId_('PAY'), ref_type: 'application', ref_id: appId, member_id: member12Id,
      amount: LICENSE_RENEWAL_FEE, slip_url: 'https://example.com/demo-slip-placeholder.jpg',
      method: 'โอนเงิน', status: 'รอตรวจสอบ', verified_by: '', verified_at: '', paid_at: ''
    });
  }

  if (member13Id && user13Id) {
    const appId = newId_('REN');
    appendRow_('Applications', {
      application_id: appId, applicant_type: 'ต่ออายุใบอนุญาต', applicant_user_id: user13Id,
      ref_data_json: JSON.stringify({ member_id: member13Id, method: 'exam', isAfterExpiry: true, email: 'narong.demo@example.com', name: 'นาย ณรงค์ รอชำระค่าต่ออายุ' }),
      documents_json: '[]', status: 'รอชำระเงิน', reject_reason: '', fee_amount: LICENSE_REGISTRATION_FEE,
      submitted_at: addYearsIso_(nowIso_(), 0), reviewed_by: 'demo_staff', reviewed_at: nowIso_(), approved_member_id: ''
    });
    appendRow_('Payments', {
      payment_id: newId_('PAY'), ref_type: 'application', ref_id: appId, member_id: member13Id,
      amount: LICENSE_REGISTRATION_FEE, slip_url: '', method: '', status: 'รอชำระ', verified_by: '', verified_at: '', paid_at: ''
    });
  }

  if (member14Id) {
    const appId = newId_('REN');
    const payId = newId_('PAY');
    const oldLicenseNo = 'LC2567-DEMO14-OLD';
    const newLicenseNo = 'LC' + (new Date().getFullYear() + 543) + '-DEMO14';
    appendRow_('Applications', {
      application_id: appId, applicant_type: 'ต่ออายุใบอนุญาต', applicant_user_id: demoUserIds['1100000000014'] || '',
      ref_data_json: JSON.stringify({ member_id: member14Id, method: 'CE_after', isAfterExpiry: true, email: 'benjamas.demo@example.com', name: 'นางสาว เบญจมาศ ต่ออายุสำเร็จ' }),
      documents_json: '[]', status: 'ต่ออายุสำเร็จ', reject_reason: '', fee_amount: LICENSE_RENEWAL_FEE,
      submitted_at: addYearsIso_(nowIso_(), 0), reviewed_by: 'demo_staff', reviewed_at: nowIso_(), approved_member_id: member14Id
    });
    appendRow_('Payments', {
      payment_id: payId, ref_type: 'application', ref_id: appId, member_id: member14Id,
      amount: LICENSE_RENEWAL_FEE, slip_url: 'https://example.com/demo-slip-placeholder.jpg',
      method: 'โอนเงิน', channel: 'Mobile Banking', status: 'ชำระแล้ว', verified_by: 'demo_staff', verified_at: nowIso_(), paid_at: nowIso_()
    });
    try { issueReceiptForPayment_(payId, 'demo_staff'); } catch (e) { /* ข้ามถ้าออกใบเสร็จเดโมไม่สำเร็จ ไม่กระทบการ seed ข้อมูลหลัก */ }
    appendRow_('VoidedLicenses', {
      void_id: newId_('VOID'), old_license_no: oldLicenseNo, new_license_no: newLicenseNo,
      member_id: member14Id, member_name: 'นางสาว เบญจมาศ ต่ออายุสำเร็จ',
      void_date: nowIso_(), reason: 'ต่ออายุหลังใบอนุญาตหมดอายุ (ถือเป็นการขึ้นทะเบียนใหม่) — ข้อมูลเดโม', actor_user_id: 'demo_staff'
    });
    appendRow_('LicenseRenewalHistory', {
      renewal_id: newId_('LR'), license_id: 'LIC-DEMO14', member_id: member14Id, method: 'CE_after',
      old_expire_date: addYearsIso_(nowIso_(), -0.05), new_expire_date: addYearsIso_(nowIso_(), 2),
      renewed_by: 'demo_staff', renewed_at: nowIso_(), note: 'ต่อหลังหมดอายุ — ออกเลขใหม่ ' + newLicenseNo + ' (ข้อมูลเดโม)'
    });
  }

  // ---------- เปลี่ยนประเภทสมาชิก (สมทบ → สามัญ) — TOR ข้อ 1.1(7) ----------
  const member16Id = demoMemberIds['1100000000016']; // พิมพ์ใจ — ยื่นเปลี่ยนประเภท รอตรวจสอบ
  const member17Id = demoMemberIds['1100000000017']; // วิสุทธิ์ — เปลี่ยนประเภทสำเร็จแล้ว (ประวัติ)
  const user16Id = demoUserIds['1100000000016'];
  const user17Id = demoUserIds['1100000000017'];

  if (member16Id && user16Id) {
    const tcAppId = newId_('TC');
    appendRow_('Applications', {
      application_id: tcAppId, applicant_type: 'เปลี่ยนประเภท', applicant_user_id: user16Id,
      ref_data_json: JSON.stringify({
        member_id: member16Id, from_type: 'สมทบ', to_type: 'สามัญ', duration: '5',
        email: 'pimjai.demo@example.com', name: 'นางสาว พิมพ์ใจ รอเปลี่ยนประเภท'
      }),
      documents_json: '[]', status: 'รอตรวจสอบ', reject_reason: '',
      fee_amount: MEMBERSHIP_FEE_TABLE['สามัญ']['5'],
      submitted_at: nowIso_(), reviewed_by: '', reviewed_at: '', approved_member_id: ''
    });
  }

  if (member17Id && user17Id) {
    const tcAppId2 = newId_('TC');
    const payId2 = newId_('PAY');
    appendRow_('Applications', {
      application_id: tcAppId2, applicant_type: 'เปลี่ยนประเภท', applicant_user_id: user17Id,
      ref_data_json: JSON.stringify({
        member_id: member17Id, from_type: 'สมทบ', to_type: 'สามัญ', duration: '5',
        email: 'wisut.demo@example.com', name: 'นาย วิสุทธิ์ เปลี่ยนประเภทสำเร็จ'
      }),
      documents_json: '[]', status: 'เปลี่ยนประเภทสำเร็จ', reject_reason: '',
      fee_amount: MEMBERSHIP_FEE_TABLE['สามัญ']['5'],
      submitted_at: addYearsIso_(nowIso_(), -0.1), reviewed_by: 'demo_staff', reviewed_at: addYearsIso_(nowIso_(), -0.08),
      approved_member_id: member17Id
    });
    appendRow_('Payments', {
      payment_id: payId2, ref_type: 'application', ref_id: tcAppId2, member_id: member17Id,
      amount: MEMBERSHIP_FEE_TABLE['สามัญ']['5'], slip_url: 'https://example.com/demo-slip-placeholder.jpg',
      method: 'โอนเงิน', channel: 'โอนเงินผ่านธนาคาร (ตรวจสอบสลิปด้วยมือ)', status: 'ชำระแล้ว',
      verified_by: 'demo_staff', verified_at: addYearsIso_(nowIso_(), -0.08), paid_at: addYearsIso_(nowIso_(), -0.08)
    });
    try { issueReceiptForPayment_(payId2, 'demo_staff'); } catch (e) { /* ข้ามถ้าออกใบเสร็จเดโมไม่สำเร็จ ไม่กระทบการ seed ข้อมูลหลัก */ }
  }

  // ---------- ระบบขอเปลี่ยนชื่อ-นามสกุล (ภาษาไทย) — ครอบคลุมทุกสถานะ ----------
  const member19Id = demoMemberIds['1100000000019']; // จันทร์เพ็ญ — รอตรวจสอบ
  const member20Id = demoMemberIds['1100000000020']; // อำไพ — รอเอกสารเพิ่มเติม
  const member21Id = demoMemberIds['1100000000021']; // ศิริพร — ไม่อนุมัติ
  const user19Id = demoUserIds['1100000000019'];
  const user20Id = demoUserIds['1100000000020'];
  const user21Id = demoUserIds['1100000000021'];

  if (member19Id && user19Id) {
    appendRow_('Applications', {
      application_id: newId_('NAMECHG'), applicant_type: 'name_change', applicant_user_id: user19Id,
      ref_data_json: JSON.stringify({
        member_id: member19Id,
        patch: { prefix_th: 'นาง', first_name_th: 'จันทร์เพ็ญ', last_name_th: 'สมรสใหม่' },
        old_name: { prefix_th: 'นางสาว', first_name_th: 'จันทร์เพ็ญ', last_name_th: 'รอเปลี่ยนชื่อ' },
        reason: 'สมรส'
      }),
      documents_json: JSON.stringify([{ name: 'ใบสำคัญการสมรส.pdf', doc_type: 'หลักฐานประกอบการเปลี่ยนชื่อ-นามสกุล', url: 'https://example.com/demo-doc-placeholder.pdf' }]),
      status: 'รอตรวจสอบ', reject_reason: '', fee_amount: 0,
      submitted_at: nowIso_(), reviewed_by: '', reviewed_at: '', approved_member_id: ''
    });
  }

  if (member20Id && user20Id) {
    appendRow_('Applications', {
      application_id: newId_('NAMECHG'), applicant_type: 'name_change', applicant_user_id: user20Id,
      ref_data_json: JSON.stringify({
        member_id: member20Id,
        patch: { prefix_th: 'นาง', first_name_th: 'อำไพ', last_name_th: 'เปลี่ยนนามสกุลใหม่' },
        old_name: { prefix_th: 'นาง', first_name_th: 'อำไพ', last_name_th: 'รอเอกสารเพิ่มเติม' },
        reason: 'เปลี่ยนชื่อตามกฎหมาย'
      }),
      documents_json: JSON.stringify([{ name: 'คำร้องขอเปลี่ยนชื่อ.jpg', doc_type: 'หลักฐานประกอบการเปลี่ยนชื่อ-นามสกุล', url: 'https://example.com/demo-doc-placeholder.jpg' }]),
      status: 'รอเอกสารเพิ่มเติม',
      reject_reason: 'เอกสารที่แนบมาเป็นเพียงคำร้องขอ ยังไม่ใช่ใบสำคัญที่ทางราชการออกให้ กรุณาแนบใบสำคัญการเปลี่ยนชื่อ-นามสกุลฉบับที่นายทะเบียนออกให้ (ตัวจริงหรือสำเนารับรอง)',
      fee_amount: 0, submitted_at: addYearsIso_(nowIso_(), -0.02), reviewed_by: 'demo_staff', reviewed_at: nowIso_(), approved_member_id: ''
    });
  }

  if (member21Id && user21Id) {
    appendRow_('Applications', {
      application_id: newId_('NAMECHG'), applicant_type: 'name_change', applicant_user_id: user21Id,
      ref_data_json: JSON.stringify({
        member_id: member21Id,
        patch: { prefix_th: 'นางสาว', first_name_th: 'ศิริพร', last_name_th: 'ชื่อใหม่ตามใจชอบ' },
        old_name: { prefix_th: 'นางสาว', first_name_th: 'ศิริพร', last_name_th: 'ถูกไม่อนุมัติ' },
        reason: 'อื่นๆ'
      }),
      documents_json: JSON.stringify([{ name: 'หนังสือชี้แจง.pdf', doc_type: 'หลักฐานประกอบการเปลี่ยนชื่อ-นามสกุล', url: 'https://example.com/demo-doc-placeholder.pdf' }]),
      status: 'ไม่อนุมัติ',
      reject_reason: 'เอกสารที่แนบมาไม่ใช่หลักฐานทางราชการที่รับรองการเปลี่ยนชื่อ-นามสกุล (ระบุเหตุผลว่า "เปลี่ยนตามใจชอบ" โดยไม่มีใบสำคัญรับรองจากนายทะเบียน) ไม่เข้าเงื่อนไขตามระเบียบ',
      fee_amount: 0, submitted_at: addYearsIso_(nowIso_(), -0.05), reviewed_by: 'demo_staff', reviewed_at: addYearsIso_(nowIso_(), -0.03), approved_member_id: ''
    });
  }

  // สมาชิกที่คำขอเปลี่ยนชื่อ "อนุมัติแล้ว" — ชื่อปัจจุบันคือชื่อใหม่ที่อนุมัติผ่านระบบจริง พร้อมสร้าง AuditLog จริง (ไม่ใช่แค่ appendRow_ ตรงๆ)
  // เพื่อให้หน้า "ประวัติการเปลี่ยนแปลง" ทั้งฝั่งสมาชิกและแอดมินมีตัวอย่างประวัติที่ผ่านการอนุมัติแล้วให้ดูจริง
  const member22Id = demoMemberIds['1100000000022'];
  const user22Id = demoUserIds['1100000000022'];
  if (member22Id && user22Id) {
    const oldNameData = { prefix_th: 'นางสาว', first_name_th: 'รัตนา', last_name_th: 'ก่อนเปลี่ยนชื่อ' };
    const newNamePatch = { prefix_th: 'นาง', first_name_th: 'รัตนา', last_name_th: 'เปลี่ยนชื่อสำเร็จแล้ว' };
    const appId22 = newId_('NAMECHG');
    appendRow_('Applications', {
      application_id: appId22, applicant_type: 'name_change', applicant_user_id: user22Id,
      ref_data_json: JSON.stringify({ member_id: member22Id, patch: newNamePatch, old_name: oldNameData, reason: 'สมรส' }),
      documents_json: JSON.stringify([{ name: 'ใบสำคัญการสมรส.pdf', doc_type: 'หลักฐานประกอบการเปลี่ยนชื่อ-นามสกุล', url: 'https://example.com/demo-doc-placeholder.pdf' }]),
      status: 'สำเร็จ', reject_reason: '', fee_amount: 0,
      submitted_at: addYearsIso_(nowIso_(), -0.1), reviewed_by: 'demo_staff', reviewed_at: addYearsIso_(nowIso_(), -0.08), approved_member_id: ''
    });
    // จำลองประวัติ "ก่อนเปลี่ยน" ไว้ใน AuditLog เพื่อให้หน้า History แสดงชื่อเดิม → ชื่อใหม่ ให้เห็นจริง (สมาชิกในชีตตอนนี้ใช้ชื่อใหม่แล้วตามที่ appendRow_ ตอนสร้างสมาชิกด้านบน)
    auditLog_('Members', member22Id, 'name_change_approved', oldNameData, newNamePatch, 'demo_staff');
  }

  // 4.4.1-4.4.2: กลุ่มนำเข้ารายชื่อผู้มีสิทธิ์ต่ออายุ (ตัวอย่างไว้ทดสอบ "อนุมัติแบบกลุ่ม")
  if (member1Id) {
    appendRow_('RenewalEligibility', {
      eligibility_id: newId_('ELIG'), member_id: member1Id, import_type: 'before_expiry', source: 'CSV',
      batch_name: 'รอบทดสอบเดโม', imported_at: nowIso_(), imported_by: 'demo_staff',
      status: 'รอตรวจสอบ', decided_by: '', decided_at: ''
    });
  }
  if (member6Id) {
    appendRow_('RenewalEligibility', {
      eligibility_id: newId_('ELIG'), member_id: member6Id, import_type: 'after_expiry', source: 'CSV',
      batch_name: 'รอบทดสอบเดโม', imported_at: nowIso_(), imported_by: 'demo_staff',
      status: 'รอตรวจสอบ', decided_by: '', decided_at: ''
    });
  }

  // 4.4.4: ตัวอย่างประวัติการยกเลิกเลขที่ใบอนุญาต (เหตุการณ์สมมติในอดีต ไว้ให้ตารางไม่ว่างตั้งแต่แรก)
  appendRow_('VoidedLicenses', {
    void_id: newId_('VOID'), old_license_no: 'LC2567-DEMO99', new_license_no: 'LC2568-DEMO98',
    member_id: member1Id || '', member_name: 'นาย สมชาย ใจดี (ข้อมูลเดโมย้อนหลัง)',
    void_date: addYearsIso_(nowIso_(), -1), reason: 'ต่ออายุหลังใบอนุญาตหมดอายุ (ข้อมูลเดโมย้อนหลัง)', actor_user_id: 'demo_staff'
  });

  // ประวัติการต่ออายุใบอนุญาตตัวอย่าง (ผูกกับ demo_member1)
  appendRow_('LicenseRenewalHistory', {
    renewal_id: newId_('LR'), license_id: 'LIC-DEMO-HIST', member_id: member1Id || '', method: 'CE_before',
    old_expire_date: addYearsIso_(nowIso_(), -1), new_expire_date: addYearsIso_(nowIso_(), 1),
    renewed_by: 'demo_staff', renewed_at: addYearsIso_(nowIso_(), -1), note: 'ต่อก่อนหมดอายุ — คงเลขเดิม (ข้อมูลเดโม)'
  });

  // ผู้สมัครเดโม 3 คน — ครอบคลุมทั้ง 3 สถานะสำคัญของ workflow สมัครสมาชิก (TOR ข้อ 2.2)
  function makeApplicant(username) {
    const userId = newId_('USR');
    appendRow_('Users', {
      user_id: userId, username: username, password_hash: hashPassword_('Demo1234'),
      role: 'member', member_id: '', thaid_pid: '', status: 'active',
      created_at: nowIso_(), last_login: ''
    });
    return userId;
  }

  // เดโม 1: ใบสมัครสมาชิกสามัญ สถานะ "รอตรวจสอบ" — ทดสอบหน้าคำขอสมัครสมาชิกฝั่งแอดมิน (อนุมัติ/ไม่อนุมัติ)
  const applicant1 = makeApplicant('demo_applicant1');
  appendRow_('Applications', {
    application_id: newId_('APP'), applicant_type: 'สามัญ', applicant_user_id: applicant1,
    ref_data_json: JSON.stringify({
      prefix_th: 'นางสาว', first_name_th: 'ทดสอบ', last_name_th: 'รอตรวจสอบ',
      national_id: '1100000000099', dob: '1995-05-05', phone: '0890000099',
      email: 'test.pending@example.com', work_province: 'นนทบุรี', workplace_name: 'หน่วยงานเดโม', duration: '5'
    }),
    documents_json: '[]', status: 'รอตรวจสอบ', reject_reason: '', fee_amount: MEMBERSHIP_FEE_TABLE['สามัญ']['5'] + MEMBERSHIP_ENTRANCE_FEE,
    submitted_at: nowIso_(), reviewed_by: '', reviewed_at: '', approved_member_id: ''
  });

  // เดโม 2: ใบสมัครสมาชิกสมทบ อนุมัติแล้ว รอชำระเงิน — ทดสอบหน้าแนบสลิป (ฝั่งสมาชิก) + ตรวจสลิป (ฝั่งแอดมิน)
  const applicant2 = makeApplicant('demo_applicant2');
  const app2Id = newId_('APP');
  appendRow_('Applications', {
    application_id: app2Id, applicant_type: 'สมทบ', applicant_user_id: applicant2,
    ref_data_json: JSON.stringify({
      prefix_th: 'นาย', first_name_th: 'ทดสอบ', last_name_th: 'รอชำระเงิน',
      national_id: '1100000000098', dob: '1992-02-02', phone: '0890000098',
      email: 'test.payment@example.com', work_province: 'ชลบุรี', workplace_name: 'หน่วยงานเดโม 2', duration: '1'
    }),
    documents_json: '[]', status: 'รอชำระเงิน', reject_reason: '', fee_amount: MEMBERSHIP_FEE_TABLE['สมทบ']['1'] + MEMBERSHIP_ENTRANCE_FEE,
    submitted_at: nowIso_(), reviewed_by: 'demo_staff', reviewed_at: nowIso_(), approved_member_id: ''
  });
  appendRow_('Payments', {
    payment_id: newId_('PAY'), ref_type: 'application', ref_id: app2Id, member_id: '',
    amount: MEMBERSHIP_FEE_TABLE['สมทบ']['1'] + MEMBERSHIP_ENTRANCE_FEE, slip_url: '', method: '', status: 'รอชำระ',
    verified_by: '', verified_at: '', paid_at: ''
  });

  // เดโม 2.5: ใบสมัครสมาชิกสามัญ อนุมัติแล้ว แนบสลิปแล้ว รอเจ้าหน้าที่ตรวจสอบ — ทดสอบส่วน "ตรวจสอบการชำระเงิน (คำขอที่รอสลิปตรวจสอบ)" ในหน้าคำขอสมัครสมาชิกฝั่งแอดมินโดยเฉพาะ (ต่างจากเดโม 2 ที่ยังไม่แนบสลิป)
  const applicant4 = makeApplicant('demo_applicant3');
  const app4Id = newId_('APP');
  appendRow_('Applications', {
    application_id: app4Id, applicant_type: 'สามัญ', applicant_user_id: applicant4,
    ref_data_json: JSON.stringify({
      prefix_th: 'นาง', first_name_th: 'ทดสอบ', last_name_th: 'รอตรวจสลิปสมัครสมาชิก',
      national_id: '1100000000097', dob: '1988-08-08', phone: '0890000097',
      email: 'test.slippending@example.com', work_province: 'ภูเก็ต', workplace_name: 'หน่วยงานเดโม 3', duration: '5'
    }),
    documents_json: '[]', status: 'รอชำระเงิน', reject_reason: '', fee_amount: MEMBERSHIP_FEE_TABLE['สามัญ']['5'] + MEMBERSHIP_ENTRANCE_FEE,
    submitted_at: nowIso_(), reviewed_by: 'demo_staff', reviewed_at: nowIso_(), approved_member_id: ''
  });
  appendRow_('Payments', {
    payment_id: newId_('PAY'), ref_type: 'application', ref_id: app4Id, member_id: '',
    amount: MEMBERSHIP_FEE_TABLE['สามัญ']['5'] + MEMBERSHIP_ENTRANCE_FEE, slip_url: 'https://example.com/demo-slip-placeholder.jpg',
    method: 'โอนเงิน', channel: 'โอนเงินผ่านธนาคาร (ตรวจสอบสลิปด้วยมือ)', status: 'รอตรวจสอบ',
    verified_by: '', verified_at: '', paid_at: ''
  });

  // เดโม 3: ใบสมัครสมาชิกวิสามัญ (องค์กร) สถานะ "รอตรวจสอบ" — ทดสอบฟอร์มหลายขั้นตอนและการอนุมัติสมาชิกวิสามัญ
  const applicant3 = makeApplicant('demo_org_applicant');
  appendRow_('Applications', {
    application_id: newId_('APP'), applicant_type: 'วิสามัญ', applicant_user_id: applicant3,
    ref_data_json: JSON.stringify({
      org_name_th: 'มูลนิธิเดโมเพื่อสังคม', org_type: 'องค์กรภาคเอกชน', juristic_no: '0000000000000',
      founded_date: '2015-01-01', purposes: ['ให้ความช่วยเหลือทางสังคมหรือบริการทางสังคม'],
      rep_prefix: 'นาง', rep_name: 'ผู้แทน เดโม', rep_dob: '1975-03-15', rep_ethnicity: 'ไทย', rep_nationality: 'ไทย',
      rep_religion: 'พุทธ', rep_marital_status: 'สมรส', rep_position: 'ประธานมูลนิธิ',
      rep_national_id: '1100000000097', rep_phone: '0890000097', rep_email: 'org.demo@example.com',
      addr_province: 'กรุงเทพมหานคร', office_phone: '021234567', duration: '1'
    }),
    documents_json: '[]', status: 'รอตรวจสอบ', reject_reason: '', fee_amount: MEMBERSHIP_FEE_TABLE['วิสามัญ']['1'] + MEMBERSHIP_ENTRANCE_FEE + CERTIFICATE_MAILING_FEE,
    submitted_at: nowIso_(), reviewed_by: '', reviewed_at: '', approved_member_id: ''
  });

  // เดโม 4: สมาชิกวิสามัญ (องค์กร) ที่อนุมัติแล้ว สถานะปกติ — ไว้ทดสอบมุมมองสมาชิกวิสามัญที่เป็นสมาชิกจริงแล้ว (ไม่ใช่แค่ผู้สมัคร)
  const orgUserId = newId_('USR');
  appendRow_('Users', {
    user_id: orgUserId, username: 'demo_org1', password_hash: hashPassword_('Demo1234'),
    role: 'member', member_id: '', thaid_pid: '', status: 'active', created_at: nowIso_(), last_login: ''
  });
  const orgId = newId_('ORG');
  appendRow_('MembersOrg', {
    org_id: orgId, org_name_th: 'สมาคมสวัสดิการสังคมเดโม', org_name_en: 'Demo Social Welfare Association',
    org_type: 'องค์กรภาคเอกชน', juristic_no: '0999999999999', tax_id: '', founded_date: '2010-06-01',
    purposes_json: JSON.stringify(['ป้องกันแก้ไขปัญหาให้แก่บุคคล ครอบครัว กลุ่มคนหรือชุมชน ให้กระทำหน้าที่ทางสังคมและดำรงชีวิตได้อย่างปกติสุข', 'ให้ความช่วยเหลือทางสังคมหรือบริการทางสังคม']),
    rep_prefix: 'นาง', rep_name: 'ผู้แทน องค์กรเดโม', rep_dob: '1970-08-20', rep_ethnicity: 'ไทย', rep_nationality: 'ไทย',
    rep_religion: 'พุทธ', rep_marital_status: 'สมรส', rep_position: 'นายกสมาคม',
    rep_national_id: '1100000000096', rep_phone: '0890000096', rep_phone2: '', rep_email: 'org1.demo@example.com', rep_line_id: '',
    addr_house_no: '99', addr_building: '', addr_floor: '', addr_moo: '', addr_soi: '', addr_road: 'พหลโยธิน',
    addr_subdistrict: 'จตุจักร', addr_district: 'จตุจักร', addr_province: 'กรุงเทพมหานคร', addr_zipcode: '10900',
    office_phone: '025551234', website: 'https://example.org',
    reg_no: generateOrgRegNo_(), applied_date: addYearsIso_(nowIso_(), -2), approved_date: addYearsIso_(nowIso_(), -2),
    status: 'ปกติ', end_date: addYearsIso_(nowIso_(), 3), note: '', documents_json: '[]',
    created_at: nowIso_(), updated_at: nowIso_()
  });
  updateRow_('Users', findRowById_('Users', 'user_id', orgUserId)._row, { member_id: orgId });

  // ---------- ข้อมูลเดโมสำหรับระบบสวัสดิการสมาชิก (TOR ข้อ 7) — อ้างอิงระเบียบสภาฯ ว่าด้วยสวัสดิการสมาชิกสามัญ พ.ศ. 2567 ----------
  if (member1Id) {
    appendRow_('Beneficiaries', {
      beneficiary_id: newId_('BEN'), member_id: member1Id, prefix: 'นาง', name: 'สมหมาย ใจดี',
      phone: '0891234567', relationship: 'สามีหรือภรรยา', created_at: nowIso_()
    });
    const benId2 = newId_('BEN');
    appendRow_('Beneficiaries', {
      beneficiary_id: benId2, member_id: member1Id, prefix: 'ด.ช.', name: 'สมหวัง ใจดี',
      phone: '0891234568', relationship: 'บุตรโดยชอบด้วยกฎหมาย', created_at: nowIso_()
    });

    // คำขอ 1: รอตรวจสอบ — ประเภท "อื่นๆ" (ไม่มีเพดานตายตัว แล้วแต่คณะอนุกรรมการพิจารณา)
    appendRow_('Welfare', {
      welfare_id: newId_('WEL'), member_id: member1Id,
      welfare_type: 'สวัสดิการหรือบริการอื่นๆ ตามที่คณะอนุกรรมการสวัสดิการอนุมัติ', beneficiary_id: '', incident_date: '',
      description: 'ขอรับความช่วยเหลือกรณีพิเศษ (ข้อมูลเดโม)', documents_json: '[]', status: 'รอตรวจสอบ', reject_reason: '',
      requested_at: nowIso_(), reviewed_by: '', reviewed_at: '', approved_amount: '', note: '',
      disbursed_at: '', disbursed_by: '', disbursement_doc_url: ''
    });

    // คำขอ 2: อนุมัติแล้ว รอเบิกจ่าย — ประเภทประสบภัยพิบัติ (เพดาน 10,000 บาทตามข้อ 22(2))
    appendRow_('Welfare', {
      welfare_id: newId_('WEL'), member_id: member1Id,
      welfare_type: 'การช่วยเหลือกรณีประสบภัยพิบัติหรือสาธารณภัย', beneficiary_id: '',
      incident_date: addYearsIso_(nowIso_(), -0.03),
      description: 'บ้านเสียหายจากน้ำท่วม (ข้อมูลเดโม)', documents_json: '[]', status: 'อนุมัติ', reject_reason: '',
      requested_at: addYearsIso_(nowIso_(), -0.02), reviewed_by: 'demo_staff', reviewed_at: nowIso_(),
      approved_amount: '3000', note: 'อนุมัติตามระเบียบ ข้อ 22(2) (ข้อมูลเดโม)',
      disbursed_at: '', disbursed_by: '', disbursement_doc_url: ''
    });

    // คำขอ 3: ไม่อนุมัติ — ประเภทถึงแก่กรรม (เพดาน 5,000 บาทตามข้อ 22(1)) ผูกกับทายาทลำดับที่ 1 (บุตร)
    appendRow_('Welfare', {
      welfare_id: newId_('WEL'), member_id: member1Id,
      welfare_type: 'การช่วยเหลือกรณีถึงแก่กรรม', beneficiary_id: benId2,
      incident_date: addYearsIso_(nowIso_(), -0.06),
      description: 'คำขอตัวอย่าง (ข้อมูลเดโม)', documents_json: '[]', status: 'ไม่อนุมัติ',
      reject_reason: 'เอกสารประกอบไม่ครบถ้วน (ข้อมูลเดโม)',
      requested_at: addYearsIso_(nowIso_(), -0.05), reviewed_by: 'demo_staff', reviewed_at: nowIso_(), approved_amount: '', note: '',
      disbursed_at: '', disbursed_by: '', disbursement_doc_url: ''
    });

    // คำขอ 4: อนุมัติแล้วและเบิกจ่ายเงินเรียบร้อยแล้ว (สถานะเบิกจ่ายจริง) — ไว้ทดสอบหน้า "ประวัติการเปลี่ยนแปลง" และดูเอกสารยืนยันการจ่ายเงินที่เคยออกแล้ว
    appendRow_('Welfare', {
      welfare_id: newId_('WEL'), member_id: member1Id,
      welfare_type: 'สวัสดิการหรือบริการอื่นๆ ตามที่คณะอนุกรรมการสวัสดิการอนุมัติ', beneficiary_id: '',
      incident_date: addYearsIso_(nowIso_(), -0.1),
      description: 'ขอรับความช่วยเหลือกรณีเจ็บป่วย (ข้อมูลเดโม — เบิกจ่ายแล้ว)', documents_json: '[]', status: 'อนุมัติ', reject_reason: '',
      requested_at: addYearsIso_(nowIso_(), -0.09), reviewed_by: 'demo_staff', reviewed_at: addYearsIso_(nowIso_(), -0.08),
      approved_amount: '2000', note: 'อนุมัติตามระเบียบ (ข้อมูลเดโม)',
      disbursed_at: addYearsIso_(nowIso_(), -0.07), disbursed_by: 'demo_staff', disbursement_doc_url: ''
    });
  }

  // ---------- ข้อมูลเดโมสำหรับระบบชำระเงิน (TOR ข้อ 11) — ตัวอย่างคำขอคืนเงินครบทั้ง 3 สถานะ (รอตรวจสอบ/อนุมัติรอโอนคืน/คืนเงินแล้ว) ----------
  if (member1Id) {
    const paidPayment = readAll_('Payments').find(pm => pm.member_id === member1Id && pm.status === 'ชำระแล้ว');
    if (paidPayment) {
      appendRow_('Refunds', {
        refund_id: newId_('RFD'), payment_id: paidPayment.payment_id, member_id: member1Id,
        reason: 'ชำระเงินซ้ำซ้อน (ข้อมูลเดโม)', requested_by: 'demo_staff', requested_at: nowIso_(),
        status: 'รอตรวจสอบ', approved_by: '', approved_at: '', refund_amount: paidPayment.amount, refund_note: '', refunded_at: ''
      });
    }
  }
  // คำขอคืนเงิน 2: อนุมัติแล้ว รอยืนยันการโอนคืนจริง — ผูกกับสมาชิกที่ขึ้นทะเบียนสำเร็จแล้ว (member10Id)
  if (member10Id) {
    const paidPayment10 = readAll_('Payments').find(pm => pm.member_id === member10Id && pm.status === 'ชำระแล้ว');
    if (paidPayment10) {
      appendRow_('Refunds', {
        refund_id: newId_('RFD'), payment_id: paidPayment10.payment_id, member_id: member10Id,
        reason: 'ยกเลิกการขึ้นทะเบียนหลังชำระเงิน (ข้อมูลเดโม)', requested_by: 'demo_staff', requested_at: addYearsIso_(nowIso_(), -0.02),
        status: 'อนุมัติ', approved_by: 'demo_staff', approved_at: nowIso_(),
        refund_amount: paidPayment10.amount, refund_note: 'อนุมัติคืนเงินเต็มจำนวน (ข้อมูลเดโม)', refunded_at: ''
      });
    }
  }
  // คำขอคืนเงิน 3: คืนเงินให้สมาชิกเรียบร้อยแล้ว (สถานะสุดท้าย) — ผูกกับสมาชิกที่ต่ออายุสำเร็จแล้ว (member14Id)
  if (member14Id) {
    const paidPayment14 = readAll_('Payments').find(pm => pm.member_id === member14Id && pm.status === 'คืนเงินแล้ว');
    const paidPayment14b = paidPayment14 || readAll_('Payments').find(pm => pm.member_id === member14Id && pm.status === 'ชำระแล้ว');
    if (paidPayment14b) {
      appendRow_('Refunds', {
        refund_id: newId_('RFD'), payment_id: paidPayment14b.payment_id, member_id: member14Id,
        reason: 'ชำระค่าธรรมเนียมผิดประเภท (ข้อมูลเดโม)', requested_by: 'demo_staff', requested_at: addYearsIso_(nowIso_(), -0.05),
        status: 'คืนเงินแล้ว', approved_by: 'demo_staff', approved_at: addYearsIso_(nowIso_(), -0.04),
        refund_amount: paidPayment14b.amount, refund_note: 'โอนคืนเข้าบัญชีสมาชิกเรียบร้อยแล้ว (ข้อมูลเดโม)', refunded_at: addYearsIso_(nowIso_(), -0.03)
      });
    }
  }

  // บัญชีเจ้าหน้าที่เดโม (นอกเหนือจากแอดมินคนแรกที่สร้างผ่าน setupFirstAdmin)
  const existingStaff = readAll_('Users').find(u => u.username === 'demo_staff');
  if (!existingStaff) {
    appendRow_('Users', {
      user_id: newId_('USR'), username: 'demo_staff', password_hash: hashPassword_('Demo1234'),
      role: 'admin_staff', member_id: '', thaid_pid: '', status: 'active',
      created_at: nowIso_(), last_login: ''
    });
  }
  const existingWelfareStaff = readAll_('Users').find(u => u.username === 'demo_welfare');
  if (!existingWelfareStaff) {
    appendRow_('Users', {
      user_id: newId_('USR'), username: 'demo_welfare', password_hash: hashPassword_('Demo1234'),
      role: 'admin_welfare', member_id: '', thaid_pid: '', status: 'active',
      created_at: nowIso_(), last_login: '', display_name: 'เจ้าหน้าที่สวัสดิการ (เดโม)'
    });
  }
  // บัญชีผู้ดูแลระบบเดโมเพิ่มเติม (แยกจากแอดมินคนแรกที่สร้างผ่าน setupFirstAdmin ซึ่งสงวนไว้เป็นบัญชีจริงของผู้ดูแลระบบตัวจริง)
  // สร้างผ่านการเขียนข้อมูลตรง (ไม่ผ่าน action_setupFirstAdmin ที่จำกัดให้สร้างได้ครั้งเดียว) เพื่อไว้ทดสอบสิทธิ์ระดับสูงสุดโดยไม่ต้องใช้บัญชีจริง
  const existingDemoAdmin = readAll_('Users').find(u => u.username === 'demo_admin');
  if (!existingDemoAdmin) {
    appendRow_('Users', {
      user_id: newId_('USR'), username: 'demo_admin', password_hash: hashPassword_('Demo1234'),
      role: 'admin_super', member_id: '', thaid_pid: '', status: 'active',
      created_at: nowIso_(), last_login: '', display_name: 'ผู้ดูแลระบบ (เดโม)'
    });
  }

  // PDPA: บัญชีสมาชิกเดโมถือว่าให้ความยินยอมแล้ว (ไม่ต้องกดซ้ำทุกครั้งที่สาธิต) — บัญชีที่สมัครใหม่จริงจะเห็นหน้าความยินยอมก่อนใช้งาน
  readAll_('Users').filter(u => u.role === 'member' && !latestPdpaConsentFor_(u.user_id)).forEach(u => recordPdpaConsent_(u, 'ให้ความยินยอม', 'ข้อมูลตัวอย่าง (Demo)', ''));

  return 'สร้างข้อมูลเดโมสำเร็จ: สมาชิก 22 คน (ครอบคลุมทุกสถานะของระบบขึ้นทะเบียน/ต่ออายุใบอนุญาตตั้งแต่ก่อนยื่นจนสำเร็จ รวมสมทบปกติ+เปลี่ยนประเภทสมาชิก+เปลี่ยนชื่อ-นามสกุล 4 สถานะ+สมาชิกที่ยังไม่เคยผูกบัญชี), ' +
    'สมาชิกวิสามัญ 1 องค์กร (อนุมัติแล้ว สถานะปกติ), ผู้สมัคร 4 สถานะ (รอตรวจสอบ/รอชำระเงิน/รอตรวจสลิป/วิสามัญ), กลุ่มต่ออายุตัวอย่าง 1 กลุ่ม, ' +
    'ผู้รับผลประโยชน์ 2 รายการและคำขอสวัสดิการ 4 สถานะ (รอตรวจสอบ/อนุมัติรอเบิกจ่าย/ไม่อนุมัติ/เบิกจ่ายแล้ว), คำขอคืนเงิน 3 สถานะ (รอตรวจสอบ/อนุมัติรอโอนคืน/คืนเงินแล้ว), ' +
    'บัญชีเจ้าหน้าที่เดโมครบ 3 ระดับ (ผู้ดูแลระบบ/ทะเบียน-การเงิน/สวัสดิการ) — ดูรายละเอียดบัญชีทั้งหมดใน DEMO_GUIDE.md';
}

/** ล้างข้อมูลเดโมทั้งหมด (แยกจากข้อมูลจริงด้วยเลขบัตร 1100000000xxx และ username demo_*) */
function clearDemoData() {
  // เก็บ member_id ของสมาชิกเดโมไว้ก่อน เพื่อใช้ลบข้อมูลที่เกี่ยวข้องในแท็บอื่น (Licenses, MembershipRenewals, Notifications)
  const demoMemberIds = readAll_('Members')
    .filter(m => String(m.national_id || '').indexOf('1100000000') === 0)
    .map(m => m.member_id);
  // สมาชิกกิตติมศักดิ์เดโม (สร้างโดย demo_staff) — ใช้ลบบัตร (Cards) ที่ผูกอยู่ด้วย
  const demoHonoraryIds = readHonorarySafe_().filter(h => h.created_by === 'demo_staff').map(h => h.honorary_id);

  ['Members', 'MembersOrg', 'Users', 'Applications', 'Payments', 'PdpaConsents', 'PdpaRequests'].forEach(name => {
    const sh = getSS_().getSheetByName(name);
    if (!sh) return;
    const rows = readAll_(name);
    for (let i = rows.length - 1; i >= 0; i--) {
      const r = rows[i];
      const isDemo =
        (r.national_id && String(r.national_id).indexOf('1100000000') === 0) ||
        (r.rep_national_id && String(r.rep_national_id).indexOf('1100000000') === 0) ||
        (r.username && String(r.username).indexOf('demo_') === 0) ||
        (r.ref_data_json && String(r.ref_data_json).indexOf('1100000000') !== -1) ||
        (r.ref_data_json && demoMemberIds.some(id => id && String(r.ref_data_json).indexOf(id) !== -1)) ||
        (r.member_id && demoMemberIds.indexOf(r.member_id) !== -1) ||
        (r.member_id && demoHonoraryIds.indexOf(r.member_id) !== -1);
      if (isDemo) sh.deleteRow(r._row);
    }
  });

  ['Licenses', 'LicenseSanctions', 'MembershipRenewals', 'Notifications', 'VoidedLicenses', 'LicenseRenewalHistory', 'RenewalEligibility', 'Cards', 'QrScanLog', 'LicenseDownloadLog', 'Beneficiaries', 'Welfare', 'WelfareDocLog', 'Refunds', 'ReconciliationLog', 'MembersHonorary', 'ExamRounds', 'InterviewSessions', 'InterviewGroups', 'InterviewCandidates'].forEach(name => {
    const sh = getSS_().getSheetByName(name);
    if (!sh) return;
    const rows = readAll_(name);
    for (let i = rows.length - 1; i >= 0; i--) {
      const r = rows[i];
      const isDemo =
        demoMemberIds.indexOf(r.member_id) !== -1 ||
        (r.member_id && demoHonoraryIds.indexOf(r.member_id) !== -1) ||
        (r.license_no && String(r.license_no).indexOf('DEMO') !== -1) ||
        (r.old_license_no && String(r.old_license_no).indexOf('DEMO') !== -1) ||
        (r.imported_by === 'demo_staff') || (r.renewed_by === 'demo_staff') || (r.actor_user_id === 'demo_staff') ||
        (r.performed_by === 'demo_staff') || (r.requested_by === 'demo_staff') || (r.created_by === 'demo_staff');
      if (isDemo) sh.deleteRow(r._row);
    }
  });

  return 'ล้างข้อมูลเดโมเรียบร้อย';
}

// ---------- Utility: sheet <-> JSON ----------

// ---------- PDPA: การเข้ารหัสข้อมูลอ่อนไหว (Encryption at rest) ----------
// TOR (ข้อกำหนดเพิ่มเติม PDPA): ข้อมูลอ่อนไหว (ศาสนา เชื้อชาติ เลขประจำตัวประชาชน) ต้องจัดเก็บแบบเข้ารหัส และจำกัดสิทธิ์ตามบทบาท
// วิธีการ: AES-256-CBC + HMAC-SHA256 (encrypt-then-MAC) เขียนด้วย JavaScript ล้วน (Apps Script ไม่มี AES ในตัว) เข้ารหัสทุกครั้งที่เขียนลงชีต
// และถอดรหัสทุกครั้งที่อ่านผ่าน readAll_ — โค้ดส่วนอื่นของระบบจึงเห็นค่าปกติ ส่วนในไฟล์ Google Sheets จะเห็นเป็น "enc1:..." เท่านั้น
// กุญแจเก็บใน Script Properties ชื่อ DATA_ENCRYPTION_KEY (สร้างอัตโนมัติครั้งแรก) — ห้ามลบ/เปลี่ยน มิฉะนั้นข้อมูลที่เข้ารหัสแล้วจะถอดไม่ได้ ควรสำรองกุญแจไว้ต่างหาก
const ENCRYPTED_COLUMNS = {
  Members: ['national_id', 'religion', 'ethnicity'],
  MembersOrg: ['rep_national_id', 'rep_religion', 'rep_ethnicity']
};
// คอลัมน์ JSON ที่ในเนื้อหามีข้อมูลอ่อนไหวฝังอยู่ (เช่น ข้อมูลในใบสมัคร, ค่าเดิม/ค่าใหม่ใน Audit log) — เข้ารหัสเฉพาะ "ค่า" ของคีย์ที่ระบุ โครงสร้าง JSON ยังอ่านได้
const ENCRYPTED_JSON_COLUMNS = {
  Applications: ['ref_data_json'],
  AuditLog: ['old_value_json', 'new_value_json']
};
const SENSITIVE_JSON_KEYS = ['national_id', 'religion', 'ethnicity', 'rep_national_id', 'rep_religion', 'rep_ethnicity'];
const ENC_PREFIX = 'enc1:';

const AES_SBOX_ = (function () {
  // สร้าง S-box จาก GF(2^8) ตามนิยามของ AES (ไม่ฝังตารางยาวในโค้ด)
  const s = new Array(256), inv = new Array(256);
  let p = 1, q = 1;
  do {
    p = p ^ ((p << 1) & 0xff) ^ (p & 0x80 ? 0x1b : 0);
    q ^= q << 1; q ^= q << 2; q ^= q << 4; q &= 0xff; if (q & 0x80) q ^= 0x09;
    const x = q ^ ((q << 1) | (q >> 7)) ^ ((q << 2) | (q >> 6)) ^ ((q << 3) | (q >> 5)) ^ ((q << 4) | (q >> 4));
    s[p] = (x ^ 0x63) & 0xff;
  } while (p !== 1);
  s[0] = 0x63;
  for (let i = 0; i < 256; i++) inv[s[i]] = i;
  return { s: s, inv: inv };
})();

function aesXtime_(a) { return ((a << 1) ^ (a & 0x80 ? 0x1b : 0)) & 0xff; }
function aesMul_(a, b) { let r = 0; while (b) { if (b & 1) r ^= a; a = aesXtime_(a); b >>= 1; } return r; }

function aesExpandKey_(key) { // key: 32 ไบต์ → round keys 15 ชุด (แต่ละชุด 16 ไบต์)
  const S = AES_SBOX_.s, w = [];
  for (let i = 0; i < 32; i++) w[i] = key[i];
  let rcon = 1;
  for (let i = 8; i < 60; i++) {
    let t = w.slice((i - 1) * 4, i * 4);
    if (i % 8 === 0) { t = [S[t[1]] ^ rcon, S[t[2]], S[t[3]], S[t[0]]]; rcon = aesXtime_(rcon); }
    else if (i % 8 === 4) t = t.map(function (b) { return S[b]; });
    for (let j = 0; j < 4; j++) w[i * 4 + j] = w[(i - 8) * 4 + j] ^ t[j];
  }
  return w;
}

function aesEncryptBlock_(w, blk) {
  const S = AES_SBOX_.s;
  let s = blk.slice();
  const addKey = function (r) { for (let i = 0; i < 16; i++) s[i] ^= w[r * 16 + i]; };
  addKey(0);
  for (let r = 1; r <= 14; r++) {
    s = s.map(function (b) { return S[b]; });
    const t = s.slice(); // ShiftRows (column-major)
    for (let c = 0; c < 4; c++) for (let row = 0; row < 4; row++) s[c * 4 + row] = t[((c + row) % 4) * 4 + row];
    if (r !== 14) {
      for (let c = 0; c < 4; c++) {
        const a = s.slice(c * 4, c * 4 + 4);
        s[c * 4] = aesXtime_(a[0]) ^ (aesXtime_(a[1]) ^ a[1]) ^ a[2] ^ a[3];
        s[c * 4 + 1] = a[0] ^ aesXtime_(a[1]) ^ (aesXtime_(a[2]) ^ a[2]) ^ a[3];
        s[c * 4 + 2] = a[0] ^ a[1] ^ aesXtime_(a[2]) ^ (aesXtime_(a[3]) ^ a[3]);
        s[c * 4 + 3] = (aesXtime_(a[0]) ^ a[0]) ^ a[1] ^ a[2] ^ aesXtime_(a[3]);
      }
    }
    addKey(r);
  }
  return s;
}

function aesDecryptBlock_(w, blk) {
  const IS = AES_SBOX_.inv;
  let s = blk.slice();
  const addKey = function (r) { for (let i = 0; i < 16; i++) s[i] ^= w[r * 16 + i]; };
  addKey(14);
  for (let r = 13; r >= 0; r--) {
    const t = s.slice(); // InvShiftRows
    for (let c = 0; c < 4; c++) for (let row = 0; row < 4; row++) s[((c + row) % 4) * 4 + row] = t[c * 4 + row];
    s = s.map(function (b) { return IS[b]; });
    addKey(r);
    if (r !== 0) {
      for (let c = 0; c < 4; c++) {
        const a = s.slice(c * 4, c * 4 + 4);
        s[c * 4] = aesMul_(a[0], 14) ^ aesMul_(a[1], 11) ^ aesMul_(a[2], 13) ^ aesMul_(a[3], 9);
        s[c * 4 + 1] = aesMul_(a[0], 9) ^ aesMul_(a[1], 14) ^ aesMul_(a[2], 11) ^ aesMul_(a[3], 13);
        s[c * 4 + 2] = aesMul_(a[0], 13) ^ aesMul_(a[1], 9) ^ aesMul_(a[2], 14) ^ aesMul_(a[3], 11);
        s[c * 4 + 3] = aesMul_(a[0], 11) ^ aesMul_(a[1], 13) ^ aesMul_(a[2], 9) ^ aesMul_(a[3], 14);
      }
    }
  }
  return s;
}

function sha256Bytes_(bytes) {
  const K = [0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
    0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da, 0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
    0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
    0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2];
  const H = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
  const m = bytes.slice(), l = m.length;
  m.push(0x80);
  while (m.length % 64 !== 56) m.push(0);
  const hi = Math.floor(l / 0x20000000), lo = (l << 3) >>> 0;
  m.push((hi >>> 24) & 255, (hi >>> 16) & 255, (hi >>> 8) & 255, hi & 255, (lo >>> 24) & 255, (lo >>> 16) & 255, (lo >>> 8) & 255, lo & 255);
  const rotr = function (x, n) { return (x >>> n) | (x << (32 - n)); };
  for (let i = 0; i < m.length; i += 64) {
    const w = new Array(64);
    for (let t = 0; t < 16; t++) w[t] = ((m[i + t * 4] << 24) | (m[i + t * 4 + 1] << 16) | (m[i + t * 4 + 2] << 8) | m[i + t * 4 + 3]) | 0;
    for (let t = 16; t < 64; t++) {
      const s0 = rotr(w[t - 15], 7) ^ rotr(w[t - 15], 18) ^ (w[t - 15] >>> 3);
      const s1 = rotr(w[t - 2], 17) ^ rotr(w[t - 2], 19) ^ (w[t - 2] >>> 10);
      w[t] = (w[t - 16] + s0 + w[t - 7] + s1) | 0;
    }
    let a = H[0], b = H[1], c = H[2], d = H[3], e = H[4], f = H[5], g = H[6], h = H[7];
    for (let t = 0; t < 64; t++) {
      const S1 = rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25), ch = (e & f) ^ (~e & g);
      const t1 = (h + S1 + ch + K[t] + w[t]) | 0;
      const S0 = rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22), mj = (a & b) ^ (a & c) ^ (b & c);
      const t2 = (S0 + mj) | 0;
      h = g; g = f; f = e; e = (d + t1) | 0; d = c; c = b; b = a; a = (t1 + t2) | 0;
    }
    H[0] = (H[0] + a) | 0; H[1] = (H[1] + b) | 0; H[2] = (H[2] + c) | 0; H[3] = (H[3] + d) | 0;
    H[4] = (H[4] + e) | 0; H[5] = (H[5] + f) | 0; H[6] = (H[6] + g) | 0; H[7] = (H[7] + h) | 0;
  }
  const out = [];
  H.forEach(function (x) { out.push((x >>> 24) & 255, (x >>> 16) & 255, (x >>> 8) & 255, x & 255); });
  return out;
}

function hmacSha256Bytes_(key, msg) {
  let k = key.slice();
  if (k.length > 64) k = sha256Bytes_(k);
  while (k.length < 64) k.push(0);
  const ipad = k.map(function (b) { return b ^ 0x36; }), opad = k.map(function (b) { return b ^ 0x5c; });
  return sha256Bytes_(opad.concat(sha256Bytes_(ipad.concat(msg))));
}

function utf8Bytes_(str) {
  const s = unescape(encodeURIComponent(String(str))), out = [];
  for (let i = 0; i < s.length; i++) out.push(s.charCodeAt(i));
  return out;
}
function utf8Str_(bytes) {
  let s = '';
  for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
  return decodeURIComponent(escape(s));
}
const B64_ = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
function b64Enc_(bytes) {
  let o = '';
  for (let i = 0; i < bytes.length; i += 3) {
    const n = (bytes[i] << 16) | ((bytes[i + 1] || 0) << 8) | (bytes[i + 2] || 0);
    o += B64_[(n >> 18) & 63] + B64_[(n >> 12) & 63] + (i + 1 < bytes.length ? B64_[(n >> 6) & 63] : '=') + (i + 2 < bytes.length ? B64_[n & 63] : '=');
  }
  return o;
}
function b64Dec_(str) {
  const out = [], clean = String(str).replace(/=+$/, '');
  let buf = 0, bits = 0;
  for (let i = 0; i < clean.length; i++) {
    const v = B64_.indexOf(clean[i]);
    if (v < 0) throw new Error('base64 ไม่ถูกต้อง');
    buf = (buf << 6) | v; bits += 6;
    if (bits >= 8) { bits -= 8; out.push((buf >> bits) & 255); buf &= (1 << bits) - 1; }
  }
  return out;
}
function hexToBytes_(hex) { const o = []; for (let i = 0; i < hex.length; i += 2) o.push(parseInt(hex.substr(i, 2), 16)); return o; }
function bytesToHex_(b) { return b.map(function (x) { return ('0' + x.toString(16)).slice(-2); }).join(''); }

let ENC_KEYS_CACHE_ = null;
/** คืนกุญแจ {enc, mac} (อย่างละ 32 ไบต์) — อ่านจาก Script Properties DATA_ENCRYPTION_KEY; ถ้ายังไม่มีจะสร้างและเก็บให้อัตโนมัติ */
function getEncKeys_() {
  if (ENC_KEYS_CACHE_) return ENC_KEYS_CACHE_;
  const props = PropertiesService.getScriptProperties();
  let master = props.getProperty('DATA_ENCRYPTION_KEY');
  if (!master || !/^[0-9a-f]{64}$/i.test(master)) {
    const seed = [];
    for (let i = 0; i < 4; i++) seed.push(Utilities.getUuid());
    master = bytesToHex_(sha256Bytes_(utf8Bytes_(seed.join('|') + '|' + new Date().getTime())));
    props.setProperty('DATA_ENCRYPTION_KEY', master);
  }
  const mk = hexToBytes_(master);
  ENC_KEYS_CACHE_ = {
    enc: aesExpandKey_(sha256Bytes_(mk.concat(utf8Bytes_('enc')))),
    mac: sha256Bytes_(mk.concat(utf8Bytes_('mac')))
  };
  return ENC_KEYS_CACHE_;
}

function isEncryptedValue_(v) { return typeof v === 'string' && v.indexOf(ENC_PREFIX) === 0; }

/** เข้ารหัสข้อความ → "enc1:" + base64(IV16 + ciphertext + HMAC32); ค่าว่างคืนค่าว่าง, ค่าที่เข้ารหัสแล้วไม่เข้ารหัสซ้ำ */
function encryptField_(plain) {
  if (plain === '' || plain === null || plain === undefined) return '';
  if (isEncryptedValue_(plain)) return plain;
  const k = getEncKeys_();
  const iv = sha256Bytes_(utf8Bytes_(Utilities.getUuid() + Utilities.getUuid() + Math.random())).slice(0, 16);
  const data = utf8Bytes_(plain), pad = 16 - (data.length % 16);
  for (let i = 0; i < pad; i++) data.push(pad);
  let prev = iv, ct = [];
  for (let i = 0; i < data.length; i += 16) {
    const blk = data.slice(i, i + 16).map(function (b, j) { return b ^ prev[j]; });
    prev = aesEncryptBlock_(k.enc, blk);
    ct = ct.concat(prev);
  }
  const body = iv.concat(ct);
  return ENC_PREFIX + b64Enc_(body.concat(hmacSha256Bytes_(k.mac, body)));
}

/** ถอดรหัส — ค่าที่ไม่ได้ขึ้นต้นด้วย enc1: (ข้อมูลเดิมก่อนเปิดใช้การเข้ารหัส) คืนค่าเดิม; ตรวจ HMAC ไม่ผ่านจะ throw เพื่อไม่ส่งข้อมูลที่ถูกแก้ไขต่อ */
function decryptField_(val) {
  if (!isEncryptedValue_(val)) return val;
  const k = getEncKeys_(), all = b64Dec_(val.slice(ENC_PREFIX.length));
  if (all.length < 16 + 16 + 32 || (all.length - 32) % 16 !== 0) throw new Error('ข้อมูลที่เข้ารหัสไม่ถูกต้อง');
  const body = all.slice(0, all.length - 32), mac = all.slice(all.length - 32), calc = hmacSha256Bytes_(k.mac, body);
  let diff = 0;
  for (let i = 0; i < 32; i++) diff |= mac[i] ^ calc[i];
  if (diff !== 0) throw new Error('ตรวจสอบความสมบูรณ์ของข้อมูลที่เข้ารหัสไม่ผ่าน (กุญแจไม่ตรงหรือข้อมูลถูกแก้ไข)');
  let prev = body.slice(0, 16), out = [];
  for (let i = 16; i < body.length; i += 16) {
    const blk = body.slice(i, i + 16);
    out = out.concat(aesDecryptBlock_(k.enc, blk).map(function (b, j) { return b ^ prev[j]; }));
    prev = blk;
  }
  const pad = out[out.length - 1];
  if (pad < 1 || pad > 16) throw new Error('ข้อมูลที่เข้ารหัสไม่ถูกต้อง');
  return utf8Str_(out.slice(0, out.length - pad));
}

function walkSensitiveJson_(node, fn) {
  if (Array.isArray(node)) { node.forEach(function (x) { walkSensitiveJson_(x, fn); }); return; }
  if (node && typeof node === 'object') {
    Object.keys(node).forEach(function (k) {
      if (SENSITIVE_JSON_KEYS.indexOf(k) !== -1 && typeof node[k] === 'string' && node[k] !== '') node[k] = fn(node[k]);
      else if (node[k] && typeof node[k] === 'object') walkSensitiveJson_(node[k], fn);
    });
  }
}
function transformJsonField_(text, fn) {
  if (typeof text !== 'string' || text === '' || !SENSITIVE_JSON_KEYS.some(function (k) { return text.indexOf('"' + k + '"') !== -1; })) return text;
  try { const o = JSON.parse(text); walkSensitiveJson_(o, fn); return JSON.stringify(o); } catch (e) { return text; }
}

function encryptRowObject_(sheetName, obj) {
  const cols = ENCRYPTED_COLUMNS[sheetName], jcols = ENCRYPTED_JSON_COLUMNS[sheetName];
  if (!cols && !jcols) return obj;
  const o = Object.assign({}, obj);
  (cols || []).forEach(function (c) { if (o[c] !== undefined && o[c] !== null && o[c] !== '') o[c] = encryptField_(String(o[c])); });
  (jcols || []).forEach(function (c) { if (o[c] !== undefined) o[c] = transformJsonField_(o[c], encryptField_); });
  return o;
}
function decryptRowObject_(sheetName, obj) {
  const cols = ENCRYPTED_COLUMNS[sheetName], jcols = ENCRYPTED_JSON_COLUMNS[sheetName];
  (cols || []).forEach(function (c) { if (isEncryptedValue_(obj[c])) obj[c] = decryptField_(obj[c]); });
  (jcols || []).forEach(function (c) { if (typeof obj[c] === 'string' && obj[c].indexOf(ENC_PREFIX) !== -1) obj[c] = transformJsonField_(obj[c], decryptField_); });
  return obj;
}

/** ผู้ดูแลสูงสุด: เข้ารหัสข้อมูลอ่อนไหวที่ยังเป็นข้อความธรรมดาในชีตเดิม (รันครั้งเดียวหลังอัปเดตระบบ; รันซ้ำได้ปลอดภัย) */
function action_encryptExistingSensitiveData(p, user) {
  requireRole_(user, ['admin_super']);
  const summary = {};
  const sheets = Object.keys(ENCRYPTED_COLUMNS).concat(Object.keys(ENCRYPTED_JSON_COLUMNS));
  sheets.forEach(function (sheetName) {
    const sh = getSheet_(sheetName), values = sh.getDataRange().getValues();
    if (values.length < 2) { summary[sheetName] = 0; return; }
    const headers = values[0];
    const cols = (ENCRYPTED_COLUMNS[sheetName] || []).concat(ENCRYPTED_JSON_COLUMNS[sheetName] || []).map(function (c) { return headers.indexOf(c); }).filter(function (i) { return i >= 0; });
    let changed = 0;
    const out = values.slice(1).map(function (row) {
      const obj = {};
      cols.forEach(function (i) { obj[headers[i]] = row[i]; });
      const enc = encryptRowObject_(sheetName, obj), nr = row.slice();
      let touched = false;
      cols.forEach(function (i) { if (enc[headers[i]] !== row[i]) { nr[i] = enc[headers[i]]; touched = true; } });
      if (touched) changed++;
      return nr;
    });
    if (changed) sh.getRange(2, 1, out.length, headers.length).setValues(out); // เขียนครั้งเดียวทั้งชีต (เร็วกว่าอัปเดตทีละเซลล์มาก)
    summary[sheetName] = changed;
  });
  auditLog_('System', 'encryption', 'เข้ารหัสข้อมูลอ่อนไหวที่มีอยู่เดิม (PDPA)', {}, summary, user.user_id);
  return { ok: true, rows_encrypted: summary };
}

/** ตรวจสถานะการเข้ารหัส: นับแถวที่ยังเป็นข้อความธรรมดา (ไม่คืนค่าข้อมูลจริง) — ใช้ยืนยันหลังรัน encryptExistingSensitiveData */
function action_getEncryptionStatus(p, user) {
  requireRole_(user, ['admin_super']);
  const out = {};
  Object.keys(ENCRYPTED_COLUMNS).concat(Object.keys(ENCRYPTED_JSON_COLUMNS)).forEach(function (sheetName) {
    const sh = getSheet_(sheetName), values = sh.getDataRange().getValues();
    if (values.length < 2) { out[sheetName] = { encrypted: 0, plaintext: 0 }; return; }
    const headers = values[0], plainCols = ENCRYPTED_COLUMNS[sheetName] || [], jsonCols = ENCRYPTED_JSON_COLUMNS[sheetName] || [];
    let enc = 0, plain = 0;
    values.slice(1).forEach(function (row) {
      plainCols.forEach(function (c) { const v = row[headers.indexOf(c)]; if (v === '' || v === undefined) return; if (isEncryptedValue_(v)) enc++; else plain++; });
      jsonCols.forEach(function (c) { // JSON: นับเฉพาะแถวที่มีคีย์อ่อนไหวอยู่ — ถ้ายังเห็นค่าที่ไม่ได้ขึ้นต้น enc1: ถือว่ายังเป็นข้อความธรรมดา
        const v = row[headers.indexOf(c)]; if (typeof v !== 'string' || v === '') return;
        let hasPlain = false, hasEnc = false;
        transformJsonField_(v, function (x) { if (isEncryptedValue_(x)) hasEnc = true; else hasPlain = true; return x; });
        if (hasPlain) plain++; else if (hasEnc) enc++;
      });
    });
    out[sheetName] = { encrypted: enc, plaintext: plain };
  });
  return { ok: true, columns: ENCRYPTED_COLUMNS, json_columns: ENCRYPTED_JSON_COLUMNS, status: out, key_configured: !!PropertiesService.getScriptProperties().getProperty('DATA_ENCRYPTION_KEY') };
}

// ---------- PDPA: ความยินยอม (Consent) และการใช้สิทธิของเจ้าของข้อมูล ----------
// TOR (ข้อกำหนดเพิ่มเติม PDPA): ต้องมีกลไกขอความยินยอมจากสมาชิกในการจัดเก็บและใช้ข้อมูลส่วนบุคคล พร้อมช่องทางให้ใช้สิทธิตาม พ.ร.บ.คุ้มครองข้อมูลส่วนบุคคล
// - ความยินยอมผูกกับ "บัญชีผู้ใช้" + "เวอร์ชันของประกาศ" (เปลี่ยนเวอร์ชันเมื่อใด สมาชิกต้องกดยินยอมใหม่) เก็บทุกครั้งเป็นแถวใหม่ (append-only) เพื่อเป็นหลักฐาน
// - ข้อความประกาศด้านล่างเป็นร่างตั้งต้น ฝ่ายกฎหมาย/เจ้าหน้าที่คุ้มครองข้อมูล (DPO) ของสภาฯ ต้องตรวจและแก้ไขก่อนใช้จริง (ชื่อผู้ควบคุมข้อมูล ระยะเวลาเก็บรักษา ช่องทางติดต่อ)
const PDPA_POLICY_VERSION = '1.0';
const PDPA_REQUEST_TYPES = ['ขอเข้าถึง/ขอสำเนาข้อมูล', 'ขอแก้ไขข้อมูลให้ถูกต้อง', 'ขอลบ/ทำลายข้อมูล', 'ขอระงับการใช้ข้อมูล', 'คัดค้านการประมวลผลข้อมูล', 'ขอโอนย้ายข้อมูล'];
const PDPA_REQUEST_STATUSES = ['รอดำเนินการ', 'กำลังดำเนินการ', 'ดำเนินการแล้ว', 'ปฏิเสธ'];
const PDPA_GATED_ACTIONS = ['submitApplication']; // action ที่ต้องมีความยินยอมปัจจุบันก่อนจึงจะส่งข้อมูลส่วนบุคคลเข้าระบบได้

function getPdpaPolicy_() {
  const contact = PropertiesService.getScriptProperties().getProperty('PDPA_CONTACT') || 'เจ้าหน้าที่คุ้มครองข้อมูลส่วนบุคคล สภาวิชาชีพสังคมสงเคราะห์ (กรอกอีเมล/โทรศัพท์ในค่า PDPA_CONTACT ของ Script Properties)';
  return {
    version: PDPA_POLICY_VERSION,
    title: 'ประกาศความเป็นส่วนตัวและคำขอความยินยอมเก็บ ใช้ และเปิดเผยข้อมูลส่วนบุคคล',
    controller: 'สภาวิชาชีพสังคมสงเคราะห์ (ผู้ควบคุมข้อมูลส่วนบุคคล)',
    sections: [
      { heading: 'ข้อมูลที่เก็บรวบรวม', body: 'ข้อมูลระบุตัวตน (ชื่อ-นามสกุล เลขประจำตัวประชาชน วันเกิด) ข้อมูลติดต่อ (ที่อยู่ โทรศัพท์ อีเมล) ข้อมูลการศึกษาและการทำงาน เอกสารประกอบการสมัคร และข้อมูลการชำระเงิน รวมถึงข้อมูลอ่อนไหว ได้แก่ ศาสนาและเชื้อชาติ' },
      { heading: 'วัตถุประสงค์การใช้ข้อมูล', body: 'เพื่อดำเนินการรับสมัครและขึ้นทะเบียนสมาชิก ออกบัตรสมาชิกและใบอนุญาตประกอบวิชาชีพ การต่ออายุ การสอบ การพิจารณาสวัสดิการ การออกใบเสร็จ การแจ้งเตือนที่เกี่ยวข้องกับสถานะสมาชิก/ใบอนุญาต และการจัดทำสถิติ' },
      { heading: 'การเปิดเผยข้อมูล', body: 'ข้อมูลที่เปิดเผยต่อสาธารณะผ่านหน้าค้นหา/ตรวจสอบ QR จำกัดเฉพาะ ชื่อ-นามสกุล เลขที่ใบอนุญาต สถานะ และวันหมดอายุ ข้อมูลอื่นเปิดให้เฉพาะเจ้าหน้าที่ตามบทบาทที่ได้รับมอบหมาย' },
      { heading: 'การคุ้มครองข้อมูลอ่อนไหว', body: 'เลขประจำตัวประชาชน ศาสนา และเชื้อชาติ ถูกเก็บแบบเข้ารหัส (Encryption at rest) และจำกัดสิทธิ์การเข้าถึงตามบทบาทผู้ใช้งาน การเข้าถึงและแก้ไขข้อมูลมีบันทึกประวัติการใช้งาน' },
      { heading: 'ระยะเวลาการเก็บรักษา', body: 'เก็บตลอดระยะเวลาที่เป็นสมาชิกและตามระยะเวลาที่กฎหมาย/ระเบียบของสภาฯ กำหนด [ฝ่ายกฎหมายของสภาฯ โปรดระบุจำนวนปีที่แน่นอน]' },
      { heading: 'สิทธิของเจ้าของข้อมูล', body: 'ท่านมีสิทธิขอเข้าถึง/ขอสำเนา แก้ไข ลบหรือทำลาย ระงับการใช้ คัดค้านการประมวลผล และขอโอนย้ายข้อมูล รวมถึงถอนความยินยอมเมื่อใดก็ได้ ผ่านเมนู “ข้อมูลส่วนบุคคล (PDPA)” ในระบบนี้ การถอนความยินยอมอาจทำให้ไม่สามารถใช้บริการที่ต้องอาศัยข้อมูลดังกล่าวได้' },
      { heading: 'ช่องทางติดต่อ', body: contact }
    ],
    consent_statement: 'ข้าพเจ้าได้อ่านและเข้าใจประกาศความเป็นส่วนตัวนี้ และยินยอมให้สภาวิชาชีพสังคมสงเคราะห์เก็บ ใช้ และเปิดเผยข้อมูลส่วนบุคคลของข้าพเจ้า รวมถึงข้อมูลอ่อนไหว (ศาสนา เชื้อชาติ เลขประจำตัวประชาชน) ตามวัตถุประสงค์ข้างต้น',
    request_types: PDPA_REQUEST_TYPES
  };
}

function action_getPdpaPolicy(p) { return getPdpaPolicy_(); }

function latestPdpaConsentFor_(userId) {
  const rows = readAll_('PdpaConsents').filter(function (r) { return r.user_id === userId; });
  return rows.length ? rows[rows.length - 1] : null; // append-only: แถวสุดท้ายคือสถานะล่าสุด
}

function hasCurrentPdpaConsent_(userId) {
  const c = latestPdpaConsentFor_(userId);
  return !!(c && c.status === 'ให้ความยินยอม' && String(c.policy_version) === PDPA_POLICY_VERSION);
}

function recordPdpaConsent_(user, status, channel, reason) {
  const row = {
    consent_id: newId_('PDC'), user_id: user.user_id, username: user.username || '', member_id: user.member_id || '',
    policy_version: PDPA_POLICY_VERSION, status: status, channel: channel || 'เว็บไซต์', reason: reason || '', created_at: nowIso_()
  };
  appendRow_('PdpaConsents', row);
  auditLog_('PdpaConsents', row.consent_id, status === 'ให้ความยินยอม' ? 'pdpa_consent_granted' : 'pdpa_consent_withdrawn', {}, { policy_version: PDPA_POLICY_VERSION, channel: row.channel }, user.user_id);
  return row;
}

/** ใช้ใน action ที่เก็บข้อมูลส่วนบุคคล: สมาชิกต้องมีความยินยอมตามเวอร์ชันปัจจุบัน */
function requirePdpaConsent_(user) {
  if (user && user.role === 'member' && !hasCurrentPdpaConsent_(user.user_id)) {
    throw new Error('ต้องให้ความยินยอมเก็บและใช้ข้อมูลส่วนบุคคล (PDPA) ก่อนดำเนินการต่อ — ไปที่เมนู “ข้อมูลส่วนบุคคล (PDPA)”');
  }
}

function action_getMyPdpaStatus(p, user) {
  if (user.role !== 'member') return { required: false, accepted: true, policy_version: PDPA_POLICY_VERSION };
  const c = latestPdpaConsentFor_(user.user_id);
  return {
    required: true, policy_version: PDPA_POLICY_VERSION, accepted: hasCurrentPdpaConsent_(user.user_id),
    last_status: c ? c.status : '', last_version: c ? c.policy_version : '', last_at: c ? c.created_at : ''
  };
}

function action_acceptPdpaConsent(p, user) {
  requireRole_(user, ['member']);
  if (p.accept !== true) throw new Error('กรุณาติ๊กช่องยืนยันการให้ความยินยอม');
  if (String(p.policy_version || '') !== PDPA_POLICY_VERSION) throw new Error('ประกาศความเป็นส่วนตัวมีการอัปเดต กรุณาโหลดหน้านี้ใหม่แล้วอ่านอีกครั้ง');
  recordPdpaConsent_(user, 'ให้ความยินยอม', 'เว็บไซต์ (หน้าความยินยอม)', '');
  return { ok: true, accepted: true, policy_version: PDPA_POLICY_VERSION };
}

function action_withdrawPdpaConsent(p, user) {
  requireRole_(user, ['member']);
  const reason = String(p.reason || '').trim();
  if (!reason) throw new Error('กรุณาระบุเหตุผลการถอนความยินยอม');
  if (!latestPdpaConsentFor_(user.user_id)) throw new Error('ยังไม่มีประวัติการให้ความยินยอม');
  recordPdpaConsent_(user, 'ถอนความยินยอม', 'เว็บไซต์ (หน้าความยินยอม)', reason);
  return { ok: true, accepted: false };
}

function action_listMyPdpaHistory(p, user) {
  requireRole_(user, ['member']);
  return {
    consents: readAll_('PdpaConsents').filter(function (r) { return r.user_id === user.user_id; }).map(function (r) { return { policy_version: r.policy_version, status: r.status, channel: r.channel, reason: r.reason, created_at: r.created_at }; }).reverse(),
    requests: readAll_('PdpaRequests').filter(function (r) { return r.user_id === user.user_id; }).sort(function (a, b) { return String(b.created_at).localeCompare(String(a.created_at)); })
  };
}

function action_submitPdpaRequest(p, user) {
  requireRole_(user, ['member']);
  if (PDPA_REQUEST_TYPES.indexOf(p.request_type) === -1) throw new Error('ประเภทคำขอไม่ถูกต้อง');
  const detail = String(p.detail || '').trim();
  if (detail.length < 5) throw new Error('กรุณาระบุรายละเอียดคำขออย่างน้อย 5 ตัวอักษร');
  const row = {
    request_id: newId_('PDR'), user_id: user.user_id, username: user.username || '', member_id: user.member_id || '',
    request_type: p.request_type, detail: detail, status: 'รอดำเนินการ', response_note: '', handled_by: '', handled_at: '', created_at: nowIso_()
  };
  appendRow_('PdpaRequests', row);
  auditLog_('PdpaRequests', row.request_id, 'pdpa_request_submitted', {}, { request_type: row.request_type }, user.user_id);
  return { ok: true, request_id: row.request_id };
}

/** สิทธิขอเข้าถึงข้อมูล: สมาชิกดาวน์โหลดสำเนาข้อมูลของตนเอง (เลขบัตร/ศาสนา/เชื้อชาติ ถอดรหัสแล้ว เฉพาะของตนเองเท่านั้น) */
function action_exportMyPersonalData(p, user) {
  requireRole_(user, ['member']);
  const out = { exported_at: nowIso_(), username: user.username, profile: null, consents: [], requests: [] };
  if (user.member_id) {
    out.profile = findRowById_('Members', 'member_id', user.member_id) || findRowById_('MembersOrg', 'org_id', user.member_id) || null;
    if (out.profile) { out.profile = Object.assign({}, out.profile); delete out._row; }
  }
  out.consents = readAll_('PdpaConsents').filter(function (r) { return r.user_id === user.user_id; }).map(function (r) { return { policy_version: r.policy_version, status: r.status, created_at: r.created_at }; });
  out.requests = readAll_('PdpaRequests').filter(function (r) { return r.user_id === user.user_id; });
  auditLog_('PdpaRequests', user.user_id, 'pdpa_export_my_data', {}, {}, user.user_id);
  return out;
}

// ---- ฝั่งเจ้าหน้าที่ ----
function action_listPdpaRequests(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  let rows = readAll_('PdpaRequests');
  if (p.status) rows = rows.filter(function (r) { return r.status === p.status; });
  return rows.sort(function (a, b) { return String(b.created_at).localeCompare(String(a.created_at)); }).slice(0, 500);
}

function action_handlePdpaRequest(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const req = findRowById_('PdpaRequests', 'request_id', p.request_id);
  if (!req) throw new Error('ไม่พบคำขอ');
  if (PDPA_REQUEST_STATUSES.indexOf(p.status) === -1 || p.status === 'รอดำเนินการ') throw new Error('สถานะไม่ถูกต้อง');
  const note = String(p.response_note || '').trim();
  if ((p.status === 'ดำเนินการแล้ว' || p.status === 'ปฏิเสธ') && !note) throw new Error('กรุณาระบุข้อความตอบกลับ/เหตุผล');
  updateRow_('PdpaRequests', req._row, { status: p.status, response_note: note, handled_by: user.user_id, handled_at: nowIso_() });
  auditLog_('PdpaRequests', req.request_id, 'pdpa_request_' + p.status, { status: req.status }, { status: p.status }, user.user_id);
  return { ok: true };
}

function action_listPdpaConsents(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const latest = {};
  readAll_('PdpaConsents').forEach(function (r) { latest[r.user_id] = r; }); // แถวหลังสุดของแต่ละบัญชี
  const users = readAll_('Users').filter(function (u) { return u.role === 'member'; });
  const rows = users.map(function (u) {
    const c = latest[u.user_id];
    const ok = !!(c && c.status === 'ให้ความยินยอม' && String(c.policy_version) === PDPA_POLICY_VERSION);
    return { user_id: u.user_id, username: u.username, member_id: u.member_id, status: ok ? 'ยินยอม (ฉบับปัจจุบัน)' : (c ? (c.status === 'ถอนความยินยอม' ? 'ถอนความยินยอม' : 'ยินยอมฉบับเก่า — ต้องให้ใหม่') : 'ยังไม่เคยให้ความยินยอม'), policy_version: c ? c.policy_version : '', last_at: c ? c.created_at : '' };
  });
  const filtered = p.status ? rows.filter(function (r) { return r.status.indexOf(p.status) === 0; }) : rows;
  return { policy_version: PDPA_POLICY_VERSION, total: rows.length, consented: rows.filter(function (r) { return r.status.indexOf('ยินยอม (') === 0; }).length, rows: filtered.slice(0, 1000) };
}

function getSheet_(name) {
  const sh = getSS_().getSheetByName(name);
  if (!sh) throw new Error('ไม่พบแท็บชื่อ ' + name + ' กรุณาสร้างตาม SHEET_SCHEMA.md');
  return sh;
}

function readAll_(sheetName) {
  const sh = getSheet_(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1)
    .filter(row => row.some(c => c !== ''))
    .map((row, i) => {
      const obj = { _row: i + 2 }; // เก็บเลขแถวจริงไว้ใช้ตอน update
      headers.forEach((h, idx) => obj[h] = row[idx]);
      if (sheetName === 'ExamRounds') obj.round_label = fixRoundLabel_(obj.round_label);
      return decryptRowObject_(sheetName, obj); // PDPA: ถอดรหัสคอลัมน์อ่อนไหวให้โค้ดส่วนอื่นเห็นเป็นค่าปกติ
    });
}

/** ซ่อมชื่อรอบสอบที่ Google Sheets แปลง "2/2569" เป็นวันที่ (Date หรือ ISO เช่น 2569-01-31T17:00:00.000Z) ให้กลับเป็น "เดือน/ปี" (เวลาไทย UTC+7) */
function fixRoundLabel_(v) {
  let d = null;
  if (Object.prototype.toString.call(v) === '[object Date]') d = v;
  else if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}T/.test(v)) d = new Date(v);
  if (!d || isNaN(d.getTime())) return v;
  const th = new Date(d.getTime() + 7 * 3600 * 1000);
  return (th.getUTCMonth() + 1) + '/' + th.getUTCFullYear();
}

/** แก้ round_label ใน ref_data_json ของใบสมัครสอบ (ที่อาจถูกบันทึกเป็น ISO ไว้แล้ว) ให้ตรงกับชื่อรอบจริง */
function fixExamRefRoundLabel_(app, round) {
  try {
    const d = JSON.parse(app.ref_data_json || '{}');
    const fixed = round ? round.round_label : fixRoundLabel_(d.round_label);
    if (fixed === d.round_label) return app;
    d.round_label = fixed;
    return Object.assign({}, app, { ref_data_json: JSON.stringify(d) });
  } catch (e) { return app; }
}

function appendRow_(sheetName, obj) {
  obj = encryptRowObject_(sheetName, obj); // PDPA: เข้ารหัสคอลัมน์อ่อนไหวก่อนลงชีต
  const sh = getSheet_(sheetName);
  const headers = sh.getDataRange().getValues()[0];
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : '');
  const targetRow = sh.getLastRow() + 1;

  // บังคับ format คอลัมน์ที่เสี่ยงเลข 0 นำหน้าหายเป็นข้อความล้วนที่แถวนี้โดยตรงก่อนเขียนค่าเสมอ (ไม่พึ่งพา format ที่ตั้งไว้ตอน setupAllSheets เพียงอย่างเดียว
  // เพราะแถวที่เพิ่มใหม่ผ่าน appendRow() อาจไม่สืบทอด format ของคอลัมน์ในบางกรณี — วิธีนี้ทำงานถูกต้องแน่นอนไม่ว่าจะเคยรัน setupAllSheets ใหม่หรือไม่)
  const digitCols = DIGIT_STRING_COLUMNS[sheetName];
  if (digitCols && digitCols.length) {
    digitCols.forEach(colName => {
      const colIdx = headers.indexOf(colName);
      if (colIdx !== -1) sh.getRange(targetRow, colIdx + 1).setNumberFormat('@');
    });
  }

  sh.getRange(targetRow, 1, 1, row.length).setValues([row]);
  return targetRow;
}

function updateRow_(sheetName, rowNumber, patch) {
  patch = encryptRowObject_(sheetName, patch); // PDPA: เข้ารหัสคอลัมน์อ่อนไหวก่อนลงชีต
  const sh = getSheet_(sheetName);
  const headers = sh.getDataRange().getValues()[0];
  const digitCols = DIGIT_STRING_COLUMNS[sheetName] || [];
  headers.forEach((h, idx) => {
    if (patch[h] !== undefined) {
      if (digitCols.indexOf(h) !== -1) sh.getRange(rowNumber, idx + 1).setNumberFormat('@');
      sh.getRange(rowNumber, idx + 1).setValue(patch[h]);
    }
  });
}

function deleteRow_(sheetName, rowNumber) {
  getSheet_(sheetName).deleteRow(rowNumber);
}

function findRowById_(sheetName, idField, idValue) {
  const rows = readAll_(sheetName);
  return rows.find(r => String(r[idField]) === String(idValue));
}

function newId_(prefix) {
  return prefix + '-' + Utilities.getUuid().slice(0, 8);
}

function nowIso_() {
  return new Date().toISOString();
}

function auditLog_(refType, refId, action, oldVal, newVal, actorUserId) {
  appendRow_('AuditLog', {
    log_id: newId_('LOG'),
    ref_type: refType,
    ref_id: refId,
    action: action,
    old_value_json: JSON.stringify(oldVal || {}),
    new_value_json: JSON.stringify(newVal || {}),
    actor_user_id: actorUserId || '',
    created_at: nowIso_()
  });
}

// ---------- Auth ----------

/** ตรวจสอบค่า boolean ที่เก็บไว้ในชีต — Google Sheets บางครั้งแปลงข้อความ "true"/"false" ที่เขียนเข้าไปให้กลายเป็นชนิด Boolean จริงโดยอัตโนมัติ
 * (ไม่ใช่ string "true" อีกต่อไปตอนอ่านกลับมา) ทำให้เทียบด้วย `=== 'true'` ตรงๆ อาจพลาดได้ — ฟังก์ชันนี้รองรับทั้งสองรูปแบบ
 */
function isTrue_(val) {
  return val === true || val === 'true' || val === 'TRUE' || val === 1;
}

function hashPassword_(plain) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, plain);
  return bytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

// ---------- นโยบายความปลอดภัยบัญชีผู้ใช้ (TOR ข้อ 9 "ข้อกำหนดเพิ่มเติม: ความปลอดภัยของระบบ") ----------
const PASSWORD_MIN_LENGTH = 8; // ความยาวรหัสผ่านขั้นต่ำ
const MAX_FAILED_LOGIN_ATTEMPTS = 5; // ล็อกบัญชีเมื่อพยายามเข้าสู่ระบบผิดพลาดเกินจำนวนนี้
const LOCKOUT_MINUTES = 15; // ระยะเวลาที่บัญชีถูกล็อกหลังพยายามผิดครบจำนวน

/** ตรวจสอบว่ารหัสผ่านผ่านเกณฑ์ขั้นต่ำหรือไม่ (ความยาว + ต้องมีทั้งตัวอักษรและตัวเลข) */
function validatePasswordPolicy_(password) {
  if (!password || password.length < PASSWORD_MIN_LENGTH) {
    throw new Error('รหัสผ่านต้องมีความยาวอย่างน้อย ' + PASSWORD_MIN_LENGTH + ' ตัวอักษร');
  }
  if (!/[a-zA-Z]/.test(password) || !/[0-9]/.test(password)) {
    throw new Error('รหัสผ่านต้องมีทั้งตัวอักษรและตัวเลขอย่างน้อยอย่างละ 1 ตัว');
  }
}

function makeToken_(user) {
  const token = Utilities.getUuid();
  appendRow_('Sessions', {
    token: token,
    user_id: user.user_id,
    role: user.role,
    created_at: nowIso_(),
    expires_at: new Date(Date.now() + 8 * 3600 * 1000).toISOString() // 8 ชั่วโมง
  });
  return token;
}

function getSessionUser_(token) {
  if (!token) return null;
  const session = findRowById_('Sessions', 'token', token);
  if (!session) return null;
  if (new Date(session.expires_at) < new Date()) return null;
  const user = findRowById_('Users', 'user_id', session.user_id);
  if (!user || user.status !== 'active') return null;
  return user;
}

function requireRole_(user, roles) {
  if (!user || roles.indexOf(user.role) === -1) {
    throw new Error('ไม่มีสิทธิ์เข้าถึง (403)');
  }
}

/** สร้างแอดมินคนแรก — เรียกครั้งเดียวตอนติดตั้งระบบเท่านั้น แล้วลบ/ปิด action นี้ */
function action_setupFirstAdmin(p) {
  const props = PropertiesService.getScriptProperties();
  const setupKey = props.getProperty('ADMIN_SETUP_KEY');
  if (!setupKey || p.setupKey !== setupKey) throw new Error('setupKey ไม่ถูกต้อง');
  // ไม่นับบัญชีแอดมินเดโม (username ขึ้นต้นด้วย demo_) เป็น "มีแอดมินอยู่แล้ว" — กันปัญหาสร้างข้อมูลเดโมแล้วบล็อกไม่ให้ตั้งค่าแอดมินจริงครั้งแรกได้
  const existing = readAll_('Users').find(u => u.role.indexOf('admin') === 0 && String(u.username || '').indexOf('demo_') !== 0);
  if (existing) throw new Error('มีแอดมินอยู่แล้วในระบบ');
  validatePasswordPolicy_(p.password);
  const user = {
    user_id: newId_('USR'),
    username: p.username,
    password_hash: hashPassword_(p.password),
    role: 'admin_super',
    member_id: '',
    status: 'active',
    created_at: nowIso_(),
    last_login: '',
    failed_login_attempts: 0,
    locked_until: '',
    display_name: p.display_name || ''
  };
  appendRow_('Users', user);
  return { ok: true };
}

/** สำหรับสมาชิกสามัญ/สมทบที่มีข้อมูลอยู่แล้วในระบบ (นำเข้าจากฐานข้อมูลเดิม) แต่ยังไม่เคยสร้างบัญชีเข้าใช้งานเอง
 * ยืนยันตัวตนด้วยเลขประจำตัวประชาชน + เบอร์โทรศัพท์ที่ตรงกับข้อมูลสมาชิกเดิม แล้วผูกบัญชีใหม่เข้ากับข้อมูลนั้นทันที (ไม่ต้องรออนุมัติ)
 * ไม่ใช้กับสมาชิกวิสามัญ (องค์กร) เพราะผู้สมัครอาจเป็นสมาชิกสามัญอยู่แล้ว ใช้เลขบัตรประชาชนซ้ำกับบัญชีส่วนตัวได้ วิสามัญจึงยังคงใช้วิธีตั้งชื่อผู้ใช้งาน/รหัสผ่านเองแบบเดิม
 */
function action_claimExistingMemberAccount(p) {
  const nationalId = String(p.national_id || '').trim();
  const phone = String(p.phone || '').trim();
  if (!nationalId || !phone) throw new Error('กรุณากรอกเลขประจำตัวประชาชนและเบอร์โทรศัพท์');

  const member = readAll_('Members').find(m =>
    String(m.national_id || '').trim() === nationalId && String(m.phone || '').trim() === phone
  );
  if (!member) throw new Error('ไม่พบข้อมูลสมาชิกที่ตรงกับเลขประจำตัวประชาชนและเบอร์โทรศัพท์นี้ กรุณาตรวจสอบข้อมูลอีกครั้ง หรือติดต่อเจ้าหน้าที่หากคิดว่าข้อมูลในระบบไม่ตรงกับที่ท่านให้ไว้ตอนสมัครสมาชิก');

  const existingUser = readAll_('Users').find(u => u.member_id === member.member_id);
  if (existingUser) throw new Error('บัญชีของสมาชิกท่านนี้มีอยู่แล้วในระบบ กรุณาเข้าสู่ระบบปกติ หรือติดต่อเจ้าหน้าที่หากลืมชื่อผู้ใช้งาน/รหัสผ่าน');

  const username = String(p.username || '').trim();
  if (!username) throw new Error('กรุณาตั้งชื่อผู้ใช้งาน');
  const existingUsername = findRowById_('Users', 'username', username);
  if (existingUsername) throw new Error('ชื่อผู้ใช้งานนี้ถูกใช้แล้ว กรุณาเลือกชื่ออื่น');
  validatePasswordPolicy_(p.password);
  if (p.pdpa_accept !== true || String(p.pdpa_version || '') !== PDPA_POLICY_VERSION) throw new Error('ต้องอ่านและยินยอมประกาศความเป็นส่วนตัว (PDPA) ก่อนสมัครใช้งานบัญชี');

  const userId = newId_('USR');
  appendRow_('Users', {
    user_id: userId, username: username, password_hash: hashPassword_(p.password),
    role: 'member', member_id: member.member_id, thaid_pid: '', status: 'active',
    created_at: nowIso_(), last_login: '', failed_login_attempts: 0, locked_until: ''
  });
  recordPdpaConsent_({ user_id: userId, username: username, member_id: member.member_id }, 'ให้ความยินยอม', 'เว็บไซต์ (ตอนผูกบัญชีสมาชิกเดิม)', '');
  auditLog_('Users', userId, 'ผูกบัญชีใหม่กับข้อมูลสมาชิกเดิม (ยืนยันตัวตนด้วยเลขบัตรประชาชน+เบอร์โทรศัพท์)', {}, { member_id: member.member_id }, userId);
  return { ok: true };
}

function action_login(p) {
  const user = findRowById_('Users', 'username', p.username);
  if (!user || user.status !== 'active') throw new Error('ไม่พบบัญชีผู้ใช้งาน หรือบัญชีถูกระงับ');

  // ตรวจสอบว่าบัญชีถูกล็อกอยู่หรือไม่ (TOR ข้อ 9 — ล็อกบัญชีเมื่อพยายามเข้าสู่ระบบผิดพลาดเกินจำนวนที่กำหนด)
  if (user.locked_until && new Date(user.locked_until) > new Date()) {
    const minutesLeft = Math.ceil((new Date(user.locked_until) - new Date()) / 60000);
    throw new Error('บัญชีนี้ถูกล็อกชั่วคราวเนื่องจากพยายามเข้าสู่ระบบผิดพลาดหลายครั้ง กรุณาลองใหม่อีกครั้งใน ' + minutesLeft + ' นาที');
  }

  if (user.password_hash !== hashPassword_(p.password)) {
    const attempts = (parseInt(user.failed_login_attempts, 10) || 0) + 1;
    const patch = { failed_login_attempts: attempts };
    if (attempts >= MAX_FAILED_LOGIN_ATTEMPTS) {
      patch.locked_until = new Date(Date.now() + LOCKOUT_MINUTES * 60000).toISOString();
      updateRow_('Users', user._row, patch);
      throw new Error('รหัสผ่านไม่ถูกต้อง — พยายามผิดครบ ' + MAX_FAILED_LOGIN_ATTEMPTS + ' ครั้ง บัญชีถูกล็อกชั่วคราว ' + LOCKOUT_MINUTES + ' นาที');
    }
    updateRow_('Users', user._row, patch);
    throw new Error('รหัสผ่านไม่ถูกต้อง (เหลือโอกาสอีก ' + (MAX_FAILED_LOGIN_ATTEMPTS - attempts) + ' ครั้งก่อนบัญชีถูกล็อกชั่วคราว)');
  }

  // เข้าสู่ระบบสำเร็จ — รีเซ็ตตัวนับความผิดพลาด
  const token = makeToken_(user);
  updateRow_('Users', user._row, { last_login: nowIso_(), failed_login_attempts: 0, locked_until: '' });
  return {
    token: token,
    role: user.role,
    username: user.username,
    member_id: user.member_id,
    account_type: accountTypeOf_(user)
  };
}

/** ประเภทบัญชี: 'org' = บัญชีสำหรับองค์กร (สมาชิกวิสามัญ) | 'person' = บัญชีบุคคล
 * บัญชีเก่าที่ไม่มีค่า: ถ้าผูกกับ ORG... ถือเป็น org มิฉะนั้นเป็น person */
function accountTypeOf_(user) {
  if (user && user.account_type) return user.account_type;
  return user && String(user.member_id || '').indexOf('ORG') === 0 ? 'org' : 'person';
}

function action_register(p) {
  // สมาชิกสมัครบัญชีผู้ใช้งานของตนเอง (ไม่ผูก member_id จนกว่าใบสมัครจะได้รับอนุมัติ — ดู action_approveApplication)
  // p.account_type: 'person' (ค่าเริ่มต้น) | 'org' — บัญชีองค์กรแยกจากบัญชีบุคคล เพราะตัวแทนหน่วยงานอาจเป็นสมาชิกสภาฯ ประเภทอื่นอยู่แล้ว
  // (ชื่อผู้ใช้งานซ้ำกันไม่ได้ จึงต้องตั้ง username แยกสำหรับองค์กร)
  const accountType = p.account_type === 'org' ? 'org' : 'person';
  const existing = findRowById_('Users', 'username', p.username);
  if (existing) throw new Error('ชื่อผู้ใช้งานนี้ถูกใช้แล้ว' + (accountType === 'org' ? ' — บัญชีองค์กรต้องตั้งชื่อผู้ใช้งานแยกจากบัญชีบุคคล' : ''));
  validatePasswordPolicy_(p.password);
  if (p.pdpa_accept !== true || String(p.pdpa_version || '') !== PDPA_POLICY_VERSION) throw new Error('ต้องอ่านและยินยอมประกาศความเป็นส่วนตัว (PDPA) ก่อนสมัครใช้งานบัญชี');
  const user = {
    user_id: newId_('USR'),
    username: p.username,
    password_hash: hashPassword_(p.password),
    role: 'member',
    member_id: '',
    status: 'active',
    created_at: nowIso_(),
    last_login: '',
    failed_login_attempts: 0,
    locked_until: '',
    display_name: accountType === 'org' ? String(p.display_name || '').trim() : '',
    account_type: accountType
  };
  appendRow_('Users', user);
  recordPdpaConsent_(user, 'ให้ความยินยอม', 'เว็บไซต์ (ตอนสมัครใช้งานบัญชี)', '');
  return { ok: true };
}

// ---------- Drive uploads & email notifications (ใช้ร่วมกับระบบสมัครสมาชิก) ----------

function getDocsFolder_() {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty('DRIVE_FOLDER_ID');
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (e) { /* โฟลเดอร์เดิมหาไม่เจอ สร้างใหม่ */ }
  }
  const folder = DriveApp.createFolder('SWC_Documents');
  props.setProperty('DRIVE_FOLDER_ID', folder.getId());
  return folder;
}

/** รับไฟล์เอกสารแนบจากฟอร์มสมัครสมาชิก (base64) แล้วเก็บใน Google Drive
 * PDPA: ตั้งสิทธิ์แชร์แบบ "ทุกคนที่มีลิงก์ดูได้" เพื่อให้เจ้าหน้าที่ตรวจสอบเอกสารได้ผ่าน URL โดยตรง
 * (ยังไม่ใช่การเข้ารหัสระดับ production — ดู ROADMAP.md เรื่องการยกระดับความปลอดภัยเอกสาร)
 */
function action_uploadDocument(p, user) {
  requireRole_(user, ['member', 'admin_super', 'admin_staff']);
  if (!p.base64 || !p.filename) throw new Error('ข้อมูลไฟล์ไม่ครบ');
  const bytes = Utilities.base64Decode(p.base64);
  const blob = Utilities.newBlob(bytes, p.mimeType || 'application/octet-stream', p.filename);
  const folder = getDocsFolder_();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { name: p.filename, doc_type: p.doc_type || '', url: file.getUrl() };
}

/** ส่งอีเมลแจ้งเตือน — ห่อด้วยแม่แบบ HTML ของสภาฯ (หัวสีกรมท่า+ตราสภาฯ, กล่องเนื้อหา, ปุ่มลิงก์, ท้ายอีเมล) พร้อมข้อความล้วนสำรอง
 * body เป็นข้อความธรรมดาเหมือนเดิม: บรรทัดว่างแยกย่อหน้า, กลุ่มบรรทัดรูปแบบ "หัวข้อ: ค่า" ติดกันตั้งแต่ 2 บรรทัดขึ้นไปจะแสดงเป็นตารางรายละเอียด,
 * บรรทัดขึ้นต้นด้วย "- " เป็นรายการ, ลิงก์ http(s) ที่ยืนอยู่บรรทัดเดียวหรือท้ายบรรทัดจะกลายเป็นปุ่ม
 * opts (ไม่บังคับ): { buttonText, buttonUrl, tone: 'success'|'danger'|'info' }
 */
function mailNotify_(to, subject, body, opts) {
  if (!to) return;
  try {
    MailApp.sendEmail({ to: to, subject: subject, body: String(body || ''), htmlBody: buildEmailHtml_(subject, body, opts || {}), name: 'สภาวิชาชีพสังคมสงเคราะห์' });
  } catch (err) { /* เกิน quota หรืออีเมลไม่ถูกต้อง — ข้ามไปไม่ให้ action ล้มเหลว */ }
}

function escHtml_(s) {
  return String(s === undefined || s === null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function buildEmailHtml_(subject, body, opts) {
  const base = (PropertiesService.getScriptProperties().getProperty('FRONTEND_BASE_URL') || '').replace(/\/?$/, '/');
  const logo = base !== '/' ? base + 'logo.png' : '';
  let tone = opts.tone;
  if (!tone) {
    if (/ไม่ผ่าน|ไม่อนุมัติ|ปฏิเสธ|หมดอายุ|ถูกยกเลิก|เพิกถอน|ไม่สำเร็จ/.test(subject)) tone = 'danger';
    else if (/ผ่าน|อนุมัติ|สำเร็จ|ยืนยัน|ได้รับ/.test(subject)) tone = 'success';
    else tone = 'info';
  }
  const accent = { success: '#2f6b4f', danger: '#c23b2e', info: '#1d5589' }[tone];
  const lines = String(body || '').split(/\r?\n/);
  let buttonUrl = opts.buttonUrl || '', buttonText = opts.buttonText || '';
  const urlRe = /(https?:\/\/[^\s]+)/;
  let html = '', i = 0;
  while (i < lines.length) {
    const line = lines[i].trim();
    if (!line) { i++; continue; }
    // ลิงก์ → ปุ่ม (ใช้ลิงก์แรกที่พบ ถ้ายังไม่ได้กำหนด)
    const um = line.match(urlRe);
    if (um && !buttonUrl) {
      buttonUrl = um[1];
      const label = line.replace(um[1], '').replace(/[:：\s]+$/, '').trim();
      if (label) html += '<p style="margin:0 0 14px;line-height:1.7">' + escHtml_(label) + '</p>';
      if (!buttonText) buttonText = 'เปิดลิงก์';
      i++; continue;
    }
    // กลุ่มรายการ "หัวข้อ: ค่า"
    const kv = l => { const m = l.trim().match(/^([^:：]{1,40})[:：]\s*(.+)$/); return m && !urlRe.test(l) ? m : null; };
    if (kv(line)) {
      let j = i, rows = [];
      while (j < lines.length && lines[j].trim() && kv(lines[j])) { rows.push(kv(lines[j])); j++; }
      if (rows.length >= 2) {
        html += '<table role="presentation" style="width:100%;border-collapse:collapse;margin:0 0 16px;background:#f7f9fb;border:1px solid #e3e8ee">' +
          rows.map(r => '<tr><td style="padding:8px 12px;color:#6b7278;width:38%;border-bottom:1px solid #e3e8ee;font-size:14px">' + escHtml_(r[1]) + '</td><td style="padding:8px 12px;font-weight:600;border-bottom:1px solid #e3e8ee;font-size:14px">' + escHtml_(r[2]) + '</td></tr>').join('') + '</table>';
        i = j; continue;
      }
    }
    if (/^[-•]\s+/.test(line)) {
      let items = [];
      while (i < lines.length && /^[-•]\s+/.test(lines[i].trim())) { items.push(lines[i].trim().replace(/^[-•]\s+/, '')); i++; }
      html += '<ul style="margin:0 0 14px;padding-left:20px;line-height:1.8">' + items.map(t => '<li>' + escHtml_(t) + '</li>').join('') + '</ul>';
      continue;
    }
    // ย่อหน้าปกติ (รวมบรรทัดต่อเนื่องจนถึงบรรทัดว่าง)
    let para = [];
    while (i < lines.length && lines[i].trim() && !/^[-•]\s+/.test(lines[i].trim()) && !(urlRe.test(lines[i]) && !buttonUrl)) { para.push(escHtml_(lines[i].trim())); i++; }
    if (!para.length) { i++; continue; }
    html += '<p style="margin:0 0 14px;line-height:1.7">' + para.join('<br>') + '</p>';
  }
  const button = buttonUrl
    ? '<p style="text-align:center;margin:22px 0 8px"><a href="' + escHtml_(buttonUrl) + '" style="background:' + accent + ';color:#ffffff;text-decoration:none;padding:12px 28px;border-radius:6px;font-weight:600;display:inline-block">' + escHtml_(buttonText || 'เปิดลิงก์') + '</a></p>' : '';
  return '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#eef1f5;font-family:\'Noto Sans Thai\',Tahoma,Arial,sans-serif;color:#1b2a2e">' +
    '<table role="presentation" width="100%" style="background:#eef1f5;padding:24px 12px"><tr><td align="center">' +
    '<table role="presentation" width="600" style="max-width:600px;width:100%;background:#ffffff;border-radius:8px;overflow:hidden;border:1px solid #dde3ea">' +
    '<tr><td style="background:#122c4d;padding:20px 24px;text-align:center">' +
      (logo ? '<img src="' + escHtml_(logo) + '" alt="" width="56" height="56" style="display:block;margin:0 auto 8px;border-radius:50%;background:#fff">' : '') +
      '<div style="color:#ffffff;font-size:18px;font-weight:700">สภาวิชาชีพสังคมสงเคราะห์</div>' +
      '<div style="color:#b9c8dc;font-size:12px;margin-top:2px">Social Work Professions Council</div></td></tr>' +
    '<tr><td style="height:4px;background:' + accent + '"></td></tr>' +
    '<tr><td style="padding:28px 28px 20px">' +
      '<h2 style="margin:0 0 18px;font-size:19px;color:' + accent + '">' + escHtml_(subject) + '</h2>' + html + button + '</td></tr>' +
    '<tr><td style="background:#f5f6f8;padding:16px 24px;text-align:center;font-size:12px;color:#6b7278;line-height:1.7;border-top:1px solid #e3e8ee">' +
      'อีเมลฉบับนี้ส่งโดยอัตโนมัติจากแพลตฟอร์มบริการสมาชิกแบบเบ็ดเสร็จ กรุณาอย่าตอบกลับ<br>' +
      'สภาวิชาชีพสังคมสงเคราะห์ · 255 ถนนราชวิถี เขตราชเทวี กรุงเทพฯ 10400 · โทร 02 050 9427 · thaisocialworkcouncil@gmail.com</td></tr>' +
    '</table></td></tr></table></body></html>';
}

/** อ่านชีตสมาชิกกิตติมศักดิ์แบบไม่ล้มเมื่อชีตยังไม่ถูกสร้าง (ยังไม่ได้รัน setupAllSheets หลังอัปเดต) — กันรายงาน/แดชบอร์ดอื่นพังตามไปด้วย */
function readHonorarySafe_() {
  try { return readAll_('MembersHonorary'); } catch (e) { return []; }
}

/** เลขลำดับถัดไป (ตัวเลข 5 หลัก) — อ่านเลขสูงสุดจากคอลัมน์ที่ระบุ รองรับเลขรูปแบบเก่า เช่น MS2569-00001 / LC2568-00001 ด้วย
 * (ตามตารางรูปแบบเลขที่ TOR ข้อ 3.1: สามัญ/สมทบ/กิตติมศักดิ์/ใบอนุญาต = "ลำดับ", วิสามัญ = "ลำดับ/ปี พ.ศ.") */
function nextRunningSeq_(sheetName, col, filterFn) {
  let max = 0;
  readAll_(sheetName).forEach(r => {
    if (filterFn && !filterFn(r)) return;
    const m = String(r[col] || '').match(/^(?:[A-Z]{2}\d{4}-)?(\d{5})$/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return String(max + 1).padStart(5, '0');
}

/** เลขสมาชิกบุคคล (สามัญ/สมทบ) — เลขลำดับ 5 หลัก แยกชุดลำดับตามประเภทสมาชิก (สามัญเริ่ม 00001 / สมทบเริ่ม 00001 แยกกัน)
 * เมื่อเปลี่ยนประเภท สมทบ→สามัญ ระบบออกเลขใหม่จากชุดลำดับของสามัญให้ เพื่อไม่ให้เลขซ้ำกับสมาชิกสามัญคนอื่น */
function generateMemberNumber_(applicantType) {
  return nextRunningSeq_('Members', 'member_no', m => m.member_type === applicantType);
}

/** เลขทะเบียนสมาชิกวิสามัญ — รูปแบบ ลำดับ/ปี พ.ศ. เช่น 00001/2569 (ลำดับเริ่มใหม่ทุกปี พ.ศ.) */
function generateOrgRegNo_() {
  const year = new Date().getFullYear() + 543;
  let max = 0;
  readAll_('MembersOrg').forEach(o => {
    const v = String(o.reg_no || '');
    let m = v.match(/^(\d{5})\/(\d{4})$/);
    if (m && Number(m[2]) === year) max = Math.max(max, parseInt(m[1], 10));
    m = v.match(/^MO(\d{4})-(\d{5})$/); // รูปแบบเก่า
    if (m && Number(m[1]) === year) max = Math.max(max, parseInt(m[2], 10));
  });
  return String(max + 1).padStart(5, '0') + '/' + year;
}

/** เลขที่ใบอนุญาตประกอบวิชาชีพ (TOR ข้อ 3.2) — เลขลำดับ 5 หลัก ออกให้เฉพาะสมาชิกบุคคล (สามัญ/สมทบ) ไม่ใช่วิสามัญ */
function generateLicenseNumber_() {
  return nextRunningSeq_('Licenses', 'license_no');
}

/** บวกจำนวนปีเข้ากับวันที่ ISO แล้วคืนวันที่ ISO ใหม่ */
function addYearsIso_(isoDate, years) {
  const d = isoDate ? new Date(isoDate) : new Date();
  d.setFullYear(d.getFullYear() + years);
  return d.toISOString();
}

/** บวกจำนวนวัน (เป็นเลขทศนิยมได้) เข้ากับวันที่ ISO แล้วคืนวันที่ ISO ใหม่ — ใช้แทน addYearsIso_ เมื่อต้องการ offset แบบใกล้ๆ (หลัก วัน/สัปดาห์/เดือน)
 * เพราะ Date.setFullYear ปัดเศษปีทศนิยมทิ้ง ทำให้ addYearsIso_(x, 0.08) ได้วันที่เท่าเดิมโดยไม่ตั้งใจ */
function addDaysIso_(isoDate, days) {
  const d = isoDate ? new Date(isoDate) : new Date();
  d.setTime(d.getTime() + days * 24 * 3600 * 1000);
  return d.toISOString();
}

function daysUntil_(isoDate) {
  if (!isoDate) return null;
  const t = new Date(isoDate).getTime();
  if (isNaN(t)) return null; // ป้องกันค่าที่แปลงเป็นวันที่ไม่ได้ ทำให้เกิด NaN แล้ว eligibility check ผิดเพี้ยน
  const diffMs = t - Date.now();
  return Math.floor(diffMs / (24 * 3600 * 1000));
}

// ---------- 1. Members (TOR ข้อ 1) ----------

function action_listMembers(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  let rows = readAll_('Members');
  if (p.q) {
    const q = String(p.q).toLowerCase();
    rows = rows.filter(r =>
      String(r.national_id).includes(q) ||
      String(r.first_name_th).toLowerCase().includes(q) ||
      String(r.last_name_th).toLowerCase().includes(q) ||
      String(r.member_no).toLowerCase().includes(q) ||
      String(r.license_no).includes(q)
    );
  }
  return rows;
}

function action_listMembersOrg(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  let rows = readAll_('MembersOrg');
  if (p.q) {
    const q = String(p.q).toLowerCase();
    rows = rows.filter(r =>
      String(r.org_name_th).toLowerCase().includes(q) ||
      String(r.reg_no).toLowerCase().includes(q) ||
      String(r.rep_name).toLowerCase().includes(q)
    );
  }
  return rows;
}

function action_getMyProfile(p, user) {
  requireRole_(user, ['member']);
  if (!user.member_id) throw new Error('บัญชีนี้ยังไม่ผูกกับข้อมูลสมาชิก (รอใบสมัครอนุมัติและชำระเงิน)');
  const member = findRowById_('Members', 'member_id', user.member_id);
  if (member) {
    const license = readAll_('Licenses').find(l => l.member_id === user.member_id);
    return Object.assign({ profile_kind: 'individual' }, member, {
      license_issue_date: license ? license.issue_date : '',
      license_expire_date: license ? license.expire_date : ''
    });
  }
  const org = findRowById_('MembersOrg', 'org_id', user.member_id);
  if (org) return Object.assign({ profile_kind: 'organization' }, org);
  const honorary = findHonoraryById_(user.member_id);
  if (honorary) return Object.assign({ profile_kind: 'honorary' }, honorary);
  throw new Error('ไม่พบข้อมูลสมาชิกที่ผูกกับบัญชีนี้');
}

/** สมาชิกวิสามัญ (องค์กร) แก้ไขข้อมูลติดต่อ/ที่อยู่ของตนเองได้บางส่วน หรือแอดมินแก้ไขได้ทุกฟิลด์
 * ฟิลด์เกี่ยวกับตัวตนตามกฎหมาย (ชื่อองค์กร, เลขทะเบียนนิติบุคคล, ข้อมูลผู้แทนที่เป็นตัวตน ฯลฯ) แก้ได้เฉพาะแอดมินเท่านั้น
 * เหมือนกับกรณีชื่อ-นามสกุลสมาชิกบุคคลที่ต้องผ่านเจ้าหน้าที่ (TOR ข้อ 1.1(5)) — ที่นี่ตัดสินใจให้แอดมินแก้ตรงได้เลยไม่ต้องผ่านคำขออนุมัติ
 * เพราะข้อมูลองค์กรเปลี่ยนไม่บ่อยเท่าและมักมีเอกสารทางการยืนยันมาให้เจ้าหน้าที่ตรวจสอบก่อนแก้อยู่แล้วนอกระบบ
 */
const ORG_SELF_EDITABLE_FIELDS = [
  'rep_phone', 'rep_phone2', 'rep_email', 'rep_line_id',
  'addr_house_no', 'addr_building', 'addr_floor', 'addr_moo', 'addr_soi', 'addr_road',
  'addr_subdistrict', 'addr_district', 'addr_province', 'addr_zipcode',
  'office_phone', 'website', 'logo_url'
];

function action_updateOrgProfile(p, user) {
  const isAdmin = ['admin_super', 'admin_staff'].indexOf(user.role) !== -1;
  const orgId = isAdmin ? p.org_id : user.member_id;
  if (!orgId) throw new Error('ไม่พบ org_id');
  const before = findRowById_('MembersOrg', 'org_id', orgId);
  if (!before) throw new Error('ไม่พบข้อมูลสมาชิกวิสามัญ');

  let patch = p.patch || {};
  if (!isAdmin) {
    // สมาชิกแก้ไขเอง — จำกัดเฉพาะฟิลด์ที่อนุญาต กันแก้ข้อมูลตัวตน/ทางกฎหมายเอง
    const filtered = {};
    ORG_SELF_EDITABLE_FIELDS.forEach(f => { if (patch[f] !== undefined) filtered[f] = patch[f]; });
    patch = filtered;
  }
  patch.updated_at = nowIso_();
  updateRow_('MembersOrg', before._row, patch);
  auditLog_('MembersOrg', orgId, isAdmin ? 'admin_update' : 'self_update', before, patch, user.user_id);
  return { ok: true };
}

function action_updateMember(p, user) {
  // สมาชิกแก้ไขข้อมูลตนเอง หรือแอดมินแก้ไขให้ — ฟิลด์ที่แก้ชื่อ-นามสกุลต้องผ่านอนุมัติ (TOR ข้อ 1.1 (5))
  const isAdmin = ['admin_super', 'admin_staff'].indexOf(user.role) !== -1;
  const memberId = isAdmin ? p.member_id : user.member_id;
  if (!memberId) throw new Error('ไม่พบ member_id');
  const before = findRowById_('Members', 'member_id', memberId);
  if (!before) throw new Error('ไม่พบสมาชิก');

  // ฟิลด์ชื่อ-นามสกุลภาษาไทย ยังต้องผ่านอนุมัติเจ้าหน้าที่เสมอ (TOR ข้อ 1.1(5) — ต้องตรงกับบัตรประชาชน)
  // ส่วนชื่อ-นามสกุล/คำนำหน้าภาษาอังกฤษ ให้สมาชิกสามัญ/สมทบแก้ไขเองได้ทันที ไม่ต้องรออนุมัติ (ไม่ใช่ข้อมูลอ้างอิงตามกฎหมายเหมือนชื่อไทย)
  const nameFields = ['first_name_th', 'last_name_th', 'prefix_th'];
  const wantsNameChange = nameFields.some(f => p.patch[f] !== undefined && p.patch[f] !== before[f]);

  if (wantsNameChange && !isAdmin) {
    // สมาชิกขอเปลี่ยนชื่อ-นามสกุล ต้องแนบเอกสารหลักฐานประกอบทุกครั้ง (เช่น หนังสือรับรองการเปลี่ยนชื่อ) และสร้างเป็นคำขอรออนุมัติ ไม่แก้ไขทันที (TOR ข้อ 1.1(5))
    if (!p.documents || !p.documents.length) {
      throw new Error('กรุณาแนบเอกสารหลักฐานประกอบการขอเปลี่ยนชื่อ-นามสกุล (เช่น หนังสือรับรองการเปลี่ยนชื่อ/ใบสำคัญที่เกี่ยวข้อง) ก่อนส่งคำขอ');
    }
    const pendingExisting = readAll_('Applications').find(a =>
      a.applicant_type === 'name_change' && a.applicant_user_id === user.user_id &&
      ['รอตรวจสอบ', 'รอเอกสารเพิ่มเติม'].indexOf(a.status) !== -1
    );
    if (pendingExisting) throw new Error('มีคำขอเปลี่ยนชื่อ-นามสกุลที่ยังไม่เสร็จสิ้นอยู่แล้ว (เลขที่ ' + pendingExisting.application_id + ') กรุณารอผลการพิจารณาก่อนยื่นคำขอใหม่');

    const oldName = { prefix_th: before.prefix_th || '', first_name_th: before.first_name_th || '', last_name_th: before.last_name_th || '' };
    const app = {
      application_id: newId_('NAMECHG'),
      applicant_type: 'name_change',
      applicant_user_id: user.user_id,
      ref_data_json: JSON.stringify({ member_id: memberId, patch: p.patch, old_name: oldName, reason: p.reason || '' }),
      documents_json: JSON.stringify(p.documents || []),
      status: 'รอตรวจสอบ',
      reject_reason: '', fee_amount: 0,
      submitted_at: nowIso_(), reviewed_by: '', reviewed_at: '', approved_member_id: ''
    };
    appendRow_('Applications', app);
    const immediatePatch = {};
    Object.keys(p.patch).forEach(k => { if (nameFields.indexOf(k) === -1) immediatePatch[k] = p.patch[k]; });
    if (Object.keys(immediatePatch).length) {
      immediatePatch.updated_at = nowIso_();
      updateRow_('Members', before._row, immediatePatch);
      auditLog_('Members', memberId, 'self_update_partial', before, immediatePatch, user.user_id);
    }
    mailNotify_(before.email, 'ได้รับคำขอเปลี่ยนชื่อ-นามสกุลแล้ว',
      'ระบบได้รับคำขอเปลี่ยนชื่อ-นามสกุลของท่านแล้ว เลขที่คำขอ ' + app.application_id +
      '\nจาก "' + [oldName.prefix_th, oldName.first_name_th, oldName.last_name_th].filter(Boolean).join(' ') + '"' +
      ' เป็น "' + [p.patch.prefix_th || oldName.prefix_th, p.patch.first_name_th || oldName.first_name_th, p.patch.last_name_th || oldName.last_name_th].filter(Boolean).join(' ') + '"' +
      '\nสถานะปัจจุบัน: รอตรวจสอบ\nเจ้าหน้าที่จะตรวจสอบเอกสารและแจ้งผลกลับทางอีเมลนี้ (ท่านสามารถตรวจสอบสถานะได้แบบเรียลไทม์ในเมนู "ขอเปลี่ยนชื่อ-นามสกุล")');
    return { ok: true, nameChangePending: true, application_id: app.application_id };
  }

  const patch = Object.assign({}, p.patch, { updated_at: nowIso_() });
  updateRow_('Members', before._row, patch);
  auditLog_('Members', memberId, isAdmin ? 'admin_update' : 'self_update', before, patch, user.user_id);
  return { ok: true, nameChangePending: false };
}

// ---------- ระบบเปลี่ยนประเภทสมาชิก (TOR ข้อ 1.1(7)) — แบ่ง 2 ฝั่ง ----------
// Workflow เดียวกับระบบสมัครสมาชิก: รอตรวจสอบ -> (อนุมัติ) รอชำระเงิน -> (ตรวจสลิปผ่าน) เปลี่ยนประเภทสำเร็จ
// ใช้ action_reviewApplication และ action_submitPaymentSlip ร่วมกับระบบสมัครสมาชิกได้เลย (โครงสร้างเดียวกัน)
// เฉพาะขั้นตอนสุดท้าย (ยืนยันชำระเงิน) ที่ต้องแยก เพราะต้อง "ย้าย" สมาชิกที่มีอยู่แล้ว ไม่ใช่สร้างสมาชิกใหม่

/** 7.1 ฝั่งสมาชิก — ยื่นคำขอเปลี่ยนประเภทระหว่างสามัญ/สมทบ พร้อมแนบเอกสาร (วุฒิการศึกษา, ใบรับรองแพทย์) */
function action_submitTypeChange(p, user) {
  requireRole_(user, ['member']);
  if (!user.member_id) throw new Error('บัญชีนี้ยังไม่ผูกกับข้อมูลสมาชิก');
  const member = findRowById_('Members', 'member_id', user.member_id);
  if (!member) throw new Error('ไม่พบข้อมูลสมาชิก');

  // รองรับเฉพาะทิศทาง สมทบ -> สามัญ เท่านั้น (ตามระเบียบของสภาฯ)
  if (member.member_type !== 'สมทบ') {
    throw new Error('ระบบรองรับเฉพาะการเปลี่ยนจากสมาชิกสมทบเป็นสมาชิกสามัญเท่านั้น บัญชีนี้เป็นสมาชิก' + member.member_type + 'อยู่แล้ว');
  }
  const toType = 'สามัญ';
  const duration = p.duration || '1';
  const hasSignature = (p.documents || []).some(d => (d.doc_type || '').indexOf('ลายมือชื่อ') !== -1);
  if (!hasSignature) throw new Error('กรุณาลงลายมือชื่อผู้ยื่นคำขอก่อนส่งคำขอเปลี่ยนประเภทสมาชิก');
  // ไม่มีค่าธรรมเนียมเปลี่ยนประเภทแยกต่างหาก คิดเป็นค่าบำรุงสมาชิกสามัญตามระยะเวลาที่เลือกเลย (ไม่เก็บค่าแรกเข้าซ้ำ เพราะเป็นสมาชิกอยู่แล้ว)
  const fee = getMembershipFee_(toType, duration, member.dob);

  const pendingExisting = readAll_('Applications').find(a =>
    a.applicant_type === 'เปลี่ยนประเภท' && a.applicant_user_id === user.user_id &&
    ['รอตรวจสอบ', 'รอชำระเงิน'].indexOf(a.status) !== -1
  );
  if (pendingExisting) throw new Error('มีคำขอเปลี่ยนประเภทที่ยังไม่เสร็จสิ้นอยู่แล้ว (เลขที่ ' + pendingExisting.application_id + ')');

  const app = {
    application_id: newId_('TC'),
    applicant_type: 'เปลี่ยนประเภท',
    applicant_user_id: user.user_id,
    ref_data_json: JSON.stringify({
      member_id: user.member_id, from_type: member.member_type, to_type: toType, duration: duration,
      email: member.email, name: [member.prefix_th, member.first_name_th, member.last_name_th].filter(Boolean).join(' ')
    }),
    documents_json: JSON.stringify(p.documents || []),
    status: 'รอตรวจสอบ', reject_reason: '', fee_amount: fee,
    submitted_at: nowIso_(), reviewed_by: '', reviewed_at: '', approved_member_id: ''
  };
  appendRow_('Applications', app);

  mailNotify_(member.email, 'ได้รับคำขอเปลี่ยนประเภทสมาชิกแล้ว',
    'ระบบได้รับคำขอเปลี่ยนประเภทจาก "' + member.member_type + '" เป็น "' + toType + '" แล้ว เลขที่คำขอ ' + app.application_id +
    '\nสถานะปัจจุบัน: รอตรวจสอบ\nเจ้าหน้าที่จะตรวจสอบคุณสมบัติและเอกสารแล้วแจ้งผลกลับทางอีเมลนี้');

  return { application_id: app.application_id };
}

/** 7.2 ฝั่งแอดมิน — ยืนยันชำระเงินแล้วเปลี่ยนประเภท+ย้ายข้อมูลอัตโนมัติ พร้อมบันทึกประวัติ */
function action_verifyTypeChangePayment(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const payment = findRowById_('Payments', 'payment_id', p.payment_id);
  if (!payment || payment.ref_type !== 'application') throw new Error('ไม่พบรายการชำระเงิน');
  const app = findRowById_('Applications', 'application_id', payment.ref_id);
  if (!app || app.applicant_type !== 'เปลี่ยนประเภท') throw new Error('ไม่พบคำขอเปลี่ยนประเภทที่เกี่ยวข้อง');

  if (p.decision === 'reject') {
    updateRow_('Payments', payment._row, { status: 'ปฏิเสธ', verified_by: user.user_id, verified_at: nowIso_() });
    auditLog_('Payments', payment.payment_id, 'payment_rejected', payment, { reason: p.reason }, user.user_id);
    return { ok: true };
  }

  const refData = JSON.parse(app.ref_data_json || '{}');
  const member = findRowById_('Members', 'member_id', refData.member_id);
  if (!member) throw new Error('ไม่พบสมาชิกที่เกี่ยวข้อง (อาจถูกลบหรือแก้ไขไปแล้ว)');

  // ค่าธรรมเนียมเปลี่ยนประเภทคิดเป็นค่าบำรุงสมาชิกสามัญตามระยะเวลาที่เลือก จึงต่ออายุสมาชิกภาพไปพร้อมกันด้วย (เหมือนการต่ออายุ)
  const baseDate = member.membership_expire_date && new Date(member.membership_expire_date) > new Date() ? member.membership_expire_date : nowIso_();
  const newExpire = computeMembershipExpiry_(baseDate, refData.duration || '1');
  const oldMemberNo = member.member_no;
  const newMemberNo = generateMemberNumber_(refData.to_type);
  updateRow_('Members', member._row, { member_type: refData.to_type, member_no: newMemberNo, membership_expire_date: newExpire, updated_at: nowIso_() });

  updateRow_('Payments', payment._row, {
    status: 'ชำระแล้ว', verified_by: user.user_id, verified_at: nowIso_(), paid_at: nowIso_()
  });
  try { issueReceiptForPayment_(payment.payment_id, user.user_id); } catch (e) { /* ไม่ให้ล้มเหลวทั้งฟังก์ชันถ้าออกใบเสร็จไม่สำเร็จ */ }
  refreshMemberCard_(member.member_id, newExpire, refData.to_type); // ออก QR ใหม่เพราะวันหมดอายุเปลี่ยน (TOR ข้อ 5.5)
  updateRow_('Applications', app._row, { status: 'เปลี่ยนประเภทสำเร็จ', approved_member_id: member.member_id });

  auditLog_('Members', member.member_id, 'type_change_confirmed',
    { member_type: refData.from_type, member_no: oldMemberNo }, { member_type: refData.to_type, member_no: newMemberNo }, user.user_id);

  mailNotify_(member.email, 'เปลี่ยนประเภทสมาชิกสำเร็จ',
    'สมาชิกภาพของท่านเปลี่ยนจาก "' + refData.from_type + '" เป็น "' + refData.to_type + '" เรียบร้อยแล้ว\n' +
    'สมาชิกภาพมีอายุถึงวันที่: ' + new Date(newExpire).toLocaleDateString('th-TH') + '\n' +
    'มีผลทันที สามารถตรวจสอบข้อมูลได้ที่ Member Portal');

  return { ok: true, member_id: member.member_id, new_type: refData.to_type };
}

/** รายการคำขอเปลี่ยนประเภททั้งหมด (ฝั่งแอดมิน) */
/** แนบรูปสมาชิกให้คำขอที่ผูกกับสมาชิกที่มีอยู่แล้ว (member_id ใน ref_data_json) — ใช้ร่วมกันทั้งคำขอเปลี่ยนประเภทและเปลี่ยนชื่อ-นามสกุล */
function attachMemberPhoto_(applications) {
  const members = readAll_('Members');
  return applications.map(a => {
    let memberId = '';
    try { memberId = JSON.parse(a.ref_data_json || '{}').member_id || ''; } catch (e) {}
    const m = memberId ? members.find(x => x.member_id === memberId) : null;
    return Object.assign({}, a, { member_photo_url: m ? (m.photo_url || '') : '' });
  });
}

function action_listTypeChangeRequests(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const rows = readAll_('Applications')
    .filter(a => a.applicant_type === 'เปลี่ยนประเภท')
    .filter(a => !p.status || a.status === p.status)
    .sort((a, b) => new Date(b.submitted_at) - new Date(a.submitted_at));
  return attachMemberPhoto_(rows);
}

// ---------- 2. ระบบสมัครสมาชิก — Membership Application System (TOR ข้อ 2) ----------
// ลำดับสถานะ: รอตรวจสอบ -> (อนุมัติ) รอชำระเงิน -> (ตรวจสลิปผ่าน) เป็นสมาชิกแล้ว
//                      \-> (ไม่อนุมัติ) ไม่อนุมัติ
// หมายเหตุ: TOR ระบุสถานะ "รออนุมัติ" แยกจาก "รอตรวจสอบ" ไว้ด้วย ซึ่งหมายถึงมีผู้อนุมัติหลายระดับ (เจ้าหน้าที่/ผู้บริหาร/กรรมการ)
// scaffold นี้ยังรวมเป็นขั้นตอนอนุมัติเดียว (เจ้าหน้าที่คนเดียว) — การอนุมัติหลายระดับเป็นงานต่อยอด ดู ROADMAP.md

function action_submitApplication(p, user) {
  requireRole_(user, ['member']);
  requirePdpaConsent_(user); // PDPA: ต้องมีความยินยอมฉบับปัจจุบันก่อนส่งข้อมูลส่วนบุคคล
  const applicantType = p.applicant_type;
  if (['สามัญ', 'สมทบ', 'วิสามัญ'].indexOf(applicantType) === -1) {
    throw new Error('ประเภทการสมัครไม่ถูกต้อง');
  }
  // แยกบัญชีองค์กรกับบุคคล: สมาชิกวิสามัญต้องสมัครด้วยบัญชีที่ตั้งไว้สำหรับองค์กรโดยเฉพาะ และยังไม่ผูกกับสมาชิกรายอื่น
  const acctType = accountTypeOf_(user);
  if (applicantType === 'วิสามัญ') {
    if (user.member_id) {
      throw new Error('บัญชีนี้ผูกกับสมาชิกอยู่แล้ว (' + (acctType === 'org' ? 'องค์กร' : 'บุคคล') + ') — การสมัครสมาชิกวิสามัญ (องค์กร) ต้องใช้บัญชีผู้ใช้งานแยกสำหรับองค์กร กรุณาสร้างบัญชีองค์กรใหม่');
    }
    if (user.account_type === 'person') {
      throw new Error('บัญชีนี้เป็นบัญชีบุคคล — การสมัครสมาชิกวิสามัญ (องค์กร) ต้องใช้บัญชีที่สร้างสำหรับองค์กรโดยเฉพาะ กรุณาสร้างบัญชีองค์กรใหม่ (ตั้งชื่อผู้ใช้งานแยก)');
    }
  } else if (acctType === 'org') {
    throw new Error('บัญชีนี้เป็นบัญชีองค์กร ไม่สามารถสมัครเป็นสมาชิกบุคคลได้ กรุณาใช้บัญชีบุคคล');
  }
  const hasSignature = (p.documents || []).some(d => (d.doc_type || '').indexOf('ลายมือชื่อ') !== -1);
  if (!hasSignature) throw new Error('กรุณาลงลายมือชื่อผู้สมัคร/ผู้ยื่นคำขอก่อนส่งใบสมัคร');
  const duration = p.duration || '1';
  const dob = p.data && p.data.dob;
  const membershipFee = getMembershipFee_(applicantType, duration, dob);
  // ค่าสมัครครั้งแรก = ค่าบำรุงสมาชิกตามระยะเวลาที่เลือก + ค่าสมาชิกครั้งแรกเข้า (ทุกประเภทรวมวิสามัญ) + ค่าจัดส่งหนังสือรับรองฯ (เฉพาะวิสามัญ ตามแบบฟอร์ม สวชส.10)
  const fee = membershipFee + MEMBERSHIP_ENTRANCE_FEE + (applicantType === 'วิสามัญ' ? CERTIFICATE_MAILING_FEE : 0);
  const dataWithDuration = Object.assign({}, p.data || {}, { duration: duration });
  const app = {
    application_id: newId_('APP'),
    applicant_type: applicantType,
    applicant_user_id: user.user_id,
    ref_data_json: JSON.stringify(dataWithDuration),
    documents_json: JSON.stringify(p.documents || []),
    status: 'รอตรวจสอบ',
    reject_reason: '', fee_amount: fee,
    submitted_at: nowIso_(), reviewed_by: '', reviewed_at: '', approved_member_id: ''
  };
  appendRow_('Applications', app);

  const email = (p.data && (p.data.email || p.data.rep_email)) || '';
  mailNotify_(email, 'ได้รับใบสมัครสมาชิกสภาวิชาชีพสังคมสงเคราะห์แล้ว',
    'ระบบได้รับใบสมัครประเภท "' + applicantType + '" ของท่านแล้ว เลขที่คำขอ ' + app.application_id +
    '\nสถานะปัจจุบัน: รอตรวจสอบ\nเจ้าหน้าที่จะตรวจสอบและแจ้งผลกลับทางอีเมลนี้ภายใน 15 วันทำการ');

  return { application_id: app.application_id };
}

/** สมาชิกดูรายการคำขอของตนเองแบบเรียลไทม์ (TOR ข้อ 2.2.1) */
function action_myApplications(p, user) {
  requireRole_(user, ['member']);
  return readAll_('Applications')
    .filter(a => a.applicant_user_id === user.user_id)
    .sort((a, b) => new Date(b.submitted_at) - new Date(a.submitted_at));
}

function action_listApplications(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const EXCLUDED_TYPES = ['name_change', 'เปลี่ยนประเภท', 'ขึ้นทะเบียน', 'ต่ออายุใบอนุญาต', 'ต่ออายุสมาชิก'].concat(EXAM_APPLICANT_TYPES);
  let rows = readAll_('Applications').filter(a => EXCLUDED_TYPES.indexOf(a.applicant_type) === -1);
  if (p.status) rows = rows.filter(r => r.status === p.status);
  if (p.applicant_type) rows = rows.filter(r => r.applicant_type === p.applicant_type);
  return rows;
}

function action_reviewApplication(p, user) {
  // p.decision: 'approve' | 'reject'
  requireRole_(user, ['admin_super', 'admin_staff']);
  const app = findRowById_('Applications', 'application_id', p.application_id);
  if (!app) throw new Error('ไม่พบใบสมัคร');
  const data = JSON.parse(app.ref_data_json || '{}');
  const email = data.email || data.rep_email || '';
  const APPLICATION_LABELS = {
    'เปลี่ยนประเภท': 'คำขอเปลี่ยนประเภทสมาชิก',
    'ขึ้นทะเบียน': 'คำขอขึ้นทะเบียนรับใบอนุญาต',
    'ต่ออายุใบอนุญาต': 'คำขอต่ออายุใบอนุญาต'
  };
  const isTypeChange = app.applicant_type === 'เปลี่ยนประเภท';
  const label = APPLICATION_LABELS[app.applicant_type] || 'ใบสมัครสมาชิก';

  if (p.decision === 'reject') {
    updateRow_('Applications', app._row, {
      status: 'ไม่อนุมัติ', reject_reason: p.reason || '',
      reviewed_by: user.user_id, reviewed_at: nowIso_()
    });
    auditLog_('Applications', app.application_id, 'reject', app, { reason: p.reason }, user.user_id);
    mailNotify_(email, 'ผลการพิจารณา' + label,
      label + 'เลขที่ ' + app.application_id + ' ไม่ได้รับการอนุมัติ\nเหตุผล: ' + (p.reason || '-'));
    return { ok: true };
  }

  // approve: ยังไม่ดำเนินการทันที — ต้องรอชำระเงินก่อน (TOR ข้อ 2.2.1 / ข้อ 1.1(7))
  updateRow_('Applications', app._row, {
    status: 'รอชำระเงิน', reviewed_by: user.user_id, reviewed_at: nowIso_()
  });

  // fee_amount ถูกกำหนดไว้แล้วตั้งแต่ตอนยื่นคำขอทุกประเภท (รวมค่าธรรมเนียมตามระยะเวลาที่เลือกแล้ว) — ใช้ค่านี้เป็นหลัก
  const feeAmount = app.fee_amount || 0;
  const existingPayment = readAll_('Payments').find(pm => pm.ref_type === 'application' && pm.ref_id === app.application_id);
  if (!existingPayment) {
    appendRow_('Payments', {
      payment_id: newId_('PAY'), ref_type: 'application', ref_id: app.application_id,
      member_id: isTypeChange ? (data.member_id || '') : '', amount: feeAmount,
      slip_url: '', method: '', status: 'รอชำระ',
      verified_by: '', verified_at: '', paid_at: ''
    });
  }

  auditLog_('Applications', app.application_id, 'approve_pending_payment', app, {}, user.user_id);
  mailNotify_(email, label + 'ได้รับการอนุมัติ — กรุณาชำระเงิน',
    label + 'เลขที่ ' + app.application_id + ' ได้รับการอนุมัติแล้ว\nกรุณาชำระค่าธรรมเนียม ' + feeAmount +
    ' บาท ผ่าน Member Portal แล้วแนบสลิปการโอนเงิน');
  return { ok: true, status: 'รอชำระเงิน' };
}

/** สมาชิกแจ้งชำระเงิน แนบสลิป (TOR ข้อ 2.2.1) */
function action_submitPaymentSlip(p, user) {
  requireRole_(user, ['member']);
  const app = findRowById_('Applications', 'application_id', p.application_id);
  if (!app) throw new Error('ไม่พบใบสมัคร');
  if (app.applicant_user_id !== user.user_id) throw new Error('ไม่มีสิทธิ์ดำเนินการกับใบสมัครนี้');
  if (app.status !== 'รอชำระเงิน') throw new Error('ใบสมัครนี้ไม่ได้อยู่ในสถานะรอชำระเงิน');

  const payment = readAll_('Payments').find(pm => pm.ref_type === 'application' && pm.ref_id === app.application_id);
  if (!payment) throw new Error('ไม่พบรายการชำระเงินของใบสมัครนี้');

  updateRow_('Payments', payment._row, {
    slip_url: p.slip_url || '', method: p.method || 'โอนเงิน', channel: p.channel || 'โอนเงินผ่านธนาคาร (ตรวจสอบสลิปด้วยมือ)', status: 'รอตรวจสอบ'
  });
  auditLog_('Payments', payment.payment_id, 'slip_submitted', payment, { slip_url: p.slip_url }, user.user_id);
  return { ok: true };
}

/** แอดมินตรวจสอบสลิป — เมื่อผ่านแล้วออกเลขสมาชิก/เลขทะเบียนวิสามัญอัตโนมัติ และสร้างข้อมูลสมาชิกจริง (TOR ข้อ 2.2.2) */
function action_verifyApplicationPayment(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const payment = findRowById_('Payments', 'payment_id', p.payment_id);
  if (!payment || payment.ref_type !== 'application') throw new Error('ไม่พบรายการชำระเงิน');
  const app = findRowById_('Applications', 'application_id', payment.ref_id);
  if (!app) throw new Error('ไม่พบใบสมัครที่เกี่ยวข้อง');
  if (app.applicant_type === 'เปลี่ยนประเภท') throw new Error('รายการนี้เป็นคำขอเปลี่ยนประเภทสมาชิก กรุณาใช้เมนู "คำขอเปลี่ยนประเภทสมาชิก" แทน');
  if (app.applicant_type === 'name_change') throw new Error('รายการนี้เป็นคำขอเปลี่ยนชื่อ ไม่ใช่การชำระเงินสมัครสมาชิก');
  if (app.applicant_type === 'ต่ออายุสมาชิก') throw new Error('รายการนี้เป็นคำขอต่ออายุสมาชิก กรุณาใช้เมนู "การชำระเงิน" (ตรวจสอบการต่ออายุสมาชิก) แทน');
  if (app.applicant_type === 'ขึ้นทะเบียน' || app.applicant_type === 'ต่ออายุใบอนุญาต') throw new Error('รายการนี้เป็นคำขอเกี่ยวกับใบอนุญาต กรุณาใช้เมนู "คำขอใบอนุญาต" แทน');
  if (EXAM_APPLICANT_TYPES.indexOf(app.applicant_type) !== -1) throw new Error('รายการนี้เป็นค่าสมัครสอบ กรุณาใช้เมนู "คำขอสมัครสอบ" แทน');

  if (p.decision === 'reject') {
    updateRow_('Payments', payment._row, { status: 'ปฏิเสธ', verified_by: user.user_id, verified_at: nowIso_() });
    auditLog_('Payments', payment.payment_id, 'payment_rejected', payment, { reason: p.reason }, user.user_id);
    return { ok: true };
  }

  const data = JSON.parse(app.ref_data_json || '{}');
  const documents = app.documents_json || '[]';
  const expireDate = computeMembershipExpiry_(nowIso_(), data.duration || '1');
  let newId, email, issuedNo;

  if (app.applicant_type === 'วิสามัญ') {
    newId = newId_('ORG');
    issuedNo = generateOrgRegNo_();
    appendRow_('MembersOrg', Object.assign({}, data, {
      org_id: newId, reg_no: issuedNo, applied_date: app.submitted_at, approved_date: nowIso_(),
      status: 'ปกติ', end_date: expireDate, documents_json: documents, created_at: nowIso_(), updated_at: nowIso_(),
      purposes_json: JSON.stringify(data.purposes || [])
    }));
    email = data.rep_email || '';
  } else {
    newId = newId_('MBR');
    issuedNo = generateMemberNumber_(app.applicant_type);
    appendRow_('Members', Object.assign({}, data, {
      member_id: newId, member_no: issuedNo, member_type: app.applicant_type, member_status: 'ปกติ',
      applied_date: app.submitted_at, membership_expire_date: expireDate,
      license_no: '', license_status: '', documents_json: documents,
      created_at: nowIso_(), updated_at: nowIso_()
    }));
    // หมายเหตุ: ไม่ออกใบอนุญาตอัตโนมัติที่ขั้นตอนนี้ — การขึ้นทะเบียนรับใบอนุญาตเป็นคนละขั้นตอนตาม TOR ข้อ 4.1
    // (เฉพาะสมาชิกสามัญที่สถานะปกติเท่านั้นที่ยื่นขอขึ้นทะเบียนได้ ผ่านเมนู "ขึ้นทะเบียนใบอนุญาต" ใน Member Portal)
    email = data.email || '';
  }

  const applicant = findRowById_('Users', 'user_id', app.applicant_user_id);
  let accountWarning = '';
  if (applicant) {
    if (applicant.member_id && applicant.member_id !== newId) {
      // บัญชีนี้ผูกกับสมาชิกรายอื่นอยู่แล้ว (เช่น ตัวแทนองค์กรเป็นสมาชิกบุคคลอยู่) — ห้ามทับ เพื่อไม่ให้สมาชิกบุคคลเสียสิทธิ์เข้าใช้งาน
      accountWarning = 'บัญชีผู้สมัคร (' + applicant.username + ') ผูกกับสมาชิกรายอื่นอยู่แล้ว จึงไม่ได้ผูกกับ' + (app.applicant_type === 'วิสามัญ' ? 'สมาชิกวิสามัญ' : 'สมาชิกใหม่') + ' — ต้องให้ผู้สมัครสร้างบัญชีองค์กรแยกและให้ผู้ดูแลระบบผูกบัญชีให้';
      auditLog_('Users', applicant.user_id, 'bind_skipped_already_bound', { member_id: applicant.member_id }, { new_id: newId }, user.user_id);
    } else {
      updateRow_('Users', applicant._row, { member_id: newId, account_type: app.applicant_type === 'วิสามัญ' ? 'org' : 'person' });
    }
  }

  refreshMemberCard_(newId, app.applicant_type === 'วิสามัญ' ? expireDate : expireDate, app.applicant_type); // ออกบัตร/QR ครั้งแรก (TOR ข้อ 5)

  updateRow_('Payments', payment._row, {
    status: 'ชำระแล้ว', verified_by: user.user_id, verified_at: nowIso_(), paid_at: nowIso_()
  });
  try { issueReceiptForPayment_(payment.payment_id, user.user_id); } catch (e) { /* ไม่ให้ล้มเหลวทั้งฟังก์ชันถ้าออกใบเสร็จไม่สำเร็จ */ }
  updateRow_('Applications', app._row, { status: 'เป็นสมาชิกแล้ว', approved_member_id: newId });

  auditLog_('Applications', app.application_id, 'membership_confirmed', app, { id: newId }, user.user_id);
  mailNotify_(email, 'ยินดีต้อนรับสู่การเป็นสมาชิกสภาวิชาชีพสังคมสงเคราะห์',
    'การชำระเงินของท่านได้รับการยืนยันแล้ว ท่านเป็นสมาชิกเรียบร้อย\n' +
    (app.applicant_type === 'วิสามัญ' ? 'เลขทะเบียนสมาชิกวิสามัญ' : 'เลขสมาชิก') + ': ' + issuedNo +
    (app.applicant_type !== 'วิสามัญ' ? '\nสมาชิกภาพมีอายุถึงวันที่: ' + new Date(expireDate).toLocaleDateString('th-TH') : '') +
    (app.applicant_type === 'สามัญ' ? '\nท่านสามารถยื่นขอขึ้นทะเบียนรับใบอนุญาตประกอบวิชาชีพได้แล้วผ่านเมนู "ขึ้นทะเบียนใบอนุญาต" ใน Member Portal' : '') +
    '\nสามารถเข้าสู่ระบบ Member Portal เพื่อดูข้อมูลและบัตรสมาชิกได้ทันที');

  return { ok: true, id: newId, warning: accountWarning || undefined };
}

/** ฝั่งแอดมิน — พิจารณาคำขอเปลี่ยนชื่อ-นามสกุล: อนุมัติ / ไม่อนุมัติ (ระบุเหตุผล) / ขอเอกสารเพิ่มเติม (ระบุข้อความ) */
function action_reviewNameChange(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const app = findRowById_('Applications', 'application_id', p.application_id);
  if (!app || app.applicant_type !== 'name_change') throw new Error('ไม่พบคำขอเปลี่ยนชื่อนี้');
  const refData = JSON.parse(app.ref_data_json || '{}');

  const refMember = findRowById_('Members', 'member_id', refData.member_id);
  if (p.decision === 'reject') {
    if (!p.reason) throw new Error('กรุณาระบุเหตุผลที่ไม่อนุมัติ');
    updateRow_('Applications', app._row, {
      status: 'ไม่อนุมัติ', reject_reason: p.reason || '', reviewed_by: user.user_id, reviewed_at: nowIso_()
    });
    mailNotify_(refMember ? refMember.email : '', 'คำขอเปลี่ยนชื่อ-นามสกุลไม่ได้รับการอนุมัติ',
      'คำขอเลขที่ ' + app.application_id + ' ไม่ได้รับการอนุมัติ\nเหตุผล: ' + p.reason);
    auditLog_('Applications', app.application_id, 'name_change_rejected', app, { reason: p.reason }, user.user_id);
    return { ok: true };
  }
  if (p.decision === 'request_more_docs') {
    if (!p.reason) throw new Error('กรุณาระบุข้อความ/รายการเอกสารที่ต้องการเพิ่มเติม');
    updateRow_('Applications', app._row, {
      status: 'รอเอกสารเพิ่มเติม', reject_reason: p.reason || '', reviewed_by: user.user_id, reviewed_at: nowIso_()
    });
    mailNotify_(refMember ? refMember.email : '', 'คำขอเปลี่ยนชื่อ-นามสกุล — เจ้าหน้าที่ขอเอกสารเพิ่มเติม',
      'คำขอเลขที่ ' + app.application_id + ' เจ้าหน้าที่ขอเอกสารเพิ่มเติม:\n' + p.reason +
      '\nกรุณาแนบเอกสารเพิ่มเติมผ่านเมนู "ขอเปลี่ยนชื่อ-นามสกุล" ในระบบ');
    auditLog_('Applications', app.application_id, 'name_change_more_docs_requested', app, { reason: p.reason }, user.user_id);
    return { ok: true };
  }
  if (!refMember) throw new Error('ไม่พบสมาชิก');
  const patch = Object.assign({}, refData.patch, { updated_at: nowIso_() });
  updateRow_('Members', refMember._row, patch);
  updateRow_('Applications', app._row, { status: 'สำเร็จ', reject_reason: '', reviewed_by: user.user_id, reviewed_at: nowIso_() });
  auditLog_('Members', refData.member_id, 'name_change_approved', refMember, patch, user.user_id);
  mailNotify_(refMember.email, 'คำขอเปลี่ยนชื่อ-นามสกุลได้รับการอนุมัติแล้ว',
    'คำขอเลขที่ ' + app.application_id + ' ได้รับการอนุมัติแล้ว ชื่อ-นามสกุลใหม่ของท่านมีผลทันที');
  return { ok: true };
}

function action_listNameChangeRequests(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const rows = readAll_('Applications').filter(a => a.applicant_type === 'name_change' && (!p.status || a.status === p.status));
  return attachMemberPhoto_(rows);
}

/** ฝั่งสมาชิก — แนบเอกสารเพิ่มเติมตามที่เจ้าหน้าที่ร้องขอ แล้วส่งกลับเข้าคิว "รอตรวจสอบ" อีกครั้ง */
function action_submitAdditionalNameChangeDocs(p, user) {
  requireRole_(user, ['member']);
  const app = findRowById_('Applications', 'application_id', p.application_id);
  if (!app || app.applicant_type !== 'name_change' || app.applicant_user_id !== user.user_id) throw new Error('ไม่พบคำขอเปลี่ยนชื่อนี้');
  if (app.status !== 'รอเอกสารเพิ่มเติม') throw new Error('คำขอนี้ไม่ได้อยู่ในสถานะรอเอกสารเพิ่มเติม');
  const existingDocs = JSON.parse(app.documents_json || '[]');
  const newDocs = existingDocs.concat(p.documents || []);
  updateRow_('Applications', app._row, {
    documents_json: JSON.stringify(newDocs), status: 'รอตรวจสอบ', reject_reason: '', submitted_at: nowIso_()
  });
  auditLog_('Applications', app.application_id, 'name_change_docs_resubmitted', app, { documents: p.documents }, user.user_id);
  return { ok: true };
}

// ---------- Dashboard (TOR ข้อ 8) — สรุปสถิติพื้นฐาน ----------

function action_dashboardSummary(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const members = readAll_('Members');
  const orgs = readAll_('MembersOrg');
  const apps = readAll_('Applications').filter(a => a.applicant_type !== 'name_change');
  const byType = {};
  members.forEach(m => { byType[m.member_type] = (byType[m.member_type] || 0) + 1; });
  if (orgs.length) byType['วิสามัญ'] = orgs.length;
  const byStatus = {};
  members.forEach(m => { byStatus[m.member_status] = (byStatus[m.member_status] || 0) + 1; });

  const LICENSE_APP_TYPES = ['ขึ้นทะเบียน', 'ต่ออายุใบอนุญาต'];
  const membershipApps = apps.filter(a => LICENSE_APP_TYPES.indexOf(a.applicant_type) === -1);
  const licenseApps = apps.filter(a => LICENSE_APP_TYPES.indexOf(a.applicant_type) !== -1);
  const pendingApps = membershipApps.filter(a => a.status === 'รอตรวจสอบ' || a.status === 'รออนุมัติ').length;
  const awaitingPayment = membershipApps.filter(a => a.status === 'รอชำระเงิน').length;
  const pendingLicenseApps = licenseApps.filter(a => a.status === 'รอตรวจสอบ' || a.status === 'รออนุมัติ').length;
  const licenseAwaitingPayment = licenseApps.filter(a => a.status === 'รอชำระเงิน').length;

  const licenses = readAll_('Licenses');
  const licenseByStatus = {};
  licenses.forEach(l => { licenseByStatus[l.status] = (licenseByStatus[l.status] || 0) + 1; });

  const honoraryCount = readHonorarySafe_().filter(h => h.status === 'ดำรงสถานะ').length;
  return {
    totalMembers: members.length + orgs.length + honoraryCount,
    honoraryCount: honoraryCount,
    byType: byType,
    byStatus: byStatus,
    pendingApplications: pendingApps,
    awaitingPayment: awaitingPayment,
    pendingLicenseApplications: pendingLicenseApps,
    licenseAwaitingPayment: licenseAwaitingPayment,
    licenseByStatus: licenseByStatus
  };
}

// ============================================================
// ระบบ Dashboard และรายงาน (TOR ข้อ 8) — ทั้งหมดอยู่ฝั่งแอดมิน/เจ้าหน้าที่
// ============================================================

// ---------- ตัวช่วยกลาง (ใช้ร่วมกันหลายรายงาน) ----------

function groupCount_(items, keyFn) {
  const counts = {};
  items.forEach(item => {
    let k = keyFn(item);
    k = (k === undefined || k === null || k === '') ? '(ไม่ระบุ)' : k;
    counts[k] = (counts[k] || 0) + 1;
  });
  return Object.entries(counts).map(([label, count]) => ({ label: label, count: count })).sort((a, b) => b.count - a.count);
}

function inDateRange_(dateIso, from, to) {
  if (!dateIso) return false;
  const d = new Date(dateIso);
  if (from && d < new Date(from)) return false;
  if (to && d > new Date(to + 'T23:59:59')) return false;
  return true;
}

function monthKey_(dateIso) {
  const d = new Date(dateIso);
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

function ageFromDob_(dobIso) {
  if (!dobIso) return null;
  const d = new Date(dobIso), now = new Date();
  let age = now.getFullYear() - d.getFullYear();
  const m = now.getMonth() - d.getMonth();
  if (m < 0 || (m === 0 && now.getDate() < d.getDate())) age--;
  return age;
}

function ageBracket_(age) {
  if (age === null) return '(ไม่ระบุ)';
  if (age < 30) return 'ต่ำกว่า 30 ปี';
  if (age < 40) return '30-39 ปี';
  if (age < 50) return '40-49 ปี';
  if (age < 60) return '50-59 ปี';
  return '60 ปีขึ้นไป';
}

// ตัวเลือกเพศ (TOR ข้อ 8.3 "รายงานสมาชิกแยกตามเพศ") — เพิ่มฟิลด์เพศจริงแล้ว ไม่ต้องประมาณจากคำนำหน้าอีกต่อไป
const GENDER_OPTIONS = ['ชาย', 'หญิง', 'LGBTQ+', 'ไม่ระบุ'];

function genderLabel_(gender) {
  return gender && GENDER_OPTIONS.indexOf(gender) !== -1 ? gender : 'ไม่ระบุ';
}

function highestEducation_(educationJson) {
  try {
    const edu = JSON.parse(educationJson || '[]');
    if (!edu.length) return '(ไม่ระบุ)';
    const order = ['ปริญญาเอก', 'ปริญญาโท', 'ปริญญาตรี', 'ปวส.', 'ปวช.', 'มัธยมศึกษา'];
    for (let i = 0; i < order.length; i++) {
      const match = edu.find(e => (e.level || '').indexOf(order[i]) !== -1);
      if (match) return order[i];
    }
    return edu[0].level || '(ไม่ระบุ)';
  } catch (e) { return '(ไม่ระบุ)'; }
}

function feeAmountByApplicantTypes_(payments, apps, types) {
  const appIds = apps.filter(a => types.indexOf(a.applicant_type) !== -1).map(a => a.application_id);
  return payments.filter(pm => pm.ref_type === 'application' && appIds.indexOf(pm.ref_id) !== -1 && pm.status === 'ชำระแล้ว');
}

// ---------- 8.1 Dashboard ผู้บริหาร ----------

function action_getExecutiveDashboard(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const members = readAll_('Members');
  const orgs = readAll_('MembersOrg');
  const licenses = readAll_('Licenses');
  const apps = readAll_('Applications');
  const payments = readAll_('Payments').filter(pm => pm.status === 'ชำระแล้ว');
  const renewalHistory = readAll_('LicenseRenewalHistory');
  const membershipRenewals = readAll_('MembershipRenewals');

  const byType = {};
  members.forEach(m => { byType[m.member_type] = (byType[m.member_type] || 0) + 1; });
  byType['วิสามัญ'] = orgs.length;

  const now = new Date();
  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1).toISOString();
  const startOfYear = new Date(now.getFullYear(), 0, 1).toISOString();

  const newMembersThisMonth = members.filter(m => inDateRange_(m.applied_date, startOfMonth, null)).length;
  const newMembersThisYear = members.filter(m => inDateRange_(m.applied_date, startOfYear, null)).length;

  const expiredMembers = members.filter(m => daysUntil_(m.membership_expire_date) !== null && daysUntil_(m.membership_expire_date) < 0).length;
  const renewedThisMonth = membershipRenewals.filter(r => inDateRange_(r.renewed_at, startOfMonth, null)).length;

  const totalLicenseHolders = licenses.length;
  const newLicensesThisMonth = apps.filter(a => a.applicant_type === 'ขึ้นทะเบียน' && a.status === 'ขึ้นทะเบียนสำเร็จ' && inDateRange_(a.reviewed_at, startOfMonth, null)).length;
  const expiredLicenses = licenses.filter(l => l.status === 'หมดอายุ').length;
  const renewedLicensesThisMonth = renewalHistory.filter(r => inDateRange_(r.renewed_at, startOfMonth, null)).length;

  const revenueThisMonth = payments.filter(pm => inDateRange_(pm.paid_at, startOfMonth, null)).reduce((s, pm) => s + (Number(pm.amount) || 0), 0);

  // ข้อมูลกราฟเปรียบเทียบ 12 เดือนล่าสุด และ 5 ปีล่าสุด (รายรับ)
  const monthlyRevenue = {};
  const yearlyRevenue = {};
  payments.forEach(pm => {
    if (!pm.paid_at) return;
    const mk = monthKey_(pm.paid_at);
    monthlyRevenue[mk] = (monthlyRevenue[mk] || 0) + (Number(pm.amount) || 0);
    const yk = String(new Date(pm.paid_at).getFullYear());
    yearlyRevenue[yk] = (yearlyRevenue[yk] || 0) + (Number(pm.amount) || 0);
  });
  const monthlyLabels = Object.keys(monthlyRevenue).sort().slice(-12);
  const yearlyLabels = Object.keys(yearlyRevenue).sort().slice(-5);

  const honoraryCount = readHonorarySafe_().filter(h => h.status === 'ดำรงสถานะ').length;
  return {
    totalMembers: members.length + orgs.length + honoraryCount,
    honoraryCount: honoraryCount,
    byType: byType,
    newMembersThisMonth: newMembersThisMonth,
    newMembersThisYear: newMembersThisYear,
    expiredMembers: expiredMembers,
    renewedThisMonth: renewedThisMonth,
    totalLicenseHolders: totalLicenseHolders,
    newLicensesThisMonth: newLicensesThisMonth,
    expiredLicenses: expiredLicenses,
    renewedLicensesThisMonth: renewedLicensesThisMonth,
    revenueThisMonth: revenueThisMonth,
    monthlyRevenueChart: { labels: monthlyLabels, values: monthlyLabels.map(k => monthlyRevenue[k]) },
    yearlyRevenueChart: { labels: yearlyLabels, values: yearlyLabels.map(k => yearlyRevenue[k]) }
  };
}

// ---------- 8.2 Dashboard เพิ่มเติม ----------

/** ข้อมูลการสอบสำหรับรายงาน — รอบสอบที่วันสอบอยู่ในช่วง from/to (ถ้าไม่ระบุ = ทุกรอบ) พร้อมใบสมัครของรอบนั้น */
function examReportData_(from, to) {
  const rounds = readAll_('ExamRounds').filter(r => !(from || to) || inDateRange_(r.exam_date, from, to));
  const roundMap = {};
  rounds.forEach(r => { roundMap[r.round_id] = r; });
  const apps = [];
  readAll_('Applications').forEach(a => {
    if (EXAM_APPLICANT_TYPES.indexOf(a.applicant_type) === -1) return;
    let data = {};
    try { data = JSON.parse(a.ref_data_json || '{}'); } catch (e) {}
    if (!roundMap[data.round_id]) return;
    apps.push(Object.assign({}, a, { data: data }));
  });
  rounds.sort((x, y) => String(y.exam_date || '').localeCompare(String(x.exam_date || '')));
  return { rounds: rounds, roundMap: roundMap, apps: apps };
}

function action_getOperationalDashboard(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const apps = readAll_('Applications');
  const members = readAll_('Members');
  const licenses = readAll_('Licenses');
  const welfare = readAll_('Welfare');

  const MEMBERSHIP_TYPES = ['สามัญ', 'สมทบ', 'วิสามัญ'];
  const membershipAppsByType = {};
  MEMBERSHIP_TYPES.forEach(t => {
    const rows = apps.filter(a => a.applicant_type === t);
    membershipAppsByType[t] = {
      pending: rows.filter(a => a.status === 'รอตรวจสอบ' || a.status === 'รออนุมัติ').length,
      awaitingPayment: rows.filter(a => a.status === 'รอชำระเงิน').length
    };
  });

  const licenseRegPending = apps.filter(a => a.applicant_type === 'ขึ้นทะเบียน' && (a.status === 'รอตรวจสอบ' || a.status === 'รออนุมัติ' || a.status === 'รอชำระเงิน')).length;
  const licenseRenewPending = apps.filter(a => a.applicant_type === 'ต่ออายุใบอนุญาต' && (a.status === 'รอตรวจสอบ' || a.status === 'รออนุมัติ' || a.status === 'รอชำระเงิน')).length;

  // SLA เฉลี่ย: เวลาจาก submitted_at ถึง reviewed_at (วัน) เฉพาะที่ผ่านการพิจารณาแล้ว
  function avgSlaDays(applicantTypes) {
    const rows = apps.filter(a => applicantTypes.indexOf(a.applicant_type) !== -1 && a.reviewed_at);
    if (!rows.length) return null;
    const totalDays = rows.reduce((s, a) => s + (new Date(a.reviewed_at) - new Date(a.submitted_at)) / (24 * 3600 * 1000), 0);
    return Math.round((totalDays / rows.length) * 10) / 10;
  }
  const slaMembership = avgSlaDays(MEMBERSHIP_TYPES);
  const slaLicense = avgSlaDays(['ขึ้นทะเบียน']);

  const memberCardExpiring90 = members.filter(m => { const d = daysUntil_(m.membership_expire_date); return d !== null && d >= 0 && d <= 90; }).length;
  const licenseExpiring180 = licenses.filter(l => { const d = daysUntil_(l.expire_date); return d !== null && d >= 0 && d <= 180; }).length;

  const nameChangePending = apps.filter(a => a.applicant_type === 'name_change' && (a.status === 'รอตรวจสอบ' || a.status === 'รออนุมัติ' || a.status === 'รอเอกสารเพิ่มเติม')).length;

  const welfarePendingByType = {};
  welfare.filter(w => w.status === 'รอตรวจสอบ' || w.status === 'รออนุมัติ').forEach(w => {
    welfarePendingByType[w.welfare_type] = (welfarePendingByType[w.welfare_type] || 0) + 1;
  });

  // สรุปการสอบ (ข้อเขียน/สัมภาษณ์) — ทุกรอบที่ยังไม่เสร็จสิ้น + ผลรวมทั้งหมด
  const examSummary = {};
  EXAM_APPLICANT_TYPES.forEach(t => {
    const list = apps.filter(a => a.applicant_type === t);
    const cnt = st => list.filter(a => a.status === st).length;
    examSummary[t] = { total: list.length, pendingReview: list.filter(a => a.status === 'รอตรวจสอบ' || a.status === 'รอชำระเงิน').length, awaitingResult: cnt('รอผลสอบ'), passed: cnt('สอบผ่าน'), failed: cnt('สอบไม่ผ่าน') };
  });
  const openRounds = readAll_('ExamRounds').filter(r => r.status !== 'เสร็จสิ้น').map(r => ({ exam_type: r.exam_type, round_label: r.round_label, exam_date: r.exam_date, status: r.status,
    applicants: apps.filter(a => { try { return JSON.parse(a.ref_data_json || '{}').round_id === r.round_id; } catch (e) { return false; } }).length }));

  return {
    examSummary: examSummary,
    examOpenRounds: openRounds,
    membershipAppsByType: membershipAppsByType,
    licenseRegPending: licenseRegPending,
    licenseRenewPending: licenseRenewPending,
    slaApprovalTargetDays: SLA_APPROVAL_DAYS,
    slaMembershipAvgDays: slaMembership,
    slaLicenseAvgDays: slaLicense,
    memberCardExpiring90: memberCardExpiring90,
    licenseExpiring180: licenseExpiring180,
    nameChangePending: nameChangePending,
    welfarePendingByType: welfarePendingByType
  };
}

/** จัดหมวดสถานะสมาชิกภาพ/ใบอนุญาต — ปกติ / หมดอายุ / ไม่มีข้อมูล (ใช้ในแดชบอร์ดสรุปจำนวนสมาชิก) */
function membershipStatusBucket_(m) {
  if (!m.membership_expire_date && !m.member_status) return 'ไม่มีข้อมูล';
  const days = daysUntil_(m.membership_expire_date);
  if (days !== null && days < 0) return 'หมดอายุ';
  return 'ปกติ';
}

function licenseStatusBucket_(l) {
  if (!l || !l.status) return 'ไม่มีข้อมูล';
  if (l.status === 'ปกติ' || l.status === 'ใกล้หมดอายุ') return 'ปกติ';
  return 'หมดอายุ'; // หมดอายุ/พักใช้/เพิกถอน/อยู่ระหว่างต่ออายุ นับรวมเป็น "ไม่ใช่สถานะปกติ" ในแดชบอร์ดสรุปนี้
}

/** แดชบอร์ดสรุปจำนวนสมาชิก แยกตาม "เข้าสู่ระบบแล้ว" (มีบัญชีผู้ใช้งานในระบบดิจิทัลนี้) กับ "ยังไม่เข้าสู่ระบบ"
 * (มีข้อมูลในทะเบียนแต่ยังไม่เคยสร้างบัญชีเข้าใช้งานเอง เช่น ข้อมูลที่เจ้าหน้าที่นำเข้า/บันทึกให้ก่อน)
 * แต่ละกลุ่มแบ่งย่อยตามสถานะ: ปกติ / หมดอายุ / ไม่มีข้อมูล — ตามรูปแบบแดชบอร์ดเดิมที่สภาฯ ใช้อยู่
 */
function action_getMemberStatusDashboard(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const members = readAll_('Members');
  const licenses = readAll_('Licenses');
  const users = readAll_('Users');
  const memberIdsWithLogin = new Set(users.filter(u => u.role === 'member' && u.member_id).map(u => u.member_id));

  function emptyBucket() { return { ปกติ: 0, หมดอายุ: 0, ไม่มีข้อมูล: 0 }; }

  // ---------- สมาชิกสามัญ ----------
  const ordinaryMembers = members.filter(m => m.member_type === 'สามัญ');
  const ordinaryLoggedIn = emptyBucket(), ordinaryNotLoggedIn = emptyBucket();
  ordinaryMembers.forEach(m => {
    const bucket = membershipStatusBucket_(m);
    if (memberIdsWithLogin.has(m.member_id)) ordinaryLoggedIn[bucket]++;
    else ordinaryNotLoggedIn[bucket]++;
  });

  // ---------- นักสังคมสงเคราะห์ที่รับอนุญาต (ผู้ถือใบอนุญาต) ----------
  const licenseLoggedIn = emptyBucket(), licenseNotLoggedIn = emptyBucket();
  licenses.forEach(l => {
    const bucket = licenseStatusBucket_(l);
    if (memberIdsWithLogin.has(l.member_id)) licenseLoggedIn[bucket]++;
    else licenseNotLoggedIn[bucket]++;
  });

  // ---------- สมาชิกสมทบ ----------
  const associateMembers = members.filter(m => m.member_type === 'สมทบ');
  const associateLoggedIn = emptyBucket(), associateNotLoggedIn = emptyBucket();
  associateMembers.forEach(m => {
    const bucket = membershipStatusBucket_(m);
    if (memberIdsWithLogin.has(m.member_id)) associateLoggedIn[bucket]++;
    else associateNotLoggedIn[bucket]++;
  });

  // ---------- เปลี่ยนสถานะสมทบไปสามัญ (สำเร็จแล้ว) ----------
  const typeChangeSuccessCount = readAll_('Applications').filter(a =>
    a.applicant_type === 'เปลี่ยนประเภท' && a.status === 'เปลี่ยนประเภทสำเร็จ'
  ).length;

  const sum = b => b.ปกติ + b.หมดอายุ + b.ไม่มีข้อมูล;

  // ---------- สมาชิกกิตติมศักดิ์ ----------
  const honoraryAll = readHonorarySafe_();
  const honoraryByType = {};
  HONORARY_QUALIFICATION_TYPES.forEach(t => { honoraryByType[t] = 0; });
  honoraryAll.filter(h => h.status === 'ดำรงสถานะ').forEach(h => { honoraryByType[h.qualification_type] = (honoraryByType[h.qualification_type] || 0) + 1; });

  return {
    honorary: {
      total: honoraryAll.length,
      active: honoraryAll.filter(h => h.status === 'ดำรงสถานะ').length,
      ended: honoraryAll.filter(h => h.status !== 'ดำรงสถานะ').length,
      byType: honoraryByType
    },
    ordinary: {
      total: ordinaryMembers.length,
      loggedIn: { total: sum(ordinaryLoggedIn), ...ordinaryLoggedIn },
      notLoggedIn: { total: sum(ordinaryNotLoggedIn), ...ordinaryNotLoggedIn }
    },
    licensed: {
      total: licenses.length,
      loggedIn: { total: sum(licenseLoggedIn), ...licenseLoggedIn },
      notLoggedIn: { total: sum(licenseNotLoggedIn), ...licenseNotLoggedIn }
    },
    associate: {
      total: associateMembers.length,
      loggedIn: { total: sum(associateLoggedIn), ...associateLoggedIn },
      notLoggedIn: { total: sum(associateNotLoggedIn), ...associateNotLoggedIn }
    },
    associateTotal: associateMembers.length,
    typeChangeSuccessCount: typeChangeSuccessCount
  };
}

/** รายชื่อสมาชิกที่มีข้อมูลอยู่แล้วในระบบ (เช่น นำเข้าจากฐานข้อมูลเดิม) แต่ยังไม่เคยสร้างบัญชีเข้าใช้งาน (Users) เลย
 * ไว้ให้เจ้าหน้าที่ติดตาม/ติดต่อให้มาสร้างบัญชีหรือใช้ฟีเจอร์ "เป็นสมาชิกอยู่แล้ว? คลิกที่นี่" ที่หน้าแรก
 * group: 'ordinary' (สมาชิกสามัญ) | 'licensed' (ผู้ถือใบอนุญาต) — ไม่ระบุ = สมาชิกสามัญทุกคนที่ยังไม่มีบัญชี
 */
function action_listMembersWithoutAccount(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const members = readAll_('Members');
  const licenses = readAll_('Licenses');
  const users = readAll_('Users');
  const memberIdsWithLogin = new Set(users.filter(u => u.role === 'member' && u.member_id).map(u => u.member_id));

  let targetMemberIds;
  if (p.group === 'licensed') {
    targetMemberIds = new Set(licenses.map(l => l.member_id).filter(id => !memberIdsWithLogin.has(id)));
  } else {
    targetMemberIds = new Set(
      members.filter(m => m.member_type === (p.group === 'associate' ? 'สมทบ' : 'สามัญ') && !memberIdsWithLogin.has(m.member_id))
        .map(m => m.member_id)
    );
  }

  return members
    .filter(m => targetMemberIds.has(m.member_id))
    .map(m => {
      const lic = licenses.find(l => l.member_id === m.member_id);
      return {
        member_id: m.member_id, member_no: m.member_no, member_type: m.member_type,
        name: [m.prefix_th, m.first_name_th, m.last_name_th].filter(Boolean).join(' '),
        phone: m.phone, email: m.email, national_id: m.national_id,
        member_status: membershipStatusBucket_(m),
        license_no: lic ? lic.license_no : '', license_status: lic ? licenseStatusBucket_(lic) : ''
      };
    })
    .slice(0, 1000); // จำกัดจำนวนต่อครั้งกันโหลดช้าเกินไปถ้ามีเยอะมาก
}

// ---------- 8.3-8.7 รายงาน (Report Catalog + Dispatcher) ----------

const REPORT_CATALOG = [
  // 8.3 รายงานสมาชิก
  { id: 'member_total', label: 'จำนวนสมาชิกทั้งหมด', category: 'สมาชิก' },
  { id: 'member_by_type', label: 'สมาชิกแยกตามประเภท', category: 'สมาชิก' },
  { id: 'member_honorary', label: 'รายชื่อสมาชิกกิตติมศักดิ์', category: 'สมาชิก' },
  { id: 'member_new', label: 'สมาชิกใหม่', category: 'สมาชิก' },
  { id: 'member_terminated', label: 'สมาชิกสิ้นสุดสมาชิกภาพ (ถูกเพิกถอน)', category: 'สมาชิก' },
  { id: 'member_expiring', label: 'สมาชิกใกล้หมดอายุ (30 วัน)', category: 'สมาชิก' },
  { id: 'member_expired', label: 'สมาชิกหมดอายุ', category: 'สมาชิก' },
  { id: 'member_by_province', label: 'สมาชิกแยกตามจังหวัด', category: 'สมาชิก' },
  { id: 'member_by_age', label: 'สมาชิกแยกตามช่วงอายุ', category: 'สมาชิก' },
  { id: 'member_by_gender', label: 'สมาชิกแยกตามเพศ', category: 'สมาชิก' },
  { id: 'member_by_education', label: 'สมาชิกแยกตามระดับการศึกษาสูงสุด', category: 'สมาชิก' },
  { id: 'member_by_position', label: 'สมาชิกแยกตามตำแหน่งงาน', category: 'สมาชิก' },
  { id: 'member_by_agency', label: 'สมาชิกแยกตามหน่วยงาน', category: 'สมาชิก' },
  { id: 'member_by_affiliation', label: 'สมาชิกแยกตามสังกัด', category: 'สมาชิก' },
  { id: 'member_by_expertise', label: 'สมาชิกแยกตามความเชี่ยวชาญ', category: 'สมาชิก' },
  // 8.4 รายงานใบอนุญาต
  { id: 'license_holders', label: 'ผู้ได้รับใบอนุญาตทั้งหมด', category: 'ใบอนุญาต' },
  { id: 'license_expiring', label: 'ใบอนุญาตใกล้หมดอายุ', category: 'ใบอนุญาต' },
  { id: 'license_expired', label: 'ใบอนุญาตหมดอายุ', category: 'ใบอนุญาต' },
  { id: 'license_renewing', label: 'ใบอนุญาตอยู่ระหว่างต่ออายุ', category: 'ใบอนุญาต' },
  { id: 'license_suspended', label: 'ใบอนุญาตถูกพักใช้', category: 'ใบอนุญาต' },
  { id: 'license_revoked', label: 'ใบอนุญาตถูกเพิกถอน', category: 'ใบอนุญาต' },
  { id: 'license_issued_monthly', label: 'การออกใบอนุญาตรายเดือน', category: 'ใบอนุญาต' },
  { id: 'license_renewed_monthly', label: 'การต่ออายุใบอนุญาตรายเดือน', category: 'ใบอนุญาต' },
  { id: 'license_renewed_yearly', label: 'การต่ออายุใบอนุญาตรายปี', category: 'ใบอนุญาต' },
  // 8.5 รายงานการเงิน
  { id: 'fee_registration', label: 'ค่าธรรมเนียมสมัครสมาชิก', category: 'การเงิน' },
  { id: 'fee_membership_renewal', label: 'ค่าธรรมเนียมต่ออายุสมาชิก', category: 'การเงิน' },
  { id: 'fee_license_registration', label: 'ค่าธรรมเนียมขึ้นทะเบียนผู้ประกอบวิชาชีพ', category: 'การเงิน' },
  { id: 'fee_license_renewal', label: 'ค่าธรรมเนียมต่ออายุใบอนุญาต', category: 'การเงิน' },
  { id: 'fee_type_change', label: 'ค่าธรรมเนียมเปลี่ยนประเภทสมาชิก', category: 'การเงิน' },
  { id: 'revenue_daily', label: 'รายรับรายวัน', category: 'การเงิน' },
  { id: 'revenue_monthly', label: 'รายรับรายเดือน', category: 'การเงิน' },
  { id: 'revenue_yearly', label: 'รายรับรายปี', category: 'การเงิน' },
  { id: 'revenue_by_service', label: 'รายรับแยกตามประเภทบริการ', category: 'การเงิน' },
  { id: 'payments_overdue', label: 'ค้างชำระ', category: 'การเงิน' },
  { id: 'revenue_summary', label: 'สรุปรายรับทั้งหมด', category: 'การเงิน' },
  // 8.6 รายงานสวัสดิการ
  { id: 'welfare_applicants', label: 'จำนวนผู้ขอรับสวัสดิการ', category: 'สวัสดิการ' },
  { id: 'welfare_by_type', label: 'สวัสดิการแยกตามประเภท', category: 'สวัสดิการ' },
  { id: 'welfare_amount_paid', label: 'จำนวนเงินที่จ่าย', category: 'สวัสดิการ' },
  { id: 'welfare_by_status', label: 'สวัสดิการแยกตามสถานะการอนุมัติ', category: 'สวัสดิการ' },
  { id: 'welfare_monthly', label: 'รายงานสวัสดิการประจำเดือน', category: 'สวัสดิการ' },
  { id: 'welfare_yearly', label: 'รายงานสวัสดิการประจำปี', category: 'สวัสดิการ' },
  // 8.7 รายงานอื่นๆ
  { id: 'application_summary', label: 'สรุปคำขอสมัครสมาชิก แยกสถานะและประเภท', category: 'อื่นๆ' },
  { id: 'license_renewal_eligible', label: 'สมาชิกมีสิทธิ์ต่ออายุใบอนุญาตใน 90/180 วันข้างหน้า', category: 'อื่นๆ' },
  { id: 'name_change_history', label: 'ประวัติการเปลี่ยนชื่อ-นามสกุล พร้อมผู้อนุมัติ', category: 'อื่นๆ' },
  // 8.8 รายงานการสอบ (ข้อเขียน/สัมภาษณ์) — ช่วงวันที่กรองตามวันสอบของรอบ
  { id: 'exam_round_summary', label: 'สรุปการสอบข้อเขียน/สัมภาษณ์ แยกรอบสอบ', category: 'อื่นๆ' },
  { id: 'exam_written_results', label: 'ผลสอบข้อเขียนรายบุคคล (คะแนนทฤษฎี/ปฏิบัติ)', category: 'อื่นๆ' },
  { id: 'exam_interview_results', label: 'ผลสอบสัมภาษณ์รายบุคคล (คะแนนกรรมการ/เฉลี่ย)', category: 'อื่นๆ' }
];

function action_listReportCatalog(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  return REPORT_CATALOG;
}

/** ตัวจ่ายงานหลักของทุกรายงาน — คืนค่า {title, columns, rows} เสมอ เพื่อให้ export ใช้รูปแบบเดียวกันได้ทั้งหมด
 * p.report_id, p.from (ISO date, optional), p.to (ISO date, optional)
 */
function action_getReport(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const catalogEntry = REPORT_CATALOG.find(r => r.id === p.report_id);
  if (!catalogEntry) throw new Error('ไม่พบรายงานนี้');

  const members = readAll_('Members');
  const orgs = readAll_('MembersOrg');
  const honoraryAll = readHonorarySafe_();
  const honoraryActive = honoraryAll.filter(h => h.status === 'ดำรงสถานะ');
  const licenses = readAll_('Licenses');
  const apps = readAll_('Applications');
  const payments = readAll_('Payments');
  const welfare = readAll_('Welfare');
  const renewalHistory = readAll_('LicenseRenewalHistory');
  const from = p.from || null, to = p.to || null;

  function memberName(m) { return [m.prefix_th, m.first_name_th, m.last_name_th].filter(Boolean).join(' '); }

  switch (p.report_id) {
    case 'member_total':
      return { title: catalogEntry.label, columns: ['รายการ', 'จำนวน'], rows: [
        ['สมาชิกสามัญ', members.filter(m => m.member_type === 'สามัญ').length],
        ['สมาชิกสมทบ', members.filter(m => m.member_type === 'สมทบ').length],
        ['สมาชิกวิสามัญ (องค์กร)', orgs.length],
        ['สมาชิกกิตติมศักดิ์ (ดำรงสถานะ)', honoraryActive.length],
        ['รวมทั้งหมด', members.length + orgs.length + honoraryActive.length]
      ] };

    case 'member_by_type':
      return { title: catalogEntry.label, columns: ['ประเภท', 'จำนวน'], rows: groupCount_(members, m => m.member_type).map(r => [r.label, r.count]).concat([['วิสามัญ', orgs.length], ['กิตติมศักดิ์', honoraryActive.length]]) };

    case 'member_honorary': {
      const rows = honoraryAll.filter(h => inDateRange_(h.appointed_date, from, to));
      return { title: catalogEntry.label, columns: ['เลขทะเบียน', 'ชื่อ-นามสกุล', 'คุณสมบัติ', 'ตำแหน่ง', 'หน่วยงาน', 'วันที่แต่งตั้ง', 'สถานะ'],
        rows: rows.map(h => [h.honorary_no, [h.prefix_th, h.first_name_th, h.last_name_th].filter(Boolean).join(' '), h.qualification_type, h.position, h.organization, h.appointed_date, h.status]) };
    }

    case 'member_new': {
      const rows = members.filter(m => inDateRange_(m.applied_date, from, to));
      return { title: catalogEntry.label, columns: ['เลขสมาชิก', 'ชื่อ-นามสกุล', 'ประเภท', 'วันที่สมัคร'], rows: rows.map(m => [m.member_no, memberName(m), m.member_type, m.applied_date]) };
    }

    case 'member_terminated': {
      const rows = members.filter(m => m.member_status === 'เพิกถอน');
      return { title: catalogEntry.label, columns: ['เลขสมาชิก', 'ชื่อ-นามสกุล', 'ประเภท', 'วันหมดอายุ'], rows: rows.map(m => [m.member_no, memberName(m), m.member_type, m.membership_expire_date]) };
    }

    case 'member_expiring': {
      const rows = members.filter(m => { const d = daysUntil_(m.membership_expire_date); return d !== null && d >= 0 && d <= EXPIRY_WARNING_DAYS; });
      return { title: catalogEntry.label, columns: ['เลขสมาชิก', 'ชื่อ-นามสกุล', 'วันหมดอายุ', 'เหลือ (วัน)'], rows: rows.map(m => [m.member_no, memberName(m), m.membership_expire_date, daysUntil_(m.membership_expire_date)]) };
    }

    case 'member_expired': {
      const rows = members.filter(m => { const d = daysUntil_(m.membership_expire_date); return d !== null && d < 0; });
      return { title: catalogEntry.label, columns: ['เลขสมาชิก', 'ชื่อ-นามสกุล', 'วันหมดอายุ', 'เกินกำหนด (วัน)'], rows: rows.map(m => [m.member_no, memberName(m), m.membership_expire_date, Math.abs(daysUntil_(m.membership_expire_date))]) };
    }

    case 'member_by_province':
      return { title: catalogEntry.label, columns: ['จังหวัด', 'จำนวน'], rows: groupCount_(members, m => m.work_province).map(r => [r.label, r.count]) };

    case 'member_by_age':
      return { title: catalogEntry.label, columns: ['ช่วงอายุ', 'จำนวน'], rows: groupCount_(members, m => ageBracket_(ageFromDob_(m.dob))).map(r => [r.label, r.count]) };

    case 'member_by_gender':
      return { title: catalogEntry.label, columns: ['เพศ', 'จำนวน'], rows: groupCount_(members, m => genderLabel_(m.gender)).map(r => [r.label, r.count]) };

    case 'member_by_education':
      return { title: catalogEntry.label, columns: ['ระดับการศึกษาสูงสุด', 'จำนวน'], rows: groupCount_(members, m => highestEducation_(m.education_json)).map(r => [r.label, r.count]) };

    case 'member_by_position':
      return { title: catalogEntry.label, columns: ['ตำแหน่งงาน', 'จำนวน'], rows: groupCount_(members, m => m.position).map(r => [r.label, r.count]) };

    case 'member_by_agency':
      return { title: catalogEntry.label, columns: ['หน่วยงาน', 'จำนวน'], rows: groupCount_(members, m => m.agency).map(r => [r.label, r.count]) };

    case 'member_by_affiliation':
      return { title: catalogEntry.label, columns: ['สังกัด', 'จำนวน'], rows: groupCount_(members, m => m.affiliation).map(r => [r.label, r.count]) };

    case 'member_by_expertise':
      return { title: catalogEntry.label, columns: ['ความเชี่ยวชาญ', 'จำนวน'], rows: groupCount_(members, m => m.expertise).map(r => [r.label, r.count]) };

    case 'license_holders':
      return { title: catalogEntry.label, columns: ['เลขที่ใบอนุญาต', 'สมาชิก', 'วันออก', 'วันหมดอายุ', 'สถานะ'], rows: licenses.map(l => {
        const m = members.find(x => x.member_id === l.member_id);
        return [l.license_no, m ? memberName(m) : l.member_id, l.issue_date, l.expire_date, l.status];
      }) };

    case 'license_expiring':
    case 'license_expired':
    case 'license_renewing':
    case 'license_suspended':
    case 'license_revoked': {
      const statusMap = { license_expiring: 'ใกล้หมดอายุ', license_expired: 'หมดอายุ', license_renewing: 'อยู่ระหว่างต่ออายุ', license_suspended: 'พักใช้', license_revoked: 'เพิกถอน' };
      const rows = licenses.filter(l => l.status === statusMap[p.report_id]);
      return { title: catalogEntry.label, columns: ['เลขที่ใบอนุญาต', 'สมาชิก', 'วันหมดอายุ'], rows: rows.map(l => {
        const m = members.find(x => x.member_id === l.member_id);
        return [l.license_no, m ? memberName(m) : l.member_id, l.expire_date];
      }) };
    }

    case 'license_issued_monthly': {
      const rows = groupCount_(licenses.filter(l => inDateRange_(l.issue_date, from, to)), l => monthKey_(l.issue_date));
      return { title: catalogEntry.label, columns: ['เดือน', 'จำนวนที่ออก'], rows: rows.sort((a, b) => a.label.localeCompare(b.label)).map(r => [r.label, r.count]) };
    }

    case 'license_renewed_monthly': {
      const rows = groupCount_(renewalHistory.filter(r => inDateRange_(r.renewed_at, from, to)), r => monthKey_(r.renewed_at));
      return { title: catalogEntry.label, columns: ['เดือน', 'จำนวนที่ต่ออายุ'], rows: rows.sort((a, b) => a.label.localeCompare(b.label)).map(r => [r.label, r.count]) };
    }

    case 'license_renewed_yearly': {
      const rows = groupCount_(renewalHistory.filter(r => inDateRange_(r.renewed_at, from, to)), r => String(new Date(r.renewed_at).getFullYear()));
      return { title: catalogEntry.label, columns: ['ปี', 'จำนวนที่ต่ออายุ'], rows: rows.sort((a, b) => a.label.localeCompare(b.label)).map(r => [r.label, r.count]) };
    }

    case 'fee_registration':
    case 'fee_membership_renewal':
    case 'fee_license_registration':
    case 'fee_license_renewal':
    case 'fee_type_change': {
      const typeMap = {
        fee_registration: ['สามัญ', 'สมทบ', 'วิสามัญ'], fee_membership_renewal: ['ต่ออายุสมาชิก'],
        fee_license_registration: ['ขึ้นทะเบียน'], fee_license_renewal: ['ต่ออายุใบอนุญาต'], fee_type_change: ['เปลี่ยนประเภท']
      };
      const rows = feeAmountByApplicantTypes_(payments, apps, typeMap[p.report_id]).filter(pm => inDateRange_(pm.paid_at, from, to));
      const total = rows.reduce((s, r) => s + (Number(r.amount) || 0), 0);
      return { title: catalogEntry.label, columns: ['เลขที่ชำระเงิน', 'จำนวนเงิน', 'วันที่ชำระ'], rows: rows.map(r => [r.payment_id, r.amount, r.paid_at]).concat([['รวม', total, '']]) };
    }

    case 'revenue_daily': {
      const paid = payments.filter(pm => pm.status === 'ชำระแล้ว' && inDateRange_(pm.paid_at, from, to));
      const rows = groupSum_(paid, pm => new Date(pm.paid_at).toLocaleDateString('th-TH'), pm => Number(pm.amount) || 0);
      return { title: catalogEntry.label, columns: ['วันที่', 'รายรับ (บาท)'], rows: rows.map(r => [r.label, r.sum]) };
    }

    case 'revenue_monthly': {
      const paid = payments.filter(pm => pm.status === 'ชำระแล้ว' && inDateRange_(pm.paid_at, from, to));
      const rows = groupSum_(paid, pm => monthKey_(pm.paid_at), pm => Number(pm.amount) || 0);
      return { title: catalogEntry.label, columns: ['เดือน', 'รายรับ (บาท)'], rows: rows.sort((a, b) => a.label.localeCompare(b.label)).map(r => [r.label, r.sum]) };
    }

    case 'revenue_yearly': {
      const paid = payments.filter(pm => pm.status === 'ชำระแล้ว' && inDateRange_(pm.paid_at, from, to));
      const rows = groupSum_(paid, pm => String(new Date(pm.paid_at).getFullYear()), pm => Number(pm.amount) || 0);
      return { title: catalogEntry.label, columns: ['ปี', 'รายรับ (บาท)'], rows: rows.sort((a, b) => a.label.localeCompare(b.label)).map(r => [r.label, r.sum]) };
    }

    case 'revenue_by_service': {
      const paid = payments.filter(pm => pm.status === 'ชำระแล้ว' && inDateRange_(pm.paid_at, from, to) && pm.ref_type === 'application');
      const rows = groupSum_(paid, pm => { const a = apps.find(x => x.application_id === pm.ref_id); return a ? a.applicant_type : '(ไม่ระบุ)'; }, pm => Number(pm.amount) || 0);
      return { title: catalogEntry.label, columns: ['ประเภทบริการ', 'รายรับ (บาท)'], rows: rows.map(r => [r.label, r.sum]) };
    }

    case 'payments_overdue': {
      const rows = payments.filter(pm => pm.status === 'รอชำระ' || pm.status === 'รอตรวจสอบ');
      return { title: catalogEntry.label, columns: ['เลขที่ชำระเงิน', 'อ้างอิง', 'จำนวนเงิน', 'สถานะ'], rows: rows.map(pm => [pm.payment_id, pm.ref_type + ' — ' + pm.ref_id, pm.amount, pm.status]) };
    }

    case 'revenue_summary': {
      const paid = payments.filter(pm => pm.status === 'ชำระแล้ว' && inDateRange_(pm.paid_at, from, to));
      const total = paid.reduce((s, pm) => s + (Number(pm.amount) || 0), 0);
      const byType = groupSum_(paid, pm => { const a = apps.find(x => x.application_id === pm.ref_id); return a ? a.applicant_type : '(ไม่ระบุ)'; }, pm => Number(pm.amount) || 0);
      return { title: catalogEntry.label, columns: ['ประเภทบริการ', 'รายรับ (บาท)'], rows: byType.map(r => [r.label, r.sum]).concat([['รวมทั้งหมด', total]]) };
    }

    case 'welfare_applicants': {
      const uniqueMembers = [...new Set(welfare.map(w => w.member_id))];
      return { title: catalogEntry.label, columns: ['รายการ', 'จำนวน'], rows: [['จำนวนคำขอทั้งหมด', welfare.length], ['จำนวนสมาชิกที่เคยยื่นขอ', uniqueMembers.length]] };
    }

    case 'welfare_by_type':
      return { title: catalogEntry.label, columns: ['ประเภทสวัสดิการ', 'จำนวนคำขอ'], rows: groupCount_(welfare, w => w.welfare_type).map(r => [r.label, r.count]) };

    case 'welfare_amount_paid': {
      const disbursed = welfare.filter(w => w.disbursed_at && inDateRange_(w.disbursed_at, from, to));
      const total = disbursed.reduce((s, w) => s + (Number(w.approved_amount) || 0), 0);
      return { title: catalogEntry.label, columns: ['เลขที่คำขอ', 'ประเภท', 'จำนวนเงิน', 'วันที่จ่าย'], rows: disbursed.map(w => [w.welfare_id, w.welfare_type, w.approved_amount, w.disbursed_at]).concat([['รวม', '', total, '']]) };
    }

    case 'welfare_by_status':
      return { title: catalogEntry.label, columns: ['สถานะ', 'จำนวน'], rows: groupCount_(welfare, w => w.status).map(r => [r.label, r.count]) };

    case 'welfare_monthly': {
      const rows = groupCount_(welfare.filter(w => inDateRange_(w.requested_at, from, to)), w => monthKey_(w.requested_at));
      return { title: catalogEntry.label, columns: ['เดือน', 'จำนวนคำขอ'], rows: rows.sort((a, b) => a.label.localeCompare(b.label)).map(r => [r.label, r.count]) };
    }

    case 'welfare_yearly': {
      const rows = groupCount_(welfare.filter(w => inDateRange_(w.requested_at, from, to)), w => String(new Date(w.requested_at).getFullYear()));
      return { title: catalogEntry.label, columns: ['ปี', 'จำนวนคำขอ'], rows: rows.sort((a, b) => a.label.localeCompare(b.label)).map(r => [r.label, r.count]) };
    }

    case 'application_summary': {
      const relevant = apps.filter(a => ['สามัญ', 'สมทบ', 'วิสามัญ'].indexOf(a.applicant_type) !== -1);
      const rows = [];
      ['สามัญ', 'สมทบ', 'วิสามัญ'].forEach(type => {
        const rowsOfType = relevant.filter(a => a.applicant_type === type);
        rows.push([type, 'เป็นสมาชิกแล้ว', rowsOfType.filter(a => a.status === 'เป็นสมาชิกแล้ว').length]);
        rows.push([type, 'ไม่อนุมัติ', rowsOfType.filter(a => a.status === 'ไม่อนุมัติ').length]);
        rows.push([type, 'อยู่ระหว่างดำเนินการ', rowsOfType.filter(a => ['เป็นสมาชิกแล้ว', 'ไม่อนุมัติ'].indexOf(a.status) === -1).length]);
      });
      return { title: catalogEntry.label, columns: ['ประเภทสมาชิก', 'สถานะ', 'จำนวน'], rows: rows };
    }

    case 'license_renewal_eligible': {
      const rows = licenses.filter(l => { const d = daysUntil_(l.expire_date); return d !== null && d >= 0 && d <= RENEWAL_WINDOW_DAYS; });
      return { title: catalogEntry.label, columns: ['เลขที่ใบอนุญาต', 'สมาชิก', 'วันหมดอายุ', 'เหลือ (วัน)'], rows: rows.map(l => {
        const m = members.find(x => x.member_id === l.member_id);
        return [l.license_no, m ? memberName(m) : l.member_id, l.expire_date, daysUntil_(l.expire_date)];
      }) };
    }

    case 'exam_round_summary': {
      const ex = examReportData_(from, to);
      const rows = ex.rounds.map(r => {
        const list = ex.apps.filter(a => a.data.round_id === r.round_id);
        const cnt = f => list.filter(f).length;
        const pass = cnt(a => a.status === 'สอบผ่าน'), fail = cnt(a => a.status === 'สอบไม่ผ่าน');
        return [r.exam_type, r.round_label, r.exam_date ? fmtThaiDate_(r.exam_date) : '-', r.status, list.length,
          cnt(a => a.status === 'รอผลสอบ'), pass, fail, pass + fail ? Math.round(pass * 1000 / (pass + fail)) / 10 + '%' : '-',
          cnt(a => ['สอบผ่าน', 'สอบไม่ผ่าน', 'รอผลสอบ'].indexOf(a.status) === -1)];
      });
      return { title: catalogEntry.label, columns: ['ประเภทสอบ', 'รอบ', 'วันสอบ', 'สถานะรอบ', 'ผู้สมัครทั้งหมด', 'รอผลสอบ', 'สอบผ่าน', 'สอบไม่ผ่าน', 'อัตราผ่าน', 'อื่นๆ (รอตรวจสอบ/ชำระเงิน/ปฏิเสธ/คืนเงิน)'], rows: rows };
    }

    case 'exam_written_results': {
      const ex = examReportData_(from, to);
      const rows = [];
      ex.apps.filter(a => a.applicant_type === 'สมัครสอบข้อเขียน').forEach(a => {
        const r = a.data.exam_result || {}, sc = r.scores || {}, rd = ex.roundMap[a.data.round_id] || {};
        rows.push([rd.round_label || a.data.round_label || '-', a.data.candidate_no || '', a.data.name || '', a.status,
          sc[EXAM_SUBJECTS[0]] !== undefined ? sc[EXAM_SUBJECTS[0]] : '-', sc[EXAM_SUBJECTS[1]] !== undefined ? sc[EXAM_SUBJECTS[1]] : '-', r.result || '-', r.exam_date ? fmtThaiDate_(r.exam_date) : (rd.exam_date ? fmtThaiDate_(rd.exam_date) : '-')]);
      });
      rows.sort((x, y) => String(x[0]).localeCompare(String(y[0])) || String(x[1]).localeCompare(String(y[1])));
      return { title: catalogEntry.label, columns: ['รอบ', 'รหัสสอบ', 'ชื่อ-สกุล', 'สถานะ', 'คะแนน: ' + EXAM_SUBJECTS[0], 'คะแนน: ' + EXAM_SUBJECTS[1], 'ผลสอบ', 'วันสอบ'], rows: rows };
    }

    case 'exam_interview_results': {
      const ex = examReportData_(from, to);
      const sessions = readAll_('InterviewSessions'), groups = readAll_('InterviewGroups'), ics = readAll_('InterviewCandidates');
      const rows = [];
      ex.apps.filter(a => a.applicant_type === 'สมัครสอบสัมภาษณ์').forEach(a => {
        const rd = ex.roundMap[a.data.round_id] || {};
        const ic = ics.find(c => c.application_id === a.application_id);
        const g = ic && groups.find(x => x.group_id === ic.group_id);
        const sess = sessions.find(x => x.round_id === a.data.round_id);
        const t = ic ? interviewTotals_(ic, interviewPassScore_(sess)) : null;
        const er = a.data.exam_result || {};
        rows.push([rd.round_label || a.data.round_label || '-', g ? 'กลุ่ม ' + g.group_no : '-', a.data.candidate_no || '', a.data.name || '',
          t && t.s1 !== null ? t.s1 : '-', t && t.s2 !== null ? t.s2 : '-', t && t.average !== null ? t.average : (er.score !== undefined && er.score !== null ? er.score : '-'),
          t && t.status === 'ไม่เข้าสอบ' ? 'ไม่เข้าสอบ' : (er.result || (t ? t.status : '-')), a.status]);
      });
      rows.sort((x, y) => String(x[0]).localeCompare(String(y[0])) || String(x[1]).localeCompare(String(y[1])) || String(x[2]).localeCompare(String(y[2])));
      return { title: catalogEntry.label, columns: ['รอบ', 'กลุ่ม', 'รหัสสอบ', 'ชื่อ-สกุล', 'คะแนน ก.1', 'คะแนน ก.2', 'คะแนนเฉลี่ย', 'ผลการประเมิน', 'สถานะในระบบ'], rows: rows };
    }

    case 'name_change_history': {
      const rows = apps.filter(a => a.applicant_type === 'name_change' && inDateRange_(a.submitted_at, from, to));
      return { title: catalogEntry.label, columns: ['เลขที่คำขอ', 'สมาชิก', 'วันที่ยื่น', 'สถานะ', 'ผู้อนุมัติ', 'วันที่อนุมัติ'], rows: rows.map(a => {
        const d = JSON.parse(a.ref_data_json || '{}');
        return [a.application_id, d.member_id || '-', a.submitted_at, a.status, a.reviewed_by || '-', a.reviewed_at || '-'];
      }) };
    }

    default:
      throw new Error('รายงานนี้ยังไม่รองรับ');
  }
}

function groupSum_(items, keyFn, valueFn) {
  const sums = {};
  items.forEach(item => {
    let k = keyFn(item);
    k = (k === undefined || k === null || k === '') ? '(ไม่ระบุ)' : k;
    sums[k] = (sums[k] || 0) + valueFn(item);
  });
  return Object.entries(sums).map(([label, sum]) => ({ label: label, sum: sum }));
}

// ---------- ส่งออกรายงาน (Excel / PDF) — CSV ทำฝั่ง client ได้เลยไม่ต้องเรียก backend ----------

/** ส่งออกรายงานเป็น Excel (.xlsx) จริง ผ่านการสร้าง Google Sheet ชั่วคราวแล้ว export */
function action_exportReportExcel(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const report = action_getReport(p, user);
  const ss = SpreadsheetApp.create(report.title);
  const sh = ss.getSheets()[0];
  sh.getRange(1, 1, 1, report.columns.length).setValues([report.columns]).setFontWeight('bold');
  if (report.rows.length) {
    sh.getRange(2, 1, report.rows.length, report.columns.length).setValues(report.rows.map(r => report.columns.map((c, i) => r[i] !== undefined ? r[i] : '')));
  }
  sh.autoResizeColumns(1, report.columns.length);

  const file = DriveApp.getFileById(ss.getId());
  const excelBlob = file.getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  const folder = getDocsFolder_();
  const excelFile = folder.createFile(excelBlob);
  excelFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  file.setTrashed(true);
  return { url: excelFile.getUrl() };
}

/** ส่งออกรายงานเป็น PDF ผ่าน Google Docs (ตารางแบบง่าย) */
function action_exportReportPdf(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const report = action_getReport(p, user);
  const doc = DocumentApp.create(report.title);
  const body = doc.getBody();
  body.appendParagraph(report.title).setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('สภาวิชาชีพสังคมสงเคราะห์ — พิมพ์เมื่อ ' + new Date().toLocaleString('th-TH'));
  body.appendParagraph(' ');
  const table = body.appendTable([report.columns].concat(report.rows.map(r => r.map(v => String(v !== undefined ? v : '')))));
  table.getRow(0).editAsText().setBold(true);
  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  const pdfBlob = file.getAs('application/pdf');
  const folder = getDocsFolder_();
  const pdfFile = folder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  file.setTrashed(true);
  return { url: pdfFile.getUrl() };
}

// ---------- 8.8 การเชื่อมโยงสถิติ — endpoint สาธารณะให้เว็บไซต์หลัก/เว็บ CE ดึงไปแสดงแบบ Realtime ----------
/** PUBLIC ไม่ต้อง login — คืนเฉพาะจำนวนสมาชิกรวม/แยกประเภท (ไม่มีข้อมูลส่วนบุคคล) ให้เว็บไซต์ภายนอกเรียกใช้ได้
 * หมายเหตุ: endpoint นี้แค่ "เปิดให้ดึงได้" เท่านั้น — การนำไปแสดงผลจริงบนเว็บไซต์หลักของสภาฯ/เว็บ CE ต้องให้ผู้ดูแลเว็บไซต์นั้นเขียนโค้ดเรียก endpoint นี้เอง (ดู REPORTS_SYSTEM.md)
 */
function action_publicStats(p) {
  const members = readAll_('Members');
  const orgs = readAll_('MembersOrg');
  const licenses = readAll_('Licenses');
  const byType = {};
  members.forEach(m => { byType[m.member_type] = (byType[m.member_type] || 0) + 1; });
  const licensedCount = licenses.filter(l => l.status === 'ปกติ').length;
  let honoraryCount = 0;
  try { honoraryCount = readAll_('MembersHonorary').filter(h => h.status === 'ดำรงสถานะ').length; } catch (e) { /* ชีตยังไม่ถูกสร้าง — ยังไม่ได้รัน setupAllSheets */ }
  return {
    totalMembers: members.length + orgs.length + honoraryCount,
    honoraryCount: honoraryCount,
    byType: byType,
    orgCount: orgs.length,
    licensedCount: licensedCount,
    updatedAt: nowIso_()
  };
}

// ============================================================
// ระบบสิทธิ์ผู้ใช้งานและความปลอดภัย (TOR ข้อ 9)
// ============================================================

const ADMIN_ROLES = ['admin_super', 'admin_staff', 'admin_welfare'];
const ROLE_LABELS = {
  admin_super: 'ผู้ดูแลระบบ (สิทธิ์สูงสุด)',
  admin_staff: 'เจ้าหน้าที่ฝ่ายทะเบียน/การเงิน',
  admin_welfare: 'เจ้าหน้าที่สวัสดิการ',
  member: 'สมาชิก'
};

// ---------- Role & Permission ----------

/** รายชื่อบัญชีผู้ใช้งานฝั่งแอดมิน/เจ้าหน้าที่เท่านั้น (ไม่รวมบัญชีสมาชิกทั่วไป) — จำกัดเฉพาะ admin_super */
function action_listAdminUsers(p, user) {
  requireRole_(user, ['admin_super']);
  return readAll_('Users').filter(u => ADMIN_ROLES.indexOf(u.role) !== -1).map(u => ({
    user_id: u.user_id, username: u.username, role: u.role, role_label: ROLE_LABELS[u.role] || u.role,
    display_name: u.display_name || '', status: u.status, created_at: u.created_at, last_login: u.last_login,
    is_locked: !!(u.locked_until && new Date(u.locked_until) > new Date())
  }));
}

/** สร้างบัญชีเจ้าหน้าที่/แอดมินใหม่ — จำกัดเฉพาะ admin_super (Role & Permission ตาม TOR ข้อ 9) */
function action_createAdminUser(p, user) {
  requireRole_(user, ['admin_super']);
  if (ADMIN_ROLES.indexOf(p.role) === -1) throw new Error('บทบาทไม่ถูกต้อง ต้องเป็น admin_super, admin_staff หรือ admin_welfare');
  const existing = findRowById_('Users', 'username', p.username);
  if (existing) throw new Error('ชื่อผู้ใช้งานนี้ถูกใช้แล้ว');
  validatePasswordPolicy_(p.password);
  const newUser = {
    user_id: newId_('USR'), username: p.username, password_hash: hashPassword_(p.password),
    role: p.role, member_id: '', status: 'active', created_at: nowIso_(), last_login: '',
    failed_login_attempts: 0, locked_until: '', display_name: p.display_name || ''
  };
  appendRow_('Users', newUser);
  auditLog_('Users', newUser.user_id, 'สร้างบัญชีเจ้าหน้าที่ (' + ROLE_LABELS[p.role] + ')', {}, { username: p.username, role: p.role }, user.user_id);
  return { user_id: newUser.user_id };
}

/** แก้ไขบทบาท/สถานะ/ชื่อแสดงผลของบัญชีเจ้าหน้าที่ — จำกัดเฉพาะ admin_super */
function action_updateAdminUser(p, user) {
  requireRole_(user, ['admin_super']);
  const target = findRowById_('Users', 'user_id', p.user_id);
  if (!target) throw new Error('ไม่พบบัญชีผู้ใช้งาน');
  if (ADMIN_ROLES.indexOf(target.role) === -1) throw new Error('บัญชีนี้ไม่ใช่บัญชีเจ้าหน้าที่');
  if (target.user_id === user.user_id && p.status === 'inactive') throw new Error('ไม่สามารถระงับบัญชีของตนเองได้');

  const patch = {};
  if (p.role) {
    if (ADMIN_ROLES.indexOf(p.role) === -1) throw new Error('บทบาทไม่ถูกต้อง');
    patch.role = p.role;
  }
  if (p.status) patch.status = p.status;
  if (p.display_name !== undefined) patch.display_name = p.display_name;
  if (p.unlock) { patch.locked_until = ''; patch.failed_login_attempts = 0; }

  updateRow_('Users', target._row, patch);
  auditLog_('Users', target.user_id, 'แก้ไขบัญชีเจ้าหน้าที่', target, patch, user.user_id);
  return { ok: true };
}

/** แอดมินตั้งรหัสผ่านใหม่ให้บัญชีเจ้าหน้าที่คนอื่น (กรณีลืมรหัสผ่าน) — จำกัดเฉพาะ admin_super */
function action_resetAdminUserPassword(p, user) {
  requireRole_(user, ['admin_super']);
  const target = findRowById_('Users', 'user_id', p.user_id);
  if (!target) throw new Error('ไม่พบบัญชีผู้ใช้งาน');
  if (ADMIN_ROLES.indexOf(target.role) === -1) throw new Error('บัญชีนี้ไม่ใช่บัญชีเจ้าหน้าที่');
  validatePasswordPolicy_(p.new_password);
  updateRow_('Users', target._row, { password_hash: hashPassword_(p.new_password), failed_login_attempts: 0, locked_until: '' });
  auditLog_('Users', target.user_id, 'รีเซ็ตรหัสผ่านโดยผู้ดูแลระบบ', {}, {}, user.user_id);
  return { ok: true };
}

/** ผู้ใช้งานเปลี่ยนรหัสผ่านของตนเอง (สมาชิกและเจ้าหน้าที่ใช้ร่วมกันได้) */
function action_changeMyPassword(p, user) {
  const account = findRowById_('Users', 'user_id', user.user_id);
  if (!account) throw new Error('ไม่พบบัญชีผู้ใช้งาน');
  if (account.password_hash !== hashPassword_(p.current_password)) throw new Error('รหัสผ่านปัจจุบันไม่ถูกต้อง');
  validatePasswordPolicy_(p.new_password);
  updateRow_('Users', account._row, { password_hash: hashPassword_(p.new_password) });
  auditLog_('Users', account.user_id, 'เปลี่ยนรหัสผ่านตนเอง', {}, {}, user.user_id);
  return { ok: true };
}

// ---------- Audit Log ----------

/** ดูประวัติการใช้งาน/แก้ไขข้อมูลในระบบ — จำกัดเฉพาะ admin_super (ข้อมูลอ่อนไหว เห็นการกระทำของเจ้าหน้าที่ทุกคน) */
function action_listAuditLog(p, user) {
  requireRole_(user, ['admin_super']);
  let rows = readAll_('AuditLog').sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
  if (p.ref_type) rows = rows.filter(r => r.ref_type === p.ref_type);
  if (p.actor_user_id) rows = rows.filter(r => r.actor_user_id === p.actor_user_id);
  if (p.from) rows = rows.filter(r => new Date(r.created_at) >= new Date(p.from));
  if (p.to) rows = rows.filter(r => new Date(r.created_at) <= new Date(p.to + 'T23:59:59'));
  return rows.slice(0, 500); // จำกัดจำนวนแถวต่อครั้งกันโหลดช้าเกินไปถ้าประวัติเยอะมาก
}

function action_listAuditLogRefTypes(p, user) {
  requireRole_(user, ['admin_super']);
  const types = [...new Set(readAll_('AuditLog').map(r => r.ref_type))].filter(Boolean);
  return types.sort();
}

// ---------- ประวัติการเปลี่ยนแปลง (History) แบบสรุปอ่านง่าย — ให้ทั้งสมาชิก (ดูของตัวเอง) และแอดมิน (ดูของสมาชิกคนไหนก็ได้) ----------
// ต่างจาก action_listAuditLog (ดิบ, admin_super เท่านั้น, ทุก ref_type ปนกัน) ตรงที่ฟังก์ชันนี้ "รวบรวม" ทุกรายการที่เกี่ยวกับสมาชิกคนเดียว
// จากหลายตาราง (Members, Applications, Welfare, Refunds, Licenses, Payments) มาเรียงเป็นไทม์ไลน์เดียว พร้อมแปลป้ายกำกับให้อ่านง่าย

const HISTORY_ACTION_LABELS = {
  self_update: 'สมาชิกแก้ไขข้อมูลของตนเอง',
  self_update_partial: 'สมาชิกแก้ไขข้อมูลบางส่วน (ระหว่างรอผลคำขอเปลี่ยนชื่อ)',
  admin_update: 'เจ้าหน้าที่แก้ไขข้อมูล',
  name_change_approved: 'อนุมัติคำขอเปลี่ยนชื่อ-นามสกุล',
  name_change_rejected: 'ไม่อนุมัติคำขอเปลี่ยนชื่อ-นามสกุล',
  name_change_more_docs_requested: 'ขอเอกสารเพิ่มเติม (คำขอเปลี่ยนชื่อ-นามสกุล)',
  name_change_docs_resubmitted: 'แนบเอกสารเพิ่มเติม (คำขอเปลี่ยนชื่อ-นามสกุล)',
  type_change_confirmed: 'เปลี่ยนประเภทสมาชิกสำเร็จ',
  membership_renewed: 'ต่ออายุสมาชิกภาพ',
  membership_confirmed: 'ยืนยันการเป็นสมาชิก (ชำระเงินสำเร็จ)',
  registration_confirmed: 'ขึ้นทะเบียนรับใบอนุญาตสำเร็จ',
  renewal_confirmed: 'ต่ออายุใบอนุญาตสำเร็จ',
  reject: 'ไม่อนุมัติคำขอ',
  approve_pending_payment: 'อนุมัติคำขอ (รอชำระเงิน)',
  payment_rejected: 'ปฏิเสธหลักฐานการชำระเงิน',
  slip_submitted: 'แนบหลักฐานการชำระเงิน',
  welfare_approved: 'อนุมัติคำขอรับสวัสดิการ',
  welfare_rejected: 'ไม่อนุมัติคำขอรับสวัสดิการ',
  welfare_disbursed: 'บันทึกการเบิกจ่ายเงินสวัสดิการ',
  refund_approved: 'อนุมัติคำขอคืนเงิน',
  refund_rejected: 'ไม่อนุมัติคำขอคืนเงิน',
  refund_completed: 'ยืนยันคืนเงินเรียบร้อยแล้ว'
};

/** แปลง actor_user_id ที่บันทึกไว้ให้เป็นข้อความอ่านง่าย — รองรับทั้งกรณีเป็น user_id จริง และกรณีข้อมูลเดโมที่ seed เป็น username ตรงๆ */
function historyActorLabel_(actorUserId, selfUserIds) {
  if (!actorUserId) return 'ระบบ';
  if (selfUserIds.indexOf(actorUserId) !== -1) return 'ตัวท่านเอง';
  const byId = findRowById_('Users', 'user_id', actorUserId);
  if (byId) return (ROLE_LABELS[byId.role] || byId.role) + ' (' + byId.username + ')';
  const byUsername = readAll_('Users').find(u => u.username === actorUserId);
  if (byUsername) return (ROLE_LABELS[byUsername.role] || byUsername.role) + ' (' + byUsername.username + ')';
  return 'เจ้าหน้าที่ (' + actorUserId + ')';
}

/** เทียบ old/new value เฉพาะฟิลด์ที่มีความหมายกับสมาชิก แล้วสรุปเป็นข้อความ "จาก ... เป็น ..." อ่านง่าย ไม่แสดง JSON ดิบ */
function summarizeHistoryChange_(oldVal, newVal) {
  const INTERESTING_FIELDS = {
    prefix_th: 'คำนำหน้า', first_name_th: 'ชื่อ', last_name_th: 'นามสกุล',
    member_type: 'ประเภทสมาชิก', member_status: 'สถานะสมาชิก',
    license_status: 'สถานะใบอนุญาต', membership_expire_date: 'วันหมดอายุสมาชิกภาพ',
    status: 'สถานะ', reject_reason: 'เหตุผล', approved_amount: 'จำนวนเงินที่อนุมัติ',
    reason: 'เหตุผล', new_expire_date: 'วันหมดอายุใหม่'
  };
  const parts = [];
  Object.keys(INTERESTING_FIELDS).forEach(f => {
    const ov = oldVal ? oldVal[f] : undefined;
    const nv = newVal ? newVal[f] : undefined;
    if (nv !== undefined && nv !== '' && ov !== nv) {
      parts.push(INTERESTING_FIELDS[f] + ': ' + (ov !== undefined && ov !== '' ? ('"' + ov + '" → ') : '') + '"' + nv + '"');
    }
  });
  return parts.join(' / ');
}

/** รวบรวมประวัติการเปลี่ยนแปลงทั้งหมดที่เกี่ยวกับสมาชิกคนหนึ่ง จากทุกตารางที่เกี่ยวข้อง มาเรียงเป็นไทม์ไลน์เดียว */
function getMemberChangeHistory_(memberId) {
  if (!memberId) return [];
  const usersOfMember = readAll_('Users').filter(u => u.member_id === memberId);
  const selfUserIds = usersOfMember.map(u => u.user_id);

  const apps = readAll_('Applications').filter(a => {
    let refMemberId = a.approved_member_id;
    if (!refMemberId) {
      try { refMemberId = JSON.parse(a.ref_data_json || '{}').member_id; } catch (e) { /* ignore */ }
    }
    return refMemberId === memberId || selfUserIds.indexOf(a.applicant_user_id) !== -1;
  }).map(a => a.application_id);

  const welfareIds = readAll_('Welfare').filter(w => w.member_id === memberId).map(w => w.welfare_id);
  const refundIds = readAll_('Refunds').filter(r => r.member_id === memberId).map(r => r.refund_id);
  const licenses = readAll_('Licenses').filter(l => l.member_id === memberId);
  const licenseRefIds = licenses.map(l => l.license_id).concat(licenses.map(l => l.license_no)).filter(Boolean);
  const paymentIds = readAll_('Payments').filter(pm => pm.member_id === memberId).map(pm => pm.payment_id);

  const matchSet = new Set([memberId].concat(apps, welfareIds, refundIds, licenseRefIds, paymentIds));

  const rows = readAll_('AuditLog')
    .filter(r => matchSet.has(r.ref_id))
    .sort((a, b) => new Date(b.created_at) - new Date(a.created_at))
    .slice(0, 200);

  return rows.map(r => {
    let oldVal = {}, newVal = {};
    try { oldVal = JSON.parse(r.old_value_json || '{}'); } catch (e) { /* ignore */ }
    try { newVal = JSON.parse(r.new_value_json || '{}'); } catch (e) { /* ignore */ }
    return {
      log_id: r.log_id, ref_type: r.ref_type, ref_id: r.ref_id,
      action: r.action, action_label: HISTORY_ACTION_LABELS[r.action] || r.action,
      summary: summarizeHistoryChange_(oldVal, newVal),
      actor_label: historyActorLabel_(r.actor_user_id, selfUserIds),
      created_at: r.created_at
    };
  });
}

/** ฝั่งสมาชิก — ดูประวัติการเปลี่ยนแปลงของตนเองแบบเรียลไทม์ (ทุกการแก้ไข/อนุมัติที่เกี่ยวกับตนเอง ไม่ว่าจะทำโดยตนเองหรือเจ้าหน้าที่) */
function action_getMyChangeHistory(p, user) {
  requireRole_(user, ['member']);
  if (!user.member_id) return [];
  return getMemberChangeHistory_(user.member_id);
}

/** ฝั่งแอดมิน — ดูประวัติการเปลี่ยนแปลงของสมาชิกคนไหนก็ได้ (ใช้ในหน้าข้อมูลสมาชิก) — เปิดให้ staff ทั่วไปดูได้ ไม่จำกัดเฉพาะ admin_super
 * เพราะเป็นประวัติที่กรองเฉพาะรายบุคคลแล้ว ต่างจาก action_listAuditLog ที่เป็น log ดิบทั้งระบบ (จำกัด admin_super)
 */
function action_getMemberChangeHistory(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff', 'admin_welfare']);
  if (!p.member_id) throw new Error('กรุณาระบุ member_id');
  return getMemberChangeHistory_(p.member_id);
}

// ---------- จัดการนายกสภาฯ/เลขาธิการฯ/เหรัญญิกฯ (ใช้แสดงในบัตรสมาชิก/บัตรใบอนุญาต/ใบอนุญาต/ใบรับรองสมาชิกวิสามัญ/ใบเสร็จ) ----------

const COUNCIL_OFFICER_POSITION_TYPES = ['นายก', 'เลขาธิการ', 'เหรัญญิก'];
const COUNCIL_OFFICER_POSITION_LABELS = {
  นายก: 'นายกสภาวิชาชีพสังคมสงเคราะห์',
  เลขาธิการ: 'เลขาธิการสภาวิชาชีพสังคมสงเคราะห์',
  เหรัญญิก: 'เหรัญญิกสภาวิชาชีพสังคมสงเคราะห์'
};

function action_listCouncilOfficers(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  return readAll_('CouncilOfficers').sort((a, b) => {
    if (a.position_type !== b.position_type) return a.position_type.localeCompare(b.position_type, 'th');
    return new Date(b.term_start) - new Date(a.term_start);
  });
}

/** เพิ่มวาระใหม่ให้นายก/เลขาธิการ/เหรัญญิก — จำกัดเฉพาะ admin_super เพราะมีผลต่อเอกสารทางการที่ออกโดยระบบทั้งหมด */
function action_addCouncilOfficer(p, user) {
  requireRole_(user, ['admin_super']);
  if (COUNCIL_OFFICER_POSITION_TYPES.indexOf(p.position_type) === -1) throw new Error('ตำแหน่งไม่ถูกต้อง');
  if (!p.name_th) throw new Error('กรุณากรอกชื่อ-นามสกุล (ไทย)');
  if (!p.term_start) throw new Error('กรุณาระบุวันที่เริ่มวาระ');
  const row = {
    officer_id: newId_('OFC'), position_type: p.position_type, name_th: p.name_th, name_en: p.name_en || '',
    position_label: COUNCIL_OFFICER_POSITION_LABELS[p.position_type],
    signature_url: p.signature_url || '', term_start: p.term_start, term_end: p.term_end || '',
    note: p.note || '', created_at: nowIso_(), updated_at: nowIso_()
  };
  appendRow_('CouncilOfficers', row);
  auditLog_('CouncilOfficers', row.officer_id, 'เพิ่มวาระใหม่ (' + p.position_type + ': ' + p.name_th + ')', {}, row, user.user_id);
  return row;
}

function action_updateCouncilOfficer(p, user) {
  requireRole_(user, ['admin_super']);
  const row = findRowById_('CouncilOfficers', 'officer_id', p.officer_id);
  if (!row) throw new Error('ไม่พบข้อมูลนี้');
  const patch = {
    name_th: p.name_th !== undefined ? p.name_th : row.name_th,
    name_en: p.name_en !== undefined ? p.name_en : row.name_en,
    signature_url: p.signature_url !== undefined ? p.signature_url : row.signature_url,
    term_start: p.term_start !== undefined ? p.term_start : row.term_start,
    term_end: p.term_end !== undefined ? p.term_end : row.term_end,
    note: p.note !== undefined ? p.note : row.note,
    updated_at: nowIso_()
  };
  updateRow_('CouncilOfficers', row._row, patch);
  auditLog_('CouncilOfficers', p.officer_id, 'แก้ไขข้อมูลวาระ', row, patch, user.user_id);
  return { ok: true };
}

function action_deleteCouncilOfficer(p, user) {
  requireRole_(user, ['admin_super']);
  const row = findRowById_('CouncilOfficers', 'officer_id', p.officer_id);
  if (!row) throw new Error('ไม่พบข้อมูลนี้');
  deleteRow_('CouncilOfficers', row._row);
  auditLog_('CouncilOfficers', p.officer_id, 'ลบวาระ (' + row.position_type + ': ' + row.name + ')', row, {}, user.user_id);
  return { ok: true };
}

// ---------- สมาชิกกิตติมศักดิ์ ----------
const HONORARY_QUALIFICATION_TYPES = [
  'ผู้ทรงคุณวุฒิ ผู้รอบรู้ด้านสังคมสงเคราะห์ศาสตร์',
  'ผู้ที่มีผลงานเชิงประจักษ์ด้านสังคมสงเคราะห์ศาสตร์'
];
const HONORARY_STATUSES = ['ดำรงสถานะ', 'สิ้นสุดสถานะ'];

/** เลขทะเบียนสมาชิกกิตติมศักดิ์ — เลขลำดับ 5 หลัก (ตามตารางรูปแบบเลขที่ TOR ข้อ 3.1) */
function nextHonoraryNo_() {
  return nextRunningSeq_('MembersHonorary', 'honorary_no');
}

function action_listHonoraryMembers(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff', 'admin_welfare']);
  const hUsers = readAll_('Users').filter(u => u.role === 'member' && u.member_id);
  return {
    items: readAll_('MembersHonorary').map(h => {
      const acc = hUsers.find(u => u.member_id === h.honorary_id);
      return Object.assign({}, h, { account_username: acc ? acc.username : '', account_status: acc ? acc.status : '' });
    }).sort((a, b) => String(b.honorary_no).localeCompare(String(a.honorary_no))),
    qualification_types: HONORARY_QUALIFICATION_TYPES,
    statuses: HONORARY_STATUSES
  };
}

function validateHonorary_(p) {
  if (!p.first_name_th || !p.last_name_th) throw new Error('กรุณากรอกชื่อและนามสกุล (ไทย)');
  if (HONORARY_QUALIFICATION_TYPES.indexOf(p.qualification_type) === -1) throw new Error('กรุณาเลือกคุณสมบัติของสมาชิกกิตติมศักดิ์');
  if (p.status && HONORARY_STATUSES.indexOf(p.status) === -1) throw new Error('สถานะไม่ถูกต้อง');
}

function action_addHonoraryMember(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  validateHonorary_(p);
  const row = {
    honorary_id: newId_('HNR'), honorary_no: nextHonoraryNo_(), prefix_th: p.prefix_th || '',
    first_name_th: p.first_name_th, last_name_th: p.last_name_th,
    first_name_en: p.first_name_en || '', last_name_en: p.last_name_en || '',
    qualification_type: p.qualification_type, expertise: p.expertise || '', position: p.position || '',
    organization: p.organization || '', phone: p.phone || '', email: p.email || '', photo_url: p.photo_url || '',
    appointed_date: p.appointed_date || '', term_end: p.term_end || '', status: p.status || 'ดำรงสถานะ',
    note: p.note || '', created_by: user.user_id, created_at: nowIso_(), updated_at: nowIso_()
  };
  appendRow_('MembersHonorary', row);
  auditLog_('MembersHonorary', row.honorary_id, 'เพิ่มสมาชิกกิตติมศักดิ์ (' + row.first_name_th + ' ' + row.last_name_th + ')', {}, row, user.user_id);
  return row;
}

function action_updateHonoraryMember(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const row = findRowById_('MembersHonorary', 'honorary_id', p.honorary_id);
  if (!row) throw new Error('ไม่พบข้อมูลสมาชิกกิตติมศักดิ์');
  const merged = Object.assign({}, row, p);
  validateHonorary_(merged);
  const fields = ['prefix_th', 'first_name_th', 'last_name_th', 'first_name_en', 'last_name_en', 'qualification_type', 'expertise',
    'position', 'organization', 'phone', 'email', 'photo_url', 'appointed_date', 'term_end', 'status', 'note'];
  const patch = { updated_at: nowIso_() };
  fields.forEach(f => { patch[f] = p[f] !== undefined ? p[f] : row[f]; });
  updateRow_('MembersHonorary', row._row, patch);
  auditLog_('MembersHonorary', p.honorary_id, 'แก้ไขข้อมูลสมาชิกกิตติมศักดิ์', row, patch, user.user_id);
  // สถานะสมาชิกเปลี่ยน → เปิด/ปิดบัญชีเข้าสู่ระบบให้สอดคล้อง และอัปเดตวันหมดอายุบนบัตร
  const acc = findHonoraryAccount_(p.honorary_id);
  if (acc) updateRow_('Users', acc._row, { status: patch.status === 'ดำรงสถานะ' ? 'active' : 'inactive' });
  const hc = readAll_('Cards').find(c => c.member_id === p.honorary_id);
  if (hc) updateRow_('Cards', hc._row, { valid_until: patch.term_end || '' });
  return { ok: true };
}

function findHonoraryById_(id) {
  if (!id) return null;
  try { return findRowById_('MembersHonorary', 'honorary_id', id); } catch (e) { return null; }
}

function findHonoraryAccount_(honoraryId) {
  return readAll_('Users').find(u => u.role === 'member' && u.member_id === honoraryId) || null;
}

/** แอดมินสร้างบัญชีเข้าสู่ระบบให้สมาชิกกิตติมศักดิ์ (ผูก member_id = honorary_id) — สมาชิกใช้ดู "ข้อมูลของฉัน" และ "บัตรสมาชิก" */
function action_createHonoraryAccount(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const h = findRowById_('MembersHonorary', 'honorary_id', p.honorary_id);
  if (!h) throw new Error('ไม่พบข้อมูลสมาชิกกิตติมศักดิ์');
  if (findHonoraryAccount_(h.honorary_id)) throw new Error('สมาชิกท่านนี้มีบัญชีแล้ว — ใช้ปุ่ม "รีเซ็ตรหัสผ่าน" แทน');
  const username = String(p.username || '').trim();
  if (!username) throw new Error('กรุณากรอกชื่อผู้ใช้งาน');
  if (findRowById_('Users', 'username', username)) throw new Error('ชื่อผู้ใช้งานนี้ถูกใช้แล้ว');
  validatePasswordPolicy_(p.password);
  const u = {
    user_id: newId_('USR'), username: username, password_hash: hashPassword_(p.password),
    role: 'member', member_id: h.honorary_id, thaid_pid: '',
    status: h.status === 'ดำรงสถานะ' ? 'active' : 'inactive',
    created_at: nowIso_(), last_login: '', failed_login_attempts: 0, locked_until: '',
    display_name: [h.prefix_th, h.first_name_th, h.last_name_th].filter(Boolean).join(' ')
  };
  appendRow_('Users', u);
  auditLog_('Users', u.user_id, 'สร้างบัญชีเข้าสู่ระบบให้สมาชิกกิตติมศักดิ์ ' + h.honorary_no, {}, { username: username, honorary_id: h.honorary_id }, user.user_id);
  return { ok: true, username: username };
}

/** แอดมินตั้งรหัสผ่านใหม่/ปลดล็อกบัญชีสมาชิกกิตติมศักดิ์ */
function action_resetHonoraryPassword(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const acc = findHonoraryAccount_(p.honorary_id);
  if (!acc) throw new Error('สมาชิกท่านนี้ยังไม่มีบัญชี');
  validatePasswordPolicy_(p.new_password);
  updateRow_('Users', acc._row, { password_hash: hashPassword_(p.new_password), failed_login_attempts: 0, locked_until: '' });
  auditLog_('Users', acc.user_id, 'รีเซ็ตรหัสผ่านสมาชิกกิตติมศักดิ์โดยเจ้าหน้าที่', {}, {}, user.user_id);
  return { ok: true };
}

/** สมาชิกกิตติมศักดิ์แก้ข้อมูลติดต่อของตนเองได้บางส่วน (ชื่อ/คุณสมบัติ/สถานะแก้ได้เฉพาะเจ้าหน้าที่) */
function action_updateHonoraryProfile(p, user) {
  requireRole_(user, ['member']);
  const h = findHonoraryById_(user.member_id);
  if (!h) throw new Error('บัญชีนี้ไม่ใช่สมาชิกกิตติมศักดิ์');
  const allowed = ['phone', 'email', 'photo_url', 'position', 'organization', 'expertise'];
  const patch = { updated_at: nowIso_() };
  allowed.forEach(f => { if (p.patch && p.patch[f] !== undefined) patch[f] = p.patch[f]; });
  updateRow_('MembersHonorary', h._row, patch);
  auditLog_('MembersHonorary', h.honorary_id, 'สมาชิกแก้ไขข้อมูลติดต่อของตนเอง', h, patch, user.user_id);
  return { ok: true };
}

function action_deleteHonoraryMember(p, user) {
  requireRole_(user, ['admin_super']);
  const row = findRowById_('MembersHonorary', 'honorary_id', p.honorary_id);
  if (!row) throw new Error('ไม่พบข้อมูลสมาชิกกิตติมศักดิ์');
  const delAcc = findHonoraryAccount_(p.honorary_id);
  if (delAcc) updateRow_('Users', delAcc._row, { status: 'inactive' }); // ปิดบัญชี ไม่ลบประวัติ
  deleteRow_('MembersHonorary', row._row);
  auditLog_('MembersHonorary', p.honorary_id, 'ลบสมาชิกกิตติมศักดิ์ (' + row.first_name_th + ' ' + row.last_name_th + ')', row, {}, user.user_id);
  return { ok: true };
}

// ---------- Backup (Export) ----------
// หมายเหตุ: สำหรับ "Restore" แนะนำให้ใช้ฟีเจอร์ "ดูประวัติเวอร์ชัน" (Version History) ของ Google Sheets เอง
// (File > ดูประวัติเวอร์ชัน) ซึ่งปลอดภัยกว่าการเขียนฟังก์ชันเขียนทับข้อมูลทั้งหมดเอง — ดู SECURITY_SYSTEM.md

/** ส่งออกข้อมูลทุกตารางเป็น JSON ไฟล์เดียว ไว้สำรองข้อมูล (Backup) — จำกัดเฉพาะ admin_super */
function action_exportFullBackup(p, user) {
  requireRole_(user, ['admin_super']);
  const ss = getSS_();
  const backup = { exportedAt: nowIso_(), exportedBy: user.user_id, sheets: {} };
  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    if (name === 'Sessions') return; // ไม่ต้องสำรอง session token ชั่วคราว
    const values = sh.getDataRange().getValues();
    backup.sheets[name] = values;
  });

  const blob = Utilities.newBlob(JSON.stringify(backup), 'application/json', 'backup_' + nowIso_().slice(0, 10) + '.json');
  const folder = getDocsFolder_();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  auditLog_('System', 'backup', 'ส่งออกข้อมูลสำรองทั้งระบบ', {}, {}, user.user_id);
  return { url: file.getUrl() };
}

// ============================================================
// ระบบชำระเงิน (Payment Gateway) — TOR ข้อ 11
// ============================================================
// สถานะการเชื่อมต่อ: ยังไม่ได้เชื่อมต่อ Payment Gateway จริง เพราะ TOR ขอให้ระบุผู้ให้บริการก่อนเริ่มพัฒนา
// (เช่น K-Payment, 2C2P, GB PromptPay) แต่ยังไม่มีการระบุ/ให้ข้อมูล credential มาในงานนี้ — ดู PAYMENT_SYSTEM.md
// สิ่งที่ทำงานได้จริงแล้วในรอบนี้ (ไม่ต้องรอเชื่อมต่อ Gateway จริง): ใบเสร็จอิเล็กทรอนิกส์+ตรวจสอบ, ตัวรับ Webhook,
// ระบบคืนเงิน (workflow ภายใน), เครื่องมือกระทบยอดรายรับ (นำเข้ารายงานจาก Gateway เป็น CSV ได้)

/** เริ่มชำระเงินผ่าน Payment Gateway จริง (บัตร/พร้อมเพย์/Mobile Banking) — STUB
 * ยังไม่เชื่อมต่อจริง เพราะยังไม่ได้เลือกผู้ให้บริการ+ยังไม่มี credential ให้ในงานนี้
 */
function action_initiatePaymentGateway(p, user) {
  throw new Error('ยังไม่ได้เชื่อมต่อ Payment Gateway จริง — TOR ขอให้ระบุผู้ให้บริการ (เช่น K-Payment, 2C2P, GB PromptPay) ก่อนเริ่มพัฒนา ' +
    'แต่ยังไม่ได้รับข้อมูลนี้ ระหว่างนี้ใช้วิธีโอนเงินผ่านธนาคารแล้วแนบสลิปให้เจ้าหน้าที่ตรวจสอบแทนไปก่อน (ดู PAYMENT_SYSTEM.md)');
}

/** ตัวรับ Webhook/Callback จาก Payment Gateway — เรียกจาก doPost เมื่อพบ query param ?webhook=true
 * โครงสร้าง payload เป็นแบบทั่วไป (generic) รอปรับให้ตรงกับผู้ให้บริการจริงเมื่อเลือกแล้ว (ดู PAYMENT_SYSTEM.md)
 * คาดหวัง payload อย่างน้อย: { payment_id หรือ gateway_ref, status: 'success'|'failed', amount, transaction_ref }
 */
function handlePaymentWebhook_(rawBody) {
  const logId = newId_('WHK');
  let payload = {};
  let result = 'ok';
  let matchedPaymentId = '';
  try {
    payload = JSON.parse(rawBody || '{}');
    const payment = payload.payment_id
      ? findRowById_('Payments', 'payment_id', payload.payment_id)
      : readAll_('Payments').find(pm => pm.gateway_ref === payload.gateway_ref);

    if (!payment) {
      result = 'ไม่พบรายการชำระเงินที่ตรงกับ webhook นี้ (payment_id/gateway_ref ไม่ตรง)';
    } else if (payload.status === 'success') {
      updateRow_('Payments', payment._row, {
        status: 'ชำระแล้ว', paid_at: nowIso_(), gateway_ref: payload.gateway_ref || payment.gateway_ref || '',
        gateway_raw_json: JSON.stringify(payload)
      });
      matchedPaymentId = payment.payment_id;
      try { issueReceiptForPayment_(payment.payment_id, 'ระบบ (Webhook อัตโนมัติ)'); } catch (e) { /* ออกใบเสร็จไม่สำเร็จไม่ทำให้ webhook ล้มเหลวทั้งหมด */ }
      result = 'อัปเดตสถานะสำเร็จ + ออกใบเสร็จอัตโนมัติ';
    } else {
      updateRow_('Payments', payment._row, { status: 'ปฏิเสธ', gateway_raw_json: JSON.stringify(payload) });
      matchedPaymentId = payment.payment_id;
      result = 'บันทึกสถานะปฏิเสธ/ล้มเหลวจาก Gateway';
    }
  } catch (err) {
    result = 'เกิดข้อผิดพลาด: ' + err.message;
  }
  appendRow_('WebhookLog', { log_id: logId, received_at: nowIso_(), raw_payload: rawBody || '', processed: 'true', payment_id: matchedPaymentId, result: result });
  return { ok: true };
}

// ---------- ใบเสร็จรับเงินอิเล็กทรอนิกส์ ----------

/** ออกใบเสร็จ PDF ให้ Payment ที่ชำระแล้ว — เรียกอัตโนมัติทุกครั้งที่มีการยืนยันชำระเงิน (ทั้งจากแอดมินตรวจสลิป และจาก webhook)
 * และเรียกซ้ำได้ (ออกใหม่/พิมพ์ซ้ำ) จากฝั่งแอดมิน
 */
/** สร้างที่อยู่แบบเต็มจาก field ที่อยู่ปัจจุบันของสมาชิก (ใช้ในใบเสร็จ) */
function formatMemberAddress_(member) {
  if (!member || member.org_name_th) return ''; // องค์กรไม่ใช้ที่อยู่แบบบุคคล
  const parts = [
    member.addr_current_house_no ? 'เลขที่ ' + member.addr_current_house_no : '',
    member.addr_current_moo ? 'หมู่ที่ ' + member.addr_current_moo : '',
    member.addr_current_soi ? 'ซอย' + member.addr_current_soi : '',
    member.addr_current_road ? 'ถนน' + member.addr_current_road : '',
    member.addr_current_subdistrict ? 'ตำบล/แขวง' + member.addr_current_subdistrict : '',
    member.addr_current_district ? 'อำเภอ/เขต' + member.addr_current_district : '',
    member.addr_current_province ? 'จังหวัด' + member.addr_current_province : '',
    member.addr_current_zipcode || ''
  ].filter(Boolean);
  return parts.join(' ');
}

/** รายการช่องทำเครื่องหมาย "เป็นค่า" ในใบเสร็จ — คืนค่ารายการที่ต้อง ☑ ตามประเภทของรายการที่ชำระ พร้อมข้อความเต็มของแต่ละบรรทัด (ตรงตามแบบฟอร์มใบเสร็จจริง) */
function getReceiptCheckboxLines_(payment, app, member) {
  const lines = [
    { key: 'entrance', text: 'สมาชิกครั้งแรกเข้า ' + MEMBERSHIP_ENTRANCE_FEE + ' บาท' },
    { key: 'assoc1', text: 'ค่าบำรุงสมาชิกสมทบ 1 ปี ' + MEMBERSHIP_FEE_TABLE['สมทบ']['1'] + ' บาท' },
    { key: 'assoc5', text: 'ค่าบำรุงสมาชิกสมทบ 5 ปี ' + MEMBERSHIP_FEE_TABLE['สมทบ']['5'] + ' บาท' },
    { key: 'ord1', text: 'ค่าบำรุงสมาชิกสามัญ 1 ปี ' + MEMBERSHIP_FEE_TABLE['สามัญ']['1'] + ' บาท' },
    { key: 'ord5', text: 'ค่าบำรุงสมาชิกสามัญ 5 ปี ' + MEMBERSHIP_FEE_TABLE['สามัญ']['5'].toLocaleString() + ' บาท' },
    { key: 'ordLife', text: 'ค่าบำรุงสมาชิกสามัญ ตลอดชีพ ' + MEMBERSHIP_FEE_TABLE['สามัญ']['lifetime'].toLocaleString() + ' บาท (อายุ ' + LIFETIME_MIN_AGE + ' ปีขึ้นไป)' },
    { key: 'assoc_org1', text: 'ค่าบำรุงสมาชิกวิสามัญ 1 ปี ' + MEMBERSHIP_FEE_TABLE['วิสามัญ']['1'] + ' บาท' },
    { key: 'assoc_org5', text: 'ค่าบำรุงสมาชิกวิสามัญ 5 ปี ' + MEMBERSHIP_FEE_TABLE['วิสามัญ']['5'].toLocaleString() + ' บาท' },
    { key: 'certMailing', text: 'ค่าจัดส่งหนังสือรับรองการเป็นสมาชิก ' + CERTIFICATE_MAILING_FEE + ' บาท' },
    { key: 'examSubject1', text: 'ค่าธรรมเนียมสอบข้อเขียน ' + EXAM_SUBJECTS[0] + ' ' + EXAM_SUBJECT_FEE + ' บาท' },
    { key: 'examSubject2', text: 'ค่าธรรมเนียมสอบข้อเขียน ' + EXAM_SUBJECTS[1] + ' ' + EXAM_SUBJECT_FEE + ' บาท' },
    { key: 'licenseReg', text: 'ค่าขึ้นทะเบียนรับใบอนุญาตเป็นผู้ประกอบวิชาชีพฯ ' + LICENSE_REGISTRATION_FEE.toLocaleString() + ' บาท' },
    { key: 'licenseRenew', text: 'ค่าต่ออายุใบอนุญาตฯ ' + LICENSE_RENEWAL_FEE.toLocaleString() + ' บาท' },
    { key: 'licenseRenewAfter', text: 'ค่าต่ออายุหลังใบอนุญาตหมดอายุ ' + LICENSE_RENEWAL_FEE_AFTER_EXPIRY.toLocaleString() + ' บาท' },
    { key: 'other', text: 'อื่นๆ ระบุ: ' + (payment.ref_type || '') + ' ' + (payment.ref_id || '') }
  ];
  const checked = {};
  const type = app ? app.applicant_type : '';
  const refData = app && app.ref_data_json ? JSON.parse(app.ref_data_json || '{}') : {};
  const duration = refData.duration || '5';
  const durationKey = duration === 'lifetime' ? 'Life' : duration;

  if (type === 'สามัญ') { checked.entrance = true; checked['ord' + durationKey] = true; }
  else if (type === 'สมทบ') { checked.entrance = true; checked['assoc' + durationKey] = true; }
  else if (type === 'วิสามัญ') { checked.entrance = true; checked['assoc_org' + durationKey] = true; checked.certMailing = true; }
  else if (type === 'เปลี่ยนประเภท') { checked['ord' + durationKey] = true; } // ไม่เก็บค่าแรกเข้าซ้ำ
  else if (type === 'ต่ออายุสมาชิก') {
    const memberType = member ? member.member_type : 'สามัญ';
    checked[(memberType === 'สามัญ' ? 'ord' : 'assoc') + durationKey] = true;
  } else if (type === 'ขึ้นทะเบียน') { checked.licenseReg = true; }
  else if (type === 'ต่ออายุใบอนุญาต') {
    if (refData.method === 'exam') { checked.licenseReg = true; } // ค่าสมัครสอบข้อเขียนเก็บแยกต่างหากตอนสมัครสอบแล้ว ไม่เก็บซ้ำในใบเสร็จขั้นตอนนี้
    else if (refData.isAfterExpiry) checked.licenseRenewAfter = true;
    else checked.licenseRenew = true;
  } else { checked.other = true; }

  return lines.map(l => ({ text: l.text, checked: !!checked[l.key] }));
}

function issueReceiptForPayment_(paymentId, issuedByLabel) {
  const payment = findRowById_('Payments', 'payment_id', paymentId);
  if (!payment) throw new Error('ไม่พบรายการชำระเงิน');
  if (payment.status !== 'ชำระแล้ว') throw new Error('ออกใบเสร็จได้เฉพาะรายการที่ชำระเงินสำเร็จแล้วเท่านั้น');

  const member = findRowById_('Members', 'member_id', payment.member_id) || findRowById_('MembersOrg', 'org_id', payment.member_id);
  const memberName = member ? ([member.prefix_th, member.first_name_th, member.last_name_th].filter(Boolean).join(' ') || member.org_name_th) : payment.member_id;
  const memberAddress = formatMemberAddress_(member);
  const app = payment.ref_type === 'application' ? findRowById_('Applications', 'application_id', payment.ref_id) : null;

  const receiptNo = payment.receipt_no || generateReceiptNumber_();
  const paidDate = payment.paid_at || nowIso_();
  const treasurer = getTreasurerForDate_(paidDate);

  const doc = DocumentApp.create('ใบเสร็จรับเงิน_' + receiptNo);
  const body = doc.getBody();
  body.setPageWidth(792).setPageHeight(612); // แนวนอน (Letter landscape) ให้ตรงกับรูปแบบใบอนุญาต/ใบรับรองสมาชิกวิสามัญ เพื่อความสวยงามสม่ำเสมอกันทั้งระบบ

  body.appendParagraph('เลขที่ ' + receiptNo).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  // ตราสภาฯ — ดึงจากเว็บไซต์จริงมาแทรกด้านบน ให้ตรงกับแบบฟอร์มใบเสร็จจริง
  try {
    const frontendBase = PropertiesService.getScriptProperties().getProperty('FRONTEND_BASE_URL') || '';
    if (frontendBase) {
      const logoResp = UrlFetchApp.fetch(frontendBase + 'logo.png', { muteHttpExceptions: true });
      if (logoResp.getResponseCode() === 200) {
        const logoPara = body.appendParagraph('');
        logoPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        logoPara.appendInlineImage(logoResp.getBlob()).setWidth(60).setHeight(60);
      }
    }
  } catch (err) { /* ถ้าดึงตราไม่สำเร็จ ข้ามไป ไม่ทำให้การออกใบเสร็จล้มเหลวทั้งหมด */ }

  body.appendParagraph('สภาวิชาชีพสังคมสงเคราะห์').setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('สภาวิชาชีพสังคมสงเคราะห์ เลขที่ 255 ตึกกรมกิจการสตรีและสถาบันครอบครัว แขวงทุ่งพญาไท เขตราชเทวี กรุงเทพมหานคร 10400')
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(10);
  body.appendParagraph('โทรศัพท์ 02-050-9427 และโทรสาร 02-354-5176 Website: www.swpc.or.th E-mail: thaisocialworkcouncil@gmail.com')
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(10);
  body.appendParagraph('เลขประจำตัวผู้เสียภาษีอากร 099-4-00097599-6').setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(10);
  body.appendParagraph(' ');
  body.appendParagraph('ใบเสร็จรับเงิน').setHeading(DocumentApp.ParagraphHeading.SUBTITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('วัน ' + new Date(paidDate).getDate() + ' เดือน ' + THAI_MONTHS_[new Date(paidDate).getMonth()] + ' พ.ศ. ' + (new Date(paidDate).getFullYear() + 543)).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph(' ');
  body.appendParagraph('ได้รับเงินจาก ' + memberName);
  if (memberAddress) body.appendParagraph('ที่อยู่ ' + memberAddress).setFontSize(10);
  body.appendParagraph(' ');
  body.appendParagraph('เป็นค่า:').setBold(true);
  getReceiptCheckboxLines_(payment, app, member).forEach(line => {
    body.appendParagraph((line.checked ? '☑ ' : '☐ ') + line.text);
  });
  body.appendParagraph(' ');
  body.appendParagraph('รวมเป็นเงิน ' + Number(payment.amount).toLocaleString() + ' บาท (ตัวอักษร ' + thaiBahtText_(payment.amount) + ')').setBold(true);
  body.appendParagraph('ช่องทางการชำระเงิน: ' + (payment.channel || payment.method || 'โอนเงินผ่านธนาคาร')).setFontSize(10);
  if (payment.gateway_ref) body.appendParagraph('เลขอ้างอิงจาก Payment Gateway: ' + payment.gateway_ref).setFontSize(10);
  body.appendParagraph(' ');

  const sigTable = body.appendTable([['', '']]);
  sigTable.setBorderWidth(0);
  const treasurerCell = sigTable.getCell(0, 0);
  const treasurerFirstPara = treasurerCell.getChild(0).asParagraph();
  treasurerFirstPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  let treasurerSigInserted = false;
  if (treasurer.signature_url) {
    try {
      const sigBlob = getDriveBlobFromUrl_(treasurer.signature_url);
      if (sigBlob) { treasurerFirstPara.appendInlineImage(sigBlob).setWidth(100).setHeight(40); treasurerSigInserted = true; }
    } catch (e) { /* ถ้าแนบรูปลายเซ็นไม่สำเร็จ ใช้ข้อความเดิมแทน ไม่ทำให้การออกใบเสร็จล้มเหลว */ }
  }
  if (!treasurerSigInserted) treasurerFirstPara.setText('(ลงชื่อ)........................................เหรัญญิก');
  treasurerCell.appendParagraph('(' + treasurer.name + ')').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  const receiverCell = sigTable.getCell(0, 1);
  const receiverFirstPara = receiverCell.getChild(0).asParagraph();
  receiverFirstPara.setText('(ลงชื่อ)........................................ผู้รับเงิน');
  receiverFirstPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  receiverCell.appendParagraph('(' + issuedByLabel + ')').setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph(' ');
  body.appendParagraph('เอกสารนี้ได้จัดทำและส่งข้อมูลให้แก่กรมสรรพากรด้วยวิธีการทางอิเล็กทรอนิกส์')
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(9);
  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  const pdfBlob = file.getAs('application/pdf');
  const folder = getDocsFolder_();
  const pdfFile = folder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  file.setTrashed(true);

  updateRow_('Payments', payment._row, { receipt_no: receiptNo, receipt_url: pdfFile.getUrl(), receipt_issued_at: nowIso_() });
  return { receipt_no: receiptNo, receipt_url: pdfFile.getUrl() };
}

function generateReceiptNumber_() {
  const prefix = 'RCT';
  const year = new Date().getFullYear() + 543;
  const existing = readAll_('Payments').filter(pm => String(pm.receipt_no || '').indexOf(prefix + year) === 0);
  let max = 0;
  existing.forEach(pm => {
    const n = parseInt(String(pm.receipt_no).slice((prefix + year).length + 1), 10);
    if (!isNaN(n) && n > max) max = n;
  });
  return prefix + year + '-' + String(max + 1).padStart(5, '0');
}

/** ออกใบเสร็จใหม่/พิมพ์ซ้ำ — ฝั่งแอดมิน */
function action_generateReceipt(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  return issueReceiptForPayment_(p.payment_id, user.user_id);
}

/** สมาชิกดาวน์โหลดใบเสร็จของตนเอง — ถ้ายังไม่เคยออก จะออกให้อัตโนมัติตอนนี้เลย (เผื่อกรณีเก่าที่ชำระผ่านแอดมินตรวจสลิปแบบเดิมก่อนมีระบบใบเสร็จ) */
function action_downloadMyReceipt(p, user) {
  requireRole_(user, ['member']);
  const payment = findRowById_('Payments', 'payment_id', p.payment_id);
  if (!payment) throw new Error('ไม่พบรายการชำระเงิน');
  if (payment.member_id !== user.member_id) throw new Error('ไม่มีสิทธิ์เข้าถึงใบเสร็จนี้');
  if (payment.receipt_url) return { receipt_no: payment.receipt_no, receipt_url: payment.receipt_url };
  return issueReceiptForPayment_(p.payment_id, 'ระบบ (ออกย้อนหลัง)');
}

function action_myReceipts(p, user) {
  requireRole_(user, ['member']);
  if (!user.member_id) throw new Error('บัญชีนี้ยังไม่ผูกกับข้อมูลสมาชิก');
  return readAll_('Payments').filter(pm => pm.member_id === user.member_id && pm.status === 'ชำระแล้ว')
    .sort((a, b) => new Date(b.paid_at || 0) - new Date(a.paid_at || 0));
}

/** ตรวจสอบใบเสร็จรับเงินอิเล็กทรอนิกส์ที่ออก — PUBLIC ไม่ต้อง login (TOR ข้อ 11 "ระบบตรวจสอบใบเสร็จ") */
function action_verifyReceipt(p) {
  const payment = readAll_('Payments').find(pm => pm.receipt_no === p.receipt_no);
  if (!payment) return { valid: false, message: 'ไม่พบเลขที่ใบเสร็จนี้ในระบบ' };
  const member = findRowById_('Members', 'member_id', payment.member_id) || findRowById_('MembersOrg', 'org_id', payment.member_id);
  const memberName = member ? ([member.prefix_th, member.first_name_th, member.last_name_th].filter(Boolean).join(' ') || member.org_name_th) : '-';
  return {
    valid: true, receipt_no: payment.receipt_no, name: memberName, amount: payment.amount,
    ref_type: payment.ref_type, paid_at: payment.paid_at, channel: payment.channel || payment.method
  };
}

// ---------- ระบบคืนเงิน (Refund) ----------

/** ยื่นคำขอคืนเงิน — ต้องผ่านตรวจสอบ+อนุมัติจากเจ้าหน้าที่การเงินก่อนดำเนินการคืนเงินทุกครั้ง (TOR ข้อ 11) */
function action_submitRefundRequest(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const payment = findRowById_('Payments', 'payment_id', p.payment_id);
  if (!payment) throw new Error('ไม่พบรายการชำระเงิน');
  if (payment.status !== 'ชำระแล้ว') throw new Error('คืนเงินได้เฉพาะรายการที่ชำระเงินสำเร็จแล้วเท่านั้น');
  const existing = readAll_('Refunds').find(r => r.payment_id === p.payment_id && ['รอตรวจสอบ', 'อนุมัติ'].indexOf(r.status) !== -1);
  if (existing) throw new Error('มีคำขอคืนเงินสำหรับรายการนี้อยู่แล้ว (เลขที่ ' + existing.refund_id + ')');

  const refundId = newId_('RFD');
  appendRow_('Refunds', {
    refund_id: refundId, payment_id: p.payment_id, member_id: payment.member_id,
    reason: p.reason || '', requested_by: user.user_id, requested_at: nowIso_(),
    status: 'รอตรวจสอบ', approved_by: '', approved_at: '', refund_amount: p.refund_amount || payment.amount, refund_note: '', refunded_at: ''
  });
  return { refund_id: refundId };
}

function action_listRefunds(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  let rows = readAll_('Refunds');
  if (p && p.status) rows = rows.filter(r => r.status === p.status);
  return rows.sort((a, b) => new Date(b.requested_at) - new Date(a.requested_at));
}

/** เจ้าหน้าที่การเงินตรวจสอบ+อนุมัติ/ปฏิเสธคำขอคืนเงิน */
function action_reviewRefund(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const refund = findRowById_('Refunds', 'refund_id', p.refund_id);
  if (!refund) throw new Error('ไม่พบคำขอคืนเงินนี้');
  if (refund.status !== 'รอตรวจสอบ') throw new Error('คำขอนี้ผ่านการพิจารณาไปแล้ว');

  if (p.decision === 'reject') {
    const patch = { status: 'ปฏิเสธ', approved_by: user.user_id, approved_at: nowIso_(), refund_note: p.note || '' };
    updateRow_('Refunds', refund._row, patch);
    auditLog_('Refunds', refund.refund_id, 'refund_rejected', refund, patch, user.user_id);
    return { ok: true };
  }
  const patch = {
    status: 'อนุมัติ', approved_by: user.user_id, approved_at: nowIso_(),
    refund_amount: p.refund_amount || refund.refund_amount, refund_note: p.note || ''
  };
  updateRow_('Refunds', refund._row, patch);
  auditLog_('Refunds', refund.refund_id, 'refund_approved', refund, patch, user.user_id);
  return { ok: true };
}

/** ยืนยันว่าคืนเงินจริงเรียบร้อยแล้ว (ทำรายการคืนเงินจริงผ่านหน้าเว็บของ Payment Gateway/ธนาคารเองนอกระบบ แล้วมาบันทึกยืนยันที่นี่)
 * เหตุผลที่แยกขั้นตอนนี้ออกจาก "อนุมัติ": TOR ต้องการให้อนุมัติก่อน แล้วจึง "ดำเนินการคืนเงิน" ซึ่งเป็นคนละขั้นตอนกัน
 * และเนื่องจากยังไม่ได้เชื่อมต่อ Payment Gateway จริง การคืนเงินจริงจึงต้องทำผ่านช่องทางของผู้ให้บริการเองในตอนนี้
 */
function action_markRefundCompleted(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const refund = findRowById_('Refunds', 'refund_id', p.refund_id);
  if (!refund) throw new Error('ไม่พบคำขอคืนเงินนี้');
  if (refund.status !== 'อนุมัติ') throw new Error('ต้องอนุมัติคำขอก่อนจึงจะยืนยันการคืนเงินได้');
  updateRow_('Refunds', refund._row, { status: 'คืนเงินแล้ว', refunded_at: nowIso_() });
  auditLog_('Refunds', refund.refund_id, 'refund_completed', refund, { status: 'คืนเงินแล้ว' }, user.user_id);

  const payment = findRowById_('Payments', 'payment_id', refund.payment_id);
  if (payment) updateRow_('Payments', payment._row, { status: 'คืนเงินแล้ว' });
  // ถ้าเป็นการคืนเงินค่าสมัครสอบข้อเขียน -> ยกเลิกใบสมัครสอบนั้นด้วย (ตัดสิทธิ์เข้าสอบ)
  if (payment && payment.ref_type === 'application') {
    const examApp = findRowById_('Applications', 'application_id', payment.ref_id);
    if (examApp && examApp.applicant_type === 'สมัครสอบข้อเขียน' && examApp.status === 'รอผลสอบ') {
      updateRow_('Applications', examApp._row, { status: 'ยกเลิก (คืนเงินแล้ว)' });
      auditLog_('Applications', examApp.application_id, 'exam_application_cancelled_refunded', examApp, { refund_id: refund.refund_id }, user.user_id);
      try { mailNotify_(JSON.parse(examApp.ref_data_json || '{}').email, 'ยกเลิกใบสมัครสอบข้อเขียนและคืนเงินแล้ว', 'ใบสมัครเลขที่ ' + examApp.application_id + ' ถูกยกเลิก และสภาฯ ดำเนินการคืนเงินค่าสมัครสอบเรียบร้อยแล้ว'); } catch (e) {}
    }
  }
  return { ok: true };
}

// ---------- ระบบกระทบยอดรายรับ (Reconciliation) ----------

/** รายการชำระเงินของระบบในวันที่กำหนด — ใช้เทียบกับรายงานจาก Payment Gateway (TOR ข้อ 11) */
function action_getReconciliationData(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const date = p.date; // YYYY-MM-DD
  const rows = readAll_('Payments').filter(pm => pm.status === 'ชำระแล้ว' && pm.paid_at && new Date(pm.paid_at).toISOString().slice(0, 10) === date);
  return rows.map(pm => ({
    payment_id: pm.payment_id, amount: pm.amount, channel: pm.channel || pm.method,
    gateway_ref: pm.gateway_ref || '', paid_at: pm.paid_at, reconciled: isTrue_(pm.reconciled)
  }));
}

/** นำเข้ารายงานธุรกรรมจาก Payment Gateway (CSV) มาจับคู่อัตโนมัติกับรายการในระบบ (ตามจำนวนเงิน+เลขอ้างอิงถ้ามี)
 * p.date, p.rows: [{ amount, gateway_ref }]
 */
function action_importGatewayStatementCsv(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const systemRows = readAll_('Payments').filter(pm => pm.status === 'ชำระแล้ว' && pm.paid_at && new Date(pm.paid_at).toISOString().slice(0, 10) === p.date);
  let matched = 0, unmatched = [];

  (p.rows || []).forEach(gwRow => {
    let match = null;
    if (gwRow.gateway_ref) match = systemRows.find(pm => pm.gateway_ref === gwRow.gateway_ref && !isTrue_(pm.reconciled));
    if (!match) match = systemRows.find(pm => Number(pm.amount) === Number(gwRow.amount) && !isTrue_(pm.reconciled));
    if (match) {
      updateRow_('Payments', match._row, { reconciled: 'true', reconciled_at: nowIso_(), gateway_ref: match.gateway_ref || gwRow.gateway_ref || '' });
      matched++;
    } else {
      unmatched.push(gwRow);
    }
  });

  const unreconciledSystemRows = systemRows.filter(pm => !isTrue_(pm.reconciled)).length;

  appendRow_('ReconciliationLog', {
    recon_id: newId_('RCN'), recon_date: p.date, source: 'CSV จาก Payment Gateway',
    matched_count: matched, unmatched_count: unmatched.length,
    performed_by: user.user_id, performed_at: nowIso_(), notes: p.notes || ''
  });

  return { matched: matched, unmatchedFromGateway: unmatched, unreconciledInSystem: unreconciledSystemRows };
}

function action_listReconciliationLog(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  return readAll_('ReconciliationLog').sort((a, b) => new Date(b.performed_at) - new Date(a.performed_at));
}

/** กระทบยอดด้วยมือทีละรายการ (กรณีไม่ได้นำเข้าไฟล์ CSV) */
function action_markPaymentReconciled(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const payment = findRowById_('Payments', 'payment_id', p.payment_id);
  if (!payment) throw new Error('ไม่พบรายการชำระเงิน');
  updateRow_('Payments', payment._row, { reconciled: p.reconciled ? 'true' : 'false', reconciled_at: p.reconciled ? nowIso_() : '' });
  return { ok: true };
}

// ---------- 🧩 โมดูลที่ยังไม่พัฒนาเต็มรูปแบบ — โครง endpoint ไว้ให้ต่อ ----------
// แท็บ Sheet มีให้แล้วใน SHEET_SCHEMA.md (Licenses, Cards, Welfare, Payments)
// ตัวอย่างการต่อยอด: ทำตามแพทเทิร์นของ action_listMembers / action_submitApplication ด้านบน

// ---------- 3. ระบบทะเบียนสมาชิก — ฝั่งแอดมิน (TOR ข้อ 3) ----------

/** รายการทะเบียนสมาชิกรวม (TOR ข้อ 3.1-3.2) — ใช้แสดงตารางหลักของหน้าทะเบียน/ใบอนุญาต */
function action_listRegistry(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const members = readAll_('Members');
  const licenses = readAll_('Licenses');
  return members.map(m => {
    const lic = licenses.find(l => l.member_id === m.member_id);
    return {
      member_id: m.member_id, member_no: m.member_no,
      name: [m.prefix_th, m.first_name_th, m.last_name_th].filter(Boolean).join(' '),
      photo_url: m.photo_url || '',
      member_type: m.member_type, member_status: m.member_status,
      applied_date: m.applied_date, membership_expire_date: m.membership_expire_date,
      license_no: m.license_no,
      license_status: lic ? lic.status : m.license_status,
      license_issue_date: lic ? lic.issue_date : '',
      license_expire_date: lic ? lic.expire_date : '',
      days_until_expire: daysUntil_(m.membership_expire_date),
      license_days_until_expire: lic ? daysUntil_(lic.expire_date) : null
    };
  });
}

/** ทะเบียนของสมาชิกรายคน: โปรไฟล์ + ใบอนุญาต + ประวัติต่ออายุ + ประวัติชำระเงิน (TOR ข้อ 3.1) */
function action_getMemberRegistry(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const member = findRowById_('Members', 'member_id', p.member_id);
  if (!member) throw new Error('ไม่พบสมาชิก');
  const license = readAll_('Licenses').find(l => l.member_id === p.member_id);
  const renewals = readAll_('MembershipRenewals').filter(r => r.member_id === p.member_id)
    .sort((a, b) => new Date(b.renewed_at) - new Date(a.renewed_at));
  const payments = readAll_('Payments').filter(pm => pm.member_id === p.member_id)
    .sort((a, b) => new Date(b.paid_at || 0) - new Date(a.paid_at || 0));
  const notifications = readAll_('Notifications').filter(n => n.member_id === p.member_id)
    .sort((a, b) => new Date(b.sent_at) - new Date(a.sent_at));
  const typeChangeHistory = readAll_('AuditLog')
    .filter(l => l.ref_type === 'Members' && l.ref_id === p.member_id && l.action === 'type_change_confirmed')
    .sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
  return { member, license, renewals, payments, notifications, typeChangeHistory };
}

/** ต่ออายุสมาชิกภาพด้วยตนเอง (บันทึกโดยเจ้าหน้าที่) — TOR ข้อ 3.1 "ประวัติการต่ออายุ"
 * หมายเหตุ: นี่คือการบันทึกต่ออายุแบบเจ้าหน้าที่ดำเนินการเอง (manual) ยังไม่ใช่ระบบสมาชิกยื่นขอต่ออายุออนไลน์เต็มรูปแบบตาม TOR ข้อ 4.2 (ดู ROADMAP.md)
 */
function action_renewMembership(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const member = findRowById_('Members', 'member_id', p.member_id);
  if (!member) throw new Error('ไม่พบสมาชิก');
  const years = parseInt(p.years, 10) || 1;
  const baseDate = member.membership_expire_date && new Date(member.membership_expire_date) > new Date()
    ? member.membership_expire_date : nowIso_();
  const newExpire = addYearsIso_(baseDate, years);

  updateRow_('Members', member._row, { membership_expire_date: newExpire, member_status: 'ปกติ', updated_at: nowIso_() });

  const license = readAll_('Licenses').find(l => l.member_id === p.member_id);
  if (license) {
    updateRow_('Licenses', license._row, { expire_date: newExpire, status: 'ปกติ', renewed_at: nowIso_() });
  }

  appendRow_('MembershipRenewals', {
    renewal_id: newId_('RNW'), member_id: p.member_id, renewed_at: nowIso_(),
    old_expire_date: member.membership_expire_date || '', new_expire_date: newExpire,
    years: years, payment_id: p.payment_id || '', renewed_by: user.user_id, note: p.note || ''
  });

  auditLog_('Members', p.member_id, 'membership_renewed', member, { new_expire_date: newExpire }, user.user_id);
  refreshMemberCard_(p.member_id, newExpire, member.member_type); // ออก QR ใหม่ (TOR ข้อ 5.5)
  mailNotify_(member.email, 'ต่ออายุสมาชิกภาพสำเร็จ',
    'สมาชิกภาพของท่านได้รับการต่ออายุแล้ว\nวันหมดอายุใหม่: ' + new Date(newExpire).toLocaleDateString('th-TH'));

  return { ok: true, new_expire_date: newExpire };
}

function action_listLicenses(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  return readAll_('Licenses');
}

/** แอดมินแก้ไขข้อมูลใบอนุญาตของสมาชิก (เลขที่, วันออก, วันหมดอายุ, สถานะ) — ใช้คู่กับ action_updateMember ตอนแก้ไขโปรไฟล์เต็มรูปแบบ */
function action_updateLicense(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  let license = readAll_('Licenses').find(l => l.member_id === p.member_id);
  const patch = Object.assign({}, p.patch);
  // พักใช้/เพิกถอน ต้องทำผ่านเมนู "พักใช้/เพิกถอนใบอนุญาต" (มีเลขที่มติ กก.จรรยาบรรณ + หลักฐาน) ไม่ให้ตั้งสถานะตรงๆ
  if (patch.status !== undefined && SANCTION_TYPES.indexOf(patch.status) !== -1 && (!license || license.status !== patch.status)) {
    throw new Error('การ' + patch.status + 'ใบอนุญาตต้องบันทึกผ่านเมนู "พักใช้/เพิกถอนใบอนุญาต" พร้อมแนบหลักฐานจากคณะกรรมการจรรยาบรรณ');
  }
  if (patch.status !== undefined && license && patch.status !== license.status && getActiveSanction_(p.member_id)) {
    throw new Error('ใบอนุญาตนี้มีคำสั่งพักใช้/เพิกถอนที่ยังมีผลอยู่ กรุณาสิ้นสุดหรือยกเลิกคำสั่งในเมนู "พักใช้/เพิกถอนใบอนุญาต" ก่อน');
  }

  if (!license) {
    // สมาชิกรายนี้ยังไม่เคยมีใบอนุญาต (เช่น สมาชิกที่นำเข้าข้อมูลมาจากระบบเดิม) — สร้างใหม่ให้
    appendRow_('Licenses', Object.assign({
      license_id: newId_('LIC'), member_id: p.member_id, license_no: '', issue_date: '', expire_date: '',
      status: 'ปกติ', renewal_docs_json: '[]', renewed_at: '', revoked_reason: ''
    }, patch));
    return { ok: true, created: true };
  }

  updateRow_('Licenses', license._row, patch);
  auditLog_('Licenses', license.license_id, 'admin_update', license, patch, user.user_id);

  // ให้ข้อมูลใบอนุญาตในตาราง Members (ใช้แสดงในหน้ารายชื่อ/ทะเบียน) sync ตามไปด้วย
  const member = findRowById_('Members', 'member_id', p.member_id);
  if (member) {
    const memberPatch = {};
    if (patch.license_no !== undefined) memberPatch.license_no = patch.license_no;
    if (patch.status !== undefined) memberPatch.license_status = patch.status;
    if (patch.expire_date !== undefined) memberPatch.membership_expire_date = patch.expire_date;
    if (Object.keys(memberPatch).length) updateRow_('Members', member._row, memberPatch);
  }
  return { ok: true, created: false };
}

// ============================================================
// พักใช้ / เพิกถอนใบอนุญาต (มติ/หนังสือคณะกรรมการจรรยาบรรณ)
// เจ้าหน้าที่ทะเบียน (admin_staff/admin_super) บันทึกคำสั่ง + แนบหลักฐาน
// เจ้าหน้าที่ทุกบทบาท (รวม admin_welfare) ดูข้อมูลและหลักฐานได้ — ดูไม่ได้เฉพาะสมาชิกทั่วไป (สมาชิกเห็นเฉพาะสรุปสถานะของตนเอง)
// ============================================================

const SANCTION_TYPES = ['พักใช้', 'เพิกถอน'];
const SANCTION_READ_ROLES = ['admin_super', 'admin_staff', 'admin_welfare'];

/** แปลงวันที่จาก Sheets (Date) หรือข้อความ ให้เป็น 'YYYY-MM-DD' (เวลาไทย) */
function ymd_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, 'Asia/Bangkok', 'yyyy-MM-dd');
  }
  const m = String(v).match(/^(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : String(v);
}
function todayYmd_() { return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd'); }
function isYmd_(v) { return /^\d{4}-\d{2}-\d{2}$/.test(String(v || '')) && !isNaN(new Date(v).getTime()); }

function parseJsonArr_(v) { try { const a = JSON.parse(v || '[]'); return Array.isArray(a) ? a : []; } catch (e) { return []; } }

/** ตรวจ/จัดรูปหลักฐานแนบ — ต้องเป็นไฟล์ที่อัปโหลดผ่าน action_uploadDocument แล้ว (มี url) */
function normalizeEvidence_(docs, user, fallbackType) {
  return (docs || []).filter(d => d && d.url).map(d => ({
    name: String(d.name || 'เอกสารแนบ'), url: String(d.url), doc_type: String(d.doc_type || fallbackType || 'หลักฐาน'),
    note: String(d.note || ''), uploaded_by: user.username || user.user_id, uploaded_at: nowIso_()
  }));
}

function getActiveSanction_(memberId) {
  if (!memberId) return null;
  return readAll_('LicenseSanctions').find(s => s.member_id === memberId && s.status === 'มีผลบังคับใช้') || null;
}

/** สถานะใบอนุญาตตามวันหมดอายุ (ใช้ตอนสิ้นสุดคำสั่งพักใช้/เพิกถอน) */
function licenseStatusByExpiry_(expireDate) {
  const days = daysUntil_(expireDate);
  if (days === null) return 'ปกติ';
  return days < 0 ? 'หมดอายุ' : (days <= EXPIRY_WARNING_DAYS ? 'ใกล้หมดอายุ' : 'ปกติ');
}

function memberDisplayName_(m) {
  return [m.prefix_th, m.first_name_th, m.last_name_th].filter(Boolean).join(' ').replace(/^(\S+) /, (a, b) => (['นาย', 'นาง', 'นางสาว'].indexOf(b) !== -1 ? b : b + ' '));
}

function sanctionView_(s, withEvidence) {
  const out = Object.assign({}, s);
  delete out._row;
  ['decision_date', 'effective_date', 'suspend_until'].forEach(k => out[k] = ymd_(s[k]));
  const ev = parseJsonArr_(s.evidence_json), lev = parseJsonArr_(s.lift_evidence_json);
  out.evidence_count = ev.length + lev.length;
  if (withEvidence) { out.evidence = ev; out.lift_evidence = lev; }
  delete out.evidence_json; delete out.lift_evidence_json;
  return out;
}

/** บันทึกคำสั่งพักใช้/เพิกถอน — ต้องแนบหลักฐานจากคณะกรรมการจรรยาบรรณอย่างน้อย 1 ไฟล์ */
function action_createLicenseSanction(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  if (SANCTION_TYPES.indexOf(p.action_type) === -1) throw new Error('ประเภทคำสั่งไม่ถูกต้อง (พักใช้ / เพิกถอน)');
  const member = findRowById_('Members', 'member_id', p.member_id);
  if (!member) throw new Error('ไม่พบสมาชิก');
  const license = readAll_('Licenses').find(l => l.member_id === p.member_id);
  if (!license || !license.license_no) throw new Error('สมาชิกรายนี้ยังไม่มีใบอนุญาต');
  if (license.status === 'เพิกถอน') throw new Error('ใบอนุญาตนี้ถูกเพิกถอนอยู่แล้ว');
  if (getActiveSanction_(p.member_id)) throw new Error('สมาชิกรายนี้มีคำสั่งที่ยังมีผลบังคับใช้อยู่แล้ว กรุณาสิ้นสุด/ยกเลิกคำสั่งเดิมก่อน');

  const caseNo = String(p.ethics_case_no || '').trim();
  const reason = String(p.reason || '').trim();
  if (!caseNo) throw new Error('กรุณาระบุเลขที่หนังสือ/มติคณะกรรมการจรรยาบรรณ');
  if (!reason) throw new Error('กรุณาระบุเหตุผล/ฐานความผิดโดยสรุป');
  if (!isYmd_(p.decision_date)) throw new Error('กรุณาระบุวันที่มติ/คำสั่ง');
  if (!isYmd_(p.effective_date)) throw new Error('กรุณาระบุวันที่มีผลบังคับใช้');
  if (p.effective_date > todayYmd_()) throw new Error('วันที่มีผลบังคับใช้ต้องไม่เกินวันนี้ (ระบบมีผลทันทีที่บันทึก) — ถ้ามีผลในอนาคต ให้บันทึกในวันที่มีผล');
  if (p.action_type === 'พักใช้') {
    if (!isYmd_(p.suspend_until)) throw new Error('กรุณาระบุวันสิ้นสุดการพักใช้');
    if (p.suspend_until <= p.effective_date) throw new Error('วันสิ้นสุดการพักใช้ต้องหลังวันที่มีผล');
  }
  const evidence = normalizeEvidence_(p.documents, user, 'มติ/หนังสือคณะกรรมการจรรยาบรรณ');
  if (!evidence.length) throw new Error('กรุณาแนบหลักฐานจากคณะกรรมการจรรยาบรรณอย่างน้อย 1 ไฟล์ (มติ/หนังสือ/รายงานการประชุม)');

  const sanctionId = newId_('SNC');
  const name = memberDisplayName_(member);
  appendRow_('LicenseSanctions', {
    sanction_id: sanctionId, member_id: member.member_id, license_id: license.license_id, license_no: license.license_no,
    member_name: name, action_type: p.action_type, status: 'มีผลบังคับใช้',
    ethics_case_no: caseNo, ethics_meeting_no: String(p.ethics_meeting_no || '').trim(),
    decision_date: p.decision_date, effective_date: p.effective_date,
    suspend_until: p.action_type === 'พักใช้' ? p.suspend_until : '',
    reason: reason, remarks: String(p.remarks || '').trim(),
    evidence_json: JSON.stringify(evidence), prev_license_status: license.status || '',
    lifted_at: '', lifted_by: '', lift_type: '', lift_reason: '', lift_evidence_json: '[]',
    created_by: user.username || user.user_id, created_at: nowIso_(), updated_at: nowIso_(), updated_by: user.username || user.user_id
  });

  updateRow_('Licenses', license._row, { status: p.action_type, revoked_reason: p.action_type === 'เพิกถอน' ? reason : license.revoked_reason });
  updateRow_('Members', member._row, { license_status: p.action_type });
  auditLog_('LicenseSanctions', sanctionId, 'sanction_create', { license_status: license.status }, { action_type: p.action_type, case: caseNo, until: p.suspend_until || '' }, user.user_id);

  mailNotify_(member.email, 'แจ้ง' + p.action_type + 'ใบอนุญาตประกอบวิชาชีพสังคมสงเคราะห์',
    'สภาวิชาชีพสังคมสงเคราะห์ขอแจ้งว่า ใบอนุญาตเลขที่ ' + license.license_no + ' ของท่านถูก' + p.action_type +
    ' ตามมติ/หนังสือคณะกรรมการจรรยาบรรณ เลขที่ ' + caseNo + '\n' +
    'มีผลตั้งแต่วันที่ ' + p.effective_date + (p.action_type === 'พักใช้' ? ' ถึงวันที่ ' + p.suspend_until : '') + '\n' +
    'เหตุผลโดยสรุป: ' + reason + '\nหากมีข้อสงสัยกรุณาติดต่อสำนักงานสภาวิชาชีพสังคมสงเคราะห์', { tone: 'danger' });
  return { ok: true, sanction_id: sanctionId };
}

/** แนบหลักฐานเพิ่มเติมให้คำสั่งเดิม (เจ้าหน้าที่ทะเบียน) */
function action_addLicenseSanctionEvidence(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const s = findRowById_('LicenseSanctions', 'sanction_id', p.sanction_id);
  if (!s) throw new Error('ไม่พบคำสั่ง');
  const add = normalizeEvidence_(p.documents, user, 'หลักฐานเพิ่มเติม');
  if (!add.length) throw new Error('กรุณาแนบไฟล์อย่างน้อย 1 ไฟล์');
  const all = parseJsonArr_(s.evidence_json).concat(add);
  updateRow_('LicenseSanctions', s._row, { evidence_json: JSON.stringify(all), updated_at: nowIso_(), updated_by: user.username || user.user_id });
  auditLog_('LicenseSanctions', s.sanction_id, 'sanction_evidence_add', {}, { added: add.map(d => d.name) }, user.user_id);
  return { ok: true, count: all.length };
}

/** แก้ไขข้อมูลคำสั่ง (ไม่เปลี่ยนประเภทคำสั่ง) */
function action_updateLicenseSanction(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const s = findRowById_('LicenseSanctions', 'sanction_id', p.sanction_id);
  if (!s) throw new Error('ไม่พบคำสั่ง');
  const patch = {};
  ['ethics_case_no', 'ethics_meeting_no', 'reason', 'remarks'].forEach(k => { if (p.patch && p.patch[k] !== undefined) patch[k] = String(p.patch[k] || '').trim(); });
  ['decision_date', 'effective_date', 'suspend_until'].forEach(k => {
    if (p.patch && p.patch[k] !== undefined) {
      if (k === 'suspend_until' && s.action_type !== 'พักใช้') return;
      if (!isYmd_(p.patch[k])) throw new Error('รูปแบบวันที่ไม่ถูกต้อง');
      patch[k] = p.patch[k];
    }
  });
  if (patch.ethics_case_no === '') throw new Error('เลขที่หนังสือ/มติต้องไม่ว่าง');
  if (patch.reason === '') throw new Error('เหตุผลต้องไม่ว่าง');
  const eff = patch.effective_date || ymd_(s.effective_date);
  const until = patch.suspend_until || ymd_(s.suspend_until);
  if (s.action_type === 'พักใช้' && until <= eff) throw new Error('วันสิ้นสุดการพักใช้ต้องหลังวันที่มีผล');
  patch.updated_at = nowIso_(); patch.updated_by = user.username || user.user_id;
  updateRow_('LicenseSanctions', s._row, patch);
  if (patch.reason && s.action_type === 'เพิกถอน') {
    const lic = readAll_('Licenses').find(l => l.member_id === s.member_id);
    if (lic) updateRow_('Licenses', lic._row, { revoked_reason: patch.reason });
  }
  auditLog_('LicenseSanctions', s.sanction_id, 'sanction_update', {}, patch, user.user_id);
  return { ok: true };
}

/** สิ้นสุดคำสั่ง — คืนสถานะใบอนุญาตตามวันหมดอายุ
 *  liftType 'ครบกำหนด' (เฉพาะพักใช้ เมื่อถึงวันสิ้นสุดแล้ว) | 'ยกเลิกคำสั่ง' (ต้องมีเหตุผล + หลักฐาน; ยกเลิกการเพิกถอนเฉพาะ admin_super) */
function applySanctionLift_(s, liftType, reason, evidence, actorLabel, actorUserId) {
  const lic = readAll_('Licenses').find(l => l.member_id === s.member_id);
  const member = findRowById_('Members', 'member_id', s.member_id);
  const restored = lic ? licenseStatusByExpiry_(lic.expire_date) : 'ปกติ';
  if (lic) updateRow_('Licenses', lic._row, { status: restored, revoked_reason: s.action_type === 'เพิกถอน' ? '' : lic.revoked_reason });
  if (member) updateRow_('Members', member._row, { license_status: restored });
  updateRow_('LicenseSanctions', s._row, {
    status: liftType === 'ครบกำหนด' ? 'สิ้นสุดแล้ว' : 'ยกเลิกคำสั่ง',
    lifted_at: nowIso_(), lifted_by: actorLabel, lift_type: liftType, lift_reason: reason || '',
    lift_evidence_json: JSON.stringify(evidence || []), updated_at: nowIso_(), updated_by: actorLabel
  });
  auditLog_('LicenseSanctions', s.sanction_id, 'sanction_lift', { status: s.status }, { lift_type: liftType, restored: restored }, actorUserId);
  if (member) {
    mailNotify_(member.email, 'แจ้งการสิ้นสุด' + s.action_type + 'ใบอนุญาต',
      'ใบอนุญาตเลขที่ ' + s.license_no + ' ของท่านสิ้นสุดการ' + s.action_type + 'แล้ว (' + liftType + ') สถานะปัจจุบัน: ' + restored +
      (restored === 'หมดอายุ' ? '\nใบอนุญาตหมดอายุแล้ว กรุณายื่นต่ออายุผ่าน Member Portal' : ''), { tone: 'success' });
  }
  return restored;
}

function action_liftLicenseSanction(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const s = findRowById_('LicenseSanctions', 'sanction_id', p.sanction_id);
  if (!s) throw new Error('ไม่พบคำสั่ง');
  if (s.status !== 'มีผลบังคับใช้') throw new Error('คำสั่งนี้สิ้นสุดแล้ว');
  if (['ครบกำหนด', 'ยกเลิกคำสั่ง'].indexOf(p.lift_type) === -1) throw new Error('ประเภทการสิ้นสุดไม่ถูกต้อง');
  if (p.lift_type === 'ครบกำหนด') {
    if (s.action_type !== 'พักใช้') throw new Error('การเพิกถอนไม่มี "ครบกำหนด" — ต้องยกเลิกคำสั่งพร้อมหลักฐาน');
    if (ymd_(s.suspend_until) > todayYmd_()) throw new Error('ยังไม่ถึงวันสิ้นสุดการพักใช้ (' + ymd_(s.suspend_until) + ') — หากต้องการสิ้นสุดก่อนกำหนดให้เลือก "ยกเลิกคำสั่ง" พร้อมหลักฐาน');
  } else {
    if (s.action_type === 'เพิกถอน') requireRole_(user, ['admin_super']); // ยกเลิกการเพิกถอน: ผู้ดูแลระบบเท่านั้น
    if (!String(p.lift_reason || '').trim()) throw new Error('กรุณาระบุเหตุผลการยกเลิกคำสั่ง');
  }
  const evidence = normalizeEvidence_(p.documents, user, 'หลักฐานการสิ้นสุดคำสั่ง');
  if (p.lift_type === 'ยกเลิกคำสั่ง' && !evidence.length) throw new Error('กรุณาแนบหนังสือ/มติคณะกรรมการจรรยาบรรณที่ให้ยกเลิกคำสั่ง');
  const restored = applySanctionLift_(s, p.lift_type, String(p.lift_reason || '').trim(), evidence, user.username || user.user_id, user.user_id);
  return { ok: true, restored_status: restored };
}

/** ตรวจรายวัน: พักใช้ที่ถึงวันสิ้นสุดแล้ว → คืนสถานะอัตโนมัติ */
function endExpiredSuspensions_() {
  const today = todayYmd_();
  readAll_('LicenseSanctions').forEach(s => {
    if (s.status === 'มีผลบังคับใช้' && s.action_type === 'พักใช้' && s.suspend_until && ymd_(s.suspend_until) <= today) {
      applySanctionLift_(s, 'ครบกำหนด', 'ครบกำหนดพักใช้ (ระบบสิ้นสุดอัตโนมัติ)', [], 'system', 'system');
    }
  });
}

/** รายการคำสั่ง — เจ้าหน้าที่ทุกบทบาทดูได้ (p.status, p.action_type, p.member_id, p.q) */
function action_listLicenseSanctions(p, user) {
  requireRole_(user, SANCTION_READ_ROLES);
  let rows = readAll_('LicenseSanctions');
  if (p.status) rows = rows.filter(r => r.status === p.status);
  if (p.action_type) rows = rows.filter(r => r.action_type === p.action_type);
  if (p.member_id) rows = rows.filter(r => r.member_id === p.member_id);
  if (p.q) {
    const q = String(p.q).toLowerCase();
    rows = rows.filter(r => [r.member_name, r.license_no, r.ethics_case_no].some(v => String(v || '').toLowerCase().indexOf(q) !== -1));
  }
  rows.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
  return rows.map(r => sanctionView_(r, !!p.with_evidence));
}

/** รายละเอียดคำสั่ง พร้อมหลักฐานทั้งหมด — เจ้าหน้าที่ทุกบทบาทดูได้ */
function action_getLicenseSanction(p, user) {
  requireRole_(user, SANCTION_READ_ROLES);
  const s = findRowById_('LicenseSanctions', 'sanction_id', p.sanction_id);
  if (!s) throw new Error('ไม่พบคำสั่ง');
  const history = readAll_('LicenseSanctions').filter(r => r.member_id === s.member_id && r.sanction_id !== s.sanction_id)
    .sort((a, b) => new Date(b.created_at) - new Date(a.created_at)).map(r => sanctionView_(r, false));
  return { sanction: sanctionView_(s, true), history: history, can_edit: ['admin_super', 'admin_staff'].indexOf(user.role) !== -1 };
}

/** ผูกบัญชีองค์กรที่สร้างแยกไว้ เข้ากับสมาชิกวิสามัญ (กรณีบัญชีที่ใช้สมัครผูกกับสมาชิกบุคคลอยู่แล้ว ตอนอนุมัติจึงไม่ได้ผูกให้) */
function action_bindOrgAccount(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const org = findRowById_('MembersOrg', 'org_id', p.org_id);
  if (!org) throw new Error('ไม่พบสมาชิกวิสามัญ');
  const acct = findRowById_('Users', 'username', String(p.username || '').trim());
  if (!acct || acct.role !== 'member') throw new Error('ไม่พบบัญชีผู้ใช้งานสมาชิกนี้');
  if (acct.member_id) throw new Error('บัญชีนี้ผูกกับสมาชิกอยู่แล้ว กรุณาให้ผู้แทนสร้างบัญชีองค์กรใหม่ (ชื่อผู้ใช้งานแยก)');
  const already = readAll_('Users').find(u => u.member_id === org.org_id);
  if (already) throw new Error('สมาชิกวิสามัญรายนี้มีบัญชีผูกอยู่แล้ว (' + already.username + ')');
  updateRow_('Users', acct._row, { member_id: org.org_id, account_type: 'org', display_name: acct.display_name || org.org_name_th || '' });
  auditLog_('Users', acct.user_id, 'bind_org_account', {}, { org_id: org.org_id }, user.user_id);
  return { ok: true };
}

/** ระบบแจ้งเตือน (TOR ข้อ 3.3) — ดูประวัติการแจ้งเตือนที่ส่งไปแล้ว เพื่อติดตาม */
function action_listNotifications(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  return readAll_('Notifications').sort((a, b) => new Date(b.sent_at) - new Date(a.sent_at));
}

/** สั่งรันตรวจสอบ+แจ้งเตือนด้วยตนเองจากหน้าแอดมิน (นอกเหนือจาก trigger รายวัน) */
function action_runExpiryCheckNow(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  recomputeExpiryStatuses();
  const sent = sendExpiryNotifications();
  return { ok: true, notifications_sent: sent };
}

/** ระบบค้นหาสมาชิกสำหรับประชาชนทั่วไป (TOR ข้อ 3.4) — PUBLIC ไม่ต้อง login
 * คืนเฉพาะฟิลด์ที่เปิดเผยได้: ชื่อ-นามสกุล เลขใบอนุญาต สถานะ วันหมดอายุ
 */
/** ค้นหาสมาชิก/ผู้ได้รับใบอนุญาตแบบสาธารณะ (ไม่ต้อง login) — รองรับ 2 โหมด:
 * 1) ค้นหาด้วยชื่อ-นามสกุล: p.first_name และ/หรือ p.last_name (แยกช่องตามรูปแบบระบบตรวจสอบแพทยสภา)
 * 2) ค้นหาด้วยเลขที่ใบอนุญาต: p.license_no (ค้นแบบตรงเป๊ะ)
 * ยังรองรับ p.q แบบเดิม (ค้นแบบคำเดียวปนกันทุกฟิลด์) ไว้เพื่อความเข้ากันได้ย้อนหลัง
 */
/** จับคู่ชื่อ-นามสกุลแบบยืดหยุ่น — รองรับทั้งกรณีกรอกแยกช่องถูกต้อง และกรณีพิมพ์ชื่อเต็ม (ชื่อ+นามสกุล) ลงในช่องเดียว
 * (ผู้ใช้บางคนอาจไม่ทันสังเกตว่ามี 2 ช่องแยกกัน แล้วพิมพ์ "สมชาย ใจดี" ลงในช่อง "ชื่อ" ช่องเดียว)
 */
function memberMatchesNameQuery_(m, fnQuery, lnQuery) {
  const firstTh = String(m.first_name_th || '').toLowerCase();
  const lastTh = String(m.last_name_th || '').toLowerCase();
  const firstEn = String(m.first_name_en || '').toLowerCase();
  const lastEn = String(m.last_name_en || '').toLowerCase();
  const fullTh = (firstTh + ' ' + lastTh).trim();
  const fullEn = (firstEn + ' ' + lastEn).trim();

  if (fnQuery && lnQuery) {
    const matchFn = firstTh.includes(fnQuery) || firstEn.includes(fnQuery);
    const matchLn = lastTh.includes(lnQuery) || lastEn.includes(lnQuery);
    return matchFn && matchLn;
  }
  // กรอกมาช่องเดียว — ลองจับคู่ทั้งแบบชื่อ/นามสกุลเดี่ยวๆ และแบบชื่อเต็มรวมกัน (เผื่อพิมพ์ทั้งชื่อ-นามสกุลในช่องเดียว)
  const q = fnQuery || lnQuery;
  return firstTh.includes(q) || lastTh.includes(q) || firstEn.includes(q) || lastEn.includes(q) ||
    fullTh.includes(q) || fullEn.includes(q);
}

function action_publicSearchMembers(p) {
  const members = readAll_('Members');
  const licenses = readAll_('Licenses');
  let results = [];

  if (p.license_no) {
    const q = String(p.license_no).trim().toLowerCase();
    if (!q) throw new Error('กรุณากรอกเลขที่ใบอนุญาต');
    results = members.filter(m => String(m.license_no || '').toLowerCase() === q);
  } else if (p.first_name || p.last_name) {
    const fn = String(p.first_name || '').trim().toLowerCase();
    const ln = String(p.last_name || '').trim().toLowerCase();
    if (!fn && !ln) throw new Error('กรุณากรอกชื่อหรือนามสกุลอย่างน้อย 1 ช่อง');
    results = members.filter(m => memberMatchesNameQuery_(m, fn, ln));
  } else {
    const q = String(p.q || '').trim().toLowerCase();
    if (!q || q.length < 2) throw new Error('กรุณากรอกคำค้นหาอย่างน้อย 2 ตัวอักษร');
    results = members.filter(m =>
      memberMatchesNameQuery_(m, q, '') ||
      String(m.license_no || '').toLowerCase().includes(q) ||
      String(m.member_no || '').toLowerCase().includes(q)
    );
  }

  return results.slice(0, 20).map(m => {
    const lic = licenses.find(l => l.member_id === m.member_id);
    const licenseStatus = lic ? lic.status : m.license_status;
    const displayName = formatMemberDisplayName_(m, licenseStatus);
    // แสดงรูปถ่ายต่อสาธารณะเฉพาะสมาชิกที่ให้ความยินยอมไว้เท่านั้น (photo_public_consent) — ป้องกันการเปิดเผยรูปถ่ายโดยไม่ได้รับอนุญาต
    const showPhoto = isTrue_(m.photo_public_consent);
    // หน้านี้คือ "ตรวจสอบผู้ประกอบวิชาชีพรับอนุญาต" — ต้องระบุชัดว่าใครไม่มีใบอนุญาตที่ใช้งานได้ในปัจจุบัน (ปกติ/ใกล้หมดอายุ ถือว่ายังใช้งานได้)
    const hasActiveLicense = licenseStatus === 'ปกติ' || licenseStatus === 'ใกล้หมดอายุ';
    return {
      name: displayName.th, name_en: displayName.en, usedProfessionalTitle: displayName.usedProfessionalTitle,
      member_type: m.member_type,
      license_no: m.license_no || '-',
      license_issue_date: lic ? lic.issue_date : '',
      status: licenseStatus,
      hasActiveLicense: hasActiveLicense,
      expire_date: lic ? lic.expire_date : (m.membership_expire_date || ''),
      expertise: m.expertise || '',
      photo_url: showPhoto ? (m.photo_url || '') : ''
    };
  });
}

/** อัปเดตสถานะ "ปกติ / ใกล้หมดอายุ / หมดอายุ" อัตโนมัติจากวันหมดอายุ (TOR ข้อ 3.2)
 * ไม่แตะสถานะที่เจ้าหน้าที่ตั้งเองด้วยมือ (พักใช้/เพิกถอน/อยู่ระหว่างต่ออายุ)
 * เรียกได้ทั้งจากปุ่มในหน้าแอดมิน (action_runExpiryCheckNow) และจาก time-driven trigger รายวัน
 */
function recomputeExpiryStatuses() {
  endExpiredSuspensions_(); // พักใช้ที่ครบกำหนด → คืนสถานะก่อนคำนวณวันหมดอายุ
  const manualStatuses = ['พักใช้', 'เพิกถอน', 'อยู่ระหว่างต่ออายุ', 'ถูกพักใช้', 'ถูกเพิกถอน'];
  const licenses = readAll_('Licenses');
  licenses.forEach(lic => {
    if (manualStatuses.indexOf(lic.status) !== -1) return;
    const days = daysUntil_(lic.expire_date);
    if (days === null) return;
    const newStatus = days < 0 ? 'หมดอายุ' : (days <= EXPIRY_WARNING_DAYS ? 'ใกล้หมดอายุ' : 'ปกติ');
    if (newStatus !== lic.status) updateRow_('Licenses', lic._row, { status: newStatus });
  });

  const members = readAll_('Members');
  members.forEach(m => {
    if (manualStatuses.indexOf(m.license_status) !== -1) return;
    const days = daysUntil_(m.membership_expire_date);
    if (days === null) return;
    const newStatus = days < 0 ? 'หมดอายุ' : (days <= EXPIRY_WARNING_DAYS ? 'ใกล้หมดอายุ' : 'ปกติ');
    if (newStatus !== m.license_status) updateRow_('Members', m._row, { license_status: newStatus });
  });
}

/** ส่งอีเมลแจ้งเตือนสมาชิก/ใบอนุญาตใกล้หมดอายุและหมดอายุ (TOR ข้อ 3.3) กันส่งซ้ำด้วยการเช็ค log ใน Notifications */
function sendExpiryNotifications() {
  const members = readAll_('Members');
  const existingNotifs = readAll_('Notifications');
  let sentCount = 0;

  members.forEach(m => {
    const days = daysUntil_(m.membership_expire_date);
    if (days === null) return;
    let type = null;
    if (days < 0) type = 'expired';
    else if (days <= EXPIRY_WARNING_DAYS) type = 'near_expiry';
    if (!type || !m.email) return;

    const alreadySent = existingNotifs.some(n =>
      n.member_id === m.member_id && n.type === type &&
      daysUntil_(n.sent_at) !== null && Math.abs(daysUntil_(n.sent_at)) < 25
    );
    if (alreadySent) return;

    const subject = type === 'expired' ? 'สมาชิกภาพของท่านหมดอายุแล้ว' : 'แจ้งเตือน: สมาชิกภาพของท่านใกล้หมดอายุ';
    const body = type === 'expired'
      ? 'สมาชิกภาพของท่านหมดอายุเมื่อวันที่ ' + new Date(m.membership_expire_date).toLocaleDateString('th-TH') + ' กรุณาติดต่อสภาฯ เพื่อต่ออายุ'
      : 'สมาชิกภาพของท่านจะหมดอายุในวันที่ ' + new Date(m.membership_expire_date).toLocaleDateString('th-TH') + ' (อีก ' + days + ' วัน) กรุณาติดต่อสภาฯ เพื่อต่ออายุล่วงหน้า';

    mailNotify_(m.email, subject, body);
    appendRow_('Notifications', {
      notification_id: newId_('NTF'), ref_type: 'membership', ref_id: m.member_id, member_id: m.member_id,
      channel: 'email', type: type, email_to: m.email, subject: subject, sent_at: nowIso_(), status: 'sent'
    });
    sentCount++;
  });

  return sentCount;
}

/** รันครั้งเดียวเพื่อตั้งเวลาให้ระบบตรวจสอบ+แจ้งเตือนวันหมดอายุอัตโนมัติทุกวัน (TOR ข้อ 3.3)
 * วิธีรัน: เลือกฟังก์ชันนี้จาก dropdown ใน Apps Script editor > กด Run ครั้งเดียว
 */
function createDailyExpiryTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyExpiryJob_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyExpiryJob_').timeBased().everyDays(1).atHour(7).create();
  return 'ตั้งเวลาแจ้งเตือนรายวันสำเร็จ (รันทุกวันประมาณ 07:00 น.)';
}

function dailyExpiryJob_() {
  recomputeExpiryStatuses();
  sendExpiryNotifications();
}

// ---------- 4. ระบบใบอนุญาต (TOR ข้อ 4) ----------
// สถาปัตยกรรม: ใช้ Applications + Payments ร่วมกับระบบสมัครสมาชิก/เปลี่ยนประเภทเหมือนเดิม
// (applicant_type = 'ขึ้นทะเบียน' หรือ 'ต่ออายุใบอนุญาต') — reviewApplication และ submitPaymentSlip ใช้ร่วมกันได้เลย
// เฉพาะขั้นตอนสุดท้าย (ยืนยันชำระเงิน) ที่แยกเป็นฟังก์ชันเฉพาะ เพราะต้องออก/ต่อเลขใบอนุญาตจริง

/** ฟังก์ชันกลาง: ทำให้ใบอนุญาตของสมาชิกคนหนึ่ง "ต่ออายุสำเร็จ" ไม่ว่าจะมาจาก flow ไหน (self-service / group approval)
 * ถ้าเป็นการต่อ "หลังหมดอายุ" (isAfterExpiry) จะ Void เลขเดิมแล้วออกเลขใหม่ตาม TOR ข้อ 4.4.4
 */
function finalizeLicenseRenewal_(memberId, isAfterExpiry, method, actorUserId) {
  const member = findRowById_('Members', 'member_id', memberId);
  if (!member) throw new Error('ไม่พบสมาชิก ' + memberId);
  const license = readAll_('Licenses').find(l => l.member_id === memberId);
  if (!license) throw new Error('สมาชิกรายนี้ยังไม่เคยขึ้นทะเบียนรับใบอนุญาต ต้องขึ้นทะเบียนก่อนจึงจะต่ออายุได้');

  const oldExpire = license.expire_date;
  let newLicenseNo = license.license_no;
  const newExpire = addYearsIso_(isAfterExpiry ? nowIso_() : (oldExpire || nowIso_()), LICENSE_DURATION_YEARS);

  if (isAfterExpiry) {
    // Void เลขเดิม ออกเลขใหม่ (TOR ข้อ 4.4.4) — ไม่นำเลขที่ถูกยกเลิกกลับมาใช้ซ้ำ
    newLicenseNo = generateLicenseNumber_();
    appendRow_('VoidedLicenses', {
      void_id: newId_('VOID'), old_license_no: license.license_no, new_license_no: newLicenseNo,
      member_id: memberId, member_name: [member.prefix_th, member.first_name_th, member.last_name_th].filter(Boolean).join(' '),
      void_date: nowIso_(), reason: 'ต่ออายุหลังใบอนุญาตหมดอายุ (ถือเป็นการขึ้นทะเบียนใหม่)', actor_user_id: actorUserId
    });
    updateRow_('Licenses', license._row, {
      license_no: newLicenseNo, issue_date: nowIso_(), expire_date: newExpire, status: 'ปกติ', renewed_at: nowIso_()
    });
  } else {
    updateRow_('Licenses', license._row, { expire_date: newExpire, status: 'ปกติ', renewed_at: nowIso_() });
  }

  updateRow_('Members', member._row, { license_no: newLicenseNo, license_status: 'ปกติ', updated_at: nowIso_() });

  appendRow_('LicenseRenewalHistory', {
    renewal_id: newId_('LR'), license_id: license.license_id, member_id: memberId, method: method,
    old_expire_date: oldExpire, new_expire_date: newExpire, renewed_by: actorUserId, renewed_at: nowIso_(),
    note: isAfterExpiry ? 'ต่อหลังหมดอายุ — ออกเลขใหม่ ' + newLicenseNo : 'ต่อก่อนหมดอายุ — คงเลขเดิม'
  });

  mailNotify_(member.email, 'ต่ออายุใบอนุญาตสำเร็จ',
    'ใบอนุญาตของท่านได้รับการต่ออายุแล้ว\nเลขที่ใบอนุญาต: ' + newLicenseNo +
    '\nวันหมดอายุใหม่: ' + new Date(newExpire).toLocaleDateString('th-TH') +
    (isAfterExpiry ? '\n(หมายเหตุ: เนื่องจากดำเนินการหลังใบอนุญาตเดิมหมดอายุ ระบบออกเลขที่ใบอนุญาตใหม่ให้)' : ''));

  return { license_no: newLicenseNo, new_expire_date: newExpire };
}

// ---------- 4.1 ขึ้นทะเบียนรับใบอนุญาต ----------

function action_getRegistrationEligibility(p, user) {
  requireRole_(user, ['member']);
  if (!user.member_id) return { eligible: false, reason: 'บัญชีนี้ยังไม่ผูกกับข้อมูลสมาชิก' };
  const member = findRowById_('Members', 'member_id', user.member_id);
  if (!member) return { eligible: false, reason: 'ไม่พบข้อมูลสมาชิก' };
  if (member.member_type !== 'สามัญ') return { eligible: false, reason: 'ต้องเป็นสมาชิกสามัญเท่านั้นจึงจะขึ้นทะเบียนรับใบอนุญาตได้ (TOR ข้อ 4.1.1)' };
  if (member.member_status !== 'ปกติ') return { eligible: false, reason: 'สถานะสมาชิกสามัญต้องเป็น "ปกติ" ก่อนจึงจะยื่นคำขอขึ้นทะเบียนได้ กรุณาต่ออายุสมาชิกภาพก่อน' };
  const existingLicense = readAll_('Licenses').find(l => l.member_id === user.member_id);
  if (existingLicense) return { eligible: false, reason: 'สมาชิกรายนี้มีใบอนุญาตอยู่แล้ว (เลขที่ ' + existingLicense.license_no + ') กรุณาใช้เมนูต่ออายุใบอนุญาตแทน' };
  const pending = readAll_('Applications').find(a => a.applicant_type === 'ขึ้นทะเบียน' && a.applicant_user_id === user.user_id && ['รอตรวจสอบ', 'รอชำระเงิน'].indexOf(a.status) !== -1);
  if (pending) return { eligible: false, reason: 'มีคำขอขึ้นทะเบียนที่ยังไม่เสร็จสิ้นอยู่แล้ว (เลขที่ ' + pending.application_id + ')' };
  return { eligible: true, examSubjects: EXAM_SUBJECTS, interviewDocLabel: EXAM_INTERVIEW_LABEL };
}

/** ต้องแนบผลสอบข้อเขียนครบ 2 วิชา + ผลสอบสัมภาษณ์ ก่อนยื่นขึ้นทะเบียนรับใบอนุญาตครั้งแรกได้ (TOR ข้อ 4.1.1) — ตรวจซ้ำฝั่ง server กันข้ามการเช็ค UI */
function action_submitLicenseRegistration(p, user) {
  requireRole_(user, ['member']);
  const check = action_getRegistrationEligibility(p, user);
  if (!check.eligible) throw new Error(check.reason);
  const member = findRowById_('Members', 'member_id', user.member_id);

  const docs = p.documents || [];
  const hasSubject1 = docs.some(d => d.doc_type === 'ผลสอบข้อเขียน ' + EXAM_SUBJECTS[0]);
  const hasSubject2 = docs.some(d => d.doc_type === 'ผลสอบข้อเขียน ' + EXAM_SUBJECTS[1]);
  const hasInterview = docs.some(d => d.doc_type === EXAM_INTERVIEW_LABEL);
  if (!hasSubject1 || !hasSubject2 || !hasInterview) {
    throw new Error('ต้องแนบผลสอบข้อเขียนครบทั้ง 2 วิชา (' + EXAM_SUBJECTS[0] + ', ' + EXAM_SUBJECTS[1] + ') และผลการสอบสัมภาษณ์ ก่อนยื่นคำขอขึ้นทะเบียนได้');
  }

  const app = {
    application_id: newId_('REG'), applicant_type: 'ขึ้นทะเบียน', applicant_user_id: user.user_id,
    ref_data_json: JSON.stringify({ member_id: user.member_id, email: member.email, name: [member.prefix_th, member.first_name_th, member.last_name_th].filter(Boolean).join(' ') }),
    documents_json: JSON.stringify(docs), status: 'รอตรวจสอบ', reject_reason: '',
    fee_amount: LICENSE_REGISTRATION_FEE, submitted_at: nowIso_(), reviewed_by: '', reviewed_at: '', approved_member_id: ''
  };
  appendRow_('Applications', app);
  mailNotify_(member.email, 'ได้รับคำขอขึ้นทะเบียนรับใบอนุญาตแล้ว',
    'ระบบได้รับคำขอขึ้นทะเบียนของท่านแล้ว เลขที่คำขอ ' + app.application_id + '\nสถานะปัจจุบัน: รอตรวจสอบ');
  return { application_id: app.application_id };
}

/** แอดมินยืนยันชำระเงิน -> ออกเลขใบอนุญาตใหม่ครั้งแรก (TOR ข้อ 4.1.2) */
function action_verifyRegistrationPayment(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const payment = findRowById_('Payments', 'payment_id', p.payment_id);
  if (!payment || payment.ref_type !== 'application') throw new Error('ไม่พบรายการชำระเงิน');
  const app = findRowById_('Applications', 'application_id', payment.ref_id);
  if (!app || app.applicant_type !== 'ขึ้นทะเบียน') throw new Error('ไม่พบคำขอขึ้นทะเบียนที่เกี่ยวข้อง');

  if (p.decision === 'reject') {
    updateRow_('Payments', payment._row, { status: 'ปฏิเสธ', verified_by: user.user_id, verified_at: nowIso_() });
    return { ok: true };
  }

  const refData = JSON.parse(app.ref_data_json || '{}');
  const member = findRowById_('Members', 'member_id', refData.member_id);
  if (!member) throw new Error('ไม่พบสมาชิกที่เกี่ยวข้อง');

  const licenseNo = generateLicenseNumber_();
  const issueDate = nowIso_();
  const expireDate = addYearsIso_(issueDate, LICENSE_DURATION_YEARS);
  appendRow_('Licenses', {
    license_id: newId_('LIC'), member_id: member.member_id, license_no: licenseNo,
    issue_date: issueDate, expire_date: expireDate, status: 'ปกติ',
    renewal_docs_json: '[]', renewed_at: '', revoked_reason: ''
  });
  updateRow_('Members', member._row, { license_no: licenseNo, license_status: 'ปกติ', updated_at: nowIso_() });

  updateRow_('Payments', payment._row, { status: 'ชำระแล้ว', verified_by: user.user_id, verified_at: nowIso_(), paid_at: nowIso_() });
  try { issueReceiptForPayment_(payment.payment_id, user.user_id); } catch (e) { /* ไม่ให้ล้มเหลวทั้งฟังก์ชันถ้าออกใบเสร็จไม่สำเร็จ */ }
  updateRow_('Applications', app._row, { status: 'ขึ้นทะเบียนสำเร็จ', approved_member_id: member.member_id });
  auditLog_('Licenses', licenseNo, 'registration_confirmed', {}, { license_no: licenseNo }, user.user_id);

  mailNotify_(member.email, 'ขึ้นทะเบียนรับใบอนุญาตสำเร็จ',
    'ท่านได้รับใบอนุญาตประกอบวิชาชีพเรียบร้อยแล้ว\nเลขที่ใบอนุญาต: ' + licenseNo +
    '\nวันที่ออก: ' + new Date(issueDate).toLocaleDateString('th-TH') +
    '\nวันหมดอายุ: ' + new Date(expireDate).toLocaleDateString('th-TH'));

  return { ok: true, license_no: licenseNo };
}

// ---------- 4.2 ต่ออายุใบอนุญาต (Self-service) ----------

function action_getRenewalEligibility(p, user) {
  requireRole_(user, ['member']);
  if (!user.member_id) return { eligible: false, reason: 'บัญชีนี้ยังไม่ผูกกับข้อมูลสมาชิก' };
  const member = findRowById_('Members', 'member_id', user.member_id);
  if (!member) return { eligible: false, reason: 'ไม่พบข้อมูลสมาชิก' };
  const license = readAll_('Licenses').find(l => l.member_id === user.member_id);
  const activeSanction = getActiveSanction_(user.member_id);
  if (activeSanction || license && SANCTION_TYPES.indexOf(license.status) !== -1) {
    const t = activeSanction ? activeSanction.action_type : license.status;
    return { eligible: false, sanction: true, reason: 'ใบอนุญาตของท่านอยู่ในสถานะ' + t +
      (activeSanction && activeSanction.action_type === 'พักใช้' ? ' (ถึงวันที่ ' + ymd_(activeSanction.suspend_until) + ')' : '') +
      ' จึงยังไม่สามารถยื่นต่ออายุได้ กรุณาติดต่อสำนักงานสภาวิชาชีพสังคมสงเคราะห์' };
  }
  if (!license) return { eligible: false, reason: 'ยังไม่เคยขึ้นทะเบียนรับใบอนุญาต กรุณาขึ้นทะเบียนก่อน' };
  if (member.member_status !== 'ปกติ') {
    return { eligible: false, reason: 'สถานะสมาชิกสามัญของท่านไม่ปกติ (' + member.member_status + ') กรุณาต่ออายุสมาชิกภาพให้เป็นปกติก่อน จึงจะต่ออายุใบอนุญาตได้ (TOR ข้อ 4.2.1)', license: license };
  }
  const windowOpen = (PropertiesService.getScriptProperties().getProperty('RENEWAL_WINDOW_OPEN') || 'true') === 'true';
  if (!windowOpen) return { eligible: false, reason: 'เจ้าหน้าที่ยังไม่เปิดช่วงเวลาให้ต่ออายุใบอนุญาตในขณะนี้', license: license };
  const days = daysUntil_(license.expire_date);
  const isAfterExpiry = days !== null && days < 0;
  if (days !== null && days > RENEWAL_WINDOW_DAYS) {
    return { eligible: false, reason: 'ยังไม่ถึงช่วงต่ออายุ — กดต่ออายุล่วงหน้าได้ไม่เกิน ' + RENEWAL_WINDOW_DAYS + ' วันก่อนวันหมดอายุ (อีก ' + (days - RENEWAL_WINDOW_DAYS) + ' วันจะเปิดให้ต่ออายุ)', license: license };
  }
  const pending = readAll_('Applications').find(a => a.applicant_type === 'ต่ออายุใบอนุญาต' && a.applicant_user_id === user.user_id && ['รอตรวจสอบ', 'รอชำระเงิน'].indexOf(a.status) !== -1);
  if (pending) return { eligible: false, reason: 'มีคำขอต่ออายุที่ยังไม่เสร็จสิ้นอยู่แล้ว (เลขที่ ' + pending.application_id + ')', license: license };

  const cat1 = Number(member.ce_score_cat1) || 0;
  const cat2 = Number(member.ce_score_cat2) || 0;
  const ceQualified = cat1 >= CE_MINIMUM_CATEGORY1 && cat2 >= CE_MINIMUM_CATEGORY2;

  return {
    eligible: true, license: license, isAfterExpiry: isAfterExpiry, daysUntilExpire: days,
    ceQualified: ceQualified, ceCat1: cat1, ceCat2: cat2,
    ceMinCat1: CE_MINIMUM_CATEGORY1, ceMinCat2: CE_MINIMUM_CATEGORY2,
    examSubjects: EXAM_SUBJECTS, examSubjectFee: EXAM_SUBJECT_FEE, interviewDocLabel: EXAM_INTERVIEW_LABEL,
    licenseRegistrationFee: LICENSE_REGISTRATION_FEE
  };
}

/** p.method: 'CE_before' | 'CE_after' | 'exam' — วิธี CE (ทั้งก่อน/หลังหมดอายุ) ต้องผ่านเกณฑ์ CE ทั้ง 2 หมวดจริงจึงจะยื่นได้ (ตรวจสอบซ้ำฝั่ง server กันข้ามการเช็คฝั่ง UI) */
function action_submitLicenseRenewal(p, user) {
  requireRole_(user, ['member']);
  const check = action_getRenewalEligibility(p, user);
  if (!check.eligible) throw new Error(check.reason);
  const member = findRowById_('Members', 'member_id', user.member_id);

  if (['CE_before', 'CE_after', 'exam'].indexOf(p.method) === -1) throw new Error('วิธีต่ออายุไม่ถูกต้อง');
  if (p.method === 'CE_before' && check.isAfterExpiry) throw new Error('ใบอนุญาตหมดอายุแล้ว กรุณาเลือกวิธี "ต่อหลังหมดอายุด้วยผล CE" หรือ "สอบข้อเขียนใหม่" แทน');
  if ((p.method === 'CE_before' || p.method === 'CE_after') && !check.ceQualified) {
    throw new Error('ยังไม่ผ่านเกณฑ์การศึกษาต่อเนื่อง (CE) — หมวด 1: ' + check.ceCat1 + '/' + check.ceMinCat1 +
      ' หมวด 2: ' + check.ceCat2 + '/' + check.ceMinCat2 + ' กรุณาเลือกต่ออายุด้วยการสอบข้อเขียนใหม่แทน');
  }

  const docs = p.documents || [];
  if (p.method === 'CE_after') {
    const hasInterview = docs.some(d => d.doc_type === EXAM_INTERVIEW_LABEL);
    if (!hasInterview) throw new Error('กรณีต่ออายุหลังใบอนุญาตหมดอายุด้วยผล CE ต้องแนบผลการสอบสัมภาษณ์เพิ่มเติมด้วย');
  }
  if (p.method === 'exam') {
    const hasSubject1 = docs.some(d => d.doc_type === 'ผลสอบข้อเขียน ' + EXAM_SUBJECTS[0]);
    const hasSubject2 = docs.some(d => d.doc_type === 'ผลสอบข้อเขียน ' + EXAM_SUBJECTS[1]);
    const hasInterview = docs.some(d => d.doc_type === EXAM_INTERVIEW_LABEL);
    if (!hasSubject1 || !hasSubject2 || !hasInterview) {
      throw new Error('กรณีต่ออายุด้วยการสอบข้อเขียนใหม่ ต้องแนบผลสอบข้อเขียนครบ 2 วิชา และผลสอบสัมภาษณ์');
    }
  }

  const refData = {
    member_id: user.member_id, method: p.method, isAfterExpiry: check.isAfterExpiry,
    email: member.email, name: [member.prefix_th, member.first_name_th, member.last_name_th].filter(Boolean).join(' ')
  };

  // กรณีต่อก่อนหมดอายุด้วยผล CE — TOR ให้กดต่อได้เลยและชำระเงินได้ทันที ไม่ต้องรอเจ้าหน้าที่อนุมัติก่อน (ตรวจสอบผ่านคะแนน CE จริงแล้วด้านบน)
  if (p.method === 'CE_before') {
    const app = {
      application_id: newId_('REN'), applicant_type: 'ต่ออายุใบอนุญาต', applicant_user_id: user.user_id,
      ref_data_json: JSON.stringify(refData), documents_json: JSON.stringify(p.documents || []),
      status: 'รอชำระเงิน', reject_reason: '', fee_amount: LICENSE_RENEWAL_FEE,
      submitted_at: nowIso_(), reviewed_by: 'ระบบ (ผ่านเกณฑ์ CE อัตโนมัติ)', reviewed_at: nowIso_(), approved_member_id: ''
    };
    appendRow_('Applications', app);
    appendRow_('Payments', {
      payment_id: newId_('PAY'), ref_type: 'application', ref_id: app.application_id, member_id: user.member_id,
      amount: LICENSE_RENEWAL_FEE, slip_url: '', method: '', status: 'รอชำระ', verified_by: '', verified_at: '', paid_at: ''
    });
    mailNotify_(member.email, 'ยื่นต่ออายุใบอนุญาตสำเร็จ — กรุณาชำระเงิน',
      'ระบบตรวจสอบคุณสมบัติผ่านแล้ว กรุณาชำระค่าธรรมเนียม ' + LICENSE_RENEWAL_FEE + ' บาท ผ่าน Member Portal');
    return { application_id: app.application_id, status: 'รอชำระเงิน' };
  }

  // กรณีต่อหลังหมดอายุด้วยผล CE หรือสอบข้อเขียนใหม่ — ต้องรอเจ้าหน้าที่อนุมัติก่อนจึงชำระเงิน
  // วิธีสอบข้อเขียนถือเป็นการขึ้นทะเบียนใหม่ด้วยผลสอบ จึงเก็บแค่ค่าขึ้นทะเบียนรับใบอนุญาต (เหมือนขึ้นทะเบียนครั้งแรก) ไม่ใช้ค่าต่ออายุแบบ CE —
  // ค่าสมัครสอบข้อเขียน (EXAM_SUBJECT_FEE ต่อวิชา) ต้องเก็บแยกต่างหากตอนสมัครสอบ ก่อนสอบผ่านแล้วค่อยมายื่นขึ้นทะเบียน/ต่ออายุ ไม่เก็บซ้ำที่ขั้นตอนนี้
  const renewalFee = p.method === 'exam' ? LICENSE_REGISTRATION_FEE
    : (check.isAfterExpiry ? LICENSE_RENEWAL_FEE_AFTER_EXPIRY : LICENSE_RENEWAL_FEE);
  const app = {
    application_id: newId_('REN'), applicant_type: 'ต่ออายุใบอนุญาต', applicant_user_id: user.user_id,
    ref_data_json: JSON.stringify(refData), documents_json: JSON.stringify(p.documents || []),
    status: 'รอตรวจสอบ', reject_reason: '', fee_amount: renewalFee,
    submitted_at: nowIso_(), reviewed_by: '', reviewed_at: '', approved_member_id: ''
  };
  appendRow_('Applications', app);
  mailNotify_(member.email, 'ได้รับคำขอต่ออายุใบอนุญาตแล้ว',
    'ระบบได้รับคำขอต่ออายุของท่านแล้ว เลขที่คำขอ ' + app.application_id + '\nสถานะปัจจุบัน: รอตรวจสอบ');
  return { application_id: app.application_id, status: 'รอตรวจสอบ' };
}

/** แอดมินยืนยันชำระเงินค่าต่ออายุ -> ต่ออายุ/ออกเลขใหม่จริง */
function action_verifyLicenseRenewalPayment(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const payment = findRowById_('Payments', 'payment_id', p.payment_id);
  if (!payment || payment.ref_type !== 'application') throw new Error('ไม่พบรายการชำระเงิน');
  const app = findRowById_('Applications', 'application_id', payment.ref_id);
  if (!app || app.applicant_type !== 'ต่ออายุใบอนุญาต') throw new Error('ไม่พบคำขอต่ออายุที่เกี่ยวข้อง');

  if (p.decision === 'reject') {
    updateRow_('Payments', payment._row, { status: 'ปฏิเสธ', verified_by: user.user_id, verified_at: nowIso_() });
    return { ok: true };
  }

  const refData = JSON.parse(app.ref_data_json || '{}');
  const result = finalizeLicenseRenewal_(refData.member_id, refData.isAfterExpiry, refData.method, user.user_id);

  updateRow_('Payments', payment._row, { status: 'ชำระแล้ว', verified_by: user.user_id, verified_at: nowIso_(), paid_at: nowIso_() });
  try { issueReceiptForPayment_(payment.payment_id, user.user_id); } catch (e) { /* ไม่ให้ล้มเหลวทั้งฟังก์ชันถ้าออกใบเสร็จไม่สำเร็จ */ }
  updateRow_('Applications', app._row, { status: 'ต่ออายุสำเร็จ', approved_member_id: refData.member_id });
  auditLog_('Applications', app.application_id, 'renewal_confirmed', app, result, user.user_id);
  return { ok: true, license_no: result.license_no };
}

function action_myLicense(p, user) {
  requireRole_(user, ['member']);
  if (!user.member_id) throw new Error('บัญชีนี้ยังไม่ผูกกับข้อมูลสมาชิก');
  const license = readAll_('Licenses').find(l => l.member_id === user.member_id);
  return license || null;
}

/** ระบบแสดงใบอนุญาต (TOR ข้อ 6) — แพ็กเกจข้อมูลเต็มสำหรับหน้าบัตรใบอนุญาตดิจิทัลของสมาชิก
 * รวม Access Control ตาม TOR ข้อ 6.2: ถ้าสถานะสมาชิกสามัญไม่ปกติ ต้องปิดเมนูนี้และแจ้งให้ไปต่ออายุสมาชิกก่อน
 */
function action_getMyLicenseCard(p, user) {
  requireRole_(user, ['member']);
  if (!user.member_id) throw new Error('บัญชีนี้ยังไม่ผูกกับข้อมูลสมาชิก');
  const member = findRowById_('Members', 'member_id', user.member_id);
  if (!member) throw new Error('เมนูนี้ใช้สำหรับสมาชิกบุคคลเท่านั้น (สมาชิกวิสามัญไม่มีใบอนุญาตส่วนบุคคล)');

  // TOR ข้อ 6.2 — Access Control: บัตรสมาชิกสามัญหมดอายุ/ไม่ปกติ ต้องปิดเมนูใบอนุญาต
  if (member.member_status !== 'ปกติ') {
    return {
      accessBlocked: true,
      reason: 'สถานะสมาชิกสามัญของท่านไม่ปกติ (' + member.member_status + ') กรุณาต่ออายุสมาชิกสามัญให้เป็นปกติก่อน จึงจะกลับมาใช้งานเมนูใบอนุญาตได้ตามปกติ'
    };
  }
  const membershipDays = daysUntil_(member.membership_expire_date);
  if (membershipDays !== null && membershipDays < 0) {
    return {
      accessBlocked: true,
      reason: 'บัตรสมาชิกสามัญของท่านหมดอายุแล้ว กรุณาต่ออายุสมาชิกสามัญก่อน จึงจะกลับมาใช้งานเมนูใบอนุญาตได้ตามปกติ'
    };
  }

  const license = readAll_('Licenses').find(l => l.member_id === user.member_id);
  let card = readAll_('Cards').find(c => c.member_id === user.member_id);
  if (!card) {
    refreshMemberCard_(user.member_id, member.membership_expire_date, member.member_type);
    card = readAll_('Cards').find(c => c.member_id === user.member_id);
  }

  const renewalApplications = readAll_('Applications')
    .filter(a => a.applicant_type === 'ต่ออายุใบอนุญาต' && a.applicant_user_id === user.user_id)
    .sort((a, b) => new Date(b.submitted_at) - new Date(a.submitted_at));
  // ให้สมาชิกเข้าดู/ติดตามคำขอ "ขึ้นทะเบียนรับใบอนุญาต" ครั้งแรกได้ด้วย เหมือนคำขอต่ออายุ (เดิมมีแค่ renewalApplications เห็นแค่ต่ออายุ)
  const registrationApplications = readAll_('Applications')
    .filter(a => a.applicant_type === 'ขึ้นทะเบียน' && a.applicant_user_id === user.user_id)
    .sort((a, b) => new Date(b.submitted_at) - new Date(a.submitted_at));
  const payments = readAll_('Payments').filter(pm =>
    renewalApplications.some(a => a.application_id === pm.ref_id) ||
    registrationApplications.some(a => a.application_id === pm.ref_id));

  return {
    accessBlocked: false, member: member, license: license, card: card,
    ceCat1: Number(member.ce_score_cat1) || 0, ceCat2: Number(member.ce_score_cat2) || 0,
    ceMinCat1: CE_MINIMUM_CATEGORY1, ceMinCat2: CE_MINIMUM_CATEGORY2,
    renewalApplications: renewalApplications, registrationApplications: registrationApplications, payments: payments,
    signatory: getPresidentForDate_(license ? license.issue_date : nowIso_()),
    sanction: (function () {
      const sn = getActiveSanction_(user.member_id);
      return sn ? { action_type: sn.action_type, effective_date: ymd_(sn.effective_date), suspend_until: ymd_(sn.suspend_until), ethics_case_no: sn.ethics_case_no } : null;
    })()
  };
}

/** ประวัติการดาวน์โหลดใบอนุญาต (ฝั่งแอดมิน) — TOR ข้อ 6.3 */
function action_listLicenseDownloadLog(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  return readAll_('LicenseDownloadLog').sort((a, b) => new Date(b.downloaded_at) - new Date(a.downloaded_at));
}

// ---------- 4.3 สถานะใบอนุญาต ----------
// ใช้กลไกเดียวกับ TOR ข้อ 3.2 ที่ทำไว้แล้ว (recomputeExpiryStatuses คำนวณ ปกติ/ใกล้หมดอายุ/หมดอายุ อัตโนมัติ
// ส่วน อยู่ระหว่างต่ออายุ/ถูกพักใช้/ถูกเพิกถอน ตั้งด้วยมือผ่าน action_updateLicense) — ไม่ต้องเพิ่มโค้ดใหม่

// ---------- 4.4 เครื่องมือจัดการแบบกลุ่ม (Back Office) ----------

/** 4.4.1 ดึงรายชื่อจากระบบ CE — ยังไม่มีการเชื่อมต่อ API จริง (ดู LICENSE_SYSTEM.md เรื่องการขอเชื่อมต่อ) */
function action_pullCeEligibleList(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  throw new Error('ยังไม่ได้เชื่อมต่อ API ระบบการศึกษาต่อเนื่อง (CE) ที่ ' + CE_WEBSITE_URL + ' — เว็บนี้เป็นระบบล็อกอินสำหรับสมาชิก/หน่วยงาน/ผู้ดูแลระบบ ' +
    'ไม่มี API สาธารณะเปิดให้เชื่อมต่อ ต้องขอสิทธิ์เชื่อมต่อจากผู้ดูแลระบบเว็บไซต์นี้โดยตรงก่อน (ดู LICENSE_SYSTEM.md) ' +
    'ระหว่างนี้ใช้เมนู "นำเข้าไฟล์ CSV" แทนได้ (เป็นทางเลือกสำรองที่ TOR อนุญาตไว้)');
}

/** ดึงผลประเมิน CE ของสมาชิกรายคนจากเว็บไซต์ CE จริง — ยังไม่ได้เชื่อมต่อ API จริง (ดูเหตุผลใน action_pullCeEligibleList/LICENSE_SYSTEM.md)
 * เตรียมโครงไว้ให้: ถ้าได้ endpoint+key จริงในอนาคต ให้แก้ฟังก์ชันนี้เรียก UrlFetchApp.fetch() แล้วอัปเดต ce_score_cat1/cat2/ce_verified_via_api/ce_synced_at ตามผลลัพธ์จริง
 */
function action_pullCeStatusForMember(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  throw new Error('ยังไม่ได้เชื่อมต่อ API ระบบการศึกษาต่อเนื่อง (CE) ที่ ' + CE_WEBSITE_URL + ' — กรอกคะแนนด้วยมือในหน้าแก้ไขข้อมูลสมาชิกไปก่อน (ดู LICENSE_SYSTEM.md)');
}

/** 4.4.1 นำเข้ารายชื่อผู้มีสิทธิ์ต่ออายุจากไฟล์ CSV (ทางเลือกสำรอง) — p.rows: [{national_id หรือ member_no}], p.import_type, p.batch_name */
function action_importRenewalEligibilityCsv(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  if (!p.batch_name) throw new Error('กรุณาตั้งชื่อกลุ่ม (batch_name)');
  if (['before_expiry', 'after_expiry'].indexOf(p.import_type) === -1) throw new Error('ประเภทการนำเข้าไม่ถูกต้อง');
  const members = readAll_('Members');
  let matched = 0, notFound = [];

  (p.rows || []).forEach(row => {
    const key = String(row.national_id || row.member_no || '').trim();
    const member = members.find(m => m.national_id === key || m.member_no === key);
    if (!member) { notFound.push(key); return; }
    appendRow_('RenewalEligibility', {
      eligibility_id: newId_('ELIG'), member_id: member.member_id, import_type: p.import_type,
      source: 'CSV', batch_name: p.batch_name, imported_at: nowIso_(), imported_by: user.user_id,
      status: 'รอตรวจสอบ', decided_by: '', decided_at: ''
    });
    matched++;
  });

  return { ok: true, matched: matched, notFound: notFound };
}

/** สร้าง/เพิ่มเข้ากลุ่มอนุมัติแบบกลุ่ม โดยเลือกจากรายชื่อ member_id ตรงๆ (เช่น เลือกจากรายการ "คำขอใบอนุญาต" ที่รอดำเนินการ)
 * แทนที่จะต้องเตรียมไฟล์ CSV แยกต่างหากเหมือน action_importRenewalEligibilityCsv ข้างบน — ใช้กลไกเดียวกัน (ตาราง RenewalEligibility) แค่รับข้อมูลนำเข้าคนละรูปแบบ
 */
function action_addMembersToRenewalBatch(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  if (!p.batch_name) throw new Error('กรุณาตั้งชื่อกลุ่ม (batch_name)');
  if (['before_expiry', 'after_expiry'].indexOf(p.import_type) === -1) throw new Error('ประเภทการนำเข้าไม่ถูกต้อง');
  const memberIds = p.member_ids || [];
  if (!memberIds.length) throw new Error('กรุณาเลือกอย่างน้อย 1 รายการ');
  const members = readAll_('Members');
  let matched = 0, notFound = [];

  memberIds.forEach(memberId => {
    const member = members.find(m => m.member_id === memberId);
    if (!member) { notFound.push(memberId); return; }
    appendRow_('RenewalEligibility', {
      eligibility_id: newId_('ELIG'), member_id: memberId, import_type: p.import_type,
      source: 'คำขอใบอนุญาต (เลือกจากรายการ)', batch_name: p.batch_name, imported_at: nowIso_(), imported_by: user.user_id,
      status: 'รอตรวจสอบ', decided_by: '', decided_at: ''
    });
    matched++;
  });

  return { ok: true, matched: matched, notFound: notFound };
}

function action_listRenewalBatches(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const rows = readAll_('RenewalEligibility');
  const batches = {};
  rows.forEach(r => {
    if (!batches[r.batch_name]) {
      batches[r.batch_name] = { batch_name: r.batch_name, import_type: r.import_type, source: r.source, imported_at: r.imported_at, total: 0, pending: 0, decided: 0 };
    }
    batches[r.batch_name].total++;
    if (r.status === 'รอตรวจสอบ') batches[r.batch_name].pending++; else batches[r.batch_name].decided++;
  });
  return Object.values(batches).sort((a, b) => new Date(b.imported_at) - new Date(a.imported_at));
}

function action_getRenewalBatchMembers(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const rows = readAll_('RenewalEligibility').filter(r => r.batch_name === p.batch_name);
  const members = readAll_('Members');
  return rows.map(r => {
    const m = members.find(x => x.member_id === r.member_id);
    return Object.assign({}, r, {
      member_no: m ? m.member_no : '-',
      name: m ? [m.prefix_th, m.first_name_th, m.last_name_th].filter(Boolean).join(' ') : '(ไม่พบสมาชิก)'
    });
  });
}

/** 4.4.2 อนุมัติ/ไม่อนุมัติทั้งกลุ่มพร้อมกัน — เมื่ออนุมัติจะต่ออายุใบอนุญาตให้ทุกคนในกลุ่มทันที */
function action_decideRenewalBatch(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const rows = readAll_('RenewalEligibility').filter(r => r.batch_name === p.batch_name && r.status === 'รอตรวจสอบ');
  if (!rows.length) throw new Error('ไม่พบรายการรอตรวจสอบในกลุ่มนี้');

  let successCount = 0, failCount = 0;
  const results = [];

  rows.forEach(r => {
    updateRow_('RenewalEligibility', r._row, { status: p.decision === 'approve' ? 'อนุมัติกลุ่มแล้ว' : 'ไม่อนุมัติ', decided_by: user.user_id, decided_at: nowIso_() });
    if (p.decision === 'approve') {
      try {
        const result = finalizeLicenseRenewal_(r.member_id, r.import_type === 'after_expiry', r.import_type === 'after_expiry' ? 'CE_after_group' : 'CE_before_group', user.user_id);
        results.push({ member_id: r.member_id, ok: true, license_no: result.license_no });
        successCount++;
      } catch (err) {
        results.push({ member_id: r.member_id, ok: false, error: err.message });
        failCount++;
      }
    }
  });

  auditLog_('RenewalBatch', p.batch_name, 'batch_' + p.decision, { count: rows.length },
    { successCount: successCount, failCount: failCount }, user.user_id);

  return { ok: true, total: rows.length, successCount: successCount, failCount: failCount, results: results };
}

/** 4.4.3 สั่งพิมพ์ใบอนุญาต (PDF) รายบุคคล — สร้าง Google Doc ชั่วคราวแล้วแปลงเป็น PDF เก็บใน Drive
 * หมายเหตุ: ยังไม่ได้ตกแต่งรูปแบบตามแบบฟอร์มใบอนุญาตจริงของสภาฯ (โลโก้/ตราครุฑ/ลายเซ็นดิจิทัล ฯลฯ) — ดู LICENSE_SYSTEM.md
 */
/** สร้าง PDF ใบอนุญาต — เรียกได้ทั้งแอดมิน (ระบุ member_id ใครก็ได้) และสมาชิกเจ้าของใบอนุญาตเอง (ดาวน์โหลดของตนเอง)
 * ฝัง QR Code ตรวจสอบ (ผ่านบริการสาธารณะ api.qrserver.com — ไม่มีข้อมูลส่วนบุคคลอยู่ใน QR มีแค่ token อ้างอิง)
 * และบันทึก log การดาวน์โหลดทุกครั้งตาม TOR ข้อ 6.3
 */
/** เติมข้อความลายเซ็น (เส้น/ชื่อ/ตำแหน่ง) ลงในช่องตารางของ PDF ใบอนุญาต — ใช้ appendParagraph แทนการต่อ string ด้วย \n
 * เพราะ TableCell.editAsText() จัดการหลายบรรทัด+จัดกึ่งกลางได้ไม่แน่นอนใน DocumentApp
 */
function fillSignatureCell_(cell, signatory) {
  const firstPara = cell.getChild(0).asParagraph();
  firstPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  let signatureInserted = false;
  if (signatory.signature_url) {
    try {
      const sigBlob = getDriveBlobFromUrl_(signatory.signature_url);
      if (sigBlob) {
        firstPara.appendInlineImage(sigBlob).setWidth(120).setHeight(50);
        signatureInserted = true;
      }
    } catch (e) { /* ถ้าแนบรูปลายเซ็นไม่สำเร็จ ใช้เส้นขีดว่างแทน ไม่ทำให้ทั้ง PDF ล้มเหลว */ }
  }
  if (!signatureInserted) firstPara.setText('_______________________');
  cell.appendParagraph('(' + signatory.name + ')').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  cell.appendParagraph(signatory.position).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

function action_generateLicensePdf(p, user) {
  const isAdmin = ['admin_super', 'admin_staff'].indexOf(user.role) !== -1;
  let memberId = p.member_id;
  if (!isAdmin) {
    requireRole_(user, ['member']);
    if (!user.member_id) throw new Error('บัญชีนี้ยังไม่ผูกกับข้อมูลสมาชิก');
    memberId = user.member_id; // สมาชิกดาวน์โหลดได้เฉพาะใบอนุญาตของตนเองเท่านั้น
  }
  const member = findRowById_('Members', 'member_id', memberId);
  if (!member) throw new Error('ไม่พบสมาชิก');
  const license = readAll_('Licenses').find(l => l.member_id === memberId);
  if (!license) throw new Error('สมาชิกรายนี้ยังไม่มีใบอนุญาต');

  const card = readAll_('Cards').find(c => c.member_id === memberId);
  const verifyUrl = (PropertiesService.getScriptProperties().getProperty('FRONTEND_BASE_URL') || '') +
    'verify-card.html?token=' + (card ? card.qr_token : '');
  const president = getPresidentForDate_(license.issue_date);
  const secretary = getSecretaryGeneralForDate_(license.issue_date);
  const memberName = [member.prefix_th, member.first_name_th, member.last_name_th].filter(Boolean).join(' ');

  const doc = DocumentApp.create('ใบอนุญาต_' + license.license_no);
  const body = doc.getBody();
  body.setPageWidth(792).setPageHeight(612); // แนวนอน (Letter landscape) ให้ตรงกับใบอนุญาตจริง (เลขที่ 1848) และใบรับรองสมาชิกวิสามัญ

  // หัวเอกสารแบบ letterhead 3 ช่อง: เลขที่ใบอนุญาต (ซ้าย) / ตราสภาฯ (กลาง) / QR ตรวจสอบ (ขวา — ย้ายมาไว้มุมบนขวา ไม่ให้ทับกับรูปถ่ายสมาชิกที่อยู่ด้านล่าง)
  const headerTable = body.appendTable([['', '', '']]);
  headerTable.setBorderWidth(0);
  const headerLeftCell = headerTable.getCell(0, 0);
  const headerMidCell = headerTable.getCell(0, 1);
  const headerRightCell = headerTable.getCell(0, 2);
  headerLeftCell.getChild(0).asParagraph().setText('ใบอนุญาตเลขที่ ' + license.license_no);
  headerMidCell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  headerRightCell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  try {
    const frontendBase = PropertiesService.getScriptProperties().getProperty('FRONTEND_BASE_URL') || '';
    if (frontendBase) {
      const logoResp = UrlFetchApp.fetch(frontendBase + 'logo.png', { muteHttpExceptions: true });
      if (logoResp.getResponseCode() === 200) {
        headerMidCell.getChild(0).asParagraph().appendInlineImage(logoResp.getBlob()).setWidth(75).setHeight(75);
      }
    }
  } catch (err) { /* ถ้าดึงตราไม่สำเร็จ (เช่น ยังไม่ได้ตั้งค่า FRONTEND_BASE_URL) ข้ามไป ไม่ทำให้การออก PDF ล้มเหลวทั้งหมด */ }
  if (card && card.qr_token) {
    try {
      const qrApiUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=110x110&data=' + encodeURIComponent(verifyUrl);
      const qrResp = UrlFetchApp.fetch(qrApiUrl, { muteHttpExceptions: true });
      if (qrResp.getResponseCode() === 200) {
        headerRightCell.getChild(0).asParagraph().appendInlineImage(qrResp.getBlob()).setWidth(65).setHeight(65);
        const qrLabelPara = headerRightCell.appendParagraph('สแกนเพื่อตรวจสอบความถูกต้อง');
        qrLabelPara.setFontSize(8).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      }
    } catch (err) { /* ถ้าดึง QR ไม่สำเร็จ (เช่น ไม่มีอินเทอร์เน็ตขาออกชั่วคราว) ให้ข้ามไป ไม่ทำให้การออก PDF ล้มเหลวทั้งหมด */ }
  }
  body.appendParagraph(' ');

  // ใช้ heading style เดียวกับใบรับรองสมาชิกวิสามัญ (TITLE/SUBTITLE) ให้ดีไซน์ตรงกัน
  body.appendParagraph('อาศัยอำนาจตามความในพระราชบัญญัติวิชาชีพสังคมสงเคราะห์ พ.ศ. 2556').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('สภาวิชาชีพสังคมสงเคราะห์').setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('ออกใบอนุญาตนี้แก่').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(memberName).setHeading(DocumentApp.ParagraphHeading.SUBTITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(' ');
  body.appendParagraph('ซึ่งได้ขึ้นทะเบียนเป็นผู้ประกอบวิชาชีพสังคมสงเคราะห์รับอนุญาตแล้ว').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('และมีสิทธิประกอบวิชาชีพสังคมสงเคราะห์รับอนุญาต').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('ภายใต้บทบังคับแห่งกฎหมายและข้อบังคับของสภาวิชาชีพสังคมสงเคราะห์').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('ออกให้ ณ วันที่ ' + formatThaiFormalDate_(license.issue_date)).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('ใช้ได้จนถึง วันที่ ' + formatThaiFormalDate_(license.expire_date)).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(' ');

  // แถวเดียว 3 ช่อง: รูปถ่ายสมาชิก (ซ้าย) / ลงนามนายกสภาฯ (กลาง) / ลงนามเลขาธิการ (ขวา) — จัดวางให้ใกล้เคียงแบบฟอร์มจริง
  const table = body.appendTable([['', '', '']]);
  table.setBorderWidth(0);
  const photoCell = table.getCell(0, 0);
  const presidentCell = table.getCell(0, 1);
  const secretaryCell = table.getCell(0, 2);

  const photoBlob = getDriveBlobFromUrl_(member.photo_url);
  if (photoBlob) {
    try { photoCell.appendImage(photoBlob).setWidth(85).setHeight(110); } catch (e) { /* ข้ามรูปถ้าแนบไม่สำเร็จ ไม่ทำให้ทั้ง PDF ล้มเหลว */ }
  }

  fillSignatureCell_(presidentCell, president);
  fillSignatureCell_(secretaryCell, secretary);

  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  const pdfBlob = file.getAs('application/pdf');
  const folder = getDocsFolder_();
  const pdfFile = folder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  file.setTrashed(true); // ลบ Google Doc ต้นฉบับทิ้ง เหลือแต่ PDF

  appendRow_('LicenseDownloadLog', {
    log_id: newId_('DL'), license_no: license.license_no, member_id: memberId,
    downloaded_by: user.user_id, downloaded_at: nowIso_()
  });

  return { member_id: memberId, name: [member.prefix_th, member.first_name_th, member.last_name_th].filter(Boolean).join(' '), url: pdfFile.getUrl() };
}

/** สั่งพิมพ์แบบกลุ่ม — สร้างไฟล์ PDF แยกทีละคน (ยังไม่รวมเป็นไฟล์เดียว ดูหมายเหตุใน LICENSE_SYSTEM.md) */
function action_generateLicensePdfBatch(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const results = [];
  (p.member_ids || []).forEach(memberId => {
    try {
      results.push(action_generateLicensePdf({ member_id: memberId }, user));
    } catch (err) {
      results.push({ member_id: memberId, error: err.message });
    }
  });
  return results;
}

/** 4.4.4 ประวัติการยกเลิกเลขที่ใบอนุญาต */
function action_listVoidedLicenses(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  return readAll_('VoidedLicenses').sort((a, b) => new Date(b.void_date) - new Date(a.void_date));
}

/** เปิด/ปิดช่วงเวลาให้สมาชิกกดต่ออายุใบอนุญาต (สวิตช์รวมทั้งระบบ — เพิ่มเติมจากเงื่อนไข 180 วันอัตโนมัติ) */
function action_setRenewalWindow(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  PropertiesService.getScriptProperties().setProperty('RENEWAL_WINDOW_OPEN', p.open ? 'true' : 'false');
  return { ok: true, open: !!p.open };
}

function action_getRenewalWindowStatus(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  return { open: (PropertiesService.getScriptProperties().getProperty('RENEWAL_WINDOW_OPEN') || 'true') === 'true' };
}

function action_listLicenseRenewalHistory(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  return readAll_('LicenseRenewalHistory').sort((a, b) => new Date(b.renewed_at) - new Date(a.renewed_at));
}

// ---------- 5. ระบบแสดงบัตรสมาชิก (TOR ข้อ 5) ----------

/** ออก/ต่ออายุ QR token ของบัตรสมาชิก — เรียกทุกครั้งที่วันหมดอายุสมาชิกภาพเปลี่ยน (สมัครใหม่/ต่ออายุ)
 * เพื่อไม่ให้ QR เดิม (ที่อาจถูกเพิกถอนไปแล้ว) ยังสแกนผ่านได้ (TOR ข้อ 5.5)
 */
function refreshMemberCard_(memberId, validUntil, cardType) {
  const existing = readAll_('Cards').find(c => c.member_id === memberId);
  const token = Utilities.getUuid();
  if (existing) {
    updateRow_('Cards', existing._row, { qr_token: token, issued_at: nowIso_(), valid_until: validUntil || existing.valid_until, card_type: cardType || existing.card_type });
  } else {
    appendRow_('Cards', {
      card_id: newId_('CARD'), member_id: memberId, card_type: cardType || '',
      qr_token: token, issued_at: nowIso_(), valid_until: validUntil || ''
    });
  }
  return token;
}

/** ข้อมูลบัตรสมาชิกของฉัน (TOR ข้อ 5) — ใช้แสดงหน้าบัตรดิจิทัลทั้งสามัญ/สมทบ/วิสามัญ */
function action_getMyCard(p, user) {
  requireRole_(user, ['member']);
  if (!user.member_id) throw new Error('บัญชีนี้ยังไม่ผูกกับข้อมูลสมาชิก');

  // สมาชิกกิตติมศักดิ์ — บัตรไม่มีค่าบำรุง/ต่ออายุ ใช้วันสิ้นสุดสถานะ (ถ้ามี) เป็นวันหมดอายุ
  const honoraryProfile = findHonoraryById_(user.member_id);
  if (honoraryProfile) {
    let hCard = readAll_('Cards').find(c => c.member_id === user.member_id);
    if (!hCard) {
      refreshMemberCard_(user.member_id, honoraryProfile.term_end, 'กิตติมศักดิ์');
      hCard = readAll_('Cards').find(c => c.member_id === user.member_id);
    }
    return {
      isOrg: false, isHonorary: true, profile: honoraryProfile, license: null, card: hCard,
      renewals: [], renewalApplications: [], payments: [],
      signatory: getSecretaryGeneralForDate_(hCard.issued_at)
    };
  }

  const orgProfile = findRowById_('MembersOrg', 'org_id', user.member_id);
  const isOrg = !!orgProfile;
  const profile = isOrg ? orgProfile : findRowById_('Members', 'member_id', user.member_id);
  if (!profile) throw new Error('ไม่พบข้อมูลสมาชิก');

  let card = readAll_('Cards').find(c => c.member_id === user.member_id);
  if (!card) {
    // เผื่อกรณีเป็นสมาชิกอยู่ก่อนระบบนี้จะถูกสร้าง — ออกบัตรให้ครั้งแรกตอนเข้าดูหน้านี้
    const validUntil = isOrg ? profile.end_date : profile.membership_expire_date;
    refreshMemberCard_(user.member_id, validUntil, isOrg ? 'วิสามัญ' : profile.member_type);
    card = readAll_('Cards').find(c => c.member_id === user.member_id);
  }

  const license = isOrg ? null : readAll_('Licenses').find(l => l.member_id === user.member_id);
  const renewals = readAll_('MembershipRenewals').filter(r => r.member_id === user.member_id)
    .sort((a, b) => new Date(b.renewed_at) - new Date(a.renewed_at));
  const renewalApplications = readAll_('Applications')
    .filter(a => a.applicant_type === 'ต่ออายุสมาชิก' && a.applicant_user_id === user.user_id)
    .sort((a, b) => new Date(b.submitted_at) - new Date(a.submitted_at));
  const payments = readAll_('Payments').filter(pm =>
    renewalApplications.some(a => a.application_id === pm.ref_id)
  );

  return {
    isOrg: isOrg, profile: profile, license: license, card: card,
    renewals: renewals, renewalApplications: renewalApplications, payments: payments,
    signatory: getSecretaryGeneralForDate_(card.issued_at)
  };
}

/** 5.1 ตรวจสอบสิทธิ์กดปุ่ม "ต่ออายุสมาชิก" — ล่วงหน้าได้ไม่เกิน 90 วันก่อนบัตรหมดอายุ */
function action_getMembershipRenewalEligibility(p, user) {
  requireRole_(user, ['member']);
  if (!user.member_id) return { eligible: false, reason: 'บัญชีนี้ยังไม่ผูกกับข้อมูลสมาชิก' };
  const member = findRowById_('Members', 'member_id', user.member_id);
  if (!member) return { eligible: false, reason: 'เมนูนี้ใช้สำหรับสมาชิกบุคคลเท่านั้น (สมาชิกวิสามัญติดต่อเจ้าหน้าที่โดยตรง)' };
  const days = daysUntil_(member.membership_expire_date);
  const isExpired = days !== null && days < 0;
  if (days !== null && days > MEMBERSHIP_RENEWAL_WINDOW_DAYS) {
    return {
      eligible: false, daysUntilExpire: days,
      reason: 'โปรดต่ออายุบัตรสมาชิกล่วงหน้า 90 วัน ก่อนบัตรหมดอายุ (อีก ' + (days - MEMBERSHIP_RENEWAL_WINDOW_DAYS) + ' วันจะเปิดให้ต่ออายุ)'
    };
  }
  const pending = readAll_('Applications').find(a =>
    a.applicant_type === 'ต่ออายุสมาชิก' && a.applicant_user_id === user.user_id &&
    ['รอตรวจสอบ', 'รอชำระเงิน'].indexOf(a.status) !== -1
  );
  if (pending) return { eligible: false, reason: 'มีคำขอต่ออายุสมาชิกที่ยังไม่เสร็จสิ้นอยู่แล้ว (เลขที่ ' + pending.application_id + ')', daysUntilExpire: days };
  return { eligible: true, daysUntilExpire: days, isExpired: isExpired };
}

/** 5.1 กดต่ออายุสมาชิก — สร้างรายการรอชำระเงินทันที (ไม่ต้องรออนุมัติ เพราะเป็นแค่การต่ออายุตามปกติ) */
function action_submitMembershipRenewal(p, user) {
  requireRole_(user, ['member']);
  const check = action_getMembershipRenewalEligibility(p, user);
  if (!check.eligible) throw new Error(check.reason);
  const member = findRowById_('Members', 'member_id', user.member_id);
  const duration = p.duration || '1';
  const fee = getMembershipFee_(member.member_type, duration, member.dob);

  const appId = newId_('MREN');
  appendRow_('Applications', {
    application_id: appId, applicant_type: 'ต่ออายุสมาชิก', applicant_user_id: user.user_id,
    ref_data_json: JSON.stringify({ member_id: user.member_id, duration: duration, email: member.email, name: [member.prefix_th, member.first_name_th, member.last_name_th].filter(Boolean).join(' ') }),
    documents_json: '[]', status: 'รอชำระเงิน', reject_reason: '', fee_amount: fee,
    submitted_at: nowIso_(), reviewed_by: 'ระบบ (ต่ออายุสมาชิกโดยตรง)', reviewed_at: nowIso_(), approved_member_id: ''
  });
  appendRow_('Payments', {
    payment_id: newId_('PAY'), ref_type: 'application', ref_id: appId, member_id: user.member_id,
    amount: fee, slip_url: '', method: '', status: 'รอชำระ', verified_by: '', verified_at: '', paid_at: ''
  });
  mailNotify_(member.email, 'ยื่นต่ออายุสมาชิกภาพสำเร็จ — กรุณาชำระเงิน',
    'กรุณาชำระค่าธรรมเนียม ' + fee + ' บาท ผ่าน Member Portal แล้วแนบสลิปการโอนเงิน');
  return { application_id: appId };
}

/** แอดมินยืนยันชำระเงินค่าต่ออายุสมาชิก -> ต่ออายุจริง + ออก QR ใหม่ */
function action_verifyMembershipRenewalPayment(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const payment = findRowById_('Payments', 'payment_id', p.payment_id);
  if (!payment || payment.ref_type !== 'application') throw new Error('ไม่พบรายการชำระเงิน');
  const app = findRowById_('Applications', 'application_id', payment.ref_id);
  if (!app || app.applicant_type !== 'ต่ออายุสมาชิก') throw new Error('ไม่พบคำขอต่ออายุสมาชิกที่เกี่ยวข้อง');

  if (p.decision === 'reject') {
    updateRow_('Payments', payment._row, { status: 'ปฏิเสธ', verified_by: user.user_id, verified_at: nowIso_() });
    return { ok: true };
  }

  const refData = JSON.parse(app.ref_data_json || '{}');
  const member = findRowById_('Members', 'member_id', refData.member_id);
  if (!member) throw new Error('ไม่พบสมาชิกที่เกี่ยวข้อง');

  const baseDate = member.membership_expire_date && new Date(member.membership_expire_date) > new Date()
    ? member.membership_expire_date : nowIso_();
  const newExpire = computeMembershipExpiry_(baseDate, refData.duration || '1');

  updateRow_('Members', member._row, { membership_expire_date: newExpire, member_status: 'ปกติ', updated_at: nowIso_() });
  appendRow_('MembershipRenewals', {
    renewal_id: newId_('RNW'), member_id: member.member_id, renewed_at: nowIso_(),
    old_expire_date: member.membership_expire_date || '', new_expire_date: newExpire,
    years: refData.duration || '1', payment_id: payment.payment_id, renewed_by: user.user_id, note: 'ต่ออายุผ่าน Member Portal (self-service)'
  });
  refreshMemberCard_(member.member_id, newExpire, member.member_type);

  updateRow_('Payments', payment._row, { status: 'ชำระแล้ว', verified_by: user.user_id, verified_at: nowIso_(), paid_at: nowIso_() });
  try { issueReceiptForPayment_(payment.payment_id, user.user_id); } catch (e) { /* ไม่ให้ล้มเหลวทั้งฟังก์ชันถ้าออกใบเสร็จไม่สำเร็จ */ }
  updateRow_('Applications', app._row, { status: 'ต่ออายุสำเร็จ', approved_member_id: member.member_id });

  mailNotify_(member.email, 'ต่ออายุสมาชิกภาพสำเร็จ',
    'สมาชิกภาพของท่านได้รับการต่ออายุแล้ว\nวันหมดอายุใหม่: ' + new Date(newExpire).toLocaleDateString('th-TH'));
  return { ok: true, new_expire_date: newExpire };
}

/** รายการรอตรวจสอบการชำระเงินต่ออายุสมาชิก (ฝั่งแอดมิน) */
function action_listMembershipRenewalPayments(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const renewalAppIds = readAll_('Applications').filter(a => a.applicant_type === 'ต่ออายุสมาชิก').map(a => a.application_id);
  return readAll_('Payments').filter(pm => pm.ref_type === 'application' && renewalAppIds.indexOf(pm.ref_id) !== -1 && pm.status === 'รอตรวจสอบ');
}

/** 5.5 ตรวจสอบความถูกต้องผ่าน QR Code — PUBLIC ไม่ต้อง login ใช้ฐานข้อมูลเดียวกับข้อ 3.4 */
function action_verifyQrToken(p) {
  const card = readAll_('Cards').find(c => c.qr_token === p.token);
  if (!card) {
    return { valid: false, message: 'ไม่พบข้อมูลนี้ในระบบ หรือ QR Code นี้ถูกยกเลิกไปแล้ว (อาจมีการออกบัตร/ต่ออายุใหม่ทับ)' };
  }
  appendRow_('QrScanLog', { scan_id: newId_('SCAN'), qr_token: p.token, member_id: card.member_id, scanned_at: nowIso_() });

  if (card.card_type === 'วิสามัญ') {
    const org = findRowById_('MembersOrg', 'org_id', card.member_id);
    if (!org) return { valid: false, message: 'ไม่พบข้อมูลองค์กรนี้ในระบบ' };
    return {
      valid: true, kind: 'organization', name: org.org_name_th, org_name_en: org.org_name_en,
      reg_no: org.reg_no, status: org.status, expire_date: org.end_date
    };
  }
  if (card.card_type === 'กิตติมศักดิ์') {
    const hn = findHonoraryById_(card.member_id);
    if (!hn) return { valid: false, message: 'ไม่พบข้อมูลสมาชิกกิตติมศักดิ์นี้ในระบบ' };
    return {
      valid: true, kind: 'honorary',
      name: [hn.prefix_th, hn.first_name_th, hn.last_name_th].filter(Boolean).join(' ').replace(/\s+/g, ' '),
      member_no: hn.honorary_no, member_type: 'กิตติมศักดิ์', qualification_type: hn.qualification_type,
      status: hn.status === 'ดำรงสถานะ' ? 'ปกติ' : 'สิ้นสุดสถานะ', expire_date: hn.term_end || ''
    };
  }
  const member = findRowById_('Members', 'member_id', card.member_id);
  if (!member) return { valid: false, message: 'ไม่พบข้อมูลสมาชิกนี้ในระบบ' };
  const displayName = formatMemberDisplayName_(member, member.license_status);
  return {
    valid: true, kind: 'individual',
    name: displayName.th, usedProfessionalTitle: displayName.usedProfessionalTitle,
    member_no: member.member_no, member_type: member.member_type,
    status: member.member_status, license_no: member.license_no, license_status: member.license_status,
    expire_date: member.membership_expire_date
  };
}

/** ใบรับรองสมาชิกวิสามัญ (TOR ข้อ 5.4) — PDF ความละเอียดสูง เรียกได้ทั้งแอดมินและองค์กรเจ้าของบัญชีเอง */
function action_generateOrgCertificatePdf(p, user) {
  const isAdmin = ['admin_super', 'admin_staff'].indexOf(user.role) !== -1;
  const orgId = isAdmin ? p.org_id : user.member_id;
  if (!isAdmin) requireRole_(user, ['member']);
  const org = findRowById_('MembersOrg', 'org_id', orgId);
  if (!org) throw new Error('ไม่พบข้อมูลองค์กร');

  const doc = DocumentApp.create('ใบรับรองสมาชิกวิสามัญ_' + org.reg_no);
  const body = doc.getBody();
  body.setPageWidth(792).setPageHeight(612); // แนวนอน (Letter landscape) ตาม TOR ข้อ 5.4
  body.appendParagraph('ใบรับรองสมาชิกวิสามัญ').setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('สภาวิชาชีพสังคมสงเคราะห์ — Social Work Professions Council').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(' ');
  body.appendParagraph('ขอรับรองว่า').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(org.org_name_th + (org.org_name_en ? ' (' + org.org_name_en + ')' : '')).setHeading(DocumentApp.ParagraphHeading.SUBTITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('เป็นสมาชิกวิสามัญของสภาวิชาชีพสังคมสงเคราะห์ เลขทะเบียน ' + org.reg_no).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('วันที่อนุมัติ: ' + new Date(org.approved_date).toLocaleDateString('th-TH') +
    '   วันสิ้นสุดสมาชิกภาพ: ' + (org.end_date ? new Date(org.end_date).toLocaleDateString('th-TH') : '-')).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(' ');
  const signatory = getSecretaryGeneralForDate_(nowIso_());
  const sigPara = body.appendParagraph('');
  sigPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  let orgSigInserted = false;
  if (signatory.signature_url) {
    try {
      const sigBlob = getDriveBlobFromUrl_(signatory.signature_url);
      if (sigBlob) { sigPara.appendInlineImage(sigBlob).setWidth(120).setHeight(50); orgSigInserted = true; }
    } catch (e) { /* ถ้าแนบรูปลายเซ็นไม่สำเร็จ ใช้เส้นขีดว่างแทน ไม่ทำให้ทั้ง PDF ล้มเหลว */ }
  }
  if (!orgSigInserted) sigPara.setText('_____________________________');
  body.appendParagraph('(' + signatory.name + ')').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(signatory.position).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  const pdfBlob = file.getAs('application/pdf');
  const folder = getDocsFolder_();
  const pdfFile = folder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  file.setTrashed(true);

  return { url: pdfFile.getUrl(), org_name: org.org_name_th };
}

function action_listLicenseApplications(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  return readAll_('Applications')
    .filter(a => ['ขึ้นทะเบียน', 'ต่ออายุใบอนุญาต'].indexOf(a.applicant_type) !== -1)
    .filter(a => !p.applicant_type || a.applicant_type === p.applicant_type)
    .sort((a, b) => new Date(b.submitted_at) - new Date(a.submitted_at));
}

// ---------- 6. ระบบสมัครสอบข้อเขียน/สอบสัมภาษณ์ (TOR ข้อ 4 — รองรับหลายรอบสอบต่อปี) ----------
// ใช้ชีต ExamRounds ใหม่จัดการรอบสอบ (แอดมินเปิด/ปิดรอบได้เอง) และใช้ชีต Applications/Payments เดิมร่วมกับฟีเจอร์อื่น
// (เพิ่ม applicant_type ใหม่ 2 ค่า 'สมัครสอบข้อเขียน'/'สมัครสอบสัมภาษณ์' — สอดคล้องกับรูปแบบที่ใช้กับคำขอขึ้นทะเบียน/ต่ออายุใบอนุญาตอยู่แล้ว)
// ค่าสมัครสอบเป็นคนละรายการกับค่าขึ้นทะเบียน/ต่ออายุใบอนุญาตเสมอ — ต้องจ่ายค่าสมัครสอบและสอบผ่านก่อน จึงยื่นขึ้นทะเบียน/ต่ออายุได้ (ดูหมายเหตุที่ action_submitLicenseRenewal)

// ----- 6.1 แอดมิน: จัดการรอบสอบ -----

/** สมาชิกและแอดมินเรียกดูได้ทั้งคู่ — สมาชิกใช้เลือกรอบตอนสมัครสอบ (ส่ง p.openOnly:true เพื่อกรองเฉพาะรอบที่เปิดรับสมัครและยังไม่เลยวันปิดรับ) */
function action_listExamRounds(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff', 'member']);
  let rows = readAll_('ExamRounds');
  if (p.exam_type) rows = rows.filter(r => r.exam_type === p.exam_type);
  if (p.status) rows = rows.filter(r => r.status === p.status);
  if (p.openOnly) {
    const today = nowIso_().slice(0, 10);
    rows = rows.filter(r => r.status === 'เปิดรับสมัคร' && (!r.close_date || r.close_date.slice(0, 10) >= today));
  }
  return rows.sort((a, b) => new Date(b.open_date || 0) - new Date(a.open_date || 0));
}

function action_createExamRound(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  if (EXAM_TYPES.indexOf(p.exam_type) === -1) throw new Error('ประเภทการสอบไม่ถูกต้อง');
  if (!p.round_label) throw new Error('กรุณาระบุชื่อ/เลขที่รอบสอบ (เช่น 2/2569)');
  if (!p.close_date) throw new Error('กรุณาระบุวันปิดรับสมัคร');
  const round = {
    round_id: newId_('EXR'), exam_type: p.exam_type, round_label: p.round_label,
    open_date: p.open_date || nowIso_(), close_date: p.close_date, exam_date: p.exam_date || '',
    format_options_json: JSON.stringify(p.format_options || EXAM_FORMAT_OPTIONS),
    location_note: p.location_note || '', status: 'เปิดรับสมัคร', note: p.note || '',
    created_by: user.user_id, created_at: nowIso_(), updated_at: nowIso_()
  };
  appendRow_('ExamRounds', round);
  auditLog_('ExamRounds', round.round_id, 'create', {}, round, user.user_id);
  return { round_id: round.round_id };
}

function action_updateExamRound(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const round = findRowById_('ExamRounds', 'round_id', p.round_id);
  if (!round) throw new Error('ไม่พบรอบสอบนี้');
  const patch = { updated_at: nowIso_() };
  if (p.round_label) patch.round_label = p.round_label;
  if (p.open_date) patch.open_date = p.open_date;
  if (p.close_date) patch.close_date = p.close_date;
  if (p.exam_date !== undefined) patch.exam_date = p.exam_date;
  if (p.format_options) patch.format_options_json = JSON.stringify(p.format_options);
  if (p.location_note !== undefined) patch.location_note = p.location_note;
  if (p.note !== undefined) patch.note = p.note;
  if (p.status) {
    if (EXAM_ROUND_STATUSES.indexOf(p.status) === -1) throw new Error('สถานะรอบสอบไม่ถูกต้อง');
    patch.status = p.status;
  }
  updateRow_('ExamRounds', round._row, patch);
  auditLog_('ExamRounds', round.round_id, 'update', round, patch, user.user_id);
  return { ok: true };
}

// ----- 6.2 สมาชิก: สมัครสอบ -----

/** ตรวจสอบสิทธิ์สมัครสอบ — excludeApplicationId ใช้ตอนแก้ไขใบสมัครเดิม (ไม่ให้เช็คชนกับใบสมัครของตัวเอง) */
function getExamApplicationEligibility_(p, user, excludeApplicationId) {
  requireRole_(user, ['member']);
  if (!user.member_id) return { eligible: false, reason: 'บัญชีนี้ยังไม่ผูกกับข้อมูลสมาชิก' };
  const member = findRowById_('Members', 'member_id', user.member_id);
  if (!member) return { eligible: false, reason: 'ไม่พบข้อมูลสมาชิก' };
  if (['สามัญ', 'สมทบ'].indexOf(member.member_type) === -1) return { eligible: false, reason: 'เมนูนี้ใช้ได้เฉพาะสมาชิกสามัญ/สมทบเท่านั้น' };
  // ต้องเป็นสมาชิกสถานะปกติ (ไม่หมดอายุ/พักใช้/เพิกถอน) จึงจะสมัครสอบได้ — ถ้าไม่ปกติบังคับให้ไปต่ออายุสมาชิกภาพก่อน
  if (member.member_status !== 'ปกติ') {
    return {
      eligible: false, needsRenewal: true,
      reason: 'สถานะสมาชิกภาพของท่านไม่ปกติ (' + member.member_status + ') กรุณาต่ออายุสมาชิกภาพให้เป็นปกติก่อน จึงจะสมัครสอบได้'
    };
  }
  if (EXAM_TYPES.indexOf(p.exam_type) === -1) return { eligible: false, reason: 'ประเภทการสอบไม่ถูกต้อง' };
  // สมัครสอบข้อเขียนต้องมีข้อมูลวุฒิการศึกษาบันทึกไว้ในโปรไฟล์ก่อน (ใช้แสดงในใบสมัครตามแบบฟอร์ม สวชส.5) — ถ้ายังไม่มีให้ไปบันทึกที่หน้า "ข้อมูลของฉัน" ก่อน
  let educationInfo = null;
  if (p.exam_type === 'เขียน') {
    let edu = [];
    try { edu = JSON.parse(member.education_json || '[]'); } catch (e) {}
    if (!edu.length) {
      return {
        eligible: false, needsProfile: true,
        reason: 'กรุณาไปบันทึกข้อมูลวุฒิการศึกษาที่หน้า "ข้อมูลของฉัน" ให้ครบก่อนสมัครสอบข้อเขียน'
      };
    }
    const order = ['ปริญญาเอก', 'ปริญญาโท', 'ปริญญาตรี', 'ปวส.', 'ปวช.', 'มัธยมศึกษา'];
    educationInfo = edu[0];
    for (let i = 0; i < order.length; i++) {
      const match = edu.find(e => (e.level || '').indexOf(order[i]) !== -1);
      if (match) { educationInfo = match; break; }
    }
  }
  const round = findRowById_('ExamRounds', 'round_id', p.round_id);
  if (!round || round.exam_type !== p.exam_type) return { eligible: false, reason: 'ไม่พบรอบสอบนี้' };
  if (round.status !== 'เปิดรับสมัคร') return { eligible: false, reason: 'รอบสอบนี้ปิดรับสมัครแล้ว' };
  const today = nowIso_().slice(0, 10);
  if (round.close_date && round.close_date.slice(0, 10) < today) return { eligible: false, reason: 'เกินวันปิดรับสมัครของรอบนี้แล้ว (' + round.close_date.slice(0, 10) + ')' };
  const applicantType = p.exam_type === 'เขียน' ? 'สมัครสอบข้อเขียน' : 'สมัครสอบสัมภาษณ์';
  const pending = readAll_('Applications').find(a => a.applicant_type === applicantType && a.applicant_user_id === user.user_id &&
    a.application_id !== excludeApplicationId &&
    ['รอตรวจสอบ', 'รอชำระเงิน', 'รอผลสอบ'].indexOf(a.status) !== -1 &&
    (JSON.parse(a.ref_data_json || '{}').round_id === p.round_id));
  if (pending) return { eligible: false, reason: 'มีใบสมัครสอบรอบนี้ที่ยังไม่เสร็จสิ้นอยู่แล้ว (เลขที่ ' + pending.application_id + ')' };
  return {
    eligible: true, round: round, examSubjects: EXAM_SUBJECTS, examSubjectFee: EXAM_SUBJECT_FEE,
    essayCategories: EXAM_INTERVIEW_ESSAY_CATEGORIES, examHistoryOptions: EXAM_HISTORY_LABELS,
    examHistoryDateRequired: EXAM_HISTORY_DATE_REQUIRED, educationDocType: member.member_type === 'สมทบ' ? 'เอกสารรับรองการศึกษา' : 'สำเนาวุฒิการศึกษา',
    member: { position: member.position || '', agency: member.agency || '', phone: member.phone || '', work_status: member.work_status || '' },
    education: educationInfo
  };
}

function action_getExamApplicationEligibility(p, user) {
  return getExamApplicationEligibility_(p, user, p.exclude_application_id || null);
}

/** รวม validation + สร้างข้อมูลใบสมัครสอบ ใช้ร่วมกันทั้งตอนสมัครใหม่และตอนแก้ไขใบสมัครเดิม
 * p.exam_type: 'เขียน'|'สัมภาษณ์', p.round_id, p.subjects (เขียน), p.format ('online'|'onsite'), p.essay_category (สัมภาษณ์)
 * p.position, p.agency, p.phone (ทั้ง 2 แบบ), p.exam_history + p.consent_no_refund (เฉพาะเขียน) */
function buildExamApplicationPayload_(p, user, check) {
  const member = findRowById_('Members', 'member_id', user.member_id);
  if (EXAM_FORMAT_OPTIONS.indexOf(p.format) === -1) throw new Error('กรุณาเลือกรูปแบบการสอบ (online/onsite)');

  const docs = p.documents || [];
  let fee = 0;
  const refData = {
    member_id: user.member_id, round_id: p.round_id, round_label: check.round.round_label,
    format: p.format, email: member.email, name: [member.prefix_th, member.first_name_th, member.last_name_th].filter(Boolean).join(' '),
    position: member.position || '', agency: member.agency || '', phone: member.phone || '', work_status: member.work_status || ''
  };

  if (p.exam_type === 'เขียน') {
    const subjects = (p.subjects || []).filter(s => EXAM_SUBJECTS.indexOf(s) !== -1);
    if (!subjects.length) throw new Error('กรุณาเลือกวิชาที่ต้องการสมัครสอบอย่างน้อย 1 วิชา');
    if (EXAM_HISTORY_OPTIONS.indexOf(p.exam_history) === -1) throw new Error('กรุณาเลือกประวัติการสอบข้อเขียนที่ผ่านมา');
    if (EXAM_HISTORY_DATE_REQUIRED.indexOf(p.exam_history) !== -1 && !p.exam_history_date) {
      throw new Error('กรุณาระบุวันที่สอบผ่านวิชาดังกล่าว');
    }
    if (!p.consent_no_refund) throw new Error('กรุณายอมรับเงื่อนไขการสมัครสอบ (เมื่อสมัครแล้วไม่สามารถขอคืนเงิน/เปลี่ยนรอบสอบได้) ก่อนยื่นใบสมัคร');
    const hasId = docs.some(d => d.doc_type === 'สำเนาบัตรประชาชน');
    const hasPhoto = docs.some(d => d.doc_type === 'รูปถ่าย');
    if (!hasId || !hasPhoto) throw new Error('กรุณาแนบสำเนาบัตรประชาชนและรูปถ่ายให้ครบ');
    // สมาชิกสมทบต้องแนบ "เอกสารรับรองการศึกษา" ส่วนสมาชิกสามัญแนบ "สำเนาวุฒิการศึกษา" — บังคับทั้งสองประเภท
    if (!docs.some(d => d.doc_type === check.educationDocType)) {
      throw new Error('กรุณาแนบ' + check.educationDocType);
    }
    fee = EXAM_SUBJECT_FEE * subjects.length;
    refData.subjects = subjects;
    refData.exam_history = p.exam_history;
    refData.exam_history_date = EXAM_HISTORY_DATE_REQUIRED.indexOf(p.exam_history) !== -1 ? p.exam_history_date : '';
    refData.consent_no_refund = true;
    refData.education = check.education || null;
  } else {
    if (!p.essay_category || EXAM_INTERVIEW_ESSAY_CATEGORIES.indexOf(p.essay_category) === -1) {
      throw new Error('กรุณาเลือกประเภทเอกสารแสดงทักษะและทัศนคติทางสังคมสงเคราะห์');
    }
    const essayTypes = examEssayDocTypes_(p.essay_category);
    if (!essayTypes.every(t => docs.some(d => d.doc_type === t))) {
      throw new Error('กรุณาแนบไฟล์เอกสารตามประเภทที่เลือกให้ครบทั้งไฟล์ PDF และไฟล์ Word (ตั้งชื่อไฟล์เป็นชื่อ-นามสกุลของท่าน ไม่ต้องใส่คำนำหน้า)');
    }
    if (!docs.some(d => d.doc_type === 'หลักฐานสอบผ่านข้อเขียน/ผลการศึกษาต่อเนื่อง (CE)')) {
      throw new Error('กรุณาแนบหลักฐานการสอบผ่านข้อเขียน 2 วิชา หรือหลักฐานผ่านเกณฑ์ CE (สำหรับผู้ต่ออายุใบอนุญาต)');
    }
    fee = 0; // สมัครสอบสัมภาษณ์ไม่มีค่าธรรมเนียม (ตามแบบฟอร์ม สวชส.8)
    refData.essay_category = p.essay_category;
  }

  const applicantType = p.exam_type === 'เขียน' ? 'สมัครสอบข้อเขียน' : 'สมัครสอบสัมภาษณ์';
  return { refData: refData, docs: docs, fee: fee, applicantType: applicantType, member: member };
}

function action_submitExamApplication(p, user) {
  requireRole_(user, ['member']);
  const check = getExamApplicationEligibility_(p, user, null);
  if (!check.eligible) throw new Error(check.reason);
  const built = buildExamApplicationPayload_(p, user, check);

  const app = {
    application_id: newId_('EXA'), applicant_type: built.applicantType, applicant_user_id: user.user_id,
    ref_data_json: JSON.stringify(built.refData), documents_json: JSON.stringify(built.docs), status: 'รอตรวจสอบ',
    reject_reason: '', fee_amount: built.fee, submitted_at: nowIso_(), reviewed_by: '', reviewed_at: '', approved_member_id: ''
  };
  appendRow_('Applications', app);
  mailNotify_(built.member.email, 'ได้รับใบสมัครสอบ' + (p.exam_type === 'เขียน' ? 'ข้อเขียน' : 'สัมภาษณ์') + 'แล้ว',
    'ระบบได้รับใบสมัครสอบของท่านแล้ว รอบ ' + check.round.round_label + ' เลขที่คำขอ ' + app.application_id +
    '\nสถานะปัจจุบัน: รอตรวจสอบ');
  return { application_id: app.application_id, fee_amount: built.fee };
}

/** สมาชิกแก้ไขใบสมัครสอบของตนเองได้ ตราบใดที่ยังอยู่สถานะ "รอตรวจสอบ" (แอดมินยังไม่พิจารณา) */
function action_updateExamApplication(p, user) {
  requireRole_(user, ['member']);
  const app = findRowById_('Applications', 'application_id', p.application_id);
  if (!app || EXAM_APPLICANT_TYPES.indexOf(app.applicant_type) === -1) throw new Error('ไม่พบใบสมัครสอบนี้');
  if (app.applicant_user_id !== user.user_id) throw new Error('ไม่มีสิทธิ์แก้ไขใบสมัครนี้');
  if (app.status !== 'รอตรวจสอบ') throw new Error('แก้ไขได้เฉพาะใบสมัครที่ยังอยู่สถานะ "รอตรวจสอบ" เท่านั้น (แอดมินพิจารณาแล้วจะแก้ไขไม่ได้)');

  const examType = app.applicant_type === 'สมัครสอบข้อเขียน' ? 'เขียน' : 'สัมภาษณ์';
  const pOverride = Object.assign({}, p, { exam_type: examType });
  const check = getExamApplicationEligibility_(pOverride, user, app.application_id);
  if (!check.eligible) throw new Error(check.reason);
  const built = buildExamApplicationPayload_(pOverride, user, check);

  updateRow_('Applications', app._row, {
    ref_data_json: JSON.stringify(built.refData), documents_json: JSON.stringify(built.docs),
    fee_amount: built.fee, submitted_at: nowIso_()
  });
  auditLog_('Applications', app.application_id, 'exam_application_edited', app, built.refData, user.user_id);
  return { application_id: app.application_id, fee_amount: built.fee };
}

/** สมาชิกดูรายการใบสมัครสอบของตนเอง พร้อมข้อมูลรอบสอบที่เกี่ยวข้อง */
function action_myExamApplications(p, user) {
  requireRole_(user, ['member']);
  const rounds = readAll_('ExamRounds');
  return readAll_('Applications')
    .filter(a => a.applicant_user_id === user.user_id && EXAM_APPLICANT_TYPES.indexOf(a.applicant_type) !== -1)
    .map(a => {
      const d = JSON.parse(a.ref_data_json || '{}');
      const round = rounds.find(r => r.round_id === d.round_id);
      return Object.assign({}, fixExamRefRoundLabel_(a, round), { round: round || null });
    })
    .sort((a, b) => new Date(b.submitted_at) - new Date(a.submitted_at));
}

// ----- 6.3 แอดมิน: ตรวจสอบคำขอ/ชำระเงิน/บันทึกผลสอบ -----

/** เลือกวุฒิการศึกษาสูงสุดจาก education_json ของสมาชิก (คืน {level, institute, degree, year, major} หรือ null ถ้าไม่มีข้อมูล) */
function pickMainEducation_(educationJson) {
  let edu = [];
  try { edu = JSON.parse(educationJson || '[]'); } catch (e) {}
  if (!edu.length) return null;
  const order = ['ปริญญาเอก', 'ปริญญาโท', 'ปริญญาตรี', 'ปวส.', 'ปวช.', 'มัธยมศึกษา'];
  for (let i = 0; i < order.length; i++) {
    const match = edu.find(e => (e.level || '').indexOf(order[i]) !== -1);
    if (match) return match;
  }
  return edu[0];
}

function action_listExamApplications(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const rounds = readAll_('ExamRounds');
  const members = readAll_('Members');
  const cards = readAll_('Cards');
  let rows = readAll_('Applications').filter(a => EXAM_APPLICANT_TYPES.indexOf(a.applicant_type) !== -1);
  if (p.applicant_type) rows = rows.filter(a => a.applicant_type === p.applicant_type);
  if (p.status) rows = rows.filter(a => a.status === p.status);
  rows = rows.map(a => {
    const d = JSON.parse(a.ref_data_json || '{}');
    const round = rounds.find(r => r.round_id === d.round_id);
    const member = d.member_id ? members.find(m => m.member_id === d.member_id) : null;
    const card = d.member_id ? cards.find(c => c.member_id === d.member_id) : null;
    return Object.assign({}, fixExamRefRoundLabel_(a, round), {
      round: round || null,
      member_education: member ? pickMainEducation_(member.education_json) : null,
      member_work: member ? { work_status: member.work_status || '', position: member.position || '', agency: member.agency || '', affiliation: member.affiliation || '' } : null,
      member_no: member ? (member.member_no || '') : '',
      membership_expire_date: member ? (member.membership_expire_date || '') : '',
      card_issued_at: card ? (card.issued_at || '') : ''
    });
  });
  if (p.round_id) rows = rows.filter(a => a.round && a.round.round_id === p.round_id);
  return attachMemberPhoto_(rows.sort((a, b) => new Date(b.submitted_at) - new Date(a.submitted_at)));
}

/** กำหนด "รหัสประจำตัวสอบ" (เลข 4 หลัก เรียงตามลำดับภายในรอบสอบ/ประเภทการสอบเดียวกัน) ให้ใบสมัครที่ผ่านการอนุมัติ/ชำระเงินแล้ว (เข้าสถานะ "รอผลสอบ")
 * เรียกตอนเข้าสถานะ "รอผลสอบ" เท่านั้น (ทั้งกรณีอนุมัติตรงๆ ไม่มีค่าธรรมเนียม และกรณียืนยันชำระเงินแล้ว) ถ้ามีรหัสอยู่แล้วคืนรหัสเดิม ไม่สร้างซ้ำ */
function assignExamCandidateNo_(applicantType, data) {
  if (data.candidate_no) return data.candidate_no;
  const count = readAll_('Applications').filter(a => {
    if (a.applicant_type !== applicantType) return false;
    try {
      const d = JSON.parse(a.ref_data_json || '{}');
      return d.round_id === data.round_id && !!d.candidate_no;
    } catch (e) { return false; }
  }).length;
  return String(count + 1).padStart(4, '0');
}

/** ตรวจสอบใบสมัครสอบ: อนุมัติ (มีค่าธรรมเนียม -> รอชำระเงิน, ไม่มีค่าธรรมเนียม -> รอผลสอบทันที) / ไม่อนุมัติ (ระบุเหตุผล) */
function action_reviewExamApplication(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const app = findRowById_('Applications', 'application_id', p.application_id);
  if (!app || EXAM_APPLICANT_TYPES.indexOf(app.applicant_type) === -1) throw new Error('ไม่พบใบสมัครสอบนี้');
  if (app.status !== 'รอตรวจสอบ') throw new Error('ใบสมัครเลขที่ ' + app.application_id + ' ไม่ได้อยู่ในสถานะรอตรวจสอบแล้ว (สถานะปัจจุบัน: ' + app.status + ')');
  const data = JSON.parse(app.ref_data_json || '{}');
  const label = app.applicant_type;

  if (p.decision === 'reject') {
    if (!p.reason) throw new Error('กรุณาระบุเหตุผลที่ไม่อนุมัติ');
    updateRow_('Applications', app._row, { status: 'ไม่อนุมัติ', reject_reason: p.reason, reviewed_by: user.user_id, reviewed_at: nowIso_() });
    auditLog_('Applications', app.application_id, 'exam_application_rejected', app, { reason: p.reason }, user.user_id);
    mailNotify_(data.email, 'ผลการพิจารณาใบสมัคร' + label, 'ใบสมัครเลขที่ ' + app.application_id + ' ไม่ได้รับการอนุมัติ\nเหตุผล: ' + p.reason);
    return { ok: true };
  }

  const feeAmount = app.fee_amount || 0;
  if (feeAmount > 0) {
    updateRow_('Applications', app._row, { status: 'รอชำระเงิน', reviewed_by: user.user_id, reviewed_at: nowIso_() });
    const existingPayment = readAll_('Payments').find(pm => pm.ref_type === 'application' && pm.ref_id === app.application_id);
    if (!existingPayment) {
      appendRow_('Payments', {
        payment_id: newId_('PAY'), ref_type: 'application', ref_id: app.application_id, member_id: data.member_id || '',
        amount: feeAmount, slip_url: '', method: '', status: 'รอชำระ', verified_by: '', verified_at: '', paid_at: ''
      });
    }
    mailNotify_(data.email, label + 'ได้รับการอนุมัติ — กรุณาชำระเงิน',
      'ใบสมัครเลขที่ ' + app.application_id + ' ได้รับการอนุมัติแล้ว กรุณาชำระค่าธรรมเนียม ' + feeAmount + ' บาท ผ่าน Member Portal แล้วแนบสลิปการโอนเงิน');
    auditLog_('Applications', app.application_id, 'exam_application_approved_pending_payment', app, {}, user.user_id);
    return { ok: true, status: 'รอชำระเงิน' };
  }

  // ไม่มีค่าธรรมเนียม (สมัครสอบสัมภาษณ์) -> ผ่านตรวจสอบแล้วรอผลสอบได้เลย ไม่ต้องรอชำระเงิน — กำหนดรหัสประจำตัวสอบตอนนี้เลย
  data.candidate_no = assignExamCandidateNo_(app.applicant_type, data);
  updateRow_('Applications', app._row, { status: 'รอผลสอบ', reviewed_by: user.user_id, reviewed_at: nowIso_(), ref_data_json: JSON.stringify(data) });
  auditLog_('Applications', app.application_id, 'exam_application_approved', app, { candidate_no: data.candidate_no }, user.user_id);
  mailNotify_(data.email, label + 'ได้รับการอนุมัติ', 'ใบสมัครเลขที่ ' + app.application_id + ' ได้รับการอนุมัติแล้ว รอผลสอบ\nรหัสประจำตัวสอบของท่าน: ' + data.candidate_no);
  return { ok: true, status: 'รอผลสอบ' };
}

/** อนุมัติ/ไม่อนุมัติใบสมัครสอบหลายรายการพร้อมกัน (แต่ละรอบมีผู้สมัครหลายร้อยคน) — p.application_ids: array, p.decision, p.reason (ใช้ร่วมกันทุกรายการกรณีไม่อนุมัติ) */
function action_bulkReviewExamApplications(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const ids = p.application_ids || [];
  if (!ids.length) throw new Error('กรุณาเลือกใบสมัครอย่างน้อย 1 รายการ');
  if (p.decision === 'reject' && !p.reason) throw new Error('กรุณาระบุเหตุผลที่ไม่อนุมัติ (ใช้ร่วมกันกับทุกรายการที่เลือก)');
  const results = ids.map(function (id) {
    try {
      const r = action_reviewExamApplication({ application_id: id, decision: p.decision, reason: p.reason }, user);
      return { application_id: id, ok: true, status: r.status || '' };
    } catch (e) {
      return { application_id: id, ok: false, error: e.message };
    }
  });
  return { results: results, successCount: results.filter(function (r) { return r.ok; }).length, failCount: results.filter(function (r) { return !r.ok; }).length };
}

/** แอดมินยืนยัน/ปฏิเสธการชำระเงินค่าสมัครสอบข้อเขียน -> ยืนยันแล้วเข้าสถานะรอผลสอบ */
function action_verifyExamPayment(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const payment = findRowById_('Payments', 'payment_id', p.payment_id);
  if (!payment || payment.ref_type !== 'application') throw new Error('ไม่พบรายการชำระเงิน');
  const app = findRowById_('Applications', 'application_id', payment.ref_id);
  if (!app || EXAM_APPLICANT_TYPES.indexOf(app.applicant_type) === -1) throw new Error('ไม่พบใบสมัครสอบที่เกี่ยวข้อง');

  if (p.decision === 'reject') {
    updateRow_('Payments', payment._row, { status: 'ปฏิเสธ', verified_by: user.user_id, verified_at: nowIso_() });
    auditLog_('Payments', payment.payment_id, 'payment_rejected', payment, { reason: p.reason }, user.user_id);
    return { ok: true };
  }

  updateRow_('Payments', payment._row, { status: 'ชำระแล้ว', verified_by: user.user_id, verified_at: nowIso_(), paid_at: nowIso_() });
  try { issueReceiptForPayment_(payment.payment_id, user.user_id); } catch (e) { /* ไม่ให้ล้มเหลวทั้งฟังก์ชันถ้าออกใบเสร็จไม่สำเร็จ */ }
  const data = JSON.parse(app.ref_data_json || '{}');
  data.candidate_no = assignExamCandidateNo_(app.applicant_type, data); // กำหนดรหัสประจำตัวสอบตอนชำระเงินครบแล้ว
  updateRow_('Applications', app._row, { status: 'รอผลสอบ', ref_data_json: JSON.stringify(data) });
  auditLog_('Applications', app.application_id, 'exam_payment_confirmed', app, { candidate_no: data.candidate_no }, user.user_id);
  mailNotify_(data.email, 'ยืนยันการชำระเงินค่าสมัครสอบแล้ว', 'ใบสมัครเลขที่ ' + app.application_id + ' ชำระเงินเรียบร้อย รอผลสอบ\nรหัสประจำตัวสอบของท่าน: ' + data.candidate_no);
  return { ok: true };
}

/** ผู้สมัครสอบข้อเขียนที่ชำระเงินแล้วและยังไม่ประกาศผล — หา payment ที่ชำระแล้วของใบสมัคร */
function findPaidExamPayment_(applicationId) {
  return readAll_('Payments').find(pm => pm.ref_type === 'application' && pm.ref_id === applicationId && pm.status === 'ชำระแล้ว');
}

/** แอดมินขอคืนเงินค่าสมัครสอบข้อเขียน (เฉพาะกรณีพิเศษที่สภาฯ กำหนด) — สร้างคำขอในระบบคืนเงินเดิม (Refunds) ให้ไปอนุมัติ/ยืนยันที่หน้า "คืนเงิน"
 * p.application_id, p.reason (บังคับ — ระบุเหตุผล/เลขที่หนังสืออนุมัติของสภาฯ), p.refund_amount (ไม่บังคับ ค่าเริ่มต้นคืนเต็มจำนวน) */
function action_requestExamRefund(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const app = findRowById_('Applications', 'application_id', p.application_id);
  if (!app || app.applicant_type !== 'สมัครสอบข้อเขียน') throw new Error('ไม่พบใบสมัครสอบข้อเขียนนี้');
  if (!p.reason || !String(p.reason).trim()) throw new Error('กรุณาระบุเหตุผล/เลขที่หนังสืออนุมัติกรณีพิเศษของสภาฯ');
  if (app.status !== 'รอผลสอบ') throw new Error('ขอคืนเงินได้เฉพาะใบสมัครที่ชำระเงินแล้วและยังไม่ประกาศผลสอบ (สถานะปัจจุบัน: ' + app.status + ')');
  const payment = findPaidExamPayment_(app.application_id);
  if (!payment) throw new Error('ไม่พบรายการชำระเงินที่ชำระแล้วของใบสมัครนี้');
  const amount = p.refund_amount ? Number(p.refund_amount) : Number(payment.amount);
  if (!(amount > 0) || amount > Number(payment.amount)) throw new Error('จำนวนเงินคืนไม่ถูกต้อง (ต้องมากกว่า 0 และไม่เกิน ' + payment.amount + ' บาท)');
  const r = action_submitRefundRequest({
    payment_id: payment.payment_id, refund_amount: amount,
    reason: '[คืนเงินค่าสมัครสอบข้อเขียน ใบสมัคร ' + app.application_id + '] ' + p.reason
  }, user);
  auditLog_('Applications', app.application_id, 'exam_refund_requested', app, { refund_id: r.refund_id, amount: amount, reason: p.reason }, user.user_id);
  return r;
}

/** แอดมินเปลี่ยนรอบสอบข้อเขียนให้ผู้ที่ชำระเงินแล้ว (เฉพาะกรณีพิเศษที่สภาฯ กำหนด) — ย้ายไปรอบใหม่ ออกรหัสประจำตัวสอบใหม่ และเก็บประวัติการเปลี่ยนรอบไว้
 * p.application_id, p.new_round_id, p.reason (บังคับ) */
function action_changeExamRound(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const app = findRowById_('Applications', 'application_id', p.application_id);
  if (!app || app.applicant_type !== 'สมัครสอบข้อเขียน') throw new Error('ไม่พบใบสมัครสอบข้อเขียนนี้');
  if (!p.reason || !String(p.reason).trim()) throw new Error('กรุณาระบุเหตุผล/เลขที่หนังสืออนุมัติกรณีพิเศษของสภาฯ');
  if (app.status !== 'รอผลสอบ') throw new Error('เปลี่ยนรอบได้เฉพาะใบสมัครที่ชำระเงินแล้วและยังไม่ประกาศผลสอบ (สถานะปัจจุบัน: ' + app.status + ')');
  const data = JSON.parse(app.ref_data_json || '{}');
  const newRound = findRowById_('ExamRounds', 'round_id', p.new_round_id);
  if (!newRound || newRound.exam_type !== 'เขียน') throw new Error('ไม่พบรอบสอบข้อเขียนที่เลือก');
  if (newRound.round_id === data.round_id) throw new Error('รอบที่เลือกเป็นรอบเดิมของผู้สมัครอยู่แล้ว');
  if (newRound.status === 'เสร็จสิ้น') throw new Error('รอบสอบที่เลือกเสร็จสิ้นแล้ว ไม่สามารถย้ายเข้าได้');
  const payment = findPaidExamPayment_(app.application_id);
  if (!payment) throw new Error('ไม่พบรายการชำระเงินที่ชำระแล้วของใบสมัครนี้');
  if (readAll_('Refunds').some(r => r.payment_id === payment.payment_id && ['รอตรวจสอบ', 'อนุมัติ'].indexOf(r.status) !== -1)) {
    throw new Error('ใบสมัครนี้มีคำขอคืนเงินค้างอยู่ ไม่สามารถเปลี่ยนรอบได้');
  }
  const clash = readAll_('Applications').find(a => a.application_id !== app.application_id && a.applicant_type === app.applicant_type &&
    a.applicant_user_id === app.applicant_user_id && ['รอตรวจสอบ', 'รอชำระเงิน', 'รอผลสอบ'].indexOf(a.status) !== -1 &&
    JSON.parse(a.ref_data_json || '{}').round_id === newRound.round_id);
  if (clash) throw new Error('ผู้สมัครมีใบสมัครสอบในรอบใหม่นี้อยู่แล้ว (เลขที่ ' + clash.application_id + ')');

  const history = data.round_changes || [];
  history.push({
    from_round_id: data.round_id, from_round_label: data.round_label, from_candidate_no: data.candidate_no || '',
    to_round_id: newRound.round_id, to_round_label: newRound.round_label,
    reason: p.reason, changed_by: user.user_id, changed_at: nowIso_()
  });
  data.round_changes = history;
  data.round_id = newRound.round_id;
  data.round_label = newRound.round_label;
  data.candidate_no = '';
  data.candidate_no = assignExamCandidateNo_(app.applicant_type, data); // ออกรหัสประจำตัวสอบใหม่ตามลำดับของรอบใหม่
  updateRow_('Applications', app._row, { ref_data_json: JSON.stringify(data) });
  auditLog_('Applications', app.application_id, 'exam_round_changed', app, { to_round: newRound.round_label, reason: p.reason, candidate_no: data.candidate_no }, user.user_id);
  mailNotify_(data.email, 'เปลี่ยนรอบสอบข้อเขียน',
    'ใบสมัครเลขที่ ' + app.application_id + ' ถูกเปลี่ยนเป็นรอบสอบ ' + newRound.round_label + ' (กรณีพิเศษตามที่สภาฯ กำหนด)\nรหัสประจำตัวสอบใหม่: ' + data.candidate_no);
  return { ok: true, round_label: newRound.round_label, candidate_no: data.candidate_no };
}

// ----- 6.4 คะแนน/ผลสอบ: เกณฑ์ผ่าน การบันทึกทีละคน/นำเข้าหลายคน การดู/พิมพ์ใบรับรอง -----

// เกณฑ์ผ่านการสอบข้อเขียน (ประกาศสภาฯ — ข้อมูลการสอบครั้งที่ 1/2569)
//  (1) ได้คะแนนแต่ละวิชาไม่น้อยกว่าร้อยละ 60 หรือ
//  (2) คะแนนเฉลี่ยรวม 2 วิชาไม่น้อยกว่าร้อยละ 65 และคะแนนรายวิชาแต่ละวิชาไม่น้อยกว่าร้อยละ 50
const EXAM_PASS_PER_SUBJECT = 60;
const EXAM_PASS_AVERAGE = 65;
const EXAM_PASS_AVERAGE_MIN_EACH = 50;

/** คำนวณผลสอบข้อเขียนจากคะแนน (ร้อยละ) — subjects: วิชาที่สมัครสอบ (ชื่อเต็มใน EXAM_SUBJECTS), scores: { [ชื่อวิชา]: ตัวเลข }
 * คืน { subject_results, result: 'ผ่าน'|'ไม่ผ่าน', basis } */
function computeWrittenExamResult_(subjects, scores) {
  const subjectResults = {};
  subjects.forEach(sub => { subjectResults[sub] = scores[sub] >= EXAM_PASS_PER_SUBJECT ? 'ผ่าน' : 'ไม่ผ่าน'; });
  if (subjects.length === 1) {
    const ok = subjectResults[subjects[0]] === 'ผ่าน';
    return { subject_results: subjectResults, result: ok ? 'ผ่าน' : 'ไม่ผ่าน', basis: ok ? 'คะแนนไม่น้อยกว่าร้อยละ ' + EXAM_PASS_PER_SUBJECT : 'คะแนนต่ำกว่าร้อยละ ' + EXAM_PASS_PER_SUBJECT };
  }
  const vals = subjects.map(sub => scores[sub]);
  if (vals.every(v => v >= EXAM_PASS_PER_SUBJECT)) {
    return { subject_results: subjectResults, result: 'ผ่าน', basis: 'ได้คะแนนทั้ง 2 วิชาไม่น้อยกว่าร้อยละ ' + EXAM_PASS_PER_SUBJECT };
  }
  const avg = vals.reduce((x, y) => x + y, 0) / vals.length;
  if (avg >= EXAM_PASS_AVERAGE && vals.every(v => v >= EXAM_PASS_AVERAGE_MIN_EACH)) {
    return { subject_results: subjectResults, result: 'ผ่าน', basis: 'คะแนนเฉลี่ยรวม 2 วิชา ' + (Math.round(avg * 100) / 100) + ' (ไม่น้อยกว่าร้อยละ ' + EXAM_PASS_AVERAGE + ') และแต่ละวิชาไม่น้อยกว่าร้อยละ ' + EXAM_PASS_AVERAGE_MIN_EACH };
  }
  return { subject_results: subjectResults, result: 'ไม่ผ่าน', basis: 'ไม่เข้าเกณฑ์ทั้งข้อ (1) และข้อ (2)' };
}

function parseScore_(v) {
  if (v === undefined || v === null || String(v).trim() === '') return null;
  const n = Number(String(v).replace(/,/g, '').trim());
  if (isNaN(n) || n < 0 || n > 100) return NaN;
  return n;
}

/** ตรวจ+คำนวณผลสอบของใบสมัครหนึ่งใบ (ยังไม่เขียนลงชีต) — input: { tp, pp, result, exam_date, seat_no, score, note }
 * คืน { exam_result, status } หรือ throw Error ถ้าข้อมูลไม่ถูกต้อง */
function buildExamResult_(app, data, input, round) {
  const examDate = input.exam_date || (round && round.exam_date) || '';
  const base = { exam_date: examDate, seat_no: input.seat_no || data.candidate_no || '', recorded_at: nowIso_() };
  const manual = input.result === 'ผ่าน' || input.result === 'ไม่ผ่าน' ? input.result : '';
  if (input.result && !manual) throw new Error('ผลสอบต้องเป็น "ผ่าน" หรือ "ไม่ผ่าน"');
  if (app.applicant_type === 'สมัครสอบข้อเขียน') {
    const subjects = (data.subjects && data.subjects.length) ? data.subjects : EXAM_SUBJECTS;
    const scores = {};
    const keyOf = sub => (sub === EXAM_SUBJECTS[0] ? 'tp' : 'pp');
    subjects.forEach(sub => {
      const sc = parseScore_(input[keyOf(sub)]);
      if (sc === null) throw new Error('ยังไม่ได้กรอกคะแนนวิชา "' + sub + '"');
      if (isNaN(sc)) throw new Error('คะแนนวิชา "' + sub + '" ต้องเป็นตัวเลข 0-100');
      scores[sub] = sc;
    });
    const calc = computeWrittenExamResult_(subjects, scores);
    const result = manual || calc.result;
    return {
      status: result === 'ผ่าน' ? 'สอบผ่าน' : 'สอบไม่ผ่าน',
      subject_results: calc.subject_results,
      exam_result: Object.assign(base, { kind: 'เขียน', scores: scores, subject_results: calc.subject_results, result: result,
        basis: manual && manual !== calc.result ? 'สภาฯ กำหนดผลสอบ (ปรับจากผลคำนวณอัตโนมัติ: ' + calc.result + ')' : calc.basis })
    };
  }
  if (!manual) throw new Error('กรุณาระบุผลสอบสัมภาษณ์ (ผ่าน/ไม่ผ่าน)');
  const sc = parseScore_(input.score);
  if (isNaN(sc)) throw new Error('คะแนนสัมภาษณ์ต้องเป็นตัวเลข 0-100');
  return {
    status: manual === 'ผ่าน' ? 'สอบผ่าน' : 'สอบไม่ผ่าน', subject_results: null,
    exam_result: Object.assign(base, { kind: 'สัมภาษณ์', score: sc, result: manual, note: input.note || '' })
  };
}

function writeExamResult_(app, data, built, user) {
  data.exam_result = Object.assign({}, built.exam_result, { recorded_by: user.user_id });
  if (built.subject_results) data.subject_results = built.subject_results;
  updateRow_('Applications', app._row, { status: built.status, ref_data_json: JSON.stringify(data), reviewed_at: nowIso_() });
  auditLog_('Applications', app.application_id, 'exam_result_recorded', { status: app.status }, { status: built.status, exam_result: data.exam_result }, user.user_id);
  const r = data.exam_result;
  const isWritten = app.applicant_type === 'สมัครสอบข้อเขียน';
  let lines = ['เรียน ' + (data.name || 'ผู้สมัครสอบ'), '', 'สภาฯ ประกาศผลสอบ' + (isWritten ? 'ข้อเขียน' : 'สัมภาษณ์') + 'ของท่านแล้ว', ''];
  lines.push('รอบสอบ: ' + (data.round_label || '-'));
  if (r.exam_date) lines.push('วันที่สอบ: ' + fmtThaiDate_(r.exam_date));
  if (r.seat_no) lines.push('รหัสประจำตัวสอบ: ' + r.seat_no);
  if (isWritten) Object.keys(r.scores).forEach(sub => lines.push(sub + ': ' + r.scores[sub] + ' (' + r.subject_results[sub] + ')'));
  else if (r.score !== null && r.score !== undefined && !isNaN(r.score)) lines.push('คะแนนสัมภาษณ์: ' + r.score);
  lines.push('ผลการสอบ: ' + r.result);
  lines.push('', 'ท่านสามารถดูคะแนนและพิมพ์ใบรับรองผลสอบได้ที่เมนู "สมัครสอบข้อเขียน/สัมภาษณ์" ในพื้นที่สมาชิก');
  const base = (PropertiesService.getScriptProperties().getProperty('FRONTEND_BASE_URL') || '');
  mailNotify_(data.email, 'ประกาศผลสอบ' + (isWritten ? 'ข้อเขียน' : 'สัมภาษณ์') + ' — ' + r.result, lines.join('\n'),
    base ? { buttonText: 'ดูผลสอบและพิมพ์ใบรับรอง', buttonUrl: base.replace(/\/?$/, '/') + 'member-exam-apply.html', tone: r.result === 'ผ่าน' ? 'success' : 'danger' } : { tone: r.result === 'ผ่าน' ? 'success' : 'danger' });
}

function fmtThaiDate_(iso) {
  if (!iso) return '-';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return String(iso);
  const dd = ('0' + d.getDate()).slice(-2), mm = ('0' + (d.getMonth() + 1)).slice(-2);
  return dd + '/' + mm + '/' + (d.getFullYear() + 543);
}

/** บันทึกผลสอบทีละคน (หลังวันสอบ) — สอบข้อเขียน: p.tp, p.pp (คะแนนร้อยละ) ระบบคำนวณผลตามเกณฑ์ (p.result ใช้ปรับเฉพาะกรณีสภาฯ กำหนด)
 * สอบสัมภาษณ์: p.result 'ผ่าน'|'ไม่ผ่าน' (+ p.score, p.note ไม่บังคับ) */
function action_recordExamResult(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const app = findRowById_('Applications', 'application_id', p.application_id);
  if (!app || EXAM_APPLICANT_TYPES.indexOf(app.applicant_type) === -1) throw new Error('ไม่พบใบสมัครสอบนี้');
  if (app.status !== 'รอผลสอบ' && !p.overwrite) throw new Error('ใบสมัครนี้ไม่ได้อยู่ในสถานะรอผลสอบ');
  if (['รอผลสอบ', 'สอบผ่าน', 'สอบไม่ผ่าน'].indexOf(app.status) === -1) throw new Error('ใบสมัครนี้ไม่ได้อยู่ในสถานะที่บันทึกผลสอบได้');
  const data = JSON.parse(app.ref_data_json || '{}');
  const round = findRowById_('ExamRounds', 'round_id', data.round_id);
  const built = buildExamResult_(app, data, p, round);
  writeExamResult_(app, data, built, user);
  return { ok: true, status: built.status, exam_result: data.exam_result };
}

/** นำเข้าผลสอบทีละหลายคน (TOR: ผลสอบข้อเขียน/สัมภาษณ์) — p.exam_type 'เขียน'|'สัมภาษณ์', p.round_id, p.rows[], p.dry_run (ตรวจอย่างเดียว ไม่บันทึก), p.overwrite
 * แต่ละแถว: { candidate_no | national_id | application_id, tp, pp, result, exam_date, seat_no } (สัมภาษณ์: { ..., result, score, note })
 * คืนผลรายแถว {row, ok, message, name, result, ...} เพื่อให้แอดมินตรวจก่อนยืนยันบันทึกจริง */
function action_importExamResults(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const applicantType = p.exam_type === 'เขียน' ? 'สมัครสอบข้อเขียน' : (p.exam_type === 'สัมภาษณ์' ? 'สมัครสอบสัมภาษณ์' : '');
  if (!applicantType) throw new Error('ประเภทการสอบไม่ถูกต้อง');
  if (!p.round_id) throw new Error('กรุณาเลือกรอบสอบ');
  const rows = p.rows || [];
  if (!rows.length) throw new Error('ไม่พบข้อมูลที่จะนำเข้า');
  if (rows.length > 500) throw new Error('นำเข้าได้ครั้งละไม่เกิน 500 แถว');
  const round = findRowById_('ExamRounds', 'round_id', p.round_id);
  if (!round) throw new Error('ไม่พบรอบสอบ');
  const members = readAll_('Members');
  const apps = readAll_('Applications').filter(a => a.applicant_type === applicantType && ['รอผลสอบ', 'สอบผ่าน', 'สอบไม่ผ่าน'].indexOf(a.status) !== -1)
    .map(a => { const d = JSON.parse(a.ref_data_json || '{}'); return { app: a, data: d }; })
    .filter(x => x.data.round_id === p.round_id);
  const nidOf = x => { const m = members.find(mm => mm.member_id === x.data.member_id); return m ? String(m.national_id || '') : ''; };

  const seen = {};
  const results = [];
  let okCount = 0;
  rows.forEach((row, idx) => {
    const out = { row: idx + 1, ok: false, message: '' };
    try {
      const cand = String(row.candidate_no || '').trim();
      const nid = String(row.national_id || '').replace(/\D/g, '');
      const aid = String(row.application_id || '').trim();
      if (!cand && !nid && !aid) throw new Error('ไม่มีรหัสประจำตัวสอบ/เลขบัตรประชาชน/เลขที่ใบสมัคร');
      let matches = apps.filter(x => (aid && x.app.application_id === aid) || (!aid && cand && String(x.data.candidate_no) === cand) || (!aid && !cand && nid && nidOf(x) === nid));
      if (!matches.length) throw new Error('ไม่พบผู้สมัครในรอบสอบนี้ที่ตรงกับข้อมูลที่ระบุ (ต้องเป็นผู้ที่อนุมัติ/ชำระเงินแล้ว)');
      if (matches.length > 1) throw new Error('พบผู้สมัครมากกว่า 1 รายที่ตรงกับข้อมูลที่ระบุ');
      const x = matches[0];
      out.name = x.data.name || '';
      out.candidate_no = x.data.candidate_no || '';
      if (seen[x.app.application_id]) throw new Error('ซ้ำกับแถวที่ ' + seen[x.app.application_id] + ' ในไฟล์เดียวกัน');
      seen[x.app.application_id] = idx + 1;
      if (x.app.status !== 'รอผลสอบ' && !p.overwrite) throw new Error('บันทึกผลสอบไว้แล้ว (' + x.app.status + ') — ติ๊ก "ทับผลเดิม" หากต้องการแก้ไข');
      const built = buildExamResult_(x.app, x.data, row, round);
      out.result = built.exam_result.result;
      out.basis = built.exam_result.basis || '';
      out.scores = built.exam_result.scores || (built.exam_result.score !== undefined ? { สัมภาษณ์: built.exam_result.score } : null);
      if (!p.dry_run) writeExamResult_(x.app, x.data, built, user);
      out.ok = true; okCount++;
      out.message = p.dry_run ? 'พร้อมบันทึก' : 'บันทึกแล้ว';
    } catch (err) {
      out.message = err.message;
    }
    results.push(out);
  });
  if (!p.dry_run) auditLog_('ExamRounds', p.round_id, 'exam_results_imported', {}, { type: applicantType, total: rows.length, ok: okCount }, user.user_id);
  return { dry_run: !!p.dry_run, total: rows.length, ok_count: okCount, error_count: rows.length - okCount, results: results };
}

/** รายการที่บันทึกผลสอบแล้ว (แอดมิน) — p.exam_type, p.round_id (ไม่บังคับ) */
function action_listExamResults(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const applicantType = p.exam_type === 'สัมภาษณ์' ? 'สมัครสอบสัมภาษณ์' : 'สมัครสอบข้อเขียน';
  const members = readAll_('Members');
  const out = [];
  readAll_('Applications').forEach(a => {
    if (a.applicant_type !== applicantType || ['สอบผ่าน', 'สอบไม่ผ่าน'].indexOf(a.status) === -1) return;
    const d = JSON.parse(a.ref_data_json || '{}');
    if (p.round_id && d.round_id !== p.round_id) return;
    const m = members.find(mm => mm.member_id === d.member_id);
    out.push({ application_id: a.application_id, name: d.name || '', national_id: m ? m.national_id : '', round_label: d.round_label || '',
      candidate_no: d.candidate_no || '', status: a.status, exam_result: d.exam_result || null, subject_results: d.subject_results || null });
  });
  return out.sort((x, y) => String(x.candidate_no).localeCompare(String(y.candidate_no)));
}

/** ข้อมูลใบรับรองผลสอบของใบสมัครตนเอง (สมาชิก) — ใช้พิมพ์ใบรับรอง */
function action_getExamCertificate(p, user) {
  requireRole_(user, ['member']);
  const app = findRowById_('Applications', 'application_id', p.application_id);
  if (!app || app.applicant_user_id !== user.user_id || EXAM_APPLICANT_TYPES.indexOf(app.applicant_type) === -1) throw new Error('ไม่พบใบสมัครสอบนี้');
  const d = JSON.parse(app.ref_data_json || '{}');
  if (!d.exam_result) throw new Error('ยังไม่มีการประกาศผลสอบของใบสมัครนี้');
  const member = findRowById_('Members', 'member_id', d.member_id);
  const officer = getPresidentForDate_(d.exam_result.recorded_at || nowIso_());
  return {
    application_id: app.application_id, kind: d.exam_result.kind, name: d.name || '', national_id: member ? member.national_id : '',
    round_label: d.round_label || '', exam_result: d.exam_result, subject_results: d.subject_results || null,
    signatory: officer ? { name_th: officer.name_th, position_label: officer.position_label, signature_url: officer.signature_url } : null,
    issued_at: nowIso_()
  };
}

/** ค้นหาผลสอบข้อเขียนจากเลขบัตรประชาชน — PUBLIC (ใช้กับหน้า exam-result.html และเชื่อมกับเว็บไซต์หลักของสภาฯ ได้)
 * ซ่อนนามสกุลบางส่วนเพื่อคุ้มครองข้อมูลส่วนบุคคล คืนเฉพาะผลที่ประกาศแล้ว */
function action_publicExamResult(p) {
  const nid = String(p.national_id || '').replace(/\D/g, '');
  if (nid.length !== 13) throw new Error('กรุณากรอกเลขบัตรประจำตัวประชาชน 13 หลัก');
  const members = readAll_('Members').filter(m => String(m.national_id) === nid);
  if (!members.length) return { found: false, results: [] };
  const ids = members.map(m => m.member_id);
  const results = [];
  readAll_('Applications').forEach(a => {
    if (a.applicant_type !== 'สมัครสอบข้อเขียน' || ['สอบผ่าน', 'สอบไม่ผ่าน'].indexOf(a.status) === -1) return;
    const d = JSON.parse(a.ref_data_json || '{}');
    if (ids.indexOf(d.member_id) === -1 || !d.exam_result) return;
    const m = members.find(mm => mm.member_id === d.member_id);
    const last = String(m.last_name_th || '');
    results.push({
      name: [m.prefix_th, m.first_name_th, last.slice(0, 2) + '***'].filter(Boolean).join(' '),
      exam_date: d.exam_result.exam_date, seat_no: d.exam_result.seat_no, scores: d.exam_result.scores, subject_results: d.exam_result.subject_results,
      result: d.exam_result.result, round_label: d.round_label || ''
    });
  });
  results.sort((x, y) => new Date(y.exam_date) - new Date(x.exam_date));
  return { found: true, subjects: EXAM_SUBJECTS, results: results };
}

// ---------- 6.5 ระบบจัดการสอบสัมภาษณ์ (แอดมิน) + หน้าให้คะแนนของกรรมการ (URL แยก ไม่ต้องเข้าระบบเจ้าหน้าที่) ----------
// โครงสร้างข้อมูล (ใช้ร่วมกับรอบสอบ ExamRounds และใบสมัครสอบสัมภาษณ์ใน Applications เดิม):
//   InterviewSessions   1 แถวต่อรอบสอบสัมภาษณ์ — เวลาสอบ, เกณฑ์ผ่าน (คะแนนเฉลี่ย), สถานะเปิด/ปิดรับคะแนน
//   InterviewGroups     กลุ่มสอบ — กรรมการ 2 ท่านต่อกลุ่ม + รหัสผ่านของกลุ่ม (เก็บแบบ hash; เว้นว่างได้ = ไม่ต้องใช้รหัสผ่าน)
//   InterviewCandidates ผู้สมัครที่จัดเข้ากลุ่ม/เวลาแล้ว + คะแนนของกรรมการแต่ละท่าน (scores_json / remarks_json)
// กรรมการเข้าใช้งานที่หน้า interviewer.html: เลือกรอบ → เลือกกลุ่ม → เลือกกรรมการ 1/2 → ใส่รหัสผ่านกลุ่ม (เหมือนระบบให้คะแนนเดิม) — ไม่มีบัญชีผู้ใช้/ไม่เข้าระบบเจ้าหน้าที่ — token เก็บใน CacheService อายุ 6 ชม.
// กรรมการเห็นเฉพาะผู้เข้าสอบกลุ่มตนเองและ "คะแนนของตนเอง" เท่านั้น (ไม่เห็นคะแนนของกรรมการอีกท่าน เพื่อให้การให้คะแนนเป็นอิสระต่อกัน)
// เมื่อกรรมการให้คะแนนครบ แอดมินกด "บันทึกผลสอบ" เพื่อส่งผลเข้าระบบผลสอบเดิม (สอบผ่าน/สอบไม่ผ่าน + อีเมลแจ้งผู้สมัคร + ใบรับรองผลสอบ)

const INTERVIEW_CRITERIA = [
  { key: 'skill', label: 'ทักษะทางสังคมสงเคราะห์ในการปฏิบัติงาน', max: 30, hint: 'ประเมินจากทักษะต่างๆ 13 ด้าน เช่น การสร้างสัมพันธภาพ การสื่อสาร การค้นหาข้อเท็จจริง การเยี่ยมบ้าน การจัดการรายกรณี และการแก้ปัญหา เป็นต้น' },
  { key: 'attitude', label: 'ทัศนคติทางสังคมสงเคราะห์', max: 30, hint: 'ประเมินจาก 8 ด้าน เช่น การไม่ตำหนิหรือตัดสิน การเคารพความหลากหลาย ความเห็นอกเห็นใจ ความเป็นธรรมทางสังคม และการคำนึงถึงประโยชน์สูงสุดของผู้ใช้บริการ' },
  { key: 'writing', label: 'การเขียน', max: 20, hint: 'ประเมินจากการเขียนรายงานที่อธิบายเนื้อหาสอดคล้อง สมเหตุสมผล เรียบเรียงเข้าใจง่าย ภาษาเหมาะสม มีส่วนนำ เนื้อหา และสรุปชัดเจน รวมถึงมีการอ้างอิงตามระบบวิชาการ' },
  { key: 'presentation', label: 'การนำเสนอ', max: 20, hint: 'ประเมินจากการอธิบายและวิเคราะห์ความเชื่อมโยงของเนื้อหาได้ชัดเจนและตรงประเด็น' }
];
const INTERVIEW_DEFAULT_PASS_SCORE = 60; // คะแนนเฉลี่ยของกรรมการ 2 ท่าน (เต็ม 100) ไม่น้อยกว่านี้จึงผ่าน — แก้ไขได้ต่อรอบ
const INTERVIEW_TOKEN_TTL_SEC = 21600; // 6 ชั่วโมง (ค่าสูงสุดของ CacheService)
const INTERVIEW_SLOTS = ['inv1', 'inv2'];
const INTERVIEW_ELIGIBLE_STATUSES = ['รอผลสอบ', 'สอบผ่าน', 'สอบไม่ผ่าน']; // สถานะใบสมัครที่จัดเข้ากลุ่มสอบได้ (ผ่านการอนุมัติแล้ว)

function interviewJson_(s, fallback) {
  try { const v = JSON.parse(s); return (v === null || v === undefined) ? fallback : v; } catch (e) { return fallback; }
}

function interviewWithLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { try { lock.releaseLock(); } catch (e) {} }
}

function interviewGetRound_(roundId) {
  const round = findRowById_('ExamRounds', 'round_id', roundId);
  if (!round || round.exam_type !== 'สัมภาษณ์') throw new Error('ไม่พบรอบสอบสัมภาษณ์นี้');
  return round;
}

/** ดึง (หรือสร้างถ้า create=true) แถวตั้งค่ารอบสอบสัมภาษณ์ */
function interviewGetSession_(roundId, create, actor) {
  let s = readAll_('InterviewSessions').find(x => x.round_id === roundId);
  if (!s && create) {
    interviewGetRound_(roundId);
    appendRow_('InterviewSessions', {
      session_id: newId_('IVS'), round_id: roundId, time_text: '', pass_score: INTERVIEW_DEFAULT_PASS_SCORE,
      scoring_closed: false, results_applied_at: '', note: '', created_by: actor || '', created_at: nowIso_(), updated_at: nowIso_()
    });
    s = readAll_('InterviewSessions').find(x => x.round_id === roundId);
  }
  return s || null;
}

function interviewPassScore_(session) {
  const n = session ? Number(session.pass_score) : NaN;
  return isNaN(n) || n <= 0 ? INTERVIEW_DEFAULT_PASS_SCORE : n;
}

/** ผู้สมัครสอบสัมภาษณ์ของรอบนี้ที่พร้อมเข้าสอบ (อนุมัติแล้ว) — เรียงตามรหัสประจำตัวสอบ */
function interviewEligibleApps_(roundId) {
  const out = [];
  readAll_('Applications').forEach(a => {
    if (a.applicant_type !== 'สมัครสอบสัมภาษณ์' || INTERVIEW_ELIGIBLE_STATUSES.indexOf(a.status) === -1) return;
    const d = interviewJson_(a.ref_data_json, {});
    if (d.round_id !== roundId) return;
    out.push({ app: a, data: d });
  });
  return out.sort((x, y) => String(x.data.candidate_no || '').localeCompare(String(y.data.candidate_no || '')));
}

function interviewTotals_(ic, passScore) {
  const sc = interviewJson_(ic.scores_json, {});
  const s1 = sc.inv1 && sc.inv1.total !== undefined ? Number(sc.inv1.total) : null;
  const s2 = sc.inv2 && sc.inv2.total !== undefined ? Number(sc.inv2.total) : null;
  let average = null, status = 'รอประเมิน';
  if (isTrue_(ic.absent)) status = 'ไม่เข้าสอบ';
  else if (s1 !== null && s2 !== null) {
    average = Number(((s1 + s2) / 2).toFixed(2));
    status = average >= passScore ? 'ผ่าน' : 'ไม่ผ่าน';
  } else if (s1 !== null || s2 !== null) status = 'ประเมินแล้ว 1 ท่าน';
  return { s1: s1, s2: s2, average: average, status: status, scores: sc, remarks: interviewJson_(ic.remarks_json, {}) };
}

function interviewValidateScores_(input) {
  if (!input || typeof input !== 'object') throw new Error('ไม่พบข้อมูลคะแนน');
  const clean = {};
  let total = 0;
  INTERVIEW_CRITERIA.forEach(c => {
    const raw = input[c.key];
    if (raw === undefined || raw === null || String(raw).trim() === '') throw new Error('กรุณากรอกคะแนนหัวข้อ "' + c.label + '"');
    const n = Number(raw);
    if (!isFinite(n) || Math.floor(n) !== n) throw new Error('คะแนนหัวข้อ "' + c.label + '" ต้องเป็นจำนวนเต็ม');
    if (n < 0 || n > c.max) throw new Error('คะแนนหัวข้อ "' + c.label + '" ต้องอยู่ระหว่าง 0-' + c.max);
    clean[c.key] = n;
    total += n;
  });
  clean.total = total;
  clean.saved_at = nowIso_();
  return clean;
}

function interviewPad2_(n) { return (n < 10 ? '0' : '') + n; }
function interviewClock_(mins) { return interviewPad2_(Math.floor(mins / 60) % 24) + '.' + interviewPad2_(mins % 60); }
/** ช่วงเวลาสอบรายคนรูปแบบเดียวกับระบบเดิม เช่น "09.30 - 10.00" */
function interviewTimeRange_(startMin, minutes) { return interviewClock_(startMin) + ' - ' + interviewClock_(startMin + minutes); }

function interviewCandidateView_(ic, passScore, withRemarks) {
  const t = interviewTotals_(ic, passScore);
  const view = {
    ic_id: ic.ic_id, application_id: ic.application_id, candidate_no: ic.candidate_no || '', name: ic.candidate_name || '',
    group_id: ic.group_id, order_no: Number(ic.order_no) || 0, time_slot: ic.time_slot || '', absent: isTrue_(ic.absent),
    inv1: t.scores.inv1 || null, inv2: t.scores.inv2 || null, s1: t.s1, s2: t.s2, average: t.average, status: t.status
  };
  if (withRemarks) view.remarks = t.remarks;
  return view;
}

function interviewSortCandidates_(a, b) {
  return (Number(a.order_no) || 0) - (Number(b.order_no) || 0) || String(a.candidate_no).localeCompare(String(b.candidate_no));
}

// ----- แอดมิน -----

/** ข้อมูลทั้งหมดของรอบสอบสัมภาษณ์หนึ่งรอบ: ตั้งค่ารอบ, กลุ่ม+กรรมการ+รหัสเข้าใช้, ผู้เข้าสอบ+คะแนน, สถิติ */
function action_getInterviewSession(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const round = interviewGetRound_(p.round_id);
  const session = interviewGetSession_(p.round_id, false);
  const passScore = interviewPassScore_(session);
  const eligible = interviewEligibleApps_(p.round_id);
  const eligMap = {};
  eligible.forEach(x => { eligMap[x.app.application_id] = x; });
  const groups = readAll_('InterviewGroups').filter(g => g.round_id === p.round_id).sort((a, b) => Number(a.group_no) - Number(b.group_no));
  const ics = readAll_('InterviewCandidates').filter(c => c.round_id === p.round_id && eligMap[c.application_id]);
  const assignedIds = {};
  const stats = { total: eligible.length, assigned: 0, unassigned: 0, absent: 0, scored_both: 0, passed: 0, failed: 0, waiting: 0, applied: 0 };
  const groupViews = groups.map(g => ({
    group_id: g.group_id, group_no: Number(g.group_no), time_range: g.time_range || '',
    inv1_name: g.inv1_name || '', inv2_name: g.inv2_name || '', has_password: !!g.password_hash,
    candidates: []
  }));
  const groupById = {};
  groupViews.forEach(g => { groupById[g.group_id] = g; });
  ics.sort(interviewSortCandidates_).forEach(ic => {
    const g = groupById[ic.group_id];
    if (!g) return;
    const v = interviewCandidateView_(ic, passScore, true);
    v.app_status = eligMap[ic.application_id].app.status;
    g.candidates.push(v);
    assignedIds[ic.application_id] = true;
    stats.assigned++;
    if (v.absent) stats.absent++;
    else if (v.average !== null) { stats.scored_both++; if (v.status === 'ผ่าน') stats.passed++; else stats.failed++; }
    else stats.waiting++;
    if (v.app_status === 'สอบผ่าน' || v.app_status === 'สอบไม่ผ่าน') stats.applied++;
  });
  const unassigned = eligible.filter(x => !assignedIds[x.app.application_id]).map(x => ({
    application_id: x.app.application_id, candidate_no: x.data.candidate_no || '', name: x.data.name || '', app_status: x.app.status
  }));
  stats.unassigned = unassigned.length;
  return {
    round: { round_id: round.round_id, round_label: round.round_label, exam_date: round.exam_date || '', location_note: round.location_note || '', status: round.status },
    session: {
      time_text: session ? (session.time_text || '') : '', pass_score: passScore,
      scoring_closed: session ? isTrue_(session.scoring_closed) : false,
      results_applied_at: session ? (session.results_applied_at || '') : '', note: session ? (session.note || '') : ''
    },
    criteria: INTERVIEW_CRITERIA, groups: groupViews, unassigned: unassigned, stats: stats
  };
}

function action_saveInterviewSession(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  return interviewWithLock_(() => {
    const s = interviewGetSession_(p.round_id, true, user.user_id);
    const patch = { updated_at: nowIso_() };
    if (p.time_text !== undefined) patch.time_text = String(p.time_text || '');
    if (p.note !== undefined) patch.note = String(p.note || '');
    if (p.pass_score !== undefined) {
      const n = Number(p.pass_score);
      if (!isFinite(n) || n < 1 || n > 100) throw new Error('เกณฑ์ผ่านต้องเป็นตัวเลข 1-100');
      patch.pass_score = n;
    }
    updateRow_('InterviewSessions', s._row, patch);
    auditLog_('InterviewSessions', s.session_id, 'update', s, patch, user.user_id);
    return { ok: true };
  });
}

function action_setInterviewScoringClosed(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  return interviewWithLock_(() => {
    const s = interviewGetSession_(p.round_id, true, user.user_id);
    updateRow_('InterviewSessions', s._row, { scoring_closed: p.closed === true, updated_at: nowIso_() });
    auditLog_('InterviewSessions', s.session_id, p.closed === true ? 'scoring_closed' : 'scoring_reopened', {}, {}, user.user_id);
    return { ok: true, scoring_closed: p.closed === true };
  });
}

/** เพิ่ม/แก้ไขกลุ่มสอบ — p.password: ตั้ง/เปลี่ยนรหัสผ่านของกลุ่ม (ไม่ส่ง/ว่าง = คงรหัสเดิม), p.clear_password = true เพื่อเลิกใช้รหัสผ่าน
 * เปลี่ยนรหัสผ่านแล้ว เซสชันกรรมการที่เข้าอยู่จะหมดอายุทันที */
function action_saveInterviewGroup(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  return interviewWithLock_(() => {
    interviewGetRound_(p.round_id);
    interviewGetSession_(p.round_id, true, user.user_id);
    const fields = {
      time_range: String(p.time_range || '').trim(), inv1_name: String(p.inv1_name || '').trim(), inv2_name: String(p.inv2_name || '').trim()
    };
    const pw = String(p.password || '');
    if (pw) {
      if (pw.length < 4) throw new Error('รหัสผ่านกลุ่มต้องมีอย่างน้อย 4 ตัวอักษร');
      fields.password_hash = hashPassword_(pw);
    } else if (p.clear_password === true) fields.password_hash = '';
    if (p.group_id) {
      const g = findRowById_('InterviewGroups', 'group_id', p.group_id);
      if (!g || g.round_id !== p.round_id) throw new Error('ไม่พบกลุ่มสอบนี้');
      const patch = Object.assign({ updated_at: nowIso_() }, fields);
      updateRow_('InterviewGroups', g._row, patch);
      auditLog_('InterviewGroups', g.group_id, 'update', { has_password: !!g.password_hash }, { has_password: fields.password_hash === undefined ? !!g.password_hash : !!fields.password_hash, name1: fields.inv1_name, name2: fields.inv2_name }, user.user_id);
      return { ok: true, group_id: g.group_id };
    }
    const existing = readAll_('InterviewGroups').filter(g => g.round_id === p.round_id);
    const nextNo = existing.reduce((m, g) => Math.max(m, Number(g.group_no) || 0), 0) + 1;
    const row = Object.assign({
      group_id: newId_('IVG'), round_id: p.round_id, group_no: nextNo, password_hash: '',
      created_by: user.user_id, created_at: nowIso_(), updated_at: nowIso_()
    }, fields);
    appendRow_('InterviewGroups', row);
    auditLog_('InterviewGroups', row.group_id, 'create', {}, { round_id: p.round_id, group_no: nextNo, has_password: !!row.password_hash }, user.user_id);
    return { ok: true, group_id: row.group_id, group_no: nextNo };
  });
}

/** ลบกลุ่มสอบ — ถ้ามีผู้เข้าสอบที่กรรมการให้คะแนนแล้วต้องยืนยัน (p.force) คะแนนจะถูกลบพร้อมกลุ่ม */
function action_deleteInterviewGroup(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  return interviewWithLock_(() => {
    const g = findRowById_('InterviewGroups', 'group_id', p.group_id);
    if (!g) throw new Error('ไม่พบกลุ่มสอบนี้');
    const ics = readAll_('InterviewCandidates').filter(c => c.group_id === g.group_id);
    const scored = ics.filter(c => { const t = interviewTotals_(c, INTERVIEW_DEFAULT_PASS_SCORE); return t.s1 !== null || t.s2 !== null; });
    if (scored.length && !p.force) throw new Error('กลุ่มนี้มีผู้เข้าสอบที่กรรมการให้คะแนนแล้ว ' + scored.length + ' คน — ยืนยันการลบจะลบคะแนนเหล่านั้นด้วย');
    const rows = ics.map(c => c._row).concat([]).sort((a, b) => b - a);
    rows.forEach(r => deleteRow_('InterviewCandidates', r));
    deleteRow_('InterviewGroups', g._row);
    auditLog_('InterviewGroups', g.group_id, 'delete', g, { removed_candidates: ics.length }, user.user_id);
    return { ok: true, removed_candidates: ics.length };
  });
}

/** จัดผู้เข้าสอบเข้ากลุ่ม/เวลา (หลายคนพร้อมกัน) — items: [{ application_id, group_id ('' = ถอดออกจากกลุ่ม), time_slot, order_no }]
 * การย้ายกลุ่ม/ถอดออก ถ้าผู้เข้าสอบมีคะแนนแล้วต้องใส่ p.force (คะแนนเดิมจะถูกล้าง) */
function action_assignInterviewCandidates(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  if (!p.round_id || !Array.isArray(p.items) || !p.items.length) throw new Error('ไม่พบรายการที่จะจัดกลุ่ม');
  return interviewWithLock_(() => {
    interviewGetRound_(p.round_id);
    const eligMap = {};
    interviewEligibleApps_(p.round_id).forEach(x => { eligMap[x.app.application_id] = x; });
    const groups = {};
    readAll_('InterviewGroups').filter(g => g.round_id === p.round_id).forEach(g => { groups[g.group_id] = g; });
    const ics = readAll_('InterviewCandidates').filter(c => c.round_id === p.round_id);
    const toDelete = [];
    const maxOrder = {};
    ics.forEach(c => { maxOrder[c.group_id] = Math.max(maxOrder[c.group_id] || 0, Number(c.order_no) || 0); });
    let changed = 0;
    p.items.forEach(it => {
      const x = eligMap[it.application_id];
      if (!x) throw new Error('ใบสมัครเลขที่ ' + it.application_id + ' ไม่อยู่ในรอบนี้ หรือยังไม่ผ่านการอนุมัติ');
      const ic = ics.find(c => c.application_id === it.application_id);
      const hasScore = ic && (() => { const t = interviewTotals_(ic, INTERVIEW_DEFAULT_PASS_SCORE); return t.s1 !== null || t.s2 !== null; })();
      if (!it.group_id) {
        if (!ic) return;
        if (hasScore && !p.force) throw new Error((x.data.name || it.application_id) + ' มีคะแนนจากกรรมการแล้ว — ยืนยันเพื่อล้างคะแนนและถอดออกจากกลุ่ม');
        toDelete.push(ic._row);
        changed++;
        return;
      }
      if (!groups[it.group_id]) throw new Error('ไม่พบกลุ่มสอบที่เลือก');
      const time = it.time_slot !== undefined ? String(it.time_slot || '').trim() : undefined;
      if (ic) {
        const moving = ic.group_id !== it.group_id;
        if (moving && hasScore && !p.force) throw new Error((x.data.name || it.application_id) + ' มีคะแนนจากกรรมการแล้ว — ยืนยันเพื่อล้างคะแนนและย้ายกลุ่ม');
        const patch = { updated_at: nowIso_(), updated_by: user.user_id };
        if (moving) {
          patch.group_id = it.group_id; patch.scores_json = '{}'; patch.remarks_json = '{}';
          maxOrder[it.group_id] = (maxOrder[it.group_id] || 0) + 1;
          patch.order_no = it.order_no !== undefined && it.order_no !== '' ? Number(it.order_no) : maxOrder[it.group_id];
        } else if (it.order_no !== undefined && it.order_no !== '') patch.order_no = Number(it.order_no);
        if (time !== undefined) patch.time_slot = time;
        updateRow_('InterviewCandidates', ic._row, patch);
        changed++;
      } else {
        maxOrder[it.group_id] = (maxOrder[it.group_id] || 0) + 1;
        appendRow_('InterviewCandidates', {
          ic_id: newId_('IVC'), round_id: p.round_id, group_id: it.group_id, application_id: it.application_id,
          candidate_no: x.data.candidate_no || '', candidate_name: x.data.name || '',
          order_no: it.order_no !== undefined && it.order_no !== '' ? Number(it.order_no) : maxOrder[it.group_id],
          time_slot: time || '', absent: false, scores_json: '{}', remarks_json: '{}',
          created_by: user.user_id, created_at: nowIso_(), updated_at: nowIso_(), updated_by: user.user_id
        });
        changed++;
      }
    });
    toDelete.sort((a, b) => b - a).forEach(r => deleteRow_('InterviewCandidates', r));
    auditLog_('InterviewSessions', p.round_id, 'candidates_assigned', {}, { changed: changed }, user.user_id);
    return { ok: true, changed: changed };
  });
}

/** จัดกลุ่มอัตโนมัติ: แบ่งผู้ที่ยังไม่มีกลุ่มเรียงตามรหัสประจำตัวสอบ กระจายเป็นช่วงเท่าๆ กันตามจำนวนกลุ่ม พร้อมกำหนดเวลาต่อคน
 * p.start_time 'HH:MM', p.minutes (นาทีต่อคน), p.reset = true จะจัดใหม่ทั้งหมดเฉพาะผู้ที่ยังไม่มีคะแนน/ไม่ได้ระบุว่าไม่เข้าสอบ */
function action_autoAssignInterviewCandidates(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const m = String(p.start_time || '09:00').match(/^(\d{1,2})[:.](\d{2})$/);
  if (!m) throw new Error('เวลาเริ่มต้องอยู่ในรูปแบบ HH:MM เช่น 09:30');
  const minutes = Number(p.minutes || 30);
  if (!isFinite(minutes) || minutes < 5 || minutes > 240) throw new Error('เวลาต่อคนต้องอยู่ระหว่าง 5-240 นาที');
  return interviewWithLock_(() => {
    interviewGetRound_(p.round_id);
    const groups = readAll_('InterviewGroups').filter(g => g.round_id === p.round_id).sort((a, b) => Number(a.group_no) - Number(b.group_no));
    if (!groups.length) throw new Error('กรุณาเพิ่มกลุ่มสอบก่อน');
    const eligible = interviewEligibleApps_(p.round_id);
    let ics = readAll_('InterviewCandidates').filter(c => c.round_id === p.round_id);
    if (p.reset) {
      const removable = ics.filter(c => { const t = interviewTotals_(c, INTERVIEW_DEFAULT_PASS_SCORE); return t.s1 === null && t.s2 === null && !isTrue_(c.absent); });
      removable.map(c => c._row).sort((a, b) => b - a).forEach(r => deleteRow_('InterviewCandidates', r));
      ics = readAll_('InterviewCandidates').filter(c => c.round_id === p.round_id);
    }
    const assigned = {};
    ics.forEach(c => { assigned[c.application_id] = true; });
    const todo = eligible.filter(x => !assigned[x.app.application_id]);
    if (!todo.length) return { ok: true, assigned: 0 };
    const startMin = Number(m[1]) * 60 + Number(m[2]);
    const per = Math.ceil(todo.length / groups.length);
    const counters = {};
    ics.forEach(c => { counters[c.group_id] = Math.max(counters[c.group_id] || 0, Number(c.order_no) || 0); });
    let n = 0;
    todo.forEach((x, idx) => {
      const g = groups[Math.min(Math.floor(idx / per), groups.length - 1)];
      counters[g.group_id] = (counters[g.group_id] || 0) + 1;
      const mins = startMin + (counters[g.group_id] - 1) * minutes;
      appendRow_('InterviewCandidates', {
        ic_id: newId_('IVC'), round_id: p.round_id, group_id: g.group_id, application_id: x.app.application_id,
        candidate_no: x.data.candidate_no || '', candidate_name: x.data.name || '', order_no: counters[g.group_id],
        time_slot: interviewTimeRange_(mins, minutes), absent: false, scores_json: '{}', remarks_json: '{}',
        created_by: user.user_id, created_at: nowIso_(), updated_at: nowIso_(), updated_by: user.user_id
      });
      n++;
    });
    auditLog_('InterviewSessions', p.round_id, 'candidates_auto_assigned', {}, { assigned: n, groups: groups.length }, user.user_id);
    return { ok: true, assigned: n };
  });
}

/** แอดมินระบุ/ยกเลิก "ไม่เข้าสอบ" ของผู้เข้าสอบ (ทำได้แม้ปิดรับคะแนนแล้ว) */
function action_adminSetInterviewAbsent(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  return interviewWithLock_(() => {
    const ic = readAll_('InterviewCandidates').find(c => c.application_id === p.application_id);
    if (!ic) throw new Error('ผู้เข้าสอบรายนี้ยังไม่ได้จัดเข้ากลุ่ม');
    updateRow_('InterviewCandidates', ic._row, { absent: p.absent === true, updated_at: nowIso_(), updated_by: user.user_id });
    auditLog_('InterviewCandidates', ic.ic_id, p.absent === true ? 'absent_set' : 'absent_cleared', {}, {}, user.user_id);
    return { ok: true };
  });
}

/** แอดมินบันทึก/แก้ไข/ล้างคะแนนแทนกรรมการ (เช่น กรอกจากใบคะแนนกระดาษ) — p.slot 'inv1'|'inv2', p.scores = null เพื่อล้างคะแนน */
function action_adminSaveInterviewScore(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  if (INTERVIEW_SLOTS.indexOf(p.slot) === -1) throw new Error('ระบุกรรมการไม่ถูกต้อง');
  return interviewWithLock_(() => {
    const ic = readAll_('InterviewCandidates').find(c => c.application_id === p.application_id);
    if (!ic) throw new Error('ผู้เข้าสอบรายนี้ยังไม่ได้จัดเข้ากลุ่ม');
    const scores = interviewJson_(ic.scores_json, {});
    const remarks = interviewJson_(ic.remarks_json, {});
    if (p.scores === null) { delete scores[p.slot]; delete remarks[p.slot]; }
    else { scores[p.slot] = Object.assign(interviewValidateScores_(p.scores), { by_admin: user.user_id }); remarks[p.slot] = String(p.remark || ''); }
    updateRow_('InterviewCandidates', ic._row, { scores_json: JSON.stringify(scores), remarks_json: JSON.stringify(remarks), updated_at: nowIso_(), updated_by: user.user_id });
    auditLog_('InterviewCandidates', ic.ic_id, p.scores === null ? 'score_cleared_by_admin' : 'score_saved_by_admin', {}, { slot: p.slot, total: p.scores === null ? null : scores[p.slot].total }, user.user_id);
    return { ok: true };
  });
}

/** บันทึกผลสอบสัมภาษณ์เข้าระบบผลสอบเดิม (สอบผ่าน/สอบไม่ผ่าน + ส่งอีเมลแจ้งผู้สมัคร + ใบรับรอง) ของผู้ที่กรรมการให้คะแนนครบ 2 ท่านแล้ว
 * p.dry_run = ตรวจอย่างเดียว, p.overwrite = ทับผลที่บันทึกไปแล้ว, p.application_ids = เลือกเฉพาะบางคน (ไม่ส่ง = ทั้งรอบ) */
function action_applyInterviewResults(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  const round = interviewGetRound_(p.round_id);
  const session = interviewGetSession_(p.round_id, false);
  const passScore = interviewPassScore_(session);
  const eligMap = {};
  interviewEligibleApps_(p.round_id).forEach(x => { eligMap[x.app.application_id] = x; });
  let ics = readAll_('InterviewCandidates').filter(c => c.round_id === p.round_id && eligMap[c.application_id]);
  if (Array.isArray(p.application_ids) && p.application_ids.length) ics = ics.filter(c => p.application_ids.indexOf(c.application_id) !== -1);
  ics.sort(interviewSortCandidates_);
  const results = [];
  let okCount = 0;
  ics.forEach(ic => {
    const x = eligMap[ic.application_id];
    const t = interviewTotals_(ic, passScore);
    const out = { application_id: ic.application_id, candidate_no: ic.candidate_no || '', name: ic.candidate_name || '', s1: t.s1, s2: t.s2, average: t.average, ok: false, message: '' };
    try {
      if (t.status === 'ไม่เข้าสอบ') throw new Error('ไม่เข้าสอบ — ไม่บันทึกผล');
      if (t.average === null) throw new Error('กรรมการให้คะแนนยังไม่ครบ 2 ท่าน');
      if (x.app.status !== 'รอผลสอบ' && !p.overwrite) throw new Error('บันทึกผลสอบไว้แล้ว (' + x.app.status + ') — เลือก "ทับผลเดิม" หากต้องการแก้ไข');
      const built = buildExamResult_(x.app, x.data, {
        result: t.status, score: t.average, note: 'คะแนนเฉลี่ยจากกรรมการ 2 ท่าน (เกณฑ์ผ่านไม่น้อยกว่า ' + passScore + ' คะแนน)'
      }, round);
      out.result = built.exam_result.result;
      if (!p.dry_run) writeExamResult_(x.app, x.data, built, user);
      out.ok = true; okCount++;
      out.message = p.dry_run ? 'พร้อมบันทึก' : 'บันทึกแล้ว';
    } catch (err) {
      out.message = err.message;
    }
    results.push(out);
  });
  if (!p.dry_run && okCount) {
    const s = interviewGetSession_(p.round_id, true, user.user_id);
    updateRow_('InterviewSessions', s._row, { results_applied_at: nowIso_(), updated_at: nowIso_() });
    auditLog_('InterviewSessions', p.round_id, 'interview_results_applied', {}, { ok: okCount, total: ics.length }, user.user_id);
  }
  return { dry_run: !!p.dry_run, total: ics.length, ok_count: okCount, error_count: ics.length - okCount, pass_score: passScore, results: results };
}

// ----- กรรมการ (PUBLIC — ไม่ต้องมีบัญชี: เลือกรอบ → กลุ่ม → กรรมการ 1/2 → รหัสผ่านกลุ่ม แล้วได้ token ชั่วคราว) -----

/** ตรวจ token กรรมการ + ว่ารหัสผ่านกลุ่มยังเป็นรหัสเดิม (ถ้าแอดมินเปลี่ยนรหัส เซสชันเดิมจะหมดอายุทันที) — คืน { group, slot, session, round } */
function interviewAuth_(token, opts) {
  const expired = 'เซสชันกรรมการหมดอายุ กรุณาเข้าสู่ระบบใหม่';
  if (!token) throw new Error(expired);
  const raw = CacheService.getScriptCache().get('ivtok_' + token);
  const payload = raw ? interviewJson_(raw, null) : null;
  if (!payload) throw new Error(expired);
  const group = findRowById_('InterviewGroups', 'group_id', payload.group_id);
  if (!group || String(group.password_hash || '') !== String(payload.ph || '')) throw new Error(expired);
  const session = interviewGetSession_(group.round_id, false);
  const round = findRowById_('ExamRounds', 'round_id', group.round_id);
  if (!round) throw new Error('ไม่พบรอบสอบ');
  if (!(opts && opts.allowClosed) && session && isTrue_(session.scoring_closed)) throw new Error('รอบสอบนี้ปิดรับคะแนนแล้ว ไม่สามารถบันทึกหรือแก้ไขได้');
  return { group: group, slot: payload.slot, session: session, round: round, payload: payload };
}

function interviewerContext_(a) {
  return {
    round_id: a.round.round_id, round_label: a.round.round_label, exam_date: a.round.exam_date || '', location_note: a.round.location_note || '',
    time_text: a.session ? (a.session.time_text || '') : '', group_no: Number(a.group.group_no), group_time_range: a.group.time_range || '',
    slot: a.slot, name: (a.slot === 'inv1' ? a.group.inv1_name : a.group.inv2_name) || (a.slot === 'inv1' ? 'กรรมการ 1' : 'กรรมการ 2'),
    partner_name: (a.slot === 'inv1' ? a.group.inv2_name : a.group.inv1_name) || '',
    scoring_closed: a.session ? isTrue_(a.session.scoring_closed) : false
  };
}

/** รอบสอบที่เปิดรับคะแนน + กลุ่ม + ชื่อกรรมการ (สำหรับหน้าเข้าสู่ระบบของกรรมการ) — ไม่ส่งรหัสผ่าน/รายชื่อผู้เข้าสอบ มีเพียงบอกว่ากลุ่มมีรหัสผ่านหรือไม่ */
function action_interviewerSessions(p) {
  const groups = readAll_('InterviewGroups');
  const rounds = readAll_('ExamRounds');
  const out = [];
  readAll_('InterviewSessions').forEach(s => {
    if (isTrue_(s.scoring_closed)) return;
    const round = rounds.find(r => r.round_id === s.round_id);
    const gs = groups.filter(g => g.round_id === s.round_id).sort((a, b) => Number(a.group_no) - Number(b.group_no));
    if (!round || round.status === 'เสร็จสิ้น' || !gs.length) return;
    out.push({
      round_id: round.round_id, round_label: round.round_label, exam_date: round.exam_date || '', time_text: s.time_text || '',
      groups: gs.map(g => ({ group_id: g.group_id, group_no: Number(g.group_no), inv1_name: g.inv1_name || '', inv2_name: g.inv2_name || '', has_password: !!g.password_hash }))
    });
  });
  return out.sort((x, y) => new Date(y.exam_date || 0) - new Date(x.exam_date || 0));
}

function action_interviewerLogin(p) {
  if (INTERVIEW_SLOTS.indexOf(p.slot) === -1) throw new Error('กรุณาเลือกว่าเป็นกรรมการท่านที่ 1 หรือ 2');
  const group = findRowById_('InterviewGroups', 'group_id', p.group_id);
  if (!group) throw new Error('ไม่พบกลุ่มสอบนี้');
  const session = interviewGetSession_(group.round_id, false);
  if (session && isTrue_(session.scoring_closed)) throw new Error('รอบสอบนี้ปิดรับคะแนนแล้ว');
  const cache = CacheService.getScriptCache();
  const failKey = 'ivfail_' + group.group_id;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= 10) throw new Error('กรอกรหัสผ่านผิดหลายครั้งเกินไป กรุณารอประมาณ 5 นาทีแล้วลองใหม่');
  if (group.password_hash) {
    if (!p.password || hashPassword_(String(p.password)) !== String(group.password_hash)) {
      cache.put(failKey, String(fails + 1), 300);
      throw new Error('รหัสผ่านไม่ถูกต้อง');
    }
  }
  const token = Utilities.getUuid();
  cache.put('ivtok_' + token, JSON.stringify({ group_id: group.group_id, slot: p.slot, ph: String(group.password_hash || '') }), INTERVIEW_TOKEN_TTL_SEC);
  const a = interviewAuth_(token, { allowClosed: true });
  return Object.assign({ token: token, criteria: INTERVIEW_CRITERIA }, interviewerContext_(a));
}

/** รายชื่อผู้เข้าสอบกลุ่มของกรรมการ + คะแนนของตนเองเท่านั้น (ไม่ส่งเลขบัตรประชาชน/คะแนนของกรรมการอีกท่านออกไป) */
function action_interviewerGetData(p) {
  const a = interviewAuth_(p.token, { allowClosed: true });
  const passScore = interviewPassScore_(a.session);
  const ics = readAll_('InterviewCandidates').filter(c => c.group_id === a.group.group_id).sort(interviewSortCandidates_);
  const candidates = ics.map(ic => {
    const sc = interviewJson_(ic.scores_json, {});
    const rm = interviewJson_(ic.remarks_json, {});
    return {
      application_id: ic.application_id, candidate_no: ic.candidate_no || '', name: ic.candidate_name || '', time_slot: ic.time_slot || '',
      order_no: Number(ic.order_no) || 0, absent: isTrue_(ic.absent), my_score: sc[a.slot] || null, my_remark: rm[a.slot] || ''
    };
  });
  return Object.assign({ criteria: INTERVIEW_CRITERIA, pass_score: passScore, candidates: candidates }, interviewerContext_(a));
}

function interviewerFindOwnCandidate_(a, applicationId) {
  const ic = readAll_('InterviewCandidates').find(c => c.application_id === applicationId);
  if (!ic || ic.group_id !== a.group.group_id) throw new Error('ไม่มีสิทธิ์แก้ไขผู้เข้าสอบรายนี้');
  return ic;
}

function action_interviewerSaveScore(p) {
  return interviewWithLock_(() => {
    const a = interviewAuth_(p.token);
    const ic = interviewerFindOwnCandidate_(a, p.application_id);
    if (isTrue_(ic.absent)) throw new Error('ผู้เข้าสอบรายนี้ถูกระบุว่า "ไม่เข้าสอบ" กรุณายกเลิกสถานะก่อนจึงจะให้คะแนนได้');
    const clean = interviewValidateScores_(p.scores);
    const scores = interviewJson_(ic.scores_json, {});
    const remarks = interviewJson_(ic.remarks_json, {});
    scores[a.slot] = clean;
    remarks[a.slot] = String(p.remark || '').slice(0, 1000);
    updateRow_('InterviewCandidates', ic._row, { scores_json: JSON.stringify(scores), remarks_json: JSON.stringify(remarks), updated_at: nowIso_(), updated_by: 'กรรมการกลุ่ม ' + a.group.group_no + ' (' + a.slot + ')' });
    auditLog_('InterviewCandidates', ic.ic_id, 'score_saved_by_interviewer', {}, { group_no: a.group.group_no, slot: a.slot, total: clean.total }, 'interviewer:' + a.group.group_no + ':' + a.slot);
    return { ok: true, my_score: clean, my_remark: remarks[a.slot] };
  });
}

function action_interviewerSetAbsent(p) {
  return interviewWithLock_(() => {
    const a = interviewAuth_(p.token);
    const ic = interviewerFindOwnCandidate_(a, p.application_id);
    updateRow_('InterviewCandidates', ic._row, { absent: p.absent === true, updated_at: nowIso_(), updated_by: 'กรรมการกลุ่ม ' + a.group.group_no + ' (' + a.slot + ')' });
    auditLog_('InterviewCandidates', ic.ic_id, p.absent === true ? 'absent_set_by_interviewer' : 'absent_cleared_by_interviewer', {}, {}, 'interviewer:' + a.group.group_no + ':' + a.slot);
    return { ok: true, absent: p.absent === true };
  });
}

// ---------- 7. ระบบข้อมูลสวัสดิการสมาชิก (TOR ข้อ 7) ----------

// ---------- 7.1 ข้อมูลผู้รับผลประโยชน์ (รองรับหลายรายการ) ----------

function action_listMyBeneficiaries(p, user) {
  requireRole_(user, ['member']);
  if (!user.member_id) throw new Error('บัญชีนี้ยังไม่ผูกกับข้อมูลสมาชิก');
  return readAll_('Beneficiaries').filter(b => b.member_id === user.member_id)
    .sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
}

function action_addBeneficiary(p, user) {
  requireRole_(user, ['member']);
  if (!user.member_id) throw new Error('บัญชีนี้ยังไม่ผูกกับข้อมูลสมาชิก');
  if (!p.name || !p.phone || !p.relationship) throw new Error('กรุณากรอกชื่อ-นามสกุล เบอร์โทรศัพท์ และความสัมพันธ์ให้ครบ');
  const row = {
    beneficiary_id: newId_('BEN'), member_id: user.member_id,
    prefix: p.prefix || '', name: p.name, phone: p.phone, relationship: p.relationship,
    created_at: nowIso_()
  };
  appendRow_('Beneficiaries', row);
  return row;
}

function action_updateBeneficiary(p, user) {
  requireRole_(user, ['member']);
  const row = findRowById_('Beneficiaries', 'beneficiary_id', p.beneficiary_id);
  if (!row || row.member_id !== user.member_id) throw new Error('ไม่พบข้อมูลผู้รับผลประโยชน์นี้');
  updateRow_('Beneficiaries', row._row, {
    prefix: p.prefix || '', name: p.name, phone: p.phone, relationship: p.relationship
  });
  return { ok: true };
}

function action_deleteBeneficiary(p, user) {
  requireRole_(user, ['member']);
  const row = findRowById_('Beneficiaries', 'beneficiary_id', p.beneficiary_id);
  if (!row || row.member_id !== user.member_id) throw new Error('ไม่พบข้อมูลผู้รับผลประโยชน์นี้');
  const inUse = readAll_('Welfare').some(w => w.beneficiary_id === p.beneficiary_id && ['รอตรวจสอบ', 'รออนุมัติ', 'อนุมัติ'].indexOf(w.status) !== -1);
  if (inUse) throw new Error('ไม่สามารถลบได้ เนื่องจากมีคำขอสวัสดิการที่ยังดำเนินการอยู่ผูกกับผู้รับผลประโยชน์รายนี้');
  deleteRow_('Beneficiaries', row._row);
  return { ok: true };
}

// ---------- 7.2 ขอรับสวัสดิการ ----------

function action_getWelfareTypes(p) {
  return { types: WELFARE_TYPES, maxAmounts: WELFARE_MAX_AMOUNTS, claimDeadlineDays: WELFARE_CLAIM_DEADLINE_DAYS };
}

function action_submitWelfareRequest(p, user) {
  requireRole_(user, ['member']);
  if (!user.member_id) throw new Error('บัญชีนี้ยังไม่ผูกกับข้อมูลสมาชิก');
  const member = findRowById_('Members', 'member_id', user.member_id);
  if (!member) throw new Error('เมนูนี้ใช้สำหรับสมาชิกบุคคลเท่านั้น');
  if (member.member_type !== 'สามัญ') throw new Error('สิทธิรับสวัสดิการตามระเบียบนี้สงวนไว้สำหรับสมาชิกประเภทสามัญเท่านั้น (ข้อ 5 ของระเบียบ)');
  if (WELFARE_TYPES.indexOf(p.welfare_type) === -1) throw new Error('กรุณาเลือกประเภทสวัสดิการให้ถูกต้อง');
  if (p.beneficiary_id) {
    const ben = findRowById_('Beneficiaries', 'beneficiary_id', p.beneficiary_id);
    if (!ben || ben.member_id !== user.member_id) throw new Error('ไม่พบข้อมูลผู้รับผลประโยชน์ที่เลือก');
  }
  if (p.incident_date) {
    const days = daysUntil_(p.incident_date);
    if (days !== null && -days > WELFARE_CLAIM_DEADLINE_DAYS) {
      throw new Error('พ้นกำหนดสิทธิขอรับสวัสดิการแล้ว (เกิน ' + WELFARE_CLAIM_DEADLINE_DAYS + ' วันนับแต่วันเกิดเหตุ/ถึงแก่กรรม ตามข้อ 7 ของระเบียบ)');
    }
  }

  const welfareId = newId_('WEL');
  const row = {
    welfare_id: welfareId, member_id: user.member_id, welfare_type: p.welfare_type,
    beneficiary_id: p.beneficiary_id || '', description: p.description || '', incident_date: p.incident_date || '',
    documents_json: JSON.stringify(p.documents || []), status: 'รอตรวจสอบ', reject_reason: '',
    requested_at: nowIso_(), reviewed_by: '', reviewed_at: '', approved_amount: '', note: '',
    disbursed_at: '', disbursed_by: '', disbursement_doc_url: ''
  };
  appendRow_('Welfare', row);
  (p.documents || []).forEach(d => logWelfareDoc_(welfareId, 'upload', d.doc_type, d.url, user));

  if (member) {
    mailNotify_(member.email, 'ได้รับคำขอรับสวัสดิการแล้ว',
      'ระบบได้รับคำขอรับ' + p.welfare_type + ' ของท่านแล้ว เลขที่คำขอ ' + welfareId + '\nสถานะปัจจุบัน: รอตรวจสอบ');
  }
  return { welfare_id: welfareId };
}

function logWelfareDoc_(welfareId, action, docType, url, user) {
  appendRow_('WelfareDocLog', {
    log_id: newId_('WDL'), welfare_id: welfareId, action: action, doc_type: docType || '', url: url || '',
    actor_user_id: user.user_id, actor_role: user.role, at: nowIso_()
  });
}

/** เพิ่ม/แทนที่เอกสารประกอบคำขอสวัสดิการที่ยื่นไปแล้ว — ใช้ได้ทั้งสมาชิก (เจ้าของคำขอ) และแอดมิน (TOR ข้อ 7.3) */
function action_addWelfareDocument(p, user) {
  const isAdmin = ['admin_super', 'admin_staff', 'admin_welfare'].indexOf(user.role) !== -1;
  const row = findRowById_('Welfare', 'welfare_id', p.welfare_id);
  if (!row) throw new Error('ไม่พบคำขอสวัสดิการนี้');
  if (!isAdmin) {
    requireRole_(user, ['member']);
    if (row.member_id !== user.member_id) throw new Error('ไม่มีสิทธิ์แก้ไขคำขอนี้');
  }
  const docs = JSON.parse(row.documents_json || '[]');
  docs.push({ doc_type: p.doc_type, url: p.url });
  updateRow_('Welfare', row._row, { documents_json: JSON.stringify(docs) });
  logWelfareDoc_(p.welfare_id, 'upload', p.doc_type, p.url, user);
  return { ok: true };
}

function action_myWelfareRequests(p, user) {
  requireRole_(user, ['member']);
  if (!user.member_id) throw new Error('บัญชีนี้ยังไม่ผูกกับข้อมูลสมาชิก');
  return readAll_('Welfare').filter(w => w.member_id === user.member_id)
    .sort((a, b) => new Date(b.requested_at) - new Date(a.requested_at));
}

// ---------- ฝั่งแอดมิน (เจ้าหน้าที่สวัสดิการ) ----------

function action_listWelfare(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff', 'admin_welfare']);
  let rows = readAll_('Welfare');
  if (p && p.status) rows = rows.filter(w => w.status === p.status);
  // แนบชื่อ-รูปสมาชิกไปกับแต่ละแถวเลย (แทนที่จะโชว์แค่ member_id ดิบ) ไว้ให้หน้าตารางแสดงรูปสมาชิกได้ทันทีโดยไม่ต้อง join เพิ่มฝั่งหน้าเว็บ
  const members = readAll_('Members');
  rows = rows.map(w => {
    const m = members.find(x => x.member_id === w.member_id);
    return Object.assign({}, w, {
      member_name: m ? [m.prefix_th, m.first_name_th, m.last_name_th].filter(Boolean).join(' ') : w.member_id,
      member_photo_url: m ? (m.photo_url || '') : ''
    });
  });
  return rows.sort((a, b) => new Date(b.requested_at) - new Date(a.requested_at));
}

function action_getWelfareDetail(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff', 'admin_welfare']);
  const row = findRowById_('Welfare', 'welfare_id', p.welfare_id);
  if (!row) throw new Error('ไม่พบคำขอสวัสดิการนี้');
  const member = findRowById_('Members', 'member_id', row.member_id);
  const beneficiary = row.beneficiary_id ? findRowById_('Beneficiaries', 'beneficiary_id', row.beneficiary_id) : null;
  const docLog = readAll_('WelfareDocLog').filter(l => l.welfare_id === row.welfare_id)
    .sort((a, b) => new Date(b.at) - new Date(a.at));
  return { welfare: row, member: member, beneficiary: beneficiary, docLog: docLog };
}

/** แอดมินอนุมัติ/ไม่อนุมัติคำขอ พร้อมระบุจำนวนเงินที่อนุมัติและหมายเหตุ (TOR ข้อ 7.2) */
function action_reviewWelfareRequest(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff', 'admin_welfare']);
  const row = findRowById_('Welfare', 'welfare_id', p.welfare_id);
  if (!row) throw new Error('ไม่พบคำขอสวัสดิการนี้');
  if (row.status !== 'รอตรวจสอบ' && row.status !== 'รออนุมัติ') throw new Error('คำขอนี้ผ่านการพิจารณาไปแล้ว');

  const member = findRowById_('Members', 'member_id', row.member_id);

  if (p.decision === 'reject') {
    const patch = { status: 'ไม่อนุมัติ', reject_reason: p.reject_reason || '', reviewed_by: user.user_id, reviewed_at: nowIso_(), note: p.note || '' };
    updateRow_('Welfare', row._row, patch);
    auditLog_('Welfare', row.welfare_id, 'welfare_rejected', row, patch, user.user_id);
    if (member) mailNotify_(member.email, 'ผลการพิจารณาคำขอรับสวัสดิการ',
      'คำขอรับ' + row.welfare_type + ' เลขที่ ' + row.welfare_id + ' ไม่ได้รับการอนุมัติ\nเหตุผล: ' + (p.reject_reason || '-'));
    return { ok: true };
  }

  if (!p.approved_amount || isNaN(Number(p.approved_amount))) throw new Error('กรุณาระบุจำนวนเงินที่อนุมัติ');
  const cap = WELFARE_MAX_AMOUNTS[row.welfare_type];
  if (cap && Number(p.approved_amount) > cap && !p.override_cap) {
    throw new Error('จำนวนเงินที่อนุมัติ (' + p.approved_amount + ' บาท) เกินวงเงินสูงสุดตามระเบียบข้อ 22 สำหรับ "' + row.welfare_type + '" (ไม่เกิน ' + cap + ' บาท) — ถ้ามีเหตุผลพิเศษให้ยืนยันการอนุมัติเกินวงเงินอีกครั้ง (override_cap)');
  }
  const approvePatch = {
    status: 'อนุมัติ', approved_amount: p.approved_amount, note: p.note || '',
    reviewed_by: user.user_id, reviewed_at: nowIso_()
  };
  updateRow_('Welfare', row._row, approvePatch);
  auditLog_('Welfare', row.welfare_id, 'welfare_approved', row, approvePatch, user.user_id);
  if (member) mailNotify_(member.email, 'คำขอรับสวัสดิการได้รับการอนุมัติ',
    'คำขอรับ' + row.welfare_type + ' เลขที่ ' + row.welfare_id + ' ได้รับการอนุมัติ\nจำนวนเงิน: ' + p.approved_amount + ' บาท\nรอการเบิกจ่ายจากเจ้าหน้าที่');
  return { ok: true };
}

/** บันทึกการเบิกจ่ายเงินสวัสดิการ + ออกเอกสารยืนยันการจ่ายเงิน (PDF) ให้สมาชิก (TOR ข้อ 7.2) */
function action_recordWelfareDisbursement(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff', 'admin_welfare']);
  const row = findRowById_('Welfare', 'welfare_id', p.welfare_id);
  if (!row) throw new Error('ไม่พบคำขอสวัสดิการนี้');
  if (row.status !== 'อนุมัติ') throw new Error('ต้องอนุมัติคำขอก่อนจึงจะบันทึกการเบิกจ่ายได้');

  const member = findRowById_('Members', 'member_id', row.member_id);
  const memberName = member ? [member.prefix_th, member.first_name_th, member.last_name_th].filter(Boolean).join(' ') : row.member_id;

  const doc = DocumentApp.create('ใบยืนยันการจ่ายเงินสวัสดิการ_' + row.welfare_id);
  const body = doc.getBody();
  body.appendParagraph('ใบยืนยันการจ่ายเงินสวัสดิการ').setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('สภาวิชาชีพสังคมสงเคราะห์').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(' ');
  body.appendParagraph('เลขที่คำขอ: ' + row.welfare_id);
  body.appendParagraph('ชื่อ-นามสกุลสมาชิก: ' + memberName);
  body.appendParagraph('ประเภทสวัสดิการ: ' + row.welfare_type);
  body.appendParagraph('จำนวนเงินที่จ่าย: ' + row.approved_amount + ' บาท');
  body.appendParagraph('วันที่จ่ายเงิน: ' + new Date().toLocaleDateString('th-TH'));
  if (row.note) body.appendParagraph('หมายเหตุ: ' + row.note);
  body.appendParagraph(' ');
  body.appendParagraph('_____________________________').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('เจ้าหน้าที่ผู้เบิกจ่าย').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  const pdfBlob = file.getAs('application/pdf');
  const folder = getDocsFolder_();
  const pdfFile = folder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  file.setTrashed(true);

  const disbursePatch = { disbursed_at: nowIso_(), disbursed_by: user.user_id, disbursement_doc_url: pdfFile.getUrl() };
  updateRow_('Welfare', row._row, disbursePatch);
  auditLog_('Welfare', row.welfare_id, 'welfare_disbursed', row, disbursePatch, user.user_id);
  if (member) mailNotify_(member.email, 'จ่ายเงินสวัสดิการเรียบร้อยแล้ว',
    'ระบบได้บันทึกการจ่ายเงินสวัสดิการ ' + row.approved_amount + ' บาท ให้ท่านเรียบร้อยแล้ว\nดาวน์โหลดเอกสารยืนยันได้ที่ Member Portal');

  return { ok: true, disbursement_doc_url: pdfFile.getUrl() };
}

function action_listPayments(p, user) {
  requireRole_(user, ['admin_super', 'admin_staff']);
  return readAll_('Payments'); // TODO (TOR ข้อ 11): เชื่อม Payment Gateway จริง + ตรวจสลิป
}

// ---------- ThaID (DOPA Digital ID) — เข้าสู่ระบบด้วยแอป ThaID ----------
// อ้างอิงมาตรฐาน OpenID Connect ที่ ThaID ใช้จริง (ตัวอย่างทางการจาก ETDA: github.com/ETDA/ThaiD-PHP-RP)
// ต้องลงทะเบียนเป็น Relying Party กับกรมการปกครองก่อน — ดู docs/THAID_SETUP.md ว่าต้องขอที่ไหน/ใช้เวลาแค่ไหน
// ตั้งค่าที่ต้องเพิ่มใน Script Properties หลังได้รับอนุมัติ:
//   THAID_CLIENT_ID, THAID_CLIENT_SECRET, THAID_SCOPE (ปกติคือ "pid"), FRONTEND_BASE_URL (เช่น https://swpcth.github.io/demo_SWPC_Database/)
//   PROMPTPAY_ID (เลขพร้อมเพย์ของสภาฯ สำหรับสร้าง QR รับชำระเงิน — เบอร์มือถือ 10 หลัก หรือเลขประจำตัวผู้เสียภาษี/นิติบุคคล 13 หลัก ที่ผูกกับพร้อมเพย์ไว้แล้ว)
// Callback URL ที่ต้องแจ้งกรมการปกครองตอนลงทะเบียน = Web App URL ของสคริปต์นี้เอง (Deploy > Manage deployments)

const THAID_WELLKNOWN_URL = 'https://imauth.bora.dopa.go.th/.well-known/openid-configuration';

function getThaidWellKnown_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('thaid_wellknown');
  if (cached) return JSON.parse(cached);
  const res = UrlFetchApp.fetch(THAID_WELLKNOWN_URL, { muteHttpExceptions: true });
  const json = JSON.parse(res.getContentText());
  cache.put('thaid_wellknown', JSON.stringify(json), 3600);
  return json;
}

function decodeJwtPayload_(jwt) {
  const parts = jwt.split('.');
  if (parts.length < 2) throw new Error('id_token จาก ThaID มีรูปแบบไม่ถูกต้อง');
  let payload = parts[1].replace(/-/g, '+').replace(/_/g, '/');
  while (payload.length % 4) payload += '=';
  const decoded = Utilities.newBlob(Utilities.base64Decode(payload)).getDataAsString();
  return JSON.parse(decoded);
}

/** ขั้นที่ 1: ผู้ใช้กดปุ่ม "เข้าสู่ระบบด้วย ThaID" > เว็บพาไปที่ URL นี้ > พาไปหน้า ThaID ต่อ */
function thaidRedirectResponse_(role) {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('THAID_CLIENT_ID');
  const scope = props.getProperty('THAID_SCOPE') || 'pid';
  const frontendBase = props.getProperty('FRONTEND_BASE_URL') || '';

  if (!clientId) {
    return HtmlService.createHtmlOutput(
      '<p>ยังไม่ได้ตั้งค่า ThaID ในระบบ (ขาด THAID_CLIENT_ID ใน Script Properties) — ' +
      'ต้องลงทะเบียนเป็น Relying Party กับกรมการปกครองก่อน ดู docs/THAID_SETUP.md</p>'
    );
  }
  try {
    const wk = getThaidWellKnown_();
    const redirectUri = ScriptApp.getService().getUrl();
    const state = role + ':' + Utilities.getUuid();
    const authUrl = wk.authorization_endpoint +
      '?response_type=code' +
      '&client_id=' + encodeURIComponent(clientId) +
      '&redirect_uri=' + encodeURIComponent(redirectUri) +
      '&scope=' + encodeURIComponent(scope) +
      '&state=' + encodeURIComponent(state);
    return HtmlService.createHtmlOutput('<script>window.location.replace(' + JSON.stringify(authUrl) + ');</script>กำลังนำท่านไปยัง ThaID...');
  } catch (err) {
    return htmlRedirectWithError_(frontendBase + (role === 'admin' ? 'admin-login.html' : 'index.html'), 'เชื่อมต่อ ThaID ไม่สำเร็จ: ' + err.message);
  }
}

/** ขั้นที่ 2: ThaID ส่ง browser กลับมาที่ Web App URL นี้พร้อม ?code=...&state=... */
function thaidCallbackResponse_(code, state) {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('THAID_CLIENT_ID');
  const clientSecret = props.getProperty('THAID_CLIENT_SECRET');
  const frontendBase = props.getProperty('FRONTEND_BASE_URL') || '';
  const role = (String(state || '').split(':')[0] === 'admin') ? 'admin' : 'member';
  const loginPage = frontendBase + (role === 'admin' ? 'admin-login.html' : 'index.html');

  try {
    const wk = getThaidWellKnown_();
    const redirectUri = ScriptApp.getService().getUrl();
    const basicAuth = Utilities.base64Encode(clientId + ':' + clientSecret);
    const tokenRes = UrlFetchApp.fetch(wk.token_endpoint, {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      headers: { Authorization: 'Basic ' + basicAuth },
      payload: { grant_type: 'authorization_code', code: code, redirect_uri: redirectUri },
      muteHttpExceptions: true
    });
    const tokenData = JSON.parse(tokenRes.getContentText());
    if (!tokenData.id_token) throw new Error('ThaID ไม่คืน id_token: ' + tokenRes.getContentText());

    const claims = decodeJwtPayload_(tokenData.id_token);
    const pid = claims.pid || claims.cid || claims.sub; // ชื่อ claim อาจต้องปรับตามเอกสารจริงที่ได้รับตอนลงทะเบียน
    if (!pid) throw new Error('ไม่พบเลขประจำตัวประชาชนใน id_token ของ ThaID');

    let user = readAll_('Users').find(u => u.thaid_pid && String(u.thaid_pid) === String(pid));

    if (!user && role === 'member') {
      // สมาชิก/ประชาชนทั่วไปที่ยืนยันตัวตนด้วย ThaID ครั้งแรก — สร้างบัญชีอัตโนมัติ
      // (ยังไม่ผูกกับข้อมูลสมาชิกจนกว่าจะยื่นใบสมัครและได้รับอนุมัติ)
      const newUser = {
        user_id: newId_('USR'), username: 'thaid_' + pid, password_hash: '',
        role: 'member', member_id: '', thaid_pid: pid, status: 'active',
        created_at: nowIso_(), last_login: nowIso_()
      };
      appendRow_('Users', newUser);
      user = newUser;
    }

    if (!user) {
      throw new Error('บัญชีนี้ยังไม่ได้ผูกกับ ThaID ในระบบ กรุณาติดต่อผู้ดูแลระบบให้ผูกบัญชีก่อนใช้งาน');
    }
    if (role === 'admin' && user.role.indexOf('admin') !== 0) {
      throw new Error('บัญชีนี้ไม่มีสิทธิ์เจ้าหน้าที่ (admin)');
    }

    const token = makeToken_(user);
    if (user._row) updateRow_('Users', user._row, { last_login: nowIso_() });

    const target = frontendBase + 'auth-callback.html#token=' + encodeURIComponent(token) +
      '&role=' + encodeURIComponent(user.role) + '&username=' + encodeURIComponent(user.username);
    return HtmlService.createHtmlOutput('<script>window.location.replace(' + JSON.stringify(target) + ');</script>เข้าสู่ระบบสำเร็จ กำลังนำท่านเข้าสู่ระบบ...');
  } catch (err) {
    return htmlRedirectWithError_(loginPage, err.message);
  }
}

function htmlRedirectWithError_(url, message) {
  const target = url + (url.indexOf('#') === -1 ? '#' : '&') + 'error=' + encodeURIComponent(message);
  return HtmlService.createHtmlOutput('<script>window.location.replace(' + JSON.stringify(target) + ');</script>' + message);
}

// ---------- Router ----------

const PUBLIC_ACTIONS = ['login', 'register', 'getPdpaPolicy', 'claimExistingMemberAccount', 'setupFirstAdmin', 'publicSearchMembers', 'verifyQrToken', 'publicStats', 'publicExamResult', 'interviewerSessions', 'interviewerLogin', 'interviewerGetData', 'interviewerSaveScore', 'interviewerSetAbsent', 'verifyReceipt', 'getBackendVersion', 'getPromptPayConfig'];

function action_getBackendVersion(p) {
  return { version: BACKEND_VERSION };
}

/** คืนค่าเลข PromptPay ของสภาฯ ให้หน้าเว็บใช้สร้าง QR การชำระเงิน (ไม่ใช่ข้อมูลอ่อนไหว เหมือนเลขบัญชีธนาคารที่เปิดเผยอยู่แล้ว) */
function action_getPromptPayConfig(p) {
  const id = PropertiesService.getScriptProperties().getProperty('PROMPTPAY_ID') || '';
  return { promptPayId: id };
}

function handle_(p) {
  if (p.__webhook__) {
    return handlePaymentWebhook_(p.__rawBody__);
  }
  const action = p.action;
  const fnName = 'action_' + action;
  if (typeof this[fnName] !== 'function') {
    throw new Error('ไม่รู้จัก action: ' + action);
  }
  if (PUBLIC_ACTIONS.indexOf(action) !== -1) {
    return this[fnName](p);
  }
  const user = getSessionUser_(p.token);
  if (!user) throw new Error('กรุณาเข้าสู่ระบบใหม่ (token หมดอายุ)');
  return this[fnName](p, user);
}

function doGet(e) {
  const p = e.parameter || {};

  // ThaID ส่ง code กลับมาที่ URL นี้ตรงๆ (ขั้นที่ 2)
  if (p.code) {
    return thaidCallbackResponse_(p.code, p.state);
  }
  // ผู้ใช้กดปุ่ม "เข้าสู่ระบบด้วย ThaID" จากหน้าเว็บ (ขั้นที่ 1)
  if (p.action === 'thaidRedirect') {
    return thaidRedirectResponse_(p.role === 'admin' ? 'admin' : 'member');
  }
  // ผู้ใช้กด "ยกเลิก" ในหน้า ThaID — บาง provider ส่งกลับมาเป็น ?error=...
  if (p.error) {
    const frontendBase = PropertiesService.getScriptProperties().getProperty('FRONTEND_BASE_URL') || '';
    return htmlRedirectWithError_(frontendBase + 'index.html', 'ยกเลิกการเข้าสู่ระบบด้วย ThaID');
  }

  return respond_(safeHandle_(p));
}

function doPost(e) {
  // TOR ข้อ 11: รองรับ Webhook/Callback จาก Payment Gateway — แยกเส้นทางออกจาก action ปกติของเว็บเรา
  // เรียกผ่าน URL ...?webhook=true (ผู้ให้บริการ Gateway จริงจะ POST มาที่ path นี้เมื่อเชื่อมต่อแล้ว)
  if (e.parameter && e.parameter.webhook === 'true') {
    return respond_(safeHandle_({ __webhook__: true, __rawBody__: e.postData ? e.postData.contents : '' }));
  }
  let p = {};
  try {
    p = JSON.parse(e.postData.contents);
  } catch (err) {
    p = e.parameter;
  }
  return respond_(safeHandle_(p));
}

function safeHandle_(p) {
  try {
    const data = handle_(p);
    return { ok: true, data: data };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function respond_(obj) {
  // ใช้ text/plain เพื่อเลี่ยง CORS preflight (ฝั่งเว็บ fetch แบบ POST + text/plain ได้)
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

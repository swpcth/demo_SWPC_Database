# PDPA: ความยินยอม + การเข้ารหัสข้อมูลอ่อนไหว (v21 — `2026-11-15-pdpa-v21`)

ตอบ TOR "ข้อกำหนดเพิ่มเติม: การคุ้มครองข้อมูลส่วนบุคคล (PDPA)" (หัวข้อ 3.1.2)

## 1. การเข้ารหัสข้อมูลอ่อนไหว (Encryption at rest)
- เข้ารหัส **AES-256-CBC + HMAC-SHA256** (เขียนด้วย JavaScript ล้วนใน `Code.gs` เพราะ Apps Script ไม่มี AES ในตัว) ผูกกับ `readAll_` / `appendRow_` / `updateRow_` — โค้ดส่วนอื่นเห็นค่าปกติ แต่ในไฟล์ Google Sheets เป็น `enc1:...`
- ฟิลด์ที่เข้ารหัส: `Members.national_id / religion / ethnicity`, `MembersOrg.rep_national_id / rep_religion / rep_ethnicity`, และค่าของคีย์เหล่านี้ที่ฝังใน JSON ของ `Applications.ref_data_json` และ `AuditLog.old_value_json / new_value_json`
- กุญแจ: Script Properties `DATA_ENCRYPTION_KEY` (hex 64 ตัว) — **สร้างอัตโนมัติครั้งแรกที่บันทึกข้อมูล หรือกำหนดเองได้** *ห้ามลบ/เปลี่ยน และต้องสำรองกุญแจแยกจากไฟล์ข้อมูล* ถ้ากุญแจหาย ข้อมูลที่เข้ารหัสแล้วถอดไม่ได้
- ถอดรหัสตรวจ HMAC ทุกครั้ง — ข้อมูลที่ถูกแก้ในชีตหรือกุญแจไม่ตรงจะแจ้ง error (ไม่ส่งข้อมูลที่ผิดต่อ)
- ข้อมูลเดิมที่ยังเป็นข้อความธรรมดายังอ่านได้ตามปกติ — ผู้ดูแลสูงสุดกด **"เข้ารหัสข้อมูลเดิม"** ที่หน้า `admin-pdpa.html` (หรือ `action_encryptExistingSensitiveData`) หนึ่งครั้งหลังอัปเดต รันซ้ำได้ปลอดภัย และตรวจสถานะได้ด้วย `getEncryptionStatus`
- การจำกัดสิทธิ์ตามบทบาท: รายชื่อ/รายละเอียดสมาชิกเรียกได้เฉพาะ admin_super / admin_staff (admin_welfare ไม่เห็น); สมาชิกเห็นเฉพาะข้อมูลของตนเอง; หน้าค้นหาสาธารณะแสดงเฉพาะ ชื่อ เลขที่ใบอนุญาต สถานะ วันหมดอายุ
- **ข้อจำกัดที่ควรรู้**: ผู้ที่เปิดสคริปต์/Script Properties ได้ (เจ้าของโปรเจกต์ Apps Script) เข้าถึงกุญแจได้ — การเข้ารหัสป้องกันการรั่วจากไฟล์ชีต/การแชร์ชีต/สำเนาชีต ไม่ใช่จากผู้ดูแลสคริปต์ ส่วนระดับ production ควรย้ายไปฐานข้อมูล + KMS

## 2. ความยินยอม (Consent)
- ชีตใหม่ `PdpaConsents` (append-only; แถวล่าสุดของบัญชี = สถานะปัจจุบัน) และ `PdpaRequests` — รัน `setupAllSheets` เพื่อสร้างแท็บ
- ข้อความประกาศ/เวอร์ชันอยู่ใน `getPdpaPolicy_()` (`PDPA_POLICY_VERSION = '1.0'`) — **เป็นร่างตั้งต้น ฝ่ายกฎหมาย/DPO ของสภาฯ ต้องตรวจแก้** (ระยะเวลาเก็บรักษา ช่องทางติดต่อ — ตั้ง `PDPA_CONTACT` ใน Script Properties) เปลี่ยนเวอร์ชันเมื่อใด สมาชิกทุกคนต้องกดยินยอมใหม่
- สมัครใช้งานบัญชี (`register`, `claimExistingMemberAccount`) ต้องติ๊กยอมรับ — backend ปฏิเสธถ้าไม่ส่ง `pdpa_accept` + `pdpa_version` ตรงเวอร์ชัน
- บัญชีเดิม/ThaID ที่ยังไม่เคยยินยอม: `api.js` พาไป `pdpa-consent.html` ก่อนใช้หน้าอื่น และ backend บังคับที่ `submitApplication` (`requirePdpaConsent_`)
- ข้อมูลเดโมถูกบันทึกว่ายินยอมแล้ว (ช่อง channel = "ข้อมูลตัวอย่าง (Demo)")

## 3. การใช้สิทธิของเจ้าของข้อมูล
หน้า `pdpa-consent.html`: ถอนความยินยอม (ต้องระบุเหตุผล), ยื่นคำขอ 6 ประเภท (เข้าถึง/แก้ไข/ลบ/ระงับ/คัดค้าน/โอนย้าย), ดาวน์โหลดสำเนาข้อมูลตนเอง (JSON), ดูประวัติ
ฝั่งเจ้าหน้าที่ `admin-pdpa.html` (staff/super): ภาพรวมความยินยอม, รับ/ตอบคำขอ (ต้องมีข้อความตอบเมื่อ "ดำเนินการแล้ว/ปฏิเสธ"), ผู้ดูแลสูงสุดเห็นสถานะการเข้ารหัส
ทุกการกระทำบันทึกใน AuditLog

## Actions ใหม่
`getPdpaPolicy` (public), `getMyPdpaStatus`, `acceptPdpaConsent`, `withdrawPdpaConsent`, `listMyPdpaHistory`, `submitPdpaRequest`, `exportMyPersonalData`, `listPdpaRequests`, `handlePdpaRequest`, `listPdpaConsents`, `encryptExistingSensitiveData` (super), `getEncryptionStatus` (super)

## ไฟล์ที่เปลี่ยน/เพิ่ม
`Code.gs`, `api.js` (เวอร์ชัน + PDPA gate), `index.html` (ช่องติ๊กยินยอม), `pdpa-consent.html` / `pdpa-policy.html` / `admin-pdpa.html` (ใหม่), nav ของ member-*.html ทั้ง 11 ไฟล์ และ admin-*.html ทั้ง 23 ไฟล์ — **ต้อง Deploy Apps Script เวอร์ชันใหม่ และรัน `setupAllSheets`**

## การทดสอบที่ทำ (ชุดจำลองใน Node/Playwright ไม่ใช่ Apps Script จริง)
AES/SHA/HMAC ตรวจเทียบกับ Node `crypto` (ถอดรหัสข้ามกันได้), ชุดทดสอบ 36 กรณีด้วย `Code.gs` จริงกับชีตจำลอง (เข้ารหัสทุกเส้นทางเขียน, อ่านข้อมูลเดิม, migration, consent/request/role) ผ่านทั้งหมด — ควรทดสอบบน Apps Script จริงก่อนใช้งาน โดยเฉพาะเวลาประมวลผลเมื่อมีข้อมูลหลายพันแถว

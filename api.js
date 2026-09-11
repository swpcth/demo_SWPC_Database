// ตั้งค่า URL ของ Apps Script Web App ที่ deploy แล้ว (ลงท้ายด้วย /exec)
const API_BASE_URL = 'https://script.google.com/macros/s/AKfycbyzhzXrSnGYr-i9HkPBWtyLPhBJce2HcBOKjQzxbufvYWNaLVvKmhu34AJOvGok88nOIQ/exec';

const Auth = {
  getToken() { return localStorage.getItem('swc_token'); },
  getRole() { return localStorage.getItem('swc_role'); },
  getUsername() { return localStorage.getItem('swc_username'); },
  setSession({ token, role, username }) {
    localStorage.setItem('swc_token', token);
    localStorage.setItem('swc_role', role);
    localStorage.setItem('swc_username', username);
  },
  clear() {
    localStorage.removeItem('swc_token');
    localStorage.removeItem('swc_role');
    localStorage.removeItem('swc_username');
  },
  isLoggedIn() { return !!this.getToken(); },
  requireLogin(redirectTo = '/index.html') {
    if (!this.isLoggedIn()) window.location.href = redirectTo;
  },
  requireRole(roles, redirectTo = '/index.html') {
    this.requireLogin(redirectTo);
    if (!roles.includes(this.getRole())) {
      alert('บัญชีนี้ไม่มีสิทธิ์เข้าถึงหน้านี้');
      window.location.href = redirectTo;
    }
  }
};

/**
 * เรียก API ฝั่ง Apps Script
 * ใช้ Content-Type: text/plain เพื่อเลี่ยง CORS preflight (ข้อจำกัดของ Apps Script Web App)
 */
async function apiCall(action, payload = {}) {
  const body = Object.assign({ action, token: Auth.getToken() }, payload);
  const res = await fetch(API_BASE_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body)
  });
  const rawText = await res.text();
  let json;
  try {
    json = JSON.parse(rawText);
  } catch (parseErr) {
    // เซิร์ฟเวอร์ตอบกลับมาไม่ใช่ JSON (มักเป็นหน้า HTML ของ Google เช่น หน้าขออนุญาต/ข้อผิดพลาดของ Apps Script)
    // ทำให้ผู้ใช้เห็นข้อความที่เข้าใจง่ายกว่า error แปลก ๆ จากการ parse โดยตรง
    console.error('apiCall: การตอบกลับไม่ใช่ JSON', rawText.slice(0, 300));
    throw new Error('เชื่อมต่อเซิร์ฟเวอร์ไม่สำเร็จ (ได้รับข้อมูลรูปแบบไม่ถูกต้อง) — Apps Script อาจยังไม่ได้ Deploy เวอร์ชันล่าสุด หรือมีปัญหาชั่วคราว กรุณาลองใหม่อีกครั้ง');
  }
  if (!json.ok) {
    if (String(json.error).includes('token หมดอายุ')) {
      Auth.clear();
      window.location.href = '/index.html';
    }
    throw new Error(json.error || 'เกิดข้อผิดพลาดไม่ทราบสาเหตุ');
  }
  return json.data;
}

/** อ่านไฟล์จาก <input type="file"> แล้วแปลงเป็น base64 (ไม่รวม prefix data:...;base64,) */
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

/** อัปโหลดไฟล์เอกสารแนบไปยัง Drive ผ่าน backend แล้วคืน {name, doc_type, url} */
async function uploadDocument(file, docType) {
  const base64 = await fileToBase64(file);
  return apiCall('uploadDocument', { base64, filename: file.name, mimeType: file.type, doc_type: docType });
}

/**
 * ซ่อนเมนู "สมัครสมาชิก" / "เปลี่ยนประเภทสมาชิก" ในแถบเมนูฝั่งสมาชิกให้เหมาะกับสถานะบัญชี:
 * - เป็นสมาชิกแล้ว (มีโปรไฟล์ผูกอยู่): ซ่อน "สมัครสมาชิก" (ไม่ต้องสมัครซ้ำ) และซ่อน "เปลี่ยนประเภทสมาชิก" ถ้าเป็นสามัญ/วิสามัญอยู่แล้ว
 * - ยังไม่เป็นสมาชิก (ยังไม่มีโปรไฟล์ผูก): ซ่อน "เปลี่ยนประเภทสมาชิก" (ยังไม่มีอะไรให้เปลี่ยน) คง "สมัครสมาชิก" ไว้
 * เรียกจากทุกหน้าในกลุ่ม Member Portal (member-dashboard.html, member-apply.html, member-type-change.html)
 */
async function applyMemberNavVisibility() {
  const navApply = document.getElementById('navApply');
  const navTypeChange = document.getElementById('navTypeChange');
  const navLicense = document.getElementById('navLicense');
  const navCard = document.getElementById('navCard');
  try {
    const profile = await apiCall('getMyProfile');
    if (navApply) navApply.style.display = 'none';
    const isOrg = profile.profile_kind === 'organization';
    if (navTypeChange && (isOrg || profile.member_type === 'สามัญ')) navTypeChange.style.display = 'none';
    if (isOrg) {
      // สมาชิกวิสามัญไม่มีใบอนุญาต/บัตรสมาชิกแบบบุคคล — ใช้ใบรับรององค์กรแทน (member-certificate.html)
      if (navLicense) navLicense.style.display = 'none';
      if (navCard) navCard.style.display = 'none';
    }
    return profile;
  } catch (err) {
    if (navTypeChange) navTypeChange.style.display = 'none';
    if (navLicense) navLicense.style.display = 'none';
    if (navCard) navCard.style.display = 'none';
    return null;
  }
}

// ---------- ระบบแจ้งเตือน/ยืนยันแบบสวยงาม (แทน alert/confirm/prompt ของเบราว์เซอร์) ----------

let toastContainer = null;
function ensureToastContainer() {
  if (!toastContainer) {
    toastContainer = document.createElement('div');
    toastContainer.className = 'toast-container';
    document.body.appendChild(toastContainer);
  }
  return toastContainer;
}

/** แสดงข้อความแจ้งเตือนแบบ toast มุมขวาบน หายไปเองใน 4 วินาที — ใช้แทน alert() */
function showToast(message, type) {
  const container = ensureToastContainer();
  const el = document.createElement('div');
  el.className = 'toast-item toast-' + (type || 'info');
  el.innerHTML = '<span class="toast-icon">' + (type === 'error' ? '✕' : type === 'success' ? '✓' : 'ℹ') + '</span><span class="toast-text"></span>';
  el.querySelector('.toast-text').textContent = message;
  container.appendChild(el);
  requestAnimationFrame(() => el.classList.add('toast-show'));
  setTimeout(() => {
    el.classList.remove('toast-show');
    setTimeout(() => el.remove(), 300);
  }, 4000);
}

function ensureModalRoot() {
  let root = document.getElementById('appModalRoot');
  if (!root) {
    root = document.createElement('div');
    root.id = 'appModalRoot';
    document.body.appendChild(root);
  }
  return root;
}

/** กล่องยืนยัน (ตกลง/ยกเลิก) แบบสวยงาม — คืนค่า Promise<boolean> ใช้แทน confirm() ต้องเรียกด้วย await */
function showConfirm(message, okLabel, cancelLabel) {
  return new Promise(resolve => {
    const root = ensureModalRoot();
    root.innerHTML = `
      <div class="modal-overlay">
        <div class="modal-box">
          <p class="modal-message"></p>
          <div class="modal-actions">
            <button class="btn secondary modal-cancel">${cancelLabel || 'ยกเลิก'}</button>
            <button class="btn modal-ok">${okLabel || 'ตกลง'}</button>
          </div>
        </div>
      </div>`;
    root.querySelector('.modal-message').textContent = message;
    const close = result => { root.innerHTML = ''; resolve(result); };
    root.querySelector('.modal-ok').onclick = () => close(true);
    root.querySelector('.modal-cancel').onclick = () => close(false);
    root.querySelector('.modal-overlay').onclick = e => { if (e.target.classList.contains('modal-overlay')) close(false); };
  });
}

/** กล่องกรอกข้อความ (ตกลง/ยกเลิก) แบบสวยงาม — คืนค่า Promise<string|null> ใช้แทน prompt() ต้องเรียกด้วย await */
function showPrompt(message, defaultValue) {
  return new Promise(resolve => {
    const root = ensureModalRoot();
    root.innerHTML = `
      <div class="modal-overlay">
        <div class="modal-box">
          <p class="modal-message"></p>
          <input type="text" class="modal-input">
          <div class="modal-actions">
            <button class="btn secondary modal-cancel">ยกเลิก</button>
            <button class="btn modal-ok">ตกลง</button>
          </div>
        </div>
      </div>`;
    root.querySelector('.modal-message').textContent = message;
    const input = root.querySelector('.modal-input');
    input.value = defaultValue || '';
    input.focus();
    const close = result => { root.innerHTML = ''; resolve(result); };
    root.querySelector('.modal-ok').onclick = () => close(input.value);
    root.querySelector('.modal-cancel').onclick = () => close(null);
    input.onkeydown = e => { if (e.key === 'Enter') close(input.value); if (e.key === 'Escape') close(null); };
    root.querySelector('.modal-overlay').onclick = e => { if (e.target.classList.contains('modal-overlay')) close(null); };
  });
}

// ---------- ที่อยู่แบบ Dropdown อัตโนมัติ (จังหวัด/อำเภอ/ตำบล/รหัสไปรษณีย์) ----------
// ใช้ฐานข้อมูลสาธารณะ kongvut/thai-province-data (จังหวัด 77 + อำเภอ + ตำบลครบ) ผ่าน GitHub Raw CDN

const THAI_ADDRESS_DATA_URL = 'https://raw.githubusercontent.com/kongvut/thai-province-data/master/api_province_with_amphure_tambon.json';
let thaiAddressDataCache = null;

async function loadThaiAddressData() {
  if (thaiAddressDataCache) return thaiAddressDataCache;
  const resp = await fetch(THAI_ADDRESS_DATA_URL);
  thaiAddressDataCache = await resp.json();
  return thaiAddressDataCache;
}

function thaiAddrFillDistricts(data, districtSel, provinceName, keepDistrict) {
  const province = data.find(p => p.name_th === provinceName);
  const districts = province ? province.amphure : [];
  districtSel.innerHTML = '<option value="">— เลือกเขต/อำเภอ —</option>' +
    districts.map(d => `<option value="${d.name_th}">${d.name_th}</option>`).join('');
  if (keepDistrict && districts.some(d => d.name_th === keepDistrict)) districtSel.value = keepDistrict;
}

function thaiAddrFillSubdistricts(data, subdistrictSel, zipInput, provinceName, districtName, keepSubdistrict) {
  const province = data.find(p => p.name_th === provinceName);
  const district = province ? province.amphure.find(d => d.name_th === districtName) : null;
  const subs = district ? district.tambon : [];
  subdistrictSel.innerHTML = '<option value="">— เลือกแขวง/ตำบล —</option>' +
    subs.map(s => `<option value="${s.name_th}" data-zip="${s.zip_code || ''}">${s.name_th}</option>`).join('');
  if (keepSubdistrict && subs.some(s => s.name_th === keepSubdistrict)) {
    subdistrictSel.value = keepSubdistrict;
    if (zipInput) {
      const match = subs.find(s => s.name_th === keepSubdistrict);
      if (match) zipInput.value = match.zip_code || '';
    }
  }
}

/** ตั้งค่าจังหวัด/อำเภอ/ตำบล/รหัสไปรษณีย์ ให้ dropdown ชุดหนึ่งโดยตรง (ใช้ตอนโหลดข้อมูลเดิมมาแก้ไข หรือคัดลอกจากที่อยู่อื่น) */
async function setThaiAddressCascade(prefix, values) {
  const data = await loadThaiAddressData();
  const provinceSel = document.getElementById(prefix + '_province');
  const districtSel = document.getElementById(prefix + '_district');
  const subdistrictSel = document.getElementById(prefix + '_subdistrict');
  const zipInput = document.getElementById(prefix + '_zipcode');
  if (!provinceSel || !districtSel || !subdistrictSel) return;
  provinceSel.value = values.province || '';
  thaiAddrFillDistricts(data, districtSel, values.province, values.district);
  thaiAddrFillSubdistricts(data, subdistrictSel, zipInput, values.province, values.district, values.subdistrict);
  if (zipInput && values.zipcode && !subdistrictSel.value) zipInput.value = values.zipcode;
}

/** ตั้งค่าเริ่มต้น + ผูก event ให้ dropdown ที่อยู่ชุดหนึ่ง (province/district/subdistrict/zipcode) ทำงานแบบ cascade อัตโนมัติ
 * prefix: ต้องมี element id = `${prefix}_province`, `${prefix}_district`, `${prefix}_subdistrict`, `${prefix}_zipcode`
 * initial: { province, district, subdistrict, zipcode } (ใส่หรือไม่ใส่ก็ได้ ถ้าเป็นฟอร์มใหม่ไม่ต้องใส่)
 */
async function initThaiAddressCascade(prefix, initial) {
  let data;
  try {
    data = await loadThaiAddressData();
  } catch (e) {
    console.error('โหลดฐานข้อมูลที่อยู่ไม่สำเร็จ', e);
    return;
  }
  const provinceSel = document.getElementById(prefix + '_province');
  const districtSel = document.getElementById(prefix + '_district');
  const subdistrictSel = document.getElementById(prefix + '_subdistrict');
  const zipInput = document.getElementById(prefix + '_zipcode');
  if (!provinceSel || !districtSel || !subdistrictSel) return;

  provinceSel.innerHTML = '<option value="">— เลือกจังหวัด —</option>' +
    data.map(p => `<option value="${p.name_th}">${p.name_th}</option>`).join('');
  districtSel.innerHTML = '<option value="">— เลือกจังหวัดก่อน —</option>';
  subdistrictSel.innerHTML = '<option value="">— เลือกเขต/อำเภอก่อน —</option>';

  provinceSel.onchange = () => {
    thaiAddrFillDistricts(data, districtSel, provinceSel.value);
    subdistrictSel.innerHTML = '<option value="">— เลือกเขต/อำเภอก่อน —</option>';
    if (zipInput) zipInput.value = '';
  };
  districtSel.onchange = () => {
    thaiAddrFillSubdistricts(data, subdistrictSel, zipInput, provinceSel.value, districtSel.value);
    if (zipInput) zipInput.value = '';
  };
  subdistrictSel.onchange = () => {
    const opt = subdistrictSel.selectedOptions[0];
    if (zipInput) zipInput.value = opt ? (opt.dataset.zip || '') : '';
  };

  if (initial && initial.province) {
    provinceSel.value = initial.province;
    thaiAddrFillDistricts(data, districtSel, initial.province, initial.district);
    thaiAddrFillSubdistricts(data, subdistrictSel, zipInput, initial.province, initial.district, initial.subdistrict);
  }
}
const AFFILIATION_DATA = {"กระทรวงกลาโหม": ["สำนักงานรัฐมนตรีกระทรวงกลาโหม", "สำนักงานปลัดกระทรวงกลาโหม", "สำนักงานเลขานุการสำนักงานปลัดกระทรวงกลาโหม", "สำนักนโยบายและแผนกลาโหม", "กรมเสมียนตรา", "สำนักงบประมาณกลาโหม", "กรมพระธรรมนูญ", "ศูนย์การอุตสาหกรรมป้องกันประเทศและพลังงานทหาร", "กรมการเงินกลาโหม"], "กระทรวงการคลัง": ["สำนักงานรัฐมนตรีกระทรวงการคลัง", "สำนักงานปลัดกระทรวงการคลัง", "กรมธนารักษ์", "กรมบัญชีกลาง", "กรมศุลกากร", "กรมสรรพสามิต", "กรมสรรพากร", "สำนักงานคณะกรรมการนโยบายรัฐวิสาหกิจ", "สำนักงานบริหารหนี้สาธารณะ"], "กระทรวงการต่างประเทศ": ["สำนักงานรัฐมนตรีกระทรวงการต่างประเทศ", "สำนักงานปลัดกระทรวงการต่างประเทศ", "กรมพิธีการทูต", "กรมยุโรป", "กรมเศรษฐกิจระหว่างประเทศ", "กรมสนธิสัญญาและกฎหมาย", "กรมสารนิเทศ", "กรมองค์การระหว่างประเทศ", "กรมเอเชียตะวันออก"], "กระทรวงการท่องเที่ยวและกีฬา": ["สำนักงานรัฐมนตรีกระทรวงการท่องเที่ยวและกีฬา", "สำนักงานปลัดกระทรวงการท่องเที่ยวและกีฬา", "กรมการท่องเที่ยว", "กรมพลศึกษา", "การท่องเที่ยวแห่งประเทศไทย (ททท.)", "การกีฬาแห่งประเทศไทย (กกท.)", "องค์การบริหารการพัฒนาพื้นที่พิเศษเพื่อการท่องเที่ยวอย่างยั่งยืน (อพท. / องค์การมหาชน)"], "กระทรวงการพัฒนาสังคมและความมั่นคงของมนุษย์": ["สำนักงานรัฐมนตรีกระทรวงการพัฒนาสังคมและความมั่นคงของมนุษย์", "กรมกิจการเด็กและเยาวชน", "กรมกิจการผู้สูงอายุ", "กรมกิจการสตรีและสถาบันครอบครัว", "กรมพัฒนาสังคมและสวัสดิการ", "กรมส่งเสริมและพัฒนาคุณภาพชีวิตคนพิการ", "สำนักงานปลัดกระทรวงการพัฒนาสังคมและความมั่นคงของมนุษย์", "การเคหะแห่งชาติ", "สถาบันพัฒนาองค์กรชุมชน (องค์การมหาชน)"], "กระทรวงเกษตรและสหกรณ์": ["สำนักงานรัฐมนตรีกระทรวงเกษตรและสหกรณ์", "สำนักงานปลัดกระทรวงเกษตรและสหกรณ์", "กรมส่งเสริมการเกษตร", "กรมส่งเสริมสหกรณ์", "กรมการข้าว", "กรมประมง", "กรมปศุสัตว์", "กรมหม่อนไหม", "กรมชลประทาน"], "กระทรวงคมนาคม": ["สำนักงานรัฐมนตรีกระทรวงคมนาคม", "สำนักงานปลัดกระทรวงคมนาคม", "กรมการขนส่งทางบก", "กรมท่าอากาศยาน (กรมการบินพลเรือน)", "กรมเจ้าท่า", "กรมทางหลวง", "กรมทางหลวงชนบท", "สำนักงานนโยบายและแผนการขนส่งและจราจร", "สำนักงานการบินพลเรือนแห่งประเทศไทย"], "กระทรวงทรัพยากรธรรมชาติและสิ่งแวดล้อม": ["สำนักงานรัฐมนตรีกระทรวงทรัพยากรธรรมชาติและสิ่งแวดล้อม", "สำนักงานปลัดกระทรวงทรัพยากรธรรมชาติและสิ่งแวดล้อม", "สำนักงานนโยบายและแผนทรัพยากรธรรมชาติและสิ่งแวดล้อม", "กรมควบคุมมลพิษ", "กรมส่งเสริมคุณภาพสิ่งแวดล้อม", "กรมป่าไม้", "กรมอุทยานแห่งชาติ สัตว์ป่า และพันธุ์พืช", "กรมทรัพยากรธรณี", "กรมทรัพยากรน้ำ"], "กระทรวงดิจิทัลเพื่อเศรษฐกิจและสังคม": ["สำนักงานรัฐมนตรีกระทรวงดิจิทัลเพื่อเศรษฐกิจและสังคม", "สำนักงานปลัดกระทรวงดิจิทัลเพื่อเศรษฐกิจและสังคม", "กรมอุตุนิยมวิทยา", "สำนักงานสถิติแห่งชาติ", "สำนักงานคณะกรรมการดิจิทัลเพื่อเศรษฐกิจและสังคมแห่งชาติ", "สำนักงานส่งเสริมเศรษฐกิจดิจิทัล"], "กระทรวงพลังงาน": ["สำนักงานรัฐมนตรีกระทรวงพลังงาน", "สำนักงานปลัดกระทรวงพลังงาน", "กรมเชื้อเพลิงธรรมชาติ", "กรมธุรกิจพลังงาน", "สำนักงานนโยบายและแผนพลังงาน (สนพ.)"], "กระทรวงพาณิชย์": ["สำนักงานรัฐมนตรีกระทรวงพาณิชย์", "สำนักงานปลัดกระทรวงพาณิชย์", "กรมการค้าต่างประเทศ กระทรวงพาณิชย์", "กรมการค้าภายใน", "กรมพัฒนาธุรกิจการค้า", "กรมเจรจาการค้าระหว่างประเทศ", "กรมส่งเสริมการค้าระหว่างประเทศ กระทรวงพาณิชย์", "กรมทรัพย์สินทางปัญญา", "สำนักงานนโยบายยุทธศาสตร์การค้า"], "กระทรวงมหาดไทย": ["สำนักงานรัฐมนตรีกระทรวงมหาดไทย", "สำนักงานปลัดกระทรวงมหาดไทย", "กรมการปกครอง", "กรมการพัฒนาชุมชน", "กรมป้องกันและบรรเทาสาธารณภัย กระทรวงมหาดไทย", "กรมที่ดิน", "กรมโยธาธิการและผังเมือง", "กรมส่งเสริมการปกครองท้องถิ่น", "สำนักงานเขต"], "กระทรวงยุติธรรม": ["สำนักงานรัฐมนตรีกระทรวงยุติธรรม", "สำนักงานปลัดกระทรวงยุติธรรม", "กรมคุ้มครองสิทธิและเสรีภาพ", "กรมบังคับคดี", "กรมคุมประพฤติ", "กรมพินิจและคุ้มครองเด็กและเยาวชน", "กรมราชทัณฑ์", "กรมสอบสวนคดีพิเศษ", "สำนักงานกิจการยุติธรรม"], "กระทรวงแรงงาน": ["สำนักงานรัฐมนตรีกระทรวงแรงงาน", "สำนักงานปลัดกระทรวงแรงงาน", "กรมพัฒนาฝีมือแรงงาน", "กรมการจัดหางาน", "กรมสวัสดิการและคุ้มครองแรงงาน", "สำนักงานประกันสังคม"], "กระทรวงวัฒนธรรม": ["สำนักงานรัฐมนตรีกระทรวงวัฒนธรรม", "สำนักงานปลัดกระทรวงวัฒนธรรม", "กรมการศาสนา", "กรมศิลปากร", "กรมส่งเสริมวัฒนธรรม", "สำนักงานศิลปวัฒนธรรมร่วมสมัย", "สถาบันบัณฑิตพัฒนศิลป์"], "กระทรวงการอุดมศึกษา วิทยาศาสตร์ วิจัยและนวัตกรรม": ["สำนักงานรัฐมนตรีกระทรวงการอุดมศึกษา วิทยาศาสตร์ วิจัยและนวัตกรรม", "สำนักงานปลัดกระทรวงการอุดมศึกษา วิทยาศาสตร์ วิจัยและนวัตกรรม", "สำนักงานคณะกรรมการการอุดมศึกษา", "กรมวิทยาศาสตร์บริการ", "สำนักงานปรมาณูเพื่อสันติ", "สถาบันมาตรวิทยาแห่งชาติ", "สำนักงานพัฒนาวิทยาศาสตร์และเทคโนโลยีแห่งชาติ", "สำนักงานสภานโยบายการอุดมศึกษา วิทยาศาสตร์ วิจัยและนวัตกรรมแห่งชาติ", "สำนักงานคณะกรรมการส่งเสริมวิทยาศาสตร์ วิจัยและนวัตกรรม"], "กระทรวงศึกษาธิการ": ["สำนักงานรัฐมนตรีกระทรวงศึกษาธิการ", "สำนักงานปลัดกระทรวงศึกษาธิการ", "สำนักงานเลขาธิการสภาการศึกษา", "สำนักงานคณะกรรมการการศึกษาขั้นพื้นฐาน", "สำนักงานคณะกรรมการการอาชีวศึกษา", "นิติบุคคลในกำกับของกระทรวงศึกษาธิการ", "สำนักงานคณะกรรมการส่งเสริมสวัสดิการและสวัสดิภาพครูและบุคลากรทางการศึกษา", "สถาบันส่งเสริมการสอนวิทยาศาสตร์และเทคโนโลยี", "สำนักงานลูกเสือแห่งชาติ"], "กระทรวงสาธารณสุข": ["สำนักงานปลัดกระทรวงสาธารณสุข", "กรมการแพทย์", "กรมการแพทย์แผนไทยและการแพทย์ทางเลือก", "กรมสุขภาพจิต", "กรมควบคุมโรค", "กรมอนามัย", "กรมสนับสนุนบริการสุขภาพ", "กรมวิทยาศาสตร์การแพทย์", "สำนักงานคณะกรรมการอาหารและยา"], "กระทรวงอุตสาหกรรม": ["สำนักงานรัฐมนตรีกระทรวงอุตสาหกรรม", "สำนักงานปลัดกระทรวงอุตสาหกรรม", "สำนักงานเศรษฐกิจอุตสาหกรรม", "กรมอุตสาหกรรมพื้นฐานและการเหมืองแร่", "กรมโรงงานอุตสาหกรรม", "สำนักงานคณะกรรมการอ้อยและน้ำตาลทราย", "กรมส่งเสริมอุตสาหกรรม", "สำนักงานมาตรฐานผลิตภัณฑ์อุตสาหกรรม", "สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม"], "กรุงเทพมหานคร": ["สำนักงานเลขานุการสภากรุงเทพมหานคร", "สำนักงานคณะกรรมการข้าราชการกรุงเทพมหานคร", "สำนักงานกฎหมายและคดี", "สำนักปลัดกรุงเทพมหานคร", "สำนักป้องกันและบรรเทาสาธารณภัย", "สำนักการคลัง", "สำนักการโยธา", "สำนักการจราจรและขนส่ง", "สำนักผังเมือง", "สำนักกการระบายน้ำ", "สำนักเทศกิจ", "สำนักพัฒนาสังคม", "สำนักวัฒนธรรม กีฬา และการท่องเที่ยว", "สำนักการศึกษา", "สำนักการแพทย์", "สำนักอนามัย", "สำนักงานงบประมาณ", "สำนักสิ่งแวดล้อม", "สำนักยุทธศาสตร์และประเมินผล", "สถาบันพัฒาข้าราชการกรุงเทพมหานคร"], "สภากาชาดไทย": [], "องค์กรไม่แสวงหาผลกำไร (NGO)": ["องค์กรไม่แสวงหาผลกำไร (NGO)"], "เกษียณ /อิสระ": ["เกษียณ/อิสระ"], "สำนักงานตำรวจแห่งชาติ": [], "สำนักงานนายกรัฐมนตรี": [], "องค์การระหว่างประเทศ": ["องค์การระหว่างประเทศ"], "อื่นๆ": ["ระบุ"]};

/** ตั้งค่า dropdown สังกัด (กระทรวง) + กรม แบบ cascade จากไฟล์ข้อมูลจริงที่สภาฯ ให้มา
 * ต้องมี element id = `${prefix}_affiliation` (สังกัด) และ `${prefix}_agency` (กรม)
 */
/** ตั้งค่า dropdown "สังกัด" (กระทรวง) + ช่องกรอก "หน่วยงาน" (กรม) พร้อมคำแนะนำอัตโนมัติจากไฟล์ข้อมูลจริงที่สภาฯ ให้มา
 * สังกัด = dropdown (ให้ข้อมูลสม่ำเสมอ ใช้ทำรายงานได้แม่นยำ) มีตัวเลือก "ไม่สังกัดหน่วยงานราชการ/เอกชน" สำหรับผู้ที่ไม่ได้ทำงานราชการ
 * หน่วยงาน = ยังเป็นช่องพิมพ์เองอิสระเหมือนเดิม (ไม่บังคับเลือกจากรายการ เพราะสมาชิกจำนวนมากทำงานหน่วยงานเอกชน/มูลนิธิที่ไม่อยู่ในรายชื่อราชการ)
 * แต่ถ้าเลือกสังกัดเป็นกระทรวงจริง จะขึ้นคำแนะนำรายชื่อกรมในสังกัดนั้นให้เลือกอัตโนมัติผ่าน <datalist> (ยังพิมพ์เองทับได้เสมอ)
 * ต้องมี element id = `${prefix}_affiliation` (select) และ `${prefix}_agency` (input พร้อม list="${prefix}_agency_list")
 */
function initAffiliationCascade(prefix, initial) {
  const affSel = document.getElementById(prefix + '_affiliation');
  const agyInput = document.getElementById(prefix + '_agency');
  if (!affSel) return;
  const ministries = Object.keys(AFFILIATION_DATA);
  affSel.innerHTML = '<option value="">— เลือกสังกัด —</option>' +
    ministries.map(m => `<option value="${m}">${m}</option>`).join('') +
    '<option value="ไม่สังกัดหน่วยงานราชการ/เอกชน">ไม่สังกัดหน่วยงานราชการ (เอกชน/มูลนิธิ/อื่นๆ)</option>';

  let datalist = document.getElementById(prefix + '_agency_list');
  if (!datalist && agyInput) {
    datalist = document.createElement('datalist');
    datalist.id = prefix + '_agency_list';
    agyInput.setAttribute('list', datalist.id);
    agyInput.parentNode.appendChild(datalist);
  }

  function updateAgencySuggestions(ministry) {
    if (!datalist) return;
    const depts = AFFILIATION_DATA[ministry] || [];
    datalist.innerHTML = depts.map(d => `<option value="${d}"></option>`).join('');
  }

  affSel.onchange = () => updateAgencySuggestions(affSel.value);

  if (initial && initial.affiliation) {
    affSel.value = ministries.includes(initial.affiliation) ? initial.affiliation : 'ไม่สังกัดหน่วยงานราชการ/เอกชน';
    updateAgencySuggestions(affSel.value);
  }
}

// ---------- ระบบลายเซ็น/ประทับตราดิจิทัล (Digital Signature Pad) ----------

/** สร้างกระดานลายเซ็นดิจิทัลบน <canvas> — วาดด้วยเมาส์/นิ้วได้ รองรับทั้งเดสก์ท็อปและมือถือ
 * canvasId: id ของ <canvas> ที่เตรียมไว้ในหน้า (ควรตั้ง width/height เป็นพิกเซลจริงไว้ล่วงหน้า)
 * คืนค่า object ที่มี: isEmpty(), clear(), toFile(filename) — ใช้ toFile() แล้วส่งเข้า uploadDocument() ได้เลย (ระบบเดียวกับอัปโหลดเอกสารทั่วไป)
 */
function createSignaturePad(canvasId) {
  const canvas = document.getElementById(canvasId);
  const ctx = canvas.getContext('2d');
  let drawing = false;
  let hasSignature = false;

  function getPos(e) {
    const rect = canvas.getBoundingClientRect();
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    const clientY = e.touches ? e.touches[0].clientY : e.clientY;
    return { x: (clientX - rect.left) * (canvas.width / rect.width), y: (clientY - rect.top) * (canvas.height / rect.height) };
  }
  function start(e) {
    drawing = true; hasSignature = true;
    const pos = getPos(e);
    ctx.beginPath();
    ctx.moveTo(pos.x, pos.y);
    e.preventDefault();
  }
  function move(e) {
    if (!drawing) return;
    const pos = getPos(e);
    ctx.strokeStyle = '#1a1a2e';
    ctx.lineWidth = 2.5;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.lineTo(pos.x, pos.y);
    ctx.stroke();
    e.preventDefault();
  }
  function end() { drawing = false; }

  canvas.addEventListener('mousedown', start);
  canvas.addEventListener('mousemove', move);
  window.addEventListener('mouseup', end);
  canvas.addEventListener('touchstart', start, { passive: false });
  canvas.addEventListener('touchmove', move, { passive: false });
  canvas.addEventListener('touchend', end);

  function clear() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    hasSignature = false;
  }
  clear();

  return {
    isEmpty: () => !hasSignature,
    clear: clear,
    toFile: (filename) => new Promise(resolve => {
      canvas.toBlob(blob => resolve(new File([blob], filename || 'signature.png', { type: 'image/png' })), 'image/png');
    })
  };
}

// ---------- ตรวจสอบว่า Deploy เวอร์ชันล่าสุดของ backend แล้วหรือยัง ----------
// ต้องตรงกับ BACKEND_VERSION ใน Code.gs — อัปเดตทุกครั้งที่ส่งมอบไฟล์ Code.gs ชุดใหม่
// ป้องกันปัญหา "อัปโหลดไฟล์เว็บแล้วแต่ลืม Deploy Apps Script ใหม่" ซึ่งทำให้ฟีเจอร์ใหม่ไม่ทำงานโดยไม่รู้ตัว
const EXPECTED_BACKEND_VERSION = '2026-09-11-boolean-field-fix';

async function checkBackendVersionAndWarn() {
  try {
    const res = await apiCall('getBackendVersion');
    if (res.version !== EXPECTED_BACKEND_VERSION) {
      showBackendVersionWarning('ระบบหลังบ้าน (Apps Script) ยังไม่ได้ Deploy เป็นเวอร์ชันล่าสุด — ฟีเจอร์บางอย่างอาจยังไม่ทำงานตามที่คาดไว้ กรุณาวางไฟล์ Code.gs ชุดล่าสุดทับแล้วกด Deploy > Manage deployments > แก้ไข > Version: New version > Deploy');
    }
  } catch (err) {
    // action getBackendVersion ไม่รู้จักเลย = backend เก่ามากๆ (deploy มาก่อนที่จะมีฟีเจอร์ตรวจสอบเวอร์ชันนี้)
    showBackendVersionWarning('ไม่สามารถตรวจสอบเวอร์ชันระบบหลังบ้านได้ — ระบบหลังบ้าน (Apps Script) อาจยังเป็นเวอร์ชันเก่ามาก กรุณาวางไฟล์ Code.gs ชุดล่าสุดทับแล้วกด Deploy ใหม่');
  }
}

function showBackendVersionWarning(message) {
  const banner = document.createElement('div');
  banner.className = 'backend-version-warning';
  banner.innerHTML = '<b>⚠ แจ้งเตือนผู้ดูแลระบบ:</b> ' + message;
  document.body.insertBefore(banner, document.body.firstChild);
}

/** เปลี่ยนรหัสผ่านของตนเอง — ใช้ modal ถาม-ตอบแทน prompt ของเบราว์เซอร์ ใช้ร่วมกันได้ทั้งสมาชิกและเจ้าหน้าที่ */
async function changeMyPasswordFlow() {
  const current = await showPrompt('รหัสผ่านปัจจุบัน:');
  if (!current) return;
  const next = await showPrompt('รหัสผ่านใหม่ (อย่างน้อย 8 ตัวอักษร มีทั้งตัวอักษรและตัวเลข):');
  if (!next) return;
  try {
    await apiCall('changeMyPassword', { current_password: current, new_password: next });
    showToast('เปลี่ยนรหัสผ่านสำเร็จ', 'success');
  } catch (err) {
    showToast('เปลี่ยนรหัสผ่านไม่สำเร็จ: ' + err.message, 'error');
  }
}

/** encode ค่าข้อความก่อนแสดงผลเป็น HTML — ป้องกัน XSS เวลาแสดงข้อมูลที่มาจากฐานข้อมูล (เช่น ชื่อสมาชิก) บนหน้าเว็บสาธารณะที่ไม่ต้อง login
 * (OWASP Top 10 — A03 Injection / Cross-Site Scripting) ใช้กับข้อความที่ไม่ควรตีความเป็น HTML tag
 */
function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** แปลงลิงก์ไฟล์ Google Drive (แบบ "view" ที่ file.getUrl() คืนมา) ให้เป็นลิงก์รูปภาพที่ฝังแสดงผลตรงได้ (<img src>)
 * ลิงก์ Drive ปกติ (drive.google.com/file/d/XXX/view) เป็นหน้าตัวดูไฟล์ ไม่ใช่ไฟล์รูปดิบ ใส่ใน <img src> ตรงๆ จะไม่ขึ้นรูป
 * ต้องแปลงเป็นรูปแบบ lh3.googleusercontent.com ก่อนเสมอเวลาจะแสดงเป็นรูปภาพ (ไฟล์ต้องแชร์แบบ "ทุกคนที่มีลิงก์" ไว้แล้ว)
 */
function toEmbeddableImageUrl(url) {
  if (!url) return '';
  const m = url.match(/[-\w]{25,}/);
  if (!m) return url;
  return 'https://lh3.googleusercontent.com/d/' + m[0];
}

/** ตรวจสอบค่า boolean ที่มาจากฐานข้อมูล (ตรงกับ isTrue_ ฝั่ง backend) — Google Sheets อาจแปลง "true"/"false" ที่เก็บไว้
 * ให้กลายเป็นชนิด Boolean จริงโดยอัตโนมัติ ทำให้เทียบด้วย === 'true' ตรงๆ พลาดได้ ฟังก์ชันนี้รองรับทั้งสองรูปแบบ
 */
function isTrueVal(val) {
  return val === true || val === 'true' || val === 'TRUE' || val === 1;
}

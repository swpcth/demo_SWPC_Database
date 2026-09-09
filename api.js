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
  const json = await res.json();
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

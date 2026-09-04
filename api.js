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
  try {
    const profile = await apiCall('getMyProfile');
    if (navApply) navApply.style.display = 'none';
    if (navTypeChange && (profile.profile_kind === 'organization' || profile.member_type === 'สามัญ')) {
      navTypeChange.style.display = 'none';
    }
    return profile;
  } catch (err) {
    if (navTypeChange) navTypeChange.style.display = 'none';
    return null;
  }
}

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

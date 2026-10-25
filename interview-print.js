// พิมพ์/ส่งออกตารางสอบสัมภาษณ์ — รูปแบบเดียวกับระบบให้คะแนนสอบสัมภาษณ์เดิม
//  - ivPrint.scoreSheet(opts): "ตารางให้คะแนนของกรรมการ" (A4 แนวนอน) ออกเป็น Word (.doc) หรือสั่งพิมพ์
//  - ivPrint.summary(opts): "ตารางสรุปผลการประเมินภาพรวม" (A4 แนวตั้ง) ออกเป็น Excel (.xlsx) หรือสั่งพิมพ์
// ใช้ร่วมกันระหว่างหน้ากรรมการ (interviewer.html) และหน้าแอดมิน (admin-interviews.html)
(function () {
  const MONTHS = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
  const DAYS = ['วันอาทิตย์', 'วันจันทร์', 'วันอังคาร', 'วันพุธ', 'วันพฤหัสบดี', 'วันศุกร์', 'วันเสาร์'];
  const esc = s => String(s === undefined || s === null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  function formatDate(iso) {
    if (!iso) return 'ยังไม่ได้ระบุวันที่';
    const d = new Date(iso);
    if (isNaN(d.getTime())) return String(iso);
    return DAYS[d.getDay()] + 'ที่ ' + d.getDate() + ' ' + MONTHS[d.getMonth()] + ' พ.ศ. ' + (d.getFullYear() + 543);
  }

  const CRITERIA = [
    { key: 'skill', label: '1. ทักษะทางสังคมสงเคราะห์ในการปฏิบัติงาน', max: 30 },
    { key: 'attitude', label: '2. ทัศนคติทางสังคมสงเคราะห์', max: 30 },
    { key: 'writing', label: '3. การเขียน', max: 20 },
    { key: 'presentation', label: '4. การนำเสนอ', max: 20 }
  ];
  const TITLE = 'การสอบสัมภาษณ์เพื่อประเมินทักษะและทัศนคติทางสังคมสงเคราะห์';

  const STYLE = `
    #ivPrintRoot { position: fixed; inset: 0; z-index: 99999; background: #fff; overflow: auto; font-family: 'Sarabun','TH Sarabun New',sans-serif; color: #000; }
    #ivPrintRoot .ivp-bar { position: fixed; top: 0; left: 0; right: 0; background: #1e293b; color: #fff; padding: 12px 16px; display: flex; justify-content: space-between; align-items: center; gap: 8px; flex-wrap: wrap; z-index: 2; }
    #ivPrintRoot .ivp-bar button { border: 0; border-radius: 8px; padding: 8px 16px; color: #fff; font-size: 14px; font-weight: 500; cursor: pointer; font-family: inherit; }
    #ivPrintRoot .ivp-body { padding: 84px 32px 60px; max-width: 1100px; margin: 0 auto; }
    #ivPrintRoot .ivp-head { text-align: center; font-weight: bold; line-height: 1.7; margin-bottom: 20px; }
    #ivPrintRoot table { width: 100%; border-collapse: collapse; text-align: center; margin-bottom: 24px; }
    #ivPrintRoot th, #ivPrintRoot td { border: 1px solid #000; padding: 8px; }
    #ivPrintRoot td.l, #ivPrintRoot th.l { text-align: left; }
    #ivPrintRoot .dots { margin: 8px 0 40px; line-height: 2.4; font-size: 14px; }
    @media print {
      body > *:not(#ivPrintRoot) { display: none !important; }
      #ivPrintRoot { position: static !important; overflow: visible !important; }
      #ivPrintRoot .ivp-bar { display: none !important; }
      #ivPrintRoot .ivp-body { padding: 0 !important; max-width: none !important; }
    }`;

  function loadXlsx() {
    if (window.XLSX) return Promise.resolve();
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
      s.onload = resolve; s.onerror = () => reject(new Error('โหลดตัวส่งออก Excel ไม่สำเร็จ'));
      document.head.appendChild(s);
    });
  }

  function open(o) {
    close();
    const style = document.createElement('style');
    style.id = 'ivPrintStyle';
    style.textContent = STYLE + `@media print { @page { size: A4 ${o.landscape ? 'landscape' : 'portrait'}; margin: 15mm; } }`;
    document.head.appendChild(style);
    const root = document.createElement('div');
    root.id = 'ivPrintRoot';
    root.innerHTML = `<div class="ivp-bar"><b>${esc(o.barTitle)}</b><div style="display:flex; gap:8px; flex-wrap:wrap">
      <button style="background:#475569" data-a="close">ปิด / กลับ</button>
      <button style="background:${o.exportKind === 'excel' ? '#059669' : '#2563eb'}" data-a="export">⬇ ${o.exportKind === 'excel' ? 'ออกเป็น Excel' : 'ออกเป็น Word'}</button>
      <button style="background:#4f46e5" data-a="print">🖨 สั่งพิมพ์</button></div></div>
      <div class="ivp-body"><div id="ivPrintArea">${o.html}</div></div>`;
    document.body.appendChild(root);
    root.querySelector('[data-a=close]').onclick = close;
    root.querySelector('[data-a=print]').onclick = () => window.print();
    root.querySelector('[data-a=export]').onclick = async () => {
      try { await (o.exportKind === 'excel' ? exportExcel(o.fileName) : exportWord(o.fileName)); } catch (e) { alert(e.message); }
    };
    document.body.style.overflow = 'hidden';
  }

  function close() {
    const r = document.getElementById('ivPrintRoot'); if (r) r.remove();
    const s = document.getElementById('ivPrintStyle'); if (s) s.remove();
    document.body.style.overflow = '';
  }

  async function exportExcel(fileName) {
    await loadXlsx();
    const table = document.getElementById('summary-table');
    const wb = XLSX.utils.table_to_book(table, { sheet: 'Summary' });
    XLSX.writeFile(wb, fileName + '.xlsx');
  }

  function exportWord(fileName) {
    const header = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head><meta charset='utf-8'><title>ตารางให้คะแนน</title><style>
        @page WordSection1 { size: 841.9pt 595.3pt; mso-page-orientation: landscape; margin: 0.5in; }
        div.WordSection1 { page: WordSection1; }
        body, table, td, th, div, span { font-family: 'Sarabun', 'TH SarabunPSK', sans-serif !important; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; font-size: 14px; }
        th, td { border: 1px solid black; padding: 5px; text-align: center; }
        .l { text-align: left !important; }
        .ivp-head { text-align: center; font-size: 16px; font-weight: bold; margin-bottom: 15px; }
      </style></head><body><div class="WordSection1">`;
    const html = header + document.getElementById('ivPrintArea').innerHTML + '</div></body></html>';
    const blob = new Blob(['﻿', html], { type: 'application/msword' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = fileName + '.doc';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
  }

  /** ตารางให้คะแนนของกรรมการ (1 ท่าน) — o: { round_label, exam_date, time_text, group_no, interviewerName, slot, candidates:[{name, absent, score, remark}] } */
  function scoreSheet(o) {
    const cands = o.candidates || [];
    const rows = cands.length ? cands.map((c, i) => {
      if (c.absent) return `<tr><td>${i + 1}</td><td class="l">${esc(c.name)}</td><td colspan="5"><b>ไม่เข้าสอบ</b></td><td class="l" style="font-size:12px">${esc(c.remark)}</td></tr>`;
      const s = c.score || {};
      const v = k => (s[k] !== undefined && s[k] !== null ? esc(s[k]) : '');
      return `<tr><td>${i + 1}</td><td class="l">${esc(c.name)}</td><td>${v('skill')}</td><td>${v('attitude')}</td><td>${v('writing')}</td><td>${v('presentation')}</td><td><b>${v('total')}</b></td><td class="l" style="font-size:12px">${esc(c.remark)}</td></tr>`;
    }).join('') : '<tr><td colspan="8">ไม่มีรายชื่อผู้เข้าสอบ</td></tr>';
    const blank = Array.from({ length: Math.max(0, 8 - cands.length) }).map(() => '<tr>' + '<td style="padding:14px">&nbsp;</td>'.repeat(8) + '</tr>').join('');
    const html = `<div class="ivp-head">
        <div>ตารางให้คะแนนของกรรมการ ${TITLE} (ครั้งที่ ${esc(o.round_label)})</div>
        <div>กลุ่มที่ ${esc(o.group_no)} วันสอบ ${esc(formatDate(o.exam_date))}${o.time_text ? ' เวลาสอบ ' + esc(o.time_text) : ''}</div></div>
      <table><thead>
        <tr><th rowspan="2" style="width:5%">ที่</th><th rowspan="2" class="l" style="width:20%; text-align:center">รายชื่อผู้สอบ</th><th colspan="4">เกณฑ์การพิจารณาให้คะแนน</th><th rowspan="2" style="width:10%">รวมคะแนนทั้งหมด</th><th rowspan="2" style="width:15%">ข้อสังเกตจากกรรมการ</th></tr>
        <tr>${CRITERIA.map(c => `<th style="font-weight:normal; font-size:13px; width:12.5%">${esc(c.label)}<br>(${c.max} คะแนน)</th>`).join('')}</tr></thead>
      <tbody>${rows}${blank}</tbody></table>
      <div class="dots">ข้อเสนอแนะเพิ่มเติม ${'.'.repeat(150)}<br>${'.'.repeat(190)}<br>${'.'.repeat(190)}</div>
      <div style="display:flex; justify-content:flex-end; margin-right:48px; margin-top:40px"><div style="width:320px; text-align:center">
        <div style="text-align:left; padding-left:24px; margin-bottom:14px">ลงนามกรรมการ ${'.'.repeat(55)}</div>
        <div>(${esc(o.interviewerName) || '.'.repeat(55)})</div></div></div>`;
    open({ html: html, landscape: true, exportKind: 'word', barTitle: 'โหมดพิมพ์ตารางให้คะแนน (' + (o.slot === 'inv2' ? 'กรรมการ 2' : 'กรรมการ 1') + ')', fileName: 'ตารางให้คะแนน_' + (o.slot || 'inv1') + '_กลุ่ม' + o.group_no });
  }

  /** ตารางสรุปผลภาพรวม — o: { round_label, exam_date, group ('ทั้งหมด' | เลขกลุ่ม), pass_score, rows:[{group_no, name, s1, s2, average, status}] } */
  function summary(o) {
    const all = o.group === 'ทั้งหมด';
    const num = v => (v !== null && v !== undefined && v !== '' ? esc(v) : '-');
    const rows = (o.rows || []).length ? o.rows.map((r, i) => `<tr><td>${i + 1}</td>${all ? `<td>${esc(r.group_no)}</td>` : ''}<td class="l">${esc(r.name)}</td><td>${num(r.s1)}</td><td>${num(r.s2)}</td><td><b>${r.average !== null && r.average !== undefined ? Number(r.average).toFixed(2) : '-'}</b></td><td><b>${esc(r.status)}</b></td></tr>`).join('')
      : `<tr><td colspan="${all ? 7 : 6}">ไม่มีข้อมูล</td></tr>`;
    const html = `<div class="ivp-head" style="font-size:18px">
        <div>ตารางสรุปผลการประเมินภาพรวม ${TITLE}</div>
        <div>ครั้งที่ ${esc(o.round_label)}</div>
        <div style="font-size:16px; font-weight:normal">กลุ่มสอบ: ${esc(o.group)} | วันสอบ: ${esc(formatDate(o.exam_date))}</div></div>
      <table id="summary-table"><thead><tr><th style="width:8%">ที่</th>${all ? '<th style="width:10%">กลุ่ม</th>' : ''}<th class="l" style="text-align:center">รายชื่อผู้สอบ</th><th style="width:12%">คะแนน ก.1</th><th style="width:12%">คะแนน ก.2</th><th style="width:15%">คะแนนเฉลี่ย</th><th style="width:18%">ผลการประเมิน</th></tr></thead>
      <tbody>${rows}</tbody></table>
      ${o.pass_score ? `<div style="font-size:13px">เกณฑ์ผ่าน: คะแนนเฉลี่ยกรรมการ 2 ท่านไม่น้อยกว่า ${esc(o.pass_score)} คะแนน</div>` : ''}`;
    open({ html: html, landscape: false, exportKind: 'excel', barTitle: 'โหมดพิมพ์สรุปภาพรวม', fileName: 'สรุปผลภาพรวม_' + o.round_label.replace(/\//g, '-') + '_กลุ่ม' + o.group });
  }

  window.ivPrint = { scoreSheet: scoreSheet, summary: summary, close: close };
})();

import { prepareWithSegments, layoutWithLines } from 'https://esm.sh/@chenglou/pretext@latest';

// ─── Config ──────────────────────────────────────────────────────────────────
const COLORS = {
  bg: '#0f1b2d', panel: '#162035', panelBorder: '#1e3050',
  accent: '#2d7ef7', accentLight: '#4d9fff', gold: '#f5a623',
  green: '#27ae60', red: '#e74c3c', purple: '#8e44ad', teal: '#1abc9c',
  textPrimary: '#e8eaf0', textSecondary: '#8896aa', textMuted: '#4a5568',
  gridLine: '#1e2d44', white: '#ffffff',
};
const FONT = 'Sarabun, sans-serif';

// ─── Canvas Setup ─────────────────────────────────────────────────────────────
const canvas = document.getElementById('canvas');
const ctx = canvas.getContext('2d');
let W = 0, H = 0, scale = 1;
let DATA = null;

// ─── SheetJS Loader ───────────────────────────────────────────────────────────
function loadSheetJS() {
  return new Promise((resolve, reject) => {
    if (window.XLSX) return resolve(window.XLSX);
    const s = document.createElement('script');
    s.src = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
    s.onload = () => resolve(window.XLSX);
    s.onerror = reject;
    document.head.appendChild(s);
  });
}

// ─── Excel Parser ─────────────────────────────────────────────────────────────
function parseExcel(buffer) {
  const XLSX = window.XLSX;
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true });

  const fmtDate = (v) => {
    if (!v) return null;
    if (v instanceof Date) {
      return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, '0')}-${String(v.getDate()).padStart(2, '0')}`;
    }
    return String(v);
  };
  const num = (v) => (v == null || v === '' || isNaN(Number(v))) ? null : Math.round(Number(v) * 100) / 100;

  // Sheet 1: ตารางผ่อนชำระ
  const sh1 = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true });
  const meta = {
    accountNo: String(sh1[0]?.[1] ?? ''),
    firstDueDate: fmtDate(sh1[1]?.[1]),
    totalLoan: num(sh1[2]?.[1]) ?? 0,
  };
  const schedule = [];
  for (let i = 6; i < sh1.length; i++) {
    const r = sh1[i];
    if (r[0] == null) continue;
    schedule.push({ no: Number(r[0]), dueDate: fmtDate(r[1]), pct: num(r[2]), total: num(r[3]), interest: num(r[4]), principal: num(r[5]) });
  }

  // Sheet 2: statement
  const sh2 = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[1]], { header: 1, raw: true });
  const payments = [];
  for (let i = 1; i < sh2.length; i++) {
    const r = sh2[i];
    if (r[0] == null) continue;
    payments.push({ no: Number(r[0]), date: fmtDate(r[1]), source: String(r[4] ?? ''), amount: Math.abs(num(r[7]) ?? 0) });
  }
  payments.sort((a, b) => (a.date > b.date ? 1 : -1));

  // Sheet 3: recal
  const sh3 = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[2]], { header: 1, raw: true });
  const recal = [];
  for (let i = 1; i < sh3.length; i++) {
    const r = sh3[i];
    recal.push({
      loanBalance: num(r[1]), installmentNo: num(r[2]), dueDate: fmtDate(r[3]),
      principalBalance: num(r[6]), payDate: fmtDate(r[7]), payAmount: num(r[8]),
      reducePrincipal: num(r[10]), reduceInterest: num(r[11]), reducePenalty: num(r[12]),
      interestAccum: num(r[24]), penaltyAccum: num(r[36]),
    });
  }

  return { meta, schedule, payments, recal };
}

// ─── Stats ────────────────────────────────────────────────────────────────────
function computeStats({ meta, payments, recal }) {
  const totalLoan = meta.totalLoan;
  const totalPaid = payments.reduce((s, p) => s + p.amount, 0);
  let currentBalance = totalLoan;
  for (let i = recal.length - 1; i >= 0; i--) {
    if (recal[i].loanBalance != null) { currentBalance = recal[i].loanBalance; break; }
  }
  const principalPaid = totalLoan - currentBalance;
  const maxInterestAccum = Math.max(0, ...recal.map(r => r.interestAccum || 0));
  const maxPenaltyAccum = Math.max(0, ...recal.map(r => r.penaltyAccum || 0));

  const payByYear = {};
  payments.forEach(p => {
    const yr = (p.date || '').split('-')[0];
    if (yr) payByYear[yr] = (payByYear[yr] || 0) + p.amount;
  });

  const paidInstallments = new Set();
  recal.forEach(r => { if (r.reducePrincipal && r.installmentNo) paidInstallments.add(r.installmentNo); });

  return { totalLoan, totalPaid, currentBalance, principalPaid, maxInterestAccum, maxPenaltyAccum, payByYear, paidInstallments };
}

// ─── Pretext Text Drawing ─────────────────────────────────────────────────────
const _cache = new Map();
function drawText(text, x, y, fontSize, color, align = 'left', maxW = 0) {
  text = String(text);
  const fontStr = `${fontSize}px ${FONT}`;
  const key = `${text}|${fontStr}`;
  if (!_cache.has(key)) _cache.set(key, prepareWithSegments(text, fontStr));
  
  const { lines } = layoutWithLines(_cache.get(key), maxW || W, fontSize * 1.3);
  
  ctx.fillStyle = color;
  ctx.textBaseline = 'top';
  ctx.font = fontStr;
  lines.forEach((ln, i) => {
    const tw = ctx.measureText(ln.text).width;
    const dx = align === 'center' ? x - tw / 2 : align === 'right' ? x - tw : x;
    ctx.fillText(ln.text, dx, y + i * fontSize * 1.3);
  });
}

// ─── Primitives ───────────────────────────────────────────────────────────────
function rr(x, y, w, h, r) {
  ctx.beginPath();
  ctx.moveTo(x + r, y); ctx.lineTo(x + w - r, y); ctx.quadraticCurveTo(x + w, y, x + w, y + r);
  ctx.lineTo(x + w, y + h - r); ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
  ctx.lineTo(x + r, y + h); ctx.quadraticCurveTo(x, y + h, x, y + h - r);
  ctx.lineTo(x, y + r); ctx.quadraticCurveTo(x, y, x + r, y); ctx.closePath();
}

function panel(x, y, w, h, accent = null) {
  rr(x, y, w, h, 10); ctx.fillStyle = COLORS.panel; ctx.fill();
  ctx.strokeStyle = accent ? accent + '44' : COLORS.panelBorder; ctx.lineWidth = 1; ctx.stroke();
  if (accent) { ctx.fillStyle = accent; rr(x + 1, y + 1, w - 2, 4, 3); ctx.fill(); }
}

const fmt = (n, d = 2) => (n ?? 0).toLocaleString('th-TH', { minimumFractionDigits: d, maximumFractionDigits: d });

// ─── Sections ─────────────────────────────────────────────────────────────────
function drawHeader(x, y, w, meta) {
  drawText('ภาพรวมบัญชีเงินกู้ กยศ.', x + w / 2, y, 22, COLORS.white, 'center');
  drawText(`เลขที่บัญชี: ${meta.accountNo}   |   วงเงินกู้: ${fmt(meta.totalLoan)} บาท`, x + w / 2, y + 30, 13, COLORS.textSecondary, 'center');
}

function drawKPICards(x, y, w, s) {
  const cards = [
    { label: 'ยอดกู้ทั้งหมด',   value: fmt(s.totalLoan),        color: COLORS.accent  },
    { label: 'ชำระแล้วสะสม',    value: fmt(s.totalPaid),        color: COLORS.green   },
    { label: 'เงินต้นคงเหลือ',   value: fmt(s.currentBalance),   color: COLORS.gold    },
    { label: 'ดอกเบี้ยสะสม',    value: fmt(s.maxInterestAccum), color: COLORS.purple  },
  ];
  const cw = (w - 30) / 4;
  cards.forEach((c, i) => {
    const cx = x + i * (cw + 10);
    panel(cx, y, cw, 90, c.color);
    drawText(c.label, cx + cw / 2, y + 14, 12, COLORS.textSecondary, 'center');
    drawText(c.value, cx + cw / 2, y + 34, 17, c.color, 'center');
    drawText('บาท', cx + cw / 2, y + 58, 11, COLORS.textMuted, 'center');
  });
}

function drawProgressBar(x, y, w, label, value, total, color) {
  const pct = Math.min(value / Math.max(total, 1), 1);
  const bw = w - 20;
  drawText(label, x + 10, y, 12, COLORS.textSecondary);
  drawText(`${fmt(value)} / ${fmt(total)}  (${(pct * 100).toFixed(1)}%)`, x + w - 10, y, 11, COLORS.textMuted, 'right');
  rr(x + 10, y + 18, bw, 18, 9); ctx.fillStyle = COLORS.panelBorder; ctx.fill();
  if (pct > 0.001) {
    const g = ctx.createLinearGradient(x + 10, 0, x + 10 + bw * pct, 0);
    g.addColorStop(0, color); g.addColorStop(1, color + 'aa');
    ctx.fillStyle = g; rr(x + 10, y + 18, bw * pct, 18, 9); ctx.fill();
  }
  return 44;
}

function drawProgressSection(x, y, w, s) {
  panel(x, y, w, 132, COLORS.teal);
  drawText('ความคืบหน้าการชำระ', x + 16, y + 14, 14, COLORS.white);
  let cy = y + 38;
  cy += drawProgressBar(x, cy, w, 'เงินต้นที่ชำระแล้ว', s.principalPaid, s.totalLoan, COLORS.green) + 4;
  cy += drawProgressBar(x, cy, w, 'ยอดชำระสะสมทั้งหมด', s.totalPaid, s.totalLoan + s.maxInterestAccum, COLORS.accent) + 4;
  drawProgressBar(x, cy, w, 'ดอกเบี้ยสะสม', s.maxInterestAccum, Math.max(s.totalLoan * 0.05, s.maxInterestAccum), COLORS.purple);
}

function drawTimeline(x, y, w, h, s) {
  const yearKeys = Object.keys(s.payByYear).sort();
  panel(x, y, w, h);
  drawText('ยอดชำระจริงรายปี (บาท)', x + 16, y + 14, 14, COLORS.white);
  const p = { t: 44, b: 46, l: 62, r: 16 };
  const gw = w - p.l - p.r, gh = h - p.t - p.b, gx = x + p.l, gy = y + p.t;
  const maxVal = Math.max(...Object.values(s.payByYear), 1);
  const barW = Math.min(gw / yearKeys.length - 8, 42);

  for (let i = 0; i <= 4; i++) {
    const ly = gy + gh * (1 - i / 4);
    ctx.strokeStyle = COLORS.gridLine; ctx.lineWidth = 0.5;
    ctx.beginPath(); ctx.moveTo(gx, ly); ctx.lineTo(gx + gw, ly); ctx.stroke();
    drawText(fmt(maxVal * i / 4, 0), gx - 6, ly - 6, 9, COLORS.textMuted, 'right');
  }
  yearKeys.forEach((yr, i) => {
    const val = s.payByYear[yr];
    const bh = (val / maxVal) * gh;
    const bx = gx + (i + 0.5) * (gw / yearKeys.length) - barW / 2, by = gy + gh - bh;
    const g = ctx.createLinearGradient(bx, by, bx, gy + gh);
    g.addColorStop(0, COLORS.accentLight); g.addColorStop(1, COLORS.accent + '55');
    ctx.fillStyle = g; rr(bx, by, barW, bh, 4); ctx.fill();
    drawText(yr.slice(2), bx + barW / 2, gy + gh + 6, 10, COLORS.textSecondary, 'center');
    if (bh > 24) drawText(fmt(val, 0), bx + barW / 2, by + 5, 9, COLORS.white, 'center');
  });
  ctx.strokeStyle = COLORS.panelBorder; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(gx, gy); ctx.lineTo(gx, gy + gh); ctx.lineTo(gx + gw, gy + gh); ctx.stroke();
}

function drawPaymentList(x, y, w, h, payments) {
  panel(x, y, w, h);
  drawText('ประวัติการชำระจริง', x + 16, y + 14, 14, COLORS.white);
  const rowH = 27, maxRows = Math.floor((h - 50) / rowH);
  [...payments].reverse().slice(0, maxRows).forEach((p, i) => {
    const ry = y + 42 + i * rowH;
    if (ry + rowH > y + h - 8) return;
    if (i % 2 === 0) { ctx.fillStyle = COLORS.bg + '88'; rr(x + 8, ry, w - 16, rowH - 2, 3); ctx.fill(); }
    ctx.fillStyle = COLORS.accent;
    ctx.beginPath(); ctx.arc(x + 22, ry + rowH / 2, 4, 0, Math.PI * 2); ctx.fill();
    drawText(p.date || '', x + 34, ry + 4, 11, COLORS.textSecondary);
    drawText(fmt(p.amount) + ' บาท', x + w - 12, ry + 4, 12, COLORS.green, 'right');
    const src = (p.source || '').slice(0, 22);
    drawText(src, x + 34, ry + 16, 9, COLORS.textMuted);
  });
}

function drawScheduleTable(x, y, w, h, schedule, s) {
  panel(x, y, w, h);
  drawText('ตารางงวดผ่อนชำระ', x + 16, y + 14, 14, COLORS.white);
  const cols = [
    { label: 'งวด', w: 0.07 }, { label: 'กำหนดชำระ', w: 0.20 },
    { label: 'เงินต้น', w: 0.18 }, { label: 'ดอกเบี้ย', w: 0.17 },
    { label: 'ยอดรวม', w: 0.18 }, { label: 'สถานะ', w: 0.20 },
  ];
  const rowH = 24, tx = x + 10, tw = w - 20, sy = y + 38;
  rr(tx, sy, tw, rowH, 4); ctx.fillStyle = COLORS.panelBorder; ctx.fill();
  let cx = tx;
  cols.forEach(c => { drawText(c.label, cx + tw * c.w / 2, sy + 6, 11, COLORS.textSecondary, 'center'); cx += tw * c.w; });

  const maxRows = Math.floor((h - 70) / rowH);
  schedule.slice(0, maxRows).forEach((row, i) => {
    const ry = sy + rowH + i * rowH;
    if (ry + rowH > y + h - 10) return;
    const isPaid = s.paidInstallments.has(row.no);
    if (i % 2 === 0) { ctx.fillStyle = COLORS.bg + 'aa'; rr(tx, ry, tw, rowH, 2); ctx.fill(); }
    if (isPaid) { ctx.fillStyle = COLORS.green + '18'; rr(tx, ry, tw, rowH, 2); ctx.fill(); }
    const vals = [
      { t: row.no, a: 'center' }, { t: row.dueDate || '', a: 'center' },
      { t: fmt(row.principal), a: 'right' }, { t: fmt(row.interest), a: 'right' },
      { t: fmt(row.total), a: 'right' }, { t: isPaid ? '✓ ชำระแล้ว' : '○ รอชำระ', a: 'center' },
    ];
    cx = tx;
    vals.forEach((v, vi) => {
      const cw = tw * cols[vi].w;
      const color = vi === 5 ? (isPaid ? COLORS.green : COLORS.textMuted) : COLORS.textPrimary;
      const ax = v.a === 'right' ? cx + cw - 8 : v.a === 'center' ? cx + cw / 2 : cx + 6;
      drawText(v.t, ax, ry + 6, 11, color, v.a);
      cx += cw;
    });
  });
}

function drawInterestChart(x, y, w, h, recal) {
  panel(x, y, w, h);
  drawText('ดอกเบี้ยสะสม & เบี้ยปรับสะสม', x + 16, y + 14, 14, COLORS.white);
  const iRows = recal.filter(r => r.interestAccum != null && r.payDate);
  if (iRows.length < 2) { drawText('ข้อมูลไม่เพียงพอ', x + w / 2, y + h / 2, 13, COLORS.textMuted, 'center'); return; }
  const pRows = recal.filter(r => r.penaltyAccum != null && r.penaltyAccum > 0 && r.payDate);
  const p = { t: 40, b: 36, l: 56, r: 20 };
  const gw = w - p.l - p.r, gh = h - p.t - p.b, gx = x + p.l, gy = y + p.t;
  const maxI = Math.max(1, ...iRows.map(r => r.interestAccum));
  const dates = iRows.map(r => r.payDate).sort();
  const t0 = +new Date(dates[0]), span = Math.max(1, +new Date(dates[dates.length - 1]) - t0);
  const toX = d => gx + ((+new Date(d) - t0) / span) * gw;
  const toY = (v, m) => gy + gh - (v / m) * gh;

  for (let i = 0; i <= 4; i++) {
    const ly = gy + gh * (1 - i / 4);
    ctx.strokeStyle = COLORS.gridLine; ctx.lineWidth = 0.5; ctx.setLineDash([4, 4]);
    ctx.beginPath(); ctx.moveTo(gx, ly); ctx.lineTo(gx + gw, ly); ctx.stroke(); ctx.setLineDash([]);
    drawText(fmt(maxI * i / 4, 0), gx - 4, ly - 6, 9, COLORS.textMuted, 'right');
  }

  // Area fill
  ctx.beginPath();
  iRows.forEach((r, i) => { i === 0 ? ctx.moveTo(toX(r.payDate), toY(r.interestAccum, maxI)) : ctx.lineTo(toX(r.payDate), toY(r.interestAccum, maxI)); });
  ctx.lineTo(toX(iRows[iRows.length - 1].payDate), gy + gh); ctx.lineTo(toX(iRows[0].payDate), gy + gh);
  ctx.closePath(); ctx.fillStyle = COLORS.purple + '33'; ctx.fill();

  // Line
  ctx.beginPath();
  iRows.forEach((r, i) => { i === 0 ? ctx.moveTo(toX(r.payDate), toY(r.interestAccum, maxI)) : ctx.lineTo(toX(r.payDate), toY(r.interestAccum, maxI)); });
  ctx.strokeStyle = COLORS.purple; ctx.lineWidth = 2; ctx.stroke();

  if (pRows.length > 1) {
    ctx.beginPath();
    pRows.forEach((r, i) => { i === 0 ? ctx.moveTo(toX(r.payDate), toY(r.penaltyAccum, maxI)) : ctx.lineTo(toX(r.payDate), toY(r.penaltyAccum, maxI)); });
    ctx.strokeStyle = COLORS.red; ctx.lineWidth = 2; ctx.stroke();
  }

  ctx.fillStyle = COLORS.purple; ctx.fillRect(x + w - 130, y + 18, 14, 3);
  drawText('ดอกเบี้ย', x + w - 112, y + 12, 11, COLORS.textSecondary);
  if (pRows.length > 1) {
    ctx.fillStyle = COLORS.red; ctx.fillRect(x + w - 72, y + 18, 14, 3);
    drawText('เบี้ยปรับ', x + w - 54, y + 12, 11, COLORS.textSecondary);
  }

  ctx.strokeStyle = COLORS.panelBorder; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(gx, gy); ctx.lineTo(gx, gy + gh); ctx.lineTo(gx + gw, gy + gh); ctx.stroke();
}

// ─── Main Render ──────────────────────────────────────────────────────────────
function render() {
  if (!DATA) return;
  W = canvas.width / scale; H = canvas.height / scale;
  ctx.fillStyle = COLORS.bg; ctx.fillRect(0, 0, W, H);
  ctx.strokeStyle = COLORS.gridLine + '44'; ctx.lineWidth = 0.5;
  for (let i = 0; i < W; i += 40) { ctx.beginPath(); ctx.moveTo(i, 0); ctx.lineTo(i, H); ctx.stroke(); }
  for (let i = 0; i < H; i += 40) { ctx.beginPath(); ctx.moveTo(0, i); ctx.lineTo(W, i); ctx.stroke(); }

  const s = computeStats(DATA);
  const pad = 16, colW = W - pad * 2;
  let cy = pad;

  drawHeader(pad, cy, colW, DATA.meta); cy += 60;
  drawKPICards(pad, cy, colW, s); cy += 106;
  drawProgressSection(pad, cy, colW, s); cy += 148;

  const tlW = Math.floor(colW * 0.62), lstW = colW - tlW - 10;
  drawTimeline(pad, cy, tlW, 220, s);
  drawPaymentList(pad + tlW + 10, cy, lstW, 220, DATA.payments);
  cy += 230;

  const tbW = Math.floor(colW * 0.55), chW = colW - tbW - 10;
  const remH = Math.max(H - cy - pad, 260);
  drawScheduleTable(pad, cy, tbW, remH, DATA.schedule, s);
  drawInterestChart(pad + tbW + 10, cy, chW, remH, DATA.recal);
}

// ─── Resize ───────────────────────────────────────────────────────────────────
function resize() {
  scale = window.devicePixelRatio || 1;

  const ctr = document.getElementById('container');
  const w = ctr.clientWidth - 32;
  const h = Math.max(window.innerHeight - 20, 960);

  canvas.width = w * scale;
  canvas.height = h * scale;
  canvas.style.width = w + 'px';
  canvas.style.height = h + 'px';

  ctx.setTransform(scale, 0, 0, scale, 0, 0);

  _cache.clear();
  render();
}
window.addEventListener('resize', () => { clearTimeout(window._rt); window._rt = setTimeout(resize, 100); });

// ─── File Handling ────────────────────────────────────────────────────────────
async function handleFile(file) {
  const status = document.getElementById('drop-status');
  if (!file || !file.name.match(/\.xlsx?$/i)) {
    status.textContent = '❌ กรุณาเลือกไฟล์ .xlsx เท่านั้น';
    status.style.color = '#e74c3c'; return;
  }
  status.textContent = '⏳ กำลังโหลด…'; status.style.color = '#f5a623';
  try {
    await loadSheetJS();
    const buf = await file.arrayBuffer();
    DATA = parseExcel(buf);
    document.getElementById('dropzone').style.display = 'none';
    canvas.style.display = 'block';
    resize();
  } catch (e) {
    status.textContent = `❌ เกิดข้อผิดพลาด: ${e.message}`;
    status.style.color = '#e74c3c';
    console.error(e);
  }
}

// ─── Init ─────────────────────────────────────────────────────────────────────
window.addEventListener('DOMContentLoaded', () => {
  new FontFace('Sarabun', 'url(https://fonts.gstatic.com/s/sarabun/v14/DtVjJx26TKEr37c9YK5sulU.woff2)')
    .load().then(f => document.fonts.add(f)).catch(() => {});

  resize();

  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('file-input');
  dropzone.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', e => handleFile(e.target.files[0]));
  dropzone.addEventListener('dragover', e => { e.preventDefault(); dropzone.classList.add('drag-over'); });
  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('drag-over'));
  dropzone.addEventListener('drop', e => { e.preventDefault(); dropzone.classList.remove('drag-over'); handleFile(e.dataTransfer.files[0]); });
});

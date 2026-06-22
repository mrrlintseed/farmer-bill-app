// pdfGenerator.js
// Telugu font is loaded from Google Fonts at runtime.
// Requires internet connection on first PDF generation.
// Once loaded it is cached in the browser session.

import jsPDF from 'jspdf';
import 'jspdf-autotable';

const BILL_DATE = "2026-07-01";

function calcInterest(amount, rate, fromDate) {
  const from = new Date(fromDate);
  const to = new Date(BILL_DATE);
  const days = Math.max(1, Math.round((to - from) / 86400000));
  return { days, interest: Math.round((amount * rate * days) / (100 * 30 * 12)) };
}

const WHITE = [255, 255, 255];
const DARK  = [30, 30, 30];
const GREY1 = [60, 60, 60];
const GREY2 = [110, 110, 110];
const GREY3 = [210, 210, 210];
const GREY4 = [240, 240, 240];
const GREY5 = [170, 170, 170];

// Try to load pre-downloaded font first, then fall back to online
let cachedTeluguFont = null;
let cachedFontExt = 'ttf';

async function loadTeluguFont(doc) {
  try {
    if (!cachedTeluguFont) {
      // Try local pre-downloaded font first (fastest, works offline)
      try {
        const local = await import('./telugu-font.js');
        if (local.teluguFontBase64) {
          cachedTeluguFont = local.teluguFontBase64;
          cachedFontExt = local.teluguFontExt || 'ttf';
        }
      } catch { /* file not generated yet, will try online */ }

      // Fall back to fetching from Google Fonts
      if (!cachedTeluguFont) {
        const cssResp = await fetch(
          'https://fonts.googleapis.com/css2?family=Noto+Serif+Telugu&display=swap',
          { headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' } }
        );
        if (!cssResp.ok) throw new Error('css fetch failed');
        const css = await cssResp.text();
        const match = css.match(/url\(([^)]+\.(ttf|woff2?)[^)]*)\)/);
        if (!match) throw new Error('no font url in css');
        const fontUrl = match[1].replace(/['"]/g, '');
        cachedFontExt = match[2] || 'ttf';
        const fontResp = await fetch(fontUrl);
        if (!fontResp.ok) throw new Error('font file fetch failed');
        const buf = await fontResp.arrayBuffer();
        const bytes = new Uint8Array(buf);
        let binary = '';
        for (let i = 0; i < bytes.byteLength; i += 8192) {
          binary += String.fromCharCode(...bytes.subarray(i, i + 8192));
        }
        cachedTeluguFont = btoa(binary);
      }
    }

    const fileName = `NotoSerifTelugu.${cachedFontExt}`;
    doc.addFileToVFS(fileName, cachedTeluguFont);
    doc.addFont(fileName, 'Telugu', 'normal');
    doc.setFont('Telugu', 'normal');
    doc.setFont('helvetica', 'normal');
    return true;
  } catch (e) {
    console.warn('Telugu font not available:', e.message);
    return false;
  }
}

// Safely render Telugu text inline
function tel(doc, text, x, y, opts = {}) {
  try {
    doc.setFont('Telugu', 'normal');
    doc.text(text, x, y, opts);
  } catch { /* skip */ }
  doc.setFont('helvetica', 'normal');
}

// Draw English (bold) + " | " + Telugu (normal, smaller) on same line
function biline(doc, eng, telText, x, y, hasTe, engSz = 9.5, telSz = 8.5) {
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(engSz);
  doc.text(eng, x, y);
  if (hasTe) {
    // getTextWidth uses current font+size — must set before measuring
    const engW = doc.getTextWidth(eng);
    // Draw separator with helvetica
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(engSz);
    const sep = '  |  ';
    doc.text(sep, x + engW, y);
    const sepW = doc.getTextWidth(sep);
    // Draw Telugu at correct position
    doc.setFontSize(telSz);
    tel(doc, telText, x + engW + sepW, y);
    doc.setFontSize(engSz);
  }
  doc.setFont('helvetica', 'normal');
}

function buildPDF(doc, farmer, hasTe) {
  const PW = doc.internal.pageSize.getWidth();
  const M = 12;
  const cW = PW - M * 2;
  let y = M;

  // ── Header ────────────────────────────────────────────────
  doc.setFillColor(...GREY1);
  doc.rect(M, y, cW, 13, 'F');
  doc.setTextColor(...WHITE);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(12);
  if (hasTe) {
    // Draw English part, measure it, then draw Telugu inline
    const engPart = 'FARMER CROP BILL  |  ';
    const teluguPart = 'రైతు పంట బిల్లు';
    // Estimate total width to centre: Telugu ~30mm at 10.5pt
    const engW = doc.getTextWidth(engPart);
    const approxTeW = 28;
    const startX = (PW - engW - approxTeW) / 2;
    doc.text(engPart, startX, y + 8.5);
    doc.setFontSize(10.5);
    tel(doc, teluguPart, startX + engW, y + 8.5);
  } else {
    doc.text('FARMER CROP BILL', PW / 2, y + 8.5, { align: 'center' });
  }
  y += 13 + 4;

  // ── Farmer Info ───────────────────────────────────────────
  const fiH = 22;
  doc.setFillColor(...GREY4);
  doc.rect(M, y, cW, fiH, 'F');
  doc.setDrawColor(...GREY5); doc.setLineWidth(0.4);
  doc.rect(M, y, cW, fiH, 'S');
  doc.setTextColor(...DARK);
  biline(doc, 'FARMER DETAILS', 'రైతు వివరాలు', M + 3, y + 6, hasTe);
  if (farmer.farmerNo) {
    doc.setFillColor(...GREY1);
    doc.roundedRect(PW - M - 42, y + 2, 40, 9, 2, 2, 'F');
    doc.setTextColor(...WHITE); doc.setFont('helvetica', 'bold'); doc.setFontSize(9);
    doc.text(`Farmer No: ${farmer.farmerNo}`, PW - M - 22, y + 7.5, { align: 'center' });
  }
  doc.setFont('helvetica', 'normal'); doc.setTextColor(...DARK); doc.setFontSize(9.5);
  const c3 = cW / 3, dY = y + 15;
  doc.text(`Name: ${farmer.name || ''}`, M + 3, dY);
  doc.text(`Father: ${farmer.fatherName || ''}`, M + c3 + 3, dY);
  doc.text(`Village: ${farmer.village || ''}`, M + c3 * 2 + 3, dY);
  y += fiH + 4;

  // ── Advances ─────────────────────────────────────────────
  const advances = farmer.advances || [];
  if (advances.length > 0) {
    doc.setTextColor(...DARK);
    biline(doc, 'ADVANCE DETAILS', 'అడ్వాన్స్ వివరాలు', M, y, hasTe);
    y += 5;

    let tA = 0, tI = 0;
    const aRows = advances.map((a, i) => {
      const { days, interest } = calcInterest(a.amount, a.interestRate, a.date);
      tA += a.amount; tI += interest;
      return [String(i+1), a.date||'—', a.note||'—',
        `Rs.${a.amount.toLocaleString('en-IN')}`, String(days),
        `Rs.${interest.toLocaleString('en-IN')}`,
        `Rs.${(a.amount+interest).toLocaleString('en-IN')}`];
    });
    aRows.push([
      {content:'TOTAL',colSpan:3,styles:{fontStyle:'bold',fillColor:GREY3,halign:'center'}},
      {content:`Rs.${tA.toLocaleString('en-IN')}`,styles:{fontStyle:'bold',fillColor:GREY3}},
      {content:'',styles:{fillColor:GREY3}},
      {content:`Rs.${tI.toLocaleString('en-IN')}`,styles:{fontStyle:'bold',fillColor:GREY3}},
      {content:`Rs.${(tA+tI).toLocaleString('en-IN')}`,styles:{fontStyle:'bold',fillColor:GREY3}},
    ]);
    doc.autoTable({
      startY:y, head:[['S.No','Date','Note','Amount (Rs.)','Days','Interest (Rs.)','Total (Rs.)']],
      body:aRows, margin:{left:M,right:M},
      styles:{fontSize:9,cellPadding:2.8,halign:'center',textColor:DARK,lineColor:GREY5,lineWidth:0.3},
      headStyles:{fillColor:GREY2,textColor:WHITE,fontStyle:'bold',fontSize:9},
      alternateRowStyles:{fillColor:GREY4},
      columnStyles:{0:{cellWidth:10},1:{cellWidth:26},2:{cellWidth:28},3:{cellWidth:30},4:{cellWidth:14},5:{cellWidth:30},6:{cellWidth:28}},
    });
    y = doc.lastAutoTable.finalY + 6;
  }

  // ── Crops ─────────────────────────────────────────────────
  const crops = farmer.crops || [];
  if (crops.length > 0) {
    doc.setTextColor(...DARK);
    biline(doc, 'CROP DETAILS', 'పంట వివరాలు', M, y, hasTe);
    y += 5;

    let tC = 0, tF = 0, tT = 0;
    const cRows = crops.map((c, i) => {
      const qty = parseFloat(c.quantity)||0, area = parseFloat(c.area)||0;
      const val = c.result==='Pass' ? qty*(parseFloat(c.ratePerUnit)||0) : 0;
      tC += val; tF += area*1000; tT += qty;
      return [String(i+1), c.variety||'—', c.lotNo||'—',
        area>0?`${area} Ac`:'—', String(qty),
        c.result==='Pass'
          ? {content:'PASS',styles:{fontStyle:'bold',fillColor:GREY4}}
          : {content:'FAIL',styles:{fontStyle:'bold',fillColor:GREY3}},
        c.cropType||'KMS',
        c.result==='Pass'?`Rs.${(parseFloat(c.ratePerUnit)||0).toLocaleString('en-IN')}`:'—',
        c.result==='Pass'?{content:`Rs.${val.toLocaleString('en-IN')}`,styles:{fontStyle:'bold'}}:'—',
      ];
    });
    cRows.push([
      {content:'TOTAL CROP VALUE',colSpan:8,styles:{fontStyle:'bold',fillColor:GREY3,halign:'center'}},
      {content:`Rs.${tC.toLocaleString('en-IN')}`,styles:{fontStyle:'bold',fillColor:GREY3}},
    ]);
    doc.autoTable({
      startY:y, head:[['S.No','Variety','Lot No','Area','Qty','Result','Type','Rate (Rs.)','Value (Rs.)']],
      body:cRows, margin:{left:M,right:M},
      styles:{fontSize:9,cellPadding:2.8,halign:'center',textColor:DARK,lineColor:GREY5,lineWidth:0.3},
      headStyles:{fillColor:GREY1,textColor:WHITE,fontStyle:'bold',fontSize:9},
      alternateRowStyles:{fillColor:GREY4},
      columnStyles:{0:{cellWidth:10},1:{cellWidth:35},2:{cellWidth:26},3:{cellWidth:16},4:{cellWidth:12},5:{cellWidth:16},6:{cellWidth:13},7:{cellWidth:22},8:{cellWidth:36}},
    });
    y = doc.lastAutoTable.finalY + 6;

    // Jamma
    const jE = farmer.jammaEnabled ? (farmer.jammaEntries||[]) : [];
    if (jE.length > 0) {
      doc.setTextColor(...DARK);
      biline(doc, 'JAMMA DETAILS', 'జమ్మా వివరాలు', M, y, hasTe);
      y += 5;
      let tJ=0, tJI=0;
      const jRows = jE.map((j,i) => {
        const amt=parseFloat(j.amount)||0;
        const {days,interest}=calcInterest(amt,parseFloat(j.interestRate)||0,j.date||BILL_DATE);
        tJ+=amt; tJI+=interest;
        return [String(i+1),j.date||'—',j.note||'—',`Rs.${amt.toLocaleString('en-IN')}`,String(days),`Rs.${interest.toLocaleString('en-IN')}`,`Rs.${(amt+interest).toLocaleString('en-IN')}`];
      });
      jRows.push([
        {content:'TOTAL JAMMA',colSpan:3,styles:{fontStyle:'bold',fillColor:GREY3,halign:'center'}},
        {content:`Rs.${tJ.toLocaleString('en-IN')}`,styles:{fontStyle:'bold',fillColor:GREY3}},
        {content:'',styles:{fillColor:GREY3}},
        {content:`Rs.${tJI.toLocaleString('en-IN')}`,styles:{fontStyle:'bold',fillColor:GREY3}},
        {content:`Rs.${(tJ+tJI).toLocaleString('en-IN')}`,styles:{fontStyle:'bold',fillColor:GREY3}},
      ]);
      doc.autoTable({
        startY:y, head:[['S.No','Date','Note','Amount (Rs.)','Days','Interest (Rs.)','Total (Rs.)']],
        body:jRows, margin:{left:M,right:M},
        styles:{fontSize:9,cellPadding:2.8,halign:'center',textColor:DARK,lineColor:GREY5,lineWidth:0.3},
        headStyles:{fillColor:GREY2,textColor:WHITE,fontStyle:'bold',fontSize:9},
        alternateRowStyles:{fillColor:GREY4},
        columnStyles:{0:{cellWidth:10},1:{cellWidth:26},2:{cellWidth:28},3:{cellWidth:30},4:{cellWidth:14},5:{cellWidth:30},6:{cellWidth:28}},
      });
      y = doc.lastAutoTable.finalY + 6;
    }

    // Settlement
    const advWI=(farmer.advances||[]).reduce((s,a)=>{const{interest}=calcInterest(a.amount,a.interestRate,a.date);return s+a.amount+interest;},0);
    const jamWI=(farmer.jammaEnabled?(farmer.jammaEntries||[]):[]).reduce((s,j)=>{const{interest}=calcInterest(parseFloat(j.amount)||0,parseFloat(j.interestRate)||0,j.date||BILL_DATE);return s+(parseFloat(j.amount)||0)+interest;},0);
    const balance = tC - advWI + jamWI - tF - tT;
    const isP = balance >= 0;
    const lines = [
      ['Total Crop Value',`Rs.${tC.toLocaleString('en-IN')}`],
      ['Total Advance + Interest',`- Rs.${advWI.toLocaleString('en-IN')}`],
    ];
    if (farmer.jammaEnabled && jamWI>0) lines.push(['Jamma + Interest',`+ Rs.${jamWI.toLocaleString('en-IN')}`]);
    if (tF>0) lines.push(['Foundation',`- Rs.${tF.toLocaleString('en-IN')}`]);
    if (tT>0) lines.push(['Transportation',`- Rs.${tT.toLocaleString('en-IN')}`]);

    // Extra height for Telugu in balance row
    const boxH = 14 + lines.length*9 + 14;
    doc.setFillColor(...GREY4); doc.setDrawColor(...GREY1); doc.setLineWidth(0.8);
    doc.rect(M, y, cW, boxH, 'FD');

    doc.setTextColor(...DARK);
    biline(doc, 'SETTLEMENT SUMMARY', 'తీర్పు సారాంశం', M+4, y+7, hasTe);
    doc.setDrawColor(...GREY5); doc.setLineWidth(0.3);
    doc.line(M+4, y+9, M+cW-4, y+9);
    let ry = y + 16;
    lines.forEach(([label, val]) => {
      doc.setFont('helvetica','normal'); doc.setFontSize(9.5); doc.setTextColor(...GREY1);
      doc.text(label, M+6, ry);
      doc.setFont('helvetica','bold'); doc.setTextColor(...DARK);
      doc.text(val, M+cW-5, ry, {align:'right'});
      ry += 9;
    });
    doc.setDrawColor(...DARK); doc.setLineWidth(0.6);
    doc.line(M+4, ry, M+cW-4, ry); ry += 5;

    const balEng = isP?'Balance Payable to Farmer':'Balance Due from Farmer';
    const balTe  = isP?'రైతుకు చెల్లించవలసిన మొత్తం':'రైతు నుండి రావలసిన మొత్తం';
    biline(doc, balEng, balTe, M+6, ry, hasTe, 10, 8.5);
    doc.setFont('helvetica','bold'); doc.setFontSize(13); doc.setTextColor(...DARK);
    doc.text(`Rs.${Math.abs(balance).toLocaleString('en-IN')}`, M+cW-5, ry, {align:'right'});
    y += boxH + 6;

    // Footer
    doc.setDrawColor(...GREY5); doc.setLineWidth(0.4);
    doc.line(M, y, M+cW, y); y += 6;
    doc.setFont('helvetica','normal'); doc.setFontSize(9); doc.setTextColor(...GREY2);
    doc.text('_________________________', M+10, y+8);
    doc.setTextColor(...DARK);
    biline(doc, 'Farmer Signature', 'రైతు సంతకం', M+10, y+14, hasTe, 9, 8.5);
    doc.setFont('helvetica','normal'); doc.setFontSize(9); doc.setTextColor(...GREY2);
    if (hasTe) {
      const tyEng = 'Thank you for your cooperation  |  ';
      doc.setFont('helvetica','normal'); doc.setFontSize(9);
      const engW = doc.getTextWidth(tyEng);
      const approxTeW = 42;
      const startX = (PW - engW - approxTeW) / 2;
      doc.text(tyEng, startX, y+10);
      doc.setFontSize(8.5);
      tel(doc, 'మీ సహకారానికి ధన్యవాదాలు', startX + engW, y+10);
    } else {
      doc.text('Thank you for your cooperation', PW/2, y+10, {align:'center'});
    }
  }
}

export async function generateFarmerBillPDF(farmer) {
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const hasTe = await loadTeluguFont(doc);
  buildPDF(doc, farmer, hasTe);
  return doc;
}

export async function downloadFarmerBillPDF(farmer) {
  const doc = await generateFarmerBillPDF(farmer);
  const name = (farmer.name || 'farmer').replace(/\s+/g, '_');
  doc.save(`bill_${farmer.farmerNo || name}.pdf`);
}

export async function downloadAllBillsPDF(farmers) {
  if (!farmers || farmers.length === 0) return;
  for (const f of farmers) {
    const d = await generateFarmerBillPDF(f);
    const name = (f.name || 'farmer').replace(/\s+/g, '_');
    d.save(`bill_${f.farmerNo || name}.pdf`);
  }
}

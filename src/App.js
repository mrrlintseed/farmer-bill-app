import { useState, useRef, useEffect, useCallback } from "react";
import { saveToCloud, loadFromCloud, saveSnapshot, getSnapshots, restoreSnapshot } from "./firebase";
import * as XLSX from "xlsx";
import { storage } from "./storage";
import { downloadFarmerBillPDF, downloadAllBillsPDF } from "./pdfGenerator";

const DEFAULT_BILL_DATE = "2026-07-01";
// App version: 2026-07-02-v3-foundation-seeds-fix
// Global BILL_DATE reads from localStorage so outside-App functions can use it
let BILL_DATE = (() => { try { return localStorage.getItem("app_bill_date")||DEFAULT_BILL_DATE; } catch { return DEFAULT_BILL_DATE; } })();

const sampleFarmers = [{
  id: 1, farmerNo: "001", name: "Ravi Kumar", fatherName: "Suresh Kumar", village: "Nandyal",
  advances: [
    { date: "2024-01-15", amount: 5000, interestRate: 24, note: "Seed advance" },
    { date: "2024-03-10", amount: 3000, interestRate: 24, note: "Fertilizer advance" },
  ],
  crops: [
    { variety: "Paddy - BPT 5204", area: "2.5", quantity: 20, cropType: "KMS", ratePerUnit: 550, result: "Pass" },
    { variety: "Groundnut", area: "1", quantity: 5, cropType: "GMS", ratePerUnit: 400, result: "Fail" },
  ],
  jammaEnabled: false, jammaEntries: [],
}];

// Format date string (YYYY-MM-DD) without timezone shift
function fmtDate(dateStr) {
  if (!dateStr) return "—";
  const parts = String(dateStr).split("-");
  if (parts.length === 3) {
    return `${parseInt(parts[2])}/${parseInt(parts[1])}/${parts[0]}`;
  }
  return dateStr;
}

// Parse date string as local date (avoids UTC midnight timezone shift)
function parseLocalDate(dateStr) {
  if (!dateStr) return new Date();
  const parts = String(dateStr).split("-");
  if (parts.length === 3) {
    return new Date(parseInt(parts[0]), parseInt(parts[1])-1, parseInt(parts[2]));
  }
  return new Date(dateStr);
}

function calcInterest(amount, rate, fromDate, billDate) {
  const from = parseLocalDate(fromDate);
  const to = parseLocalDate(billDate || BILL_DATE);
  const days = Math.max(1, Math.round((to - from) / 86400000));
  return { days, interest: Math.round((amount * rate * days) / (100 * 30 * 12)) };

// Compound interest — compounds yearly
function calcCompoundInterest(amount, rate, fromDate, billDate) {
  const from = parseLocalDate(fromDate);
  const to = parseLocalDate(billDate || BILL_DATE);
  const totalDays = Math.max(1, Math.round((to - from) / 86400000));
  const fullYears = Math.floor(totalDays / 365);
  const remainingDays = totalDays % 365;
  let principal = amount;
  for (let y = 0; y < fullYears; y++) {
    const yearInterest = Math.round((principal * rate * 365) / (100 * 30 * 12));
    principal = principal + yearInterest;
  }
  const remainingInterest = remainingDays > 0 ? Math.round((principal * rate * remainingDays) / (100 * 30 * 12)) : 0;
  const totalInterest = principal + remainingInterest - amount;
  return { days: totalDays, interest: Math.round(totalInterest) };
}
}

function printBill(elementId, filename) {
  const el = document.getElementById(elementId);
  if (!el) { alert("Please go to the Preview tab first, then click Print."); return; }

  const origTitle = document.title;
  document.title = filename || "farmer-bill";

  // Get the bill HTML with all inline styles (React renders inline styles so they copy perfectly)
  const billHTML = el.outerHTML;

  // Save original body
  const origBody = document.body.innerHTML;

  // Replace body with just the bill
  document.body.innerHTML = `
    <link href="https://fonts.googleapis.com/css2?family=Noto+Serif+Telugu&family=Noto+Serif:wght@400;600;700&display=swap" rel="stylesheet"/>
    <style>
      * { box-sizing: border-box; margin: 0; padding: 0; }
      body { font-family: 'Noto Serif', Georgia, serif; padding: 10px; background: #fff; }
      table { border-collapse: collapse; width: 100%; }
      * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
      @page { margin: 8mm; size: A4 portrait; }
    </style>
    ${billHTML}
  `;

  // Use setTimeout to let the DOM update before printing
  setTimeout(() => {
    window.print();
    // After print dialog closes, reload to restore the app
    setTimeout(() => {
      document.title = origTitle;
      window.location.reload();
    }, 500);
  }, 300);
}

// ─── Farmer Bill ────────────────────────────────────────────────
function BillPreview({ farmer, varietySettings, getVarietyBillDate, isVarietyPaid, getVarietyRate, getVarietyType }) {
  // Safe fallbacks when props not passed
  const _isVarietyPaid = isVarietyPaid || (() => true);
  const _getVarietyBillDate = getVarietyBillDate || (() => BILL_DATE);
  const _getVarietyRate = getVarietyRate || (() => null);
  const _getVarietyType = getVarietyType || (() => null);

  const billNo = `BILL-${farmer.farmerNo || farmer.id || "001"}-2026`;
  const advCalc = (farmer.advances || []).map(a => {
    // Use earliest paid variety bill date for advance interest
    const paidDates = (farmer.crops||[]).filter(c=>c.result==="Pass"&&_isVarietyPaid(c.variety)).map(c=>_getVarietyBillDate(c.variety));
    const advBillDate = a.tillDate || (paidDates.length > 0 ? paidDates.reduce((a,b)=>a<b?a:b) : BILL_DATE);
    const { days, interest } = (a.compound?calcCompoundInterest:calcInterest)(a.amount, a.interestRate, a.date, advBillDate);
    return { ...a, days, interest, total: a.amount + interest };
  });
  const totalAdv = advCalc.reduce((s, a) => s + a.amount, 0);
  const totalAdvInt = advCalc.reduce((s, a) => s + a.interest, 0);
  const totalAdvWithInt = totalAdv + totalAdvInt;

  const cropsCalc = (farmer.crops || []).map(c => {
    const area = parseFloat(c.area) || 0;
    const qty = parseFloat(c.quantity) || 0;
    const vBillDate = _getVarietyBillDate(c.variety);
    const vPaid = _isVarietyPaid(c.variety);
    const vRate = (c.rateOverride === true) ? (parseFloat(c.ratePerUnit) || 0) : (_getVarietyRate(c.variety) || (parseFloat(c.ratePerUnit) || 0));
    const vType = _getVarietyType(c.variety) || c.cropType || "KMS";
    const isPaidPass = c.result === "Pass" && vPaid;
    return { ...c, cropType: vType, rateUsed: vRate, value: isPaidPass ? qty * vRate : 0, pendingValue: (c.result === "Pass" && !vPaid) ? qty * vRate : 0, foundation: area * 1000, transportation: qty, vBillDate, vPaid };
  });
  const totalCropValue = cropsCalc.reduce((s, c) => s + c.value, 0);
  const totalFoundation = cropsCalc.reduce((s, c) => s + c.foundation, 0);
  const totalTransport = cropsCalc.reduce((s, c) => s + c.transportation, 0);

  const jammaCalc = (farmer.jammaEnabled ? farmer.jammaEntries || [] : []).map(j => {
    const jamBillDate = j.tillDate || BILL_DATE;
    const { days, interest } = calcInterest(parseFloat(j.amount) || 0, parseFloat(j.interestRate) || 0, j.date || BILL_DATE, jamBillDate);
    return { ...j, days, interest, total: (parseFloat(j.amount) || 0) + interest };
  });
  const totalJamma = jammaCalc.reduce((s, j) => s + (parseFloat(j.amount) || 0), 0);
  const totalJammaInt = jammaCalc.reduce((s, j) => s + j.interest, 0);
  const totalJammaWithInt = totalJamma + totalJammaInt;
  const balance = totalCropValue - totalAdvWithInt + totalJammaWithInt - totalFoundation - totalTransport;

  const TH = ({ ch }) => <th style={{ padding: "4px 6px", textAlign: "center", fontWeight: 600, whiteSpace: "nowrap", fontSize: 11 }}>{ch}</th>;
  const TD = ({ ch, s }) => <td style={{ padding: "4px 6px", textAlign: "center", whiteSpace: "nowrap", fontSize: 11, ...s }}>{ch}</td>;

  return (
    <div className="bill-print" style={{ fontFamily: "Noto Serif Telugu,Georgia,serif", background: "#fff", color: "#1a1a1a", width: "100%", border: "2px solid #2d6a2d", borderRadius: 4 }}>
      <div style={{ background: "linear-gradient(135deg,#1a4a1a,#2d6a2d)", color: "#fff", padding: "14px 20px", display: "flex", justifyContent: "center", alignItems: "center" }}>
        <span style={{ fontSize: 16, fontWeight: 700 }}>FARMER CROP BILL | రైతు పంట బిల్లు</span>
      </div>
      <div style={{ padding: "14px 20px" }}>
        {/* Farmer Info */}
        <div style={{ background: "#f0f7f0", borderRadius: 6, padding: "10px 14px", marginBottom: 12, border: "1px solid #b8d8b8" }}>
          <div style={{ fontWeight: 700, color: "#1a4a1a", marginBottom: 6, fontSize: 13, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span>FARMER DETAILS | రైతు వివరాలు</span>
            {farmer.farmerNo && <span style={{ background: "#1a4a1a", color: "#fff", padding: "2px 14px", borderRadius: 20, fontSize: 13, fontWeight: 800 }}>No: {farmer.farmerNo}</span>}
          </div>
          <div style={{ display: "grid", gridTemplateColumns: farmer.careOf ? "1fr 1fr 1fr 1fr" : "1fr 1fr 1fr", gap: "5px 14px", fontSize: 13 }}>
            <div><span style={{ color: "#555" }}>Name: </span><strong>{farmer.name}</strong></div>
            <div><span style={{ color: "#555" }}>Father: </span><strong>{farmer.fatherName}</strong></div>
            <div><span style={{ color: "#555" }}>Village: </span><strong>{farmer.village}</strong></div>
            {farmer.careOf && <div><span style={{ color: "#555" }}>C/o: </span><strong>{farmer.careOf}</strong></div>}
          </div>
        </div>

        {/* Advances */}
        {advCalc.length > 0 && (
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontWeight: 700, color: "#1a4a1a", marginBottom: 5, fontSize: 12 }}>ADVANCE DETAILS | అడ్వాన్స్ వివరాలు</div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 480 }}>
                <thead><tr style={{ background: "#2d6a2d", color: "#fff" }}><TH ch="S.No" /><TH ch="Date" /><TH ch="Note" /><TH ch="Amount(₹)" /><TH ch="Days" /><TH ch="Interest(₹)" /><TH ch="Total(₹)" /></tr></thead>
                <tbody>
                  {advCalc.map((a, i) => (
                    <tr key={i} style={{ background: i % 2 === 0 ? "#f9fdf9" : "#fff", borderBottom: "1px solid #d4e8d4" }}>
                      <TD ch={i + 1} />
                      <TD ch={fmtDate(a.date)} />
                      <TD ch={a.note || "—"} />
                      <TD ch={`₹${a.amount.toLocaleString("en-IN")}`} />
                      <TD ch={a.days} />
                      <TD ch={`₹${a.interest.toLocaleString("en-IN")}`} s={{ color: "#c0392b" }} />
                      <TD ch={`₹${a.total.toLocaleString("en-IN")}`} s={{ fontWeight: 600 }} />
                    </tr>
                  ))}
                  <tr style={{ background: "#e8f5e9", fontWeight: 700 }}>
                    <td colSpan={3} style={{ padding: "4px 6px", fontSize: 11, color: "#1a4a1a" }}>TOTAL</td>
                    <TD ch={`₹${totalAdv.toLocaleString("en-IN")}`} /><TD ch="" />
                    <TD ch={`₹${totalAdvInt.toLocaleString("en-IN")}`} s={{ color: "#c0392b" }} />
                    <TD ch={`₹${totalAdvWithInt.toLocaleString("en-IN")}`} />
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* Crops */}
        {cropsCalc.length > 0 && (
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontWeight: 700, color: "#1a4a1a", marginBottom: 5, fontSize: 12 }}>CROP DETAILS | పంట వివరాలు</div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 500 }}>
                <thead><tr style={{ background: "#1a4a1a", color: "#fff" }}><TH ch="S.No" /><TH ch="Variety" /><TH ch="Lot No" /><TH ch="Area(Ac)" /><TH ch="Qty" /><TH ch="Result" /><TH ch="Type" /><TH ch="Rate(₹)" /><TH ch="Value(₹)" /></tr></thead>
                <tbody>
                  {cropsCalc.map((c, i) => (
                    <tr key={i} style={{ background: i % 2 === 0 ? "#f9fdf9" : "#fff", borderBottom: c.note ? "none" : "1px solid #d4e8d4" }}>
                      <TD ch={i + 1} />
                      <TD ch={c.variety} />
                      <TD ch={c.lotNo || "—"} />
                      <TD ch={c.area ? `${c.area} Ac` : "—"} />
                      <TD ch={c.quantity} />
                      <TD ch={<span style={{ padding: "2px 7px", borderRadius: 10, fontSize: 10, fontWeight: 700, background: c.result === "Pass" ? "#d4edda" : "#fdecea", color: c.result === "Pass" ? "#155724" : "#721c24" }}>{c.result === "Pass" ? "✓ Pass" : "✗ Fail"}</span>} />
                      <TD ch={<span style={{ padding: "2px 7px", borderRadius: 8, fontSize: 10, fontWeight: 700, background: c.cropType === "GMS" ? "#fff3cd" : "#e8f0ff", color: c.cropType === "GMS" ? "#856404" : "#1a3a8a" }}>{c.cropType || "KMS"}</span>} />
                      <TD ch={c.result === "Pass" ? `₹${(c.rateUsed || parseFloat(c.ratePerUnit) || 0).toLocaleString("en-IN")}` : "—"} s={{ color: c.result === "Pass" ? "#1a1a1a" : "#aaa" }} />
                      <TD ch={c.result === "Pass" ? (c.vPaid !== false ? `₹${c.value.toLocaleString("en-IN")}` : `⏳ Pending`) : "—"} s={{ fontWeight: 600, color: c.result === "Pass" ? (c.vPaid !== false ? "#1a6a1a" : "#856404") : "#aaa", background: (c.result === "Pass" && c.vPaid === false) ? "#fff3cd" : "transparent" }} />
                    </tr>
                  )).reduce((acc, row, i) => {
                    acc.push(row);
                    const c = cropsCalc[i];
                    if (c.note) {
                      acc.push(
                        <tr key={"note"+i} style={{ background: i % 2 === 0 ? "#f9fdf9" : "#fff", borderBottom: "1px solid #d4e8d4" }}>
                          <td colSpan={9} style={{ padding: "1px 8px 4px 24px", fontSize: 10, color: "#856404", fontStyle: "italic" }}>📝 {c.note}</td>
                        </tr>
                      );
                    }
                    return acc;
                  }, [])}
                  <tr style={{ background: "#e8f5e9", fontWeight: 700 }}>
                    <td colSpan={6} style={{ padding: "4px 6px", fontSize: 11, color: "#1a4a1a" }}>TOTAL CROP VALUE</td>
                    <TD ch={`₹${totalCropValue.toLocaleString("en-IN")}`} s={{ color: "#1a6a1a" }} />
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* Jamma */}
        {farmer.jammaEnabled && jammaCalc.length > 0 && (
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontWeight: 700, color: "#856404", marginBottom: 5, fontSize: 12 }}>JAMMA DETAILS | జమ్మా వివరాలు</div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 440 }}>
                <thead><tr style={{ background: "#856404", color: "#fff" }}><TH ch="S.No" /><TH ch="Date" /><TH ch="Note" /><TH ch="Amount(₹)" /><TH ch="Days" /><TH ch="Interest(₹)" /><TH ch="Total(₹)" /></tr></thead>
                <tbody>
                  {jammaCalc.map((j, i) => (
                    <tr key={i} style={{ background: i % 2 === 0 ? "#fffdf0" : "#fff", borderBottom: "1px solid #ffeeba" }}>
                      <TD ch={i + 1} />
                      <TD ch={j.date ? fmtDate(j.date) : "—"} />
                      <TD ch={j.note || "—"} />
                      <TD ch={`₹${(parseFloat(j.amount) || 0).toLocaleString("en-IN")}`} s={{ color: "#856404", fontWeight: 600 }} />
                      <TD ch={j.days} />
                      <TD ch={`₹${j.interest.toLocaleString("en-IN")}`} s={{ color: "#856404" }} />
                      <TD ch={`₹${j.total.toLocaleString("en-IN")}`} s={{ fontWeight: 700, color: "#5a3e00" }} />
                    </tr>
                  ))}
                  <tr style={{ background: "#fff3cd", fontWeight: 700 }}>
                    <td colSpan={3} style={{ padding: "4px 6px", fontSize: 11, color: "#856404" }}>TOTAL JAMMA</td>
                    <TD ch={`₹${totalJamma.toLocaleString("en-IN")}`} s={{ color: "#856404" }} /><TD ch="" />
                    <TD ch={`₹${totalJammaInt.toLocaleString("en-IN")}`} s={{ color: "#856404" }} />
                    <TD ch={`₹${totalJammaWithInt.toLocaleString("en-IN")}`} s={{ color: "#5a3e00" }} />
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* Settlement */}
        <div style={{ background: balance >= 0 ? "#e8f5e9" : "#fdecea", borderRadius: 6, padding: "12px 16px", border: `2px solid ${balance >= 0 ? "#2d6a2d" : "#e74c3c"}` }}>
          <div style={{ fontWeight: 700, color: "#1a4a1a", marginBottom: 8, fontSize: 13 }}>SETTLEMENT SUMMARY | తీర్పు సారాంశం</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr auto", gap: "4px 16px", fontSize: 13 }}>
            <div style={{ color: "#555" }}>Total Crop Value (Paid)</div><div style={{ textAlign: "right", fontWeight: 600, color: "#1a6a1a" }}>₹{totalCropValue.toLocaleString("en-IN")}</div>
            {cropsCalc.some(c=>c.pendingValue>0) && <>
              <div style={{ color: "#856404" }}>⏳ Pending (not counted)</div>
              <div style={{ textAlign: "right", fontWeight: 600, color: "#856404" }}>₹{cropsCalc.reduce((s,c)=>s+(c.pendingValue||0),0).toLocaleString("en-IN")}</div>
            </>}
            <div style={{ color: "#555" }}>Total Advance + Interest</div><div style={{ textAlign: "right", fontWeight: 600, color: "#c0392b" }}>− ₹{totalAdvWithInt.toLocaleString("en-IN")}</div>
            {farmer.jammaEnabled && totalJammaWithInt > 0 && <><div style={{ color: "#856404" }}>Jamma + Interest</div><div style={{ textAlign: "right", fontWeight: 600, color: "#856404" }}>+ ₹{totalJammaWithInt.toLocaleString("en-IN")}</div></>}
            {totalFoundation > 0 && <><div style={{ color: "#555" }}>Foundation</div><div style={{ textAlign: "right", fontWeight: 600, color: "#c0392b" }}>− ₹{totalFoundation.toLocaleString("en-IN")}</div></>}
            {totalTransport > 0 && <><div style={{ color: "#555" }}>Transportation</div><div style={{ textAlign: "right", fontWeight: 600, color: "#c0392b" }}>− ₹{totalTransport.toLocaleString("en-IN")}</div></>}
            <div style={{ borderTop: "2px solid #2d6a2d", paddingTop: 6, marginTop: 4, fontWeight: 700, fontSize: 14, color: balance >= 0 ? "#1a4a1a" : "#c0392b" }}>
              {balance >= 0 ? "Balance Payable to Farmer | రైతుకు చెల్లించవలసిన మొత్తం" : "Balance Due from Farmer | రైతు నుండి రావలసిన మొత్తం"}
            </div>
            <div style={{ borderTop: "2px solid #2d6a2d", paddingTop: 6, marginTop: 4, textAlign: "right", fontWeight: 800, fontSize: 18, color: balance >= 0 ? "#1a4a1a" : "#c0392b" }}>
              ₹{Math.abs(balance).toLocaleString("en-IN")}
            </div>
          </div>
        </div>

        {farmer.comment && farmer.comment.trim() && (
          <div style={{ marginTop: 14, background: "#fff9e6", border: "1.5px solid #c8a000", borderRadius: 6, padding: "10px 14px" }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: "#856404", marginBottom: 2 }}>📝 Note | గమనిక</div>
            <div style={{ fontSize: 12, color: "#5c4400" }}>{farmer.comment}</div>
          </div>
        )}

        <div style={{ marginTop: 16, display: "flex", justifyContent: "space-between", fontSize: 11, color: "#666", borderTop: "1px solid #d4e8d4", paddingTop: 10 }}>
          <div style={{ textAlign: "center" }}><div style={{ borderTop: "1px solid #aaa", width: 120, marginBottom: 4 }}></div><div>Farmer Signature | రైతు సంతకం</div></div>
          <div style={{ textAlign: "center", color: "#2d6a2d", fontSize: 12 }}><div>మీ సహకారానికి ధన్యవాదాలు</div><div>Thank you for your cooperation</div></div>
        </div>
      </div>
    </div>
  );
}

// ─── Sub-Org Bill ───────────────────────────────────────────────
function SubOrgBill({ so, isSubOrgVarietyPaid, isSubOrgVarietySettled, isSubOrgVarietyToPay, getSubOrgVarietyBillDate, getSubOrgVarietyRate, getSubOrgVarietyType, soVarieties, billMode, selectedVarieties, settledVarieties }) {
  const _isPaid = isSubOrgVarietyPaid || (() => true);
  // Use per-sub-org settledVarieties array instead of global settled status
  const _isSettled = (v) => (settledVarieties||[]).includes(v);
  const _isToPay = (v) => _isPaid(v) && !_isSettled(v);
  const _getBillDate = getSubOrgVarietyBillDate || (() => BILL_DATE);
  const _getRate = getSubOrgVarietyRate || (() => null);
  const _getType = getSubOrgVarietyType || (() => null);
  const _soVarieties = soVarieties || [];
  const _billMode = billMode || "final"; // "partial" or "final"
  const _selVars = selectedVarieties || [];

  // For partial bill: use earliest selected variety bill date for interest
  // For final bill: use earliest paid variety bill date
  const getAdvBillDate = () => {
    // Always use the manually set bill date if available
    if (so._billDate) return so._billDate;
    if (_billMode === "partial") {
      const dates = _selVars.map(v=>_getBillDate(v)).filter(Boolean);
      return dates.length>0 ? dates.reduce((a,b)=>a>b?a:b) : BILL_DATE;
    }
    const dates = _soVarieties.filter(v=>_isPaid(v)).map(v=>_getBillDate(v)).filter(Boolean);
    return dates.length>0 ? dates.reduce((a,b)=>a<b?a:b) : BILL_DATE;
  };
  const advCalc = (_billMode === "partial" ? [] : (so.advances || [])).map(a => {
    const advBillDate = a.tillDate || getAdvBillDate();
    const { days, interest } = (a.compound?calcCompoundInterest:calcInterest)(parseFloat(a.amount) || 0, parseFloat(a.interestRate) || 0, a.date, advBillDate);
    return { ...a, days, interest, total: (parseFloat(a.amount) || 0) + interest };
  });
  const totalAdv = advCalc.reduce((s, a) => s + (parseFloat(a.amount) || 0), 0);
  const totalAdvInt = advCalc.reduce((s, a) => s + a.interest, 0);
  const totalAdvWithInt = totalAdv + totalAdvInt;
  const growers = (_billMode === "partial"
    ? (so.growers || []).filter(g => _selVars.includes(g.variety))
    : (so.growers || [])
  ).slice().sort((a,b) => {
    const na = parseFloat(a.sNo)||0, nb = parseFloat(b.sNo)||0;
    if (na && nb) return na - nb;
    return String(a.sNo||"").localeCompare(String(b.sNo||""));
  });
  const passGrowers = growers.filter(g => g.result === "Pass");
  // Apply variety settings to growers
  const growersCalc = growers.map(g => {
    const paid = _isPaid(g.variety);
    const settled = _isSettled(g.variety);
    const toPay = _isToPay(g.variety);
    const vRate = _getRate(g.variety) || parseFloat(g.rate) || 0;
    const vType = _getType(g.variety) || g.type || "KMS";
    const amt = g.result === "Pass" ? (parseFloat(g.packets)||0) * vRate : 0;
    const pendingAmt = (g.result === "Pass" && !paid) ? amt : 0;
    // Only count "To Pay" in seed amount — settled already handled via Jamma
    const calcAmt = (g.result === "Pass" && paid) ? amt : 0;
    return { ...g, type: vType, rateUsed: vRate, calcAmt, pendingAmt, vPaid: paid, vSettled: settled, vToPay: toPay };
  });
  const totalSeedAmt = growersCalc.filter(g=>g.vPaid).reduce((s,g) => s + g.calcAmt, 0);
  const totalSettledAmt = growersCalc.filter(g=>g.vSettled&&g.result==="Pass").reduce((s,g)=>s+g.calcAmt,0);
  const totalToPayAmt = growersCalc.filter(g=>g.vToPay&&g.result==="Pass").reduce((s,g)=>s+g.calcAmt,0);
  const totalPendingAmt = growersCalc.reduce((s,g) => s + g.pendingAmt, 0);
  // Foundation from foundationSeeds (set in Excel Sheet2)
  const totalFoundation = (so.foundationSeeds||[]).reduce((s,fs) => s + (parseFloat(fs.area)||0) * 1000, 0);
  const totalTransport = passGrowers.reduce((s, g) => s + (parseFloat(g.packets) || 0), 0);
  // Jamma calculation — only in final bill
  const jammaCalcSO = (_billMode === "partial" ? [] : (so.jammaEntries||[])).map(j => {
    const amt = parseFloat(j.amount)||0;
    const jBillDate = j.tillDate || getAdvBillDate();
    const {interest} = calcInterest(amt, parseFloat(j.interestRate)||0, j.date||BILL_DATE, jBillDate);
    return {...j, interest, total: amt+interest};
  });
  const totalJammaSO = jammaCalcSO.reduce((s,j)=>s+(parseFloat(j.amount)||0),0);
  const totalJammaIntSO = jammaCalcSO.reduce((s,j)=>s+j.interest,0);
  const totalJammaWithIntSO = totalJammaSO + totalJammaIntSO;
  const balance = totalToPayAmt - totalAdvWithInt + totalJammaWithIntSO - totalFoundation - totalTransport;

  const TH = ({ ch }) => <th style={{ padding: "4px 6px", textAlign: "center", fontWeight: 600, fontSize: 11, whiteSpace: "nowrap" }}>{ch}</th>;
  const TD = ({ ch, s }) => <td style={{ padding: "3px 6px", textAlign: "center", fontSize: 11, whiteSpace: "nowrap", ...s }}>{ch}</td>;
  // Smaller version for growers table to fit A4
  const GTH = ({ ch }) => <th style={{ padding: "3px 4px", textAlign: "center", fontWeight: 600, fontSize: 9, whiteSpace: "nowrap" }}>{ch}</th>;
  const GTD = ({ ch, s }) => <td style={{ padding: "2px 4px", textAlign: "center", fontSize: 9, whiteSpace: "nowrap", ...s }}>{ch}</td>;

  return (
    <div className="bill-print" style={{ fontFamily: "Noto Serif Telugu,Georgia,serif", background: "#fff", border: "2px solid #2d5a8a", borderRadius: 4, width: "100%" }}>
      <div style={{ background: "linear-gradient(135deg,#1a2a4a,#2d5a8a)", color: "#fff", padding: "14px 20px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <span style={{ fontSize: 16, fontWeight: 700 }}>SUB-ORGANIZER BILL | సబ్ ఆర్గనైజర్ బిల్లు</span>
        <div style={{ textAlign: "right" }}>
          <div style={{ fontSize: 12, fontWeight: 600, background: "rgba(255,255,255,0.2)", padding: "3px 10px", borderRadius: 20 }}>Acc No: {so.accNo || "—"}</div>
          <div style={{ fontSize: 11, marginTop: 3, opacity: 0.85 }}>Date: {fmtDate(so._billDate||BILL_DATE)}</div>
        </div>
      </div>
      <div style={{ padding: "14px 20px" }}>
        <div style={{ background: "#f0f5ff", borderRadius: 6, padding: "10px 14px", marginBottom: 12, border: "1px solid #b0c8e0" }}>
          <div style={{ display: "grid", gridTemplateColumns: so.fatherName ? "1fr 1fr 1fr" : "1fr 1fr", gap: "5px 14px", fontSize: 13 }}>
            <div><span style={{ color: "#555" }}>Sub-Organizer: </span><strong>{so.name}</strong></div>
            {so.fatherName && <div><span style={{ color: "#555" }}>Father: </span><strong>{so.fatherName}</strong></div>}
            <div><span style={{ color: "#555" }}>Village (MSP): </span><strong>{so.village || "MSP"}</strong></div>
          </div>
        </div>
        {advCalc.length > 0 && (
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontWeight: 700, color: "#1a2a4a", marginBottom: 5, fontSize: 12 }}>ADVANCES</div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 460 }}>
                <thead><tr style={{ background: "#2d5a8a", color: "#fff" }}><TH ch="S.No" /><TH ch="Date" /><TH ch="Amount(₹)" /><TH ch="Days" /><TH ch="Interest(₹)" /><TH ch="Note" /><TH ch="Total(₹)" /></tr></thead>
                <tbody>
                  {advCalc.map((a, i) => (
                    <tr key={i} style={{ background: i % 2 === 0 ? "#f5faff" : "#fff", borderBottom: "1px solid #d0e4f4" }}>
                      <TD ch={i + 1} /><TD ch={fmtDate(a.date)} />
                      <TD ch={`₹${(parseFloat(a.amount) || 0).toLocaleString("en-IN")}`} />
                      <TD ch={a.days} />
                      <TD ch={`₹${a.interest.toLocaleString("en-IN")}`} s={{ color: "#c0392b" }} />
                      <TD ch={a.note || "—"} />
                      <TD ch={`₹${a.total.toLocaleString("en-IN")}`} s={{ fontWeight: 600 }} />
                    </tr>
                  ))}
                  <tr style={{ background: "#e8f0ff", fontWeight: 700 }}>
                    <td colSpan={2} style={{ padding: "4px 6px", fontSize: 11, color: "#1a2a4a" }}>TOTAL</td>
                    <TD ch={`₹${totalAdv.toLocaleString("en-IN")}`} /><TD ch="" />
                    <TD ch={`₹${totalAdvInt.toLocaleString("en-IN")}`} s={{ color: "#c0392b" }} /><TD ch="" />
                    <TD ch={`₹${totalAdvWithInt.toLocaleString("en-IN")}`} />
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        )}
        {growers.length > 0 && (
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontWeight: 700, color: "#1a2a4a", marginBottom: 5, fontSize: 12 }}>GROWERS SUMMARY | రైతుల సారాంశం</div>
            <table style={{ width: "100%", borderCollapse: "collapse", tableLayout: "fixed" }}>
              <colgroup>
                <col style={{ width: "4%" }} />
                <col style={{ width: "8%" }} />
                <col style={{ width: "16%" }} />
                <col style={{ width: "14%" }} />
                <col style={{ width: "14%" }} />
                <col style={{ width: "10%" }} />
                <col style={{ width: "6%" }} />
                <col style={{ width: "7%" }} />
                <col style={{ width: "6%" }} />
                <col style={{ width: "7%" }} />
                <col style={{ width: "8%" }} />
              </colgroup>
              <thead>
                <tr style={{ background: "#1a2a4a", color: "#fff" }}>
                  <GTH ch="S.No" /><GTH ch="LOT No" /><GTH ch="Grower" /><GTH ch="Father" />
                  <GTH ch="Village" /><GTH ch="Variety" /><GTH ch="Pkts" />
                  <GTH ch="Result" /><GTH ch="Type" /><GTH ch="Rate(₹)" /><GTH ch="Amt(₹)" />
                </tr>
              </thead>
              <tbody>
                {/* ── SETTLED GROWERS ── */}
                {growersCalc.filter(g => _isSettled(g.variety)).length > 0 && (
                  <tr><td colSpan={11} style={{ background:"#1a5c1a",color:"#fff",padding:"4px 6px",fontSize:9,fontWeight:700 }}>
                    ✔ SETTLED — Already paid to Sub-Org — {growersCalc.filter(g=>_isSettled(g.variety)&&g.result==="Pass").length} growers
                  </td></tr>
                )}
                {growersCalc.filter(g => _isSettled(g.variety)).map((g, i) => {
                  const ip = g.result === "Pass";
                  return (
                    <>
                    <tr key={"s"+i} style={{ background: !ip ? "#fdecea" : i%2===0?"#e8f5e8":"#f5fff5", borderBottom:"1px solid #b8ddb8", opacity:0.85 }}>
                      <GTD ch={i + 1} />
                      <GTD ch={g.lotNo||"—"} />
                      <GTD ch={<span>{g.name}{g.note?<span title={g.note} style={{marginLeft:3,fontSize:9,color:"#856404"}}>📝</span>:null}</span>} s={{textAlign:"left",overflow:"hidden",textOverflow:"ellipsis"}} />
                      <GTD ch={g.fatherName||"—"} s={{textAlign:"left",overflow:"hidden",textOverflow:"ellipsis"}} />
                      <GTD ch={g.village||"—"} s={{textAlign:"left",overflow:"hidden",textOverflow:"ellipsis"}} />
                      <GTD ch={g.variety} />
                      <GTD ch={g.packets} />
                      <GTD ch={ip?"✓P":"✗F"} s={{fontWeight:700,color:ip?"#155724":"#721c24"}} />
                      <GTD ch={g.type||"KMS"} s={{fontWeight:700,color:(g.type||"KMS")==="GMS"?"#856404":"#1a3a8a"}} />
                      <GTD ch={ip?`₹${(g.rateUsed||0).toLocaleString("en-IN")}`:"—"} />
                      <GTD ch={ip?`₹${g.calcAmt.toLocaleString("en-IN")}`:"—"} s={{fontWeight:600,color:"#1a5c1a"}} />
                    </tr>
                    {g.note && <tr key={"sn"+i}><td colSpan={11} style={{padding:"1px 8px 3px 28px",fontSize:10,color:"#856404",fontStyle:"italic",borderBottom:"1px solid #b8ddb8",background:"#fffdf0"}}>📝 {g.note}</td></tr>}
                    </>
                  );
                })}
                {/* ── TO PAY GROWERS ── */}
                {growersCalc.filter(g => _isToPay(g.variety)).length > 0 && (
                  <tr><td colSpan={11} style={{ background:"#2d5a8a",color:"#fff",padding:"4px 6px",fontSize:9,fontWeight:700 }}>
                    💰 TO PAY — Company paid us, need to pay Sub-Org — {growersCalc.filter(g=>_isToPay(g.variety)&&g.result==="Pass").length} growers
                  </td></tr>
                )}
                {growersCalc.filter(g => _isToPay(g.variety)).map((g, i) => {
                  const ip = g.result === "Pass";
                  const offset = growersCalc.filter(x=>_isSettled(x.variety)).length;
                  return (
                    <>
                    <tr key={"tp"+i} style={{ background: !ip?"#fdecea":i%2===0?"#f0fff0":"#fff", borderBottom:"1px solid #d4edd4" }}>
                      <GTD ch={offset + i + 1} />
                      <GTD ch={g.lotNo||"—"} />
                      <GTD ch={<span>{g.name}{g.note?<span title={g.note} style={{marginLeft:3,fontSize:9,color:"#856404"}}>📝</span>:null}</span>} s={{textAlign:"left",overflow:"hidden",textOverflow:"ellipsis"}} />
                      <GTD ch={g.fatherName||"—"} s={{textAlign:"left",overflow:"hidden",textOverflow:"ellipsis"}} />
                      <GTD ch={g.village||"—"} s={{textAlign:"left",overflow:"hidden",textOverflow:"ellipsis"}} />
                      <GTD ch={g.variety} />
                      <GTD ch={g.packets} />
                      <GTD ch={ip?"✓P":"✗F"} s={{fontWeight:700,color:ip?"#155724":"#721c24"}} />
                      <GTD ch={g.type||"KMS"} s={{fontWeight:700,color:(g.type||"KMS")==="GMS"?"#856404":"#1a3a8a"}} />
                      <GTD ch={ip?`₹${(g.rateUsed||0).toLocaleString("en-IN")}`:"—"} />
                      <GTD ch={ip?`₹${g.calcAmt.toLocaleString("en-IN")}`:"—"} s={{fontWeight:600,color:"#1a6a1a"}} />
                    </tr>
                    {g.note && <tr key={"tn"+i}><td colSpan={11} style={{padding:"1px 8px 3px 28px",fontSize:10,color:"#856404",fontStyle:"italic",borderBottom:"1px solid #d4edd4",background:"#fffdf0"}}>📝 {g.note}</td></tr>}
                    </>
                  );
                })}
                {/* ── PENDING GROWERS ── */}
                {growersCalc.filter(g => !_isPaid(g.variety)).length > 0 && (
                  <tr><td colSpan={11} style={{ background:"#856404",color:"#fff",padding:"4px 6px",fontSize:9,fontWeight:700 }}>
                    ⏳ PENDING — Company hasn't paid us yet — {growersCalc.filter(g=>!_isPaid(g.variety)&&g.result==="Pass").length} growers
                  </td></tr>
                )}
                {growersCalc.filter(g => !_isPaid(g.variety)).map((g, i) => {
                  const ip = g.result === "Pass";
                  const offset = growersCalc.filter(x=>_isPaid(x.variety)).length;
                  return (
                    <>
                    <tr key={"pd"+i} style={{ background:"#fff8e6", borderBottom:"1px solid #f0d080" }}>
                      <GTD ch={offset + i + 1} />
                      <GTD ch={g.lotNo||"—"} />
                      <GTD ch={<span>{g.name}{g.note?<span title={g.note} style={{marginLeft:3,fontSize:9,color:"#856404"}}>📝</span>:null}</span>} s={{textAlign:"left",overflow:"hidden",textOverflow:"ellipsis"}} />
                      <GTD ch={g.fatherName||"—"} s={{textAlign:"left",overflow:"hidden",textOverflow:"ellipsis"}} />
                      <GTD ch={g.village||"—"} s={{textAlign:"left",overflow:"hidden",textOverflow:"ellipsis"}} />
                      <GTD ch={g.variety} />
                      <GTD ch={g.packets} />
                      <GTD ch={ip?"✓P":"✗F"} s={{fontWeight:700,color:ip?"#856404":"#721c24"}} />
                      <GTD ch={g.type||"KMS"} s={{fontWeight:700,color:"#856404"}} />
                      <GTD ch={ip?`₹${(g.rateUsed||0).toLocaleString("en-IN")}`:"—"} s={{color:"#856404"}} />
                      <GTD ch={ip?"⏳":"—"} s={{fontWeight:600,color:"#856404"}} />
                    </tr>
                    {g.note && <tr key={"pn"+i} style={{background:"#fffdf0"}}><td colSpan={11} style={{padding:"1px 8px 3px 28px",fontSize:10,color:"#856404",fontStyle:"italic",borderBottom:"1px solid #f0d080"}}>📝 {g.note}</td></tr>}
                    </>
                  );
                })}
                <tr style={{ background: "#e8f0ff", fontWeight: 700 }}>
                  <td colSpan={6} style={{ padding:"3px 4px",fontSize:9,color:"#1a2a4a" }}>TOTAL</td>
                  <GTD ch={growers.reduce((s,g)=>s+(parseFloat(g.packets)||0),0)} />
                  <GTD ch={`${passGrowers.length}P/${growers.length-passGrowers.length}F`} />
                  <GTD ch="" /><GTD ch="" />
                  <GTD ch={`₹${totalToPayAmt.toLocaleString("en-IN")}`} s={{ color:"#2d5a8a" }} />
                </tr>
              </tbody>
            </table>
          </div>
        )}

        {/* ── JAMMA TABLE (final bill only) ── */}
        {_billMode !== "partial" && jammaCalcSO.length > 0 && (
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontWeight: 700, color: "#856404", marginBottom: 5, fontSize: 12 }}>JAMMA DETAILS (Partial Payments Received from Sub-Org)</div>
            <table style={{ width:"100%", borderCollapse:"collapse" }}>
              <thead>
                <tr style={{ background:"#856404", color:"#fff" }}>
                  {["S.No","Date","Note","Amount(₹)","Days","Interest(₹)","Total(₹)"].map(h=>(
                    <th key={h} style={{ padding:"4px 8px", textAlign:"center", fontSize:11, fontWeight:600 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {jammaCalcSO.map((j,i)=>(
                  <tr key={i} style={{ background:i%2===0?"#fff9e6":"#fff", borderBottom:"1px solid #f0d080" }}>
                    {[i+1, fmtDate(j.date||BILL_DATE), j.note||"—",
                      `₹${(parseFloat(j.amount)||0).toLocaleString("en-IN")}`,
                      j.days||"—",
                      `₹${(j.interest||0).toLocaleString("en-IN")}`,
                      `₹${(j.total||0).toLocaleString("en-IN")}`
                    ].map((v,ci)=>(
                      <td key={ci} style={{ padding:"3px 8px", textAlign:"center", fontSize:11, color: ci===5?"#856404":ci===6?"#1a6a1a":"#333", fontWeight: ci>=5?600:400 }}>{v}</td>
                    ))}
                  </tr>
                ))}
                <tr style={{ background:"#fff3cd", fontWeight:700 }}>
                  <td colSpan={3} style={{ padding:"4px 8px", fontSize:11, color:"#856404" }}>TOTAL JAMMA</td>
                  <td style={{ textAlign:"center", fontSize:11, color:"#856404" }}>₹{totalJammaSO.toLocaleString("en-IN")}</td>
                  <td></td>
                  <td style={{ textAlign:"center", fontSize:11, color:"#856404" }}>₹{totalJammaIntSO.toLocaleString("en-IN")}</td>
                  <td style={{ textAlign:"center", fontSize:11, color:"#1a6a1a", fontWeight:700 }}>₹{totalJammaWithIntSO.toLocaleString("en-IN")}</td>
                </tr>
              </tbody>
            </table>
          </div>
        )}

        {/* ── VARIETY SUMMARY TABLE (final bill only) ── */}
        {_billMode !== "partial" && (() => {
          // Build variety list in grower order (unique, preserving first appearance)
          const growerVarieties = [];
          (so.growers||[]).forEach(g => {
            if (g.variety && !growerVarieties.includes(g.variety)) growerVarieties.push(g.variety);
          });
          // Also add foundation-only varieties (zero growers) not already in list
          (so.foundationSeeds||[]).forEach(f => {
            if (f.variety && !growerVarieties.includes(f.variety)) growerVarieties.push(f.variety);
          });
          if (growerVarieties.length === 0) return null;
          const foundMap = {};
          (so.foundationSeeds||[]).forEach(f => { if(f.variety) foundMap[f.variety] = parseFloat(f.area)||0; });
          return (
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontWeight: 700, color: "#1a2a4a", marginBottom: 5, fontSize: 12 }}>VARIETY SUMMARY | వెరైటీ వివరాలు</div>
            <table style={{ width:"100%", borderCollapse:"collapse" }}>
              <thead>
                <tr style={{ background:"#1a2a4a", color:"#fff" }}>
                  {["S.No","Variety","Area (Ac)","Quantity (Pkts)","Status","Settled On","Notes"].map(h=>(
                    <th key={h} style={{ padding:"4px 10px", textAlign:"center", fontSize:11, fontWeight:600 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {growerVarieties.map((variety,i)=>{
                  const area = foundMap[variety] || 0;
                  const qty = (so.growers||[]).filter(g=>g.variety===variety&&g.result==="Pass").reduce((s,g)=>s+(parseFloat(g.packets)||0),0);
                  const isSettledVar = (settledVarieties||[]).includes(variety);
                  const isPaidVar = _isPaid(variety);
                  const status = isSettledVar ? "settled" : isPaidVar ? "topay" : "pending";
                  const statusLabel = status==="settled" ? "✔ Settled" : status==="topay" ? "💰 To Pay" : "⏳ Pending";
                  const statusColor = status==="settled" ? "#1a5c1a" : status==="topay" ? "#2d5a8a" : "#856404";
                  const statusBg = status==="settled" ? "#e8f5e9" : status==="topay" ? "#e8f0ff" : "#fff8e6";
                  const settledDate = (so._settledDates||{})[variety] || "";
                  const notes = (so._varietyNotes||{})[variety] || "";
                  return (
                    <tr key={i} style={{ background:i%2===0?"#f5f8ff":"#fff", borderBottom:"1px solid #d0e4f4" }}>
                      <td style={{ padding:"3px 10px", textAlign:"center", fontSize:11 }}>{i+1}</td>
                      <td style={{ padding:"3px 10px", textAlign:"left", fontSize:11, fontWeight:600 }}>{variety}</td>
                      <td style={{ padding:"3px 10px", textAlign:"center", fontSize:11 }}>{area > 0 ? `${area} Ac` : "—"}</td>
                      <td style={{ padding:"3px 10px", textAlign:"center", fontSize:11, fontWeight:600 }}>{qty.toLocaleString("en-IN")}</td>
                      <td style={{ padding:"3px 10px", textAlign:"center", fontSize:11, fontWeight:700, color:statusColor, background:statusBg }}>{statusLabel}</td>
                      <td style={{ padding:"3px 10px", textAlign:"center", fontSize:11 }}>{isSettledVar ? fmtDate(settledDate) : "—"}</td>
                      <td style={{ padding:"3px 10px", textAlign:"left", fontSize:11, color:"#555" }}>{notes||"—"}</td>
                    </tr>
                  );
                })}
                <tr style={{ background:"#e8f0ff", fontWeight:700 }}>
                  <td colSpan={2} style={{ padding:"4px 10px", fontSize:11, color:"#1a2a4a" }}>TOTAL</td>
                  <td style={{ textAlign:"center", fontSize:11 }}>{Object.values(foundMap).reduce((s,a)=>s+a,0)} Ac</td>
                  <td style={{ textAlign:"center", fontSize:11 }}>
                    {growerVarieties.reduce((s,variety)=>{
                      return s+(so.growers||[]).filter(g=>g.variety===variety&&g.result==="Pass").reduce((ss,g)=>ss+(parseFloat(g.packets)||0),0);
                    },0).toLocaleString("en-IN")}
                  </td>
                  <td colSpan={3}></td>
                </tr>
              </tbody>
            </table>
          </div>
          );
        })()}

        <div style={{ background: _billMode==="partial" ? "#f0f5ff" : (balance >= 0 ? "#e8f5e9" : "#fdecea"), borderRadius: 6, padding: "12px 16px", border: `2px solid ${_billMode==="partial"?"#2d5a8a":balance >= 0 ? "#2d6a2d" : "#e74c3c"}` }}>
          {_billMode === "partial" ? (
            <>
              <div style={{ fontWeight:700, color:"#2d5a8a", marginBottom:8, fontSize:13 }}>
                🧾 PARTIAL PAYMENT — {_selVars.join(", ")}
              </div>
              <div style={{ display:"grid", gridTemplateColumns:"1fr auto", gap:"4px 16px", fontSize:13 }}>
                <div style={{ color:"#555" }}>Seed Amount for Selected Varieties</div>
                <div style={{ textAlign:"right", fontWeight:800, fontSize:18, color:"#2d5a8a" }}>₹{totalSeedAmt.toLocaleString("en-IN")}</div>
                <div style={{ color:"#888", fontSize:11 }}>Note: Record this amount as Jamma when sub-org returns it</div>
                <div></div>
              </div>
            </>
          ) : (
            <>
              <div style={{ fontWeight: 700, color: "#1a2a4a", marginBottom: 8, fontSize: 13 }}>SETTLEMENT SUMMARY</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr auto", gap: "4px 16px", fontSize: 13 }}>

                {/* To Pay — main billing amount */}
                <div style={{ color: "#2d5a8a", fontWeight: 600 }}>💰 Seed Amount — To Pay Now</div>
                <div style={{ textAlign: "right", fontWeight: 700, color: "#2d5a8a" }}>₹{totalToPayAmt.toLocaleString("en-IN")}</div>

                {/* Settled — reference only */}
                {totalSettledAmt > 0 && <>
                  <div style={{ color: "#1a5c1a", fontSize: 12 }}>✔ Already Settled (previous bills)</div>
                  <div style={{ textAlign: "right", fontWeight: 600, color: "#1a5c1a", fontSize: 12 }}>₹{totalSettledAmt.toLocaleString("en-IN")}</div>
                </>}

                {/* Pending — not counted */}
                {totalPendingAmt > 0 && <>
                  <div style={{ color: "#856404", fontSize: 12 }}>⏳ Pending — Company hasn't paid yet (not counted)</div>
                  <div style={{ textAlign: "right", fontWeight: 600, color: "#856404", fontSize: 12 }}>₹{totalPendingAmt.toLocaleString("en-IN")}</div>
                </>}

                {/* Divider */}
                <div style={{ borderTop: "1px dashed #c0c0c0", gridColumn: "1/-1", margin: "4px 0" }}></div>

                {/* Deductions */}
                <div style={{ color: "#555" }}>Advance + Interest</div>
                <div style={{ textAlign: "right", fontWeight: 600, color: "#c0392b" }}>− ₹{totalAdvWithInt.toLocaleString("en-IN")}</div>

                {totalJammaWithIntSO > 0 && <>
                  <div style={{ color: "#1a6a1a" }}>Jamma + Interest (partial payments received)</div>
                  <div style={{ textAlign: "right", fontWeight: 600, color: "#1a6a1a" }}>+ ₹{totalJammaWithIntSO.toLocaleString("en-IN")}</div>
                </>}

                {totalFoundation > 0 && <>
                  <div style={{ color: "#555" }}>Foundation (see table above)</div>
                  <div style={{ textAlign: "right", fontWeight: 600, color: "#c0392b" }}>− ₹{totalFoundation.toLocaleString("en-IN")}</div>
                </>}

                {totalTransport > 0 && <>
                  <div style={{ color: "#555" }}>Transportation</div>
                  <div style={{ textAlign: "right", fontWeight: 600, color: "#c0392b" }}>− ₹{totalTransport.toLocaleString("en-IN")}</div>
                </>}

                {/* Final balance */}
                <div style={{ borderTop: "2px solid #2d5a8a", paddingTop: 6, marginTop: 4, fontWeight: 700, fontSize: 14, color: balance >= 0 ? "#1a4a1a" : "#c0392b" }}>
                  {balance >= 0 ? "Payable to Sub-Org (This Bill)" : "Due from Sub-Org (This Bill)"}
                </div>
                <div style={{ borderTop: "2px solid #2d5a8a", paddingTop: 6, marginTop: 4, textAlign: "right", fontWeight: 800, fontSize: 18, color: balance >= 0 ? "#1a4a1a" : "#c0392b" }}>
                  ₹{Math.abs(balance).toLocaleString("en-IN")}
                </div>
              </div>
            </>
          )}
        </div>
        {so.comment && so.comment.trim() && (
          <div style={{ marginTop: 14, background: "#fff9e6", border: "1.5px solid #c8a000", borderRadius: 6, padding: "10px 14px" }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: "#856404", marginBottom: 2 }}>📝 Note</div>
            <div style={{ fontSize: 12, color: "#5c4400" }}>{so.comment}</div>
          </div>
        )}
        <div style={{ marginTop: 16, display: "flex", justifyContent: "space-between", fontSize: 11, color: "#666", borderTop: "1px solid #d0e4f4", paddingTop: 10 }}>
          <div style={{ textAlign: "center" }}><div style={{ borderTop: "1px solid #aaa", width: 120, marginBottom: 4 }}></div><div>Sub-Organizer Signature</div></div>
          <div style={{ textAlign: "center", color: "#2d5a8a", fontSize: 12 }}><div>Thank you for your cooperation</div></div>
        </div>
      </div>
    </div>
  );
}

// ─── Farmer Form ────────────────────────────────────────────────
function FarmerForm({ farmer, index, onChange, onRemove, varietySettings, getVarietyRate, getVarietyType, farmers, subOrgs, pesticideList }) {
  const MAX_ADV = 10, MAX_CROP = 3;
  const inp = { style: { width: "100%", padding: "5px 8px", border: "1px solid #c8dfc8", borderRadius: 4, fontSize: 13, background: "#fafffe", boxSizing: "border-box" } };

  const updAdv = (i, f, v) => { const a = [...(farmer.advances || [])]; a[i] = { ...a[i], [f]: ["amount","interestRate"].includes(f) ? parseFloat(v)||0 : v }; onChange({ ...farmer, advances: a }); };
  const updCrop = (i, f, v) => { const c = [...(farmer.crops || [])]; c[i] = { ...c[i], [f]: ["quantity","ratePerUnit"].includes(f) ? parseFloat(v)||0 : v }; onChange({ ...farmer, crops: c }); };
  const updJamma = (i, f, v) => { const j = [...(farmer.jammaEntries || [])]; j[i] = { ...j[i], [f]: ["amount","interestRate"].includes(f) ? parseFloat(v)||0 : v }; onChange({ ...farmer, jammaEntries: j }); };
  const advCount = (farmer.advances || []).length;
  const cropCount = (farmer.crops || []).length;

  return (
    <div style={{ border: "1.5px solid #b8d8b8", borderRadius: 8, padding: 16, marginBottom: 14, background: "#f8fdf8" }} className="farmer-form-pad">
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, flexWrap:"wrap", gap:8 }}>
        <div style={{ fontWeight: 700, color: "#1a4a1a", fontSize: 14 }}>Farmer #{index + 1}</div>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <label style={{ display:"flex", alignItems:"center", gap:6, cursor:"pointer", background: farmer.billingDone?"#e8f5e9":"#fff3cd", border:`1.5px solid ${farmer.billingDone?"#2d6a2d":"#c8a000"}`, borderRadius:6, padding:"4px 10px" }}>
            <input type="checkbox" checked={!!farmer.billingDone}
              onChange={e=>onChange({...farmer, billingDone:e.target.checked, billingDoneDate: e.target.checked ? new Date().toISOString().split("T")[0] : ""})}
              style={{ width:14, height:14, cursor:"pointer" }}
            />
            <span style={{ fontSize:12, fontWeight:700, color: farmer.billingDone?"#2d6a2d":"#856404" }}>
              {farmer.billingDone ? `✔ Billed${farmer.billingDoneDate?" — "+fmtDate(farmer.billingDoneDate):""}` : "⏳ Billing Pending"}
            </span>
          </label>
          <button onClick={onRemove} style={{ background: "#e74c3c", color: "#fff", border: "none", borderRadius: 4, padding: "3px 10px", cursor: "pointer", fontSize: 12 }}>Remove</button>
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "0.6fr 1fr 1fr 1fr", gap: 8, marginBottom: 8 }} className="mobile-grid-4">
        <div><label style={{ fontSize: 11, color: "#555", display: "block", marginBottom: 2 }}>Farmer No</label><input {...inp} value={farmer.farmerNo||""} onChange={e => onChange({...farmer,farmerNo:e.target.value})} placeholder="001" style={{...inp.style,fontWeight:700}} /></div>
        {[["name","Name"],["fatherName","Father"],["village","Village"]].map(([f,l]) => (
          <div key={f}><label style={{ fontSize: 11, color: "#555", display: "block", marginBottom: 2 }}>{l}</label><input {...inp} value={farmer[f]||""} onChange={e => onChange({...farmer,[f]:e.target.value})} /></div>
        ))}
      </div>
      <div style={{ marginBottom: 8 }}>
        <label style={{ fontSize: 11, color: "#888", display: "block", marginBottom: 2 }}>C/o (Care of) — optional, only if someone else is responsible</label>
        <input {...inp} value={farmer.careOf||""} onChange={e => onChange({...farmer,careOf:e.target.value})} placeholder="e.g. Ramesh (in-charge)" style={{...inp.style, maxWidth:320}} />
      </div>
      <div style={{ marginBottom: 12 }}>
        <input
          placeholder="📝 Comment for this farmer's bill (optional) — e.g. Quality was low, Need to pay the due, Someone else will pay..."
          value={farmer.comment||""}
          onChange={e=>onChange({...farmer,comment:e.target.value})}
          style={{width:"100%",padding:"7px 10px",border:"1.5px dashed #c8a000",borderRadius:5,fontSize:12,background:"#fffdf5",color:"#856404"}}
        />
      </div>

      {/* Advances */}
      <div style={{ marginBottom: 12 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
          <span style={{ fontWeight: 600, fontSize: 12, color: "#2d6a2d" }}>Advances / అడ్వాన్స్లు</span>
          <span style={{ fontSize: 11, background: advCount>=MAX_ADV?"#fdecea":"#e8f5e9", color: advCount>=MAX_ADV?"#c0392b":"#1a4a1a", padding:"1px 8px", borderRadius:10, fontWeight:600 }}>{advCount}/{MAX_ADV}</span>
        </div>
        {(farmer.advances||[]).map((a,i) => (
          <div key={i} style={{ marginBottom:8 }}>
            <div style={{ display:"grid", gridTemplateColumns:"1.2fr 1fr 0.7fr 1fr 32px", gap:6, alignItems:"end" }} className="mobile-adv-grid">
              <div><label style={{fontSize:10,color:"#666",display:"block"}}>Date</label><input {...inp} type="date" value={a.date} onChange={e=>updAdv(i,"date",e.target.value)} /></div>
              <div><label style={{fontSize:10,color:"#666",display:"block"}}>Amount ₹</label><input {...inp} type="number" value={a.amount} onChange={e=>updAdv(i,"amount",e.target.value)} /></div>
              <div><label style={{fontSize:10,color:"#666",display:"block"}}>Interest %</label><input {...inp} type="number" step="0.5" value={a.interestRate} onChange={e=>updAdv(i,"interestRate",e.target.value)} /></div>
              <div><label style={{fontSize:10,color:"#666",display:"block"}}>Note</label><input {...inp} value={a.note||""} onChange={e=>updAdv(i,"note",e.target.value)} /></div>
              <button onClick={()=>onChange({...farmer,advances:farmer.advances.filter((_,j)=>j!==i)})} style={{background:"#fdecea",color:"#e74c3c",border:"1px solid #e74c3c",borderRadius:4,padding:"5px 6px",cursor:"pointer",fontSize:11}}>✕</button>
            </div>
            {/* Till Date override */}
            <div style={{display:"flex",alignItems:"center",gap:8,marginTop:4,paddingLeft:4}}>
              {a.tillDate ? (
                <>
                  <label style={{fontSize:10,color:"#e67e22",fontWeight:600}}>⚠ Interest Till:</label>
                  <input type="date" value={a.tillDate} onChange={e=>updAdv(i,"tillDate",e.target.value)}
                    style={{padding:"2px 6px",border:"1px solid #e67e22",borderRadius:4,fontSize:11,color:"#e67e22"}} />
                  <button onClick={()=>updAdv(i,"tillDate","")} style={{fontSize:10,padding:"2px 8px",borderRadius:4,border:"1px solid #aaa",background:"#fff",color:"#888",cursor:"pointer"}}>↩ Use Bill Date</button>
                  <span style={{fontSize:10,color:"#888"}}>
                    ({Math.max(0,Math.round((parseLocalDate(a.tillDate)-parseLocalDate(a.date))/86400000))} days)
                  </span>
                </>
              ) : (
                <button onClick={()=>updAdv(i,"tillDate",a.date)} style={{fontSize:10,padding:"2px 8px",borderRadius:4,border:"1px dashed #e67e22",background:"#fff9f0",color:"#e67e22",cursor:"pointer"}}>
                  + Set custom interest end date
                </button>
              )}
              {/* Compound interest toggle — only show if advance is 1+ year old */}
              {(()=>{
                const advDays = Math.round((parseLocalDate(a.tillDate||BILL_DATE)-parseLocalDate(a.date))/86400000);
                if (advDays < 365) return null;
                return (
                  <button onClick={()=>updAdv(i,"compound",!a.compound)}
                    style={{fontSize:10,padding:"2px 10px",borderRadius:4,border:a.compound?"1px solid #6a0dad":"1px dashed #9b59b6",background:a.compound?"#f3e5ff":"#fdf5ff",color:a.compound?"#6a0dad":"#9b59b6",cursor:"pointer",fontWeight:a.compound?700:400}}>
                    {a.compound?"✓ Compound Yearly":"⟳ Compound Yearly?"}
                  </button>
                );
              })()}
            </div>
          </div>
        ))}
        <button disabled={advCount>=MAX_ADV} onClick={()=>onChange({...farmer,advances:[...(farmer.advances||[]),{date:new Date().toISOString().split("T")[0],amount:0,interestRate:24,note:""}]})} style={{background:advCount>=MAX_ADV?"#eee":"#e8f5e9",color:advCount>=MAX_ADV?"#999":"#2d6a2d",border:`1px dashed ${advCount>=MAX_ADV?"#ccc":"#2d6a2d"}`,borderRadius:4,padding:"4px 12px",cursor:advCount>=MAX_ADV?"not-allowed":"pointer",fontSize:12}}>{advCount>=MAX_ADV?"Max 10 reached":"+ Add Advance"}</button>
        {pesticideList && pesticideList.filter(p=>p.name&&(p.sizes||[]).some(s=>s.size||s.price)).length > 0 && advCount < MAX_ADV && (
          <div style={{marginTop:8,background:"#fff8f0",border:"1px solid #f0a040",borderRadius:6,padding:"8px 10px"}}>
            <div style={{fontSize:11,color:"#b35c00",fontWeight:700,marginBottom:6}}>🧪 Add Pesticide as Advance</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 0.7fr auto",gap:6,alignItems:"end"}}>
              <div>
                <label style={{fontSize:10,color:"#888",display:"block",marginBottom:2}}>Date</label>
                <input id={`pest-date-${farmer.id}`} type="date"
                  defaultValue={new Date().toISOString().split("T")[0]}
                  style={{width:"100%",padding:"6px 8px",border:"1px solid #f0a040",borderRadius:4,fontSize:13}} />
              </div>
              <div>
                <label style={{fontSize:10,color:"#888",display:"block",marginBottom:2}}>Pesticide &amp; Size</label>
                <select id={`pest-sel-${farmer.id}`} style={{width:"100%",padding:"6px 8px",border:"1px solid #f0a040",borderRadius:4,fontSize:13,background:"#fffdf5"}}>
                  <option value="">— Select —</option>
                  {pesticideList.filter(p=>p.name).flatMap((p,pi)=>
                    (p.sizes||[]).filter(s=>s.size||s.price).map((s,si)=>(
                      <option key={`${pi}-${si}`} value={`${pi}-${si}`}>{p.name} {s.size} (₹{s.price})</option>
                    ))
                  )}
                </select>
              </div>
              <div>
                <label style={{fontSize:10,color:"#888",display:"block",marginBottom:2}}>Packets</label>
                <input id={`pest-qty-${farmer.id}`} type="number" min="1" defaultValue="1"
                  style={{width:"100%",padding:"6px 8px",border:"1px solid #f0a040",borderRadius:4,fontSize:13}} />
              </div>
              <button onClick={()=>{
                const sel = document.getElementById(`pest-sel-${farmer.id}`);
                const qtyEl = document.getElementById(`pest-qty-${farmer.id}`);
                const dateEl = document.getElementById(`pest-date-${farmer.id}`);
                const val = sel.value;
                const qty = parseFloat(qtyEl.value)||1;
                const date = dateEl.value || new Date().toISOString().split("T")[0];
                if (!val) { alert("Please select a pesticide"); return; }
                const [pi, si] = val.split("-").map(Number);
                const p = pesticideList[pi];
                const s = p.sizes[si];
                const amount = s.price * qty;
                const note = `${p.name} ${s.size} ×${qty}`;
                onChange({...farmer, advances:[...(farmer.advances||[]),{
                  date, amount, interestRate: 0, note
                }]});
                sel.value=""; qtyEl.value="1";
              }} style={{background:"#e67e22",color:"#fff",border:"none",borderRadius:4,padding:"6px 14px",fontSize:13,fontWeight:700,cursor:"pointer",whiteSpace:"nowrap",marginTop:16}}>
                ➕ Add
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Crops */}
      <div style={{ marginBottom: 12 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
          <span style={{ fontWeight: 600, fontSize: 12, color: "#2d6a2d" }}>Crops / పంటలు</span>
          <span style={{ fontSize:11, background:cropCount>=MAX_CROP?"#fdecea":"#e8f5e9", color:cropCount>=MAX_CROP?"#c0392b":"#1a4a1a", padding:"1px 8px", borderRadius:10, fontWeight:600 }}>{cropCount}/{MAX_CROP}</span>
        </div>
        {(farmer.crops||[]).map((c,i) => {
          const ip = c.result==="Pass";
          const area = parseFloat(c.area)||0;
          const qty = parseFloat(c.quantity)||0;
          return (
            <div key={i} style={{background:"#fff",border:"1px solid #c8dfc8",borderRadius:6,padding:"8px 10px",marginBottom:8}}>
              <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 0.7fr 0.7fr 1fr",gap:6,alignItems:"end",marginBottom:6}} className="mobile-crop-grid">
                <div><label style={{fontSize:10,color:"#666",display:"block"}}>Crop Variety</label><input {...inp} value={c.variety} onChange={e=>{
                  const variety = e.target.value;
                  const vpRate = getVarietyRate ? getVarietyRate(variety) : null;
                  const vpType = getVarietyType ? getVarietyType(variety) : null;
                  const vsEntry = varietySettings && varietySettings[variety];
                  const a=[...(farmer.crops||[])];
                  a[i]={...a[i], variety,
                    ...(vsEntry?.rate ? {ratePerUnit: parseFloat(vsEntry.rate)} : {}),
                    ...(vsEntry?.type ? {cropType: vsEntry.type} : {})
                  };
                  onChange({...farmer,crops:a});
                }} /></div>
                <div><label style={{fontSize:10,color:"#666",display:"block"}}>Lot No</label><input {...inp} value={c.lotNo||""} onChange={e=>updCrop(i,"lotNo",e.target.value)} placeholder="e.g. 815841" style={{
                  ...inp.style,
                  borderColor: (() => {
                    try {
                      const lot = (c.lotNo||"").trim();
                      if (!lot) return "#c8dfc8";
                      const usedByFarmer = (farmers||[]).some((f,fi) => fi !== index && (f.crops||[]).some(cr => (cr.lotNo||"").trim() === lot));
                      const usedBySubOrg = (subOrgs||[]).some(so => (so.growers||[]).some(g => (g.lotNo||"").trim() === lot));
                      return (usedByFarmer || usedBySubOrg) ? "#e74c3c" : "#2d6a2d";
                    } catch { return "#c8dfc8"; }
                  })(),
                  boxShadow: (() => {
                    try {
                      const lot = (c.lotNo||"").trim();
                      if (!lot) return "none";
                      const usedByFarmer = (farmers||[]).some((f,fi) => fi !== index && (f.crops||[]).some(cr => (cr.lotNo||"").trim() === lot));
                      const usedBySubOrg = (subOrgs||[]).some(so => (so.growers||[]).some(g => (g.lotNo||"").trim() === lot));
                      return (usedByFarmer || usedBySubOrg) ? "0 0 0 2px rgba(231,76,60,0.3)" : "none";
                    } catch { return "none"; }
                  })()
                }} /></div>
                <div><label style={{fontSize:10,color:"#666",display:"block"}}>Area (Acres)</label><input {...inp} type="number" step="0.1" value={c.area||""} onChange={e=>updCrop(i,"area",e.target.value)} /></div>
                <div><label style={{fontSize:10,color:"#666",display:"block"}}>Qty</label><input {...inp} type="number" value={c.quantity} onChange={e=>updCrop(i,"quantity",e.target.value)} /></div>
                <div>
                  <label style={{fontSize:10,color:"#666",display:"block"}}>Type</label>
                  {(() => {
                    const vsEntry = varietySettings?.[c.variety];
                    const vpType = vsEntry?.type;
                    if (vpType) {
                      // From Variety Pay — show as read-only box
                      return (
                        <div>
                          <div style={{...inp.style, background:"#f0fff0", borderColor:"#2d6a2d", color:"#1a4a1a", fontWeight:700, display:"flex", alignItems:"center", padding:"5px 8px", fontSize:13}}>
                            {vpType}
                          </div>
                          <div style={{fontSize:9,color:"#2d6a2d",marginTop:2}}>✅ Variety Pay</div>
                        </div>
                      );
                    }
                    // No Variety Pay setting — show simple select
                    return (
                      <select value={c.cropType||"KMS"} onChange={e=>updCrop(i,"cropType",e.target.value)}
                        style={{...inp.style, cursor:"pointer"}}>
                        <option value="KMS">KMS</option>
                        <option value="GMS">GMS</option>
                      </select>
                    );
                  })()}
                </div>
                {ip ? (() => {
                  const vpRate = varietySettings?.[c.variety]?.rate ? parseFloat(varietySettings[c.variety].rate) : null;
                  const isOverride = c.rateOverride === true;
                  const displayRate = (!isOverride && vpRate) ? vpRate : (c.ratePerUnit||0);
                  return (
                    <div>
                      <label style={{fontSize:10,color:isOverride?"#e67e22":"#666",fontWeight:isOverride?700:400,display:"block"}}>
                        Rate ₹ {vpRate && !isOverride ? <span style={{color:"#2d6a2d",fontWeight:700}}>= ₹{vpRate} (Variety Pay ✅)</span> : isOverride ? "⚠ Manual Override" : ""}
                      </label>
                      {vpRate && !isOverride ? (
                        <div style={{display:"flex",alignItems:"center",gap:6}}>
                          <div style={{...inp.style,background:"#f0fff0",borderColor:"#2d6a2d",color:"#1a4a1a",fontWeight:700,display:"flex",alignItems:"center",padding:"5px 8px",fontSize:13,minWidth:60}}>₹{vpRate}</div>
                          <button onClick={()=>{const a=[...(farmer.crops||[])];a[i]={...a[i],rateOverride:true,ratePerUnit:vpRate};onChange({...farmer,crops:a});}} style={{fontSize:10,padding:"4px 8px",borderRadius:4,border:"1px solid #e67e22",background:"#fff9f0",color:"#e67e22",cursor:"pointer",whiteSpace:"nowrap"}}>✏ Override</button>
                        </div>
                      ) : (
                        <div style={{display:"flex",alignItems:"center",gap:6}}>
                          <input {...inp} type="number" value={c.ratePerUnit||0} onChange={e=>updCrop(i,"ratePerUnit",e.target.value)} style={{...inp.style,borderColor:"#e67e22",background:"#fff9f0"}} />
                          {vpRate && <button onClick={()=>{const a=[...(farmer.crops||[])];a[i]={...a[i],rateOverride:false,ratePerUnit:vpRate};onChange({...farmer,crops:a});}} style={{fontSize:10,padding:"4px 8px",borderRadius:4,border:"1px solid #2d6a2d",background:"#f0fff0",color:"#2d6a2d",cursor:"pointer",whiteSpace:"nowrap"}}>↩ Use VP</button>}
                        </div>
                      )}
                    </div>
                  );
                })() : <div style={{display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,color:"#aaa",fontStyle:"italic",paddingTop:16}}>No rate (Fail)</div>
                }
              </div>
              <div style={{display:"flex",alignItems:"center",gap:8,justifyContent:"space-between",flexWrap:"wrap"}}>
                <div style={{display:"flex",alignItems:"center",gap:6}}>
                  <span style={{fontSize:11,fontWeight:600,color:"#555"}}>Result:</span>
                  <button onClick={()=>updCrop(i,"result","Pass")} style={{padding:"3px 12px",borderRadius:10,border:"none",fontWeight:700,fontSize:11,cursor:"pointer",background:ip?"#155724":"#e8f5e9",color:ip?"#fff":"#155724"}}>✓ Pass</button>
                  <button onClick={()=>updCrop(i,"result","Fail")} style={{padding:"3px 12px",borderRadius:10,border:"none",fontWeight:700,fontSize:11,cursor:"pointer",background:!ip?"#721c24":"#fdecea",color:!ip?"#fff":"#721c24"}}>✗ Fail</button>
                </div>
                <div style={{fontSize:11,padding:"2px 10px",background:ip?"#e8f5e9":"#f5f5f5",borderRadius:5,color:ip?"#1a4a1a":"#999"}}>
                  {ip ? <strong>{qty} × ₹{(c.ratePerUnit||0).toLocaleString("en-IN")} = ₹{(qty*(c.ratePerUnit||0)).toLocaleString("en-IN")}</strong> : <em>No amount — Fail</em>}
                </div>
                <button onClick={()=>onChange({...farmer,crops:farmer.crops.filter((_,j)=>j!==i)})} style={{background:"#fdecea",color:"#e74c3c",border:"1px solid #e74c3c",borderRadius:4,padding:"4px 8px",cursor:"pointer",fontSize:11}}>✕ Remove</button>
              </div>
              <div style={{marginTop:6}}>
                <input
                  placeholder="📝 Note for this crop (optional) — e.g. Low quality LOT, Rejected by company..."
                  value={c.note||""}
                  onChange={e=>updCrop(i,"note",e.target.value)}
                  style={{width:"100%",padding:"5px 8px",border:"1px dashed #c8a000",borderRadius:4,fontSize:11,background:"#fffdf5",color:"#856404"}}
                />
              </div>
              <div style={{marginTop:5,fontSize:11,padding:"3px 8px",background:"#fdecea",borderRadius:4,color:"#721c24",display:"flex",gap:14,flexWrap:"wrap"}}>
                <span>⚠ Always charged:</span>
                <span>Foundation: {area>0?`${area} × ₹1000 = ₹${(area*1000).toLocaleString("en-IN")}`:"Enter area"}</span>
                <span>Transport: {qty} × ₹1 = <strong>₹{qty.toLocaleString("en-IN")}</strong></span>
              </div>
            </div>
          );
        })}
        <button disabled={cropCount>=MAX_CROP} onClick={()=>onChange({...farmer,crops:[...(farmer.crops||[]),{variety:"",area:"",quantity:0,cropType:"KMS",ratePerUnit:550,result:"Fail"}]})} style={{background:cropCount>=MAX_CROP?"#eee":"#e8f5e9",color:cropCount>=MAX_CROP?"#999":"#2d6a2d",border:`1px dashed ${cropCount>=MAX_CROP?"#ccc":"#2d6a2d"}`,borderRadius:4,padding:"4px 12px",cursor:cropCount>=MAX_CROP?"not-allowed":"pointer",fontSize:12}}>{cropCount>=MAX_CROP?"Max 3 crops":"+ Add Crop"}</button>
      </div>

      {/* Jamma */}
      <div style={{borderTop:"1.5px dashed #c8a000",paddingTop:10}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:farmer.jammaEnabled?8:0}}>
          <span style={{fontWeight:700,fontSize:13,color:"#856404"}}>జమ్మా / Jamma</span>
          <span style={{fontSize:11,color:"#666"}}>(Farmer pays us)</span>
          <button onClick={()=>onChange({...farmer,jammaEnabled:!farmer.jammaEnabled,jammaEntries:farmer.jammaEntries||[]})} style={{padding:"3px 14px",borderRadius:12,border:"none",fontWeight:700,fontSize:12,cursor:"pointer",background:farmer.jammaEnabled?"#856404":"#f5f5f5",color:farmer.jammaEnabled?"#fff":"#856404",outline:farmer.jammaEnabled?"none":"1px solid #c8a000"}}>{farmer.jammaEnabled?"✓ Enabled":"Enable Jamma"}</button>
        </div>
        {farmer.jammaEnabled && (
          <div style={{background:"#fffdf0",border:"1px solid #ffeeba",borderRadius:6,padding:"10px 12px"}}>
            {(farmer.jammaEntries||[]).map((j,i) => {
              const amt=parseFloat(j.amount)||0, rate=parseFloat(j.interestRate)||0;
              const jamTillDate = j.tillDate || "";
              const {days,interest}=calcInterest(amt,rate,j.date||BILL_DATE, jamTillDate||BILL_DATE);
              return (
                <div key={i} style={{marginBottom:8,background:"#fff",border:"1px solid #ffeeba",borderRadius:5,padding:"8px 10px"}}>
                  <div style={{display:"grid",gridTemplateColumns:"1.2fr 1fr 0.7fr 1fr 32px",gap:6,alignItems:"end",marginBottom:5}}>
                    <div><label style={{fontSize:10,color:"#666",display:"block"}}>Date</label><input {...inp} type="date" value={j.date||""} onChange={e=>updJamma(i,"date",e.target.value)} /></div>
                    <div><label style={{fontSize:10,color:"#666",display:"block"}}>Amount ₹</label><input {...inp} type="number" value={j.amount||""} onChange={e=>updJamma(i,"amount",e.target.value)} /></div>
                    <div><label style={{fontSize:10,color:"#666",display:"block"}}>Interest %</label><input {...inp} type="number" step="0.5" value={j.interestRate||""} onChange={e=>updJamma(i,"interestRate",e.target.value)} /></div>
                    <div><label style={{fontSize:10,color:"#666",display:"block"}}>Note</label><input {...inp} value={j.note||""} onChange={e=>updJamma(i,"note",e.target.value)} /></div>
                    <button onClick={()=>onChange({...farmer,jammaEntries:farmer.jammaEntries.filter((_,k)=>k!==i)})} style={{background:"#fdecea",color:"#e74c3c",border:"1px solid #e74c3c",borderRadius:4,padding:"5px 6px",cursor:"pointer",fontSize:11}}>✕</button>
                  </div>
                  {/* Till Date override */}
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:5}}>
                    {jamTillDate ? (
                      <>
                        <label style={{fontSize:10,color:"#e67e22",fontWeight:600}}>⚠ Interest Till:</label>
                        <input type="date" value={jamTillDate} onChange={e=>updJamma(i,"tillDate",e.target.value)}
                          style={{padding:"2px 6px",border:"1px solid #e67e22",borderRadius:4,fontSize:11,color:"#e67e22"}} />
                        <button onClick={()=>updJamma(i,"tillDate","")} style={{fontSize:10,padding:"2px 8px",borderRadius:4,border:"1px solid #aaa",background:"#fff",color:"#888",cursor:"pointer"}}>↩ Use Bill Date</button>
                        <span style={{fontSize:10,color:"#888"}}>({Math.max(0,Math.round((parseLocalDate(jamTillDate)-parseLocalDate(j.date||BILL_DATE))/86400000))} days)</span>
                      </>
                    ) : (
                      <button onClick={()=>updJamma(i,"tillDate",j.date||BILL_DATE)} style={{fontSize:10,padding:"2px 8px",borderRadius:4,border:"1px dashed #e67e22",background:"#fff9f0",color:"#e67e22",cursor:"pointer"}}>
                        + Set custom interest end date
                      </button>
                    )}
                  </div>
                  <div style={{fontSize:11,padding:"2px 8px",background:"#fff3cd",borderRadius:4,color:"#856404"}}>
                    ₹{amt.toLocaleString("en-IN")} × {rate}% × {days} days = Interest: ₹{interest.toLocaleString("en-IN")} | <strong>Total: ₹{(amt+interest).toLocaleString("en-IN")}</strong>
                  </div>
                </div>
              );
            })}
            <button onClick={()=>onChange({...farmer,jammaEntries:[...(farmer.jammaEntries||[]),{date:new Date().toISOString().split("T")[0],amount:0,interestRate:0,note:""}]})} style={{background:"#fff3cd",color:"#856404",border:"1px dashed #c8a000",borderRadius:4,padding:"4px 12px",cursor:"pointer",fontSize:12}}>+ Add Jamma Entry</button>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── Main App ───────────────────────────────────────────────────
export default function App() {
  const [farmers, setFarmers] = useState(null);
  const [BILL_DATE, setBILL_DATE_STATE] = useState(()=>{
    try { return localStorage.getItem("app_bill_date")||"2026-07-01"; } catch { return "2026-07-01"; }
  });
  const setBILL_DATE = (d) => { setBILL_DATE_STATE(d); try { localStorage.setItem("app_bill_date",d); BILL_DATE = d; } catch {} };
  const [showSettings, setShowSettings] = useState(false);
  const [cloudStatus, setCloudStatus] = useState("idle"); // idle | saving | saved | error
  const [cloudLastSaved, setCloudLastSaved] = useState("");
  const [snapshots, setSnapshots] = useState([]);
  const [showRestore, setShowRestore] = useState(false);
  const cloudSaveTimer = useRef(null);
  const [undoItem, setUndoItem] = useState(null); // {type, data, idx, timeout}
  const deleteWithUndo = (type, idx) => {
    if (type === "farmer") {
      const deleted = farmers[idx];
      const newList = farmers.filter((_,i)=>i!==idx);
      updateFarmers(newList);
      setSelectedIdx(0);
      if (undoItem?.timeout) clearTimeout(undoItem.timeout);
      const t = setTimeout(()=>setUndoItem(null), 8000);
      setUndoItem({type:"farmer", data:deleted, idx, timeout:t});
    } else {
      const deleted = subOrgs[idx];
      const newList = subOrgs.filter((_,i)=>i!==idx);
      updateSubOrgs(newList);
      setSelectedSubOrgIdx(0);
      if (undoItem?.timeout) clearTimeout(undoItem.timeout);
      const t = setTimeout(()=>setUndoItem(null), 8000);
      setUndoItem({type:"suborg", data:deleted, idx, timeout:t});
    }
  };
  const undoDelete = () => {
    if (!undoItem) return;
    clearTimeout(undoItem.timeout);
    if (undoItem.type === "farmer") {
      const newList = [...farmers];
      newList.splice(undoItem.idx, 0, undoItem.data);
      updateFarmers(newList);
      setSelectedIdx(undoItem.idx);
    } else {
      const newList = [...subOrgs];
      newList.splice(undoItem.idx, 0, undoItem.data);
      updateSubOrgs(newList);
      setSelectedSubOrgIdx(undoItem.idx);
    }
    setUndoItem(null);
  };
  const [lastBackupDate, setLastBackupDate] = useState(()=>{
    try { return localStorage.getItem("last_backup_date")||""; } catch { return ""; }
  });
  const markBackupDone = () => {
    const today = new Date().toISOString().split("T")[0];
    setLastBackupDate(today);
    try { localStorage.setItem("last_backup_date", today); } catch {}
  };
  const daysSinceBackup = lastBackupDate ? Math.floor((new Date()-new Date(lastBackupDate))/86400000) : 999;
  const [subOrgs, setSubOrgs] = useState([]);
  const [mode, setMode] = useState("farmers");
  const [selectedVillage, setSelectedVillage] = useState(null);
  const [villageSearch, setVillageSearch] = useState("");
  const [selectedCareOf, setSelectedCareOf] = useState(null);
  const [selectedIdx, setSelectedIdx] = useState(0);
  const [tab, setTab] = useState("form");
  const [saveStatus, setSaveStatus] = useState("idle");
  const [villagesPanelOpen, setVillagesPanelOpen] = useState(true);
  const [farmersPanelOpen, setFarmersPanelOpen] = useState(true);
  const [farmerSearch, setFarmerSearch] = useState("");
  const [farmerPage, setFarmerPage] = useState(0);
  const FARMERS_PER_PAGE = 25;
  const [selectedSubOrgIdx, setSelectedSubOrgIdx] = useState(0);
  const [subOrgTab, setSubOrgTab] = useState("form");
  const [subOrgPanelOpen, setSubOrgPanelOpen] = useState(true);
  const [printQueue, setPrintQueue] = useState([]);
  const [printQueueIdx, setPrintQueueIdx] = useState(-1);
  const [printQueueTotal, setPrintQueueTotal] = useState(0);

  const [billingHistory, setBillingHistory] = useState(() => {
    try { const s=localStorage.getItem("billing_history"); return s?JSON.parse(s):{}; } catch { return {}; }
  });
  const saveBillingHistory = (h) => {
    setBillingHistory(h);
    try { localStorage.setItem("billing_history",JSON.stringify(h)); } catch {}
  };
  const recordBillingHistory = (soAccNo, paidVarieties) => {
    const key = soAccNo;
    const prev = billingHistory[key] || [];
    const entry = {
      date: new Date().toLocaleDateString("en-IN"),
      time: new Date().toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit"}),
      varieties: paidVarieties,
    };
    saveBillingHistory({...billingHistory,[key]:[...prev,entry]});
  };
  const [selectedPrintVarieties, setSelectedPrintVarieties] = useState([]); // empty = all varieties
  const [pesticideList, setPesticideList] = useState(() => {
    try { return JSON.parse(localStorage.getItem("pesticide_list") || "[]"); } catch { return []; }
  });
  const savePesticideList = (list) => {
    setPesticideList(list);
    localStorage.setItem("pesticide_list", JSON.stringify(list));
  };
  const [varietySettings, setVarietySettings] = useState(() => {
    try { const s = localStorage.getItem("variety_settings"); return s ? JSON.parse(s) : {}; } catch { return {}; }
  });
  const saveVarietySettings = (vs) => {
    setVarietySettings(vs);
    try { localStorage.setItem("variety_settings", JSON.stringify(vs)); } catch {}
    // Save to Firebase
    if (cloudSaveTimer.current) clearTimeout(cloudSaveTimer.current);
    cloudSaveTimer.current = setTimeout(async () => {
      setCloudStatus("saving");
      const saved = await saveToCloud(farmers||[], subOrgs||[], vs);
      setCloudStatus(saved ? "saved" : "error");
      setTimeout(() => setCloudStatus("idle"), 3000);
    }, 2000);
  };
  // Get all unique varieties from farmers
  const allVarieties = [...new Set([
    ...(farmers||[]).flatMap(f => (f.crops||[]).map(c => c.variety).filter(Boolean)),
    ...(subOrgs||[]).flatMap(so => (so.growers||[]).map(g => g.variety).filter(Boolean))
  ])].sort();
  const allSubOrgVarieties = [...new Set((subOrgs||[]).flatMap(so => [
    ...(so.growers||[]).map(g => g.variety),
    ...(so.foundationSeeds||[]).map(f => f.variety)
  ].filter(Boolean)))].sort();
  // Sub-org variety settings stored separately under "so_" prefix in varietySettings
  // Status: "pending" = company hasn't paid | "paid" = company paid, not yet settled with sub-org | "settled" = already paid to sub-org
  const isSubOrgVarietyPaid = (variety) => {
    const s = varietySettings["so_"+variety];
    if (!s) return true;
    return s.status === "paid" || s.status === "settled";
  };
  const isSubOrgVarietySettled = (variety) => {
    const s = varietySettings["so_"+variety];
    return s?.status === "settled";
  };
  const isSubOrgVarietyToPay = (variety) => {
    const s = varietySettings["so_"+variety];
    if (!s) return true;
    return s.status === "paid"; // paid by company but not yet given to sub-org
  };
  const getSubOrgVarietyBillDate = (variety) => {
    const s = varietySettings["so_"+variety];
    return (s && s.billDate) ? s.billDate : BILL_DATE;
  };
  const getSubOrgVarietyRate = (variety) => {
    const s = varietySettings["so_"+variety] || varietySettings[variety];
    return (s && s.rate) ? parseFloat(s.rate) : null;
  };
  const getSubOrgVarietyType = (variety) => {
    const s = varietySettings["so_"+variety] || varietySettings[variety];
    if (s && s.type) return s.type;
    for (const so of (subOrgs||[])) {
      const g = (so.growers||[]).find(g=>g.variety===variety);
      if (g) return g.type||"KMS";
    }
    return "KMS";
  };
  // Helper: get variety bill date (falls back to global BILL_DATE)
  const getVarietyBillDate = (variety) => varietySettings[variety]?.billDate || BILL_DATE;
  // Helper: is variety paid
  const isVarietyPaid = (variety) => {
    const s = varietySettings[variety];
    if (!s) return true; // default: paid (so existing bills don't break)
    return s.status === "paid";
  };
  // Helper: get variety rate (returns null if not set — falls back to farmer's own rate)
  const getVarietyRate = (variety) => {
    const s = varietySettings[variety];
    if (!s || !s.rate) return null;
    return parseFloat(s.rate) || null;
  };
  // Helper: get variety type (from settings, falls back to farmer data)
  const getVarietyType = (variety) => {
    const s = varietySettings[variety];
    if (s && s.type) return s.type;
    // Fall back to first farmer with this variety
    for (const f of (farmers||[])) {
      const c = (f.crops||[]).find(c=>c.variety===variety);
      if (c) return c.cropType||"KMS";
    }
    return "KMS";
  };
  // Helper: calculate a farmer's net balance (shared logic used across tabs)
  const getFarmerBalance = (f) => {
    let cropVal=0, found=0, trans=0;
    (f.crops||[]).forEach(c=>{
      const qty=parseFloat(c.quantity)||0;
      const vRate=(c.rateOverride===true)?(parseFloat(c.ratePerUnit)||0):(getVarietyRate(c.variety)||parseFloat(c.ratePerUnit)||0);
      const paid=isVarietyPaid(c.variety);
      if(c.result==="Pass"&&paid) cropVal+=qty*vRate;
      found+=(parseFloat(c.area)||0)*1000;
      trans+=qty;
    });
    const paidDates=(f.crops||[]).filter(c=>c.result==="Pass"&&isVarietyPaid(c.variety)).map(c=>getVarietyBillDate(c.variety));
    const billDate=paidDates.length>0?paidDates.reduce((a,b)=>a<b?a:b):BILL_DATE;
    const advWI=(f.advances||[]).reduce((s,a)=>{const bd=a.tillDate||billDate;const {interest}=(a.compound?calcCompoundInterest:calcInterest)(a.amount,a.interestRate,a.date,bd);return s+a.amount+interest;},0);
    const jamWI=(f.jammaEnabled?f.jammaEntries||[]:[]).reduce((s,j)=>{const bd=j.tillDate||billDate;const {interest}=calcInterest(parseFloat(j.amount)||0,parseFloat(j.interestRate)||0,j.date||BILL_DATE,bd);return s+(parseFloat(j.amount)||0)+interest;},0);
    return cropVal-advWI+jamWI-found-trans;
  };
  // Helper: detect default rate from farmers (used as placeholder)
  const getVarietyInfo = (variety) => {
    for (const f of (farmers||[])) {
      for (const c of (f.crops||[])) {
        if (c.variety === variety) {
          return { type: c.cropType||"KMS", rate: parseFloat(c.ratePerUnit)||0 };
        }
      }
    }
    return { type: "KMS", rate: 550 };
  };
  const saveTimer = useRef(null);

  useEffect(() => {
    const loadData = async () => {
      setCloudStatus("saving");
      const cloud = await loadFromCloud();
      if (cloud && cloud.farmers && cloud.farmers.length > 0) {
        setFarmers(cloud.farmers);
        setSubOrgs(cloud.subOrgs || []);
        if (cloud.varietySettings && Object.keys(cloud.varietySettings).length > 0) {
          setVarietySettings(cloud.varietySettings);
          localStorage.setItem("variety_settings", JSON.stringify(cloud.varietySettings));
        }
        storage.saveFarmers(cloud.farmers);
        storage.saveSubOrgs(cloud.subOrgs || []);
      } else {
        const saved = storage.getFarmers();
        setFarmers(saved && saved.length > 0 ? saved : sampleFarmers);
        setSubOrgs(storage.getSubOrgs() || []);
      }
      setCloudStatus("idle");
    };
    loadData();
  }, []);

  const saveFarmers = async (data) => {
    const ok = storage.saveFarmers(data);
    setSaveStatus(ok ? "saved" : "error");
    setTimeout(() => setSaveStatus("idle"), 2000);
    // Also save to Firebase
    setCloudStatus("saving");
    const currentSubOrgs = subOrgs;
    const currentVS = JSON.parse(localStorage.getItem("variety_settings") || "{}");
    const saved = await saveToCloud(data, currentSubOrgs, currentVS);
    setCloudStatus(saved ? "saved" : "error");
    if (saved) setCloudLastSaved(new Date().toLocaleTimeString("en-IN", {hour:"2-digit", minute:"2-digit"}));
    setTimeout(() => setCloudStatus("idle"), 3000);
  };
  const updateFarmers = (data) => {
    setFarmers(data);
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => saveFarmers(data), 800);
  };
  const updateSubOrgs = async (data) => {
    setSubOrgs(data);
    storage.saveSubOrgs(data);
    // Also save to Firebase
    setCloudStatus("saving");
    const currentFarmers = farmers;
    const currentVS = JSON.parse(localStorage.getItem("variety_settings") || "{}");
    const saved = await saveToCloud(currentFarmers, data, currentVS);
    setCloudStatus(saved ? "saved" : "error");
    if (saved) setCloudLastSaved(new Date().toLocaleTimeString("en-IN", {hour:"2-digit", minute:"2-digit"}));
    setTimeout(() => setCloudStatus("idle"), 3000);
  };

  const addFarmer = (village = "") => {
    const newF = { id: Date.now(), farmerNo: "", name: "", fatherName: "", village, careOf: "", advances: [], crops: [], jammaEnabled: false, jammaEntries: [] };
    const updated = [...(farmers || []), newF];
    updateFarmers(updated); setSelectedIdx(updated.length - 1); setTab("form");
  };
  const addSubOrg = () => {
    const newS = { id: Date.now(), accNo: "", name: "", fatherName: "", village: "", advances: [], growers: [] };
    const updated = [...subOrgs, newS];
    updateSubOrgs(updated); setSelectedSubOrgIdx(updated.length - 1); setSubOrgTab("form");
  };



  




  const downloadSubOrgTemplate = () => {
    // GROWERS sheet — one row per grower, simple to fill
    const growerRows = [
      { "Acc No":"695","Sub-Org Name":"D K Ramudu","S-Org Father Name":"","Village":"MSP","S.No":1,"LOT No":"11069","Grower Name":"Mabbu Narasimhulu","Father Name":"Santenna","Grower Village":"Maddelabanda","Variety":"Royal-999","Packets":385,"Result":"Pass","Note":"" },
      { "Acc No":"695","Sub-Org Name":"D K Ramudu","S-Org Father Name":"","Village":"MSP","S.No":2,"LOT No":"514822","Grower Name":"Gopi Nayak","Father Name":"Lakshman Nayak","Grower Village":"Chinna Thanda","Variety":"Rasi-202","Packets":124,"Result":"Pass","Note":"" },
    ];
    // ADVANCES sheet — one row per advance
    const advRows = [
      { "Acc No":"695","Sub-Org Name":"D K Ramudu","Village":"MSP","Date":"2025-07-31","Amount":200000,"Interest %":24,"Note":"" },
      { "Acc No":"695","Sub-Org Name":"D K Ramudu","Village":"MSP","Date":"2025-08-19","Amount":80000,"Interest %":24,"Note":"" },
    ];
    // FOUNDATION sheet — one row per variety
    const foundRows = [
      { "Acc No":"695","Sub-Org Name":"D K Ramudu","Village":"MSP","Variety":"Royal-999","Area (Ac)":10 },
      { "Acc No":"695","Sub-Org Name":"D K Ramudu","Village":"MSP","Variety":"Rasi-202","Area (Ac)":8 },
    ];
    // SUB-ORG INFO sheet — for the bill comment (optional)
    const infoRows = [
      { "Acc No":"695","Sub-Org Name":"D K Ramudu","S-Org Father Name":"","Village":"MSP","Sub-Org Comment (optional)":"" },
    ];

    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.json_to_sheet(growerRows); ws1["!cols"]=Object.keys(growerRows[0]).map(k=>({wch:k.includes("Name")||k.includes("Village")||k.includes("Note")?20:12}));
    const ws2 = XLSX.utils.json_to_sheet(advRows); ws2["!cols"]=Object.keys(advRows[0]).map(()=>({wch:14}));
    const ws3 = XLSX.utils.json_to_sheet(foundRows); ws3["!cols"]=Object.keys(foundRows[0]).map(()=>({wch:14}));
    const ws4 = XLSX.utils.json_to_sheet(infoRows); ws4["!cols"]=[{wch:12},{wch:30}];
    XLSX.utils.book_append_sheet(wb, ws1, "Growers");
    XLSX.utils.book_append_sheet(wb, ws2, "Advances");
    XLSX.utils.book_append_sheet(wb, ws3, "Foundation");
    XLSX.utils.book_append_sheet(wb, ws4, "SubOrg Info");
    XLSX.writeFile(wb, "suborg_template.xlsx");
  };

  const handleSubOrgTemplateUpload = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "binary", cellDates: true });
        const clean = (v) => v==null ? "" : String(v).replace(/^['"]+|['"]+$/g,"").trim();
        const getNum = (v) => { if (v==null) return 0; if (typeof v==="number") return v; return parseFloat(String(v).replace(/[^0-9.-]/g,""))||0; };
        const toDate = (v) => {
          if (!v) return "";
          if (v instanceof Date) { const y=v.getFullYear(),m=String(v.getMonth()+1).padStart(2,"0"),d=String(v.getDate()).padStart(2,"0"); return y+"-"+m+"-"+d; }
          const s=clean(v); const m=s.match(/(\d{4})-(\d{2})-(\d{2})/); return m?m[0]:s;
        };

        const growersWs = wb.Sheets["Growers"];
        const advWs = wb.Sheets["Advances"];
        const foundWs = wb.Sheets["Foundation"];
        const infoWs = wb.Sheets["SubOrg Info"];
        if (!growersWs) { alert("⚠️ 'Growers' sheet not found. Use the downloaded template."); return; }

        const growerRows = XLSX.utils.sheet_to_json(growersWs, {raw:true});
        const advRowsRaw = advWs ? XLSX.utils.sheet_to_json(advWs, {raw:true}) : [];
        const foundRowsRaw = foundWs ? XLSX.utils.sheet_to_json(foundWs, {raw:true}) : [];
        const infoRowsRaw = infoWs ? XLSX.utils.sheet_to_json(infoWs, {raw:true}) : [];

        // Group by Acc No
        const subOrgMap = {};
        growerRows.forEach(r => {
          const accNo = clean(r["Acc No"]);
          if (!accNo) return;
          if (!subOrgMap[accNo]) {
            subOrgMap[accNo] = {
              id: Date.now()+Math.random(),
              accNo, name: clean(r["Sub-Org Name"]), fatherName: clean(r["S-Org Father Name"])||"", village: clean(r["Village"]),
              advances: [], growers: [], foundationSeeds: [], jammaEntries: [], comment: "",
            };
          }
          subOrgMap[accNo].growers.push({
            sNo: clean(r["S.No"]), lotNo: clean(r["LOT No"]), name: clean(r["Grower Name"]),
            fatherName: clean(r["Father Name"]), village: clean(r["Grower Village"]),
            variety: clean(r["Variety"]), packets: getNum(r["Packets"]),
            result: clean(r["Result"])==="Pass"?"Pass":"Fail", type: "KMS",
            rate: 0, note: clean(r["Note"]),
          });
        });

        advRowsRaw.forEach(r => {
          const accNo = clean(r["Acc No"]);
          if (!accNo || !subOrgMap[accNo]) return;
          const amt = getNum(r["Amount"]);
          if (amt<=0) return;
          subOrgMap[accNo].advances.push({
            date: toDate(r["Date"])||BILL_DATE, amount: amt,
            interestRate: getNum(r["Interest %"])||24, note: clean(r["Note"]),
          });
        });

        foundRowsRaw.forEach(r => {
          const accNo = clean(r["Acc No"]);
          if (!accNo || !subOrgMap[accNo]) return;
          const variety = clean(r["Variety"]);
          const area = getNum(r["Area (Ac)"]);
          if (!variety || area<=0) return;
          subOrgMap[accNo].foundationSeeds.push({ variety, area: String(area) });
        });

        infoRowsRaw.forEach(r => {
          const accNo = clean(r["Acc No"]);
          if (!accNo || !subOrgMap[accNo]) return;
          const comment = clean(r["Sub-Org Comment (optional)"]);
          if (comment) subOrgMap[accNo].comment = comment;
          const soFather = clean(r["S-Org Father Name"]);
          if (soFather) subOrgMap[accNo].fatherName = soFather;
        });

        const imported = Object.values(subOrgMap);
        if (imported.length === 0) { alert("⚠️ No sub-orgs found. Check the Acc No column is filled."); return; }

        // Merge: replace existing sub-orgs with same accNo, keep others
        const existing = (subOrgs||[]).filter(so => !imported.some(n=>n.accNo===so.accNo));
        updateSubOrgs([...existing, ...imported]);
        setSelectedSubOrgIdx(0);
        alert(`✅ Imported ${imported.length} sub-org(s)!`);
      } catch(err) { alert("⚠️ Error: "+err.message); }
    };
    reader.readAsBinaryString(file);
  };

  const downloadTemplate = () => {
    const data = [
      { "Farmer No":"001","Farmer Name":"Ravi Kumar","Father Name":"Suresh Kumar","Village":"Nandyal","C/o (optional)":"",
        "Adv1 Date":"2025-07-15","Adv1 Amount":50000,"Adv1 Interest%":24,"Adv1 Note":"",
        "Adv2 Date":"","Adv2 Amount":"","Adv2 Interest%":"","Adv2 Note":"",
        "Crop1 Variety":"NC-4605","Crop1 LOT No":"5579","Crop1 Area":"1","Crop1 Qty":410,"Crop1 Type":"GMS","Crop1 Rate":400,"Crop1 Result":"Pass","Crop1 Note":"",
        "Crop2 Variety":"","Crop2 LOT No":"","Crop2 Area":"","Crop2 Qty":"","Crop2 Type":"","Crop2 Rate":"","Crop2 Result":"","Crop2 Note":"",
        "Jamma Enabled":"No",
        "Jamma1 Date":"","Jamma1 Amount":"","Jamma1 Interest%":"","Jamma1 Note":"",
        "Jamma2 Date":"","Jamma2 Amount":"","Jamma2 Interest%":"","Jamma2 Note":"",
        "Farmer Comment":"" },
    ];
    const ws = XLSX.utils.json_to_sheet(data); ws["!cols"] = Object.keys(data[0]).map(k => ({ wch: k.includes("Name")||k.includes("Village")||k.includes("Comment")||k.includes("Note") ? 20 : 13 }));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Template"); XLSX.writeFile(wb, "farmer_template.xlsx");
  };

  const exportAllData = () => {
    const maxAdv  = Math.max(1, ...(farmers||[]).map(f=>(f.advances||[]).length));
    const maxCrop = Math.max(1, ...(farmers||[]).map(f=>(f.crops||[]).length));
    const maxJam  = Math.max(0, ...(farmers||[]).map(f=>(f.jammaEntries||[]).length));

    // Helper: calculate bill date for a farmer
    const getFarmerBillDate = (f) => {
      const paidDates = (f.crops||[]).filter(c=>c.result==="Pass"&&isVarietyPaid(c.variety)).map(c=>getVarietyBillDate(c.variety));
      return paidDates.length > 0 ? paidDates.reduce((a,b)=>a<b?a:b) : BILL_DATE;
    };

    // Build headers
    const headers = ["Farmer No","Farmer Name","Father Name","Village","C/o","Bill Date"];
    for (let i=1;i<=maxAdv;i++)  headers.push(`Adv${i} Date`,`Adv${i} Amount`,`Adv${i} Days`,`Adv${i} Interest`,`Adv${i} Total`,`Adv${i} Note`);
    for (let i=1;i<=maxCrop;i++) headers.push(`Crop${i} Variety`,`Crop${i} LOT No`,`Crop${i} Qty`,`Crop${i} Type`,`Crop${i} Rate`,`Crop${i} Result`,`Crop${i} Value`,`Crop${i} Note`);
    if (maxJam>0) {
      for (let i=1;i<=maxJam;i++) headers.push(`Jamma${i} Date`,`Jamma${i} Amount`,`Jamma${i} Days`,`Jamma${i} Interest`,`Jamma${i} Total`,`Jamma${i} Note`);
    }
    headers.push("Total Advance","Total Interest","Total Adv+Int","Total Crop Value","Foundation","Transportation","Balance","Farmer Comment");

    const rows = (farmers||[]).map(f => {
      const billDate = getFarmerBillDate(f);
      const row = {};
      row["Farmer No"]   = f.farmerNo||"";
      row["Farmer Name"] = f.name||"";
      row["Father Name"] = f.fatherName||"";
      row["Village"]     = f.village||"";
      row["C/o"]         = f.careOf||"";
      row["Bill Date"]   = fmtDate(billDate);

      // Calculate advances with interest
      let totalAdv=0, totalInt=0;
      for (let i=0;i<maxAdv;i++) {
        const a = (f.advances||[])[i];
        if (a) {
          const advBillDate = a.tillDate || billDate;
          const {days, interest} = (a.compound?calcCompoundInterest:calcInterest)(parseFloat(a.amount)||0, parseFloat(a.interestRate)||0, a.date, advBillDate);
          const total = (parseFloat(a.amount)||0) + interest;
          totalAdv += parseFloat(a.amount)||0;
          totalInt += interest;
          row[`Adv${i+1} Date`]     = fmtDate(a.date);
          row[`Adv${i+1} Amount`]   = parseFloat(a.amount)||0;
          row[`Adv${i+1} Days`]     = days;
          row[`Adv${i+1} Interest`] = interest;
          row[`Adv${i+1} Total`]    = total;
          row[`Adv${i+1} Note`]     = a.note||"";
        } else {
          row[`Adv${i+1} Date`]=""; row[`Adv${i+1} Amount`]=""; row[`Adv${i+1} Days`]="";
          row[`Adv${i+1} Interest`]=""; row[`Adv${i+1} Total`]=""; row[`Adv${i+1} Note`]="";
        }
      }

      // Calculate crops
      let totalCrop=0, totalFound=0, totalTrans=0;
      for (let i=0;i<maxCrop;i++) {
        const c = (f.crops||[])[i];
        if (c) {
          const vRate = (c.rateOverride===true) ? (parseFloat(c.ratePerUnit)||0) : (getVarietyRate(c.variety)||parseFloat(c.ratePerUnit)||0);
          const qty = parseFloat(c.quantity)||0;
          const area = parseFloat(c.area)||0;
          const paid = isVarietyPaid(c.variety);
          const cropVal = (c.result==="Pass"&&paid) ? qty*vRate : 0;
          totalCrop += cropVal;
          totalFound += area*1000;
          totalTrans += qty;
          row[`Crop${i+1} Variety`] = c.variety||"";
          row[`Crop${i+1} LOT No`]  = c.lotNo||"";
          row[`Crop${i+1} Qty`]     = qty;
          row[`Crop${i+1} Type`]    = getVarietyType(c.variety)||c.cropType||"KMS";
          row[`Crop${i+1} Rate`]    = vRate;
          row[`Crop${i+1} Result`]  = c.result||"";
          row[`Crop${i+1} Value`]   = cropVal;
          row[`Crop${i+1} Note`]    = c.note||"";
        } else {
          row[`Crop${i+1} Variety`]=""; row[`Crop${i+1} LOT No`]=""; row[`Crop${i+1} Qty`]="";
          row[`Crop${i+1} Type`]=""; row[`Crop${i+1} Rate`]=""; row[`Crop${i+1} Result`]=""; row[`Crop${i+1} Value`]=""; row[`Crop${i+1} Note`]="";
        }
      }

      // Jamma
      if (maxJam>0) {
        for (let i=0;i<maxJam;i++) {
          const j = (f.jammaEntries||[])[i];
          if (j) {
            const {days,interest} = calcInterest(parseFloat(j.amount)||0, parseFloat(j.interestRate)||0, j.date||BILL_DATE, billDate);
            row[`Jamma${i+1} Date`]     = fmtDate(j.date||BILL_DATE);
            row[`Jamma${i+1} Amount`]   = parseFloat(j.amount)||0;
            row[`Jamma${i+1} Days`]     = days;
            row[`Jamma${i+1} Interest`] = interest;
            row[`Jamma${i+1} Total`]    = (parseFloat(j.amount)||0)+interest;
            row[`Jamma${i+1} Note`]     = j.note||"";
          } else {
            row[`Jamma${i+1} Date`]=""; row[`Jamma${i+1} Amount`]=""; row[`Jamma${i+1} Days`]="";
            row[`Jamma${i+1} Interest`]=""; row[`Jamma${i+1} Total`]=""; row[`Jamma${i+1} Note`]="";
          }
        }
      }
      // Summary at end
      const advWI = totalAdv + totalInt;
      const balance = totalCrop - advWI - totalFound - totalTrans;
      row["Total Advance"]    = totalAdv;
      row["Total Interest"]   = totalInt;
      row["Total Adv+Int"]    = advWI;
      row["Total Crop Value"] = totalCrop;
      row["Foundation"]       = totalFound;
      row["Transportation"]   = totalTrans;
      row["Balance"]          = balance;
      row["Farmer Comment"]   = f.comment||"";

      return row;
    });

    const ws = XLSX.utils.json_to_sheet(rows, {header: headers});
    ws["!cols"] = headers.map(h => ({
      wch: h.includes("Name")||h.includes("Village")||h.includes("Note")||h.includes("Variety") ? 20
           : h.includes("Date") ? 12
           : h.includes("Balance")||h.includes("Total")||h.includes("Value") ? 14
           : 10
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "All Farmers");
    XLSX.writeFile(wb, `farmer_backup_${new Date().toISOString().split("T")[0]}.xlsx`);
  };

  if (farmers === null) return (
    <div style={{ display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:"#f0f7f0",flexDirection:"column",gap:16 }}>
      <div style={{ fontSize:40 }}>🌾</div>
      <div style={{ fontSize:18,fontWeight:700,color:"#1a4a1a" }}>Loading your data...</div>
      <div style={{ fontSize:13,color:"#666" }}>Fetching from cloud storage</div>
    </div>
  );

  const currentFarmer = farmers[selectedIdx] || farmers[0];
  const currentSubOrg = subOrgs[selectedSubOrgIdx] || subOrgs[0];
  const btnStyle = { background:"rgba(255,255,255,0.15)",color:"#fff",border:"1px solid rgba(255,255,255,0.4)",borderRadius:5,padding:"6px 12px",cursor:"pointer",fontSize:12 };
  const inp2 = { style:{ width:"100%",padding:"5px 8px",border:"1px solid #b0c8e0",borderRadius:4,fontSize:13,background:"#fafdff",boxSizing:"border-box" } };


  const exportSubOrgData = () => {
    if (!subOrgs || subOrgs.length === 0) { alert("No sub-organizer data to export."); return; }
    const wb = XLSX.utils.book_new();

    subOrgs.forEach(so => {
      const sheetName = (so.name||"SubOrg").substring(0,31).replace(/[/*?:[\]]/g,"_");
      const billDate = so._billDate || BILL_DATE;
      const rows = [];

      // Header
      rows.push([`SUB-ORGANIZER BILL`, ``, ``, ``, ``, ``, ``, `Acc No: ${so.accNo||""}`, `Bill Date: ${fmtDate(billDate)}`]);
      rows.push([`Sub-Organizer: ${so.name||""}`, ``, `Village: ${so.village||""}`]);
      rows.push([]);

      // Advances
      rows.push([`ADVANCES`]);
      rows.push([`S.No`,`Date`,`Amount (Rs)`,`Days`,`Interest (Rs)`,`Note`,`Total (Rs)`]);
      let ta=0,ti=0;
      (so.advances||[]).forEach((adv,i)=>{
        const amt=parseFloat(adv.amount)||0;
        const d=Math.max(1,Math.round((parseLocalDate(billDate)-parseLocalDate(adv.date))/86400000));
        const interest=Math.round(amt*(parseFloat(adv.interestRate)||0)*d/(100*30*12));
        ta+=amt; ti+=interest;
        rows.push([i+1, fmtDate(adv.date), amt, d, interest, adv.note||"—", amt+interest]);
      });
      rows.push([`TOTAL`, ``, ta, ``, ti, ``, ta+ti]);
      rows.push([]);

      // Jamma
      const jamma=so.jammaEntries||[];
      let tj=0,tji=0;
      if(jamma.length>0){
        rows.push([`JAMMA DETAILS (Partial Payments Received from Sub-Org)`]);
        rows.push([`S.No`,`Date`,`Amount (Rs)`,`Days`,`Interest (Rs)`,`Note`,`Total (Rs)`]);
        jamma.forEach((j,i)=>{
          const amt=parseFloat(j.amount)||0;
          const d=Math.max(0,Math.round((parseLocalDate(billDate)-parseLocalDate(j.date||billDate))/86400000));
          const interest=Math.round(amt*(parseFloat(j.interestRate)||0)*d/(100*30*12));
          tj+=amt; tji+=interest;
          rows.push([i+1,fmtDate(j.date||billDate),amt,d,interest,j.note||"—",amt+interest]);
        });
        rows.push([`TOTAL`,``,tj,``,tji,``,tj+tji]);
        rows.push([]);
      }

      // Foundation
      const seeds=so.foundationSeeds||[];
      let tf=0;
      if(seeds.length>0){
        rows.push([`FOUNDATION`]);
        rows.push([`S.No`,`Variety`,`Area (Ac)`,`Quantity (Pkts)`]);
        seeds.forEach((f,i)=>{
          const area=parseFloat(f.area)||0;
          const qty=(so.growers||[]).filter(g=>g.variety===f.variety&&g.result==="Pass").reduce((s,g)=>s+(parseFloat(g.packets)||0),0);
          rows.push([i+1,f.variety||"",area,qty]);
        });
        rows.push([`TOTAL`,``,seeds.reduce((s,f)=>s+(parseFloat(f.area)||0),0),seeds.reduce((s,f)=>{
          const qty=(so.growers||[]).filter(g=>g.variety===f.variety&&g.result==="Pass").reduce((ss,g)=>ss+(parseFloat(g.packets)||0),0);
          return s+qty;
        },0)]);
        rows.push([]);
      }

      // Growers
      const growers=(so.growers||[]).slice().sort((a,b)=>(parseFloat(a.sNo)||0)-(parseFloat(b.sNo)||0));
      const gRate=(v,gr)=>parseFloat((varietySettings["so_"+v]||varietySettings[v]||{}).rate||gr||0);
      const gType=(v,gt)=>(varietySettings["so_"+v]||varietySettings[v]||{}).type||gt||"KMS";
      const gStatus=(v)=>(varietySettings["so_"+v]||varietySettings[v]||{}).status||"paid";
      const settled_g=growers.filter(g=>gStatus(g.variety)==="settled");
      const topay_g=growers.filter(g=>gStatus(g.variety)==="paid");
      const pend_g=growers.filter(g=>gStatus(g.variety)==="pending");
      const gCols=[`S.No`,`LOT No`,`Grower Name`,`Father Name`,`Village`,`Variety`,`Packets`,`Result`,`Type`,`Rate (Rs)`,`Amount (Rs)`,`Note`];
      let sno=1,ts=0,tt=0,tp=0;

      if(settled_g.length>0){
        rows.push([`SETTLED — Already paid to Sub-Org — ${settled_g.filter(g=>g.result==="Pass").length} growers`]);
        rows.push(gCols);
        settled_g.forEach(g=>{
          const pkts=parseFloat(g.packets)||0,rt=gRate(g.variety,g.rate),gt=gType(g.variety,g.type);
          const amt=g.result==="Pass"?pkts*rt:0;
          if(g.result==="Pass"){ts+=amt;tt+=pkts;}
          rows.push([sno++,String(g.lotNo||""),g.name||"",g.fatherName||"—",g.village||"",g.variety||"",pkts,g.result||"",gt,g.result==="Pass"?rt:"—",g.result==="Pass"?amt:"—",g.note||""]);
        });
      }
      if(topay_g.length>0){
        rows.push([`TO PAY — Company paid us, need to pay Sub-Org — ${topay_g.filter(g=>g.result==="Pass").length} growers`]);
        rows.push(gCols);
        topay_g.forEach(g=>{
          const pkts=parseFloat(g.packets)||0,rt=gRate(g.variety,g.rate),gt=gType(g.variety,g.type);
          const amt=g.result==="Pass"?pkts*rt:0;
          if(g.result==="Pass"){ts+=amt;tt+=pkts;}
          rows.push([sno++,String(g.lotNo||""),g.name||"",g.fatherName||"—",g.village||"",g.variety||"",pkts,g.result||"",gt,g.result==="Pass"?rt:"—",g.result==="Pass"?amt:"—",g.note||""]);
        });
      }
      if(pend_g.length>0){
        rows.push([`PENDING VARIETIES — ${pend_g.filter(g=>g.result==="Pass").length} growers — Payment not received yet`]);
        rows.push(gCols);
        pend_g.forEach(g=>{
          const pkts=parseFloat(g.packets)||0,rt=gRate(g.variety,g.rate),gt=gType(g.variety,g.type);
          if(g.result==="Pass") tp+=pkts*rt;
          rows.push([sno++,String(g.lotNo||""),g.name||"",g.fatherName||"—",g.village||"",g.variety||"",pkts,g.result||"",gt,g.result==="Pass"?rt:"—",g.result==="Pass"?"Pending":"—",g.note||""]);
        });
      }
      const apAll=growers.filter(g=>g.result==="Pass");
      rows.push([`TOTAL`,``,``,``,``,``,growers.reduce((s,g)=>s+(parseFloat(g.packets)||0),0),`${apAll.length}P/${growers.length-apAll.length}F`,``,``,ts]);
      rows.push([]);

      // Settlement
      rows.push([]);
      rows.push([`SETTLEMENT SUMMARY`, ``]);
      const awI=ta+ti, jwI=tj+tji, bal=ts-awI+jwI-tf-tt;
      rows.push([`Seed Amount (Paid)`, ts]);
      if(tp>0) rows.push([`Pending (not counted yet)`, tp]);
      rows.push([`Advance + Interest`, -awI]);
      if(jwI>0) rows.push([`Jamma + Interest (partial payments)`, jwI]);
      if(tf>0) rows.push([`Foundation Cost`, -tf]);
      if(tt>0) rows.push([`Transportation`, -tt]);
      rows.push([bal>=0?`PAYABLE TO SUB-ORG`:`DUE FROM SUB-ORG`, Math.abs(bal)]);

      const ws = XLSX.utils.aoa_to_sheet(rows);

      // Auto-fit column widths based on actual content
      const maxCols = rows.reduce((m, r) => Math.max(m, r.length), 0);
      const colWidths = Array(maxCols).fill(8);
      rows.forEach(row => {
        row.forEach((cell, ci) => {
          if (cell !== null && cell !== undefined && cell !== "") {
            const len = String(cell).length;
            if (len > colWidths[ci]) colWidths[ci] = Math.min(len + 2, 40);
          }
        });
      });
      ws["!cols"] = colWidths.map(w => ({wch: w}));
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });

    XLSX.writeFile(wb, "suborg_bill_"+new Date().toISOString().split("T")[0]+".xlsx");
  };







  const handleSubOrgExcelUpload = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "binary", cellDates: true });

        // Strip surrounding quotes and clean whitespace
        const clean = (v) => {
          if (v == null) return "";
          return String(v).replace(/^['"]+|['"]+$/g,"").replace(/\n/g," ").trim();
        };
        const getNum = (v) => {
          if (v == null) return 0;
          if (typeof v === "number") return v;
          return parseFloat(String(v).replace(/[^0-9.-]/g,"")) || 0;
        };
        const toDate = (v) => {
          if (!v) return "";
          if (v instanceof Date) {
            const y=v.getFullYear(), m=String(v.getMonth()+1).padStart(2,"0"), d=String(v.getDate()).padStart(2,"0");
            return y+"-"+m+"-"+d;
          }
          // Handle datetime string from openpyxl format
          const s = String(v);
          const match = s.match(/(\d{4})-(\d{2})-(\d{2})/);
          if (match) return match[0];
          return "";
        };

        // ── SHEET 2: Acc No, Name, Advances, Foundation Seeds ──
        const ws2 = wb.Sheets[wb.SheetNames[1]];
        const s2 = XLSX.utils.sheet_to_json(ws2, { header:1, raw:true });

        // Row 1 [0]: col 0="Acc No:", col 1=695
        const accNo = String(getNum(s2[0]?.[1]) || clean(s2[0]?.[1])).replace(/['"]/g,"").trim();
        // Row 2 [1]: col 0="Sub-Org:", col 1="D K Ramudu"
        const soName = clean(s2[1]?.[1]).replace(/['"]/g,"").trim();

        // Rows 4+ [3+]: advances - col 1=Date, col 2=Amount
        // Foundation seeds start after blank rows - col 1=Variety, col 2=Area
        const advances = [];
        const foundationSeeds = [];
        let foundFoundation = false;

        for (let i = 3; i < s2.length; i++) {
          const row = s2[i]; if (!row) continue;
          const col1 = row[1]; const col2 = row[2];

          // Detect foundation section header - handle quotes
          const col1str = clean(col1).replace(/['"]/g,"").trim();
          if (col1str === "Foundation Cost" || col1str === "Varitey:" || col1str === "Varitey") {
            foundFoundation = true; continue;
          }

          if (foundFoundation) {
            // Foundation seed row: col1=Variety, col2=Area
            const variety = clean(col1).replace(/['"]/g,"").trim();
            const area = getNum(col2);
            if (variety && variety !== "Area" && area > 0) {
              foundationSeeds.push({ variety, area: String(area) });
            }
          } else {
            // Advance row: col1=Date (Date object), col2=Amount
            if (col1 instanceof Date && getNum(col2) > 0) {
              advances.push({
                date: toDate(col1),
                amount: getNum(col2),
                interestRate: 24,
                note: clean(row[4]) || "",
              });
            }
          }
        }

        // ── SHEET 1: Sub-org name, Growers ───────────────────
        // Columns: [0]=S.No, [1]=LOT No, [2]=Name, [3]=Father, [4]=Village,
        //          [5]=Variety, [6]=Packets, [7]=Result, [8]=Note
        const ws1 = wb.Sheets[wb.SheetNames[0]];
        const s1 = XLSX.utils.sheet_to_json(ws1, { header:1, raw:true });

        // Row 4 [3]: col 2 = "D K Ramudu(MSP)" — extract village from brackets
        const soNameFull = clean(s1[3]?.[2]) || soName;
        const villageMatch = soNameFull.match(/\(([^)]+)\)/);
        const soVillage = villageMatch ? villageMatch[1].trim() : "";

        // Row 5 [4]: headers — skip
        // Rows 6+ [5+]: grower data
        const growers = [];
        for (let i = 5; i < s1.length; i++) {
          const row = s1[i]; if (!row) continue;

          const sNo = getNum(row[0]);
          const lotNo = clean(row[1]);
          const name = clean(row[2]);
          const father = clean(row[3]);
          const village = clean(row[4]);
          const variety = clean(row[5]);
          const packets = getNum(row[6]);
          const result = clean(row[7]) === "Pass" ? "Pass" : "Fail";

          // Must have a name or LOT No to be a real grower
          if (!name && !lotNo) continue;
          // Stop at summary rows (variety totals at bottom — no name, no lotNo, has packets)
          if (!name && !lotNo && packets > 0) continue;
          // Skip rows that look like totals/summaries
          if (name.includes("Total") || name.includes("GMS") || name.includes("CONV")) continue;
          // Skip if no variety (summary rows at bottom)
          if (!variety && !name) continue;

          growers.push({
            sNo: String(sNo || growers.length + 1),
            lotNo,
            name,
            fatherName: father,
            village,
            variety: variety || "",
            area: "",
            packets,
            result,
            type: "KMS",  // default — can be changed in Variety Pay settings
            rate: 550,    // default — can be changed in Variety Pay settings
          });
        }

        // Build sub-org object
        const newSO = {
          id: Date.now(),
          accNo,
          name: soName,
          village: soVillage,
          advances,
          growers,
          foundationSeeds,
        };

        // If sub-org with same accNo already exists, replace it
        const existingIdx = subOrgs.findIndex(s => s.accNo === newSO.accNo);
        let updated;
        if (existingIdx >= 0) {
          updated = [...subOrgs];
          updated[existingIdx] = newSO;
        } else {
          updated = [...subOrgs, newSO];
        }
        updateSubOrgs(updated);
        setSelectedSubOrgIdx(existingIdx >= 0 ? existingIdx : updated.length - 1);
        setSubOrgTab("growers");
        alert(
          "✅ Imported: " + soName + " (" + soVillage + ")\n" +
          "Acc No: " + accNo + "\n" +
          "Growers: " + growers.length + "\n" +
          "Advances: " + advances.length + "\n" +
          "Foundation Seeds: " + foundationSeeds.length
        );
      } catch(err) {
        alert("⚠️ Error: " + err.message);
      }
    };
    reader.readAsBinaryString(file);
  };

  // ── Smart C/o Update: match by Farmer No, update only C/o field ──
  const handleCareOfUpdate = (e) => {
    const file = e.target.files[0]; if (!file) return;
    e.target.value = "";
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "binary", cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
        const clean = (v) => { if (v==null) return ""; return String(v).replace(/^['\"]+|['\"]+$/g,"").trim(); };

        // Find header row
        let headerIdx = 0;
        for (let i=0;i<Math.min(5,allRows.length);i++) {
          if (allRows[i]?.some(c => { const s=clean(c).toLowerCase(); return s.includes("farmer name")||s==="farmer no"; })) { headerIdx=i; break; }
        }
        const headers = (allRows[headerIdx]||[]).map(h=>clean(h));
        const hIdx = {}; headers.forEach((h,i)=>{ hIdx[h]=i; });
        const get = (row, name) => { const idx=hIdx[name]??-1; if (idx<0||row[idx]==null) return ""; return clean(row[idx]); };

        // Build map of farmerNo → careOf from Excel
        const careOfMap = {};
        for (let ri=headerIdx+1; ri<allRows.length; ri++) {
          const row = allRows[ri]; if (!row||row.length===0) continue;
          const farmerNo = get(row,"Farmer No"); if (!farmerNo) continue;
          const careOf = get(row,"C/O:")||get(row,"C/o (optional)")||get(row,"C/o")||get(row,"C/O")||"";
          if (careOf) careOfMap[farmerNo] = careOf;
        }

        const total = Object.keys(careOfMap).length;
        if (total === 0) { alert("⚠️ No C/o values found in the file. Make sure the file has a 'C/O:' column with values."); return; }

        // Update only farmers whose Farmer No matches AND C/o is filled in Excel
        let updated = 0;
        const newFarmers = (farmers||[]).map(f => {
          const co = careOfMap[f.farmerNo];
          if (co !== undefined) { updated++; return {...f, careOf: co}; }
          return f;
        });

        updateFarmers(newFarmers);
        alert(`✅ Updated C/o for ${updated} farmers out of ${total} found in Excel.\n\nNo farmers were added or removed — only C/o field was updated.`);
      } catch(err) {
        alert("⚠️ Error reading file: " + err.message);
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleExcelUpload = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "binary", cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });

        const clean = (v) => { if (v==null) return ""; return String(v).replace(/^['"]+|['"]+$/g,"").trim(); };
        const getNum = (v) => { if (v==null) return 0; if (typeof v==="number") return v; return parseFloat(String(v).replace(/[^0-9.-]/g,""))||0; };
        const toDate = (v) => {
          if (!v) return "";
          if (v instanceof Date) { const y=v.getFullYear(),m=String(v.getMonth()+1).padStart(2,"0"),d=String(v.getDate()).padStart(2,"0"); return y+"-"+m+"-"+d; }
          const s=clean(v); if (s.match(/^\d{4}-\d{2}-\d{2}/)) return s.substring(0,10); return s;
        };

        // Find header row
        let headerIdx = 0;
        for (let i=0;i<Math.min(5,allRows.length);i++) {
          if (allRows[i] && allRows[i].some(c => { const s=clean(c).toLowerCase(); return s.includes("farmer name")||s==="farmer no"; })) { headerIdx=i; break; }
        }
        const headers = (allRows[headerIdx]||[]).map(h => clean(h));
        const hIdx = {}; headers.forEach((h,i) => { hIdx[h]=i; });
        const get = (row, name) => { const idx=hIdx[name]??-1; if (idx<0||row[idx]==null) return ""; return clean(row[idx]); };

        // Detect format: new (single-row) or old (multi-row)
        const isNewFormat = headers.some(h => h.match(/^Adv\d+ Date$/));

        const imported = [];

        if (isNewFormat) {
          // ── NEW FORMAT: one row per farmer, all advances/crops in columns ──
          for (let ri=headerIdx+1;ri<allRows.length;ri++) {
            const row=allRows[ri]; if (!row||row.length===0) continue;
            const farmerName=get(row,"Farmer Name"); if (!farmerName) continue;

            const advances=[], crops=[], jammaEntries=[];
            // Read up to 10 advances
            for (let i=1;i<=10;i++) {
              const date=toDate(row[hIdx[`Adv${i} Date`]??-1]);
              const amount=getNum(row[hIdx[`Adv${i} Amount`]??-1]);
              if (!date && !amount) break;
              advances.push({ date:date||BILL_DATE, amount, interestRate:getNum(row[hIdx[`Adv${i} Interest%`]??-1])||24, note:get(row,`Adv${i} Note`) });
            }
            // Read up to 3 crops
            for (let i=1;i<=3;i++) {
              const variety=get(row,`Crop${i} Variety`); if (!variety) break;
              crops.push({ variety, lotNo:get(row,`Crop${i} LOT No`), area:get(row,`Crop${i} Area`)||"0", quantity:getNum(row[hIdx[`Crop${i} Qty`]??-1]), cropType:get(row,`Crop${i} Type`)||"KMS", ratePerUnit:getNum(row[hIdx[`Crop${i} Rate`]??-1])||550, result:get(row,`Crop${i} Result`)==="Pass"?"Pass":"Fail", note:get(row,`Crop${i} Note`) });
            }
            // Read up to 5 jamma
            const jammaEnabled = get(row,"Jamma Enabled")==="Yes";
            for (let i=1;i<=5;i++) {
              const date=toDate(row[hIdx[`Jamma${i} Date`]??-1]);
              const amount=getNum(row[hIdx[`Jamma${i} Amount`]??-1]);
              if (!date && !amount) break;
              jammaEntries.push({ date:date||BILL_DATE, amount, interestRate:getNum(row[hIdx[`Jamma${i} Interest%`]??-1])||getNum(row[hIdx[`Jamma${i} Interest`]??-1])||0, note:get(row,`Jamma${i} Note`) });
            }            imported.push({ id:Date.now()+ri, farmerNo:get(row,"Farmer No"), name:farmerName, fatherName:get(row,"Father Name"), village:get(row,"Village"), careOf:get(row,"C/o (optional)")||get(row,"C/o")||get(row,"C/O:")||get(row,"C/O")||"", advances, crops, jammaEnabled, jammaEntries, comment:get(row,"Farmer Comment") });
          }
        } else {
          // ── OLD FORMAT: multi-row per farmer ──
          let cur = null;
          for (let ri=headerIdx+1;ri<allRows.length;ri++) {
            const row=allRows[ri]; if (!row||row.length===0) continue;
            const farmerName=get(row,"Farmer Name");
            if (farmerName) {
              cur={ id:Date.now()+ri, farmerNo:get(row,"Farmer No"), name:farmerName, fatherName:get(row,"Father Name"), village:get(row,"Village"), careOf:get(row,"C/o")||"", advances:[], crops:[], jammaEnabled:get(row,"Jamma Enabled")==="Yes", jammaEntries:[] };
              imported.push(cur);
            }
            if (!cur) continue;
            const advDate=toDate(row[hIdx["Advance Date"]??-1]);
            const advAmount=getNum(row[hIdx["Advance Amount"]??-1]);
            if (advDate||advAmount>0) cur.advances.push({ date:advDate||BILL_DATE, amount:advAmount, interestRate:getNum(row[hIdx["Interest Rate %"]??-1])||24, note:get(row,"Note") });
            const variety=get(row,"Crop Variety");
            if (variety) cur.crops.push({ variety, lotNo:get(row,"LOT No:")||"", area:get(row,"Crop Area")||"0", quantity:getNum(row[hIdx["Quantity"]??-1]), cropType:get(row,"Type")||"KMS", ratePerUnit:getNum(row[hIdx["Rate Per Unit"]??-1])||550, result:get(row,"Result")==="Pass"?"Pass":"Fail" });
            const jamDate=toDate(row[hIdx["Jamma Date"]??-1]);
            const jamAmount=getNum(row[hIdx["Jamma Amount"]??-1]);
            if (jamDate||jamAmount>0) { cur.jammaEnabled=true; cur.jammaEntries.push({ date:jamDate||BILL_DATE, amount:jamAmount, interestRate:getNum(row[hIdx["Jamma Interest %"]??-1])||0, note:get(row,"Jamma Note") }); }
          }
        }

        if (imported.length>0) { updateFarmers(imported); setSelectedIdx(0); setSelectedVillage(null); alert("✅ Imported "+imported.length+" farmers!"); }
        else alert("⚠️ No farmers found. Check your file has a Farmer Name column.");
      } catch(err) { alert("⚠️ Error: "+err.message); }
    };
    reader.readAsBinaryString(file);
  };

  return (
    <div style={{ fontFamily:"Noto Serif,Georgia,serif", minHeight:"100vh", background:"#f0f7f0" }}>
      <link href="https://fonts.googleapis.com/css2?family=Noto+Serif+Telugu&family=Noto+Serif:wght@400;600;700&display=swap" rel="stylesheet" />

      <style>{`
        @media (max-width: 600px) {
          .mobile-toolbar { flex-direction: column !important; align-items: stretch !important; gap: 10px !important; }
          .mobile-toolbar-btns { display: flex !important; flex-wrap: wrap !important; gap: 8px !important; }
          .mobile-toolbar-btns button, .mobile-toolbar-btns label {
            flex: 1 1 calc(50% - 8px) !important;
            text-align: center !important;
            font-size: 13px !important;
            padding: 10px 6px !important;
          }
          .mobile-nav { width: 100% !important; justify-content: space-between !important; }
          .mobile-nav button { flex: 1 !important; font-size: 11px !important; padding: 7px 4px !important; text-align: center !important; }
          .mobile-grid-4 { grid-template-columns: 1fr 1fr !important; }
          .mobile-adv-grid { grid-template-columns: 1fr 1fr !important; gap: 8px !important; }
          .mobile-crop-grid { grid-template-columns: 1fr 1fr !important; gap: 8px !important; }
          .mobile-panel { width: 100% !important; min-width: unset !important; flex-direction: row !important; flex-wrap: wrap !important; }
          .farmer-form-pad { padding: 12px !important; }
          input, select { font-size: 16px !important; min-height: 42px !important; }
          input[type=checkbox] { min-height: unset !important; width: 18px !important; height: 18px !important; }
          button { min-height: 38px !important; }
        }
      `}</style>
      {/* Top Bar */}
      <div style={{ background:"linear-gradient(135deg,#1a4a1a,#2d6a2d)",color:"#fff",padding:"10px 16px",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8 }} className="mobile-toolbar">
        <div style={{ display:"flex",alignItems:"center",gap:14 }}>
          <div>
            <div style={{ fontSize:18,fontWeight:700 }}>🌾 Farmer Bill Generator</div>
            <div style={{ fontSize:11,opacity:0.8 }}>రైతు పంట బిల్లు జనరేటర్</div>
          </div>
          <div style={{ display:"flex",background:"rgba(0,0,0,0.25)",borderRadius:8,padding:3,gap:2,flexWrap:"wrap" }} className="mobile-nav">
            {[["farmers","👨‍🌾 Farmers"],["suborgs","🏢 Sub-Orgs"],["careof","🤝 C/o Groups"],["dashboard","📊 Dashboard"],["variety","🌾 Variety Pay"]].map(([m,l]) => (
              <button key={m} onClick={()=>setMode(m)} style={{ padding:"5px 12px",borderRadius:6,border:"none",cursor:"pointer",fontWeight:700,fontSize:12,background:mode===m?"#fff":"transparent",color:mode===m?"#1a4a1a":"#ccc" }}>{l}</button>
            ))}
          </div>
        </div>
        <div style={{ display:"flex",gap:6,alignItems:"center",flexWrap:"wrap" }} className="mobile-toolbar-btns">
          <div style={{ fontSize:12,padding:"4px 12px",borderRadius:20,background:"rgba(255,255,255,0.15)",color:"#fff",minWidth:90,textAlign:"center" }}>
            {saveStatus==="saving"&&"💾 Saving..."}
            {saveStatus==="saved"&&"✅ Saved"}
            {saveStatus==="error"&&"⚠️ Failed"}
            {saveStatus==="idle"&&`🌾 ${farmers.length} farmers`}
          </div>
          <button onClick={mode==="suborgs" ? exportSubOrgData : exportAllData} style={btnStyle}>
            📊 {mode==="suborgs" ? "Export Sub-Org" : "Export Excel"}
          </button>
          {/* Cloud sync status */}
          <div style={{display:"flex",alignItems:"center",gap:6,padding:"4px 10px",borderRadius:6,background:"rgba(0,0,0,0.3)",fontSize:11,color:"#fff"}}>
            {cloudStatus==="saving" && <><span style={{animation:"spin 1s linear infinite",display:"inline-block"}}>⟳</span> Saving...</>}
            {cloudStatus==="saved" && <><span style={{color:"#7dd87d"}}>☁️ ✓</span> {cloudLastSaved}</>}
            {cloudStatus==="error" && <><span style={{color:"#e74c3c"}}>☁️ ✗</span> Sync error</>}
            {cloudStatus==="idle" && <><span style={{color:"#aaa"}}>☁️</span> Cloud ready</>}
          </div>
          <button onClick={()=>setShowSettings(true)} style={{...btnStyle, background:"rgba(80,80,80,0.6)"}}>⚙️ Settings</button>
          <button onClick={mode==="suborgs"?downloadSubOrgTemplate:downloadTemplate} style={btnStyle}>📥 Template</button>
          <label style={{...btnStyle,display:"inline-block"}}>📤 Upload Excel<input type="file" accept=".xlsx,.xls,.csv" onChange={mode==="suborgs"?handleSubOrgExcelUpload:handleExcelUpload} style={{display:"none"}} /></label>
          {mode==="farmers" && <label style={{...btnStyle,display:"inline-block",background:"rgba(230,126,34,0.7)"}}>🔄 Update C/o<input type="file" accept=".xlsx,.xls,.csv" onChange={handleCareOfUpdate} style={{display:"none"}} /></label>}
        </div>
      </div>

      {/* Save Failed Banner */}
      {saveStatus === "error" && (
        <div style={{ background:"#c0392b",color:"#fff",padding:"10px 20px",display:"flex",alignItems:"center",justifyContent:"space-between",gap:12,flexWrap:"wrap" }}>
          <div style={{ display:"flex",alignItems:"center",gap:10 }}>
            <span style={{ fontSize:20 }}>⚠️</span>
            <div>
              <div style={{ fontWeight:700,fontSize:13 }}>Cloud save failed — your data is NOT saved!</div>
              <div style={{ fontSize:11,opacity:0.9 }}>Please export your data immediately to avoid losing it.</div>
            </div>
          </div>
          <button onClick={()=>{exportAllData();setSaveStatus("idle");}} style={{ background:"#fff",color:"#c0392b",border:"none",borderRadius:6,padding:"8px 18px",fontWeight:700,fontSize:13,cursor:"pointer" }}>💾 Export Data Now</button>
        </div>
      )}


      {/* LOT No Duplicate Warning Banner */}
      {(() => {
        try {
          const lotMap = {};
          // Add farmer crop LOT Nos
          (farmers || []).forEach(f => {
            (f.crops || []).forEach(c => {
              const lot = (c.lotNo || "").trim();
              if (!lot) return;
              if (!lotMap[lot]) lotMap[lot] = [];
              lotMap[lot].push({ type: "farmer", farmerNo: f.farmerNo, name: f.name, village: f.village });
            });
          });
          // Add sub-org grower LOT Nos
          (subOrgs || []).forEach(so => {
            (so.growers || []).forEach(g => {
              const lot = (g.lotNo || "").trim();
              if (!lot) return;
              if (!lotMap[lot]) lotMap[lot] = [];
              lotMap[lot].push({ type: "suborg", soName: so.name, name: g.name, village: g.village });
            });
          });
          const duplicates = Object.entries(lotMap).filter(([, arr]) => arr.length > 1);
          if (duplicates.length === 0) return null;
          return (
            <div style={{ background:"#856404",color:"#fff",padding:"10px 20px" }}>
              <div style={{ display:"flex",alignItems:"center",gap:10,marginBottom:8 }}>
                <span style={{ fontSize:20 }}>⚠️</span>
                <div style={{ fontWeight:700,fontSize:13 }}>
                  LOT No Repeated! {duplicates.length} duplicate LOT No{duplicates.length > 1 ? "s" : ""} found across Farmers and Sub-Orgs
                </div>
              </div>
              <div style={{ display:"flex",flexWrap:"wrap",gap:8 }}>
                {duplicates.map(([lot, arr]) => (
                  <div key={lot} style={{ background:"rgba(0,0,0,0.25)",borderRadius:6,padding:"6px 12px",fontSize:12 }}>
                    <div style={{ fontWeight:700,marginBottom:4 }}>LOT No: {lot}</div>
                    {arr.map((entry, i) => (
                      <div key={i} style={{ opacity:0.9 }}>
                        {entry.type === "farmer"
                          ? `• 👨‍🌾 Farmer #${entry.farmerNo} ${entry.name} — ${entry.village}`
                          : `• 🏢 Sub-Org [${entry.soName}] ${entry.name} — ${entry.village}`
                        }
                      </div>
                    ))}
                  </div>
                ))}
              </div>
            </div>
          );
        } catch { return null; }
      })()}

      {/* ── FARMERS MODE ── */}
      {mode === "farmers" && (
        <div style={{ display:"flex",height:"calc(100vh - 58px)" }}>

          {/* Villages Panel */}
          <div style={{ width:villagesPanelOpen?150:40,background:"#0f2a0f",color:"#fff",flexShrink:0,transition:"width 0.2s ease",display:"flex",flexDirection:"column",borderRight:"1px solid #2d5a2d" }}>
            <button onClick={()=>setVillagesPanelOpen(!villagesPanelOpen)} style={{ width:"100%",background:"#0a1f0a",color:"#8bc88b",border:"none",borderBottom:"1px solid #2d5a2d",padding:"8px 0",cursor:"pointer",fontSize:15,textAlign:"center" }}>{villagesPanelOpen?"◀":"▶"}</button>
            {villagesPanelOpen && (() => {
              const villages = [...new Set(farmers.map(f => f.village?.trim()||""))].sort();
              const filteredVillages = villages.filter(v=>v.toLowerCase().includes(villageSearch.toLowerCase()));
              return (
                <>
                  <div style={{ padding:"6px 10px",fontSize:10,color:"#8bc88b",fontWeight:700,letterSpacing:1,borderBottom:"1px solid #2d5a2d" }}>VILLAGES ({villages.length})</div>
                  <div style={{ padding:"6px 8px", borderBottom:"1px solid #2d5a2d" }}>
                    <input
                      value={villageSearch}
                      onChange={e=>{setVillageSearch(e.target.value);setFarmerPage(0);}}
                      placeholder="🔍 Search village..."
                      style={{ width:"100%",padding:"5px 8px",borderRadius:4,border:"1px solid #3d8a3d",background:"rgba(255,255,255,0.1)",color:"#fff",fontSize:11,boxSizing:"border-box",outline:"none" }}
                    />
                  </div>
                  <div style={{ flex:1,overflowY:"auto" }}>
                    <div onClick={()=>{setSelectedVillage(null);setFarmerSearch("");setFarmerPage(0);setVillageSearch("");}} style={{ padding:"8px 10px",cursor:"pointer",background:selectedVillage===null?"#2d6a2d":"transparent",borderLeft:selectedVillage===null?"3px solid #7dd87d":"3px solid transparent",fontSize:12,fontWeight:600,borderBottom:"1px solid #1a3a1a" }}>🌾 All ({farmers.length})</div>
                    {filteredVillages.map(v => {
                      const count = farmers.filter(f=>(f.village?.trim()||"")===v).length;
                      return (
                        <div key={v} onClick={()=>{setSelectedVillage(v);setFarmerSearch("");setFarmerPage(0);}} style={{ padding:"8px 10px",cursor:"pointer",background:selectedVillage===v?"#2d6a2d":"transparent",borderLeft:selectedVillage===v?"3px solid #7dd87d":"3px solid transparent",borderBottom:"1px solid #1a3a1a" }}>
                          <div style={{ fontSize:12,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>📍 {v||"No Village"}</div>
                          <div style={{ fontSize:10,color:"#8bc88b" }}>{count} farmer{count!==1?"s":""}</div>
                        </div>
                      );
                    })}
                    {filteredVillages.length===0 && (
                      <div style={{ padding:"12px 10px",fontSize:11,color:"#8bc88b",fontStyle:"italic" }}>No villages found</div>
                    )}
                  </div>
                  <div style={{ padding:"8px 10px",borderTop:"1px solid #2d5a2d" }}>
                    <button onClick={()=>{const v=prompt("Enter village name:");if(v?.trim()){addFarmer(v.trim());setSelectedVillage(v.trim());setVillagesPanelOpen(true);setFarmersPanelOpen(true);}}} style={{ width:"100%",background:"rgba(255,255,255,0.08)",color:"#fff",border:"1px dashed rgba(255,255,255,0.3)",borderRadius:4,padding:"5px",cursor:"pointer",fontSize:11 }}>+ New Village</button>
                  </div>
                </>
              );
            })()}
            {!villagesPanelOpen && <div style={{ display:"flex",flexDirection:"column",alignItems:"center",paddingTop:8,gap:4 }}><div style={{ fontSize:10,color:"#8bc88b",fontWeight:700 }}>{[...new Set(farmers.map(f=>f.village?.trim()||""))].length}</div><div style={{ fontSize:9,color:"#8bc88b" }}>vil</div></div>}
          </div>

          {/* Farmers Panel */}
          {(() => {
            const searchLower = farmerSearch.toLowerCase().trim();
            const vf = farmers.map((f,i)=>({...f,_idx:i}))
              .filter(f=>selectedVillage===null||(f.village?.trim()||"")===selectedVillage)
              .filter(f=>!searchLower||
                f.name?.toLowerCase().includes(searchLower)||
                f.farmerNo?.toLowerCase().includes(searchLower)||
                f.fatherName?.toLowerCase().includes(searchLower)||
                f.village?.toLowerCase().includes(searchLower)||
                (f.crops||[]).some(c=>c.variety?.toLowerCase().includes(searchLower))
              );
            return (
              <div style={{ width:farmersPanelOpen?190:40,background:"#1a3a1a",color:"#fff",flexShrink:0,display:"flex",flexDirection:"column",borderRight:"1px solid #2d5a2d",transition:"width 0.2s ease",overflow:"hidden" }}>
                <button onClick={()=>setFarmersPanelOpen(!farmersPanelOpen)} style={{ width:"100%",background:"#142e14",color:"#8bc88b",border:"none",borderBottom:"1px solid #2d5a2d",padding:"8px 0",cursor:"pointer",fontSize:15,textAlign:"center" }}>{farmersPanelOpen?"◀":"▶"}</button>
                {farmersPanelOpen && (
                  <>
                    <div style={{ padding:"6px 10px",background:"#142e14",borderBottom:"1px solid #2d5a2d" }}>
                      <div style={{ fontSize:11,color:"#7dd87d",fontWeight:700,marginBottom:5 }}>{selectedVillage?`📍 ${selectedVillage}`:"🌾 All"} ({vf.length})</div>
                      <input value={farmerSearch} onChange={e=>{setFarmerSearch(e.target.value);setFarmerPage(0);}} placeholder="🔍 Search name, village, farmer no..." style={{ width:"100%",padding:"4px 7px",borderRadius:5,border:"1px solid #3a6a3a",background:"#0f2a0f",color:"#fff",fontSize:11,boxSizing:"border-box",outline:"none" }} />
                    </div>
                    <div style={{ flex:1,overflowY:"auto" }}>
                      {vf.length===0&&<div style={{ padding:14,textAlign:"center",color:"#8bc88b",fontSize:12 }}>No farmers found</div>}
                      {vf.slice(farmerPage*FARMERS_PER_PAGE,(farmerPage+1)*FARMERS_PER_PAGE).map(f => (
                        <div key={f._idx} onClick={()=>{setSelectedIdx(f._idx);setTab("form");}} style={{ padding:"7px 10px",cursor:"pointer",background:selectedIdx===f._idx?"#2d6a2d":"transparent",borderLeft:selectedIdx===f._idx?"3px solid #7dd87d":"3px solid transparent",borderBottom:"1px solid #1f4a1f" }}>
                          <div style={{ display:"flex",alignItems:"center",gap:4 }}>
                            {f.farmerNo&&<span style={{ fontSize:9,background:"rgba(255,255,255,0.2)",padding:"1px 5px",borderRadius:8,fontWeight:700 }}>#{f.farmerNo}</span>}
                            <span style={{ fontSize:12,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>{f.name||`Farmer #${f._idx+1}`}</span>
                          </div>
                          {f.fatherName&&<div style={{ fontSize:10,color:"#8bc88b",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>S/o {f.fatherName}</div>}
                          {f.careOf&&<div style={{ fontSize:10,color:"#ffb74d",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>C/o {f.careOf}</div>}
                          {f.billingDone&&<div style={{ fontSize:9,color:"#7dd87d" }}>✔ Billed</div>}
                        </div>
                      ))}
                      {/* Pagination controls */}
                      {vf.length > FARMERS_PER_PAGE && (
                        <div style={{ padding:"6px 4px",borderTop:"1px solid #2d5a2d",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:2 }}>
                          <button onClick={()=>setFarmerPage(p=>Math.max(0,p-1))} disabled={farmerPage===0}
                            style={{ background:"rgba(255,255,255,0.1)",color:farmerPage===0?"#555":"#fff",border:"none",borderRadius:4,padding:"3px 8px",cursor:farmerPage===0?"default":"pointer",fontSize:12 }}>◀</button>
                          <span style={{ fontSize:10,color:"#8bc88b" }}>{farmerPage*FARMERS_PER_PAGE+1}-{Math.min((farmerPage+1)*FARMERS_PER_PAGE,vf.length)} of {vf.length}</span>
                          <button onClick={()=>setFarmerPage(p=>Math.min(Math.ceil(vf.length/FARMERS_PER_PAGE)-1,p+1))} disabled={(farmerPage+1)*FARMERS_PER_PAGE>=vf.length}
                            style={{ background:"rgba(255,255,255,0.1)",color:(farmerPage+1)*FARMERS_PER_PAGE>=vf.length?"#555":"#fff",border:"none",borderRadius:4,padding:"3px 8px",cursor:(farmerPage+1)*FARMERS_PER_PAGE>=vf.length?"default":"pointer",fontSize:12 }}>▶</button>
                        </div>
                      )}
                    </div>
                    <div style={{ padding:"8px 10px",borderTop:"1px solid #2d5a2d" }}>
                      <button onClick={()=>addFarmer(selectedVillage||"")} style={{ width:"100%",background:"rgba(255,255,255,0.1)",color:"#fff",border:"1px dashed rgba(255,255,255,0.3)",borderRadius:4,padding:"6px",cursor:"pointer",fontSize:12 }}>+ Add Farmer{selectedVillage?` to ${selectedVillage}`:""}</button>
                    </div>
                  </>
                )}
                {!farmersPanelOpen&&<div style={{ display:"flex",flexDirection:"column",alignItems:"center",paddingTop:8,gap:4 }}><div style={{ fontSize:10,color:"#8bc88b",fontWeight:700 }}>{vf.length}</div><div style={{ fontSize:9,color:"#8bc88b" }}>fmr</div></div>}
              </div>
            );
          })()}

          {/* Main Area */}
          <div style={{ flex:1,overflowY:"auto",padding:"14px 16px" }}>

            <div style={{ display:"flex",borderBottom:"2px solid #b8d8b8",marginBottom:14 }}>
              {[["form","✏️ Edit"],["preview","👁 Preview"],["all","📋 All Bills"]].map(([t,l]) => (
                <button key={t} onClick={()=>setTab(t)} style={{ padding:"7px 16px",border:"none",borderBottom:tab===t?"3px solid #2d6a2d":"3px solid transparent",background:"transparent",fontWeight:tab===t?700:400,color:tab===t?"#1a4a1a":"#555",cursor:"pointer",fontSize:13,marginBottom:-2 }}>{l}</button>
              ))}
            </div>
            {tab==="form"&&currentFarmer&&<FarmerForm farmer={currentFarmer} index={selectedIdx} onChange={updated=>{const copy=[...farmers];copy[selectedIdx]=updated;updateFarmers(copy);}} onRemove={()=>deleteWithUndo("farmer", selectedIdx)} varietySettings={varietySettings} getVarietyRate={getVarietyRate} getVarietyType={getVarietyType} farmers={farmers} subOrgs={subOrgs} pesticideList={pesticideList} />}
            {tab==="preview"&&currentFarmer&&(
              <div>
                <button onClick={()=>printBill('bill-single',`bill_${currentFarmer.farmerNo||currentFarmer.name||'farmer'}`)} style={{ marginBottom:12,background:"#2d6a2d",color:"#fff",border:"none",borderRadius:5,padding:"7px 18px",cursor:"pointer",fontSize:13 }}>🖨 Print / Save as PDF</button>
                <div id="bill-single"><BillPreview farmer={currentFarmer} varietySettings={varietySettings} getVarietyBillDate={getVarietyBillDate} isVarietyPaid={isVarietyPaid} getVarietyRate={getVarietyRate} getVarietyType={getVarietyType} /></div>
              </div>
            )}
            {tab==="all"&&(()=>{
              // Filter farmers based on selected varieties
              // A farmer is included if they have at least one crop whose variety is in selectedPrintVarieties (and is paid+pass)
              // If nothing selected, show all farmers
              const noSeedSelected = selectedPrintVarieties.includes("__NO_SEED__");
              const paidVarSelected = selectedPrintVarieties.filter(v=>v!=="__NO_SEED__");
              const filteredFarmers = selectedPrintVarieties.length === 0
                ? farmers
                : farmers.filter(f => {
                    const hasSelectedVariety = paidVarSelected.length>0 && (f.crops||[]).some(c=>paidVarSelected.includes(c.variety)&&isVarietyPaid(c.variety));
                    const isNoSeed = noSeedSelected && (f.advances||[]).length>0 && !(f.crops||[]).some(c=>c.variety&&c.variety.trim());
                    return hasSelectedVariety || isNoSeed;
                  });

              // Paid varieties for checkboxes
              const paidVarieties = allVarieties.filter(v => isVarietyPaid(v));

              return (
                <div>
                  {/* ── Billing Status Overview ── */}
                  {(() => {
                    const doneFarmers = farmers.filter(f=>f.billingDone);
                    const pendingFarmers = farmers.filter(f=>!f.billingDone);

                    // Categorize pending farmers by balance type (computed same way as Variety Pay tab)
                    const getFarmerBalance = (f) => {
                      let cropVal=0, found=0, trans=0;
                      (f.crops||[]).forEach(c=>{
                        const qty=parseFloat(c.quantity)||0;
                        const vRate=(c.rateOverride===true)?(parseFloat(c.ratePerUnit)||0):(getVarietyRate(c.variety)||parseFloat(c.ratePerUnit)||0);
                        const paid=isVarietyPaid(c.variety);
                        if(c.result==="Pass"&&paid) cropVal+=qty*vRate;
                        found+=(parseFloat(c.area)||0)*1000;
                        trans+=qty;
                      });
                      const paidDates=(f.crops||[]).filter(c=>c.result==="Pass"&&isVarietyPaid(c.variety)).map(c=>getVarietyBillDate(c.variety));
                      const billDate=paidDates.length>0?paidDates.reduce((a,b)=>a<b?a:b):BILL_DATE;
                      const advWI=(f.advances||[]).reduce((s,a)=>{const bd=a.tillDate||billDate;const {interest}=(a.compound?calcCompoundInterest:calcInterest)(a.amount,a.interestRate,a.date,bd);return s+a.amount+interest;},0);
                      const jamWI=(f.jammaEnabled?f.jammaEntries||[]:[]).reduce((s,j)=>{const bd=j.tillDate||billDate;const {interest}=calcInterest(parseFloat(j.amount)||0,parseFloat(j.interestRate)||0,j.date||BILL_DATE,bd);return s+(parseFloat(j.amount)||0)+interest;},0);
                      return cropVal-advWI+jamWI-found-trans;
                    };

                    const pendingWithBalance = pendingFarmers.map(f=>({f, balance:getFarmerBalance(f)}));
                    const pendingDue = pendingWithBalance.filter(x=>x.balance<0);
                    const pendingPay = pendingWithBalance.filter(x=>x.balance>=0);

                    const exportPendingFarmers = () => {
                      const noSeedIds = new Set(farmers.filter(f=>(f.advances||[]).length>0&&!(f.crops||[]).some(c=>c.variety&&c.variety.trim())).map(f=>f.id));
                      const allPending = [
                        ...pendingWithBalance.filter(({f})=>!noSeedIds.has(f.id)),
                        // Also add No Seed farmers (they have advances but no crops)
                        ...(farmers.filter(f=>(f.advances||[]).length>0&&!(f.crops||[]).some(c=>c.variety&&c.variety.trim())&&!f.billingDone)).map(f=>{
                          const billDate = BILL_DATE;
                          const advWI = (f.advances||[]).reduce((s,a)=>{
                            const bd=a.tillDate||billDate;
                            const {interest}=(a.compound?calcCompoundInterest:calcInterest)(parseFloat(a.amount)||0,parseFloat(a.interestRate)||0,a.date,bd);
                            return s+(parseFloat(a.amount)||0)+interest;
                          },0);
                          return {f, balance:-advWI, noSeed:true};
                        })
                      ];
                      const rows = allPending
                        .sort((a,b)=>a.balance-b.balance)
                        .map(({f,balance,noSeed})=>({
                          "Farmer No": f.farmerNo||"",
                          "Name": f.name||"",
                          "Father Name": f.fatherName||"",
                          "Village": f.village||"",
                          "Category": noSeed ? "🚫 No Seed — Advance Only" : balance<0 ? "Balance Due (they owe you)" : "Balance to Pay (you owe them)",
                          "Amount (₹)": Math.round(Math.abs(balance)),
                        }));
                      if (rows.length===0) { alert("No pending farmers! All billing is done."); return; }
                      const ws = XLSX.utils.json_to_sheet(rows);
                      ws["!cols"] = [{wch:10},{wch:22},{wch:20},{wch:18},{wch:28},{wch:14}];
                      const wb = XLSX.utils.book_new();
                      XLSX.utils.book_append_sheet(wb, ws, "Pending Farmers");
                      XLSX.writeFile(wb, `pending_farmers_${new Date().toISOString().split("T")[0]}.xlsx`);
                    };

                    return (
                      <div style={{ background:"#fff",border:"1px solid #d8e8d8",borderRadius:8,padding:"14px 16px",marginBottom:14 }}>
                        <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,flexWrap:"wrap",gap:8 }}>
                          <div style={{ fontSize:13,fontWeight:700,color:"#1a4a1a" }}>📋 Billing Status</div>
                          <button onClick={exportPendingFarmers} style={{ background:"#2d5a8a",color:"#fff",border:"none",borderRadius:6,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer" }}>
                            📊 Export Pending Farmers ({pendingFarmers.length})
                          </button>
                        </div>
                        <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:10 }}>
                          <div style={{ background:"#e8f5e9",border:"1.5px solid #2d6a2d",borderRadius:8,padding:"10px 12px" }}>
                            <div style={{ fontSize:11,color:"#2d6a2d",fontWeight:700 }}>✔ Billing Done</div>
                            <div style={{ fontSize:20,fontWeight:800,color:"#1a4a1a" }}>{doneFarmers.length}</div>
                          </div>
                          <div style={{ background:"#fdecea",border:"1.5px solid #e74c3c",borderRadius:8,padding:"10px 12px" }}>
                            <div style={{ fontSize:11,color:"#c0392b",fontWeight:700 }}>⏳ Pending — Balance Due</div>
                            <div style={{ fontSize:20,fontWeight:800,color:"#c0392b" }}>{pendingDue.length}</div>
                          </div>
                          <div style={{ background:"#fff3cd",border:"1.5px solid #c8a000",borderRadius:8,padding:"10px 12px" }}>
                            <div style={{ fontSize:11,color:"#856404",fontWeight:700 }}>⏳ Pending — Balance to Pay</div>
                            <div style={{ fontSize:20,fontWeight:800,color:"#856404" }}>{pendingPay.length}</div>
                          </div>
                          <div style={{ background:"#fff0f0",border:"1.5px solid #c0392b",borderRadius:8,padding:"10px 12px" }}>
                            <div style={{ fontSize:11,color:"#c0392b",fontWeight:700 }}>🚫 No Seed — Advance Only</div>
                            <div style={{ fontSize:20,fontWeight:800,color:"#c0392b" }}>{farmers.filter(f=>(f.advances||[]).length>0&&!(f.crops||[]).some(c=>c.variety&&c.variety.trim())&&!f.billingDone).length}</div>
                          </div>
                        </div>
                      </div>
                    );
                  })()}

                  {/* ── Variety Filter ── */}
                  {/* No Seed filter */}
                  {(() => {
                    const noSeedFarmers = farmers.filter(f=>(f.advances||[]).length>0 && !(f.crops||[]).some(c=>c.variety&&c.variety.trim()));
                    if (noSeedFarmers.length === 0) return null;
                    const noSeedSelected = selectedPrintVarieties.includes("__NO_SEED__");
                    return (
                      <div onClick={()=>setSelectedPrintVarieties(noSeedSelected?selectedPrintVarieties.filter(v=>v!=="__NO_SEED__"):[...selectedPrintVarieties,"__NO_SEED__"])}
                        style={{display:"flex",alignItems:"center",gap:8,padding:"8px 12px",borderRadius:8,cursor:"pointer",
                          border:`2px solid ${noSeedSelected?"#c0392b":"#e0b0b0"}`,
                          background:noSeedSelected?"#fdecea":"#fff",marginBottom:8}}>
                        <div style={{width:16,height:16,borderRadius:3,border:`2px solid ${noSeedSelected?"#c0392b":"#aaa"}`,background:noSeedSelected?"#c0392b":"#fff",display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontSize:12,fontWeight:700,flexShrink:0}}>{noSeedSelected?"✓":""}</div>
                        <div>
                          <div style={{fontWeight:700,fontSize:12,color:"#c0392b"}}>🚫 No Seed Farmers</div>
                          <div style={{fontSize:10,color:"#888"}}>{noSeedFarmers.length} farmers — advance only</div>
                        </div>
                      </div>
                    );
                  })()}
                  <div style={{ background:"#f0f7f0",border:"1px solid #c8dfc8",borderRadius:8,padding:"12px 16px",marginBottom:14 }}>
                    <div style={{ fontSize:12,fontWeight:700,color:"#1a4a1a",marginBottom:8 }}>
                      🖨 Select Varieties to Print — Only farmers with these varieties will be included
                    </div>
                    <div style={{ display:"flex",flexWrap:"wrap",gap:8,marginBottom:10 }}>
                      {paidVarieties.length === 0 ? (
                        <div style={{ fontSize:12,color:"#aaa" }}>No paid varieties yet. Go to 🌾 Variety Pay to mark varieties as paid.</div>
                      ) : paidVarieties.map(v => {
                        const checked = selectedPrintVarieties.includes(v);
                        const billDate = getVarietyBillDate(v);
                        const fCount = farmers.filter(f=>(f.crops||[]).some(c=>c.variety===v&&c.result==="Pass")).length;
                        return (
                          <div key={v} onClick={()=>{
                            setSelectedPrintVarieties(prev =>
                              prev.includes(v) ? prev.filter(x=>x!==v) : [...prev, v]
                            );
                          }} style={{
                            display:"flex",alignItems:"center",gap:8,
                            padding:"8px 14px",borderRadius:8,cursor:"pointer",
                            border:`2px solid ${checked?"#2d6a2d":"#c8dfc8"}`,
                            background:checked?"#e8f5e9":"#fff",
                            transition:"all 0.15s"
                          }}>
                            <div style={{
                              width:18,height:18,borderRadius:4,border:`2px solid ${checked?"#2d6a2d":"#aaa"}`,
                              background:checked?"#2d6a2d":"#fff",display:"flex",alignItems:"center",justifyContent:"center",
                              color:"#fff",fontSize:13,fontWeight:700,flexShrink:0
                            }}>{checked?"✓":""}</div>
                            <div>
                              <div style={{ fontWeight:700,fontSize:13,color:"#1a4a1a" }}>{v}</div>
                              <div style={{ fontSize:10,color:"#666" }}>
                                📅 {billDate} · {fCount} farmers
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                    <div style={{ display:"flex",gap:8,alignItems:"center",flexWrap:"wrap" }}>
                      <button onClick={()=>setSelectedPrintVarieties(paidVarieties)} style={{ background:"#2d6a2d",color:"#fff",border:"none",borderRadius:5,padding:"5px 12px",fontSize:12,cursor:"pointer",fontWeight:600 }}>☑ Select All</button>
                      <button onClick={()=>setSelectedPrintVarieties([])} style={{ background:"#fff",color:"#555",border:"1px solid #c8dfc8",borderRadius:5,padding:"5px 12px",fontSize:12,cursor:"pointer" }}>☐ Clear</button>
                      <span style={{ fontSize:12,color:"#666" }}>
                        {selectedPrintVarieties.length === 0
                          ? `Showing all ${farmers.length} farmers`
                          : `${filteredFarmers.length} farmers selected for printing`}
                      </span>
                      {selectedPrintVarieties.length > 0 && (
                        <button onClick={()=>{ setPrintQueue(filteredFarmers); setPrintQueueIdx(0); setPrintQueueTotal(filteredFarmers.length); }} style={{ background:"#2d6a2d",color:"#fff",border:"none",borderRadius:5,padding:"7px 16px",cursor:"pointer",fontSize:13,fontWeight:700,marginLeft:"auto" }}>
                          🖨 Print {filteredFarmers.length} Bills (One by One)
                        </button>
                      )}
                      {selectedPrintVarieties.length === 0 && (
                        <button onClick={()=>{ setPrintQueue(farmers); setPrintQueueIdx(0); setPrintQueueTotal(farmers.length); }} style={{ background:"#555",color:"#fff",border:"none",borderRadius:5,padding:"7px 16px",cursor:"pointer",fontSize:13,marginLeft:"auto" }}>
                          🖨 Print All {farmers.length} Bills (One by One)
                        </button>
                      )}
                    </div>
                  </div>

                  {/* Bills */}
                  <div id="bill-all">
                    {filteredFarmers.map((f,i)=>(
                      <div key={i} style={{ marginBottom:28,pageBreakAfter:"always" }}>
                        <BillPreview farmer={f} varietySettings={varietySettings} getVarietyBillDate={getVarietyBillDate} isVarietyPaid={isVarietyPaid} getVarietyRate={getVarietyRate} getVarietyType={getVarietyType} />
                      </div>
                    ))}
                    {filteredFarmers.length === 0 && (
                      <div style={{ textAlign:"center",padding:40,color:"#aaa",fontSize:14 }}>
                        No farmers found for selected varieties. Select at least one variety above.
                      </div>
                    )}
                  </div>
                </div>
              );
            })()}
          </div>
        </div>
      )}

      {/* ── SUB-ORGANIZERS MODE ── */}
      {mode === "suborgs" && (
        <div style={{ display:"flex",height:"calc(100vh - 58px)" }}>
          <div style={{ width:subOrgPanelOpen?210:40,background:"#1a2a3a",color:"#fff",flexShrink:0,transition:"width 0.2s ease",display:"flex",flexDirection:"column",borderRight:"1px solid #2d4a6a" }}>
            <button onClick={()=>setSubOrgPanelOpen(!subOrgPanelOpen)} style={{ width:"100%",background:"#111f2e",color:"#7ab8e8",border:"none",borderBottom:"1px solid #2d4a6a",padding:"8px 0",cursor:"pointer",fontSize:15,textAlign:"center" }}>{subOrgPanelOpen?"◀":"▶"}</button>
            {subOrgPanelOpen&&(
              <>
                <div style={{ padding:"6px 10px",fontSize:10,color:"#7ab8e8",fontWeight:700,letterSpacing:1,borderBottom:"1px solid #2d4a6a" }}>SUB-ORGANIZERS ({subOrgs.length})</div>
                <div style={{ flex:1,overflowY:"auto" }}>
                  {subOrgs.length===0&&<div style={{ padding:14,textAlign:"center",color:"#7ab8e8",fontSize:12 }}>No sub-organizers yet</div>}
                  {subOrgs.map((s,i)=>(
                    <div key={i} onClick={()=>{setSelectedSubOrgIdx(i);setSubOrgTab("form");}} style={{ padding:"8px 10px",cursor:"pointer",background:selectedSubOrgIdx===i?"#2d5a8a":"transparent",borderLeft:selectedSubOrgIdx===i?"3px solid #7ab8e8":"3px solid transparent",borderBottom:"1px solid #1f3a5a" }}>
                      <div style={{ display:"flex",alignItems:"center",gap:5 }}>
                        {s.accNo&&<span style={{ fontSize:9,background:"rgba(255,255,255,0.2)",padding:"1px 5px",borderRadius:8,fontWeight:700 }}>#{s.accNo}</span>}
                        <span style={{ fontSize:12,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>{s.name||`Sub-Org #${i+1}`}</span>
                      </div>
                      <div style={{ fontSize:10,color:"#7ab8e8",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>📍 {s.village||"—"} · {(s.growers||[]).length} growers</div>
                    </div>
                  ))}
                </div>
                <div style={{ padding:"8px 10px",borderTop:"1px solid #2d4a6a" }}>
                  <button onClick={addSubOrg} style={{ width:"100%",background:"rgba(255,255,255,0.1)",color:"#fff",border:"1px dashed rgba(255,255,255,0.3)",borderRadius:4,padding:"6px",cursor:"pointer",fontSize:12,marginBottom:4 }}>+ Add Sub-Organizer</button>
                  <button onClick={()=>setSubOrgTab("summary")} style={{width:"100%",background:subOrgTab==="summary"?"rgba(255,200,0,0.3)":"rgba(255,200,0,0.1)",color:"#ffd700",border:"1px dashed #ffd700",borderRadius:4,padding:"6px",cursor:"pointer",fontSize:12,display:"block",textAlign:"center",marginBottom:4}}>
                    📊 All Sub-Org Summary
                  </button>
                  <button onClick={exportSubOrgData} style={{width:"100%",background:"rgba(45,90,138,0.4)",color:"#7ab8e8",border:"1px dashed #7ab8e8",borderRadius:4,padding:"6px",cursor:"pointer",fontSize:12,display:"block",textAlign:"center",marginBottom:4}}>
                    📊 Export Sub-Org Excel
                  </button>
                  <button onClick={downloadSubOrgTemplate} style={{width:"100%",background:"rgba(45,106,45,0.4)",color:"#9ee8a0",border:"1px dashed #9ee8a0",borderRadius:4,padding:"6px",cursor:"pointer",fontSize:12,display:"block",textAlign:"center",marginBottom:4}}>
                    📥 Download Blank Template
                  </button>
                  <label style={{width:"100%",background:"rgba(45,106,45,0.2)",color:"#9ee8a0",border:"1px dashed #9ee8a0",borderRadius:4,padding:"6px",cursor:"pointer",fontSize:12,display:"block",textAlign:"center",marginBottom:4}}>
                    📤 Upload Filled Template
                    <input type="file" accept=".xlsx,.xls" onChange={handleSubOrgTemplateUpload} style={{display:"none"}} />
                  </label>
                <label style={{width:"100%",background:"rgba(122,184,232,0.2)",color:"#7ab8e8",border:"1px dashed #7ab8e8",borderRadius:4,padding:"6px",cursor:"pointer",fontSize:12,display:"block",textAlign:"center"}}>
                  📤 Upload Old-Format Excel
                  <input type="file" accept=".xlsx,.xls" onChange={handleSubOrgExcelUpload} style={{display:"none"}} />
                </label>
                </div>
              </>
            )}
            {!subOrgPanelOpen&&<div style={{ display:"flex",flexDirection:"column",alignItems:"center",paddingTop:8,gap:4 }}><div style={{ fontSize:10,color:"#7ab8e8",fontWeight:700 }}>{subOrgs.length}</div><div style={{ fontSize:9,color:"#7ab8e8" }}>org</div></div>}
          </div>

          <div style={{ flex:1,overflowY:"auto",padding:"14px 16px" }}>
            {subOrgs.length===0?(
              <div style={{ textAlign:"center",marginTop:60,color:"#555" }}>
                <div style={{ fontSize:40 }}>🏢</div>
                <div style={{ fontSize:16,fontWeight:600,marginTop:10 }}>No Sub-Organizers yet</div>
                <div style={{ fontSize:13,color:"#888",marginTop:6 }}>Click "+ Add Sub-Organizer" in the panel</div>
              </div>
            ):subOrgTab==="summary"?(()=>{
              // Calculate balance for each sub-org
              const fmt = n => `₹${Math.abs(Math.round(n)).toLocaleString("en-IN")}`;
              const soSummary = subOrgs.map(so => {
                const totalAdvWI = (so.advances||[]).reduce((s,a) => {
                  const {interest} = (a.compound?calcCompoundInterest:calcInterest)(a.amount, a.interestRate, a.date, BILL_DATE);
                  return s + a.amount + interest;
                }, 0);
                const passGrowers = (so.growers||[]).filter(g=>g.result==="Pass");
                const totalQty = passGrowers.reduce((s,g)=>s+(parseFloat(g.packets)||0),0);
                const totalCropVal = passGrowers.reduce((s,g)=>{
                  const rate = getSubOrgVarietyRate(g.variety)||parseFloat(g.rate)||0;
                  return s+(parseFloat(g.packets)||0)*rate;
                },0);
                const totalFound = (so.foundationSeeds||[]).reduce((s,f)=>s+(parseFloat(f.area)||0)*1000,0);
                const totalTrans = totalQty;
                const totalJammaWI = (so.jammaEntries||[]).reduce((s,j)=>{
                  const {interest} = calcInterest(parseFloat(j.amount)||0, parseFloat(j.interestRate)||0, j.date||BILL_DATE, BILL_DATE);
                  return s+(parseFloat(j.amount)||0)+interest;
                },0);
                const balance = totalCropVal - totalAdvWI + totalJammaWI - totalFound - totalTrans;
                return { so, totalAdvWI, totalQty, totalCropVal, balance };
              });
              const grandAdv = soSummary.reduce((s,x)=>s+x.totalAdvWI,0);
              const grandQty = soSummary.reduce((s,x)=>s+x.totalQty,0);
              const grandCrop = soSummary.reduce((s,x)=>s+x.totalCropVal,0);
              const grandPay = soSummary.filter(x=>x.balance>=0).reduce((s,x)=>s+x.balance,0);
              const grandDue = soSummary.filter(x=>x.balance<0).reduce((s,x)=>s+Math.abs(x.balance),0);
              return (
                <div>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexWrap:"wrap",gap:10}}>
                    <div style={{fontWeight:800,fontSize:18,color:"#1a2a4a"}}>📊 Sub-Org Summary</div>
                    <button onClick={()=>{
                      const rows = soSummary.map((x,i)=>{
                        const {so,totalAdvWI,totalQty,totalCropVal,balance}=x;
                        return `<tr style="background:${i%2===0?"#f5f8ff":"#fff"}">
                          <td style="padding:6px 10px;border:1px solid #ddd;">${so.accNo||"—"}</td>
                          <td style="padding:6px 10px;border:1px solid #ddd;font-weight:600;">${so.name||"—"}</td>
                          <td style="padding:6px 10px;border:1px solid #ddd;">${so.village||"—"}</td>
                          <td style="padding:6px 10px;border:1px solid #ddd;text-align:right;">${(so.growers||[]).length}</td>
                          <td style="padding:6px 10px;border:1px solid #ddd;text-align:right;">${totalQty.toLocaleString("en-IN")}</td>
                          <td style="padding:6px 10px;border:1px solid #ddd;text-align:right;">₹${Math.round(totalCropVal).toLocaleString("en-IN")}</td>
                          <td style="padding:6px 10px;border:1px solid #ddd;text-align:right;color:#c0392b;">₹${Math.round(totalAdvWI).toLocaleString("en-IN")}</td>
                          <td style="padding:6px 10px;border:1px solid #ddd;text-align:right;font-weight:700;color:${balance>=0?"#1a5c1a":"#c0392b"};">₹${Math.abs(Math.round(balance)).toLocaleString("en-IN")} ${balance>=0?"(Pay)":"(Due)"}</td>
                        </tr>`;
                      }).join("");
                      const html=`<!DOCTYPE html><html><head><meta charset="UTF-8"/><title>Sub-Org Summary</title>
                        <style>body{font-family:Georgia,serif;padding:20px;}h2{color:#1a2a4a;}table{border-collapse:collapse;width:100%;font-size:12px;}th{background:#1a2a4a;color:#fff;padding:7px 10px;border:1px solid #ddd;}.total-row td{font-weight:800;background:#e8f0ff;border-top:2px solid #1a2a4a;}@media print{@page{margin:8mm;size:A4 landscape;}}</style>
                        </head><body><h2>📊 Sub-Org Summary</h2>
                        <div style="font-size:13px;color:#555;margin-bottom:10px;">Bill Date: ${fmtDate(BILL_DATE)} | Total Sub-Orgs: ${subOrgs.length}</div>
                        <table><thead><tr><th>Acc No</th><th>Sub-Org Name</th><th>Village</th><th>Growers</th><th>Qty</th><th>Crop Value</th><th>Adv+Int</th><th>Bal to Pay</th><th>Bal Due</th></tr></thead>
                        <tbody>${rows}
                        <tr class="total-row"><td colspan="3">GRAND TOTAL</td>
                          <td style="text-align:right;padding:6px 10px;border:1px solid #ddd;">${subOrgs.reduce((s,so)=>s+(so.growers||[]).length,0)}</td>
                          <td style="text-align:right;padding:6px 10px;border:1px solid #ddd;">${grandQty.toLocaleString("en-IN")}</td>
                          <td style="text-align:right;padding:6px 10px;border:1px solid #ddd;">₹${Math.round(grandCrop).toLocaleString("en-IN")}</td>
                          <td style="text-align:right;padding:6px 10px;border:1px solid #ddd;color:#c0392b;">₹${Math.round(grandAdv).toLocaleString("en-IN")}</td>
                          <td style="text-align:right;padding:6px 10px;border:1px solid #ddd;font-weight:800;color:#1a5c1a;">₹${Math.round(grandPay).toLocaleString("en-IN")}</td>
          <td style="text-align:right;padding:6px 10px;border:1px solid #ddd;font-weight:800;color:#c0392b;">₹${Math.round(grandDue).toLocaleString("en-IN")}</td>
                        </tr></tbody></table></body></html>`;
                      const w=window.open("","_blank"); w.document.write(html); w.document.close(); setTimeout(()=>w.print(),400);
                    }} style={{background:"#1a2a4a",color:"#fff",border:"none",borderRadius:6,padding:"9px 16px",fontWeight:700,fontSize:13,cursor:"pointer"}}>🖨️ Print Summary</button>
                  </div>

                  {/* Summary cards */}
                  <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:16}}>
                    {[
                      ["Total Sub-Orgs",subOrgs.length,"#1a2a4a","#e8f0ff"],
                      ["Total Crop Value",fmt(grandCrop),"#1a5c1a","#e8f5e9"],
                      ["Total Bal to Pay",fmt(grandPay),"#1a5c1a","#e8f5e9"],
                      ["Total Bal Due",fmt(grandDue),"#c0392b","#fdecea"],
                    ].map(([label,val,color,bg])=>(
                      <div key={label} style={{background:bg,border:`1px solid ${color}30`,borderRadius:8,padding:"12px 14px"}}>
                        <div style={{fontSize:11,color:"#555",marginBottom:4}}>{label}</div>
                        <div style={{fontSize:18,fontWeight:800,color}}>{val}</div>
                      </div>
                    ))}
                  </div>

                  {/* Summary table */}
                  <div style={{background:"#fff",border:"1px solid #c8d8e8",borderRadius:8,overflow:"hidden"}}>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                      <thead>
                        <tr style={{background:"#1a2a4a",color:"#fff"}}>
                          {["Acc No","Sub-Org Name","Village","Growers","Total Qty","Crop Value","Adv+Int","Bal to Pay","Bal Due"].map(h=>(
                            <th key={h} style={{padding:"8px 10px",textAlign:h==="Sub-Org Name"||h==="Village"?"left":"center",fontWeight:600,fontSize:12}}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {soSummary.map(({so,totalAdvWI,totalQty,totalCropVal,balance},i)=>(
                          <tr key={i} onClick={()=>{setSelectedSubOrgIdx(i);setSubOrgTab("form");}} style={{background:i%2===0?"#f5f8ff":"#fff",cursor:"pointer",borderBottom:"1px solid #d0e4f4"}}
                            onMouseEnter={e=>e.currentTarget.style.background="#e8f0ff"}
                            onMouseLeave={e=>e.currentTarget.style.background=i%2===0?"#f5f8ff":"#fff"}>
                            <td style={{padding:"8px 10px",textAlign:"center",fontWeight:700,color:"#2d5a8a"}}>{so.accNo||"—"}</td>
                            <td style={{padding:"8px 10px",fontWeight:600}}>{so.name||"—"}</td>
                            <td style={{padding:"8px 10px",color:"#555"}}>{so.village||"—"}</td>
                            <td style={{padding:"8px 10px",textAlign:"center"}}>{(so.growers||[]).length}</td>
                            <td style={{padding:"8px 10px",textAlign:"center",fontWeight:600}}>{totalQty.toLocaleString("en-IN")}</td>
                            <td style={{padding:"8px 10px",textAlign:"right",fontWeight:600,color:"#1a5c1a"}}>{fmt(totalCropVal)}</td>
                            <td style={{padding:"8px 10px",textAlign:"right",color:"#c0392b"}}>{fmt(totalAdvWI)}</td>
                            <td style={{padding:"8px 10px",textAlign:"right",fontWeight:700,color:"#1a5c1a"}}>
                              {balance>=0?fmt(balance):"—"}
                            </td>
                            <td style={{padding:"8px 10px",textAlign:"right",fontWeight:700,color:"#c0392b"}}>
                              {balance<0?fmt(Math.abs(balance)):"—"}
                            </td>
                          </tr>
                        ))}
                        <tr style={{background:"#e8f0ff",fontWeight:800,borderTop:"2px solid #1a2a4a"}}>
                          <td colSpan={3} style={{padding:"8px 10px",color:"#1a2a4a"}}>GRAND TOTAL</td>
                          <td style={{padding:"8px 10px",textAlign:"center"}}>{subOrgs.reduce((s,so)=>s+(so.growers||[]).length,0)}</td>
                          <td style={{padding:"8px 10px",textAlign:"center"}}>{grandQty.toLocaleString("en-IN")}</td>
                          <td style={{padding:"8px 10px",textAlign:"right",color:"#1a5c1a"}}>{fmt(grandCrop)}</td>
                          <td style={{padding:"8px 10px",textAlign:"right",color:"#c0392b"}}>{fmt(grandAdv)}</td>
                          <td style={{padding:"8px 10px",textAlign:"right",fontWeight:700,color:"#1a5c1a"}}>{fmt(grandPay)}</td>
                          <td style={{padding:"8px 10px",textAlign:"right",fontWeight:700,color:"#c0392b"}}>{fmt(grandDue)}</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>
              );
            })():currentSubOrg&&(()=>{
              const so=currentSubOrg;
              const growers=so.growers||[];
              const updateSO=(updated)=>{const copy=[...subOrgs];copy[selectedSubOrgIdx]=updated;updateSubOrgs(copy);};
              const updGrower=(i,f,v)=>{const g=[...growers];g[i]={...g[i],[f]:["packets","rate"].includes(f)?parseFloat(v)||0:v};updateSO({...so,growers:g});};
              const th2={padding:"4px 7px",textAlign:"center",fontWeight:600,fontSize:11,whiteSpace:"nowrap"};
              const td2={padding:"4px 7px",textAlign:"center",fontSize:11,whiteSpace:"nowrap"};
              return (
                <>
                  <div style={{ display:"flex",borderBottom:"2px solid #b0c8e0",marginBottom:14 }}>
                    {[["form","✏️ Edit"],["growers","👥 Growers"],["bill","📄 Bill"]].map(([t,l])=>(
                      <button key={t} onClick={()=>setSubOrgTab(t)} style={{ padding:"7px 16px",border:"none",borderBottom:subOrgTab===t?"3px solid #2d5a8a":"3px solid transparent",background:"transparent",fontWeight:subOrgTab===t?700:400,color:subOrgTab===t?"#1a2a4a":"#555",cursor:"pointer",fontSize:13,marginBottom:-2 }}>{l}</button>
                    ))}
                    <button onClick={()=>deleteWithUndo("suborg", selectedSubOrgIdx)} style={{ marginLeft:"auto",background:"#fdecea",color:"#c0392b",border:"1px solid #f5c6cb",borderRadius:4,padding:"4px 12px",cursor:"pointer",fontSize:12 }}>🗑 Remove</button>
                  </div>
                  {/* Billing History */}
                  {(billingHistory[so.accNo||so.name]||[]).length > 0 && (
                    <div style={{marginTop:12,background:"#f0f5ff",borderRadius:8,padding:"10px 14px",border:"1px solid #b0c8e0"}}>
                      <div style={{fontWeight:700,fontSize:12,color:"#2d5a8a",marginBottom:8}}>📋 Billing History</div>
                      {(billingHistory[so.accNo||so.name]||[]).map((h,hi)=>(
                        <div key={hi} style={{display:"flex",gap:10,alignItems:"flex-start",padding:"6px 0",borderBottom:"1px solid #d0e4f4",fontSize:12}}>
                          <div style={{minWidth:120,color:"#555",fontWeight:600}}>
                            Bill #{hi+1} — {h.date} {h.time}
                          </div>
                          <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                            {(h.varieties||[]).map(v=>(
                              <span key={v} style={{background:"#2d5a8a",color:"#fff",borderRadius:10,padding:"2px 8px",fontSize:11}}>{v}</span>
                            ))}
                            {(h.varieties||[]).length===0&&<span style={{color:"#aaa",fontSize:11}}>All varieties</span>}
                          </div>
                        </div>
                      ))}
                      <button onClick={()=>{const h={...billingHistory};delete h[so.accNo||so.name];saveBillingHistory(h);}} style={{marginTop:6,background:"none",border:"none",color:"#c0392b",cursor:"pointer",fontSize:11}}>🗑 Clear history</button>
                    </div>
                  )}

                  {subOrgTab==="form"&&(
                    <div style={{ border:"1.5px solid #b0c8e0",borderRadius:8,padding:16,background:"#f5faff" }}>
                      <div style={{ display:"grid",gridTemplateColumns:"0.6fr 1fr 1fr 1fr",gap:8,marginBottom:14 }}>
                        <div><label style={{ fontSize:11,color:"#555",display:"block",marginBottom:2 }}>Acc No</label><input {...inp2} value={so.accNo||""} onChange={e=>updateSO({...so,accNo:e.target.value})} placeholder="e.g. 695" style={{...inp2.style,fontWeight:700}} /></div>
                        <div><label style={{ fontSize:11,color:"#555",display:"block",marginBottom:2 }}>Sub-Organizer Name</label><input {...inp2} value={so.name||""} onChange={e=>updateSO({...so,name:e.target.value})} /></div>
                        <div><label style={{ fontSize:11,color:"#555",display:"block",marginBottom:2 }}>Father Name</label><input {...inp2} value={so.fatherName||""} onChange={e=>updateSO({...so,fatherName:e.target.value})} /></div>
                        <div><label style={{ fontSize:11,color:"#555",display:"block",marginBottom:2 }}>Village (MSP)</label><input {...inp2} value={so.village||""} onChange={e=>updateSO({...so,village:e.target.value})} /></div>
                      </div>
                      <div style={{ marginBottom:14 }}>
                        <input
                          placeholder="📝 Comment for this sub-org's bill (optional) — e.g. Discussed quality concerns, Promised to settle next visit..."
                          value={so.comment||""}
                          onChange={e=>updateSO({...so,comment:e.target.value})}
                          style={{width:"100%",padding:"7px 10px",border:"1.5px dashed #c8a000",borderRadius:5,fontSize:12,background:"#fffdf5",color:"#856404"}}
                        />
                      </div>
                      <div>
                        <div style={{ fontWeight:600,fontSize:12,color:"#2d5a8a",marginBottom:6 }}>Advances</div>
                        {(so.advances||[]).map((a,i)=>(
                          <div key={i} style={{ marginBottom:8 }}>
                            <div style={{ display:"grid",gridTemplateColumns:"1.2fr 1fr 0.7fr 1fr 32px",gap:6,alignItems:"end" }}>
                              <div><label style={{fontSize:10,color:"#666",display:"block"}}>Date</label><input {...inp2} type="date" value={a.date||""} onChange={e=>{const arr=[...so.advances];arr[i]={...arr[i],date:e.target.value};updateSO({...so,advances:arr});}} /></div>
                              <div><label style={{fontSize:10,color:"#666",display:"block"}}>Amount ₹</label><input {...inp2} type="number" value={a.amount||""} onChange={e=>{const arr=[...so.advances];arr[i]={...arr[i],amount:parseFloat(e.target.value)||0};updateSO({...so,advances:arr});}} /></div>
                              <div><label style={{fontSize:10,color:"#666",display:"block"}}>Interest %</label><input {...inp2} type="number" step="0.5" value={a.interestRate||""} onChange={e=>{const arr=[...so.advances];arr[i]={...arr[i],interestRate:parseFloat(e.target.value)||0};updateSO({...so,advances:arr});}} /></div>
                              <div><label style={{fontSize:10,color:"#666",display:"block"}}>Note</label><input {...inp2} value={a.note||""} onChange={e=>{const arr=[...so.advances];arr[i]={...arr[i],note:e.target.value};updateSO({...so,advances:arr});}} /></div>
                              <button onClick={()=>updateSO({...so,advances:so.advances.filter((_,j)=>j!==i)})} style={{background:"#fdecea",color:"#e74c3c",border:"1px solid #e74c3c",borderRadius:4,padding:"5px 6px",cursor:"pointer",fontSize:11}}>✕</button>
                            </div>
                            {/* Till Date override */}
                            <div style={{display:"flex",alignItems:"center",gap:8,marginTop:4,paddingLeft:4}}>
                              {a.tillDate ? (
                                <>
                                  <label style={{fontSize:10,color:"#e67e22",fontWeight:600}}>⚠ Interest Till:</label>
                                  <input type="date" value={a.tillDate} onChange={e=>{const arr=[...so.advances];arr[i]={...arr[i],tillDate:e.target.value};updateSO({...so,advances:arr});}}
                                    style={{padding:"2px 6px",border:"1px solid #e67e22",borderRadius:4,fontSize:11,color:"#e67e22"}} />
                                  <button onClick={()=>{const arr=[...so.advances];arr[i]={...arr[i],tillDate:""};updateSO({...so,advances:arr});}} style={{fontSize:10,padding:"2px 8px",borderRadius:4,border:"1px solid #aaa",background:"#fff",color:"#888",cursor:"pointer"}}>↩ Use Bill Date</button>
                                  <span style={{fontSize:10,color:"#888"}}>
                                    ({Math.max(0,Math.round((parseLocalDate(a.tillDate)-parseLocalDate(a.date))/86400000))} days)
                                  </span>
                                </>
                              ) : (
                                <button onClick={()=>{const arr=[...so.advances];arr[i]={...arr[i],tillDate:a.date};updateSO({...so,advances:arr});}} style={{fontSize:10,padding:"2px 8px",borderRadius:4,border:"1px dashed #e67e22",background:"#fff9f0",color:"#e67e22",cursor:"pointer"}}>
                                  + Set custom interest end date
                                </button>
                              )}
                            </div>
                              {/* Compound toggle for sub-org advances */}
                              {(()=>{
                                const advDays = Math.round((parseLocalDate(a.tillDate||BILL_DATE)-parseLocalDate(a.date))/86400000);
                                if (advDays < 365) return null;
                                return (
                                  <button onClick={()=>{const arr=[...so.advances];arr[i]={...arr[i],compound:!a.compound};updateSO({...so,advances:arr});}}
                                    style={{fontSize:10,padding:"2px 10px",borderRadius:4,border:a.compound?"1px solid #6a0dad":"1px dashed #9b59b6",background:a.compound?"#f3e5ff":"#fdf5ff",color:a.compound?"#6a0dad":"#9b59b6",cursor:"pointer",fontWeight:a.compound?700:400}}>
                                    {a.compound?"✓ Compound Yearly":"⟳ Compound Yearly?"}
                                  </button>
                                );
                              })()}
                          </div>
                        ))}
                        <button onClick={()=>updateSO({...so,advances:[...(so.advances||[]),{date:new Date().toISOString().split("T")[0],amount:0,interestRate:24,note:""}]})} style={{background:"#e8f0ff",color:"#2d5a8a",border:"1px dashed #2d5a8a",borderRadius:4,padding:"4px 12px",cursor:"pointer",fontSize:12}}>+ Add Advance</button>
                      </div>


                      {/* ── FOUNDATION SEEDS ── */}
                      <div style={{marginTop:16,paddingTop:12,borderTop:"2px solid #2d6a2d"}}>
                        <div style={{fontWeight:700,fontSize:13,color:"#2d6a2d",marginBottom:10}}>
                          🌱 Foundation Seeds &nbsp;
                          <span style={{fontWeight:400,fontSize:11,color:"#888"}}>(₹1000 per acre deducted)</span>
                          <button
                            onClick={()=>{const fs=[...(so.foundationSeeds||[]),{variety:"",area:""}];updateSO({...so,foundationSeeds:fs});}}
                            style={{float:"right",background:"#e8f5e9",color:"#2d6a2d",border:"1px solid #2d6a2d",borderRadius:4,padding:"3px 12px",fontSize:12,cursor:"pointer",fontWeight:700}}>
                            + Add Row
                          </button>
                        </div>
                        {(so.foundationSeeds||[]).length===0 ? (
                          <div style={{color:"#aaa",fontSize:12,fontStyle:"italic",padding:"6px 0"}}>No foundation seeds yet — click + Add Row to add</div>
                        ) : (
                          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                            <thead>
                              <tr style={{background:"#e8f5e9"}}>
                                <th style={{padding:"5px 8px",textAlign:"left",border:"1px solid #c8e6c9"}}>Variety Name</th>
                                <th style={{padding:"5px 8px",textAlign:"center",border:"1px solid #c8e6c9",width:100}}>Area (Ac)</th>
                                <th style={{padding:"5px 8px",textAlign:"center",border:"1px solid #c8e6c9",width:80}}>Cost ₹</th>
                                <th style={{padding:"5px 8px",textAlign:"center",border:"1px solid #c8e6c9",width:40}}></th>
                              </tr>
                            </thead>
                            <tbody>
                              {(so.foundationSeeds||[]).map((fs,fi)=>(
                                <tr key={fi} style={{background:fi%2===0?"#fff":"#f9fdf9"}}>
                                  <td style={{padding:"4px 6px",border:"1px solid #c8e6c9"}}>
                                    <input value={fs.variety||""} onChange={e=>{const a=[...(so.foundationSeeds||[])];a[fi]={...a[fi],variety:e.target.value};updateSO({...so,foundationSeeds:a});}}
                                      placeholder="e.g. Bio-7511"
                                      style={{width:"100%",padding:"4px 6px",border:"1px solid #b0c8b0",borderRadius:4,fontSize:13}} />
                                  </td>
                                  <td style={{padding:"4px 6px",border:"1px solid #c8e6c9"}}>
                                    <input type="number" step="0.5" value={fs.area||""} onChange={e=>{const a=[...(so.foundationSeeds||[])];a[fi]={...a[fi],area:e.target.value};updateSO({...so,foundationSeeds:a});}}
                                      placeholder="0"
                                      style={{width:"100%",padding:"4px 6px",border:"1px solid #b0c8b0",borderRadius:4,fontSize:13,textAlign:"center"}} />
                                  </td>
                                  <td style={{padding:"4px 6px",border:"1px solid #c8e6c9",textAlign:"center",color:"#c0392b",fontWeight:600}}>
                                    ₹{((parseFloat(fs.area)||0)*1000).toLocaleString("en-IN")}
                                  </td>
                                  <td style={{padding:"4px 6px",border:"1px solid #c8e6c9",textAlign:"center"}}>
                                    <button onClick={()=>{const a=(so.foundationSeeds||[]).filter((_,k)=>k!==fi);updateSO({...so,foundationSeeds:a});}}
                                      style={{background:"#fdecea",color:"#e74c3c",border:"1px solid #e74c3c",borderRadius:4,padding:"2px 8px",cursor:"pointer",fontSize:12}}>✕</button>
                                  </td>
                                </tr>
                              ))}
                              <tr style={{background:"#e8f5e9",fontWeight:700}}>
                                <td style={{padding:"5px 8px",border:"1px solid #c8e6c9"}}>TOTAL</td>
                                <td style={{padding:"5px 8px",border:"1px solid #c8e6c9",textAlign:"center"}}>{(so.foundationSeeds||[]).reduce((s,f)=>s+(parseFloat(f.area)||0),0)} Ac</td>
                                <td style={{padding:"5px 8px",border:"1px solid #c8e6c9",textAlign:"center",color:"#c0392b"}}>₹{((so.foundationSeeds||[]).reduce((s,f)=>s+(parseFloat(f.area)||0),0)*1000).toLocaleString("en-IN")}</td>
                                <td style={{border:"1px solid #c8e6c9"}}></td>
                              </tr>
                            </tbody>
                          </table>
                        )}
                      </div>

                      {/* Jamma section */}
                      <div style={{marginTop:14,borderTop:"1px dashed #856404",paddingTop:10}}>
                        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
                          <span style={{fontWeight:700,fontSize:13,color:"#856404"}}>జమ్మా / Jamma (Sub-Org pays us)</span>
                          <button onClick={()=>updateSO({...so,jammaEntries:[...(so.jammaEntries||[]),{date:new Date().toISOString().split("T")[0],amount:0,interestRate:0,note:""}]})} style={{background:"rgba(133,100,4,0.15)",color:"#856404",border:"1px dashed #856404",borderRadius:4,padding:"3px 10px",cursor:"pointer",fontSize:11}}>+ Add Jamma</button>
                        </div>
                        {(so.jammaEntries||[]).map((j,ji)=>(
                          <div key={ji} style={{marginBottom:8,background:"#fff",border:"1px solid #ffeeba",borderRadius:5,padding:"8px 10px"}}>
                            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr auto",gap:6,alignItems:"end",marginBottom:5}}>
                              <div><label style={{fontSize:10,color:"#856404",display:"block"}}>Date</label><input type="date" value={j.date||""} onChange={e=>{const je=[...(so.jammaEntries||[])];je[ji]={...je[ji],date:e.target.value};updateSO({...so,jammaEntries:je});}} style={{width:"100%",padding:"4px 6px",border:"1px solid #856404",borderRadius:4,fontSize:12}} /></div>
                              <div><label style={{fontSize:10,color:"#856404",display:"block"}}>Amount ₹</label><input type="number" value={j.amount||""} onChange={e=>{const je=[...(so.jammaEntries||[])];je[ji]={...je[ji],amount:parseFloat(e.target.value)||0};updateSO({...so,jammaEntries:je});}} style={{width:"100%",padding:"4px 6px",border:"1px solid #856404",borderRadius:4,fontSize:12}} /></div>
                              <div><label style={{fontSize:10,color:"#856404",display:"block"}}>Interest %</label><input type="number" step="0.5" value={j.interestRate||""} onChange={e=>{const je=[...(so.jammaEntries||[])];je[ji]={...je[ji],interestRate:parseFloat(e.target.value)||0};updateSO({...so,jammaEntries:je});}} style={{width:"100%",padding:"4px 6px",border:"1px solid #856404",borderRadius:4,fontSize:12}} /></div>
                              <div><label style={{fontSize:10,color:"#856404",display:"block"}}>Note</label><input value={j.note||""} onChange={e=>{const je=[...(so.jammaEntries||[])];je[ji]={...je[ji],note:e.target.value};updateSO({...so,jammaEntries:je});}} style={{width:"100%",padding:"4px 6px",border:"1px solid #856404",borderRadius:4,fontSize:12}} /></div>
                              <button onClick={()=>updateSO({...so,jammaEntries:(so.jammaEntries||[]).filter((_,k)=>k!==ji)})} style={{background:"#c0392b",color:"#fff",border:"none",borderRadius:4,padding:"4px 8px",cursor:"pointer",fontSize:11}}>✕</button>
                            </div>
                            {/* Till Date override */}
                            <div style={{display:"flex",alignItems:"center",gap:8}}>
                              {j.tillDate ? (
                                <>
                                  <label style={{fontSize:10,color:"#e67e22",fontWeight:600}}>⚠ Interest Till:</label>
                                  <input type="date" value={j.tillDate} onChange={e=>{const je=[...(so.jammaEntries||[])];je[ji]={...je[ji],tillDate:e.target.value};updateSO({...so,jammaEntries:je});}}
                                    style={{padding:"2px 6px",border:"1px solid #e67e22",borderRadius:4,fontSize:11,color:"#e67e22"}} />
                                  <button onClick={()=>{const je=[...(so.jammaEntries||[])];je[ji]={...je[ji],tillDate:""};updateSO({...so,jammaEntries:je});}} style={{fontSize:10,padding:"2px 8px",borderRadius:4,border:"1px solid #aaa",background:"#fff",color:"#888",cursor:"pointer"}}>↩ Use Bill Date</button>
                                  <span style={{fontSize:10,color:"#888"}}>({Math.max(0,Math.round((parseLocalDate(j.tillDate)-parseLocalDate(j.date||BILL_DATE))/86400000))} days)</span>
                                </>
                              ) : (
                                <button onClick={()=>{const je=[...(so.jammaEntries||[])];je[ji]={...je[ji],tillDate:j.date||BILL_DATE};updateSO({...so,jammaEntries:je});}} style={{fontSize:10,padding:"2px 8px",borderRadius:4,border:"1px dashed #e67e22",background:"#fff9f0",color:"#e67e22",cursor:"pointer"}}>
                                  + Set custom interest end date
                                </button>
                              )}
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {subOrgTab==="growers"&&(
                    <div>
                      <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10 }}>
                        <div style={{ fontWeight:700,color:"#1a2a4a",fontSize:14 }}>Growers for {so.name||"Sub-Organizer"} ({growers.length})</div>
                        <button onClick={()=>updateSO({...so,growers:[...growers,{sNo:growers.length+1,lotNo:"",name:"",fatherName:"",village:so.village||"",variety:"",area:"",packets:0,result:"Pass",type:"KMS",rate:550}]})} style={{background:"#2d5a8a",color:"#fff",border:"none",borderRadius:5,padding:"6px 14px",cursor:"pointer",fontSize:12}}>+ Add Grower</button>
                      </div>
                      <div style={{ overflowX:"auto" }}>
                        <table style={{ width:"100%",borderCollapse:"collapse",fontSize:11,minWidth:720 }}>
                          <thead>
                            <tr style={{ background:"#2d5a8a",color:"#fff" }}>
                              {["S.No","LOT No","Grower","Father","Village","Variety","Packets","Result","Type","Rate(₹)","Amount(₹)","Note",""].map(h=><th key={h} style={th2}>{h}</th>)}
                            </tr>
                          </thead>
                          <tbody>
                            {growers.map((g,i)=>{
                              const ip=g.result==="Pass";
                              const amt=ip?(parseFloat(g.packets)||0)*(parseFloat(g.rate)||0):0;
                              return (
                                <tr key={i} style={{ background:!ip?"#fdecea":i%2===0?"#f5faff":"#fff",borderBottom:"1px solid #d0e4f4" }}>
                                  <td style={td2}><input value={g.sNo||""} onChange={e=>updGrower(i,"sNo",e.target.value)} style={{width:36,padding:"2px 4px",border:"1px solid #b0c8e0",borderRadius:3,fontSize:11,textAlign:"center"}} /></td>
                                  <td style={td2}><input value={g.lotNo||""} onChange={e=>updGrower(i,"lotNo",e.target.value)} style={{width:60,padding:"2px 4px",border:"1px solid #b0c8e0",borderRadius:3,fontSize:11}} /></td>
                                  <td style={td2}><input value={g.name||""} onChange={e=>updGrower(i,"name",e.target.value)} style={{width:110,padding:"2px 4px",border:"1px solid #b0c8e0",borderRadius:3,fontSize:11}} /></td>
                                  <td style={td2}><input value={g.fatherName||""} onChange={e=>updGrower(i,"fatherName",e.target.value)} style={{width:110,padding:"2px 4px",border:"1px solid #b0c8e0",borderRadius:3,fontSize:11}} /></td>
                                  <td style={td2}><input value={g.village||""} onChange={e=>updGrower(i,"village",e.target.value)} style={{width:90,padding:"2px 4px",border:"1px solid #b0c8e0",borderRadius:3,fontSize:11}} /></td>
                                  <td style={td2}><input value={g.variety||""} onChange={e=>updGrower(i,"variety",e.target.value)} style={{width:80,padding:"2px 4px",border:"1px solid #b0c8e0",borderRadius:3,fontSize:11}} /></td>
                                  <td style={td2}><input type="number" value={g.packets||0} onChange={e=>updGrower(i,"packets",e.target.value)} style={{width:55,padding:"2px 4px",border:"1px solid #b0c8e0",borderRadius:3,fontSize:11,textAlign:"center"}} /></td>
                                  <td style={td2}>
                                    <button onClick={()=>updGrower(i,"result","Pass")} style={{padding:"2px 7px",borderRadius:8,border:"none",fontWeight:700,fontSize:10,cursor:"pointer",background:ip?"#155724":"#e8f5e9",color:ip?"#fff":"#155724"}}>✓P</button>
                                    <button onClick={()=>updGrower(i,"result","Fail")} style={{padding:"2px 7px",borderRadius:8,border:"none",fontWeight:700,fontSize:10,cursor:"pointer",marginLeft:2,background:!ip?"#721c24":"#fdecea",color:!ip?"#fff":"#721c24"}}>✗F</button>
                                  </td>
                                  <td style={td2}>
                                    <button onClick={()=>{const gg=[...growers];gg[i]={...gg[i],type:"KMS",rate:550};updateSO({...so,growers:gg});}} style={{padding:"2px 6px",borderRadius:6,border:"none",fontWeight:700,fontSize:10,cursor:"pointer",background:(g.type||"KMS")==="KMS"?"#1a3a8a":"#e8f0ff",color:(g.type||"KMS")==="KMS"?"#fff":"#1a3a8a"}}>KMS</button>
                                    <button onClick={()=>{const gg=[...growers];gg[i]={...gg[i],type:"GMS",rate:400};updateSO({...so,growers:gg});}} style={{padding:"2px 6px",borderRadius:6,border:"none",fontWeight:700,fontSize:10,cursor:"pointer",marginLeft:2,background:g.type==="GMS"?"#856404":"#fff3cd",color:g.type==="GMS"?"#fff":"#856404"}}>GMS</button>
                                  </td>
                                  <td style={td2}>{ip?<input type="number" value={g.rate||0} onChange={e=>updGrower(i,"rate",e.target.value)} style={{width:55,padding:"2px 4px",border:"1px solid #b0c8e0",borderRadius:3,fontSize:11,textAlign:"center"}} />:<span style={{color:"#aaa",fontStyle:"italic",fontSize:10}}>—</span>}</td>
                                  <td style={{...td2,fontWeight:600,color:ip?"#1a6a1a":"#aaa"}}>{ip?`₹${amt.toLocaleString("en-IN")}`:"—"}</td>
                                  <td style={td2}><input value={g.note||""} onChange={e=>updGrower(i,"note",e.target.value)} placeholder="optional" style={{width:90,padding:"2px 4px",border:"1px dashed #c8a000",borderRadius:3,fontSize:10,background:"#fffdf5",color:"#856404"}} /></td>
                                  <td style={td2}><button onClick={()=>updateSO({...so,growers:growers.filter((_,j)=>j!==i)})} style={{background:"#fdecea",color:"#e74c3c",border:"1px solid #e74c3c",borderRadius:3,padding:"2px 6px",cursor:"pointer",fontSize:10}}>✕</button></td>
                                </tr>
                              );
                            })}
                          </tbody>
                          {growers.length>0&&(
                            <tfoot>
                              <tr style={{ background:"#e8f0ff",fontWeight:700 }}>
                                <td colSpan={7} style={{...td2,textAlign:"left",color:"#1a2a4a"}}>TOTALS</td>
                                <td style={td2}>{growers.reduce((s,g)=>s+(parseFloat(g.packets)||0),0)}</td>
                                <td style={td2}><span style={{color:"#155724"}}>{growers.filter(g=>g.result==="Pass").length}P</span>/<span style={{color:"#721c24"}}>{growers.filter(g=>g.result!=="Pass").length}F</span></td>
                                <td colSpan={2}></td>
                                <td style={{...td2,color:"#1a6a1a"}}>₹{growers.filter(g=>g.result==="Pass").reduce((s,g)=>s+(parseFloat(g.packets)||0)*(parseFloat(g.rate)||0),0).toLocaleString("en-IN")}</td>
                                <td></td>
                                <td></td>
                              </tr>
                            </tfoot>
                          )}
                        </table>
                      </div>
                    </div>
                  )}

                  {subOrgTab==="bill"&&(()=>{
                    const paidVars = allSubOrgVarieties.filter(v=>isSubOrgVarietyPaid(v));
                    const pendingVars = allSubOrgVarieties.filter(v=>!isSubOrgVarietyPaid(v));
                    const [billMode, setBillMode] = [so._billMode||"partial", (m)=>updateSO({...so,_billMode:m})];
                    const [selVars, setSelVars] = [so._selVars||[], (v)=>updateSO({...so,_selVars:v})];
                    // Per-sub-org settled varieties (stored on the sub-org object)
                    const settledVars = so._settledVars || [];
                    const setSettledVars = (sv) => updateSO({...so, _settledVars: sv});
                    const toPayVars = paidVars.filter(v => !settledVars.includes(v));
                    return (
                      <div>
                        {/* ── Bill Type Toggle ── */}
                        <div style={{background:"#f0f5ff",border:"1px solid #b0c8e0",borderRadius:8,padding:"12px 16px",marginBottom:14}}>
                          <div style={{fontWeight:700,fontSize:13,color:"#2d5a8a",marginBottom:10}}>Select Bill Type</div>
                          <div style={{display:"flex",gap:10,marginBottom:12}}>
                            <button onClick={()=>setBillMode("partial")} style={{padding:"8px 20px",borderRadius:8,border:"none",fontWeight:700,fontSize:13,cursor:"pointer",background:billMode==="partial"?"#2d5a8a":"#e8f0ff",color:billMode==="partial"?"#fff":"#2d5a8a"}}>
                              🧾 Partial Bill
                            </button>
                            <button onClick={()=>setBillMode("final")} style={{padding:"8px 20px",borderRadius:8,border:"none",fontWeight:700,fontSize:13,cursor:"pointer",background:billMode==="final"?"#1a4a1a":"#e8f5e9",color:billMode==="final"?"#fff":"#1a4a1a"}}>
                              📋 Final Bill
                            </button>
                          </div>

                          {billMode==="partial" && (
                            <div>
                              <div style={{fontSize:12,color:"#555",marginBottom:8}}>
                                <strong>Partial Bill</strong> — Shows only selected variety growers. No advances or settlement. Just a payment receipt. Record the amount you pay as Jamma later.
                              </div>
                              <div style={{fontSize:12,color:"#2d5a8a",fontWeight:600,marginBottom:8}}>Select varieties to include:</div>
                              <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:10}}>
                                {paidVars.map(v => {
                                  const checked = selVars.includes(v);
                                  const cnt = (so.growers||[]).filter(g=>g.variety===v&&g.result==="Pass").length;
                                  const billDate = getSubOrgVarietyBillDate(v);
                                  return (
                                    <div key={v} onClick={()=>setSelVars(checked?selVars.filter(x=>x!==v):[...selVars,v])}
                                      style={{display:"flex",alignItems:"center",gap:8,padding:"7px 12px",borderRadius:8,cursor:"pointer",border:`2px solid ${checked?"#2d5a8a":"#c8dfc8"}`,background:checked?"#e8f0ff":"#fff"}}>
                                      <div style={{width:16,height:16,borderRadius:3,border:`2px solid ${checked?"#2d5a8a":"#aaa"}`,background:checked?"#2d5a8a":"#fff",display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontSize:12,fontWeight:700,flexShrink:0}}>{checked?"✓":""}</div>
                                      <div>
                                        <div style={{fontWeight:700,fontSize:12,color:"#1a4a1a"}}>{v}</div>
                                        <div style={{fontSize:10,color:"#666"}}>📅 {billDate} · {cnt} growers</div>
                                      </div>
                                    </div>
                                  );
                                })}
                              </div>
                              <div style={{display:"flex",gap:8,alignItems:"center"}}>
                                <button onClick={()=>setSelVars(paidVars)} style={{background:"#2d5a8a",color:"#fff",border:"none",borderRadius:5,padding:"5px 12px",fontSize:12,cursor:"pointer"}}>☑ Select All</button>
                                <button onClick={()=>setSelVars([])} style={{background:"#fff",color:"#555",border:"1px solid #c8dfc8",borderRadius:5,padding:"5px 12px",fontSize:12,cursor:"pointer"}}>☐ Clear</button>
                                <span style={{fontSize:12,color:"#666"}}>{selVars.length>0?`${(so.growers||[]).filter(g=>selVars.includes(g.variety)&&g.result==="Pass").length} growers selected`:"Select varieties above"}</span>
                              </div>
                            </div>
                          )}
                          {billMode==="final" && (
                            <div>
                              <div style={{fontSize:12,color:"#555",marginBottom:10}}>
                                <strong>Final Bill</strong> — Shows all growers in 3 sections. Mark which varieties you have already paid to this sub-org as <strong>Settled</strong>.
                              </div>
                              {/* Settled varieties selector — per sub-org */}
                              {paidVars.length > 0 && (
                                <div style={{marginBottom:8}}>
                                  <div style={{fontSize:12,fontWeight:700,color:"#1a4a1a",marginBottom:6}}>
                                    Mark varieties already paid to this sub-org as Settled:
                                  </div>
                                  <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:8}}>
                                    {paidVars.map(v => {
                                      const isSettled = settledVars.includes(v);
                                      const amt = (so.growers||[]).filter(g=>g.variety===v&&g.result==="Pass").reduce((s,g)=>s+(parseFloat(g.packets)||0)*(getSubOrgVarietyRate(v)||parseFloat(g.rate)||0),0);
                                      const note = (so._varietyNotes||{})[v]||"";
                                      const settledDate = (so._settledDates||{})[v]||"";
                                      return (
                                        <div key={v} style={{display:"flex",flexDirection:"column",gap:4,padding:"8px 12px",borderRadius:8,
                                          border:`2px solid ${isSettled?"#2d6a2d":"#b0c8e0"}`,
                                          background:isSettled?"#e8f5e9":"#f0f5ff", minWidth:200}}>
                                          {/* Checkbox + variety name */}
                                          <div style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer"}} onClick={()=>setSettledVars(isSettled?settledVars.filter(x=>x!==v):[...settledVars,v])}>
                                            <div style={{width:18,height:18,borderRadius:3,border:`2px solid ${isSettled?"#2d6a2d":"#aaa"}`,background:isSettled?"#2d6a2d":"#fff",display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontSize:12,fontWeight:700,flexShrink:0}}>{isSettled?"✔":""}</div>
                                            <div>
                                              <div style={{fontWeight:700,fontSize:12,color:isSettled?"#2d6a2d":"#2d5a8a"}}>{v}</div>
                                              <div style={{fontSize:10,color:isSettled?"#2d6a2d":"#888"}}>
                                                {isSettled?"✔ Settled":"💰 To Pay"} · ₹{Math.round(amt).toLocaleString("en-IN")}
                                              </div>
                                            </div>
                                          </div>
                                          {/* Settled On date — only when settled */}
                                          {isSettled && (
                                            <div style={{display:"flex",alignItems:"center",gap:6}}>
                                              <label style={{fontSize:10,color:"#2d6a2d",fontWeight:600,whiteSpace:"nowrap"}}>📅 Settled On:</label>
                                              <input type="date" value={settledDate}
                                                onClick={e=>e.stopPropagation()}
                                                onChange={e=>{const d={...(so._settledDates||{}),[v]:e.target.value};updateSO({...so,_settledDates:d});}}
                                                style={{fontSize:10,padding:"3px 6px",border:"1px solid #2d6a2d",borderRadius:4,background:"#fff",color:"#1a4a1a",flex:1}}
                                              />
                                            </div>
                                          )}
                                          {/* Notes */}
                                          <input
                                            placeholder="Notes (optional)"
                                            value={note}
                                            onClick={e=>e.stopPropagation()}
                                            onChange={e=>{const n={...(so._varietyNotes||{}),[v]:e.target.value};updateSO({...so,_varietyNotes:n});}}
                                            style={{fontSize:10,padding:"3px 6px",border:"1px solid #c8dfc8",borderRadius:4,background:"#fff",color:"#333",width:"100%"}}
                                          />
                                        </div>
                                      );
                                    })}
                                  </div>
                                  <div style={{display:"flex",gap:8}}>
                                    <button onClick={()=>setSettledVars(paidVars)} style={{background:"#2d6a2d",color:"#fff",border:"none",borderRadius:5,padding:"5px 12px",fontSize:12,cursor:"pointer"}}>✔ All Settled</button>
                                    <button onClick={()=>setSettledVars([])} style={{background:"#fff",color:"#555",border:"1px solid #ccc",borderRadius:5,padding:"5px 12px",fontSize:12,cursor:"pointer"}}>☐ All To Pay</button>
                                  </div>
                                  {pendingVars.length > 0 && (
                                    <div style={{marginTop:8,fontSize:11,color:"#856404"}}>
                                      ⏳ Pending (company hasn't paid yet): {pendingVars.join(", ")}
                                    </div>
                                  )}
                                </div>
                              )}
                            </div>
                          )}
                        </div>

                        {/* ── Bill Date + Print Button ── */}
                        <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:12,flexWrap:"wrap"}}>
                          <div style={{display:"flex",alignItems:"center",gap:8,background:"#f0f5ff",border:"1px solid #b0c8e0",borderRadius:8,padding:"8px 14px"}}>
                            <label style={{fontSize:12,fontWeight:700,color:"#2d5a8a"}}>📅 Bill Date:</label>
                            <input type="date"
                              value={so._billDate||BILL_DATE}
                              onChange={e=>updateSO({...so,_billDate:e.target.value})}
                              style={{padding:"4px 8px",border:"1px solid #2d5a8a",borderRadius:5,fontSize:13,fontWeight:700,color:"#1a4a1a"}}
                            />
                            <span style={{fontSize:11,color:"#888"}}>Used for interest calculation</span>
                          </div>
                          <button onClick={()=>{
                            recordBillingHistory(so.accNo||so.name, billMode==="partial"?selVars:["FINAL BILL"]);
                            printBill('suborg-bill',`bill_${billMode==="partial"?"partial":"final"}_suborg_${so.accNo||so.name||'suborg'}`);
                          }} style={{background:billMode==="final"?"#1a4a1a":"#2d5a8a",color:"#fff",border:"none",borderRadius:5,padding:"8px 20px",cursor:"pointer",fontSize:13,fontWeight:700}}>
                            🖨 Print {billMode==="partial"?"Partial":"Final"} Bill
                          </button>
                        </div>

                        {/* ── Bill Preview ── */}
                        <div id="suborg-bill">
                          <SubOrgBill
                            so={so}
                            billMode={billMode}
                            selectedVarieties={selVars}
                            settledVarieties={settledVars}
                            isSubOrgVarietyPaid={isSubOrgVarietyPaid}
                            isSubOrgVarietySettled={isSubOrgVarietySettled}
                            isSubOrgVarietyToPay={isSubOrgVarietyToPay}
                            getSubOrgVarietyBillDate={getSubOrgVarietyBillDate}
                            getSubOrgVarietyRate={getSubOrgVarietyRate}
                            getSubOrgVarietyType={getSubOrgVarietyType}
                            soVarieties={allSubOrgVarieties}
                          />
                        </div>
                      </div>
                    );
                  })()}
                </>
              );
            })()}
          </div>
        </div>
      )}

      {/* ── DASHBOARD MODE ── */}
      {mode==="careof"&&(()=>{
        const allF = farmers||[];
        const careOfNames = [...new Set(allF.map(f=>(f.careOf||"").trim()).filter(Boolean))].sort();

        if (careOfNames.length === 0) {
          return (
            <div style={{ padding:40, textAlign:"center", color:"#888" }}>
              <div style={{ fontSize:40, marginBottom:10 }}>🤝</div>
              <div style={{ fontSize:16, fontWeight:700, marginBottom:6 }}>No C/o groups yet</div>
              <div style={{ fontSize:13 }}>Add a "C/o" name to any farmer in the Farmers tab to see them grouped here.</div>
            </div>
          );
        }

        const printCareOfSummary = (name, farmersList) => {
          const rows = farmersList.map(f => {
            const bal = getFarmerBalance(f);
            return `<tr>
              <td style="padding:6px 10px;border:1px solid #ccc;">${f.farmerNo||""}</td>
              <td style="padding:6px 10px;border:1px solid #ccc;">${f.name||""}</td>
              <td style="padding:6px 10px;border:1px solid #ccc;">${f.village||""}</td>
              <td style="padding:6px 10px;border:1px solid #ccc;text-align:right;font-weight:700;color:${bal>=0?'#1a7a1a':'#b30000'}">₹${Math.abs(Math.round(bal)).toLocaleString('en-IN')} ${bal>=0?'(Pay)':'(Due)'}</td>
            </tr>`;
          }).join("");
          const total = farmersList.reduce((s,f)=>s+getFarmerBalance(f),0);
          const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"/><title>C/o Summary - ${name}</title>
            <style>
              body{font-family:Georgia,serif;padding:20px;color:#222;}
              h2{color:#1a4a1a;margin-bottom:2px;}
              table{border-collapse:collapse;width:100%;margin-top:14px;}
              th{background:#1a4a1a;color:#fff;padding:8px 10px;text-align:left;border:1px solid #ccc;}
              .total-row td{font-weight:800;font-size:15px;border-top:2px solid #1a4a1a;background:#f0f7f0;}
              @media print{@page{margin:10mm;size:A4 portrait;}}
            </style></head><body>
            <h2>C/o Settlement Summary</h2>
            <div style="font-size:15px;margin-bottom:4px;"><strong>C/o:</strong> ${name}</div>
            <div style="font-size:13px;color:#555;">Bill Date: ${fmtDate(BILL_DATE)} | Farmers: ${farmersList.length}</div>
            <table>
              <thead><tr><th>Farmer No</th><th>Name</th><th>Village</th><th style="text-align:right;">Balance</th></tr></thead>
              <tbody>${rows}
                <tr class="total-row"><td colspan="3">TOTAL</td><td style="text-align:right;padding:8px 10px;border:1px solid #ccc;">₹${Math.abs(Math.round(total)).toLocaleString('en-IN')} ${total>=0?'(Pay to C/o)':'(Due from C/o)'}</td></tr>
              </tbody>
            </table>
            </body></html>`;
          const w = window.open("", "_blank");
          w.document.write(html);
          w.document.close();
          setTimeout(()=>w.print(), 400);
        };

        return (
          <div style={{ display:"flex", height:"calc(100vh - 64px)" }}>
            <div style={{ width:240, minWidth:200, background:"#1a4a1a", color:"#fff", overflowY:"auto" }} className="mobile-panel">
              <div style={{ padding:"8px 12px", fontSize:11, fontWeight:700, letterSpacing:1, color:"#8bc88b", borderBottom:"1px solid #2d5a2d" }}>
                C/o PERSONS ({careOfNames.length})
              </div>
              {careOfNames.map(name => {
                const count = allF.filter(f=>(f.careOf||"").trim()===name).length;
                const isSel = selectedCareOf === name;
                return (
                  <div key={name} onClick={()=>setSelectedCareOf(name)}
                    style={{ padding:"10px 12px", cursor:"pointer", background:isSel?"#2d6a2d":"transparent", borderLeft:isSel?"3px solid #7dd87d":"3px solid transparent", borderBottom:"1px solid #1a3a1a" }}>
                    <div style={{ fontSize:13, fontWeight:700 }}>🤝 {name}</div>
                    <div style={{ fontSize:11, color:"#8bc88b" }}>{count} farmer{count!==1?"s":""}</div>
                  </div>
                );
              })}
            </div>
            <div style={{ flex:1, overflowY:"auto", padding:"16px 20px" }}>
              {!selectedCareOf ? (
                <div style={{ padding:40, textAlign:"center", color:"#888" }}>
                  <div style={{ fontSize:34, marginBottom:8 }}>👈</div>
                  <div>Select a C/o person from the list to see their farmers and combined total.</div>
                </div>
              ) : (() => {
                const farmersList = allF.filter(f=>(f.careOf||"").trim()===selectedCareOf);
                const farmersWithBal = farmersList.map(f=>({f, balance:getFarmerBalance(f)}));
                const total = farmersWithBal.reduce((s,x)=>s+x.balance,0);
                return (
                  <div>
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:10, marginBottom:16 }}>
                      <div>
                        <div style={{ fontSize:20, fontWeight:800, color:"#1a4a1a" }}>🤝 {selectedCareOf}</div>
                        <div style={{ fontSize:13, color:"#666" }}>{farmersList.length} farmer{farmersList.length!==1?"s":""} across {[...new Set(farmersList.map(f=>f.village).filter(Boolean))].length} village(s)</div>
                      </div>
                      <button onClick={()=>printCareOfSummary(selectedCareOf, farmersList)}
                        style={{ background:"#1a4a1a", color:"#fff", border:"none", borderRadius:6, padding:"10px 18px", fontWeight:700, fontSize:13, cursor:"pointer" }}>
                        🖨️ Print Summary
                      </button>
                    </div>

                    <div style={{ background:total>=0?"#e8f5e9":"#fdecea", border:`2px solid ${total>=0?"#2d6a2d":"#e74c3c"}`, borderRadius:8, padding:"14px 18px", marginBottom:18, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                      <span style={{ fontWeight:700, fontSize:14, color:total>=0?"#1a4a1a":"#b30000" }}>Combined Total</span>
                      <span style={{ fontWeight:800, fontSize:22, color:total>=0?"#1a7a1a":"#b30000" }}>
                        ₹{Math.abs(Math.round(total)).toLocaleString('en-IN')} {total>=0?"(Pay to C/o)":"(Due from C/o)"}
                      </span>
                    </div>

                    <div style={{ background:"#fff", border:"1px solid #c8dfc8", borderRadius:8, overflow:"hidden" }}>
                      <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                        <thead>
                          <tr style={{ background:"#f0f7f0" }}>
                            <th style={{ padding:"8px 10px", textAlign:"left", borderBottom:"1px solid #c8dfc8" }}>Farmer No</th>
                            <th style={{ padding:"8px 10px", textAlign:"left", borderBottom:"1px solid #c8dfc8" }}>Name</th>
                            <th style={{ padding:"8px 10px", textAlign:"left", borderBottom:"1px solid #c8dfc8" }}>Village</th>
                            <th style={{ padding:"8px 10px", textAlign:"right", borderBottom:"1px solid #c8dfc8" }}>Balance</th>
                            <th style={{ padding:"8px 10px", textAlign:"center", borderBottom:"1px solid #c8dfc8" }}>Status</th>
                          </tr>
                        </thead>
                        <tbody>
                          {farmersWithBal.map(({f,balance}) => (
                            <tr key={f.id} style={{ cursor:"pointer" }} onClick={()=>{ const idx=farmers.findIndex(x=>x.id===f.id); setSelectedIdx(idx); setMode("farmers"); setTab("preview"); }}>
                              <td style={{ padding:"8px 10px", borderBottom:"1px solid #eee" }}>{f.farmerNo||"—"}</td>
                              <td style={{ padding:"8px 10px", borderBottom:"1px solid #eee", fontWeight:600 }}>{f.name||"—"}</td>
                              <td style={{ padding:"8px 10px", borderBottom:"1px solid #eee" }}>{f.village||"—"}</td>
                              <td style={{ padding:"8px 10px", borderBottom:"1px solid #eee", textAlign:"right", fontWeight:700, color:balance>=0?"#1a7a1a":"#b30000" }}>
                                ₹{Math.abs(Math.round(balance)).toLocaleString('en-IN')} {balance>=0?"(Pay)":"(Due)"}
                              </td>
                              <td style={{ padding:"8px 10px", borderBottom:"1px solid #eee", textAlign:"center" }}>
                                {f.billingDone ? <span style={{ color:"#1a7a1a" }}>✔ Billed</span> : <span style={{ color:"#b35c00" }}>⏳ Pending</span>}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                );
              })()}
            </div>
          </div>
        );
      })()}

      {mode==="dashboard"&&(()=>{
        const allF=farmers||[];
        const fStats=allF.map(f=>{
          const advWI=(f.advances||[]).reduce((s,a)=>{const{interest}=calcInterest(a.amount,a.interestRate,a.date);return s+a.amount+interest;},0);
          const cCalc=(f.crops||[]).map(c=>{const area=parseFloat(c.area)||0,qty=parseFloat(c.quantity)||0;return{value:c.result==="Pass"?qty*(parseFloat(c.ratePerUnit)||0):0,foundation:area*1000,transport:qty};});
          const cropVal=cCalc.reduce((s,c)=>s+c.value,0);
          const found=cCalc.reduce((s,c)=>s+c.foundation,0);
          const trans=cCalc.reduce((s,c)=>s+c.transport,0);
          const jamWI=(f.jammaEnabled?(f.jammaEntries||[]):[]).reduce((s,j)=>{const{interest}=calcInterest(parseFloat(j.amount)||0,parseFloat(j.interestRate)||0,j.date||BILL_DATE);return s+(parseFloat(j.amount)||0)+interest;},0);
          const bal=cropVal-advWI+jamWI-found-trans;
          return{...f,balance:bal,cropVal,advWI};
        });
        const totalPayable=fStats.filter(f=>f.balance>0).reduce((s,f)=>s+f.balance,0);
        const totalDue=fStats.filter(f=>f.balance<0).reduce((s,f)=>s+Math.abs(f.balance),0);
        const totalAdvances=fStats.reduce((s,f)=>s+f.advWI,0);
        const totalCropVal=fStats.reduce((s,f)=>s+f.cropVal,0);
        const allCrops=allF.flatMap(f=>f.crops||[]);
        const passC=allCrops.filter(c=>c.result==="Pass").length;
        const failC=allCrops.filter(c=>c.result==="Fail").length;
        const vMap={};
        fStats.forEach(f=>{const v=f.village?.trim()||"No Village";if(!vMap[v])vMap[v]={farmers:0,payable:0,due:0,cropVal:0};vMap[v].farmers++;if(f.balance>=0)vMap[v].payable+=f.balance;else vMap[v].due+=Math.abs(f.balance);vMap[v].cropVal+=f.cropVal;});
        const villages=Object.entries(vMap).sort((a,b)=>b[1].cropVal-a[1].cropVal);
        const maxVC=Math.max(...villages.map(([,v])=>v.cropVal),1);
        const soStats=subOrgs.map(so=>{
          const advWI=(so.advances||[]).reduce((s,a)=>{const{interest}=calcInterest(parseFloat(a.amount)||0,parseFloat(a.interestRate)||0,a.date);return s+(parseFloat(a.amount)||0)+interest;},0);
          const g=so.growers||[];
          const seedAmt=g.filter(gr=>gr.result==="Pass").reduce((s,gr)=>s+(parseFloat(gr.packets)||0)*(parseFloat(gr.rate)||0),0);
          const found=g.reduce((s,gr)=>s+(parseFloat(gr.area)||0)*1000,0);
          const trans=g.reduce((s,gr)=>s+(parseFloat(gr.packets)||0),0);
          return{...so,balance:seedAmt-advWI-found-trans,seedAmt,growerCount:g.length,passCount:g.filter(gr=>gr.result==="Pass").length};
        });
        const fmt=n=>`₹${Math.abs(n).toLocaleString("en-IN")}`;
        const Card=({icon,label,value,sub,color})=>(
          <div style={{background:"#fff",borderRadius:10,padding:"14px 18px",border:`2px solid ${color}20`,boxShadow:"0 2px 8px rgba(0,0,0,0.06)",flex:"1 1 160px",minWidth:150}}>
            <div style={{fontSize:22,marginBottom:4}}>{icon}</div>
            <div style={{fontSize:11,color:"#777",marginBottom:2}}>{label}</div>
            <div style={{fontSize:20,fontWeight:800,color}}>{value}</div>
            {sub&&<div style={{fontSize:10,color:"#999",marginTop:2}}>{sub}</div>}
          </div>
        );
        const Bar=({data,color,maxV,vf})=>(
          <div style={{display:"flex",flexDirection:"column",gap:6}}>
            {data.map(([k,v],i)=>(
              <div key={i}>
                <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:2}}>
                  <span style={{color:"#444",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:"60%"}}>{k}</span>
                  <span style={{fontWeight:600,color}}>{vf(v)}</span>
                </div>
                <div style={{background:"#f0f0f0",borderRadius:4,height:10,overflow:"hidden"}}>
                  <div style={{width:`${(v/maxV)*100}%`,height:"100%",background:color,borderRadius:4}} />
                </div>
              </div>
            ))}
          </div>
        );
        const Donut=({segs,size=120})=>{
          const total=segs.reduce((s,g)=>s+g.value,0);
          if(!total) return <div style={{width:size,height:size,borderRadius:"50%",background:"#eee",margin:"0 auto"}} />;
          let cum=0;
          const r=40,cx=60,cy=60,sw=20,circ=2*Math.PI*r;
          return (
            <svg width={size} height={size} viewBox="0 0 120 120">
              {segs.map((seg,i)=>{
                const pct=seg.value/total,dash=pct*circ,off=circ-(cum/total)*circ;
                cum+=seg.value;
                return <circle key={i} cx={cx} cy={cy} r={r} fill="none" stroke={seg.color} strokeWidth={sw} strokeDasharray={`${dash} ${circ-dash}`} strokeDashoffset={off} style={{transform:"rotate(-90deg)",transformOrigin:"60px 60px"}} />;
              })}
              <text x={cx} y={cy-6} textAnchor="middle" style={{fontSize:11,fill:"#333",fontWeight:700}}>{total}</text>
              <text x={cx} y={cy+8} textAnchor="middle" style={{fontSize:9,fill:"#777"}}>Total</text>
            </svg>
          );
        };
        return (
          <div style={{padding:"16px 20px",overflowY:"auto",height:"calc(100vh - 58px)",background:"#f0f7f0"}}>
            <div style={{fontWeight:800,fontSize:18,color:"#1a4a1a",marginBottom:4}}>📊 Dashboard Overview</div>
            <div style={{fontSize:12,color:"#666",marginBottom:16}}>Bill Date: 01 July 2026 · {allF.length} farmers · {subOrgs.length} sub-organizers</div>
            <div style={{display:"flex",flexWrap:"wrap",gap:12,marginBottom:20}}>
              <Card icon="💰" label="Total Payable to Farmers" value={fmt(totalPayable)} sub={`${fStats.filter(f=>f.balance>0).length} farmers`} color="#1a6a1a" />
              <Card icon="📥" label="Total Due from Farmers" value={fmt(totalDue)} sub={`${fStats.filter(f=>f.balance<0).length} farmers`} color="#c0392b" />
              <Card icon="🌾" label="Total Crop Value" value={fmt(totalCropVal)} sub={`${passC} crops passed`} color="#856404" />
              <Card icon="💸" label="Total Advances+Interest" value={fmt(totalAdvances)} sub={`Across ${allF.length} farmers`} color="#2d5a8a" />
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(300px,1fr))",gap:16}}>
              <div style={{background:"#fff",borderRadius:10,padding:16,boxShadow:"0 2px 8px rgba(0,0,0,0.06)"}}>
                <div style={{fontWeight:700,fontSize:13,color:"#1a4a1a",marginBottom:12}}>💰 Balance Pending — Top Farmers</div>
                {fStats.filter(f=>f.balance>0).length===0?<div style={{color:"#aaa",fontSize:12,textAlign:"center",padding:20}}>No pending balances</div>:
                  <Bar data={fStats.filter(f=>f.balance>0).sort((a,b)=>b.balance-a.balance).slice(0,8).map(f=>[`#${f.farmerNo||"?"} ${f.name}`,f.balance])} color="#1a6a1a" maxV={Math.max(...fStats.filter(f=>f.balance>0).map(f=>f.balance),1)} vf={fmt} />
                }
              </div>
              <div style={{background:"#fff",borderRadius:10,padding:16,boxShadow:"0 2px 8px rgba(0,0,0,0.06)"}}>
                <div style={{fontWeight:700,fontSize:13,color:"#1a4a1a",marginBottom:12}}>📍 Village-wise Crop Value</div>
                {villages.length===0?<div style={{color:"#aaa",fontSize:12,textAlign:"center",padding:20}}>No village data</div>:
                  <Bar data={villages.map(([v,d])=>[v,d.cropVal])} color="#856404" maxV={maxVC} vf={fmt} />
                }
                {villages.length>0&&<div style={{marginTop:12,display:"flex",flexWrap:"wrap",gap:6}}>{villages.map(([v,d])=><div key={v} style={{background:"#f0f7f0",borderRadius:6,padding:"4px 10px",fontSize:11}}><strong>{v}</strong>: {d.farmers} farmers · <span style={{color:"#1a6a1a"}}>Pay {fmt(d.payable)}</span>{d.due>0&&<span style={{color:"#c0392b"}}> · Due {fmt(d.due)}</span>}</div>)}</div>}
              </div>
              <div style={{background:"#fff",borderRadius:10,padding:16,boxShadow:"0 2px 8px rgba(0,0,0,0.06)"}}>
                <div style={{fontWeight:700,fontSize:13,color:"#1a4a1a",marginBottom:12}}>✅ Crop Pass / Fail Summary</div>
                <div style={{display:"flex",alignItems:"center",gap:20}}>
                  <Donut segs={[{value:passC,color:"#2d6a2d"},{value:failC,color:"#e74c3c"}]} size={120} />
                  <div style={{flex:1}}>
                    {[["✓ Passed",passC,"#1a4a1a","#2d6a2d"],["✗ Failed",failC,"#c0392b","#e74c3c"]].map(([l,n,tc,bc])=>(
                      <div key={l} style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                        <div style={{width:14,height:14,borderRadius:3,background:bc,flexShrink:0}} />
                        <div><div style={{fontSize:12,fontWeight:600,color:tc}}>{l}: {n}</div><div style={{fontSize:11,color:"#777"}}>{allCrops.length>0?Math.round(n/allCrops.length*100):0}% of crops</div></div>
                      </div>
                    ))}
                    <div style={{marginTop:8,padding:"6px 10px",background:"#f0f7f0",borderRadius:6,fontSize:11}}>
                      <div>KMS: {allCrops.filter(c=>(c.cropType||"KMS")==="KMS").length}</div>
                      <div>GMS: {allCrops.filter(c=>c.cropType==="GMS").length}</div>
                    </div>
                  </div>
                </div>
              </div>
              <div style={{background:"#fff",borderRadius:10,padding:16,boxShadow:"0 2px 8px rgba(0,0,0,0.06)"}}>
                <div style={{fontWeight:700,fontSize:13,color:"#1a2a4a",marginBottom:12}}>🏢 Sub-Organizer Summary</div>
                {soStats.length===0?<div style={{color:"#aaa",fontSize:12,textAlign:"center",padding:20}}>No sub-organizers added yet</div>:(
                  <>
                    <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:12}}>
                      {[["#",soStats.length,"Sub-Orgs","#2d5a8a","#e8f0ff"],["👥",soStats.reduce((s,so)=>s+so.growerCount,0),"Growers","#1a6a1a","#e8f5e9"],["🌾",fmt(soStats.reduce((s,so)=>s+so.seedAmt,0)),"Seed Amt","#856404","#fff3cd"]].map(([ic,v,l,tc,bg])=>(
                        <div key={l} style={{background:bg,borderRadius:8,padding:"8px 14px",fontSize:12,flex:1,textAlign:"center"}}><div style={{fontWeight:700,fontSize:15,color:tc}}>{ic} {v}</div><div style={{color:"#555"}}>{l}</div></div>
                      ))}
                    </div>
                    <div style={{display:"flex",flexDirection:"column",gap:6}}>
                      {soStats.map((so,i)=>(
                        <div key={i} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 12px",background:"#f5faff",borderRadius:8,border:"1px solid #d0e4f4"}}>
                          <div><div style={{fontWeight:600,fontSize:12}}>#{so.accNo||"—"} {so.name||"Sub-Org"}</div><div style={{fontSize:10,color:"#7ab8e8"}}>📍 {so.village||"—"} · {so.growerCount} growers ({so.passCount} passed)</div></div>
                          <div style={{textAlign:"right"}}><div style={{fontWeight:700,fontSize:13,color:so.balance>=0?"#1a6a1a":"#c0392b"}}>{so.balance>=0?"Pay ":"Due "}{fmt(so.balance)}</div><div style={{fontSize:10,color:"#999"}}>Seed: {fmt(so.seedAmt)}</div></div>
                        </div>
                      ))}
                    </div>
                  </>
                )}
              </div>
            </div>
          </div>
        );
      })()}

      {/* ── VARIETY PAY MODE ── */}
      {mode === "variety" && (
        <div style={{padding:"20px 28px",overflowY:"auto",height:"calc(100vh - 58px)",background:"#f0f7f0"}}>
          <div style={{fontWeight:800,fontSize:18,color:"#1a4a1a",marginBottom:4}}>🌾 Variety Payment Settings</div>
          <div style={{fontSize:12,color:"#666",marginBottom:16}}>All varieties from Farmers and Sub-Orgs in one place. Edit Status, Type, Rate and Bill Date here — changes apply to both farmer bills and sub-org bills instantly.</div>

          {allVarieties.length === 0 ? (
            <div style={{background:"#fff",borderRadius:10,padding:40,textAlign:"center",color:"#aaa"}}>No crop varieties found.</div>
          ) : (() => {
            const fmt = n => `₹${Math.abs(Math.round(n)).toLocaleString("en-IN")}`;

            // Per-variety stats — loop through every farmer for this variety
            const varStats = allVarieties.map(variety => {
              const paid = isVarietyPaid(variety);
              const billDate = getVarietyBillDate(variety);
              const farmerList = (farmers||[]).filter(f => (f.crops||[]).some(c => c.variety === variety));

              let totalCropVal = 0, totalAdvWI = 0, totalJamWI = 0, totalFound = 0, totalTrans = 0, totalQty = 0;
              let balPayTotal = 0, balDueTotal = 0;
              let payFarmers = 0, dueFarmers = 0;

              farmerList.forEach(f => {
                // Crop value for this farmer — ONLY for this variety
                // Use rate from Variety Settings if set, otherwise fall back to crop's own rate
                const settingsRate = getVarietyRate(variety);
                let fCropVal = 0, fFound = 0, fTrans = 0, fQty = 0;
                (f.crops||[]).forEach(c => {
                  if (c.variety !== variety) return;
                  const qty = parseFloat(c.quantity)||0;
                  const rateToUse = (c.rateOverride===true) ? (parseFloat(c.ratePerUnit)||0) : (settingsRate || parseFloat(c.ratePerUnit)||0);
                  if (c.result === "Pass") { fCropVal += qty * rateToUse; fQty += qty; }
                  fFound += (parseFloat(c.area)||0) * 1000;
                  fTrans += qty;
                });

                // Full advance+interest for this farmer using variety bill date
                const fAdvWI = (f.advances||[]).reduce((s,a) => {
                  const {interest} = (a.compound?calcCompoundInterest:calcInterest)(a.amount, a.interestRate, a.date, billDate);
                  return s + a.amount + interest;
                }, 0);
                const fJamWI = ((f.jammaEnabled?f.jammaEntries||[]:[]).reduce((s,j) => {
                  const {interest} = calcInterest(parseFloat(j.amount)||0, parseFloat(j.interestRate)||0, j.date||BILL_DATE, billDate);
                  return s + (parseFloat(j.amount)||0) + interest;
                }, 0));

                // Split advance proportionally across this farmer's passed crop varieties
                const fPassVarieties = Math.max((f.crops||[]).filter(c=>c.result==="Pass").length, 1);
                const fAdvShare = fAdvWI / fPassVarieties;
                const fJamShare = fJamWI / fPassVarieties;

                const fBalance = fCropVal - fAdvShare + fJamShare - fFound - fTrans;

                totalCropVal += fCropVal;
                totalAdvWI   += fAdvShare;
                totalJamWI   += fJamShare;
                totalFound   += fFound;
                totalTrans   += fTrans;
                totalQty     += fQty;

                if (fBalance >= 0) { balPayTotal += fBalance; payFarmers++; }
                else               { balDueTotal += Math.abs(fBalance); dueFarmers++; }
              });

              // Detect default type/rate from first farmer, allow override from settings
              let detectedType = "KMS", detectedRate = 0;
              for (const f of (farmers||[])) {
                const c = (f.crops||[]).find(c=>c.variety===variety);
                if (c) { detectedType = c.cropType||"KMS"; detectedRate = parseFloat(c.ratePerUnit)||0; break; }
              }
              const selectedType = varietySettings[variety]?.type || detectedType;
              const globalRate = (varietySettings[variety]?.rate) ? parseFloat(varietySettings[variety].rate) : detectedRate;

              // Sub-org stats for this variety — uses SAME status/type/rate as farmer
              const soList = (subOrgs||[]).filter(so=>(so.growers||[]).some(g=>g.variety===variety));
              const soGrowerCount = (subOrgs||[]).reduce((s,so)=>(so.growers||[]).filter(g=>g.variety===variety&&g.result==="Pass").length+s,0);
              const soRate = globalRate || detectedRate;
              const soQty = (subOrgs||[]).reduce((s,so)=>(so.growers||[]).filter(g=>g.variety===variety&&g.result==="Pass").reduce((ss,g)=>ss+(parseFloat(g.packets)||0),0)+s,0);
              const soCropVal = soQty * soRate;

              return { variety, paid, billDate, farmerCount: farmerList.length,
                totalCropVal, totalAdvWI, totalFound, totalTrans, totalQty,
                balPayTotal, balDueTotal, payFarmers, dueFarmers,
                detectedType, selectedType, detectedRate, globalRate,
                soCount: soList.length, soGrowerCount, soCropVal, soQty };
            });

            const fmtN = n => `₹${Math.abs(Math.round(n)).toLocaleString("en-IN")}`;

            const printVarietyPaySummary = () => {
              const rows = varStats.map(v => `<tr>
                <td style="padding:6px 8px;border:1px solid #ccc;">${v.variety}</td>
                <td style="padding:6px 8px;border:1px solid #ccc;text-align:center;">${v.paid?'✅ Paid':'⏳ Pending'}</td>
                <td style="padding:6px 8px;border:1px solid #ccc;text-align:center;">${v.selectedType}</td>
                <td style="padding:6px 8px;border:1px solid #ccc;text-align:right;">₹${v.globalRate||v.detectedRate||0}</td>
                <td style="padding:6px 8px;border:1px solid #ccc;text-align:center;">${v.farmerCount}</td>
                <td style="padding:6px 8px;border:1px solid #ccc;text-align:right;">${v.totalQty.toLocaleString('en-IN')}</td>
                <td style="padding:6px 8px;border:1px solid #ccc;text-align:right;">${fmtN(v.totalCropVal)}</td>
                <td style="padding:6px 8px;border:1px solid #ccc;text-align:right;color:#1a7a1a;">${fmtN(v.balPayTotal)}</td>
                <td style="padding:6px 8px;border:1px solid #ccc;text-align:right;color:#b30000;">${fmtN(v.balDueTotal)}</td>
                <td style="padding:6px 8px;border:1px solid #ccc;text-align:center;">${fmtDate(v.billDate)}</td>
                <td style="padding:6px 8px;border:1px solid #ccc;text-align:center;">${v.soCount||0}</td>
                <td style="padding:6px 8px;border:1px solid #ccc;text-align:right;">${(v.soQty||0).toLocaleString('en-IN')}</td>
                <td style="padding:6px 8px;border:1px solid #ccc;text-align:right;">${fmtN(v.soCropVal||0)}</td>
              </tr>`).join("");

              const grandQty = varStats.reduce((s,v)=>s+v.totalQty+(v.soQty||0),0);
              const grandPay = varStats.reduce((s,v)=>s+v.balPayTotal,0);
              const grandDue = varStats.reduce((s,v)=>s+v.balDueTotal,0);

              const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"/><title>Variety Payment Summary</title>
                <style>
                  body{font-family:Georgia,serif;padding:20px;color:#222;}
                  h2{color:#1a4a1a;margin-bottom:2px;}
                  table{border-collapse:collapse;width:100%;margin-top:14px;font-size:12px;}
                  th{background:#1a4a1a;color:#fff;padding:7px 8px;text-align:left;border:1px solid #ccc;font-size:11px;}
                  .total-row td{font-weight:800;border-top:2px solid #1a4a1a;background:#f0f7f0;}
                  @media print{@page{margin:8mm;size:A4 landscape;}}
                </style></head><body>
                <h2>🌾 Variety Payment Summary</h2>
                <div style="font-size:13px;color:#555;margin-bottom:6px;">Generated: ${fmtDate(new Date().toISOString().split("T")[0])} | Total Varieties: ${varStats.length}</div>
                <table>
                  <thead><tr>
                    <th>Variety</th><th>Status</th><th>Type</th><th>Rate ₹</th>
                    <th>Farmers</th><th>Qty</th><th>Crop Value</th><th>Bal Pay</th><th>Bal Due</th><th>Bill Date</th>
                    <th>Sub-Orgs</th><th>SO Qty</th><th>SO Value</th>
                  </tr></thead>
                  <tbody>${rows}
                    <tr class="total-row">
                      <td colspan="5">GRAND TOTAL</td>
                      <td style="text-align:right;padding:7px 8px;border:1px solid #ccc;">${grandQty.toLocaleString('en-IN')}</td>
                      <td></td>
                      <td style="text-align:right;padding:7px 8px;border:1px solid #ccc;color:#1a7a1a;">${fmtN(grandPay)}</td>
                      <td style="text-align:right;padding:7px 8px;border:1px solid #ccc;color:#b30000;">${fmtN(grandDue)}</td>
                      <td colspan="4"></td>
                    </tr>
                  </tbody>
                </table>
                </body></html>`;
              const w = window.open("", "_blank");
              w.document.write(html);
              w.document.close();
              setTimeout(()=>w.print(), 400);
            };


            // Summary totals
            const paid = varStats.filter(v=>v.paid);
            const pend = varStats.filter(v=>!v.paid);
            const paidPay = paid.reduce((s,v)=>s+v.balPayTotal,0);
            const paidDue = paid.reduce((s,v)=>s+v.balDueTotal,0);
            const pendPay = pend.reduce((s,v)=>s+v.balPayTotal,0);
            const pendDue = pend.reduce((s,v)=>s+v.balDueTotal,0);

            const col = "180px 70px 70px 80px 70px 90px 100px 100px 100px 100px 100px 70px 80px 90px 110px 100px";
            const th = (align="center") => ({textAlign:align,fontSize:11,fontWeight:700,padding:"10px 8px",display:"flex",alignItems:"center",justifyContent:align==="left"?"flex-start":"center",minHeight:36,boxSizing:"border-box"});
            const td = (color="#333",bold=false,align="center") => ({textAlign:align,fontSize:12,padding:"10px 8px",fontWeight:bold?700:400,color,display:"flex",alignItems:"center",justifyContent:align==="left"?"flex-start":"center",boxSizing:"border-box"});

            return (
              <div style={{display:"flex",flexDirection:"column",gap:14}}>

                {/* ── Print Button ── */}
                <div style={{display:"flex",justifyContent:"flex-end"}}>
                  <button onClick={printVarietyPaySummary}
                    style={{background:"#1a4a1a",color:"#fff",border:"none",borderRadius:6,padding:"9px 16px",fontWeight:700,fontSize:13,cursor:"pointer"}}>
                    🖨️ Print Summary
                  </button>
                </div>

                {/* ── Summary Cards ── */}
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:10}}>
                  <div style={{background:"#e8f5e9",border:"2px solid #2d6a2d",borderRadius:10,padding:"14px 16px"}}>
                    <div style={{fontSize:11,color:"#2d6a2d",fontWeight:700,marginBottom:6}}>✅ PAID — Balance to Pay Farmers</div>
                    <div style={{fontSize:24,fontWeight:800,color:"#1a4a1a"}}>{fmt(paidPay)}</div>
                    <div style={{fontSize:10,color:"#555",marginTop:4}}>Cash you need to arrange RIGHT NOW</div>
                  </div>
                  <div style={{background:"#fdecea",border:"2px solid #e74c3c",borderRadius:10,padding:"14px 16px"}}>
                    <div style={{fontSize:11,color:"#c0392b",fontWeight:700,marginBottom:6}}>✅ PAID — Balance Due from Farmers</div>
                    <div style={{fontSize:24,fontWeight:800,color:"#c0392b"}}>{fmt(paidDue)}</div>
                    <div style={{fontSize:10,color:"#555",marginTop:4}}>Farmers owe you — collect later</div>
                  </div>
                  <div style={{background:"#fff3cd",border:"2px solid #856404",borderRadius:10,padding:"14px 16px"}}>
                    <div style={{fontSize:11,color:"#856404",fontWeight:700,marginBottom:6}}>⏳ PENDING — Balance to Pay Farmers</div>
                    <div style={{fontSize:24,fontWeight:800,color:"#856404"}}>{fmt(pendPay)}</div>
                    <div style={{fontSize:10,color:"#555",marginTop:4}}>Cash needed WHEN company pays you</div>
                  </div>
                  <div style={{background:"#fef9ec",border:"2px solid #c8a000",borderRadius:10,padding:"14px 16px"}}>
                    <div style={{fontSize:11,color:"#856404",fontWeight:700,marginBottom:6}}>⏳ PENDING — Balance Due from Farmers</div>
                    <div style={{fontSize:24,fontWeight:800,color:"#856404"}}>{fmt(pendDue)}</div>
                    <div style={{fontSize:10,color:"#555",marginTop:4}}>Pending farmers who owe you — collect later</div>
                  </div>
                </div>

                {/* ── Controls ── */}
                <div style={{display:"flex",gap:8,alignItems:"center"}}>
                  <button onClick={()=>{const vs={...varietySettings};allVarieties.forEach(v=>{vs[v]={...(vs[v]||{}),status:"paid",billDate:BILL_DATE};vs["so_"+v]={...(vs["so_"+v]||{}),status:"paid",billDate:BILL_DATE};});saveVarietySettings(vs);}} style={{background:"#2d6a2d",color:"#fff",border:"none",borderRadius:6,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>✅ Mark All Paid</button>
                  <button onClick={()=>{const vs={...varietySettings};allVarieties.forEach(v=>{vs[v]={...(vs[v]||{}),status:"pending",billDate:""};vs["so_"+v]={...(vs["so_"+v]||{}),status:"pending",billDate:""};});saveVarietySettings(vs);}} style={{background:"#856404",color:"#fff",border:"none",borderRadius:6,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>⏳ Mark All Pending</button>
                  <span style={{fontSize:12,color:"#666"}}>✅ {paid.length} paid · ⏳ {pend.length} pending</span>
                </div>

                {/* ── Table ── */}
                <div style={{background:"#fff",borderRadius:10,overflowX:"auto",overflowY:"hidden",boxShadow:"0 2px 8px rgba(0,0,0,0.07)"}}>
                  <div style={{minWidth:1470}}>
                  {/* Header */}
                  <div style={{display:"grid",gridTemplateColumns:col,background:"#1a4a1a",color:"#fff",gap:0}}>
                    <div style={th("left")}>Variety</div>
                    <div style={th()}>Status</div>
                    <div style={th()}>Type</div>
                    <div style={th()}>Rate(₹)</div>
                    {/* Farmer columns */}
                    <div style={{...th(),background:"#2d5a8a",borderLeft:"2px solid #fff"}}>Farmers</div>
                    <div style={{...th(),background:"#2d5a8a"}}>Quantity</div>
                    <div style={{...th(),background:"#2d5a8a"}}>Crop Value</div>
                    <div style={{...th(),background:"#2d5a8a"}}>Adv+Int</div>
                    <div style={{...th(),background:"#2d5a8a"}}>Bal to Pay</div>
                    <div style={{...th(),background:"#2d5a8a"}}>Bal Due</div>
                    <div style={{...th(),background:"#2d5a8a"}}>Bill Date</div>
                    {/* Sub-Org columns */}
                    <div style={{...th(),background:"#856404",borderLeft:"2px solid #fff"}}>Sub-Orgs</div>
                    <div style={{...th(),background:"#856404"}}>Growers</div>
                    <div style={{...th(),background:"#856404"}}>Quantity</div>
                    <div style={{...th(),background:"#856404"}}>Crop Value</div>
                    <div style={{...th(),background:"#856404"}}>Bill Date</div>
                    {/* Total */}
                    <div style={{...th(),background:"#1a4a1a",borderLeft:"2px solid #fff"}}>Total Qty</div>
                  </div>

                  {/* Data rows */}
                  {varStats.map((v,vi)=>{
                    const inFarmer = (farmers||[]).some(f=>(f.crops||[]).some(c=>c.variety===v.variety));
                    const inSubOrg = (subOrgs||[]).some(so=>(so.growers||[]).some(g=>g.variety===v.variety));
                    const vsKey = "so_"+v.variety;
                    const soVs = varietySettings[vsKey]||{};
                    const soStatus = soVs.status||"paid";
                    const soToPay = soStatus==="paid";
                    const soBillDate = soVs.billDate||BILL_DATE;
                    const soCount = (subOrgs||[]).filter(so=>(so.growers||[]).some(g=>g.variety===v.variety)).length;
                    const soGrowerCount = (subOrgs||[]).reduce((s,so)=>(so.growers||[]).filter(g=>g.variety===v.variety&&g.result==="Pass").length+s,0);
                    const soRate = parseFloat((varietySettings[v.variety]||{}).rate||0) || (() => { for(const so of (subOrgs||[])){const g=(so.growers||[]).find(g=>g.variety===v.variety&&g.result==="Pass");if(g)return parseFloat(g.rate)||0;} return 0; })();
                    const soCropVal = (subOrgs||[]).reduce((s,so)=>(so.growers||[]).filter(g=>g.variety===v.variety&&g.result==="Pass").reduce((ss,g)=>ss+(parseFloat(g.packets)||0)*soRate,0)+s,0);
                    const rowBg = vi%2===0?"#fff":"#f9fdf9";
                    return (
                    <div key={v.variety} style={{display:"grid",gridTemplateColumns:col,borderBottom:"1px solid #e8f5e9",background:rowBg,alignItems:"center",gap:0}}>
                      {/* Variety */}
                      <div style={{...td("#1a4a1a",true,"left"),flexDirection:"column",alignItems:"flex-start",gap:2}}>
                        <div>{v.variety}</div>
                        <div style={{display:"flex",gap:3}}>
                          {inFarmer && <span style={{fontSize:9,background:"#e8f0ff",color:"#2d5a8a",borderRadius:3,padding:"1px 4px"}}>👨‍🌾</span>}
                          {inSubOrg && <span style={{fontSize:9,background:"#e8f5e9",color:"#2d6a2d",borderRadius:3,padding:"1px 4px"}}>🏢</span>}
                        </div>
                      </div>
                      {/* Status — shared (uses farmer setting as source of truth, applies to both) */}
                      <div style={{padding:"6px 4px",textAlign:"center"}}>
                        <div style={{display:"flex",gap:2,justifyContent:"center"}}>
                          <button onClick={()=>{
                            const nv={...varietySettings};
                            nv[v.variety]={...(nv[v.variety]||{}),status:"paid",billDate:(nv[v.variety]?.billDate)||BILL_DATE};
                            nv[vsKey]={...soVs,status:"paid",billDate:soVs.billDate||BILL_DATE};
                            saveVarietySettings(nv);
                          }} style={{padding:"3px 6px",borderRadius:4,border:"none",fontWeight:700,fontSize:10,cursor:"pointer",background:v.paid?"#2d6a2d":"#e8f5e9",color:v.paid?"#fff":"#2d6a2d"}}>✅ Paid</button>
                          <button onClick={()=>{
                            const nv={...varietySettings};
                            nv[v.variety]={...(nv[v.variety]||{}),status:"pending",billDate:""};
                            nv[vsKey]={...soVs,status:"pending",billDate:""};
                            saveVarietySettings(nv);
                          }} style={{padding:"3px 6px",borderRadius:4,border:"none",fontWeight:700,fontSize:10,cursor:"pointer",background:!v.paid?"#856404":"#fff3cd",color:!v.paid?"#fff":"#856404"}}>⏳</button>
                        </div>
                      </div>
                      {/* Type — shared */}
                      <div style={{padding:"4px",textAlign:"center"}}>
                        <div style={{display:"flex",gap:2,justifyContent:"center"}}>
                          {["KMS","GMS"].map(t=>{
                            const active=(varietySettings[v.variety]?.type||v.detectedType)===t;
                            return <button key={t} onClick={()=>{const nv={...varietySettings,[v.variety]:{...(varietySettings[v.variety]||{}),type:t},["so_"+v.variety]:{...(varietySettings["so_"+v.variety]||{}),type:t}};saveVarietySettings(nv);}} style={{padding:"3px 6px",borderRadius:4,border:"none",fontWeight:700,fontSize:10,cursor:"pointer",background:active?(t==="GMS"?"#856404":"#1a3a8a"):"#f0f0f0",color:active?"#fff":"#aaa"}}>{t}</button>;
                          })}
                        </div>
                      </div>
                      {/* Rate — shared */}
                      <div style={{padding:"4px 6px",textAlign:"center"}}>
                        <input type="number" value={v.globalRate||""} onChange={e=>{const nv={...varietySettings,[v.variety]:{...(varietySettings[v.variety]||{}),rate:e.target.value},["so_"+v.variety]:{...(varietySettings["so_"+v.variety]||{}),rate:e.target.value}};saveVarietySettings(nv);}}
                          style={{width:"60px",padding:"4px 5px",borderRadius:5,border:"2px solid #2d6a2d",fontSize:12,fontWeight:700,color:"#1a4a1a",textAlign:"center"}} placeholder={String(v.detectedRate)} />
                      </div>
                      {/* ── FARMER COLUMNS ── */}
                      <div style={{borderLeft:"2px solid #2d5a8a",textAlign:"center",padding:"6px 4px",fontSize:12,color:"#555"}}>{inFarmer?v.farmerCount:"—"}</div>
                      <div style={{textAlign:"center",padding:"6px 4px",fontSize:12,fontWeight:600,color:"#555"}}>{inFarmer?v.totalQty.toLocaleString("en-IN"):"—"}</div>
                      <div style={{textAlign:"center",padding:"6px 4px",fontSize:12,fontWeight:600,color:"#1a6a1a"}}>{inFarmer?fmt(v.totalCropVal):"—"}</div>
                      <div style={{textAlign:"center",padding:"6px 4px",fontSize:12,color:"#c0392b"}}>{inFarmer?fmt(v.totalAdvWI):"—"}</div>
                      <div style={{textAlign:"center",padding:"6px 4px",fontSize:11,fontWeight:600,color:"#1a6a1a",background:v.balPayTotal>0?"#f0fff0":"transparent"}}>
                        {inFarmer?(v.balPayTotal>0?<div><div>{fmt(v.balPayTotal)}</div><div style={{fontSize:9,color:"#2d6a2d"}}>{v.payFarmers}F</div></div>:"—"):"—"}
                      </div>
                      <div style={{textAlign:"center",padding:"6px 4px",fontSize:11,fontWeight:600,color:"#c0392b",background:v.balDueTotal>0?"#fff5f5":"transparent"}}>
                        {inFarmer?(v.balDueTotal>0?<div><div>{fmt(v.balDueTotal)}</div><div style={{fontSize:9,color:"#c0392b"}}>{v.dueFarmers}F</div></div>:"—"):"—"}
                      </div>
                      <div style={{padding:"4px 6px",borderRight:"2px solid #856404"}}>
                        {inFarmer&&v.paid?<input type="date" value={(varietySettings[v.variety]?.billDate)||BILL_DATE} onChange={e=>{const nv={...varietySettings,[v.variety]:{...(varietySettings[v.variety]||{}),billDate:e.target.value}};saveVarietySettings(nv);}} style={{padding:"2px 4px",borderRadius:4,border:"1px solid #c8dfc8",fontSize:10,width:"100%"}}/>:<span style={{color:"#bbb",fontSize:10}}>—</span>}
                      </div>
                      {/* ── SUB-ORG COLUMNS ── */}
                      <div style={{borderLeft:"2px solid #856404",textAlign:"center",padding:"6px 4px",fontSize:12,color:"#555"}}>{inSubOrg?soCount:"—"}</div>
                      <div style={{textAlign:"center",padding:"6px 4px",fontSize:12,color:"#555"}}>{inSubOrg?soGrowerCount:"—"}</div>
                      <div style={{textAlign:"center",padding:"6px 4px",fontSize:12,fontWeight:600,color:"#555"}}>{inSubOrg?v.soQty.toLocaleString("en-IN"):"—"}</div>
                      <div style={{textAlign:"center",padding:"6px 4px",fontSize:12,fontWeight:600,color:v.paid?"#2d5a8a":"#aaa"}}>{inSubOrg?(v.paid?fmt(soCropVal):"⏳"):"—"}</div>
                      <div style={{padding:"4px 6px"}}>
                        {inSubOrg&&v.paid?<input type="date" value={soBillDate} onChange={e=>{const nv={...varietySettings,[vsKey]:{...soVs,billDate:e.target.value}};saveVarietySettings(nv);}} style={{padding:"2px 4px",borderRadius:4,border:"1px solid #b0c8e0",fontSize:10,width:"100%"}}/>:<span style={{color:"#bbb",fontSize:10}}>—</span>}
                      </div>
                      {/* Total Quantity — Farmers + Sub-Org */}
                      <div style={{borderLeft:"2px solid #1a4a1a",textAlign:"center",padding:"6px 4px",fontSize:13,fontWeight:800,color:"#1a4a1a",background:"#f0f5ff"}}>
                        {((inFarmer?v.totalQty:0)+(inSubOrg?v.soQty:0)).toLocaleString("en-IN")}
                      </div>
                    </div>
                    );
                  })}
                  {/* ── No Seed Farmers row ── */}
                  {(() => {
                    const noSeedFarmers = (farmers||[]).filter(f=>(f.advances||[]).length>0 && !(f.crops||[]).some(c=>c.variety&&c.variety.trim()));
                    if (noSeedFarmers.length === 0) return null;
                    const billDate = BILL_DATE;
                    let totalAdv=0, totalInt=0;
                    noSeedFarmers.forEach(f => {
                      (f.advances||[]).forEach(a => {
                        const amt=parseFloat(a.amount)||0;
                        const bd=a.tillDate||billDate;
                        const {interest}=calcInterest(amt,parseFloat(a.interestRate)||0,a.date,bd);
                        totalAdv+=amt; totalInt+=interest;
                      });
                    });
                    const totalDue = totalAdv+totalInt;
                    return (
                      <div style={{display:"grid",gridTemplateColumns:col,borderBottom:"2px solid #c0392b",background:"#fff5f5",alignItems:"center",gap:0,borderTop:"2px solid #c0392b"}}>
                        <div style={{...td("#c0392b",true,"left"),flexDirection:"column",alignItems:"flex-start",gap:2,padding:"10px 8px"}}>
                          <div>🚫 No Seed Farmers</div>
                          <div style={{fontSize:10,color:"#888",fontWeight:400}}>No crop this season</div>
                        </div>
                        <div style={{padding:"6px 4px",textAlign:"center"}}>
                          <span style={{fontSize:11,color:"#888"}}>—</span>
                        </div>
                        <div style={{textAlign:"center",padding:"4px",fontSize:11,color:"#aaa"}}>—</div>
                        <div style={{textAlign:"center",padding:"4px",fontSize:11,color:"#aaa"}}>—</div>
                        {/* Farmer count */}
                        <div style={{borderLeft:"2px solid #2d5a8a",textAlign:"center",padding:"6px 4px",fontSize:12,fontWeight:700,color:"#c0392b"}}>{noSeedFarmers.length}</div>
                        {/* Quantity — blank */}
                        <div style={{textAlign:"center",fontSize:11,color:"#aaa"}}>—</div>
                        {/* Crop Value — blank */}
                        <div style={{textAlign:"center",fontSize:11,color:"#aaa"}}>—</div>
                        {/* Adv+Int */}
                        <div style={{textAlign:"center",padding:"6px 4px",fontSize:12,fontWeight:700,color:"#c0392b"}}>₹{(totalAdv+totalInt).toLocaleString("en-IN")}</div>
                        {/* Bal to Pay — blank */}
                        <div style={{textAlign:"center",fontSize:11,color:"#aaa"}}>—</div>
                        {/* Bal Due = full advance+int (no crop value to offset) */}
                        <div style={{textAlign:"center",padding:"6px 4px",fontSize:12,fontWeight:700,color:"#c0392b",background:"#fdecea"}}>
                          <div>₹{totalDue.toLocaleString("en-IN")}</div>
                          <div style={{fontSize:9,color:"#c0392b"}}>{noSeedFarmers.length}F</div>
                        </div>
                        {/* Bill Date — blank */}
                        <div style={{textAlign:"center",fontSize:11,color:"#aaa"}}>—</div>
                        {/* Sub-Org columns — blank */}
                        <div style={{borderLeft:"2px solid #856404",textAlign:"center",fontSize:11,color:"#aaa"}}>—</div>
                        <div style={{textAlign:"center",fontSize:11,color:"#aaa"}}>—</div>
                        <div style={{textAlign:"center",fontSize:11,color:"#aaa"}}>—</div>
                        <div style={{textAlign:"center",fontSize:11,color:"#aaa"}}>—</div>
                        <div style={{textAlign:"center",fontSize:11,color:"#aaa"}}>—</div>
                        {/* Total Qty — blank */}
                        <div style={{borderLeft:"2px solid #1a4a1a",textAlign:"center",fontSize:11,color:"#aaa"}}>—</div>
                      </div>
                    );
                  })()}

                  {/* Totals row */}
                  <div style={{display:"grid",gridTemplateColumns:col,padding:"0",background:"#e8f5e9",fontWeight:700,borderTop:"2px solid #2d6a2d",alignItems:"center",gap:0}}>
                    <div style={td("#1a4a1a",true,"left")}>TOTAL</div>
                    <div></div><div></div><div></div>
                    <div style={{...td("#2d5a8a",true),borderLeft:"2px solid #2d5a8a"}}>{(farmers||[]).length}</div>
                    <div style={td("#555",true)}>{varStats.reduce((s,v)=>s+v.totalQty,0).toLocaleString("en-IN")}</div>
                    <div style={td("#1a6a1a",true)}>{fmt(varStats.reduce((s,v)=>s+v.totalCropVal,0))}</div>
                    <div style={td("#c0392b",true)}>{fmt(varStats.reduce((s,v)=>s+v.totalAdvWI,0))}</div>
                    <div style={{...td("#1a6a1a",true),background:"#d4edda"}}>{fmt(paidPay+pendPay)}</div>
                    <div style={{...td("#c0392b",true),background:"#fde8e8"}}>{fmt(paidDue+pendDue)}</div>
                    <div></div>
                    <div style={{...td("#856404",true),borderLeft:"2px solid #856404"}}>
                      {(subOrgs||[]).length}
                    </div>
                    <div style={td("#555",true)}>{(subOrgs||[]).reduce((s,so)=>(so.growers||[]).filter(g=>g.result==="Pass").length+s,0)}</div>
                    <div style={td("#555",true)}>{varStats.reduce((s,v)=>s+v.soQty,0).toLocaleString("en-IN")}</div>
                    <div style={td("#2d5a8a",true)}>{fmt(allSubOrgVarieties.reduce((s,v)=>{
                      const r=parseFloat((varietySettings[v]||{}).rate||0)||0;
                      return s+(subOrgs||[]).reduce((ss,so)=>(so.growers||[]).filter(g=>g.variety===v&&g.result==="Pass").reduce((sss,g)=>sss+(parseFloat(g.packets)||0)*r,0)+ss,0);
                    },0))}</div>
                    <div></div>
                    <div style={{...td("#1a4a1a",true),borderLeft:"2px solid #1a4a1a",background:"#e8f0ff",fontSize:13}}>
                      {(varStats.reduce((s,v)=>s+v.totalQty,0)+varStats.reduce((s,v)=>s+v.soQty,0)).toLocaleString("en-IN")}
                    </div>
                  </div>
                  </div>
                </div>

                <div style={{fontSize:11,color:"#888",padding:"8px 12px",background:"#fff",borderRadius:8,border:"1px solid #e8f5e9"}}>
                  💡 <strong>Balance to Pay</strong> = farmers whose crop value exceeds their advance — you pay them. <strong>Balance Due</strong> = farmers whose advance exceeds crop value — they pay you back later. Both are shown separately per variety.
                </div>
              </div>
            );
          })()}
        </div>
      )}

      {/* ── UNDO DELETE TOAST ── */}
      {undoItem && (
        <div style={{position:"fixed",bottom:24,left:"50%",transform:"translateX(-50%)",background:"#1a2a4a",color:"#fff",borderRadius:10,padding:"12px 20px",display:"flex",alignItems:"center",gap:14,boxShadow:"0 4px 20px rgba(0,0,0,0.4)",zIndex:9997,minWidth:320}}>
          <div style={{fontSize:13}}>
            🗑 {undoItem.type==="farmer"?"Farmer":"Sub-Org"} <strong>"{undoItem.data?.name}"</strong> deleted
          </div>
          <button onClick={undoDelete} style={{background:"#f0c040",color:"#1a2a4a",border:"none",borderRadius:6,padding:"6px 14px",fontWeight:800,cursor:"pointer",fontSize:13,whiteSpace:"nowrap"}}>
            ↩ Undo
          </button>
          <button onClick={()=>{clearTimeout(undoItem.timeout);setUndoItem(null);}} style={{background:"none",border:"none",color:"#aaa",cursor:"pointer",fontSize:16}}>✕</button>
        </div>
      )}

      {/* ── SETTINGS MODAL ── */}
      {showSettings && (
        <div style={{position:"fixed",top:0,left:0,width:"100%",height:"100%",background:"rgba(0,0,0,0.6)",zIndex:9998,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:"#fff",borderRadius:12,padding:24,width:520,maxWidth:"96vw",maxHeight:"90vh",overflowY:"auto",boxShadow:"0 8px 32px rgba(0,0,0,0.3)"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
              <div style={{fontWeight:800,fontSize:18,color:"#1a2a4a"}}>⚙️ Settings</div>
              <button onClick={()=>setShowSettings(false)} style={{background:"none",border:"none",fontSize:20,cursor:"pointer",color:"#888"}}>✕</button>
            </div>

            {/* Bill Date */}
            <div style={{marginBottom:20,background:"#f0f5ff",borderRadius:8,padding:"14px 16px"}}>
              <div style={{fontWeight:700,fontSize:13,color:"#2d5a8a",marginBottom:6}}>📅 Default Bill Date</div>
              <div style={{fontSize:12,color:"#666",marginBottom:8}}>Used for all interest calculations. Change this at the start of each billing season.</div>
              <input type="date" value={BILL_DATE} onChange={e=>setBILL_DATE(e.target.value)}
                style={{padding:"8px 12px",border:"2px solid #2d5a8a",borderRadius:6,fontSize:14,fontWeight:700,color:"#1a2a4a",width:"100%"}} />
              <div style={{fontSize:11,color:"#2d5a8a",marginTop:4}}>Current: {fmtDate(BILL_DATE)}</div>
            </div>

            {/* Backup status */}
            <div style={{marginBottom:20,background:"#f0fff0",borderRadius:8,padding:"14px 16px"}}>
              <div style={{fontWeight:700,fontSize:13,color:"#2d6a2d",marginBottom:6}}>💾 Data Backup</div>
              <div style={{fontSize:12,color:"#666",marginBottom:8}}>
                Last backup: <strong>{lastBackupDate ? fmtDate(lastBackupDate) : "Never"}</strong>
                {daysSinceBackup < 999 && ` (${daysSinceBackup} days ago)`}
              </div>
              <button onClick={()=>{ storage.exportToFile(farmers, subOrgs); markBackupDone(); }}
                style={{background:"#2d6a2d",color:"#fff",border:"none",borderRadius:6,padding:"8px 16px",fontWeight:700,cursor:"pointer",fontSize:13,width:"100%"}}>
                📊 Export Backup Now
              </button>
            </div>

            {/* Data stats */}
            <div style={{background:"#f5f5f5",borderRadius:8,padding:"14px 16px",marginBottom:16}}>
              <div style={{fontWeight:700,fontSize:13,color:"#555",marginBottom:8}}>📊 Data Summary</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,fontSize:13}}>
                <div>👨‍🌾 Farmers: <strong>{(farmers||[]).length}</strong></div>
                <div>🏢 Sub-Orgs: <strong>{(subOrgs||[]).length}</strong></div>
                <div>🌾 Varieties: <strong>{allVarieties.length}</strong></div>
                <div>📦 Total Growers: <strong>{(subOrgs||[]).reduce((s,so)=>s+(so.growers||[]).length,0)}</strong></div>
              </div>
            </div>

            {/* Daily Snapshots */}
            <div style={{marginBottom:16,background:"#f0f5ff",borderRadius:8,padding:"14px 16px"}}>
              <div style={{fontWeight:700,fontSize:13,color:"#2d5a8a",marginBottom:6}}>📸 Daily Snapshots (last 7 days)</div>
              <div style={{fontSize:12,color:"#666",marginBottom:8}}>Automatic daily backup. Restore if something goes wrong.</div>
              <button onClick={async()=>{
                const ok=await saveSnapshot(farmers||[],subOrgs||[]);
                if(ok) alert("✅ Snapshot saved to cloud!");
                else alert("❌ Failed to save snapshot");
              }} style={{background:"#2d5a8a",color:"#fff",border:"none",borderRadius:5,padding:"6px 14px",fontSize:12,fontWeight:700,cursor:"pointer",marginBottom:8}}>
                📸 Save Snapshot Now
              </button>
              <button onClick={async()=>{
                const snaps=await getSnapshots();
                setSnapshots(snaps);
                setShowRestore(true);
              }} style={{background:"#fff",color:"#2d5a8a",border:"1px solid #2d5a8a",borderRadius:5,padding:"6px 14px",fontSize:12,fontWeight:700,cursor:"pointer",marginLeft:8}}>
                🔄 View & Restore
              </button>
            </div>

            {/* Pesticide Master List */}
            <div style={{marginBottom:20,background:"#fff8f0",borderRadius:8,padding:"16px",border:"1.5px solid #f0a040"}}>
              <div style={{fontWeight:700,fontSize:15,color:"#b35c00",marginBottom:4}}>🧪 Pesticide Price List</div>
              <div style={{fontSize:12,color:"#888",marginBottom:12}}>Add each pesticide with its sizes and prices. Each size is a separate entry.</div>
              {pesticideList.map((p,i) => (
                <div key={i} style={{marginBottom:12,background:"#fff",border:"1.5px solid #f0a040",borderRadius:8,padding:"10px 12px"}}>
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                    <input
                      placeholder="Pesticide name (e.g. NuziMax 13-45)"
                      value={p.name}
                      onChange={e=>{ const l=[...pesticideList]; l[i]={...l[i],name:e.target.value}; savePesticideList(l); }}
                      style={{flex:1,padding:"8px 10px",border:"1.5px solid #f0a040",borderRadius:6,fontSize:14,boxSizing:"border-box"}}
                    />
                    <button onClick={()=>savePesticideList(pesticideList.filter((_,j)=>j!==i))}
                      style={{background:"#fdecea",color:"#e74c3c",border:"1.5px solid #e74c3c",borderRadius:6,padding:"6px 10px",cursor:"pointer",fontSize:13,fontWeight:700,whiteSpace:"nowrap"}}>✕ Remove</button>
                  </div>
                  <div style={{fontSize:11,fontWeight:700,color:"#b35c00",marginBottom:4}}>Sizes &amp; Prices:</div>
                  {(p.sizes||[]).map((s,si)=>(
                    <div key={si} style={{display:"grid",gridTemplateColumns:"1fr 1fr 36px",gap:6,marginBottom:6,alignItems:"center"}}>
                      <input placeholder="Size (e.g. 250ml)" value={s.size||""} onChange={e=>{ const l=[...pesticideList]; l[i].sizes[si]={...s,size:e.target.value}; savePesticideList(l); }}
                        style={{padding:"7px 10px",border:"1px solid #f0c080",borderRadius:5,fontSize:13}} />
                      <input placeholder="Price ₹" type="number" value={s.price||""} onChange={e=>{ const l=[...pesticideList]; l[i].sizes[si]={...s,price:parseFloat(e.target.value)||0}; savePesticideList(l); }}
                        style={{padding:"7px 10px",border:"1px solid #f0c080",borderRadius:5,fontSize:13}} />
                      <button onClick={()=>{ const l=[...pesticideList]; l[i].sizes=l[i].sizes.filter((_,k)=>k!==si); savePesticideList(l); }}
                        style={{background:"#fdecea",color:"#e74c3c",border:"1px solid #e74c3c",borderRadius:4,padding:"4px",cursor:"pointer",fontSize:12,width:36,height:34}}>✕</button>
                    </div>
                  ))}
                  <button onClick={()=>{ const l=[...pesticideList]; l[i].sizes=[...(l[i].sizes||[]),{size:"",price:0}]; savePesticideList(l); }}
                    style={{background:"#fff8f0",color:"#b35c00",border:"1px dashed #f0a040",borderRadius:4,padding:"4px 10px",fontSize:12,cursor:"pointer"}}>+ Add Size</button>
                </div>
              ))}
              {pesticideList.length < 15 && (
                <button onClick={()=>savePesticideList([...pesticideList,{name:"",sizes:[{size:"",price:0}]}])}
                  style={{background:"#fff3e0",color:"#b35c00",border:"2px dashed #f0a040",borderRadius:6,padding:"10px 16px",fontSize:14,fontWeight:700,cursor:"pointer",marginTop:4,width:"100%"}}>
                  ＋ Add Pesticide
                </button>
              )}
            </div>

            <button onClick={()=>setShowSettings(false)} style={{width:"100%",background:"#1a2a4a",color:"#fff",border:"none",borderRadius:6,padding:"10px",fontWeight:700,cursor:"pointer",fontSize:14}}>
              Done
            </button>
          </div>
        </div>
      )}

      {/* ── RESTORE SNAPSHOT MODAL ── */}
      {showRestore && (
        <div style={{position:"fixed",top:0,left:0,width:"100%",height:"100%",background:"rgba(0,0,0,0.7)",zIndex:9999,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:"#fff",borderRadius:12,padding:24,width:420,maxWidth:"95vw"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
              <div style={{fontWeight:800,fontSize:16,color:"#1a2a4a"}}>🔄 Restore from Snapshot</div>
              <button onClick={()=>setShowRestore(false)} style={{background:"none",border:"none",fontSize:20,cursor:"pointer"}}>✕</button>
            </div>
            {snapshots.length===0 ? (
              <div style={{color:"#888",fontSize:13,textAlign:"center",padding:20}}>No snapshots found. Save one first!</div>
            ) : snapshots.map(s=>(
              <div key={s.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 12px",borderRadius:8,border:"1px solid #e0e0e0",marginBottom:8}}>
                <div>
                  <div style={{fontWeight:700,fontSize:13}}>{s.date}</div>
                  <div style={{fontSize:11,color:"#666"}}>{s.farmerCount} farmers · {s.subOrgCount} sub-orgs · Saved {new Date(s.savedAt).toLocaleTimeString("en-IN")}</div>
                </div>
                <button onClick={async()=>{
                  if(!window.confirm(`Restore data from ${s.date}? Current data will be replaced.`)) return;
                  const data=await restoreSnapshot(s.id);
                  if(data){
                    updateFarmers(data.farmers);
                    updateSubOrgs(data.subOrgs);
                    setShowRestore(false);
                    setShowSettings(false);
                    alert("✅ Data restored from "+s.date);
                  } else alert("❌ Restore failed");
                }} style={{background:"#2d5a8a",color:"#fff",border:"none",borderRadius:6,padding:"6px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>
                  Restore
                </button>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ── PRINT QUEUE OVERLAY ── */}
      {printQueueIdx >= 0 && printQueue.length > 0 && (()=>{
        const qf = printQueue[printQueueIdx];
        const isLast = printQueueIdx >= printQueue.length - 1;

        const openInNewTab = (farmer) => {
          // Build the full bill HTML
          const billDiv = document.createElement("div");
          billDiv.id = "tmp-bill-render";
          billDiv.style.cssText = "position:absolute;left:-9999px;top:0;width:210mm;background:#fff;";
          document.body.appendChild(billDiv);

          // We need to get the rendered HTML — use a small timeout for React to render
          return new Promise((resolve) => {
            setTimeout(() => {
              const el = document.getElementById("bill-queue-single");
              if (!el) { billDiv.remove(); resolve(); return; }
              const billHTML = el.outerHTML;
              const fname = `bill_${(farmer?.farmerNo||"").replace(/[^a-zA-Z0-9-_]/g,"_")}_${(farmer?.name||"").replace(/\s+/g,"_").substring(0,20)}`;
              const fullHTML = `<!DOCTYPE html><html><head><meta charset="UTF-8"/><title>${fname}</title>
                <link href="https://fonts.googleapis.com/css2?family=Noto+Serif+Telugu&family=Noto+Serif:wght@400;600;700&display=swap" rel="stylesheet"/>
                <style>*{box-sizing:border-box;margin:0;padding:0;}body{font-family:'Noto Serif',Georgia,serif;padding:10px;background:#fff;}table{border-collapse:collapse;width:100%;}*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;}@page{margin:8mm;size:A4 portrait;}</style>
                </head><body>${billHTML}
                <script>document.fonts.ready.then(()=>{setTimeout(()=>{window.print();},600);});<\/script>
                </body></html>`;
              const blob = new Blob([fullHTML], {type:"text/html;charset=utf-8"});
              const url = URL.createObjectURL(blob);
              const tab = window.open(url, "_blank");
              setTimeout(() => URL.revokeObjectURL(url), 30000);
              billDiv.remove();
              resolve();
            }, 400);
          });
        };

        return (
          <div style={{position:"fixed",top:0,left:0,width:"100%",height:"100%",background:"rgba(0,0,0,0.88)",zIndex:9999,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:16}}>
            <div style={{color:"#fff",fontSize:18,fontWeight:700}}>🖨 Print Queue — Bill {printQueueIdx+1} of {printQueueTotal}</div>
            {/* Progress bar */}
            <div style={{width:360,background:"rgba(255,255,255,0.15)",borderRadius:10,overflow:"hidden",height:14}}>
              <div style={{background:"#2d6a2d",height:"100%",borderRadius:10,width:`${((printQueueIdx+1)/printQueueTotal)*100}%`,transition:"width 0.3s"}}></div>
            </div>
            <div style={{color:"#ccc",fontSize:14}}>Farmer #{qf?.farmerNo} — {qf?.name} — {qf?.village}</div>

            {/* Hidden bill rendered off-screen for capturing HTML */}
            <div style={{position:"absolute",left:"-9999px",top:0,width:"210mm",background:"#fff"}}>
              <div id="bill-queue-single">
                <BillPreview farmer={qf} varietySettings={varietySettings} getVarietyBillDate={getVarietyBillDate} isVarietyPaid={isVarietyPaid} getVarietyRate={getVarietyRate} getVarietyType={getVarietyType} />
              </div>
            </div>

            <div style={{display:"flex",gap:12,marginTop:8,flexWrap:"wrap",justifyContent:"center"}}>
              {/* Open in new tab and move to next */}
              <button onClick={async ()=>{
                await openInNewTab(qf);
                if (!isLast) setPrintQueueIdx(qi=>qi+1);
                else { setPrintQueueIdx(-1); setPrintQueue([]); }
              }} style={{background:"#2d6a2d",color:"#fff",border:"none",borderRadius:8,padding:"14px 32px",fontSize:16,fontWeight:700,cursor:"pointer"}}>
                🖨 Open PDF → {isLast ? "Done" : "Next Bill"}
              </button>
              {/* Open ALL remaining in tabs */}
              <button onClick={async ()=>{
                const remaining = printQueue.slice(printQueueIdx);
                for (let i=0; i<remaining.length; i++) {
                  setPrintQueueIdx(printQueueIdx + i);
                  await new Promise(r=>setTimeout(r,600));
                  const el = document.getElementById("bill-queue-single");
                  if (!el) continue;
                  const billHTML = el.outerHTML;
                  const f = remaining[i];
                  const fname = `bill_${(f?.farmerNo||"").replace(/[^a-zA-Z0-9-_]/g,"_")}_${(f?.name||"").replace(/\s+/g,"_").substring(0,20)}`;
                  const fullHTML = `<!DOCTYPE html><html><head><meta charset="UTF-8"/><title>${fname}</title>
                    <link href="https://fonts.googleapis.com/css2?family=Noto+Serif+Telugu&family=Noto+Serif:wght@400;600;700&display=swap" rel="stylesheet"/>
                    <style>*{box-sizing:border-box;margin:0;padding:0;}body{font-family:'Noto Serif',Georgia,serif;padding:10px;background:#fff;}table{border-collapse:collapse;width:100%;}*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;}@page{margin:8mm;size:A4 portrait;}</style>
                    </head><body>${billHTML}
                    <script>document.fonts.ready.then(()=>{setTimeout(()=>{window.print();},600);});<\/script>
                    </body></html>`;
                  const blob = new Blob([fullHTML], {type:"text/html;charset=utf-8"});
                  const url = URL.createObjectURL(blob);
                  window.open(url, "_blank");
                  setTimeout(()=>URL.revokeObjectURL(url), 30000);
                  await new Promise(r=>setTimeout(r,800));
                }
                setPrintQueueIdx(-1); setPrintQueue([]);
              }} style={{background:"#856404",color:"#fff",border:"none",borderRadius:8,padding:"14px 24px",fontSize:14,fontWeight:600,cursor:"pointer"}}>
                ⚡ Open All {printQueue.length - printQueueIdx} Tabs At Once
              </button>
              {!isLast && <button onClick={()=>setPrintQueueIdx(qi=>qi+1)} style={{background:"#2d5a8a",color:"#fff",border:"none",borderRadius:8,padding:"14px 24px",fontSize:15,cursor:"pointer"}}>⏭ Skip</button>}
              <button onClick={()=>{setPrintQueueIdx(-1);setPrintQueue([]);}} style={{background:"#c0392b",color:"#fff",border:"none",borderRadius:8,padding:"14px 20px",fontSize:15,cursor:"pointer"}}>✕ Stop</button>
            </div>

            {isLast && <div style={{color:"#2ecc71",fontSize:14,fontWeight:700}}>✅ Last bill in queue!</div>}
            <div style={{color:"#aaa",fontSize:12,textAlign:"center",maxWidth:440,lineHeight:1.6}}>
              <strong style={{color:"#fff"}}>🖨 Open PDF → Next Bill</strong> — opens one tab, moves to next bill<br/>
              <strong style={{color:"#ffd700"}}>⚡ Open All Tabs</strong> — opens all remaining bills in separate tabs at once.<br/>
              Each tab auto-opens print dialog. Allow popups for localhost:3000 if asked.
            </div>
          </div>
        );
      })()}
    </div>
  );


function calcInterest(amount, rate, fromDate, billDate) {
  const from = parseLocalDate(fromDate);
  const to = parseLocalDate(billDate || BILL_DATE);
  const days = Math.max(1, Math.round((to - from) / 86400000));
  return { days, interest: Math.round((amount * rate * days) / (100 * 30 * 12)) };
}

function printBill(elementId, filename) {
  const el = document.getElementById(elementId);
  if (!el) { alert("Please go to the Preview tab first, then click Print."); return; }

  const origTitle = document.title;
  document.title = filename || "farmer-bill";

  // Get the bill HTML with all inline styles (React renders inline styles so they copy perfectly)
  const billHTML = el.outerHTML;

  // Save original body
  const origBody = document.body.innerHTML;

  // Replace body with just the bill
  document.body.innerHTML = `
    <link href="https://fonts.googleapis.com/css2?family=Noto+Serif+Telugu&family=Noto+Serif:wght@400;600;700&display=swap" rel="stylesheet"/>
    <style>
      * { box-sizing: border-box; margin: 0; padding: 0; }
      body { font-family: 'Noto Serif', Georgia, serif; padding: 10px; background: #fff; }
      table { border-collapse: collapse; width: 100%; }
      * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
      @page { margin: 8mm; size: A4 portrait; }
    </style>
    ${billHTML}
  `;

  // Use setTimeout to let the DOM update before printing
  setTimeout(() => {
    window.print();
    // After print dialog closes, reload to restore the app
    setTimeout(() => {
      document.title = origTitle;
      window.location.reload();
    }, 500);
  }, 300);
}
}

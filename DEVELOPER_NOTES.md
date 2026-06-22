# FARMER BILL APP — DEVELOPER NOTES
# Created: June 2026
# Purpose: Reference document for any developer (or AI) working on this app

## BUSINESS CONTEXT
- Owner: Gopinath, Gadwal, Telangana, India
- Business: Agricultural seed production — gives advances to farmers who grow seed crops
- Sub-Organizers (Sub-Orgs): Intermediaries who manage groups of growers
- Season: Typically April-July crop cycle, billing in June-July
- Scale: 20-50 farmers currently, planning 1000+ farmers, 10-15 sub-orgs

## APP LOCATION (LOCAL)
- Path: C:\Users\gopin\Downloads\farmer-bill-app\farmer-bill-app\
- Start: double-click start.bat OR run `npm start`
- URL: http://localhost:3000
- Main file: src/App.js (~3400 lines)
- Firebase: src/firebase.js
- Storage: src/storage.js (localStorage + Firebase)

## FIREBASE CONFIG
- Project: farmer-bill-app
- Console: https://console.firebase.google.com
- Database: Cloud Firestore (asia-south1/Mumbai)
- Collections: app_data (farmers, suborgs, variety_settings, meta), snapshots

## CORE BUSINESS LOGIC

### Interest Formula
interest = (amount × rate × days) / (100 × 30 × 12)
- Uses 30-day months (banking convention)
- days = calendar days between advance date and bill date
- Minimum 1 day always charged
- Rate is annual percentage (e.g. 24% per year)
- Example: ₹1,00,000 × 24% × 365 days = ₹24,000/year

### Bill Date
- Global BILL_DATE stored in localStorage key "app_bill_date"
- Default: "2026-07-01"
- Settable from ⚙️ Settings panel in UI
- Changed each season — no code change needed
- All interest calculated to this date unless overridden per advance

### Till Date Override (per advance)
- Each advance can have a.tillDate which overrides BILL_DATE
- If set: interest = advance date → tillDate (not bill date)
- Used when: advance was settled before billing date
- UI: "+ Set custom interest end date" button per advance row

### Foundation Cost
- Foundation = area (acres) × ₹1000 per acre
- Always charged regardless of crop result
- Deducted from farmer's balance

### Transportation
- Transportation = total quantity (packets/kg) × ₹1 per unit
- Always charged
- Deducted from farmer's balance

### Balance Calculation (Farmer)
balance = Crop Value (paid varieties only) - Advance+Interest - Foundation - Transportation + Jamma+Interest
- Positive balance = we pay the farmer (Payable to Farmer)
- Negative balance = farmer owes us (Due from Farmer)

### Balance Calculation (Sub-Org Final Bill)
balance = To Pay seed amount - Advance+Interest + Jamma+Interest - Foundation - Transportation
- "To Pay" = varieties company has paid but not yet settled with sub-org
- "Settled" = varieties already paid to sub-org (per sub-org, not global)
- "Pending" = varieties company hasn't paid yet

## DATA MODELS

### Farmer Object
{
  id: number (Date.now() + random),
  farmerNo: string (e.g. "5-A"),
  name: string,
  fatherName: string,
  village: string,
  billingDone: boolean,
  billingDoneDate: string (YYYY-MM-DD),
  comment: string (optional — shows on bill if filled),
  advances: [{
    date: string (YYYY-MM-DD),
    amount: number,
    interestRate: number (annual %),
    note: string,
    tillDate: string (optional override)
  }],
  crops: [{
    variety: string,
    lotNo: string,
    area: string,
    quantity: number,
    cropType: string (KMS or GMS),
    ratePerUnit: number,
    result: string (Pass or Fail),
    note: string (optional),
    rateOverride: boolean (if true, use ratePerUnit instead of Variety Pay rate)
  }],
  jammaEnabled: boolean,
  jammaEntries: [{
    date: string (YYYY-MM-DD),
    amount: number,
    interestRate: number,
    note: string,
    tillDate: string (optional)
  }]
}

### Sub-Org Object
{
  id: number,
  accNo: string (unique, used for Excel import matching),
  name: string,
  village: string,
  comment: string (optional — shows on bill),
  _billDate: string (per sub-org bill date override),
  _billMode: string ("partial" or "final"),
  _selVars: array (selected varieties for partial bill),
  _settledVars: array (varieties already paid to THIS sub-org),
  _settledDates: object {variety: date} (when each was settled),
  _varietyNotes: object {variety: note},
  advances: [{date, amount, interestRate, note, tillDate}],
  growers: [{
    sNo: string,
    lotNo: string,
    name: string,
    fatherName: string,
    village: string,
    variety: string,
    packets: number,
    result: string (Pass or Fail),
    type: string (KMS or GMS),
    rate: number,
    note: string (optional)
  }],
  foundationSeeds: [{variety: string, area: string}],
  jammaEntries: [{date, amount, interestRate, note, tillDate}]
}

### Variety Settings (localStorage key: "variety_settings")
{
  [variety]: {
    status: string ("paid" or "pending"),
    type: string (KMS or GMS),
    rate: string (number as string),
    billDate: string (YYYY-MM-DD)
  },
  ["so_"+variety]: {  // sub-org settings — SAME rate/type as farmer
    status: string ("paid" or "pending"),
    billDate: string (YYYY-MM-DD)
  }
}

## KEY STATE VARIABLES (App Component)
- farmers: array of farmer objects
- subOrgs: array of sub-org objects  
- varietySettings: object (rate/type/status per variety)
- mode: string ("farmers" | "suborgs" | "dashboard" | "variety")
- tab: string ("form" | "preview" | "all")
- selectedIdx: number (current farmer index)
- subOrgTab: string ("form" | "growers" | "bill")
- selectedSubOrgIdx: number
- selectedVillage: string | null (village filter)
- farmerSearch: string (search text)
- farmerPage: number (pagination, 25 per page)
- BILL_DATE: string (from localStorage, settable in Settings)
- printQueue: array (farmers to print one by one)
- printQueueIdx: number (-1 = not printing)
- undoItem: object | null (for undo delete feature)
- cloudStatus: string ("idle"|"saving"|"saved"|"error")
- billingHistory: object (per sub-org bill print history)

## KEY HELPER FUNCTIONS (outside App)
- fmtDate(dateStr): YYYY-MM-DD → D/M/YYYY (no timezone shift)
- parseLocalDate(dateStr): YYYY-MM-DD → local Date object (no UTC shift)
- calcInterest(amount, rate, fromDate, billDate): returns {days, interest}
- printBill(elementId, filename): replaces body with bill HTML, calls window.print()

## IMPORTANT: DATE HANDLING
- ALL dates stored as YYYY-MM-DD strings
- NEVER use new Date(dateStr).toLocaleDateString() — timezone shift causes -1 day bug
- ALWAYS use fmtDate() for display and parseLocalDate() for calculation
- This is because India is UTC+5:30, so midnight UTC = 5:30 AM India = previous day

## VARIETY PAY TAB
- One unified table for ALL varieties (farmer + sub-org)
- Status/Type/Rate are SHARED — change once applies to both
- Farmer section: Farmers, Quantity, Crop Value, Adv+Int, Bal to Pay, Bal Due, Bill Date
- Sub-Org section: Sub-Orgs, Growers, Quantity, Crop Value, Bill Date
- Total Qty column = Farmer Qty + Sub-Org Qty
- "No Seed Farmers" row at bottom — farmers with advances but no crops

## BILLING WORKFLOW

### Farmer Billing
1. Enter farmer data in Edit tab (advances, crops, jamma)
2. Preview bill in Preview tab
3. Go to All Bills tab, select variety, print queue
4. Print one by one — each opens in new tab with auto print dialog
5. Mark "Billing Done" checkbox after meeting farmer

### Sub-Org Billing (Multi-Round)
1. Company pays for Variety X → mark as "Paid" in Variety Pay tab
2. Go to sub-org Bill tab → select "Partial Bill" → select Variety X → Print
3. Pay sub-org → enter amount as Jamma entry
4. Repeat for other varieties as company pays
5. Final round → select "Final Bill" → mark settled varieties
   - "Settled" = already paid in previous partial bills
   - "To Pay" = paid by company, paying sub-org now
   - "Pending" = company hasn't paid yet (not included in settlement)
6. Final bill shows 3 sections in growers table

## EXCEL IMPORT/EXPORT

### Farmer Export (single row per farmer)
Headers: Farmer No, Farmer Name, Father Name, Village, Bill Date,
         Adv1 Date, Adv1 Amount, Adv1 Days, Adv1 Interest, Adv1 Total, Adv1 Note,
         Crop1 Variety, Crop1 LOT No, Crop1 Qty, Crop1 Type, Crop1 Rate, Crop1 Result, Crop1 Value,
         Jamma1 Date, Jamma1 Amount, Jamma1 Days, Jamma1 Interest, Jamma1 Total, Jamma1 Note,
         Total Advance, Total Interest, Total Adv+Int, Total Crop Value, Foundation, Transportation, Balance

### Farmer Template (for data entry)
Headers: Farmer No, Farmer Name, Father Name, Village,
         Adv1 Date, Adv1 Amount, Adv1 Interest%, Adv1 Note, (Adv2...),
         Crop1 Variety, Crop1 LOT No, Crop1 Area, Crop1 Qty, Crop1 Type, Crop1 Rate, Crop1 Result, Crop1 Note,
         Jamma Enabled (Yes/No), Jamma1 Date, Jamma1 Amount, Jamma1 Interest%, Jamma1 Note,
         Farmer Comment

### Sub-Org Template (3 sheets)
- Sheet "Growers": Acc No, Sub-Org Name, Village, S.No, LOT No, Grower Name, Father Name, Grower Village, Variety, Packets, Result, Type, Rate, Note
- Sheet "Advances": Acc No, Date, Amount, Interest%, Note
- Sheet "Foundation": Acc No, Variety, Area (Ac)
- Acc No links rows across sheets to same sub-org

### Sub-Org Export
- One sheet per sub-org
- Sections: Header, Advances (with interest calc), Jamma, Foundation, Paid Growers, Pending Growers, Settlement Summary

## LOT NO RULES
- LOT No is assigned by the company (unique per grower per season)
- NO two farmers can have same LOT No
- NO farmer and sub-org grower can have same LOT No
- Duplicate detection: red border on input + warning banner at top of app
- LOT No stored as string (can start with 0, has letters sometimes)

## PRINT SYSTEM
- printBill(): replaces body HTML with bill, calls window.print(), then reloads
- Print queue: opens each bill in a new browser tab
- Each new tab auto-triggers print dialog
- Allow popups must be enabled for localhost:3000
- Files named: bill_{farmerNo}_{name}.pdf

## Telugu LANGUAGE
- App is bilingual: English + Telugu
- Telugu text in bill headers: "రైతు పంట బిల్లు" (Farmer Crop Bill)
- Font: Noto Serif Telugu (loaded from Google Fonts)
- Telugu text hardcoded in bill HTML, not translatable in current version

## KNOWN ISSUES / LIMITATIONS
1. App.js is 3400+ lines — needs splitting into components eventually
2. All state in one component — slow with 1000+ farmers (needs virtualization)
3. Print uses page reload — loses scroll position
4. XLSX library (free version) cannot add cell colors/borders — needs xlsx-js-style or openpyxl
5. No offline support — requires internet for Firebase sync
6. Two copies of calcInterest function in file (lines ~41 and ~2363) — keep in sync
7. Variety Pay rate change affects all reprinted bills (no bill locking)

## FIREBASE STRUCTURE
Firestore Collections:
  app_data/
    farmers: {data: JSON string of farmers array}
    suborgs: {data: JSON string of subOrgs array}
    variety_settings: {data: JSON string of varietySettings object}
    meta: {lastSaved, farmerCount, subOrgCount}
  snapshots/
    {YYYY-MM-DD}: {date, savedAt, farmers, subOrgs, farmerCount, subOrgCount}

Auto-save: triggered 2 seconds after any data change (debounced)
Load: on app startup, loads from Firebase first, falls back to localStorage

## DEPLOYMENT (PLANNED)
- Host: Vercel (free tier)
- URL: https://farmer-bill-app-xxx.vercel.app
- Deploy: `vercel --prod` or GitHub push
- Security: Firebase rules need password protection before going public

## FUTURE ROADMAP
1. Split App.js into components (FarmerForm, BillPreview, SubOrgBill, VarietyPay)
2. Virtual scrolling for 1000+ farmers list
3. Add password/PIN protection
4. Multi-season support (2025-26, 2026-27 selector)
5. Phone number field for farmers
6. Payment received tracking
7. Running total dashboard
8. PWA (Progressive Web App) for offline + phone install
9. Proper Android app via React Native

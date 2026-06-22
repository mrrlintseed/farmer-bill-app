# 🌾 Farmer Bill Generator — Local App

## Setup (One Time Only)

### Step 1 — Install Node.js
Go to https://nodejs.org and download the **LTS version**.
Install it like any normal software (Next → Next → Finish).

### Step 2 — Install the app
Open **Command Prompt** (Windows) or **Terminal** (Mac/Linux).
Navigate to this folder:
```
cd path\to\farmer-bill-app
```
Then run:
```
npm install
```
This downloads the required libraries (takes 1-2 minutes, one time only).

---

## Running the App

Every time you want to use the app, run:
```
npm start
```
The app opens automatically in your browser at http://localhost:3000

To stop the app, press **Ctrl+C** in the terminal.

---

## Using Across 4-5 Devices (OneDrive Sync)

### Save data to OneDrive:
1. Click **☁️ Save to OneDrive** button in the app
2. Save the downloaded `farmer_data_YYYY-MM-DD.json` file to your OneDrive folder
3. Rename it to `farmer_data.json` for easy finding

### Load data on another device:
1. Make sure OneDrive has synced on that device
2. Open the app on that device (npm start)
3. Click **📂 Load from OneDrive** button
4. Select the `farmer_data.json` file from your OneDrive folder
5. All your farmer data loads instantly!

### Important:
- ✅ Data saves automatically to the browser's local storage
- ✅ Use Save/Load buttons to sync across devices via OneDrive
- ⚠️ Don't edit on two devices at the same time

---

## Generating PDF Bills

- Go to **👁 Preview** tab for a single farmer
- Click **🖨 Download PDF** — saves directly to your Downloads folder
- Go to **📋 All Bills** tab → Click **📥 Download All PDFs**

---

## Data Storage

- **Local (this device):** Saved in browser's localStorage — instant, no internet needed
- **OneDrive sync:** Use the Save/Load buttons — your data as a simple JSON file
- **Excel backup:** Click 📊 Export Excel anytime for a spreadsheet backup

---

## Folder Structure
```
farmer-bill-app/
├── src/
│   ├── App.js          ← Main app code
│   ├── storage.js      ← Local + OneDrive storage
│   ├── pdfGenerator.js ← PDF generation
│   └── index.js        ← Entry point
├── public/
│   └── index.html
├── package.json
└── README.md           ← This file
```

---

## Run on Windows (Double-click shortcut)

You can create a shortcut: right-click `start.bat` → Send to Desktop.
Then just double-click it to start the app anytime.

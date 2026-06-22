// storage.js
// Data is stored in localStorage for instant access.
// Use "Export to OneDrive" button to save farmer_data.json to your OneDrive folder.
// Use "Load from OneDrive" button to load it back on another device.
// This way data syncs across all your devices via OneDrive.

const FARMERS_KEY = "farmer_bill_farmers";
const SUBORGS_KEY = "farmer_bill_suborgs";

export const storage = {
  getFarmers: () => {
    try {
      const data = localStorage.getItem(FARMERS_KEY);
      return data ? JSON.parse(data) : null;
    } catch { return null; }
  },

  saveFarmers: (farmers) => {
    try {
      localStorage.setItem(FARMERS_KEY, JSON.stringify(farmers));
      return true;
    } catch { return false; }
  },

  getSubOrgs: () => {
    try {
      const data = localStorage.getItem(SUBORGS_KEY);
      return data ? JSON.parse(data) : [];
    } catch { return []; }
  },

  saveSubOrgs: (suborgs) => {
    try {
      localStorage.setItem(SUBORGS_KEY, JSON.stringify(suborgs));
      return true;
    } catch { return false; }
  },

  // Export all data as a JSON file — save this to your OneDrive folder
  exportToFile: (farmers, suborgs) => {
    const data = {
      version: "1.0",
      exportedAt: new Date().toISOString(),
      farmers,
      suborgs,
    };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `farmer_data_${new Date().toISOString().split("T")[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);
  },

  // Load from a JSON file (from OneDrive folder)
  importFromFile: (file, onSuccess, onError) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = JSON.parse(e.target.result);
        if (data.farmers) {
          localStorage.setItem(FARMERS_KEY, JSON.stringify(data.farmers));
          localStorage.setItem(SUBORGS_KEY, JSON.stringify(data.suborgs || []));
          onSuccess(data.farmers, data.suborgs || []);
        } else {
          onError("Invalid file format");
        }
      } catch (err) {
        onError(err.message);
      }
    };
    reader.readAsText(file);
  },
};

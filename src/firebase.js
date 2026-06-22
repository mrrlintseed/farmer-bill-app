import { initializeApp } from "firebase/app";
import { getFirestore, doc, setDoc, getDoc, collection, getDocs } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyBPRzt6d4vkPMtx59ho8b686WdNcGD-lPg",
  authDomain: "farmer-bill-app.firebaseapp.com",
  projectId: "farmer-bill-app",
  storageBucket: "farmer-bill-app.firebasestorage.app",
  messagingSenderId: "375394419787",
  appId: "1:375394419787:web:6d71c910d4b68800eb79ce"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);

// ── Save data to Firestore ──
export const saveToCloud = async (farmers, subOrgs, varietySettings) => {
  try {
    await setDoc(doc(db, "app_data", "farmers"), { data: JSON.stringify(farmers) });
    await setDoc(doc(db, "app_data", "suborgs"), { data: JSON.stringify(subOrgs) });
    await setDoc(doc(db, "app_data", "variety_settings"), { data: JSON.stringify(varietySettings) });
    await setDoc(doc(db, "app_data", "meta"), {
      lastSaved: new Date().toISOString(),
      farmerCount: farmers.length,
      subOrgCount: subOrgs.length,
    });
    return true;
  } catch (err) {
    console.error("Firebase save error:", err);
    return false;
  }
};

// ── Load data from Firestore ──
export const loadFromCloud = async () => {
  try {
    const [farmersDoc, suborgsDoc, vsDoc] = await Promise.all([
      getDoc(doc(db, "app_data", "farmers")),
      getDoc(doc(db, "app_data", "suborgs")),
      getDoc(doc(db, "app_data", "variety_settings")),
    ]);
    return {
      farmers: farmersDoc.exists() ? JSON.parse(farmersDoc.data().data) : null,
      subOrgs: suborgsDoc.exists() ? JSON.parse(suborgsDoc.data().data) : [],
      varietySettings: vsDoc.exists() ? JSON.parse(vsDoc.data().data) : {},
    };
  } catch (err) {
    console.error("Firebase load error:", err);
    return null;
  }
};

// ── Save a daily snapshot ──
export const saveSnapshot = async (farmers, subOrgs) => {
  try {
    const today = new Date().toISOString().split("T")[0];
    await setDoc(doc(db, "snapshots", today), {
      date: today,
      savedAt: new Date().toISOString(),
      farmers: JSON.stringify(farmers),
      subOrgs: JSON.stringify(subOrgs),
      farmerCount: farmers.length,
      subOrgCount: subOrgs.length,
    });
    return true;
  } catch (err) {
    console.error("Snapshot error:", err);
    return false;
  }
};

// ── Get list of snapshots ──
export const getSnapshots = async () => {
  try {
    const snap = await getDocs(collection(db, "snapshots"));
    return snap.docs
      .map(d => ({ id: d.id, ...d.data() }))
      .sort((a, b) => b.date.localeCompare(a.date))
      .slice(0, 7); // last 7 days
  } catch (err) {
    return [];
  }
};

// ── Restore from a snapshot ──
export const restoreSnapshot = async (snapshotId) => {
  try {
    const snap = await getDoc(doc(db, "snapshots", snapshotId));
    if (!snap.exists()) return null;
    const data = snap.data();
    return {
      farmers: JSON.parse(data.farmers),
      subOrgs: JSON.parse(data.subOrgs),
    };
  } catch (err) {
    return null;
  }
};

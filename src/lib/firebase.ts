import { initializeApp, getApps } from "firebase/app";
import {
  getFirestore,
  doc,
  setDoc,
  increment,
  serverTimestamp,
} from "firebase/firestore";

const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID,
};

const app = getApps().length === 0 ? initializeApp(firebaseConfig) : getApps()[0];
export const db = getFirestore(app);

export const DEVICE_ID = "esp32-01";

export interface DeviceSettings {
  deviceId: string;
  mode: "manual" | "auto";
  manualValveOpen: boolean;
  autoTargetMl: number;
  autoTimes: string;
  maxOpenSec: number;
  settingsVersion: number;
}

export interface DeviceStatus {
  valveOpen: boolean;
  watering: boolean;
  flowMlPerMin: number;
  sessionMl: number;
  appliedVersion: number;
  rssi: number;
  uptimeSec: number;
  lastSeenAt: number;
}

export interface WateringLog {
  id: string;
  mode: "manual" | "auto";
  volumeMl: number;
  durationSec: number;
  startedAt: number;
  endedAt: number;
}

export function deviceDocRef() {
  return doc(db, "devices", DEVICE_ID);
}

export function isOnline(lastSeenAt: number | undefined): boolean {
  if (!lastSeenAt) return false;
  return Date.now() - lastSeenAt * 1000 < 15_000;
}

export async function setMode(mode: "manual" | "auto") {
  await setDoc(
    deviceDocRef(),
    { mode, settingsVersion: increment(1), updatedAt: serverTimestamp() },
    { merge: true },
  );
}

export async function setManualValve(open: boolean) {
  await setDoc(
    deviceDocRef(),
    { mode: "manual", manualValveOpen: open, settingsVersion: increment(1), updatedAt: serverTimestamp() },
    { merge: true },
  );
}

export async function updateAutoSettings(input: {
  autoTargetMl: number;
  autoTimes: string;
  maxOpenSec: number;
}) {
  await setDoc(
    deviceDocRef(),
    { ...input, mode: "auto", settingsVersion: increment(1), updatedAt: serverTimestamp() },
    { merge: true },
  );
}

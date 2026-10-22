import { useEffect, useState } from "react";
import {
  collection,
  getDocs,
  limit as fbLimit,
  onSnapshot,
  orderBy,
  query,
  where,
} from "firebase/firestore";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Progress } from "@/components/ui/progress";
import { Separator } from "@/components/ui/separator";
import { Skeleton } from "@/components/ui/skeleton";
import { toast } from "sonner";
import * as XLSX from "xlsx";
import {
  db,
  deviceDocRef,
  isOnline,
  setMode as fsSetMode,
  setManualValve as fsSetManualValve,
  updateAutoSettings as fsUpdateAuto,
  DEVICE_ID,
  type DeviceSettings,
  type DeviceStatus,
  type WateringLog,
} from "@/lib/firebase";
import {
  Droplets,
  Gauge,
  History,
  Leaf,
  Power,
  Timer,
  Wifi,
  WifiOff,
  Plus,
  X,
  Save,
  Clock3,
  Download,
} from "lucide-react";

const DEFAULT_SETTINGS: DeviceSettings = {
  deviceId: DEVICE_ID,
  mode: "manual",
  manualValveOpen: false,
  autoTargetMl: 500,
  autoTimes: "06:00,17:00",
  maxOpenSec: 300,
  settingsVersion: 1,
};

function formatVolume(ml: number) {
  if (ml >= 1000) return `${(ml / 1000).toFixed(2)} L`;
  return `${Math.round(ml)} ml`;
}

function formatDuration(sec: number) {
  if (sec < 60) return `${sec} dtk`;
  const m = Math.floor(sec / 60);
  const s = sec % 60;
  return s > 0 ? `${m} mnt ${s} dtk` : `${m} mnt`;
}

function formatDateTime(unix: number | null) {
  if (!unix) return "—";
  return new Intl.DateTimeFormat("id-ID", {
    day: "2-digit",
    month: "short",
    hour: "2-digit",
    minute: "2-digit",
  }).format(new Date(unix * 1000));
}

function formatLastSeen(unix: number) {
  return new Intl.DateTimeFormat("id-ID", {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  }).format(new Date(unix * 1000));
}

function toDateInputValue(d: Date) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function dayStart(dateStr: string) {
  return Math.floor(new Date(`${dateStr}T00:00:00`).getTime() / 1000);
}

function dayEnd(dateStr: string) {
  return Math.floor(new Date(`${dateStr}T23:59:59.999`).getTime() / 1000);
}

function excelDateTime(unix: number) {
  if (!unix) return "";
  return new Intl.DateTimeFormat("id-ID", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  }).format(new Date(unix * 1000));
}

async function fetchLogs(from?: number, to?: number, limitCount = 1000) {
  const logsRef = collection(db, "devices", DEVICE_ID, "logs");
  const constraints = [];
  if (from) constraints.push(where("startedAt", ">=", from));
  if (to) constraints.push(where("startedAt", "<=", to));
  constraints.push(orderBy("startedAt", "desc"), fbLimit(limitCount));
  const snap = await getDocs(query(logsRef, ...constraints));
  return snap.docs.map((d) => {
    const v = d.data();
    return {
      id: d.id,
      mode: v.mode === "auto" ? "auto" : "manual",
      volumeMl: Number(v.volumeMl) || 0,
      durationSec: Number(v.durationSec) || 0,
      startedAt: Number(v.startedAt) || 0,
      endedAt: Number(v.endedAt) || 0,
    } as WateringLog;
  });
}

export default function Home() {
  const [settings, setSettings] = useState<DeviceSettings>(DEFAULT_SETTINGS);
  const [status, setStatus] = useState<DeviceStatus | null>(null);
  const [, setNowTick] = useState(0);
  const [pending, setPending] = useState(false);

  const [fromDate, setFromDate] = useState("");
  const [toDate, setToDate] = useState("");
  const [exporting, setExporting] = useState<"daily" | "monthly" | null>(null);
  const [logsData, setLogsData] = useState<WateringLog[] | null>(null);
  const [logsError, setLogsError] = useState(false);

  useEffect(() => {
    const t = setInterval(() => setNowTick((n) => n + 1), 3000);
    return () => clearInterval(t);
  }, []);

  useEffect(() => {
    const unsub = onSnapshot(
      deviceDocRef(),
      (snap) => {
        if (!snap.exists()) {
          setSettings(DEFAULT_SETTINGS);
          setStatus(null);
          return;
        }
        const d = snap.data();
        setSettings({
          deviceId: DEVICE_ID,
          mode: d.mode === "auto" ? "auto" : "manual",
          manualValveOpen: Boolean(d.manualValveOpen),
          autoTargetMl: Number(d.autoTargetMl) || 500,
          autoTimes: String(d.autoTimes ?? "06:00"),
          maxOpenSec: Number(d.maxOpenSec) || 300,
          settingsVersion: Number(d.settingsVersion) || 1,
        });
        const s = d.status;
        setStatus(
          s
            ? {
                valveOpen: Boolean(s.valveOpen),
                watering: Boolean(s.watering),
                flowMlPerMin: Number(s.flowMlPerMin) || 0,
                sessionMl: Number(s.sessionMl) || 0,
                appliedVersion: Number(s.appliedVersion) || 0,
                rssi: Number(s.rssi) || 0,
                uptimeSec: Number(s.uptimeSec) || 0,
                lastSeenAt: Number(s.lastSeenAt) || 0,
              }
            : null,
        );
      },
      (err) => {
        console.error(err);
        toast.error("Gagal membaca data perangkat", {
          description: "Periksa konfigurasi Firebase",
        });
      },
    );
    return () => unsub();
  }, []);

  useEffect(() => {
    setLogsError(false);
    const logsRef = collection(db, "devices", DEVICE_ID, "logs");
    const constraints = [];
    if (fromDate) constraints.push(where("startedAt", ">=", dayStart(fromDate)));
    if (toDate) constraints.push(where("startedAt", "<=", dayEnd(toDate)));
    constraints.push(
      orderBy("startedAt", "desc"),
      fbLimit(fromDate || toDate ? 500 : 20),
    );
    const unsub = onSnapshot(
      query(logsRef, ...constraints),
      (snap) => {
        setLogsData(
          snap.docs.map((d) => {
            const v = d.data();
            return {
              id: d.id,
              mode: v.mode === "auto" ? "auto" : "manual",
              volumeMl: Number(v.volumeMl) || 0,
              durationSec: Number(v.durationSec) || 0,
              startedAt: Number(v.startedAt) || 0,
              endedAt: Number(v.endedAt) || 0,
            } as WateringLog;
          }),
        );
      },
      (err) => {
        console.error(err);
        setLogsError(true);
      },
    );
    return () => unsub();
  }, [fromDate, toDate]);

  const online = isOnline(status?.lastSeenAt);
  const watering = status?.watering ?? false;
  const valveOpen = status?.valveOpen ?? false;
  const mode = settings.mode;
  const sessionMl = status?.sessionMl ?? 0;
  const target = settings.autoTargetMl;
  const progress = Math.min(100, Math.round((sessionMl / Math.max(target, 1)) * 100));
  const synced = status ? status.appliedVersion >= settings.settingsVersion : false;

  // --- form pengaturan otomatis ---
  const [targetMl, setTargetMl] = useState("500");
  const [times, setTimes] = useState<string[]>(["06:00", "17:00"]);
  const [newTime, setNewTime] = useState("06:00");
  const [maxOpenSec, setMaxOpenSec] = useState("300");
  const [formDirty, setFormDirty] = useState(false);

  useEffect(() => {
    if (!formDirty) {
      setTargetMl(String(settings.autoTargetMl));
      setTimes(settings.autoTimes.split(",").filter(Boolean));
      setMaxOpenSec(String(settings.maxOpenSec));
    }
  }, [settings, formDirty]);

  const runMutation = async (fn: () => Promise<void>, okMsg: string) => {
    setPending(true);
    try {
      await fn();
      toast.success(okMsg, { description: "ESP32 menerapkan dalam ±5 detik" });
    } catch (e) {
      toast.error("Gagal menyimpan", {
        description: e instanceof Error ? e.message : "Periksa koneksi Firebase",
      });
    } finally {
      setPending(false);
    }
  };

  const addTime = () => {
    if (!newTime) return;
    if (times.includes(newTime)) {
      toast.error("Jam tersebut sudah ada");
      return;
    }
    if (times.length >= 6) {
      toast.error("Maksimal 6 jadwal per hari");
      return;
    }
    setTimes([...times, newTime].sort());
    setFormDirty(true);
  };

  const removeTime = (t: string) => {
    if (times.length <= 1) {
      toast.error("Minimal harus ada 1 jadwal");
      return;
    }
    setTimes(times.filter((x) => x !== t));
    setFormDirty(true);
  };

  const saveAuto = () => {
    const ml = parseInt(targetMl, 10);
    const sec = parseInt(maxOpenSec, 10);
    if (!ml || ml < 10 || ml > 100000) {
      toast.error("Batas air harus antara 10 - 100.000 ml");
      return;
    }
    if (!sec || sec < 10 || sec > 3600) {
      toast.error("Durasi maksimal harus antara 10 - 3600 detik");
      return;
    }
    if (times.length === 0) {
      toast.error("Tambahkan minimal 1 jadwal");
      return;
    }
    runMutation(
      () =>
        fsUpdateAuto({
          autoTargetMl: ml,
          autoTimes: times.join(","),
          maxOpenSec: sec,
        }),
      "Pengaturan otomatis tersimpan",
    );
    setFormDirty(false);
  };

  // --- export excel ---
  const exportDaily = async () => {
    const dateStr = fromDate || toDateInputValue(new Date());
    setExporting("daily");
    try {
      const data = await fetchLogs(dayStart(dateStr), dayEnd(dateStr));
      if (data.length === 0) {
        toast.error("Tidak ada data pada tanggal tersebut");
        return;
      }
      const rows = data.map((log) => [
        excelDateTime(log.startedAt),
        log.endedAt
          ? new Intl.DateTimeFormat("id-ID", { hour: "2-digit", minute: "2-digit" }).format(
              new Date(log.endedAt * 1000),
            )
          : "-",
        log.mode === "auto" ? "Otomatis" : "Manual",
        Math.round(log.volumeMl),
        log.durationSec,
      ]);
      const totalMl = data.reduce((s, l) => s + l.volumeMl, 0);
      const sheet = XLSX.utils.aoa_to_sheet([
        [`Laporan Penyiraman Harian - ${dateStr}`],
        [],
        ["Waktu Mulai", "Jam Selesai", "Mode", "Volume (ml)", "Durasi (detik)"],
        ...rows,
        [],
        ["TOTAL", "", "", Math.round(totalMl), ""],
      ]);
      sheet["!cols"] = [{ wch: 20 }, { wch: 12 }, { wch: 10 }, { wch: 12 }, { wch: 14 }];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, sheet, "Harian");
      XLSX.writeFile(wb, `penyiraman-harian-${dateStr}.xlsx`);
      toast.success("Excel harian terunduh");
    } catch {
      toast.error("Gagal membuat file Excel");
    } finally {
      setExporting(null);
    }
  };

  const exportMonthly = async () => {
    const base = fromDate ? new Date(`${fromDate}T00:00:00`) : new Date();
    const y = base.getFullYear();
    const m = base.getMonth();
    const first = Math.floor(new Date(y, m, 1, 0, 0, 0).getTime() / 1000);
    const last = Math.floor(new Date(y, m + 1, 0, 23, 59, 59).getTime() / 1000);
    const monthLabel = `${y}-${String(m + 1).padStart(2, "0")}`;
    setExporting("monthly");
    try {
      const data = await fetchLogs(first, last);
      if (data.length === 0) {
        toast.error("Tidak ada data pada bulan tersebut");
        return;
      }
      const perDay = new Map<string, { sesi: number; ml: number }>();
      for (const log of data) {
        const key = log.startedAt ? toDateInputValue(new Date(log.startedAt * 1000)) : "-";
        const cur = perDay.get(key) ?? { sesi: 0, ml: 0 };
        cur.sesi += 1;
        cur.ml += log.volumeMl;
        perDay.set(key, cur);
      }
      const rekapRows = [...perDay.entries()]
        .sort(([a], [b]) => a.localeCompare(b))
        .map(([tgl, v]) => [tgl, v.sesi, Math.round(v.ml), +(v.ml / 1000).toFixed(2)]);
      const totalMl = data.reduce((s, l) => s + l.volumeMl, 0);
      const rekap = XLSX.utils.aoa_to_sheet([
        [`Rekap Penyiraman Bulanan - ${monthLabel}`],
        [],
        ["Tanggal", "Jumlah Penyiraman", "Total Volume (ml)", "Total Volume (L)"],
        ...rekapRows,
        [],
        ["TOTAL BULAN", data.length, Math.round(totalMl), +(totalMl / 1000).toFixed(2)],
      ]);
      rekap["!cols"] = [{ wch: 14 }, { wch: 18 }, { wch: 16 }, { wch: 15 }];

      const detailRows = data.map((log) => [
        excelDateTime(log.startedAt),
        log.mode === "auto" ? "Otomatis" : "Manual",
        Math.round(log.volumeMl),
        log.durationSec,
      ]);
      const detail = XLSX.utils.aoa_to_sheet([
        ["Waktu Mulai", "Mode", "Volume (ml)", "Durasi (detik)"],
        ...detailRows,
      ]);
      detail["!cols"] = [{ wch: 20 }, { wch: 10 }, { wch: 12 }, { wch: 14 }];

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, rekap, "Rekap Bulanan");
      XLSX.utils.book_append_sheet(wb, detail, "Detail Sesi");
      XLSX.writeFile(wb, `penyiraman-bulanan-${monthLabel}.xlsx`);
      toast.success("Excel bulanan terunduh");
    } catch {
      toast.error("Gagal membuat file Excel");
    } finally {
      setExporting(null);
    }
  };

  return (
    <div className="mx-auto min-h-100dvh w-full max-w-5xl px-4 pb-16 pt-6 md:px-8">
      {/* ===== HEADER ===== */}
      <header className="flex items-center justify-between gap-3">
        <div className="flex items-center gap-3">
          <div className="relative flex h-10 w-10 items-center justify-center rounded-full border border-border bg-card">
            <Leaf className="h-5 w-5 text-primary" strokeWidth={1.75} />
          </div>
          <div>
            <h1 className="font-mono text-sm font-semibold uppercase tracking-[0.18em]">
              Penyiraman Otomatis
            </h1>
            <p className="label-chrome mt-0.5">{DEVICE_ID}</p>
          </div>
        </div>
        <div className="flex items-center gap-2 rounded-full border border-border bg-card px-3 py-1.5">
          <span className={`led ${online ? "led-on" : "led-off"}`} />
          <span className="font-mono text-[11px] font-medium uppercase tracking-[0.14em]">
            {online ? "Online" : "Offline"}
          </span>
          {online ? (
            <Wifi className="h-3.5 w-3.5 text-primary" />
          ) : (
            <WifiOff className="h-3.5 w-3.5 text-destructive" />
          )}
        </div>
      </header>

      <Separator className="my-5" />

      {/* ===== STATUS UTAMA ===== */}
      <section className="grid gap-4 md:grid-cols-[1.2fr_1fr]">
        <div className="relative overflow-hidden rounded-lg border border-border bg-card p-6">
          <p className="label-chrome">Status Keran</p>
          <div className="mt-4 flex items-center gap-5">
            <div className="relative flex h-16 w-16 items-center justify-center">
              {watering && (
                <>
                  <span className="radar-ring" />
                  <span className="radar-ring" style={{ animationDelay: "1.1s" }} />
                </>
              )}
              <div
                className={`flex h-14 w-14 items-center justify-center rounded-full border ${
                  watering
                    ? "border-accent/60 bg-accent/10"
                    : "border-border bg-secondary"
                }`}
              >
                <Droplets
                  className={`h-6 w-6 ${watering ? "text-accent" : "text-muted-foreground"}`}
                  strokeWidth={1.75}
                />
              </div>
            </div>
            <div>
              <p
                className={`font-mono text-2xl font-bold uppercase tracking-[0.08em] ${
                  watering ? "text-accent" : "text-foreground"
                }`}
              >
                {watering ? "Menyiram" : valveOpen ? "Terbuka" : "Standby"}
              </p>
              <p className="mt-1 font-mono text-[11px] uppercase tracking-[0.14em] text-muted-foreground">
                {online
                  ? synced
                    ? "ESP32 tersinkron"
                    : "Menunggu ESP32 menerapkan…"
                  : status?.lastSeenAt
                    ? `Terakhir aktif ${formatLastSeen(status.lastSeenAt)}`
                    : "ESP32 belum pernah terhubung"}
              </p>
            </div>
          </div>

          {watering && (
            <div className="mt-5">
              <div className="mb-1.5 flex items-center justify-between">
                <span className="label-chrome">Volume sesi ini</span>
                <span className="font-mono text-xs text-accent">
                  {formatVolume(sessionMl)} / {formatVolume(target)}
                </span>
              </div>
              <Progress value={progress} className="h-1.5" />
            </div>
          )}
        </div>

        <div className="grid grid-cols-2 gap-4">
          <div className="rounded-lg border border-border bg-card p-4">
            <div className="flex items-center gap-1.5">
              <Gauge className="h-3.5 w-3.5 text-accent" />
              <p className="label-chrome">Debit Air</p>
            </div>
            <p className="mt-3 font-mono text-xl font-semibold">
              {Math.round(status?.flowMlPerMin ?? 0)}
              <span className="ml-1 text-xs font-normal text-muted-foreground">ml/mnt</span>
            </p>
          </div>
          <div className="rounded-lg border border-border bg-card p-4">
            <div className="flex items-center gap-1.5">
              <Droplets className="h-3.5 w-3.5 text-accent" />
              <p className="label-chrome">Volume Sesi</p>
            </div>
            <p className="mt-3 font-mono text-xl font-semibold">
              {formatVolume(sessionMl)}
            </p>
          </div>
          <div className="rounded-lg border border-border bg-card p-4">
            <div className="flex items-center gap-1.5">
              <Timer className="h-3.5 w-3.5 text-primary" />
              <p className="label-chrome">Mode Aktif</p>
            </div>
            <p className="mt-3 font-mono text-xl font-semibold uppercase">
              {mode === "auto" ? "Otomatis" : "Manual"}
            </p>
          </div>
          <div className="rounded-lg border border-border bg-card p-4">
            <div className="flex items-center gap-1.5">
              <Clock3 className="h-3.5 w-3.5 text-primary" />
              <p className="label-chrome">Jadwal</p>
            </div>
            <p className="mt-3 font-mono text-sm font-semibold leading-6">
              {settings.autoTimes.split(",").join("  ·  ")}
            </p>
          </div>
        </div>
      </section>

      {/* ===== KONTROL ===== */}
      <section className="mt-6 rounded-lg border border-border bg-card">
        <div className="grid grid-cols-2">
          <button
            type="button"
            onClick={() =>
              mode !== "manual" &&
              runMutation(() => fsSetMode("manual"), "Mode MANUAL aktif")
            }
            className={`rounded-tl-lg px-4 py-3.5 font-mono text-xs font-semibold uppercase tracking-[0.18em] transition-colors duration-200 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring ${
              mode === "manual"
                ? "bg-primary text-primary-foreground"
                : "text-muted-foreground hover:text-foreground"
            }`}
          >
            Manual
          </button>
          <button
            type="button"
            onClick={() =>
              mode !== "auto" &&
              runMutation(() => fsSetMode("auto"), "Mode OTOMATIS aktif")
            }
            className={`rounded-tr-lg px-4 py-3.5 font-mono text-xs font-semibold uppercase tracking-[0.18em] transition-colors duration-200 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring ${
              mode === "auto"
                ? "bg-primary text-primary-foreground"
                : "text-muted-foreground hover:text-foreground"
            }`}
          >
            Otomatis
          </button>
        </div>

        <Separator />

        {mode === "manual" ? (
          <div className="flex flex-col items-center gap-5 p-6 md:p-8">
            <p className="max-w-md text-center text-sm text-muted-foreground">
              Tekan tombol untuk membuka keran air. Tekan lagi untuk menutup.
              Perintah diterapkan ESP32 dalam ±5 detik.
            </p>
            <button
              type="button"
              disabled={pending}
              onClick={() =>
                runMutation(
                  () => fsSetManualValve(!settings.manualValveOpen),
                  settings.manualValveOpen
                    ? "Perintah TUTUP keran dikirim"
                    : "Perintah BUKA keran dikirim",
                )
              }
              className={`relative flex h-36 w-36 flex-col items-center justify-center gap-2 rounded-full border-2 font-mono text-sm font-bold uppercase tracking-[0.14em] transition-all duration-200 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 focus-visible:ring-offset-background active:scale-95 disabled:opacity-50 ${
                settings.manualValveOpen
                  ? "border-accent bg-accent/15 text-accent shadow-[0_0_40px_-5px_hsl(188_86%_53%/0.5)]"
                  : "border-border bg-secondary text-foreground hover:border-primary/60"
              }`}
            >
              {settings.manualValveOpen && <span className="radar-ring" />}
              <Power className="h-8 w-8" strokeWidth={2} />
              {pending
                ? "Mengirim…"
                : settings.manualValveOpen
                  ? "Tutup Air"
                  : "Buka Air"}
            </button>
            <div className="flex items-center gap-2">
              <span className={`led ${settings.manualValveOpen ? "led-water" : "led-off"}`} />
              <span className="font-mono text-[11px] uppercase tracking-[0.14em] text-muted-foreground">
                Perintah saat ini: {settings.manualValveOpen ? "Keran terbuka" : "Keran tertutup"}
              </span>
            </div>
          </div>
        ) : (
          <div className="p-6 md:p-8">
            <p className="max-w-lg text-sm text-muted-foreground">
              ESP32 menyiram otomatis sesuai jadwal. Air berhenti saat volume
              mencapai batas, atau saat durasi maksimal terlampaui (pengaman).
            </p>

            <div className="mt-6 grid gap-5 md:grid-cols-2">
              <div>
                <label htmlFor="targetMl" className="label-chrome mb-2 block">
                  Batas air keluar (ml)
                </label>
                <div className="flex items-center gap-2">
                  <Input
                    id="targetMl"
                    type="number"
                    min={10}
                    max={100000}
                    step={50}
                    value={targetMl}
                    onChange={(e) => {
                      setTargetMl(e.target.value);
                      setFormDirty(true);
                    }}
                    className="font-mono"
                  />
                  <span className="font-mono text-xs text-muted-foreground">ml</span>
                </div>
                <p className="mt-1.5 text-xs text-muted-foreground">
                  ≈ {formatVolume(parseInt(targetMl, 10) || 0)} per penyiraman
                </p>
              </div>

              <div>
                <label htmlFor="maxOpenSec" className="label-chrome mb-2 block">
                  Durasi maksimal pengaman (detik)
                </label>
                <div className="flex items-center gap-2">
                  <Input
                    id="maxOpenSec"
                    type="number"
                    min={10}
                    max={3600}
                    step={10}
                    value={maxOpenSec}
                    onChange={(e) => {
                      setMaxOpenSec(e.target.value);
                      setFormDirty(true);
                    }}
                    className="font-mono"
                  />
                  <span className="font-mono text-xs text-muted-foreground">dtk</span>
                </div>
                <p className="mt-1.5 text-xs text-muted-foreground">
                  Keran menutup paksa bila flow sensor bermasalah
                </p>
              </div>
            </div>

            <div className="mt-6">
              <p className="label-chrome mb-2">Jadwal penyiraman</p>
              <div className="flex flex-wrap items-center gap-2">
                {times.map((t) => (
                  <span
                    key={t}
                    className="flex items-center gap-1.5 rounded-full border border-border bg-secondary py-1.5 pl-3.5 pr-2 font-mono text-sm"
                  >
                    {t}
                    <button
                      type="button"
                      onClick={() => removeTime(t)}
                      aria-label={`Hapus jadwal ${t}`}
                      className="flex h-5 w-5 items-center justify-center rounded-full text-muted-foreground transition-colors hover:bg-destructive/20 hover:text-destructive focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring"
                    >
                      <X className="h-3 w-3" />
                    </button>
                  </span>
                ))}
                <div className="flex items-center gap-1.5">
                  <Input
                    type="time"
                    value={newTime}
                    onChange={(e) => setNewTime(e.target.value)}
                    aria-label="Jam baru"
                    className="h-9 w-[110px] font-mono"
                  />
                  <Button
                    type="button"
                    variant="outline"
                    size="sm"
                    onClick={addTime}
                    className="h-9 font-mono text-xs uppercase tracking-wider"
                  >
                    <Plus className="mr-1 h-3.5 w-3.5" />
                    Tambah
                  </Button>
                </div>
              </div>
            </div>

            <div className="mt-7 flex items-center gap-3">
              <Button
                type="button"
                onClick={saveAuto}
                disabled={pending || !formDirty}
                className="font-mono text-xs font-semibold uppercase tracking-[0.14em]"
              >
                <Save className="mr-2 h-4 w-4" />
                {pending ? "Menyimpan…" : "Simpan Pengaturan"}
              </Button>
              {formDirty && (
                <span className="font-mono text-[11px] uppercase tracking-[0.14em] text-amber-400">
                  Ada perubahan belum disimpan
                </span>
              )}
            </div>
          </div>
        )}
      </section>

      {/* ===== RIWAYAT ===== */}
      <section className="mt-6 rounded-lg border border-border bg-card">
        <div className="flex items-center gap-2 px-6 py-4">
          <History className="h-4 w-4 text-muted-foreground" />
          <h2 className="font-mono text-xs font-semibold uppercase tracking-[0.18em]">
            Riwayat Penyiraman
          </h2>
        </div>
        <Separator />

        <div className="flex flex-wrap items-end gap-3 px-6 py-4">
          <div>
            <label htmlFor="fromDate" className="label-chrome mb-1.5 block">Dari tanggal</label>
            <Input
              id="fromDate"
              type="date"
              value={fromDate}
              onChange={(e) => setFromDate(e.target.value)}
              className="h-9 w-[150px] font-mono text-sm"
            />
          </div>
          <div>
            <label htmlFor="toDate" className="label-chrome mb-1.5 block">Sampai</label>
            <Input
              id="toDate"
              type="date"
              value={toDate}
              onChange={(e) => setToDate(e.target.value)}
              className="h-9 w-[150px] font-mono text-sm"
            />
          </div>
          {(fromDate || toDate) && (
            <Button
              type="button"
              variant="ghost"
              size="sm"
              onClick={() => {
                setFromDate("");
                setToDate("");
              }}
              className="h-9 font-mono text-xs uppercase tracking-wider text-muted-foreground"
            >
              <X className="mr-1 h-3.5 w-3.5" />
              Reset
            </Button>
          )}
          <div className="ml-auto flex flex-wrap gap-2">
            <Button
              type="button"
              variant="outline"
              size="sm"
              onClick={exportDaily}
              disabled={exporting !== null}
              className="h-9 font-mono text-xs uppercase tracking-wider"
            >
              <Download className="mr-1.5 h-3.5 w-3.5" />
              {exporting === "daily" ? "Membuat…" : "Excel Harian"}
            </Button>
            <Button
              type="button"
              variant="outline"
              size="sm"
              onClick={exportMonthly}
              disabled={exporting !== null}
              className="h-9 font-mono text-xs uppercase tracking-wider"
            >
              <Download className="mr-1.5 h-3.5 w-3.5" />
              {exporting === "monthly" ? "Membuat…" : "Excel Bulanan"}
            </Button>
          </div>
        </div>
        <p className="px-6 pb-3 text-xs text-muted-foreground">
          Excel harian memakai tanggal pada kolom "Dari" (kosong = hari ini).
          Excel bulanan memakai bulan dari tanggal tersebut (kosong = bulan ini).
        </p>
        <Separator />

        {logsData === null && !logsError ? (
          <div className="space-y-3 p-6">
            <Skeleton className="h-8 w-full" />
            <Skeleton className="h-8 w-full" />
            <Skeleton className="h-8 w-2/3" />
          </div>
        ) : logsError ? (
          <div className="p-10 text-center">
            <p className="text-sm text-destructive">Gagal memuat riwayat</p>
            <p className="mt-1 text-xs text-muted-foreground">
              Periksa konfigurasi Firebase
            </p>
          </div>
        ) : !logsData || logsData.length === 0 ? (
          <div className="flex flex-col items-center gap-2 border-dashed p-10 text-center">
            <Droplets className="h-8 w-8 text-muted-foreground/40" strokeWidth={1.5} />
            <p className="text-sm text-muted-foreground">
              {fromDate || toDate
                ? "Tidak ada penyiraman pada rentang tanggal ini."
                : "Belum ada riwayat. Penyiraman pertama akan tercatat di sini."}
            </p>
          </div>
        ) : (
          <ul className="divide-y divide-border">
            {logsData.map((log) => (
              <li
                key={log.id}
                className="flex items-center justify-between gap-3 px-6 py-3.5"
              >
                <div className="flex items-center gap-3">
                  <span
                    className={`led ${log.mode === "auto" ? "led-on" : "led-water"}`}
                    style={{ animation: "none" }}
                  />
                  <div>
                    <p className="font-mono text-sm font-medium">
                      {formatVolume(log.volumeMl)}
                      <span className="ml-2 text-[10px] uppercase tracking-[0.14em] text-muted-foreground">
                        {log.mode === "auto" ? "Otomatis" : "Manual"}
                      </span>
                    </p>
                    <p className="text-xs text-muted-foreground">
                      {formatDateTime(log.startedAt)}
                    </p>
                  </div>
                </div>
                <span className="font-mono text-xs text-muted-foreground">
                  {formatDuration(log.durationSec)}
                </span>
              </li>
            ))}
          </ul>
        )}
      </section>

      <footer className="mt-8 text-center">
        <p className="font-mono text-[10px] uppercase tracking-[0.2em] text-muted-foreground/60">
          ESP32 memperbarui status setiap 5 detik · Relay aktif LOW · Flow YF-S201
        </p>
      </footer>
    </div>
  );
}

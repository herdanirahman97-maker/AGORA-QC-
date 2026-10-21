# -*- coding: utf-8 -*-
"""
DIGITAL MONITORING SHAG - Siloam Hospitals Agora | Instalasi Radiologi
  1) Pemeriksaan Harian Suhu & Kelembaban Ruang Radiology  (SHKL 2205.11/2)
  2) Checklist Kegiatan Pemeliharaan Peralatan Radiologi    (RAD 2209.xx/00)
"""
import base64
import calendar
import hmac
import datetime as dt
import io
import json
import math
import os
import textwrap
import time
import urllib.parse
import urllib.request
from xml.sax.saxutils import escape

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Digital Monitoring QC Harian Radiologi SHAG {}".format(2026),
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={"Get help": None, "Report a bug": None, "About": None},
)

# =====================================================================
# 0. KONFIGURASI UTAMA (aman diubah)
# =====================================================================
APP_YEAR = 2026
APP_TITLE = "Digital Monitoring QC Harian Departemen Radiologi SHAG Tahun {}".format(APP_YEAR)
HOSPITAL = "Siloam Hospitals"
SITE = "AGORA"
DEPT_LINE = "INSTALASI RADIOLOGI SILOAM HOSPITALS AGORA"
FORM_CODE_TH = "SHKL 2205.11/2"          # kode form suhu & kelembaban (sesuai gambar template)
TZ = dt.timezone(dt.timedelta(hours=7))  # WIB. Streamlit Cloud berjalan di UTC, jadi zona waktu dipaksa ke WIB
QC_DUE_HOUR = 17                         # checklist QC hari ini dianggap "jatuh tempo" setelah jam ini

DATA_DIR = "data_shag"
ENTRY_FILE = os.path.join(DATA_DIR, "entries.csv")
QC_FILE = os.path.join(DATA_DIR, "qc_entries.csv")
CFG_FILE = os.path.join(DATA_DIR, "config.json")
LOGO_FILE = "Siloam Logo.png"
FONT_REG = "Lexend-Regular.ttf"
FONT_BOLD = "Lexend-Bold.ttf"
os.makedirs(DATA_DIR, exist_ok=True)

FONT_URLS = {
    FONT_REG: [
        "https://github.com/google/fonts/raw/main/ofl/lexend/static/Lexend-Regular.ttf",
        "https://raw.githubusercontent.com/google/fonts/main/ofl/lexend/static/Lexend-Regular.ttf",
    ],
    FONT_BOLD: [
        "https://github.com/google/fonts/raw/main/ofl/lexend/static/Lexend-Bold.ttf",
        "https://raw.githubusercontent.com/google/fonts/main/ofl/lexend/static/Lexend-Bold.ttf",
    ],
}

BLN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
SHIFT_NAME = {"P": "Pagi", "S": "Siang", "M": "Malam"}
SHIFT_LABEL = {"P": "P - Pagi (08.00)", "S": "S - Siang (14.00)", "M": "M - Malam (22.00)"}
SHIFT_HOUR = {"P": 8, "S": 14, "M": 22}

# Ruang + batas (sesuai template Excel FORM-SHG-KL-11). Suhu semua ruang 20-24 C.
_ROOMS = [("Mammography", 30, 75), ("USG", 30, 80), ("Dental Panoramik", 40, 55), ("X-Ray", 20, 75),
          ("Fluoroskopi", 20, 75), ("CT Scan", 50, 60), ("Panel CT Scan", 50, 60), ("MRI", 30, 70),
          ("Panel MRI", 30, 70), ("Lemari Alkes", 30, 70)]

_SUHU = "Periksa Suhu ruangan"
_HUM = "Periksa Kelembaban ruangan"
_BERSIH = "Periksa kebersihan peralatan dan ruangan"

# Peralatan + kegiatan harian (sesuai template Excel RAD 2209.xx/00)
_DEVICES = [
    dict(nama="Konvensional", judul="PERALATAN RADIOLOGI DIAGNOSTIK KONVENSIONAL", kode="RAD 2209.07/00", ruang="X-Ray",
         info=[["Nama Pesawat", "GE XR6000"], ["Tahun Op", "2022"], ["Lokasi", "Siloam Hospitals Agora"], ["Type Model", "XR 6000"],
               ["No Seri Tabung", "210182HL2"], ["Kapasitas Maks.", "150 kV"]],
         items=[_SUHU, _HUM, "Periksa mekanik", "Periksa lampu indikator pada control table", "Periksa brake (rem) tabung x-ray",
                "Periksa pergerakan cycling pada stand tabung", "Periksa lampu dan tombol pengatur kolimator dan lakukan ekspose",
                "Periksa kerja Bucky", "Periksa Generator dan lakukan eksposi", _BERSIH, "Periksa brake (rem) pada meja pemeriksaan"]),
    dict(nama="Mobile", judul="PERALATAN RADIOLOGI DIAGNOSTIK KONVENSIONAL/INTERVENTIONAL MOBILE", kode="RAD 2209.08/00", ruang=None, qr_ruang="X-Ray",
         info=[["Nama Pesawat", "Siemens Polymobil Plus"], ["Tahun Op", "2022"], ["Lokasi", "Siloam Hospitals Agora"],
               ["Type Model", "(240) 10187800"], ["No Seri Tabung", "(21) 24063"], ["Kapasitas Maks.", "125 kV"]],
         items=[_SUHU, _HUM, "Periksa mekanik", "Periksa lampu indikator pada control table",
                "Periksa pergerakan cycling pada stand tabung", "Periksa lampu kolimator", _BERSIH]),
    dict(nama="Dental Panoramik", judul="PERALATAN RADIOLOGI DIAGNOSTIK DENTAL PANORAMIK", kode="RAD 2209.03/00", ruang="Dental Panoramik",
         info=[["Nama Pesawat", "Carestream Dental"], ["Tahun Op", "2022"], ["Lokasi", "Siloam Hospitals Agora"], ["Type Model", "CS 2200"],
               ["No Seri Tabung", "JIXV141"], ["Kapasitas Maks.", "70 kV"]],
         items=[_SUHU, _HUM, "Periksa lampu indikator pada pad control", "Periksa pergerakan cycling pada stand tabung", _BERSIH]),
    dict(nama="CT Scan", judul="PERALATAN RADIOLOGI DIAGNOSTIK CT SCAN", kode="RAD 2209.04/00", ruang="CT Scan",
         info=[["Nama Pesawat", "Incisive CT"], ["Tahun Op", "2022"], ["Lokasi", "Siloam Hospitals Agora"],
               ["Type Model", "vMRCPerformanceTube"], ["No Seri Tabung", "174175"], ["Kapasitas Maks.", "140 kV"]],
         items=[_SUHU, _HUM, "Periksa mekanik", "Periksa lampu indikator pada control table", "Periksa Laser penanda",
                "Melakukan Kalibrasi harian", _BERSIH, "Warm Up Alat"]),
    dict(nama="MRI", judul="PERALATAN RADIOLOGI DIAGNOSTIK MRI", kode="RAD 2209.05/00", ruang="MRI",
         info=[["Nama Pesawat", "MRI INGENIA PRODIVA"], ["Tahun Op", "2022"], ["Lokasi", "Siloam Hospitals Agora"]],
         items=[_SUHU, _HUM, "Pengecekan akurasi posisi meja", "Evaluasi artefak", _BERSIH, "Test Phantom", "Helium level"]),
    dict(nama="USG", judul="PERALATAN RADIOLOGI DIAGNOSTIK USG", kode="RAD 2209.06/00", ruang="USG",
         info=[["Nama Pesawat", "Mindray Consona C9"], ["Tahun Op", "2024"], ["Lokasi", "Siloam Hospitals Agora"],
               ["Type Transduser", "Linear, Curved"], ["No Seri Transduser", "KS4U3C012713, NA6C025926, KS4M3C018995"]],
         items=[_SUHU, _HUM, "Kebersihan transduser", "Koneksi Kabel Transduser", "Kebersihan monitor", "Pengecekan komunikasi PACs",
                "Pengecekan kondisi dan fungsi panel", "Pengecekan kondisi panel kendali kecerahan dan kontras monitor", _BERSIH]),
    dict(nama="Mammografi", judul="PERALATAN RADIOLOGI DIAGNOSTIK MAMMOGRAFI", kode="RAD 2209.09/00", ruang="Mammography",
         info=[["Nama Pesawat", "Seno Crystal Nova (Mammography)"], ["Tahun Op", "2022"], ["Lokasi", "Siloam Hospitals Agora"],
               ["Type Model", "Senograph Crystal Nova"], ["No Seri Tabung", "DC1SS2200074HL"], ["Kapasitas Maks.", "49 kV"]],
         items=[_SUHU, _HUM, "Periksa mekanik", "Periksa lampu indikator pada control table",
                "Periksa pergerakan cycling pada stand tabung", "Periksa lampu kolimator", _BERSIH, "Flat field Test (Mingguan)"]),
]

DEFAULT_CFG = {
    "rooms": [dict(nama=n, tekanan=False, t_lo=20.0, t_hi=24.0, h_lo=float(hl), h_hi=float(hh)) for n, hl, hh in _ROOMS],
    "staff": ["RF", "NR", "AC", "AP", "AM"],
    "devices": _DEVICES,
}


def now():
    return dt.datetime.now(TZ)


def TODAY():
    return now().date()


# =====================================================================
# 1. AUTO-DOWNLOAD FONT LEXEND
# =====================================================================
def ensure_lexend_fonts():
    reg_ok = os.path.exists(FONT_REG)
    bold_ok = os.path.exists(FONT_BOLD)
    for path, urls in FONT_URLS.items():
        if os.path.exists(path):
            continue
        for url in urls:
            try:
                req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=20) as r, open(path, "wb") as f:
                    f.write(r.read())
                if path == FONT_REG:
                    reg_ok = True
                if path == FONT_BOLD:
                    bold_ok = True
                break
            except Exception:
                continue
    return reg_ok, bold_ok


# =====================================================================
# 2. PENYIMPANAN (CSV + JSON)
# =====================================================================
ECOLS = ["id", "dibuat", "ruang", "tanggal", "shift", "suhu", "kelembapan", "petugas", "status", "temuan", "tindakan"]
QCOLS = ["id", "dibuat", "alat", "tanggal", "petugas", "hasil", "catatan"]


def save_cfg(c):
    """Simpan ruang & staff. Daftar peralatan (devices) tetap dari kode, kecuali sudah ada manual di config.json."""
    out = {k: c[k] for k in ("rooms", "staff", "base_url") if k in c}
    if os.path.exists(CFG_FILE):
        try:
            with open(CFG_FILE, encoding="utf-8") as f:
                old = json.load(f)
            if old.get("devices"):
                out["devices"] = old["devices"]
        except Exception:
            pass
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(CFG_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)


def load_cfg():
    cfg = json.loads(json.dumps(DEFAULT_CFG))
    if os.path.exists(CFG_FILE):
        try:
            with open(CFG_FILE, encoding="utf-8") as f:
                c = json.load(f)
            for k in ("rooms", "staff", "devices", "base_url"):
                if c.get(k):
                    cfg[k] = c[k]
        except Exception:
            pass
    return cfg


def rooms():
    return load_cfg()["rooms"]


def room_names():
    return [r["nama"] for r in rooms()]


def room_std(name):
    for r in rooms():
        if r["nama"] == name:
            return r
    return dict(nama=name, tekanan=False, t_lo=20.0, t_hi=24.0, h_lo=30.0, h_hi=70.0)


def devices():
    return load_cfg()["devices"]


def dev_by_name(name):
    for d in devices():
        if d["nama"] == name:
            return d
    return devices()[0]


def _read(path, cols):
    if os.path.exists(path):
        try:
            df = pd.read_csv(path, dtype=str, encoding="utf-8-sig").fillna("")
        except Exception:
            df = pd.DataFrame(columns=cols)
    else:
        df = pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    return df[cols]


def _next_id(df):
    return str(1 + max([int(x) for x in df["id"] if str(x).isdigit()] + [0]))


def load_entries():
    return _read(ENTRY_FILE, ECOLS)


def load_qc():
    return _read(QC_FILE, QCOLS)


def save_entry(**kw):
    """Simpan 1 pengukuran. Data dengan ruang+tanggal+shift yang sama ditimpa."""
    df = load_entries()
    if len(df):
        m = (df["ruang"] == kw["ruang"]) & (df["tanggal"] == kw["tanggal"]) & (df["shift"] == kw["shift"])
        df = df[~m]
    row = {c: "" for c in ECOLS}
    for k, v in kw.items():
        row[k] = str(v)
    row["id"] = _next_id(df)
    row["dibuat"] = now().strftime("%Y-%m-%d %H:%M:%S")
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    os.makedirs(DATA_DIR, exist_ok=True)
    df.to_csv(ENTRY_FILE, index=False)


def delete_entry(eid):
    df = load_entries()
    df[df["id"] != str(eid)].to_csv(ENTRY_FILE, index=False)


def save_qc(alat, tanggal, petugas, hasil, catatan):
    """Simpan checklist QC 1 alat 1 hari. Data alat+tanggal yang sama ditimpa."""
    df = load_qc()
    if len(df):
        df = df[~((df["alat"] == alat) & (df["tanggal"] == tanggal))]
    row = dict(id=_next_id(df), dibuat=now().strftime("%Y-%m-%d %H:%M:%S"), alat=alat, tanggal=tanggal,
               petugas=petugas, hasil=json.dumps(hasil, ensure_ascii=False), catatan=catatan)
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    os.makedirs(DATA_DIR, exist_ok=True)
    df.to_csv(QC_FILE, index=False)


def delete_qc(eid):
    df = load_qc()
    df[df["id"] != str(eid)].to_csv(QC_FILE, index=False)


def qc_get(alat, tgl):
    df = load_qc()
    df = df[(df["alat"] == alat) & (df["tanggal"] == tgl.isoformat())]
    if not len(df):
        return None
    r = df.iloc[-1].to_dict()
    try:
        hasil = json.loads(r["hasil"]) if r["hasil"] else {}
    except Exception:
        hasil = {}
    return dict(petugas=r["petugas"], done=set(k for k, v in hasil.items() if v), catatan=r["catatan"], id=r["id"])


# =====================================================================
# 2b. QR CODE PER RUANGAN
# =====================================================================
def default_base_url():
    cfg = load_cfg()
    if cfg.get("base_url"):
        return cfg["base_url"]
    try:   # alamat aplikasi yang sedang dibuka (Streamlit >= 1.37)
        host = st.context.headers.get("Host")
        if host:
            return ("http://" if host.startswith(("localhost", "127.")) else "https://") + host
    except Exception:
        pass
    return "https://NAMA-APP.streamlit.app"


def room_url(base, ruang):
    return "{}/ruang?ruang={}".format(base.rstrip("/"), urllib.parse.quote(ruang))


def room_devices(ruang):
    """Peralatan yang checklist QC-nya diisi lewat QR ruangan ini."""
    return [d for d in devices() if (d.get("ruang") or d.get("qr_ruang")) == ruang]


def shift_now():
    h = now().hour
    return "P" if h < 11 else ("S" if h < 18 else "M")


def qr_matrix(text):
    from reportlab.graphics.barcode.qrencoder import QRCode, QRErrorCorrectLevel
    q = QRCode(None, QRErrorCorrectLevel.M)
    q.addData(text)
    q.make()
    n = q.getModuleCount()
    return [[bool(q.isDark(r, c)) for c in range(n)] for r in range(n)]


def fnum(v):
    return "%g" % v


# =====================================================================
# 3. DATA BULANAN + ANALISA : SUHU & KELEMBABAN
# =====================================================================
def slot_due(d, s):
    n = now()
    if d < n.date():
        return True
    if d > n.date():
        return False
    return n.hour >= SHIFT_HOUR[s]


def month_data(ruang, y, m):
    """{(tanggal, 'P'|'S'|'M'): dict(suhu, hum, petugas, ket, id)}"""
    out = {}
    df = load_entries()
    df = df[df["ruang"] == ruang]
    for r in df.to_dict("records"):
        try:
            td = dt.date.fromisoformat(r["tanggal"])
            if td.year != y or td.month != m:
                continue
            ket = r["temuan"] + (" (" + r["tindakan"] + ")" if r["tindakan"] and r["temuan"] else "")
            out[(td.day, r["shift"][:1])] = dict(suhu=float(r["suhu"]), hum=float(r["kelembapan"]),
                                                 petugas=r["petugas"], ket=ket, id=r["id"])
        except Exception:
            continue
    return out


def classify(v, lo, hi):
    return "ok" if lo <= v <= hi else ("high" if v > hi else "low")


def analyze(ruang, y, m):
    std = room_std(ruang)
    data = month_data(ruang, y, m)
    n = calendar.monthrange(y, m)[1]
    slots, empty_by_day = {}, {}
    high, low, hum_bad = [], [], []
    due = filled = filled_due = ok = hum_ok = 0
    for d in range(1, n + 1):
        for s in "PSM":
            is_due = slot_due(dt.date(y, m, d), s)
            e = data.get((d, s))
            if e is None:
                slots[(d, s)] = dict(state="empty" if is_due else "future")
                if is_due:
                    due += 1
                    empty_by_day.setdefault(d, []).append(s)
                continue
            st_ = classify(e["suhu"], std["t_lo"], std["t_hi"])
            hs = classify(e["hum"], std["h_lo"], std["h_hi"])
            slots[(d, s)] = dict(state=st_, t=e["suhu"], h=e["hum"], hstate=hs, petugas=e["petugas"])
            if is_due:
                due += 1
                filled_due += 1
            filled += 1
            ok += 1 if st_ == "ok" else 0
            hum_ok += 1 if hs == "ok" else 0
            if st_ == "high":
                high.append((d, s, e["suhu"]))
            elif st_ == "low":
                low.append((d, s, e["suhu"]))
            if hs != "ok":
                hum_bad.append((d, s, e["hum"], hs))
    full_empty = [d for d, v in empty_by_day.items() if len(v) == 3]
    return dict(std=std, n=n, slots=slots, empty_by_day=empty_by_day, high=high, low=low, hum_bad=hum_bad, due=due, filled=filled,
                filled_due=filled_due, ok=ok, hum_ok=hum_ok, pct_fill=(filled_due / due * 100 if due else 0.0), pct_ok=(ok / filled * 100 if filled else 0.0),
                pct_hum=(hum_ok / filled * 100 if filled else 0.0), full_empty=sorted(full_empty), n_empty=sum(len(v) for v in empty_by_day.values()))


def abnormal_today():
    """Pengukuran HARI INI (semua ruang) yang di luar batas -> perlu hubungi teknisi."""
    df = load_entries()
    df = df[df["tanggal"] == TODAY().isoformat()]
    bad = []
    for r in df.to_dict("records"):
        try:
            std = room_std(r["ruang"])
            sv, hv = float(r["suhu"]), float(r["kelembapan"])
        except Exception:
            continue
        ts, hs = classify(sv, std["t_lo"], std["t_hi"]), classify(hv, std["h_lo"], std["h_hi"])
        if ts != "ok" or hs != "ok":
            parts = []
            if ts != "ok":
                parts.append("suhu {:g}°C ({})".format(sv, "di atas" if ts == "high" else "di bawah"))
            if hs != "ok":
                parts.append("kelembapan {:g}% ({})".format(hv, "di atas" if hs == "high" else "di bawah"))
            bad.append("<b>{}</b> - {}: {}".format(escape(r["ruang"]), SHIFT_NAME.get(r["shift"][:1], r["shift"]), ", ".join(parts)))
    return bad


# =====================================================================
# 4. DATA BULANAN + ANALISA : QC PERALATAN
# =====================================================================
def auto_kind(item):
    s = item.lower()
    if s.startswith("periksa suhu"):
        return "suhu"
    if s.startswith("periksa kelembaban"):
        return "hum"
    return None


def is_daily(item):
    return "mingguan" not in item.lower()


def auto_status(dev, tgl):
    """Untuk alat yang terhubung ke ruang: apakah suhu/kelembaban hari itu sudah diinput."""
    res = dict(suhu=False, hum=False)
    if dev.get("ruang"):
        if any(True for (d, _s) in month_data(dev["ruang"], tgl.year, tgl.month) if d == tgl.day):
            res = dict(suhu=True, hum=True)
    return res


def qc_month(dev, y, m):
    """{tanggal: dict(petugas, done=set(item), catatan)} - sudah digabung dengan item otomatis suhu/kelembaban."""
    out = {}
    df = load_qc()
    df = df[df["alat"] == dev["nama"]]
    for r in df.to_dict("records"):
        try:
            td = dt.date.fromisoformat(r["tanggal"])
            if td.year != y or td.month != m:
                continue
            hasil = json.loads(r["hasil"]) if r["hasil"] else {}
            out[td.day] = dict(petugas=r["petugas"], done=set(k for k, v in hasil.items() if v), catatan=r["catatan"], id=r["id"])
        except Exception:
            continue
    if dev.get("ruang"):
        for (d, _s) in month_data(dev["ruang"], y, m):
            ent = out.setdefault(d, dict(petugas="", done=set(), catatan="", id=""))
            for it in dev["items"]:
                if auto_kind(it):
                    ent["done"].add(it)
    return out


def qc_analyze(dev, y, m):
    data = qc_month(dev, y, m)
    n = calendar.monthrange(y, m)[1]
    daily = [i for i in dev["items"] if is_daily(i)]
    n_now = now()
    days, missing = {}, []
    counted = full = filled = item_done = item_total = 0
    for d in range(1, n + 1):
        dd = dt.date(y, m, d)
        is_due = dd < n_now.date() or (dd == n_now.date() and n_now.hour >= QC_DUE_HOUR)
        e = data.get(d)
        miss = [i for i in daily if not (e and i in e["done"])]
        dn = len(daily) - len(miss)
        if dn == len(daily):
            state = "full"
        elif dn > 0:
            state = "part"
        else:
            state = "empty" if is_due else "future"
        days[d] = dict(state=state, done=dn, total=len(daily))
        if is_due or dn > 0:
            counted += 1
            item_total += len(daily)
            item_done += dn
            if state == "full":
                full += 1
            if dn > 0:
                filled += 1
            if state in ("part", "empty"):
                missing.append((d, state, miss))
    return dict(days=days, n=n, counted=counted, full=full, filled=filled, item_done=item_done, item_total=item_total,
                pct_full=(full / counted * 100 if counted else 0.0), pct_item=(item_done / item_total * 100 if item_total else 0.0),
                missing=missing, n_empty=sum(1 for _d, s, _m in missing if s == "empty"),
                n_part=sum(1 for _d, s, _m in missing if s == "part"))


# =====================================================================
# 5. GAMBAR FORMULIR (vektor: dipakai untuk preview SVG & PDF)
# =====================================================================
BAND, TEAL, NAVY, RED, BLACK = "#cfe2f3", "#1b7895", "#004e75", "#ff0000", "#000000"
PINK = "#f4cccc"   # pita "didalam batas normal"
GRAYL = "#efefef"


def _png_dimensions(path):
    try:
        with open(path, "rb") as f:
            head = f.read(24)
        if len(head) < 24 or head[:8] != b"\x89PNG\r\n\x1a\n":
            return None, None
        return int.from_bytes(head[16:20], "big"), int.from_bytes(head[20:24], "big")
    except Exception:
        return None, None


def _logo_b64():
    if os.path.exists(LOGO_FILE):
        with open(LOGO_FILE, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None


def _fit_box(src_w, src_h, dst_w, dst_h):
    if not src_w or not src_h:
        return dst_w, dst_h
    scale = min(dst_w / src_w, dst_h / src_h)
    return src_w * scale, src_h * scale


class Sheet:
    """Daftar operasi gambar dalam 'unit' (y ke bawah). width/height diisi builder."""

    def __init__(self, width=1000.0, height=700.0):
        self.ops = []
        self.width = width
        self.height = height

    def rect(self, x, y, w, h, fill=None, stroke=None, lw=1.0):
        self.ops.append(("rect", x, y, w, h, fill, stroke, lw))

    def line(self, x1, y1, x2, y2, color=BLACK, lw=1.0):
        self.ops.append(("line", x1, y1, x2, y2, color, lw))

    def text(self, x, y, s, size=8.0, bold=False, color=BLACK, anchor="start", rot=0):
        self.ops.append(("text", x, y, str(s), size, bold, color, anchor, rot))

    def circle(self, cx, cy, r, fill=None, stroke=None, lw=1.0):
        self.ops.append(("circle", cx, cy, r, fill, stroke, lw))

    def image(self, x, y, w, h, data_b64):
        self.ops.append(("image", x, y, w, h, data_b64))

    def ctext(self, x, y, w, h, s, size=8.0, bold=False, color=BLACK):
        lines = str(s).split("\n")
        lh = size * 1.18
        base = y + (h - lh * len(lines)) / 2 + size * 0.82
        for i, ln in enumerate(lines):
            self.text(x + w / 2, base + i * lh, ln, size, bold, color, "middle")

    def ltext(self, x, y, h, s, size=8.0, bold=False, color=BLACK):
        """Teks rata kiri, rata tengah secara vertikal dalam kotak setinggi h."""
        lines = str(s).split("\n")
        lh = size * 1.18
        base = y + (h - lh * len(lines)) / 2 + size * 0.82
        for i, ln in enumerate(lines):
            self.text(x, base + i * lh, ln, size, bold, color, "start")

    def check(self, cx, cy, color=BLACK, lw=1.2):
        self.line(cx - 3.4, cy + 0.2, cx - 1.1, cy + 2.9, color, lw)
        self.line(cx - 1.1, cy + 2.9, cx + 3.6, cy - 3.0, color, lw)


def _initials(sh, x, y, w, h, t):
    """Inisial staff: tebal & besar, ditulis tegak (diputar 90 derajat) agar muat di sel sempit."""
    t = (t or "").strip()
    if not t:
        return
    size = 9.2 if len(t) <= 2 else 7.6
    sh.text(x + w / 2 + size * 0.35, y + h / 2, t, size, True, BLACK, "middle", 90)


def _logo_into(sh, x, y, box_w, box_h):
    logo = _logo_b64()
    if logo:
        sw, shh = _png_dimensions(LOGO_FILE)
        fw, fh = _fit_box(sw, shh, box_w, box_h)
        sh.image(x, y + (box_h - fh) / 2, fw, fh, logo)
    else:   # cadangan bila file logo belum ada
        sh.text(x, y + box_h * 0.7, "Siloam", 16, True, "#1d4f9c")
        sh.text(x + 52, y + box_h * 0.7, SITE, 16, True, "#3a9a3a")


# ---------- FORM 1: SUHU & KELEMBABAN (SHKL 2205.11/2) ----------
X0, AW, BW = 6.0, 26.0, 26.0
XD = X0 + AW + BW
SUBW = 10.0
DAYW = SUBW * 3
XE = XD + 31 * DAYW
LW = 0.5
CELL = 21.0   # tinggi 1 label (= 2 baris kecil)


def build_th_form(ruang, y, m, data):
    std = room_std(ruang)
    tlo, thi, hlo, hhi = std["t_lo"], std["t_hi"], std["h_lo"], std["h_hi"]
    ymax = max(27, int(math.ceil(thi)) + 1)
    ymin = min(16, int(math.floor(tlo)) - 1)
    t_labels = [("{}°C".format(v), v) for v in range(ymax, ymin - 1, -1)]
    h_labels = [("{}%".format(v), v) for v in range(100, 0, -10)]
    W = XE + 6
    sh = Sheet(W, 0)

    # ---- header ----
    _logo_into(sh, X0, 3, 96, 32)
    sh.text(W / 2, 16, "Pemeriksaan Harian Suhu & Kelembaban Ruang Radiology - " + ruang, 12, True, anchor="middle")
    for i, (lab, val) in enumerate([(HOSPITAL, SITE), ("Ruang", ruang.upper()), ("Bulan", "{} {}".format(BLN[m - 1].upper(), y))]):
        sh.text(X0, 44 + i * 11, lab, 8.5)
        sh.text(X0 + 92, 44 + i * 11, ": " + val, 8.5)
    sh.text(420, 44, "Suhu", 8.5)
    sh.text(480, 44, ": {:g} - {:g} °C".format(tlo, thi), 8.5)
    sh.text(420, 55, "Kelembaban", 8.5)
    sh.text(480, 55, ": {:g} - {:g} %".format(hlo, hhi), 8.5)
    sh.rect(740, 36, 13, 9, PINK, "#9e9e9e", 0.5)
    sh.text(757, 44, ": didalam batas normal", 8.5)

    def block(y0, kind):
        HT, HJ, HN = 13.0, 11.0, 24.0
        # tanggal & jam
        sh.rect(X0, y0, XD - X0, HT, None, BLACK, LW)
        sh.ctext(X0, y0, XD - X0, HT, "Tanggal", 7.5)
        y1 = y0 + HT
        sh.rect(X0, y1, XD - X0, HJ, None, BLACK, LW)
        sh.ctext(X0, y1, XD - X0, HJ, "Jam", 7)
        for d in range(1, 32):
            x = XD + (d - 1) * DAYW
            sh.rect(x, y0, DAYW, HT, None, BLACK, LW)
            sh.ctext(x, y0, DAYW, HT, str(d), 7.5)
            for k, ch in enumerate("PSM"):
                sh.rect(x + k * SUBW, y1, SUBW, HJ, None, BLACK, LW)
                sh.ctext(x + k * SUBW, y1, SUBW, HJ, ch, 6.5, True)
        gt = y1 + HJ
        labels = t_labels if kind == "t" else h_labels
        gh = len(labels) * CELL
        if kind == "t":
            yv = lambda v: gt + (ymax - v) * CELL + CELL / 2
            lo, hi, side = tlo, thi, "Suhu"
        else:
            yv = lambda v: gt + (100 - v) / 10.0 * CELL + CELL / 2
            lo, hi, side = hlo, hhi, "Kelembaban"
        # pita batas normal
        top, bot = max(gt, yv(hi)), min(gt + gh, yv(lo))
        if bot > top:
            sh.rect(XD, top, XE - XD, bot - top, PINK, None, 0)
        # garis grid
        for i in range(0, 2 * len(labels) + 1):
            yy = gt + i * CELL / 2
            sh.line(XD, yy, XE, yy, BLACK, LW)
        for i in range(0, 93):
            xx = XD + i * SUBW
            sh.line(xx, gt, xx, gt + gh, BLACK, 0.7 if i % 3 == 0 else LW)
        sh.rect(XD, gt, XE - XD, gh, None, BLACK, 0.9)
        # kolom label
        sh.rect(X0, gt, AW, gh, None, BLACK, LW)
        sh.text(X0 + AW / 2 + 3.2, gt + gh / 2, side, 9, False, BLACK, "middle", 90)
        for i, (lab, _v) in enumerate(labels):
            sh.rect(X0 + AW, gt + i * CELL, BW, CELL, None, BLACK, LW)
            sh.ctext(X0 + AW, gt + i * CELL, BW, CELL, lab, 5.5)
        # baris nama
        yn = gt + gh
        sh.rect(X0, yn, AW, HN, None, BLACK, LW)
        sh.ctext(X0, yn, AW, HN, "Nama", 6.5)
        sh.rect(X0 + AW, yn, BW, HN, BLACK, BLACK, LW)
        for d in range(1, 32):
            for k, s in enumerate("PSM"):
                x = XD + (d - 1) * DAYW + k * SUBW
                sh.rect(x, yn, SUBW, HN, None, BLACK, LW)
                e = data.get((d, s))
                if e:
                    _initials(sh, x, yn, SUBW, HN, e["petugas"])
        # titik data
        for (d, s), e in data.items():
            v = e["suhu"] if kind == "t" else e["hum"]
            k = "PSM".index(s)
            cx = XD + (d - 1) * DAYW + (k + .5) * SUBW
            cy = min(max(yv(v), gt + 3.0), gt + gh - 3.0)
            sh.circle(cx, cy, 2.8, BLACK if lo <= v <= hi else RED)
        return yn + HN

    yb = block(72.0, "t")
    yb = block(yb + 6, "h")

    # ---- catatan & kode ----
    sh.text(XE, yb + 11, FORM_CODE_TH, 8.5, False, BLACK, "end")
    sh.text(X0, yb + 15, "Catatan", 8.5)
    sh.text(X0 + 52, yb + 15, ":", 8.5)
    sh.text(X0 + 62, yb + 15, "1) Hubungi teknisi bila terjadi ketidaksesuaian pengukuran batas normal", 8.5)
    sh.text(X0 + 62, yb + 26, "2) Batas suhu dan kelembaban berdasarkan Permenkes No 24 Tahun 2020 dan Manual Book Alat", 8.5)
    sh.text(X0 + 62, yb + 37, "3) P = Pagi (pukul 08.00), S = Siang (pukul 14.00), M = Malam (pukul 22.00)", 8.5)
    yend = yb + 41
    notes = []
    for (d, s), e in sorted(data.items()):
        if e.get("ket"):
            notes.extend(textwrap.wrap("Tgl {}/{}: {}".format(d, SHIFT_NAME[s], e["ket"]), 190))
    if notes:
        sh.text(X0 + 62, yend + 11, "Temuan:", 8, True)
        for i, ln in enumerate(notes[:6]):
            sh.text(X0 + 62, yend + 22 + i * 10, ln, 7.5)
        yend += 22 + min(len(notes), 6) * 10
    sh.height = yend + 8
    return sh


# ---------- FORM 2: CHECKLIST PEMELIHARAAN (RAD 2209.xx/00) ----------
def build_qc_form(dev, y, m, data):
    items = dev["items"]
    X0_, NOW_, ACTW, DAYW_ = 6.0, 24.0, 235.0, 17.0
    XD_ = X0_ + NOW_ + ACTW
    XE_ = XD_ + 31 * DAYW_
    W = XE_ + 6
    sh = Sheet(W, 0)
    cy = 14.0
    for ln in ("CHECKLIST KEGIATAN PEMELIHARAAN", dev["judul"], DEPT_LINE):
        sh.text(W / 2, cy, ln, 11, True, anchor="middle")
        cy += 14
    # tabel info alat
    yt = cy + 6
    IH, LABW, VALW = 16.0, 112.0, 270.0
    for i, (lab, val) in enumerate(dev["info"]):
        yy = yt + i * IH
        sh.rect(X0_, yy, LABW, IH, None, BLACK, LW)
        sh.ltext(X0_ + 4, yy, IH, lab, 8)
        sh.rect(X0_ + LABW, yy, VALW, IH, None, BLACK, LW)
        sh.ltext(X0_ + LABW + 4, yy, IH, val, 8)
    yh = yt + len(dev["info"]) * IH + 14
    HH, RHI = 16.0, 22.0
    sh.rect(X0_, yh, NOW_, 2 * HH, GRAYL, BLACK, LW)
    sh.ctext(X0_, yh, NOW_, 2 * HH, "No.", 8, True)
    sh.rect(X0_ + NOW_, yh, ACTW, 2 * HH, GRAYL, BLACK, LW)
    sh.ctext(X0_ + NOW_, yh, ACTW, 2 * HH, "Nama Kegiatan harian", 8.5, True)
    sh.rect(XD_, yh, XE_ - XD_, HH, GRAYL, BLACK, LW)
    sh.ctext(XD_, yh, XE_ - XD_, HH, "{} {}".format(BLN[m - 1].upper(), y), 9, True)
    for d in range(1, 32):
        x = XD_ + (d - 1) * DAYW_
        sh.rect(x, yh + HH, DAYW_, HH, GRAYL, BLACK, LW)
        sh.ctext(x, yh + HH, DAYW_, HH, str(d), 7.5)
    yi = yh + 2 * HH
    for i, it in enumerate(items):
        yy = yi + i * RHI
        sh.rect(X0_, yy, NOW_, RHI, None, BLACK, LW)
        sh.ctext(X0_, yy, NOW_, RHI, str(i + 1), 8)
        sh.rect(X0_ + NOW_, yy, ACTW, RHI, None, BLACK, LW)
        sh.ltext(X0_ + NOW_ + 4, yy, RHI, "\n".join(textwrap.wrap(it, 50)[:2]), 7.5)
        for d in range(1, 32):
            x = XD_ + (d - 1) * DAYW_
            sh.rect(x, yy, DAYW_, RHI, None, BLACK, LW)
            e = data.get(d)
            if e and it in e["done"]:
                sh.check(x + DAYW_ / 2, yy + RHI / 2)
    ys = yi + len(items) * RHI
    NH = 24.0
    sh.rect(X0_, ys, NOW_ + ACTW, NH, None, BLACK, LW)
    sh.ctext(X0_, ys, NOW_ + ACTW, NH, "Nama Staff", 8.5, True)
    for d in range(1, 32):
        x = XD_ + (d - 1) * DAYW_
        sh.rect(x, ys, DAYW_, NH, None, BLACK, LW)
        e = data.get(d)
        if e and e["petugas"]:
            sh.ctext(x, ys, DAYW_, NH, e["petugas"], 8.5 if len(e["petugas"]) <= 2 else 7, True)
    yend = ys + NH
    sh.text(XE_, yend + 12, dev["kode"], 8.5, False, BLACK, "end")
    notes = []
    for d in sorted(data):
        if data[d].get("catatan"):
            notes.extend(textwrap.wrap("Tgl {}: {}".format(d, data[d]["catatan"]), 170))
    if notes:
        sh.text(X0_, yend + 12, "Catatan:", 8.5, True)
        for i, ln in enumerate(notes[:6]):
            sh.text(X0_ + 6, yend + 24 + i * 10, ln, 7.5)
        yend += 24 + min(len(notes), 6) * 10
    sh.height = yend + 20
    return sh



# ---------- LEMBAR QR CODE RUANGAN (stiker, 6 per halaman A4 portrait) ----------
def build_qr_sheets(base):
    CW, CH = 240.0, 270.0
    cards = []
    for rn in room_names():
        devs = [d["nama"] for d in room_devices(rn)]
        cards.append((rn, room_url(base, rn), devs))
    sheets = []
    for p0 in range(0, len(cards), 6):
        chunk = cards[p0:p0 + 6]
        rows = max(1, math.ceil(len(chunk) / 2))
        sh = Sheet(2 * CW, rows * CH)
        for i, (rn, url, devs) in enumerate(chunk):
            x, y = (i % 2) * CW, (i // 2) * CH
            sh.rect(x + 3, y + 3, CW - 6, CH - 6, None, "#9e9e9e", 0.6)
            _logo_into(sh, x + CW / 2 - 50, y + 10, 100, 30)
            sh.text(x + CW / 2, y + 52, "INSTALASI RADIOLOGI", 8, True, "#555555", "middle")
            sh.text(x + CW / 2, y + 77, rn, 17 if len(rn) > 12 else 21, True, BLACK, "middle")
            mat = qr_matrix(url)
            n = len(mat)
            size = 128.0
            qx, qy = x + (CW - size) / 2, y + 90
            sh.rect(qx - 7, qy - 7, size + 14, size + 14, "#ffffff", "#bdbdbd", 0.5)
            cell = size / n
            for r in range(n):   # gabungkan modul gelap berurutan jadi satu persegi panjang
                c = 0
                while c < n:
                    if mat[r][c]:
                        c2 = c
                        while c2 + 1 < n and mat[r][c2 + 1]:
                            c2 += 1
                        sh.rect(qx + c * cell, qy + r * cell, (c2 - c + 1) * cell, cell, BLACK, None, 0)
                        c = c2 + 1
                    else:
                        c += 1
            sh.text(x + CW / 2, y + 242, "SCAN UNTUK MENGISI", 9.5, True, BLACK, "middle")
            sh.text(x + CW / 2, y + 254, "Suhu & Kelembaban" + (" + QC " + ", ".join(devs) if devs else ""), 7, False, "#333333", "middle")
            sh.text(x + CW / 2, y + 264, url if len(url) < 70 else url[:67] + "...", 4.8, False, "#888888", "middle")
        sheets.append(sh)
    return sheets

# ---------- konverter: SVG (preview) & PDF ----------
def _font_css():
    css = ""
    for path, wgt in ((FONT_REG, 400), (FONT_BOLD, 700)):
        if os.path.exists(path):
            try:
                with open(path, "rb") as f:
                    b = base64.b64encode(f.read()).decode()
                css += "@font-face{{font-family:'LexendEmbed';font-style:normal;font-weight:{};src:url(data:font/ttf;base64,{}) format('truetype');}}".format(wgt, b)
            except Exception:
                pass
    return css


def sheet_to_svg(sh, width_px=1400):
    W, H = sh.width, sh.height
    p = ["<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 {:.1f} {:.1f}' width='{}' height='{}' font-family=\"LexendEmbed, Lexend, Helvetica, Arial, sans-serif\">".format(
        W, H, width_px, int(width_px * H / W))]
    css = _font_css()
    if css:
        p.append("<defs><style type='text/css'>{}</style></defs>".format(css))
    p.append("<rect x='0' y='0' width='{:.1f}' height='{:.1f}' fill='#ffffff'/>".format(W, H))
    for op in sh.ops:
        k = op[0]
        if k == "rect":
            _, x, y, w, h, fill, stroke, lw = op
            p.append("<rect x='{:.2f}' y='{:.2f}' width='{:.2f}' height='{:.2f}' fill='{}' stroke='{}' stroke-width='{}'/>".format(
                x, y, w, h, fill or "none", stroke or "none", lw))
        elif k == "line":
            _, x1, y1, x2, y2, c, lw = op
            p.append("<line x1='{:.2f}' y1='{:.2f}' x2='{:.2f}' y2='{:.2f}' stroke='{}' stroke-width='{}' stroke-linecap='round'/>".format(x1, y1, x2, y2, c, lw))
        elif k == "circle":
            _, cx, cy, r, fill, stroke, lw = op
            p.append("<circle cx='{:.2f}' cy='{:.2f}' r='{}' fill='{}' stroke='{}' stroke-width='{}'/>".format(cx, cy, r, fill or "none", stroke or "none", lw))
        elif k == "image":
            _, x, y, w, h, b64 = op
            p.append("<image x='{:.2f}' y='{:.2f}' width='{:.2f}' height='{:.2f}' preserveAspectRatio='xMidYMid meet' href='data:image/png;base64,{}'/>".format(x, y, w, h, b64))
        else:
            _, x, y, s, size, bold, color, anchor, rot = op
            tf = " transform='rotate(-{} {:.2f} {:.2f})'".format(rot, x, y) if rot else ""
            p.append("<text x='{:.2f}' y='{:.2f}' font-size='{}' font-weight='{}' fill='{}' text-anchor='{}'{}>{}</text>".format(
                x, y, size, "bold" if bold else "normal", color, anchor, tf, escape(s)))
    return "".join(p) + "</svg>"


LEXEND_STATUS = {"reg": False, "bold": False, "tried": False}


def _register_lexend():
    if LEXEND_STATUS["tried"]:
        return LEXEND_STATUS["reg"], LEXEND_STATUS["bold"]
    ensure_lexend_fonts()
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    ok = {}
    for key, name, path in (("reg", "Lexend", FONT_REG), ("bold", "Lexend-Bold", FONT_BOLD)):
        ok[key] = False
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont(name, path))
                ok[key] = True
            except Exception:
                pass
    LEXEND_STATUS.update(reg=ok["reg"], bold=ok["bold"], tried=True)
    return ok["reg"], ok["bold"]


def sheets_to_pdf(sheets, title="Formulir Radiologi - Siloam Hospitals Agora", portrait=False):
    from reportlab.lib.colors import HexColor
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen.canvas import Canvas
    from reportlab.lib.utils import ImageReader

    reg_ok, bold_ok = _register_lexend()
    fn_reg = "Lexend" if reg_ok else "Helvetica"
    fn_bold = "Lexend-Bold" if bold_ok else "Helvetica-Bold"
    buf = io.BytesIO()
    pw, ph = A4 if portrait else landscape(A4)
    
    # --- UKURAN MARGIN DIPERKECIL DARI 22.0 MENJADI 10.0 AGAR PDF LEBIH BESAR & JELAS ---
    MG = 10.0
    
    c = Canvas(buf, pagesize=(pw, ph))
    c.setTitle(title)
    c.setAuthor("Siloam Hospitals Agora - Instalasi Radiologi")

    for sh in sheets:
        s = min((pw - 2 * MG) / sh.width, (ph - 2 * MG) / sh.height)
        ox = (pw - sh.width * s) / 2
        tx = lambda x: ox + x * s
        ty = lambda y: ph - (MG + y * s)
        for op in sh.ops:
            k = op[0]
            if k == "rect":
                _, x, y, w, h, fill, stroke, lw = op
                if not fill and not stroke:
                    continue
                if fill:
                    c.setFillColor(HexColor(fill))
                if stroke:
                    c.setStrokeColor(HexColor(stroke))
                    c.setLineWidth(max(lw * s, 0.1))
                c.rect(tx(x), ty(y + h), w * s, h * s, stroke=1 if stroke else 0, fill=1 if fill else 0)
            elif k == "line":
                _, x1, y1, x2, y2, col, lw = op
                c.setStrokeColor(HexColor(col))
                c.setLineWidth(max(lw * s, 0.1))
                c.setLineCap(1)
                c.line(tx(x1), ty(y1), tx(x2), ty(y2))
            elif k == "circle":
                _, cx, cy, r, fill, stroke, lw = op
                if fill:
                    c.setFillColor(HexColor(fill))
                if stroke:
                    c.setStrokeColor(HexColor(stroke))
                    c.setLineWidth(max(lw * s, 0.1))
                c.circle(tx(cx), ty(cy), r * s, stroke=1 if stroke else 0, fill=1 if fill else 0)
            elif k == "image":
                _, x, y, w, h, b64 = op
                try:
                    img = ImageReader(io.BytesIO(base64.b64decode(b64)))
                    c.drawImage(img, tx(x), ty(y + h), w * s, h * s, mask="auto", preserveAspectRatio=True)
                except Exception:
                    pass
            else:
                _, x, y, txt, size, bold, color, anchor, rot = op
                c.saveState()
                c.setFillColor(HexColor(color))
                c.setFont(fn_bold if bold else fn_reg, size * s)
                c.translate(tx(x), ty(y))
                if rot:
                    c.rotate(rot)
                if anchor == "middle":
                    c.drawCentredString(0, 0, txt)
                elif anchor == "end":
                    c.drawRightString(0, 0, txt)
                else:
                    c.drawString(0, 0, txt)
                c.restoreState()
        c.showPage()
    c.save()
    return buf.getvalue()


# =====================================================================
# 6. STYLE & KOMPONEN UI
# =====================================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Lexend:wght@300;400;500;600;700;800&display=swap');

:root{--navy:#1f2a7c;--blue:#2b5fd9;--ink:#1c2a6b;--muted:#5f6b9a;--line:#e3e9f6;--soft:#f5f8ff;--ok:#2e9e5b;--bad:#d64545;--warn:#f0a81c;
--shadow:0 6px 22px rgba(31,42,124,.10);--shadow-hi:0 14px 34px rgba(31,42,124,.20)}
:root, html, body { color-scheme: light !important; }
html{scroll-behavior:smooth}
html, body, [class*="css"], .stApp, .stMarkdown, .stButton, .stTextInput,
.stSelectbox, .stDateInput, .stNumberInput, .stRadio, .stCheckbox, .stDataFrame,
[data-testid="stSidebar"], [data-testid="stHeader"], [data-testid="stAppViewContainer"] {
    font-family: 'Lexend', 'Helvetica Neue', Helvetica, Arial, sans-serif !important;
    color-scheme: light !important;
}
.stApp{background:#ffffff !important;color:var(--ink)}
.stApp h1,.stApp h2,.stApp h3,.stApp h4{color:var(--navy) !important;font-weight:700;letter-spacing:.2px}
.stApp [data-testid="stWidgetLabel"] *{color:var(--ink) !important;font-weight:500}
.stApp [data-testid="stCaptionContainer"], .stApp [data-testid="stCaptionContainer"] *{color:var(--muted) !important}
[data-testid="stHeader"]{background:transparent !important}
.block-container{padding-top:1.4rem;max-width:100%;animation:fadeIn .5s ease both}
[data-testid="stToolbar"], [data-testid="stDecoration"], [data-testid="stStatusWidget"], [data-testid="stAppDeployButton"],
.stAppDeployButton, #MainMenu, footer, [data-testid="stMainMenu"]{display:none !important;visibility:hidden !important}
::-webkit-scrollbar{width:10px;height:10px}::-webkit-scrollbar-thumb{background:#c9d6f5;border-radius:8px}::-webkit-scrollbar-thumb:hover{background:#9db8ff}

/* ---------- animasi ---------- */
@keyframes fadeIn{from{opacity:0}to{opacity:1}}
@keyframes fadeUp{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}
@keyframes floaty{0%,100%{transform:translateY(0)}50%{transform:translateY(-6px)}}
@keyframes shimmer{from{background-position:0% 0}to{background-position:300% 0}}
@keyframes blob{0%,100%{transform:translate(0,0) scale(1)}50%{transform:translate(26px,18px) scale(1.12)}}
@keyframes dropIn{from{opacity:0;transform:translateY(-8px) scale(.98)}to{opacity:1;transform:none}}
@keyframes pulseRing{0%{box-shadow:0 0 0 0 rgba(214,69,69,.35)}70%{box-shadow:0 0 0 14px rgba(214,69,69,0)}100%{box-shadow:0 0 0 0 rgba(214,69,69,0)}}
@property --n{syntax:'<integer>';initial-value:0;inherits:false}
@keyframes cntup{from{--n:0}}
.cnt{counter-reset:num var(--n);animation:cntup 1.2s cubic-bezier(.2,.8,.2,1) both}
.cnt::after{content:counter(num) attr(data-s)}

/* ---------- header ---------- */
.hero{position:relative;overflow:hidden;background:#fff;border:1px solid var(--line);border-radius:22px;padding:30px 28px 24px;margin-bottom:22px;text-align:center;box-shadow:var(--shadow);animation:fadeUp .6s ease both}
.hero::before{content:'';position:absolute;left:0;right:0;top:0;height:5px;background:linear-gradient(90deg,#1f2a7c,#2b5fd9,#27b5c9,#ffb81c,#2b5fd9,#1f2a7c);background-size:300% 100%;animation:shimmer 6s linear infinite}
.hero .blob{position:absolute;border-radius:50%;filter:blur(38px);opacity:.35;pointer-events:none}
.hero .b1{width:220px;height:220px;background:#9db8ff;left:-60px;top:-80px;animation:blob 9s ease-in-out infinite}
.hero .b2{width:200px;height:200px;background:#25D366;right:-50px;bottom:-90px;animation:blob 11s ease-in-out infinite reverse} /* WHATSAPP GREEN COLOUR */
.hero > *:not(.blob){position:relative;z-index:1}
.hero img{height:56px;margin-bottom:14px;animation:floaty 4s ease-in-out infinite}
.hero .ttl{font-size:clamp(20px,2.9vw,34px);font-weight:800;color:#1f2a7c;line-height:1.25;letter-spacing:.2px;margin:0 auto 10px;max-width:980px}
.hero .sub{font-size:clamp(12px,1.4vw,15px);color:#5f6b9a;letter-spacing:.6px;margin-bottom:14px}
.hero .code{display:inline-block;font-size:12px;color:#1f2a7c;background:#f5f8ff;border:1px solid #e3e9f6;padding:5px 14px;border-radius:999px}
.hero-c{display:flex;align-items:center;gap:14px;background:#fff;border:1px solid var(--line);border-radius:16px;padding:12px 16px;margin-bottom:14px;box-shadow:var(--shadow);animation:fadeUp .5s ease both}
.hero-c img{height:36px;animation:floaty 4s ease-in-out infinite}
.hero-c .ttl{font-size:14px;font-weight:800;color:#1f2a7c;line-height:1.3}
.hero-c .sub{font-size:11px;color:#5f6b9a;margin-top:2px}

/* ---------- kartu KPI (melayang) ---------- */
.kp{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:14px;margin:8px 0 18px}
.kc{--c:#1f2a7c;background:#fff;border:1px solid var(--line);border-top:4px solid var(--c);border-radius:16px;padding:14px 18px;box-shadow:var(--shadow);
transition:transform .25s cubic-bezier(.2,.8,.2,1),box-shadow .25s;animation:fadeUp .55s ease both}
.kc:nth-child(2){animation-delay:.06s}.kc:nth-child(3){animation-delay:.12s}.kc:nth-child(4){animation-delay:.18s}.kc:nth-child(5){animation-delay:.24s}.kc:nth-child(6){animation-delay:.3s}
.kc:hover{transform:translateY(-6px);box-shadow:var(--shadow-hi)}
.kc .t{font-size:12px;color:#5f6b9a;font-weight:500}.kc .v{font-size:32px;font-weight:800;margin:4px 0 2px;color:#1f2a7c;line-height:1.1}.kc .s{font-size:11px;color:#5f6b9a}

/* ---------- kalender ---------- */
.cal{border-collapse:separate;border-spacing:2px;font-size:11px}
.cal td,.cal th{text-align:center;min-width:26px;height:26px;padding:0 2px;border-radius:6px;transition:transform .15s,box-shadow .15s}
.cal th{background:#eaf0ff;color:#1f2a7c;font-weight:600}
.cal td.l{text-align:left;font-weight:700;background:#f5f8ff;color:#1f2a7c;padding:0 10px;min-width:96px;white-space:nowrap;border-radius:8px}
.cal td:not(.l):hover{transform:scale(1.35);box-shadow:0 6px 14px rgba(31,42,124,.28);position:relative;z-index:2}
.s-ok{background:#2e9e5b;color:#fff}.s-high{background:#d64545;color:#fff;font-weight:700}.s-low{background:#f0a81c;color:#2b2000;font-weight:700}
.s-empty{background:#9aa5b1;color:#fff}.s-future{background:#eef1f8;color:#bbb}
.lg2{display:flex;gap:16px;font-size:12px;margin:10px 0;flex-wrap:wrap;color:#1c2a6b}.lg2 i{display:inline-block;width:14px;height:14px;border-radius:4px;margin-right:6px;vertical-align:-3px}

/* ---------- kotak, peringatan, floating ---------- */
.box{background:#fff;border:1px solid var(--line);border-radius:16px;padding:14px 18px;box-shadow:var(--shadow);margin-bottom:12px;font-size:13px;line-height:1.75;color:#1c2a6b;
transition:transform .25s cubic-bezier(.2,.8,.2,1),box-shadow .25s;animation:fadeUp .5s ease both}
.box:hover{transform:translateY(-3px);box-shadow:var(--shadow-hi)}
.alert{background:#fff4f4;border:1px solid #f5c6c6;border-left:6px solid #d64545;border-radius:14px;padding:12px 18px;margin:4px 0 14px;font-size:13px;line-height:1.75;color:#7a1f1f;
animation:fadeUp .5s ease both,pulseRing 2.4s ease-out infinite}
.hint{font-size:12px;margin:-8px 0 8px 2px}.hint.ok{color:#2e9e5b}.hint.bad{color:#d64545}
.empty{text-align:center;padding:46px 20px;background:#f5f8ff;border:1.5px dashed #c9d6f5;border-radius:22px;color:#1f2a7c;animation:fadeUp .5s ease both}
.empty .ic{font-size:56px;display:inline-block;animation:floaty 3.2s ease-in-out infinite}
.empty .h{font-size:20px;font-weight:700;margin:8px 0 4px}.empty .p{font-size:13px;color:#5f6b9a}
.fab{position:fixed;right:18px;bottom:76px;z-index:999;background:linear-gradient(135deg,#1f2a7c,#2b5fd9);color:#fff;border-radius:999px;padding:10px 18px;font-size:13px;font-weight:600;
box-shadow:0 10px 28px rgba(31,42,124,.38);animation:floaty 3.4s ease-in-out infinite;white-space:nowrap}
.fab small{opacity:.85;font-weight:400;margin-left:6px}

/* ---------- sidebar ---------- */
[data-testid="stSidebar"]{background:#ffffff !important;border-right:1px solid var(--line)}
[data-testid="stSidebarNavLink"], [data-testid="stSidebarNav"] a{border-radius:12px !important;margin:3px 8px !important;transition:background .18s,transform .18s,box-shadow .18s}
[data-testid="stSidebarNavLink"] *, [data-testid="stSidebarNav"] a *{color:#1f2a7c !important;font-weight:500}
[data-testid="stSidebarNavLink"]:hover, [data-testid="stSidebarNav"] a:hover{background:#eef3ff !important;transform:translateX(5px)}
[data-testid="stSidebarNavLink"][aria-current="page"], [data-testid="stSidebarNav"] a[aria-current="page"]{background:linear-gradient(90deg,#e6edff,#f5f8ff) !important;box-shadow:inset 4px 0 0 #1f2a7c}
[data-testid="stSidebarNavLink"][aria-current="page"] *{font-weight:700}

/* ---------- tombol ---------- */
.stButton>button, .stDownloadButton>button{border-radius:999px;border:1.5px solid var(--line);background:#fff;color:#1f2a7c;font-weight:600;padding:.45rem 1.2rem;
transition:transform .2s,box-shadow .2s,border-color .2s,color .2s;box-shadow:0 1px 2px rgba(31,42,124,.08)}
.stButton>button:hover, .stDownloadButton>button:hover{transform:translateY(-2px);box-shadow:0 10px 22px rgba(31,42,124,.18);border-color:#2b5fd9;color:#2b5fd9}
.stButton>button:active, .stDownloadButton>button:active{transform:translateY(0) scale(.98)}
.stButton>button[kind="primary"], [data-testid="stBaseButton-primary"]{background:linear-gradient(135deg,#1f2a7c,#2b5fd9) !important;border:none !important;color:#fff !important}
.stButton>button[kind="primary"] *, [data-testid="stBaseButton-primary"] *{color:#fff !important}
.stButton>button[kind="primary"]:hover, [data-testid="stBaseButton-primary"]:hover{box-shadow:0 12px 26px rgba(43,95,217,.45)}
.stButton>button:disabled, .stDownloadButton>button:disabled{opacity:.45;transform:none;box-shadow:none}

/* ---------- isian & dropdown (bullet drop) ---------- */
div[data-baseweb="input"], div[data-baseweb="textarea"], div[data-baseweb="select"] > div{border-radius:12px !important;background:#fff !important;border:1.5px solid var(--line) !important;
transition:border-color .18s,box-shadow .18s,transform .18s}
div[data-baseweb="input"] input { color: #1c2a6b !important; -webkit-text-fill-color: #1c2a6b !important; }
div[data-baseweb="input"]:hover, div[data-baseweb="textarea"]:hover, div[data-baseweb="select"] > div:hover{border-color:#9db8ff !important}
div[data-baseweb="input"]:focus-within, div[data-baseweb="textarea"]:focus-within, div[data-baseweb="select"] > div:focus-within{border-color:#2b5fd9 !important;box-shadow:0 0 0 4px rgba(43,95,217,.16)}
div[data-baseweb="select"] *{color:#1c2a6b}
div[data-baseweb="popover"] ul[role="listbox"]{border-radius:14px;border:1px solid var(--line);padding:6px;background:#fff;box-shadow:0 16px 38px rgba(31,42,124,.20);animation:dropIn .18s ease-out}
ul[role="listbox"]{list-style:none !important;margin:0}
li[role="option"]{list-style:none !important;border-radius:10px;margin:2px 0;color:#1c2a6b;transition:background .12s,padding-left .12s}
li[role="option"]:hover, li[role="option"][aria-selected="true"]{background:#eef3ff !important;color:#1f2a7c;padding-left:18px}
li[role="option"][aria-selected="true"]{font-weight:700}
div[role="radiogroup"] > label{background:#fff;border:1.5px solid var(--line);border-radius:999px;padding:6px 16px 6px 10px;margin-right:8px;cursor:pointer;transition:all .18s}
div[role="radiogroup"] > label:hover{border-color:#2b5fd9;transform:translateY(-1px);box-shadow:0 6px 14px rgba(31,42,124,.12)}
div[role="radiogroup"] > label:has(input:checked){background:linear-gradient(135deg,#1f2a7c,#2b5fd9);border-color:transparent}
div[role="radiogroup"] > label:has(input:checked) *{color:#fff !important}
label[data-baseweb="checkbox"]{padding:4px 8px;border-radius:10px;transition:background .15s}
label[data-baseweb="checkbox"]:hover{background:#f5f8ff}

/* ---------- tab, expander, tabel ---------- */
button[data-baseweb="tab"]{font-weight:600 !important; color:#5f6b9a !important; border-radius:12px 12px 0 0; transition:color .18s,background .18s}
button[data-baseweb="tab"] * { color:#5f6b9a !important; }
button[data-baseweb="tab"]:hover{color:#1f2a7c !important;background:#f0f4ff}
button[data-baseweb="tab"]:hover * { color:#1f2a7c !important; }
button[data-baseweb="tab"][aria-selected="true"]{color:#d64545 !important;}
button[data-baseweb="tab"][aria-selected="true"] * { color:#d64545 !important; }
div[data-baseweb="tab-highlight"]{background:#d64545 !important;height:3px !important;border-radius:3px}

/* PERBAIKAN WARNA EXPANDER (Author box agar terlihat) */
[data-testid="stExpander"]{border:1px solid var(--line) !important;border-radius:16px !important;background:#fff;box-shadow:var(--shadow);transition:box-shadow .25s}
[data-testid="stExpander"]:hover{box-shadow:var(--shadow-hi)}
[data-testid="stExpander"] summary, 
[data-testid="stExpander"] summary * { font-weight:600 !important; color: var(--navy) !important; }
[data-testid="stExpander"] summary svg { fill: var(--navy) !important; }

[data-testid="stDataFrame"]{border:1px solid var(--line);border-radius:14px;overflow:hidden;box-shadow:var(--shadow)}
[data-testid="stAlert"]{border-radius:14px}
@media (prefers-reduced-motion: reduce){*{animation:none !important;transition:none !important}}
</style>
""", unsafe_allow_html=True)


def header(compact=False):
    """Header (kelas CSS di blok style di atas). Teks dibuat satu baris agar tidak dianggap code-block oleh markdown."""
    logo = _logo_b64()
    img = "<img src='data:image/png;base64,{}'>".format(logo) if logo else ""
    if compact:   # versi ringkas untuk halaman QR (layar HP)
        st.markdown("<div class='hero-c'>{}<div><div class='ttl'>{}</div><div class='sub'>Instalasi Radiologi &middot; Siloam Hospitals Agora</div></div></div>".format(
            img, escape(APP_TITLE)), unsafe_allow_html=True)
        return
    st.markdown("<div class='hero'><span class='blob b1'></span><span class='blob b2'></span>{}<div class='ttl'>{}</div>"
                "<div class='sub'>Siloam Hospitals Agora &middot; Suhu &middot; Kelembaban &middot; QC Peralatan</div>"
                "<div class='code'>Instalasi Radiologi &nbsp;|&nbsp; {} &nbsp;&middot;&nbsp; RAD 2209.xx/00</div></div>".format(
                    img, escape(APP_TITLE), FORM_CODE_TH), unsafe_allow_html=True)


def month_year_bar(prefix, cols=(1.6, 1)):
    t = TODAY()
    c2, c3 = st.columns(list(cols))
    bln = c2.selectbox("Bulan", list(range(1, 13)), format_func=lambda x: BLN[x - 1], index=t.month - 1, key=prefix + "_bln")
    yrs = list(range(t.year - 2, t.year + 2))
    thn = c3.selectbox("Tahun", yrs, index=yrs.index(t.year), key=prefix + "_thn")
    return int(bln), int(thn)


def filter_bar(prefix, options, label):
    c1, c2, c3 = st.columns([2.2, 1.6, 1])
    pilih = c1.selectbox(label, options, key=prefix + "_pilih")
    t = TODAY()
    bln = c2.selectbox("Bulan", list(range(1, 13)), format_func=lambda x: BLN[x - 1], index=t.month - 1, key=prefix + "_bln")
    yrs = list(range(t.year - 2, t.year + 2))
    thn = c3.selectbox("Tahun", yrs, index=yrs.index(t.year), key=prefix + "_thn")
    return pilih, int(bln), int(thn)


def days_text(days):
    return ", ".join(str(d) for d in days) if days else "-"



# =====================================================================
# 6b. MODE AUTHOR (password di Streamlit Secrets: AUTHOR_PASSWORD)
# =====================================================================
def _author_secret():
    # =========================================================================
    # CARA MENGATUR PASSWORD DI STREAMLIT CLOUD SECRETS:
    # 1. Buka dashboard Streamlit Cloud Anda.
    # 2. Klik ikon titik tiga pada app -> Settings -> Secrets.
    # 3. Tambahkan baris berikut (ganti dengan password rahasia Anda):
    #    AUTHOR_PASSWORD = "password_rahasia_anda"
    # 4. Klik Save. Password default di bawah ini ("admin123") hanya digunakan jika Secrets belum diset.
    # =========================================================================
    try:
        v = st.secrets.get("AUTHOR_PASSWORD")
    except Exception:
        v = None
    return str(v) if v else os.environ.get("AUTHOR_PASSWORD", "admin123")


def is_author():
    return bool(st.session_state.get("is_author"))


def author_box():
    """Kotak login author di sidebar."""
    with st.sidebar.expander("🔐 Author"):
        if is_author():
            st.success("Mode author aktif")
            if st.button("Keluar", key="auth_out"):
                st.session_state["is_author"] = False
                st.rerun()
            return
        pw = st.text_input("Password author", type="password", key="auth_pw")
        if st.button("Masuk", key="auth_btn"):
            secret = _author_secret()
            if secret and hmac.compare_digest(pw.encode("utf-8"), secret.encode("utf-8")):
                st.session_state["is_author"] = True
                st.rerun()
            else:
                time.sleep(1)
                st.error("Password salah.")


# =====================================================================
# 7. HALAMAN
# =====================================================================
def dash_th():
    st.caption("Rangkuman pencapaian pemeriksaan suhu & kelembaban (3x sehari: 08.00 · 14.00 · 22.00)")
    ruang, m, y = filter_bar("db", room_names(), "Ruangan")
    an = analyze(ruang, y, m)
    std = an["std"]
    if an["due"] == 0:
        st.info("Belum ada slot pengukuran yang jatuh tempo pada bulan ini.")
    st.markdown("""
    <div class='kp'>
      <div class='kc'><div class='t'>Persentase pengisian</div><div class='v'><span class='cnt' style='--n:{pf:.0f}' data-s='%'></span></div><div class='s'>{fd} data terisi dari {d} slot jatuh tempo</div></div>
      <div class='kc' style='border-top-color:#2e9e5b'><div class='t'>Suhu dalam batas ({lo:g}&ndash;{hi:g} &deg;C)</div><div class='v' style='color:#2e9e5b'><span class='cnt' style='--n:{po:.0f}' data-s='%'></span></div><div class='s'>{o} dari {f} data terisi</div></div>
      <div class='kc' style='border-top-color:#2e9e5b'><div class='t'>Kelembaban dalam batas ({hl:g}&ndash;{hh:g} %)</div><div class='v' style='color:#2e9e5b'><span class='cnt' style='--n:{ph:.0f}' data-s='%'></span></div><div class='s'>{ho} dari {f} data terisi</div></div>
      <div class='kc' style='border-top-color:#d64545'><div class='t'>Suhu DI ATAS batas</div><div class='v' style='color:#d64545'><span class='cnt' style='--n:{hi_n}' data-s=''></span></div><div class='s'>slot pengukuran</div></div>
      <div class='kc' style='border-top-color:#f0a81c'><div class='t'>Suhu di bawah batas</div><div class='v' style='color:#d98e00'><span class='cnt' style='--n:{lo_n}' data-s=''></span></div><div class='s'>slot pengukuran</div></div>
      <div class='kc' style='border-top-color:#9aa5b1'><div class='t'>Slot kosong / belum diisi</div><div class='v' style='color:#667'><span class='cnt' style='--n:{e}' data-s=''></span></div><div class='s'>{fe} tanggal kosong total</div></div>
    </div>""".format(pf=an["pct_fill"], f=an["filled"], fd=an["filled_due"], d=an["due"], lo=std["t_lo"], hi=std["t_hi"], po=an["pct_ok"], o=an["ok"],
                     hl=std["h_lo"], hh=std["h_hi"], ph=an["pct_hum"], ho=an["hum_ok"],
                     hi_n=len(an["high"]), lo_n=len(an["low"]), e=an["n_empty"], fe=len(an["full_empty"])), unsafe_allow_html=True)

    head = "<tr><th></th>" + "".join("<th>{}</th>".format(d) for d in range(1, an["n"] + 1)) + "</tr>"
    body = ""
    for s in "PSM":
        body += "<tr><td class='l'>{} ({:02d}.00)</td>".format(SHIFT_NAME[s], SHIFT_HOUR[s])
        for d in range(1, an["n"] + 1):
            c = an["slots"][(d, s)]
            stt = c["state"]
            tip = {"ok": "{:g}°C normal", "high": "{:g}°C DI ATAS batas", "low": "{:g}°C di bawah batas"}.get(stt, "")
            tip = "Tgl {} {}: ".format(d, SHIFT_NAME[s]) + (tip.format(c["t"]) if tip else ("KOSONG" if stt == "empty" else "belum jatuh tempo"))
            txt = {"ok": "&#10003;", "high": "&#9650;", "low": "&#9660;", "empty": "&ndash;", "future": ""}[stt]
            body += "<td class='s-{}' title='{}'>{}</td>".format(stt, tip, txt)
        body += "</tr>"
    st.markdown("<div style='overflow-x:auto'><table class='cal'>{}{}</table></div>".format(head, body), unsafe_allow_html=True)
    st.markdown("<div class='lg2'><span><i class='s-ok'></i>Dalam batas</span><span><i class='s-high'></i>&#9650; Di atas batas</span><span><i class='s-low'></i>&#9660; Di bawah batas</span>"
                "<span><i class='s-empty'></i>Kosong (belum diisi)</span><span><i class='s-future'></i>Belum jatuh tempo</span></div>", unsafe_allow_html=True)

    hi_days = sorted(set(d for d, _, _ in an["high"]))
    lo_days = sorted(set(d for d, _, _ in an["low"]))
    emp_days = sorted(an["empty_by_day"].keys())
    st.markdown("<div class='box'><b>Ringkasan {} - {} {}</b><br>"
                "&#128308; Tanggal dengan suhu <b>di atas batas</b> ({:g} &deg;C): <b>{}</b><br>"
                "&#128992; Tanggal dengan suhu di bawah batas ({:g} &deg;C): <b>{}</b><br>"
                "&#9899; Tanggal yang ada <b>slot kosong</b>: <b>{}</b><br>"
                "&#9898; Tanggal <b>kosong total</b> (Pagi, Siang, Malam belum diisi): <b>{}</b></div>".format(
                    escape(ruang), BLN[m - 1], y, std["t_hi"], days_text(hi_days), std["t_lo"], days_text(lo_days), days_text(emp_days), days_text(an["full_empty"])),
                unsafe_allow_html=True)

    data = month_data(ruang, y, m)
    if data:
        per_day = {}
        for (d, _s), e in data.items():
            per_day.setdefault(d, []).append(e)
        dfc = pd.DataFrame({"Tanggal": sorted(per_day),
                            "Suhu rata-rata (°C)": [sum(x["suhu"] for x in per_day[d]) / len(per_day[d]) for d in sorted(per_day)],
                            "Kelembaban rata-rata (%)": [sum(x["hum"] for x in per_day[d]) / len(per_day[d]) for d in sorted(per_day)]}).set_index("Tanggal")
        g1, g2 = st.columns(2)
        g1.markdown("**Tren suhu harian (rata-rata)**")
        g1.line_chart(dfc[["Suhu rata-rata (°C)"]])
        g2.markdown("**Tren kelembaban harian (rata-rata)**")
        g2.line_chart(dfc[["Kelembaban rata-rata (%)"]])

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Suhu di atas batas**")
        st.dataframe(pd.DataFrame([dict(Tanggal=d, Shift=SHIFT_NAME[s], **{"Suhu (°C)": v, "Batas atas": std["t_hi"], "Selisih": round(v - std["t_hi"], 1)}) for d, s, v in an["high"]]
                                  or [dict(Tanggal="-", Shift="-", **{"Suhu (°C)": "-", "Batas atas": "-", "Selisih": "-"})]), hide_index=True)
        if an["low"]:
            st.markdown("**Suhu di bawah batas**")
            st.dataframe(pd.DataFrame([dict(Tanggal=d, Shift=SHIFT_NAME[s], **{"Suhu (°C)": v, "Batas bawah": std["t_lo"], "Selisih": round(v - std["t_lo"], 1)}) for d, s, v in an["low"]]), hide_index=True)
    with c2:
        st.markdown("**Tanggal & shift kosong**")
        st.dataframe(pd.DataFrame([dict(Tanggal=d, **{"Shift kosong": ", ".join(SHIFT_NAME[x] for x in v), "Status": "Kosong total" if len(v) == 3 else "Sebagian"}) for d, v in sorted(an["empty_by_day"].items())]
                                  or [dict(Tanggal="-", **{"Shift kosong": "Semua terisi", "Status": "-"})]), hide_index=True)
    with c3:
        st.markdown("**Kelembaban di luar batas ({:g}–{:g} %)**".format(std["h_lo"], std["h_hi"]))
        st.dataframe(pd.DataFrame([dict(Tanggal=d, Shift=SHIFT_NAME[s], **{"Kelembaban (%)": v, "Keterangan": "Di atas" if k == "high" else "Di bawah"}) for d, s, v, k in an["hum_bad"]]
                                  or [dict(Tanggal="-", Shift="-", **{"Kelembaban (%)": "-", "Keterangan": "-"})]), hide_index=True)

    st.markdown("**Rekap semua ruangan - {} {}**".format(BLN[m - 1], y))
    rows = []
    for rn in room_names():
        a = analyze(rn, y, m)
        rows.append({"Ruangan": rn, "% Pengisian": round(a["pct_fill"]), "% Suhu dalam batas": round(a["pct_ok"]), "% Kelembaban dalam batas": round(a["pct_hum"]),
                     "Suhu di atas batas": len(a["high"]), "Suhu di bawah batas": len(a["low"]), "Slot kosong": a["n_empty"],
                     "Tanggal kosong total": days_text(a["full_empty"])})
    st.dataframe(pd.DataFrame(rows), hide_index=True)


def dash_qc():
    st.caption("Rangkuman checklist kegiatan pemeliharaan harian peralatan radiologi")
    m, y = month_year_bar("qd")
    devs = devices()
    ans = {d["nama"]: qc_analyze(d, y, m) for d in devs}
    tot_counted = sum(a["counted"] for a in ans.values())
    tot_full = sum(a["full"] for a in ans.values())
    tot_idone = sum(a["item_done"] for a in ans.values())
    tot_itotal = sum(a["item_total"] for a in ans.values())
    tot_empty = sum(a["n_empty"] for a in ans.values())
    tot_part = sum(a["n_part"] for a in ans.values())
    if tot_counted == 0:
        st.info("Belum ada hari yang jatuh tempo / terisi pada bulan ini.")
    st.markdown("""
    <div class='kp'>
      <div class='kc'><div class='t'>Hari checklist lengkap</div><div class='v'><span class='cnt' style='--n:{pf:.0f}' data-s='%'></span></div><div class='s'>{full} dari {cnt} hari-alat jatuh tempo</div></div>
      <div class='kc' style='border-top-color:#2e9e5b'><div class='t'>Kegiatan terselesaikan</div><div class='v' style='color:#2e9e5b'><span class='cnt' style='--n:{pi:.0f}' data-s='%'></span></div><div class='s'>{idn} dari {itot} kegiatan harian</div></div>
      <div class='kc' style='border-top-color:#f0a81c'><div class='t'>Hari terisi sebagian</div><div class='v' style='color:#d98e00'><span class='cnt' style='--n:{part}' data-s=''></span></div><div class='s'>hari-alat belum lengkap</div></div>
      <div class='kc' style='border-top-color:#9aa5b1'><div class='t'>Hari kosong</div><div class='v' style='color:#667'><span class='cnt' style='--n:{emp}' data-s=''></span></div><div class='s'>hari-alat belum diisi</div></div>
    </div>""".format(pf=(tot_full / tot_counted * 100 if tot_counted else 0), full=tot_full, cnt=tot_counted,
                     pi=(tot_idone / tot_itotal * 100 if tot_itotal else 0), idn=tot_idone, itot=tot_itotal, part=tot_part, emp=tot_empty),
                unsafe_allow_html=True)

    n = calendar.monthrange(y, m)[1]
    head = "<tr><th></th>" + "".join("<th>{}</th>".format(d) for d in range(1, n + 1)) + "</tr>"
    body = ""
    cls = {"full": "ok", "part": "low", "empty": "empty", "future": "future"}
    txt = {"full": "&#10003;", "part": "&#189;", "empty": "&ndash;", "future": ""}
    for dev in devs:
        a = ans[dev["nama"]]
        body += "<tr><td class='l'>{}</td>".format(escape(dev["nama"]))
        for d in range(1, n + 1):
            c = a["days"][d]
            tip = "{} tgl {}: {}/{} kegiatan".format(escape(dev["nama"]), d, c["done"], c["total"]) if c["state"] != "future" else "belum jatuh tempo"
            body += "<td class='s-{}' title='{}'>{}</td>".format(cls[c["state"]], tip, txt[c["state"]])
        body += "</tr>"
    st.markdown("<div style='overflow-x:auto'><table class='cal'>{}{}</table></div>".format(head, body), unsafe_allow_html=True)
    st.markdown("<div class='lg2'><span><i class='s-ok'></i>&#10003; Lengkap</span><span><i class='s-low'></i>&#189; Sebagian</span>"
                "<span><i class='s-empty'></i>Kosong (belum diisi)</span><span><i class='s-future'></i>Belum jatuh tempo</span></div>", unsafe_allow_html=True)

    st.markdown("**Rekap semua peralatan - {} {}**".format(BLN[m - 1], y))
    st.dataframe(pd.DataFrame([{"Peralatan": d["nama"], "Hari lengkap": ans[d["nama"]]["full"], "Hari jatuh tempo": ans[d["nama"]]["counted"],
                                "% Hari lengkap": round(ans[d["nama"]]["pct_full"]), "% Kegiatan selesai": round(ans[d["nama"]]["pct_item"]),
                                "Hari sebagian": ans[d["nama"]]["n_part"], "Hari kosong": ans[d["nama"]]["n_empty"]} for d in devs]), hide_index=True)

    pilih = st.selectbox("Detail kegiatan yang belum dikerjakan - pilih peralatan", [d["nama"] for d in devs], key="qd_detail")
    a = ans[pilih]
    rows = [{"Tanggal": d, "Status": "Kosong" if s == "empty" else "Sebagian", "Kegiatan belum dikerjakan": "; ".join(ms) if s == "part" else "Semua kegiatan"} for d, s, ms in a["missing"]]
    st.dataframe(pd.DataFrame(rows or [{"Tanggal": "-", "Status": "-", "Kegiatan belum dikerjakan": "Semua terisi lengkap"}]), hide_index=True)


def page_dashboard():
    header()
    bad = abnormal_today()
    if bad:
        st.markdown("<div class='alert'><b>&#9888; Hubungi teknisi / maintenance</b> - pengukuran hari ini di luar batas normal:<br>{}</div>".format("<br>".join(bad)), unsafe_allow_html=True)
    t1, t2 = st.tabs(["🌡️ Suhu & Kelembaban", "🛠️ QC Peralatan"])
    with t1:
        dash_th()
    with t2:
        dash_qc()


HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]


def _show_msg(key):
    msg = st.session_state.pop(key, None)
    if msg:
        (st.success if msg[0] == "success" else st.error)(msg[1])
        if msg[0] == "success":
            st.toast("Data berhasil disimpan", icon="✅")


def _save_th(p, ruang, tgl, shift_key, petugas):
    ss = st.session_state
    std = room_std(ruang)
    suhu, hum = ss.get(p + "_suhu"), ss.get(p + "_hum")
    temuan = (ss.get(p + "_temuan") or "").strip()
    tindakan = ss.get(p + "_tindakan") or ""
    if suhu is None or hum is None:
        ss[p + "_msg"] = ("error", "Suhu dan kelembaban wajib diisi.")
        return
    abn = classify(suhu, std["t_lo"], std["t_hi"]) != "ok" or classify(hum, std["h_lo"], std["h_hi"]) != "ok"
    if abn and not temuan:
        ss[p + "_msg"] = ("error", "Kolom Temuan wajib diisi karena suhu/kelembaban di luar batas standar.")
        return
    save_entry(ruang=ruang, tanggal=tgl.isoformat(), shift=shift_key, suhu=suhu, kelembapan=hum, petugas=petugas,
               status="Abnormal" if abn else "Normal", temuan=temuan, tindakan=tindakan if abn else "")
    ss[p + "_msg"] = ("success", "Tersimpan: {} - {} - {} ({} °C / {} %).{}".format(
        ruang, tgl.strftime("%d/%m/%Y"), SHIFT_NAME[shift_key], fnum(suhu), fnum(hum), " Segera hubungi teknisi." if abn else ""))
    for k in ("_suhu", "_hum", "_temuan", "_timpa"):   # kosongkan isian agar tidak tersimpan dobel di jam berikutnya
        ss.pop(p + k, None)


def th_entry_form(ruang, tgl, shift_key, petugas, p):
    """Isian suhu & kelembaban (dipakai halaman QR dan halaman manual)."""
    std = room_std(ruang)
    _show_msg(p + "_msg")
    c3, c4 = st.columns(2)
    suhu = c3.number_input("Suhu (°C) [standar {:g}–{:g}] *".format(std["t_lo"], std["t_hi"]), value=None, min_value=0.0, max_value=60.0, step=0.1,
                           format="%.1f", placeholder="hasil baca termometer", key=p + "_suhu")
    hum = c4.number_input("Kelembaban (%) [standar {:g}–{:g}] *".format(std["h_lo"], std["h_hi"]), value=None, min_value=0.0, max_value=100.0, step=1.0,
                          format="%.0f", placeholder="hasil baca higrometer", key=p + "_hum")
    ts = classify(suhu, std["t_lo"], std["t_hi"]) if suhu is not None else None
    hs = classify(hum, std["h_lo"], std["h_hi"]) if hum is not None else None
    if ts:
        c3.markdown("<div class='hint {}'>{}</div>".format("ok" if ts == "ok" else "bad", "✓ Dalam batas standar" if ts == "ok" else "⚠ {} batas standar ({:g}–{:g} °C)".format("DI ATAS" if ts == "high" else "DI BAWAH", std["t_lo"], std["t_hi"])), unsafe_allow_html=True)
    if hs:
        c4.markdown("<div class='hint {}'>{}</div>".format("ok" if hs == "ok" else "bad", "✓ Dalam batas standar" if hs == "ok" else "⚠ {} batas standar ({:g}–{:g} %)".format("DI ATAS" if hs == "high" else "DI BAWAH", std["h_lo"], std["h_hi"])), unsafe_allow_html=True)
    abn = (ts is not None and ts != "ok") or (hs is not None and hs != "ok")
    if abn:
        st.error("Di luar batas normal - segera hubungi teknisi / maintenance.")
        st.text_area("Temuan ketidaksesuaian * (wajib karena di luar batas standar)", key=p + "_temuan")
        st.selectbox("Tindakan korektif *", ["Lapor Teknisi AC / Maintenance", "Pengaturan AC / Dehumidifier", "Lainnya"], key=p + "_tindakan")
    ada = load_entries()
    dup = len(ada[(ada["ruang"] == ruang) & (ada["tanggal"] == tgl.isoformat()) & (ada["shift"] == shift_key)]) > 0
    timpa = True
    if dup:
        st.warning("Data {} - {} - {} sudah ada.".format(ruang, tgl.strftime("%d/%m/%Y"), SHIFT_NAME[shift_key]))
        timpa = st.checkbox("Timpa data yang sudah ada", key=p + "_timpa")
    st.button("Simpan suhu & kelembaban", type="primary", key=p + "_save", disabled=(dup and not timpa),
              on_click=_save_th, args=(p, ruang, tgl, shift_key, petugas))


def _set_all(keys, val):
    for k in keys:
        st.session_state[k] = val


def _save_qc_cb(pk, alat, tgl, petugas, pairs):
    ss = st.session_state
    hasil = {it: True for k, it in pairs if ss.get(k)}
    catatan = (ss.get(pk + "_cat") or "").strip()
    if not hasil and not catatan and not any(auto_status(dev_by_name(alat), tgl).values()):
        ss[pk + "_msg"] = ("error", "Belum ada kegiatan yang dicentang.")
        return
    save_qc(alat, tgl.isoformat(), petugas, hasil, catatan)
    ss[pk + "_msg"] = ("success", "Checklist tersimpan: {} - {} ({} kegiatan dicentang manual).".format(alat, tgl.strftime("%d/%m/%Y"), len(hasil)))


def qc_checklist_ui(dev, tgl, petugas, p):
    """Daftar centang QC 1 alat 1 hari (dipakai halaman QR dan halaman manual)."""
    alat = dev["nama"]
    pk = "{}_{}_{}".format(p, alat, tgl.isoformat())
    _show_msg(pk + "_msg")
    existing = qc_get(alat, tgl)
    preset = existing["done"] if existing else set()
    if existing:
        st.info("Checklist {} tanggal {} sudah ada (staff {}). Menyimpan akan memperbarui data tersebut.".format(alat, tgl.strftime("%d/%m/%Y"), existing["petugas"] or "-"))
    auto = auto_status(dev, tgl)
    pairs = []   # (key, item) untuk kegiatan yang dicentang manual
    for i, it in enumerate(dev["items"]):
        if not (dev.get("ruang") and auto_kind(it)):
            pairs.append(("{}_i{}".format(pk, i), it))
    keys = [k for k, _ in pairs]
    b1, b2, _sp = st.columns([1, 1, 2])
    b1.button("Centang semua", on_click=_set_all, args=(keys, True), key=pk + "_all")
    b2.button("Kosongkan", on_click=_set_all, args=(keys, False), key=pk + "_none")
    kmap = dict((it, k) for k, it in pairs)
    for it in dev["items"]:
        if it in kmap:
            k = kmap[it]
            if k not in st.session_state:
                st.session_state[k] = it in preset
            st.checkbox(it, key=k)
        else:
            ok = auto[auto_kind(it)]
            st.markdown("{} {} - <span style='font-size:12px;color:{}'>{}</span>".format(
                "✅" if ok else "⬜", escape(it), "#2e9e5b" if ok else "#d64545",
                "terisi otomatis dari data Suhu & Kelembaban ({})".format(escape(dev["ruang"])) if ok else "isi dulu Suhu & Kelembaban ruangan ini"),
                unsafe_allow_html=True)
    st.text_area("Catatan / temuan (opsional)", value=existing["catatan"] if existing else "", key=pk + "_cat")
    st.button("Simpan checklist QC", type="primary", key=pk + "_save", on_click=_save_qc_cb, args=(pk, alat, tgl, petugas, pairs))


def page_ruang():
    """Halaman tujuan QR code: /ruang?ruang=NAMA -> isi suhu & kelembaban + QC untuk ruangan itu."""
    header(compact=True)
    names = room_names()
    qp = st.query_params.get("ruang", "")
    if isinstance(qp, list):
        qp = qp[0] if qp else ""
    ss = st.session_state
    if qp in names:
        ss["locked_room"] = qp    # scan QR -> ruangan terkunci selama sesi (tetap terkunci walau pindah menu lalu kembali)
    locked = ss.get("locked_room")
    if is_author():               # hanya author yang boleh memilih ruangan secara manual
        ruang = st.selectbox("Ruangan (mode author)", names, index=names.index(locked) if locked in names else 0, key="qr_pilih")
    elif qp and qp not in names:
        st.error("QR code tidak dikenali (ruangan '{}'). Scan ulang QR code di ruangan.".format(qp))
        return
    elif locked in names:
        ruang = locked
    else:
        st.markdown("<div class='empty'><div class='ic'>🔳</div><div class='h'>Scan QR code di ruangan</div>"
                    "<div class='p'>Arahkan kamera HP ke QR code yang ditempel di ruangan untuk membuka formulir pengisian ruangan tersebut.</div></div>", unsafe_allow_html=True)
        return
    n = now()
    tgl = n.date()
    data = month_data(ruang, tgl.year, tgl.month)
    badges = " &nbsp; ".join("{} {} ({:02d}.00)".format("✅" if (tgl.day, s) in data else "⬜", SHIFT_NAME[s], SHIFT_HOUR[s]) for s in "PSM")
    st.markdown("<div class='box'><span style='font-size:12px;color:#667'>Ruangan</span><br><span style='font-size:26px;font-weight:700;color:#0b5cab'>{}</span><br>"
                "<span style='font-size:12px;color:#556'>{}, {} {} {} &middot; {} WIB</span><br><span style='font-size:12px'>Sudah diisi hari ini: {}</span>{}</div>".format(
                    escape(ruang), HARI[tgl.weekday()], tgl.day, BLN[tgl.month - 1], tgl.year, n.strftime("%H:%M"), badges,
                    "" if is_author() else "<br><span style='font-size:11px;color:#667'>&#128274; Terkunci ke ruangan ini. Untuk ruangan lain, scan QR code ruangan tersebut.</span>"), unsafe_allow_html=True)
    petugas = st.selectbox("Nama staff (inisial) *", load_cfg()["staff"], key="qr_petugas")
    st.markdown("### 1. Suhu & Kelembaban")
    sk = shift_now()
    shift_key = st.radio("Jam pemantauan", list(SHIFT_LABEL.keys()), index=list("PSM").index(sk), format_func=lambda k: SHIFT_LABEL[k],
                         horizontal=True, key="qr_shift_" + sk)
    st.markdown("<div class='fab'>{} {} <small>{}</small></div>".format("🔓" if is_author() else "🔒", escape(ruang), SHIFT_NAME[shift_key]), unsafe_allow_html=True)
    th_entry_form(ruang, tgl, shift_key, petugas, "qr")
    st.markdown("### 2. Checklist QC Peralatan")
    devs = room_devices(ruang)
    if not devs:
        st.caption("Tidak ada checklist QC peralatan untuk ruangan ini.")
    for dev in devs:
        with st.expander("{} - {}".format(dev["nama"], dev["kode"]), expanded=(len(devs) == 1)):
            qc_checklist_ui(dev, tgl, petugas, "qr")


def YESTERDAY():
    return TODAY() - dt.timedelta(days=1)


def backdate_th():
    st.caption("Backdate suhu & kelembaban - hanya tanggal sebelum hari ini ({}). Pengisian real-time lewat QR code di ruangan.".format(FORM_CODE_TH))
    cfg = load_cfg()
    left, right = st.columns([2.2, 1], gap="large")
    with left:
        ruang = st.selectbox("Ruangan *", room_names(), key="in_ruang")
        c1, c2 = st.columns(2)
        tgl = c1.date_input("Tanggal backdate *", value=YESTERDAY(), max_value=YESTERDAY(), key="in_tgl")
        shift_key = c2.radio("Jam pemantauan *", list(SHIFT_LABEL.keys()), format_func=lambda k: SHIFT_LABEL[k], horizontal=True, key="in_shift")
        petugas = st.selectbox("Nama staff (inisial) *", cfg["staff"], key="in_petugas")
        th_entry_form(ruang, tgl, shift_key, petugas, "in")
    with right:
        st.markdown("<div class='box'><b>Panduan pengisian</b><br>Menu ini <b>hanya untuk backdate</b> (tanggal sebelum hari ini). Pengisian real-time dilakukan lewat <b>scan QR code di ruangan</b>.<br><br>"
                    "1. Pilih ruangan, tanggal, dan jam (08.00 · 14.00 · 22.00).<br>2. Isi suhu & kelembaban sesuai hasil baca alat.<br>3. Pilih inisial staff.<br>"
                    "4. Bila di luar batas standar, temuan & tindakan korektif wajib diisi dan segera hubungi teknisi.<br><br>"
                    "Suhu & kelembaban yang diinput otomatis menandai 2 kegiatan pertama pada checklist QC peralatan terkait.</div>", unsafe_allow_html=True)
    st.markdown("**Data terakhir - {}**".format(ruang))
    df = load_entries()
    df = df[df["ruang"] == ruang]
    if len(df):
        df = df.sort_values("id", key=lambda s: s.astype(int), ascending=False).head(15)
        st.dataframe(df[["id", "tanggal", "shift", "suhu", "kelembapan", "petugas", "status", "temuan"]], hide_index=True)
        cc1, cc2 = st.columns([1, 3])
        did = cc1.selectbox("Hapus data ID", list(df["id"]), key="in_del")
        if cc2.button("Hapus data terpilih", key="in_del_btn"):
            delete_entry(did)
            st.success("Data ID {} dihapus. Muat ulang halaman untuk memperbarui tabel.".format(did))
    else:
        st.caption("Belum ada data yang diinput untuk ruangan ini.")


def backdate_qc():
    st.caption("Backdate checklist QC - hanya tanggal sebelum hari ini. Pengisian real-time lewat QR code di ruangan.")
    cfg = load_cfg()
    left, right = st.columns([2.2, 1], gap="large")
    with left:
        alat = st.selectbox("Peralatan *", [d["nama"] for d in devices()], key="qi_alat")
        dev = dev_by_name(alat)
        c1, c2 = st.columns(2)
        tgl = c1.date_input("Tanggal backdate *", value=YESTERDAY(), max_value=YESTERDAY(), key="qi_tgl")
        petugas = c2.selectbox("Nama staff (inisial) *", cfg["staff"], key="qi_petugas")
        st.caption(" · ".join("{}: {}".format(a, b) for a, b in dev["info"][:4]) + "   |   " + dev["kode"])
        qc_checklist_ui(dev, tgl, petugas, "qi")
    with right:
        st.markdown("<div class='box'><b>Panduan pengisian</b><br>Menu ini <b>hanya untuk backdate</b> (tanggal sebelum hari ini). Pengisian real-time lewat <b>scan QR code di ruangan</b>.<br><br>1. Pilih peralatan dan tanggal.<br>"
                    "2. Centang kegiatan yang sudah dikerjakan (atau <i>Centang semua</i>).<br>"
                    "3. <i>Periksa Suhu/Kelembaban ruangan</i> terisi otomatis dari data Suhu & Kelembaban. Alat <i>Mobile</i> tidak terhubung ke ruang tertentu, jadi dicentang manual.<br>"
                    "4. Catatan tampil di bawah formulir.<br>5. <i>Flat field Test (Mingguan)</i> tidak dihitung sebagai kegiatan harian di dashboard.</div>", unsafe_allow_html=True)
    st.markdown("**Data terakhir - {}**".format(alat))
    df = load_qc()
    df = df[df["alat"] == alat]
    if len(df):
        df = df.sort_values("id", key=lambda s: s.astype(int), ascending=False).head(15).copy()

        def _n(h):
            try:
                return len([1 for v in json.loads(h).values() if v])
            except Exception:
                return 0
        df["jumlah dicentang"] = df["hasil"].map(_n)
        st.dataframe(df[["id", "tanggal", "petugas", "jumlah dicentang", "catatan"]], hide_index=True)
        cc1, cc2 = st.columns([1, 3])
        did = cc1.selectbox("Hapus data ID", list(df["id"]), key="qi_del")
        if cc2.button("Hapus data terpilih", key="qi_del_btn"):
            delete_qc(did)
            st.success("Data ID {} dihapus. Muat ulang halaman untuk memperbarui tabel.".format(did))
    else:
        st.caption("Belum ada checklist yang diinput untuk peralatan ini.")


def page_backdate():
    header()
    if not is_author():
        st.error("Halaman ini hanya untuk author.")
        return
    t1, t2 = st.tabs(["🌡️ Suhu & Kelembaban", "🛠️ QC Peralatan"])
    with t1:
        backdate_th()
    with t2:
        backdate_qc()


def _preview(sh, base_px=1400, scale=0.85): # Diperbesar dari 0.6 menjadi 0.85
    """Preview formulir dengan ukuran tetap (85%)."""
    wpx = int(base_px * scale)
    components.html("<div style='overflow:auto;background:#fff;border:1px solid #d5dbe3'>{}</div>".format(sheet_to_svg(sh, wpx)),
                    height=int(wpx * sh.height / sh.width) + 30, scrolling=True)


def page_download():
    header()
    names = ["🌡️ Form Suhu & Kelembaban", "🛠️ Form Checklist QC Peralatan"]
    if is_author():   # generator QR hanya untuk author
        names.append("🔳 QR Code Ruangan")
    tabs = st.tabs(names)
    t1, t2 = tabs[0], tabs[1]
    with t1:
        ruang, m, y = filter_bar("dl", room_names(), "Ruangan")
        sh = build_th_form(ruang, y, m, month_data(ruang, y, m))
        _preview(sh)
        fname = "Form_Suhu_Kelembaban_{}_{}_{:02d}".format(ruang.replace(" ", "_"), y, m)
        c1, c2, c3 = st.columns(3)
        try:
            c1.download_button("⬇️ PDF (ruangan ini)", sheets_to_pdf([sh], "Pemeriksaan Suhu & Kelembaban - " + ruang), fname + ".pdf", "application/pdf", key="dl_pdf")
            allsh = [build_th_form(rn, y, m, month_data(rn, y, m)) for rn in room_names()]
            c2.download_button("⬇️ PDF (semua ruangan)", sheets_to_pdf(allsh, "Pemeriksaan Suhu & Kelembaban - Semua Ruangan"),
                               "Form_Suhu_Kelembaban_SEMUA_{}_{:02d}.pdf".format(y, m), "application/pdf", key="dl_pdf_all")
        except ImportError:
            st.error("Library reportlab belum terpasang. Jalankan:  pip install reportlab")
        rows = [{"Ruang": ruang, "Tanggal": dt.date(y, m, d).isoformat(), "Jam": "{:02d}.00".format(SHIFT_HOUR[s]), "Shift": SHIFT_NAME[s],
                 "Suhu (°C)": e["suhu"], "Kelembaban (%)": e["hum"], "Staff": e["petugas"], "Temuan": e["ket"]}
                for (d, s), e in sorted(month_data(ruang, y, m).items())]
        c3.download_button("⬇️ Data (CSV)", pd.DataFrame(rows).to_csv(index=False).encode("utf-8-sig"), fname + ".csv", "text/csv", key="dl_csv")
        # st.caption("Format mengikuti {}: grafik titik suhu (16-27 °C) dan kelembaban (10-100 %), pita merah muda = batas normal, titik merah = di luar batas.".format(FORM_CODE_TH)) <-- DIHAPUS
    with t2:
        alat, m2, y2 = filter_bar("dq", [d["nama"] for d in devices()], "Peralatan")
        dev = dev_by_name(alat)
        sh2 = build_qc_form(dev, y2, m2, qc_month(dev, y2, m2))
        _preview(sh2)
        fname2 = "Checklist_QC_{}_{}_{:02d}".format(alat.replace(" ", "_"), y2, m2)
        q1, q2, q3 = st.columns(3)
        try:
            q1.download_button("⬇️ PDF (peralatan ini)", sheets_to_pdf([sh2], "Checklist Pemeliharaan - " + alat), fname2 + ".pdf", "application/pdf", key="dq_pdf")
            allq = [build_qc_form(d, y2, m2, qc_month(d, y2, m2)) for d in devices()]
            q2.download_button("⬇️ PDF (semua peralatan)", sheets_to_pdf(allq, "Checklist Pemeliharaan - Semua Peralatan"),
                               "Checklist_QC_SEMUA_{}_{:02d}.pdf".format(y2, m2), "application/pdf", key="dq_pdf_all")
        except ImportError:
            st.error("Library reportlab belum terpasang. Jalankan:  pip install reportlab")
        qd = qc_month(dev, y2, m2)
        rows2 = []
        for d in sorted(qd):
            e = qd[d]
            rows2.append({"Alat": alat, "Tanggal": dt.date(y2, m2, d).isoformat(), "Staff": e["petugas"],
                          "Kegiatan selesai": len([i for i in dev["items"] if i in e["done"]]), "Total kegiatan": len(dev["items"]),
                          "Belum dikerjakan": "; ".join(i for i in dev["items"] if i not in e["done"] and is_daily(i)), "Catatan": e["catatan"]})
        q3.download_button("⬇️ Data (CSV)", pd.DataFrame(rows2).to_csv(index=False).encode("utf-8-sig"), fname2 + ".csv", "text/csv", key="dq_csv")
        st.caption("Format mengikuti {} - centang = kegiatan selesai. Kegiatan suhu & kelembaban terisi otomatis dari data ruang terkait.".format(dev["kode"]))
    if is_author():
        with tabs[2]:
            download_qr_tab()




def download_qr_tab():
    if not is_author():
        return
    st.markdown("**QR Code per ruangan** - scan membuka halaman pengisian Suhu & Kelembaban + Checklist QC untuk ruangan tersebut ({} ruangan).".format(len(room_names())))
    base = st.text_input("Alamat aplikasi (URL Streamlit)", value=default_base_url(), key="qr_base").strip()
    if "NAMA-APP" in base or not base.startswith("http"):
        st.warning("Ganti dengan alamat asli aplikasi Anda, mis. https://shag-radiologi.streamlit.app, lalu cetak QR. Jangan cetak sebelum alamatnya benar.")
    if st.button("Simpan alamat ini", key="qr_base_save"):
        cfg = load_cfg()
        cfg["base_url"] = base.rstrip("/")
        save_cfg(cfg)
        st.success("Alamat disimpan.")
    sheets = build_qr_sheets(base)
    _preview(sheets[0], 900)
    st.download_button("⬇️ PDF stiker QR (semua ruangan, A4)", sheets_to_pdf(sheets, "QR Code Ruangan - Siloam Hospitals Agora", portrait=True),
                       "QR_Ruangan_SHAG.pdf", "application/pdf", key="qr_pdf")
    st.dataframe(pd.DataFrame([{"Ruangan": r, "Checklist QC": ", ".join(d["nama"] for d in room_devices(r)) or "-", "Link": room_url(base, r)} for r in room_names()]),
                 hide_index=True)
    st.caption("Cetak, laminating, tempel di dekat termometer/higrometer. Link bisa diuji dengan membuka salah satu alamat di tabel. "
               "Untuk MRI, tempel QR di luar Zona IV (pintu/ruang kontrol) - HP tidak boleh masuk ruang magnet.")


# =====================================================================
# 8. NAVIGASI & SIDEBAR
# =====================================================================
_pages = [(page_dashboard, "Dashboard", "📊", "dashboard", True),
          (page_ruang, "Isi Ruangan (QR)", "🔳", "ruang", False)]
if is_author():   # backdate manual (tanggal sebelum hari ini) hanya untuk author
    _pages.append((page_backdate, "Input Backdate (manual)", "🕘", "backdate", False))
_pages.append((page_download, "Download & QR" if is_author() else "Download Formulir", "⬇️", "download", False))
pg = st.navigation([st.Page(f, title="{}. {}".format(i + 1, t), icon=ic, url_path=u, default=d) for i, (f, t, ic, u, d) in enumerate(_pages)])

with st.sidebar:
    if is_author():
        with st.expander("Pengaturan ruang & staff"):
            cfg = load_cfg()
            nm = st.text_input("Nama ruang baru / ruang yang diubah", key="cfg_nm")
            c1, c2 = st.columns(2)
            tl = c1.number_input("Suhu min (°C)", value=20.0, step=0.5, key="cfg_tl")
            th = c2.number_input("Suhu maks (°C)", value=24.0, step=0.5, key="cfg_th")
            hl = c1.number_input("Kelembaban min (%)", value=30.0, step=1.0, key="cfg_hl")
            hh = c2.number_input("Kelembaban maks (%)", value=70.0, step=1.0, key="cfg_hh")
            if st.button("Simpan ruang", key="cfg_add"):
                if nm.strip():
                    cfg["rooms"] = [r for r in cfg["rooms"] if r["nama"] != nm.strip()] + [dict(nama=nm.strip(), t_lo=tl, t_hi=th, h_lo=hl, h_hi=hh, tekanan=False)]
                    save_cfg(cfg)
                    st.success("Ruang disimpan.")
            stf = st.text_input("Inisial staff (pisahkan koma)", value=",".join(cfg["staff"]), key="cfg_staff")
            if st.button("Simpan staff", key="cfg_staff_btn"):
                cfg["staff"] = [x.strip().upper() for x in stf.split(",") if x.strip()] or cfg["staff"]
                save_cfg(cfg)
                st.success("Staff disimpan.")
        with st.expander("Backup & restore data"):
            st.caption("Penyimpanan CSV di Streamlit Cloud bisa hilang saat aplikasi restart/redeploy. Backup rutin (mis. akhir bulan), lalu restore bila perlu.")
            for label, path, cols, k in (("Suhu & kelembaban", ENTRY_FILE, ECOLS, "th"), ("QC peralatan", QC_FILE, QCOLS, "qc")):
                st.download_button("⬇️ Backup {} (CSV)".format(label), _read(path, cols).to_csv(index=False).encode("utf-8-sig"),
                                   os.path.basename(path), "text/csv", key="bk_" + k)
                up = st.file_uploader("Restore {}".format(label), type="csv", key="up_" + k)
                if up is not None and st.button("Restore {} (menimpa data saat ini)".format(label), key="rs_" + k):
                    try:
                        new = pd.read_csv(up, dtype=str, encoding="utf-8-sig").fillna("")
                        miss = [c for c in cols if c not in new.columns]
                        if miss:
                            st.error("Kolom tidak lengkap: " + ", ".join(miss))
                        else:
                            os.makedirs(DATA_DIR, exist_ok=True)
                            new[cols].to_csv(path, index=False)
                            st.success("Data {} dipulihkan ({} baris).".format(label, len(new)))
                    except Exception as ex:
                        st.error("Gagal restore: {}".format(ex))
    author_box()
    # st.caption("{} · RAD 2209.xx/00".format(FORM_CODE_TH))  <-- DIHAPUS SESUAI PERMINTAAN

pg.run()

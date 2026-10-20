# DIGITAL MONITORING SHAG – Siloam Hospitals Agora (Radiologi)

Streamlit app: daily room temperature/humidity (form **SHKL 2205.11/2**) + equipment maintenance checklist (forms **RAD 2209.xx/00**), with dashboard and PDF export.

## Files to put in the GitHub repo (same layout as RAD-PHPK)
```
app.py
requirements.txt
Siloam Logo.png          <- your Siloam logo (exact name; if missing, a text logo is used)
Lexend-Regular.ttf       <- copy from the Primaya repo (auto-downloaded if missing)
Lexend-Bold.ttf
```
Deploy on Streamlit Community Cloud: New app -> pick the repo -> main file `app.py`.

## Pages
1. **Dashboard** – tab *Suhu & Kelembaban* (KPIs, calendar, trend, recap) and tab *QC Peralatan*. Red alert on top if any reading today is out of range.
2. **Input Suhu & Kelembaban** – 08.00 / 14.00 / 22.00, per room.
3. **Input QC Peralatan** – daily tick list per device.
4. **Download Formulir** – PDF (one room/device or all) + CSV.

## Edit things
- Rooms & limits, staff initials: sidebar -> *Pengaturan ruang & staff*.
- Devices / checklist items / serial numbers: `_DEVICES` near the top of `app.py`.
- Form code, timezone (WIB), hospital name: "KONFIGURASI UTAMA" block in `app.py`.

## Important: data storage
Data is saved as CSV in `data_shag/`. On Streamlit Cloud these files are **erased when the app restarts or is redeployed**. Use sidebar -> *Backup & restore data* regularly (e.g. weekly / month end), or move to a database later.

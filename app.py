from flask import Flask, jsonify, request
from flask_cors import CORS
import pandas as pd
from datetime import datetime, timedelta

app = Flask(__name__)
CORS(app)

FILE_PATH = "DATA DASHBOARD HRD PT ALDZAMA.xlsx"


# =========================
# LOAD & CLEAN DATA
# =========================
def load_data():
    try:
        df = pd.read_excel(
            FILE_PATH,
            sheet_name="INTERNSHIP",
            skiprows=2,
            usecols="B:M"
        )

        df.columns = [
            "nama", "program", "institusi", "kota", "ket",
            "permohonan", "penempatan", "durasi",
            "status", "final_project", "potensi", "berakhir"
        ]

        # CLEANING
        df['institusi'] = df['institusi'].astype(str).str.strip()
        df['status'] = df['status'].astype(str).str.strip().str.lower()
        df['penempatan'] = df['penempatan'].astype(str).str.strip()

        # FORMAT TANGGAL
        df['permohonan'] = pd.to_datetime(df['permohonan'], errors='coerce')
        df['bulan'] = df['permohonan'].dt.strftime('%B %Y')
        df['berakhir'] = pd.to_datetime(df['berakhir'], dayfirst=True, errors='coerce')

        return df

    except Exception as e:
        print("❌ ERROR LOAD DATA:", e)
        return pd.DataFrame()


# =========================
# FILTER FUNCTION
# =========================
def apply_filters(df):
    institusi = request.args.get("institusi")
    penempatan = request.args.get("penempatan")
    status = request.args.get("status")
    start = request.args.get("start")
    end = request.args.get("end")

    # FILTER INSTITUSI
    if institusi:
        df = df[df['institusi'] == institusi]

    # FILTER PENEMPATAN
    if penempatan:
        df = df[df['penempatan'] == penempatan]

    # FILTER STATUS
    if status:
        df = df[df['status'] == status.lower()]

    # FILTER TANGGAL PERMOHONAN
    if start and end:
        start_date = pd.to_datetime(start)
        end_date = pd.to_datetime(end)

        df = df[
            (df['permohonan'] >= start_date) &
            (df['permohonan'] <= end_date)
        ]

    return df


# =========================
# KPI
# =========================
def calculate_kpi(df):
    today = pd.to_datetime(datetime.today().date())
    batas = today + timedelta(days=14)

    akan_berakhir = df[
        (df['berakhir'].notna()) &
        (df['berakhir'] >= today) &
        (df['berakhir'] <= batas)
    ]

    return {
        "onboard": len(df[df['status'] == 'onboard']),
        "butuh_surat_balasan": len(df[df['status'] == 'butuh surat balasan']),
        "ajukan_ulang": len(df[df['status'] == 'ajukan ulang']),
        "selesai": len(df[df['status'] == 'selesai']),
        "akan_berakhir": len(akan_berakhir),
        "total": len(df)
    }


# =========================
# DASHBOARD (ALL IN ONE)
# =========================
@app.route("/dashboard")
def get_dashboard():
    df = load_data()

    # 🔥 APPLY FILTER DI SINI
    df = apply_filters(df)

    kpi = calculate_kpi(df)

    institusi = df['institusi'].value_counts().reset_index()
    institusi.columns = ['institusi', 'jumlah']

    permohonan = df['bulan'].value_counts().sort_index().reset_index()
    permohonan.columns = ['bulan', 'jumlah']

    penempatan = df['penempatan'].value_counts().reset_index()
    penempatan.columns = ['penempatan', 'jumlah']

    ket = df['ket'].value_counts().reset_index()
    ket.columns = ['ket', 'jumlah']

    return jsonify({
        "kpi": kpi,
        "institusi": institusi.to_dict(orient='records'),
        "permohonan": permohonan.to_dict(orient='records'),
        "penempatan": penempatan.to_dict(orient='records'),
        "ket": ket.to_dict(orient='records')
    })


if __name__ == "__main__":
    app.run(debug=True) 
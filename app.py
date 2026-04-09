from flask import Flask, jsonify, request
from flask_cors import CORS
import pandas as pd
from datetime import datetime, timedelta

app = Flask(__name__)
CORS(app)

FILE_PATH = "DATA DASHBOARD HRD PT ALDZAMA.xlsx"



# LOAD INTERNSHIP DATA

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

        df['institusi'] = df['institusi'].astype(str).str.strip()
        df['status'] = df['status'].astype(str).str.strip().str.lower()
        df['penempatan'] = df['penempatan'].astype(str).str.strip()

        df['permohonan'] = pd.to_datetime(df['permohonan'], errors='coerce')
        df['bulan'] = df['permohonan'].dt.strftime('%B %Y')
        df['berakhir'] = pd.to_datetime(df['berakhir'], dayfirst=True, errors='coerce')

        return df

    except Exception as e:
        print("❌ ERROR LOAD INTERNSHIP:", e)
        return pd.DataFrame()



# LOAD MANPOWER DATA

def load_manpower():
    try:
        df = pd.read_excel(
            FILE_PATH,
            sheet_name="MAN POWER",
            usecols="B:H"
        )

        df.columns = [
            "nama",
            "jenis_kelamin",
            "jabatan",
            "branch",
            "group_project",
            "ring",
            "status_kontrak"
        ]

        # CLEANING
        df['jabatan'] = df['jabatan'].astype(str).str.strip()
        df['branch'] = df['branch'].astype(str).str.strip()
        df['group_project'] = df['group_project'].astype(str).str.strip()
        df['status_kontrak'] = df['status_kontrak'].astype(str).str.strip().str.lower()

        return df

    except Exception as e:
        print("❌ ERROR LOAD MANPOWER:", e)
        return pd.DataFrame()



# LOAD RECRUITMENT DATA

def load_recruitment():
    try:
        sheets = [
            "RECRUITMENT",
            "RECRUITMENT OKT",
            "RECRUITMENT FEB 2026"
        ]

        df_list = []

        for sheet in sheets:
            df = pd.read_excel(
                FILE_PATH,
                sheet_name=sheet,
                usecols="B:E"
            )

            df.columns = [
                "nama",
                "posisi",
                "periode",
                "status"
            ]

            df_list.append(df)

        # GABUNG SEMUA DATA
        df = pd.concat(df_list, ignore_index=True)

        # CLEANING
        df['posisi'] = df['posisi'].astype(str).str.strip()
        df['status'] = df['status'].astype(str).str.strip().str.lower()

        # NORMALISASI STATUS
        df['status'] = df['status'].replace({
            'on process user': 'on process',
            'on process hr': 'on process'
        })

        # FORMAT TANGGAL
        df['periode'] = pd.to_datetime(df['periode'], errors='coerce')
        df['bulan'] = df['periode'].dt.strftime('%B %Y')

        return df

    except Exception as e:
        print("❌ ERROR LOAD RECRUITMENT:", e)
        return pd.DataFrame()




# FILTER (INTERNSHIP)

def apply_filters(df):
    institusi = request.args.get("institusi")
    penempatan = request.args.get("penempatan")
    status = request.args.get("status")
    start = request.args.get("start")
    end = request.args.get("end")

    if institusi:
        df = df[df['institusi'] == institusi]

    if penempatan:
        df = df[df['penempatan'] == penempatan]

    if status:
        df = df[df['status'] == status.lower()]

    if start and end:
        df = df[
            (df['permohonan'] >= pd.to_datetime(start)) &
            (df['permohonan'] <= pd.to_datetime(end))
        ]

    return df


# FILTER MANPOWER

def apply_manpower_filters(df):
    branch = request.args.get("branch")
    jabatan = request.args.get("jabatan")
    status_kontrak = request.args.get("status_kontrak")
    group_project = request.args.get("group_project")

    if branch:
        df = df[df['branch'] == branch]

    if jabatan:
        df = df[df['jabatan'] == jabatan]

    if status_kontrak:
        df = df[df['status_kontrak'] == status_kontrak.lower()]

    if group_project:
        df = df[df['group_project'] == group_project]

    return df


# FILTER RECRUITMENT

def apply_recruitment_filters(df):
    status = request.args.get("status")
    posisi = request.args.get("posisi")
    start = request.args.get("start")
    end = request.args.get("end")

    # FILTER STATUS
    if status:
        df = df[df['status'] == status.lower()]

    # FILTER POSISI
    if posisi:
        df = df[df['posisi'] == posisi]

    # FILTER TANGGAL
    if start and end:
        df = df[
            (df['periode'] >= pd.to_datetime(start)) &
            (df['periode'] <= pd.to_datetime(end))
        ]

    return df


# KPI INTERNSHIP

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



# KPI MANPOWER

def calculate_manpower_kpi(df):
    return {
        "total": len(df),
        "permanent": len(df[df['status_kontrak'] == 'permanent']),
        "kontrak": len(df[df['status_kontrak'] == 'kontrak'])
    }


# KPI RECRUITMENT

def calculate_recruitment_kpi(df):
    return {
        "total": len(df),
        "accepted": len(df[df['status'] == 'accepted']),
        "rejected": len(df[df['status'] == 'rejected']),
        "on_process": len(df[df['status'] == 'on process'])
    }

# DASHBOARD INTERNSHIP

@app.route("/internship/dashboard")
def get_dashboard():
    df = load_data()
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

# DASHBOARD MAN POWER

@app.route("/manpower/dashboard")
def manpower_dashboard():
    df = load_manpower()

    # 🔥 APPLY FILTER DI SINI
    df = apply_manpower_filters(df)

    kpi = calculate_manpower_kpi(df)

    jabatan = df['jabatan'].value_counts().reset_index()
    jabatan.columns = ['jabatan', 'jumlah']

    branch = df['branch'].value_counts().reset_index()
    branch.columns = ['branch', 'jumlah']

    group_project = df['group_project'].value_counts().reset_index()
    group_project.columns = ['group_project', 'jumlah']

    status_kontrak = df['status_kontrak'].value_counts().reset_index()
    status_kontrak.columns = ['status_kontrak', 'jumlah']

    return jsonify({
        "kpi": kpi,
        "jabatan": jabatan.to_dict(orient='records'),
        "branch": branch.to_dict(orient='records'),
        "group_project": group_project.to_dict(orient='records'),
        "status_kontrak": status_kontrak.to_dict(orient='records')
    })


# DASHBOARD RECRUITMENT

@app.route("/recruitment/dashboard")
def recruitment_dashboard():
    df = load_recruitment()

    # 🔥 APPLY FILTER DI SINI
    df = apply_recruitment_filters(df)

    # KPI
    kpi = calculate_recruitment_kpi(df)

    # STATUS (PIE)
    status = df['status'].value_counts().reset_index()
    status.columns = ['status', 'jumlah']

    # POSISI (BAR)
    posisi = df['posisi'].value_counts().reset_index()
    posisi.columns = ['posisi', 'jumlah']

    # TREND (LINE)
    trend = df['bulan'].value_counts().sort_index().reset_index()
    trend.columns = ['bulan', 'jumlah']

    # FUNNEL
    funnel_order = ['on process', 'mcu', 'accepted', 'rejected', 'mengundurkan diri']
    funnel_data = df['status'].value_counts()

    funnel = []
    for f in funnel_order:
        funnel.append({
            "status": f,
            "jumlah": int(funnel_data.get(f, 0))
        })

    return jsonify({
        "kpi": kpi,
        "status": status.to_dict(orient='records'),
        "posisi": posisi.to_dict(orient='records'),
        "trend": trend.to_dict(orient='records'),
        "funnel": funnel
    })


# RUN

if __name__ == "__main__":
    app.run(debug=True)
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
    


# LOAD SALARY & PPH21

def load_salary():
    try:
        df_gaji = pd.read_excel(
            FILE_PATH,
            sheet_name="GAJI DAN PPH21",
            usecols="B:D",
            skiprows=2
        )

        df_gaji.columns = ["project", "periode", "gaji"]

        df_gaji['project'] = df_gaji['project'].astype(str).str.strip()
        df_gaji['periode'] = df_gaji['periode'].astype(str).str.strip()
        df_gaji['gaji'] = pd.to_numeric(df_gaji['gaji'], errors='coerce')

        df_pph = pd.read_excel(
            FILE_PATH,
            sheet_name="GAJI DAN PPH21",
            usecols="G:I",
            skiprows=2
        )

        df_pph.columns = ["project", "periode", "pph21"]

        df_pph['project'] = df_pph['project'].astype(str).str.strip()
        df_pph['periode'] = df_pph['periode'].astype(str).str.strip()
        df_pph['pph21'] = pd.to_numeric(df_pph['pph21'], errors='coerce')

        return df_gaji, df_pph

    except Exception as e:
        print("❌ ERROR LOAD SALARY:", e)
        return pd.DataFrame(), pd.DataFrame()
    

# LOAD TRAINING DATA

def load_training():
    try:
        df = pd.read_excel(
            FILE_PATH,
            sheet_name="INTERNAL & EXTERNAL TRAINING FE",
            usecols="B:H"
        )

        df.columns = [
            "jenis",
            "nama",
            "divisi",
            "pelatihan",
            "lembaga",
            "periode",
            "status"
        ]

        # CLEANING
        df['jenis'] = df['jenis'].astype(str).str.strip().str.lower()
        df['divisi'] = df['divisi'].astype(str).str.strip()
        df['pelatihan'] = df['pelatihan'].astype(str).str.strip()
        df['status'] = df['status'].astype(str).str.strip().str.lower()

        # FORMAT TANGGAL
        df['periode'] = pd.to_datetime(df['periode'], errors='coerce')
        df['bulan'] = df['periode'].dt.strftime('%B %Y')

        return df

    except Exception as e:
        print("❌ ERROR LOAD TRAINING:", e)
        return pd.DataFrame()
    


# LOAD BPJS DATA

def load_bpjs():
    try:
        
        # KARYAWAN
        
        df_karyawan = pd.read_excel(
            FILE_PATH,
            sheet_name="BPJS",
            usecols="B:D"
        )

        df_karyawan.columns = ["nama", "project", "jenis_bpjs"]

        df_karyawan['project'] = df_karyawan['project'].astype(str).str.strip()
        df_karyawan['jenis_bpjs'] = df_karyawan['jenis_bpjs'].astype(str).str.strip().str.lower()

        
        # BPJS KESEHATAN
        
        df_kesehatan = pd.read_excel(
            FILE_PATH,
            sheet_name="BPJS",
            usecols="G:H"
        )

        df_kesehatan.columns = ["periode", "pembayaran"]

        
        # BPJS TK
        
        df_tk1 = pd.read_excel(
            FILE_PATH,
            sheet_name="PAY BPJS TK",
            usecols="C:D"
        )

        df_tk1.columns = ["periode", "pembayaran"]

        df_tk2 = pd.read_excel(
            FILE_PATH,
            sheet_name="PAY BPJS TK",
            usecols="F:G"
        )

        df_tk2.columns = ["periode", "pembayaran"]

        
        # CLEAN FUNCTION
        
        def clean_uang(df):
            df['periode'] = df['periode'].astype(str).str.strip()

            df['pembayaran'] = (
                df['pembayaran']
                .astype(str)
                .str.replace('.', '', regex=False)
            )

            df['pembayaran'] = pd.to_numeric(df['pembayaran'], errors='coerce')

            return df.dropna(subset=['pembayaran'])

        df_kesehatan = clean_uang(df_kesehatan)
        df_tk1 = clean_uang(df_tk1)
        df_tk2 = clean_uang(df_tk2)

        # GABUNG BPJS TK
        df_tk = pd.concat([df_tk1, df_tk2], ignore_index=True)

        # TAMBAH LABEL
        df_kesehatan['jenis'] = 'kesehatan'
        df_tk['jenis'] = 'tk'

        return df_karyawan, df_kesehatan, df_tk

    except Exception as e:
        print("❌ ERROR LOAD BPJS:", e)
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()


# =========================
# LOAD OVERTIME DATA
# =========================
def load_overtime():
    try:
        # OVERTIME
        df_overtime = pd.read_excel(
            FILE_PATH,
            sheet_name="OVERTIME DAN ABSENSI",
            usecols="B:D",
            skiprows=2
        )
        df_overtime.columns = ["project", "bulan", "overtime"]

        # ABSENSI
        df_absensi = pd.read_excel(
            FILE_PATH,
            sheet_name="OVERTIME DAN ABSENSI",
            usecols="F:G",
            skiprows=2
        )
        df_absensi.columns = ["periode", "absensi"]

        # COST
        df_cost = pd.read_excel(
            FILE_PATH,
            sheet_name="OVERTIME DAN ABSENSI",
            usecols="I:M",
            skiprows=2
        )
        df_cost.columns = [
            "project",
            "bulan",
            "total_cost",
            "overtime_cost",
            "overtime_percent"
        ]

        return df_overtime, df_absensi, df_cost

    except Exception as e:
        print("❌ ERROR LOAD OVERTIME:", e)
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()


# =========================
# CLEANING
# =========================
def clean_overtime(df_overtime, df_absensi, df_cost):

    # ===== OVERTIME =====
    df_overtime = df_overtime.dropna(subset=['overtime'])
    df_overtime['bulan'] = pd.to_datetime(df_overtime['bulan'], errors='coerce')

    # ===== ABSENSI =====
    df_absensi = df_absensi.dropna(subset=['absensi'])
    df_absensi['periode'] = pd.to_datetime(df_absensi['periode'], errors='coerce')

    # AGGREGATE PER BULAN
    df_absensi['bulan'] = df_absensi['periode'].dt.to_period('M').astype(str)
    absensi_bulanan = df_absensi.groupby('bulan')['absensi'].mean().reset_index()

    # ===== COST =====
    df_cost['bulan'] = pd.to_datetime(df_cost['bulan'], errors='coerce')

    # CLEAN ANGKA
    for col in ['total_cost', 'overtime_cost']:
        df_cost[col] = (
            df_cost[col]
            .astype(str)
            .str.replace(',', '', regex=False)
        )
        df_cost[col] = pd.to_numeric(df_cost[col], errors='coerce')

    df_cost = df_cost.dropna(subset=['total_cost'])

    return df_overtime, absensi_bulanan, df_cost



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


# FILTER SALARY

def apply_salary_filters(df, column_periode):
    periode = request.args.get("periode")
    project = request.args.get("project")

    if periode:
        df = df[df[column_periode] == periode]

    if project:
        df = df[df['project'] == project]

    return df


# FILTER TRAINING

def apply_training_filters(df):
    jenis = request.args.get("jenis")
    divisi = request.args.get("divisi")
    status = request.args.get("status")
    start = request.args.get("start")
    end = request.args.get("end")

    if jenis:
        df = df[df['jenis'] == jenis.lower()]

    if divisi:
        df = df[df['divisi'] == divisi]

    if status:
        df = df[df['status'] == status.lower()]

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


# FILTER BPJS

def apply_bpjs_filters(df_karyawan, df_kesehatan, df_tk):
    project = request.args.get("project")
    jenis_bpjs = request.args.get("jenis_bpjs")
    periode = request.args.get("periode")

    # FILTER KARYAWAN
    if project:
        df_karyawan = df_karyawan[df_karyawan['project'] == project]

    if jenis_bpjs:
        df_karyawan = df_karyawan[df_karyawan['jenis_bpjs'] == jenis_bpjs.lower()]

    # FILTER PERIODE (FINANSIAL)
    if periode:
        df_kesehatan = df_kesehatan[df_kesehatan['periode'] == periode]
        df_tk = df_tk[df_tk['periode'] == periode]

    return df_karyawan, df_kesehatan, df_tk



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


# KPI SALARY

def calculate_salary_kpi(df_gaji, df_pph):
    return {
        "total_gaji": int(df_gaji['gaji'].sum()),
        "total_pph21": int(df_pph['pph21'].sum()),
        "jumlah_project": df_gaji['project'].nunique()
    }


# KPI TRAINING

def calculate_training_kpi(df):
    return {
        "total": len(df),
        "internal": len(df[df['jenis'] == 'internal']),
        "external": len(df[df['jenis'] == 'external']),
        "done": len(df[df['status'] == 'done']),
        "in_progress": len(df[df['status'] == 'in progress']),
        "cancel": len(df[df['status'] == 'cancel'])
    }


# KPI BPJS

def calculate_bpjs_kpi(df_karyawan, df_kesehatan, df_tk):
    total_kesehatan = df_kesehatan['pembayaran'].sum()
    total_tk = df_tk['pembayaran'].sum()

    return {
        "total_karyawan": len(df_karyawan),
        "total_kesehatan": int(total_kesehatan),
        "total_tk": int(total_tk),
        "total_semua": int(total_kesehatan + total_tk)
    }

# =========================
# KPI
# =========================
def calculate_overtime_kpi(df_overtime, df_absensi, df_cost):
    return {
        "avg_overtime": float(df_overtime['overtime'].mean()),
        "avg_absensi": float(df_absensi['absensi'].mean()),
        "total_cost": int(df_cost['total_cost'].sum()),
        "overtime_cost": int(df_cost['overtime_cost'].sum())
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


# DASHBOARD SALARY

@app.route("/salary/dashboard")
def salary_dashboard():
    df_gaji, df_pph = load_salary()

    # APPLY FILTER
    df_gaji = apply_salary_filters(df_gaji, "periode")
    df_pph = apply_salary_filters(df_pph, "periode")

    # KPI
    kpi = calculate_salary_kpi(df_gaji, df_pph)

    # BAR GAJI PER PROJECT
    gaji_project = df_gaji.groupby('project')['gaji'].sum().reset_index()

    # LINE TREND GAJI
    gaji_trend = df_gaji.groupby('periode')['gaji'].sum().reset_index()

    # BAR PPH21
    pph_project = df_pph.groupby('project')['pph21'].sum().reset_index()

    # TREND PPH21
    pph_trend = df_pph.groupby('periode')['pph21'].sum().reset_index()

    return jsonify({
        "kpi": kpi,
        "gaji_project": gaji_project.to_dict(orient='records'),
        "gaji_trend": gaji_trend.to_dict(orient='records'),
        "pph_project": pph_project.to_dict(orient='records'),
        "pph_trend": pph_trend.to_dict(orient='records')
    })


# DASHBOARD TRAINING

@app.route("/training/dashboard")
def training_dashboard():
    df = load_training()

    # 🔥 APPLY FILTER
    df = apply_training_filters(df)

    # KPI
    kpi = calculate_training_kpi(df)

    # JENIS (PIE)
    jenis = df['jenis'].value_counts().reset_index()
    jenis.columns = ['jenis', 'jumlah']

    # DIVISI (BAR)
    divisi = df['divisi'].value_counts().reset_index()
    divisi.columns = ['divisi', 'jumlah']

    # PELATIHAN (BAR)
    pelatihan = df['pelatihan'].value_counts().reset_index()
    pelatihan.columns = ['pelatihan', 'jumlah']

    # STATUS (PIE)
    status = df['status'].value_counts().reset_index()
    status.columns = ['status', 'jumlah']

    # TREND (LINE)
    trend = df['bulan'].value_counts().sort_index().reset_index()
    trend.columns = ['bulan', 'jumlah']

    return jsonify({
        "kpi": kpi,
        "jenis": jenis.to_dict(orient='records'),
        "divisi": divisi.to_dict(orient='records'),
        "pelatihan": pelatihan.to_dict(orient='records'),
        "status": status.to_dict(orient='records'),
        "trend": trend.to_dict(orient='records')
    })


# DASHBOARD BPJS

@app.route("/bpjs/dashboard")
def bpjs_dashboard():
    df_karyawan, df_kesehatan, df_tk = load_bpjs()

    
    df_karyawan, df_kesehatan, df_tk = apply_bpjs_filters(
        df_karyawan, df_kesehatan, df_tk
    )

    # KPI
    kpi = calculate_bpjs_kpi(df_karyawan, df_kesehatan, df_tk)

    
    # DISTRIBUSI BPJS
    
    jenis = df_karyawan['jenis_bpjs'].value_counts().reset_index()
    jenis.columns = ['jenis_bpjs', 'jumlah']

    
    # PROJECT
    
    project = df_karyawan['project'].value_counts().reset_index()
    project.columns = ['project', 'jumlah']

    
    # TREND KESEHATAN
    
    kesehatan_trend = df_kesehatan.groupby('periode')['pembayaran'].sum().reset_index()

    
    # TREND TK
    
    tk_trend = df_tk.groupby('periode')['pembayaran'].sum().reset_index()

    
    # TOTAL PER BULAN
    
    df_all = pd.concat([df_kesehatan, df_tk])
    total_trend = df_all.groupby('periode')['pembayaran'].sum().reset_index()

    return jsonify({
        "kpi": kpi,
        "jenis": jenis.to_dict(orient='records'),
        "project": project.to_dict(orient='records'),
        "kesehatan_trend": kesehatan_trend.to_dict(orient='records'),
        "tk_trend": tk_trend.to_dict(orient='records'),
        "total_trend": total_trend.to_dict(orient='records')
    })

# =========================
# DASHBOARD OVERTIME
# =========================
@app.route("/overtime/dashboard")
def overtime_dashboard():

    df_overtime, df_absensi, df_cost = load_overtime()

    # CLEAN
    df_overtime, df_absensi, df_cost = clean_overtime(
        df_overtime, df_absensi, df_cost
    )

    # KPI
    kpi = calculate_overtime_kpi(df_overtime, df_absensi, df_cost)

    # =====================
    # OVERTIME TREND (LINE)
    # =====================
    overtime_trend = df_overtime.copy()
    overtime_trend['bulan'] = overtime_trend['bulan'].dt.strftime('%Y-%m')
    overtime_trend = overtime_trend.groupby('bulan')['overtime'].mean().reset_index()

    # =====================
    # ABSENSI TREND (LINE)
    # =====================
    absensi_trend = df_absensi

    # =====================
    # OVERTIME PER PROJECT (BAR)
    # =====================
    overtime_project = df_overtime.groupby('project')['overtime'].mean().reset_index()

    # =====================
    # OVERTIME COST (BAR)
    # =====================
    overtime_cost = df_cost.groupby('bulan')['overtime_cost'].sum().reset_index()
    overtime_cost['bulan'] = overtime_cost['bulan'].dt.strftime('%Y-%m')

    # =====================
    # TOP 5 PROJECT (BOTTOM)
    # =====================
    top_project = (
        df_overtime.groupby('project')['overtime']
        .mean()
        .sort_values(ascending=False)
        .head(5)
        .reset_index()
    )

    return jsonify({
        "kpi": kpi,
        "overtime_trend": overtime_trend.to_dict(orient='records'),
        "absensi_trend": absensi_trend.to_dict(orient='records'),
        "overtime_project": overtime_project.to_dict(orient='records'),
        "overtime_cost": overtime_cost.to_dict(orient='records'),
        "top_project": top_project.to_dict(orient='records')
    })


# RUN

if __name__ == "__main__":
    app.run(debug=True)
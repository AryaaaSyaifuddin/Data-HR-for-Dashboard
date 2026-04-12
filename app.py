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

        df = pd.concat(df_list, ignore_index=True)

        df['posisi'] = df['posisi'].astype(str).str.strip()
        df['status'] = df['status'].astype(str).str.strip().str.lower()

        df['status'] = df['status'].replace({
            'on process user': 'on process',
            'on process hr': 'on process'
        })

        df['periode'] = pd.to_datetime(df['periode'], errors='coerce')
        df['bulan'] = df['periode'].dt.strftime('%B %Y')

        return df

    except Exception as e:
        print("❌ ERROR LOAD RECRUITMENT:", e)
        return pd.DataFrame()
    


# LOAD SALARY & PPH21
def load_salary():
    try:
        # === DATA GAJI ===
        df_gaji = pd.read_excel(
            FILE_PATH,
            sheet_name="GAJI DAN PPH21",
            usecols="B:D",
            skiprows=3            # ✅ lewati 3 baris header
        )
        df_gaji.columns = ["project", "periode", "gaji"]
        df_gaji['project'] = df_gaji['project'].astype(str).str.strip()
        df_gaji['gaji'] = pd.to_numeric(df_gaji['gaji'], errors='coerce')
        df_gaji = df_gaji.dropna(subset=['project', 'gaji'])

        # Konversi periode ke datetime, lalu format "Jan 2026"
        df_gaji['periode_dt'] = pd.to_datetime(df_gaji['periode'], errors='coerce')
        df_gaji['periode'] = df_gaji['periode_dt'].dt.strftime('%b %Y')
        df_gaji = df_gaji.dropna(subset=['periode'])

        # === DATA PPH21 ===
        df_pph = pd.read_excel(
            FILE_PATH,
            sheet_name="GAJI DAN PPH21",
            usecols="G:I",
            skiprows=3
        )
        df_pph.columns = ["project", "periode", "pph21"]
        df_pph['project'] = df_pph['project'].astype(str).str.strip()
        df_pph['pph21'] = pd.to_numeric(df_pph['pph21'], errors='coerce')
        df_pph = df_pph.dropna(subset=['project', 'pph21'])

        df_pph['periode_dt'] = pd.to_datetime(df_pph['periode'], errors='coerce')
        df_pph['periode'] = df_pph['periode_dt'].dt.strftime('%b %Y')
        df_pph = df_pph.dropna(subset=['periode'])

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

        df['jenis'] = df['jenis'].astype(str).str.strip().str.lower()
        df['divisi'] = df['divisi'].astype(str).str.strip()
        df['pelatihan'] = df['pelatihan'].astype(str).str.strip()
        df['status'] = df['status'].astype(str).str.strip().str.lower()

        df['periode'] = pd.to_datetime(df['periode'], errors='coerce')
        df['bulan'] = df['periode'].dt.strftime('%B %Y')

        return df

    except Exception as e:
        print("❌ ERROR LOAD TRAINING:", e)
        return pd.DataFrame()
    


# LOAD BPJS DATA
def load_bpjs():
    try:
        # =========================
        # DATA KARYAWAN
        # =========================
        df_karyawan = pd.read_excel(
            FILE_PATH,
            sheet_name="BPJS",
            usecols="B:D"
        )
        df_karyawan.columns = ["nama", "project", "jenis_bpjs"]
        df_karyawan['project'] = df_karyawan['project'].astype(str).str.strip()
        df_karyawan['jenis_bpjs'] = df_karyawan['jenis_bpjs'].astype(str).str.strip().str.lower()
        # Hapus baris yang nama-nya kosong (jika ada)
        df_karyawan = df_karyawan.dropna(subset=['nama'])

        # =========================
        # BPJS KESEHATAN
        # =========================
        df_kesehatan = pd.read_excel(
            FILE_PATH,
            sheet_name="BPJS",
            usecols="G:H"
        )
        df_kesehatan.columns = ["periode", "pembayaran"]

        # =========================
        # BPJS TK PERMANENT
        # =========================
        df_tk_permanent = pd.read_excel(
            FILE_PATH,
            sheet_name="PAY BPJS TK",
            usecols="C:D"
        )
        df_tk_permanent.columns = ["periode", "pembayaran"]

        # =========================
        # BPJS TK BORONGAN
        # =========================
        df_tk_borongan = pd.read_excel(
            FILE_PATH,
            sheet_name="PAY BPJS TK",
            usecols="F:G"
        )
        df_tk_borongan.columns = ["periode", "pembayaran"]

        # =========================
        # FUNGSI CLEANING UANG (ROBUST)
        # =========================
        def clean_uang(df):
            def fix_number(x):
                # Jika sudah numerik, langsung return sebagai int
                if isinstance(x, (int, float)):
                    return int(x) if pd.notna(x) else None
                # Jika string, bersihkan simbol dan ribuan
                x = str(x).strip()
                # Hapus 'Rp', spasi, dan titik ribuan
                x = x.replace('Rp', '').replace(' ', '').replace('.', '')
                # Jika ada koma desimal, ambil bagian sebelum koma (asumsi tanpa desimal)
                if ',' in x:
                    x = x.split(',')[0]
                try:
                    return int(x)
                except:
                    return None

            df = df.copy()
            df['pembayaran'] = df['pembayaran'].apply(fix_number)
            return df.dropna(subset=['pembayaran'])

        df_kesehatan = clean_uang(df_kesehatan)
        df_tk_permanent = clean_uang(df_tk_permanent)
        df_tk_borongan = clean_uang(df_tk_borongan)

        # Konversi periode ke datetime
        df_kesehatan['periode'] = pd.to_datetime(df_kesehatan['periode'], errors='coerce')
        df_tk_permanent['periode'] = pd.to_datetime(df_tk_permanent['periode'], errors='coerce')
        df_tk_borongan['periode'] = pd.to_datetime(df_tk_borongan['periode'], errors='coerce')

        return df_karyawan, df_kesehatan, df_tk_permanent, df_tk_borongan

    except Exception as e:
        print("❌ ERROR LOAD BPJS:", e)
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

def load_overtime():
    try:
        # === DATA OVERTIME (Kolom B:D) ===
        df_overtime = pd.read_excel(
            FILE_PATH,
            sheet_name="OVERTIME DAN ABSENSI",
            usecols="B:D",
            skiprows=3
        )
        df_overtime.columns = ["project", "bulan", "overtime"]

        # === DATA ABSENSI (Kolom F:G) ===
        df_absensi = pd.read_excel(
            FILE_PATH,
            sheet_name="OVERTIME DAN ABSENSI",
            usecols="F:G",
            skiprows=3
        )
        df_absensi.columns = ["periode", "absensi"]

        # === DATA COST (Kolom I:M) ===
        df_cost = pd.read_excel(
            FILE_PATH,
            sheet_name="OVERTIME DAN ABSENSI",
            usecols="I:M",
            skiprows=3
        )
        df_cost.columns = ["project", "bulan", "total_cost", "overtime_cost", "overtime_percent"]

        # === FUNGSI PEMBERSIH ANGKA (ROBUST) ===
        def clean_number(x):
            if pd.isna(x):
                return None
            if isinstance(x, (int, float)):
                return float(x)
            s = str(x).strip()
            # Hapus simbol Rp, spasi
            s = s.replace('Rp', '').replace(' ', '')
            # Deteksi pola: jika ada titik dan koma, titik = ribuan, koma = desimal
            if '.' in s and ',' in s:
                s = s.replace('.', '').replace(',', '.')
            elif ',' in s:
                # Hanya koma: jika jumlah koma > 1 → ribuan (hapus semua)
                if s.count(',') > 1:
                    s = s.replace(',', '')
                else:
                    # Satu koma: cek apakah ada <=2 digit setelahnya → desimal
                    parts = s.split(',')
                    if len(parts) == 2 and len(parts[1]) <= 2:
                        s = s.replace(',', '.')
                    else:
                        s = s.replace(',', '')
            elif '.' in s:
                # Hanya titik: jika jumlah titik > 1 → ribuan (hapus semua)
                if s.count('.') > 1:
                    s = s.replace('.', '')
                # satu titik biarkan sebagai desimal
            try:
                return float(s)
            except:
                return None

        # Terapkan ke kolom biaya
        for col in ['total_cost', 'overtime_cost', 'overtime_percent']:
            df_cost[col] = df_cost[col].apply(clean_number)

        # Kolom persentase lain (sudah numerik)
        df_overtime['overtime'] = pd.to_numeric(df_overtime['overtime'], errors='coerce')
        df_absensi['absensi'] = pd.to_numeric(df_absensi['absensi'], errors='coerce')

        # Konversi tanggal
        df_overtime['bulan'] = pd.to_datetime(df_overtime['bulan'], errors='coerce')
        df_absensi['periode'] = pd.to_datetime(df_absensi['periode'], errors='coerce')
        df_cost['bulan'] = pd.to_datetime(df_cost['bulan'], errors='coerce')

        # Buang baris tanpa data penting
        df_overtime = df_overtime.dropna(subset=['project', 'bulan', 'overtime'])
        df_absensi = df_absensi.dropna(subset=['periode', 'absensi'])
        df_cost = df_cost.dropna(subset=['project', 'bulan', 'total_cost', 'overtime_cost'])

        return df_overtime, df_absensi, df_cost

    except Exception as e:
        print("❌ ERROR LOAD OVERTIME:", e)
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    
# LOAD MCU DATA
def load_mcu():
    try:
        # =========================
        # DATA MCU
        # =========================
        df_mcu = pd.read_excel(
            FILE_PATH,
            sheet_name="MCU FEB 2026",
            usecols="B:H"
        )

        df_mcu.columns = [
            "nama",
            "divisi",
            "project",
            "tahun",
            "periode",
            "tipe_mcu",
            "hasil_mcu"
        ]

        df_mcu['nama'] = df_mcu['nama'].astype(str).str.strip()
        df_mcu['project'] = df_mcu['project'].astype(str).str.strip()
        df_mcu['divisi'] = df_mcu['divisi'].astype(str).str.strip()
        df_mcu['hasil_mcu'] = df_mcu['hasil_mcu'].astype(str).str.strip().str.lower()

        df_mcu['periode'] = pd.to_datetime(df_mcu['periode'], errors='coerce')
        # ✅ Ubah format bulan menjadi "Jan 2026"
        df_mcu['bulan'] = df_mcu['periode'].dt.strftime('%b %Y')

        # =========================
        # DATA KARYAWAN (ATTRIBUTE TAMBAHAN)
        # =========================
        df_karyawan = pd.read_excel(
            FILE_PATH,
            sheet_name="MCU FEB 2026",
            usecols="K:O"
        )

        df_karyawan.columns = [
            "nama",
            "branch",
            "job",
            "ring",
            "gender"
        ]

        df_karyawan['nama'] = df_karyawan['nama'].astype(str).str.strip()
        df_karyawan['gender'] = df_karyawan['gender'].astype(str).str.strip().str.lower()

        # =========================
        # DATA PAYMENT MCU
        # =========================
        df_pay = pd.read_excel(
            FILE_PATH,
            sheet_name="Pay MCU 2026",
            usecols="T:W"
        )

        df_pay.columns = [
            "nama",
            "project",
            "periode",
            "pembayaran"
        ]

        df_pay['nama'] = df_pay['nama'].astype(str).str.strip()
        df_pay['project'] = df_pay['project'].astype(str).str.strip()

        # 🔥 BUANG BARIS DENGAN PROJECT KOSONG ATAU "nan"
        df_pay = df_pay[
            (df_pay['project'].notna()) & 
            (df_pay['project'] != '') & 
            (df_pay['project'] != 'nan')
        ]

        # Fungsi pembersih uang
        def clean_pembayaran(x):
            if isinstance(x, (int, float)):
                return int(x) if pd.notna(x) else None
            x = str(x).strip()
            x = x.replace('Rp', '').replace(' ', '').replace('.', '')
            if ',' in x:
                x = x.split(',')[0]
            try:
                return int(x)
            except:
                return None

        df_pay['pembayaran'] = df_pay['pembayaran'].apply(clean_pembayaran)
        df_pay = df_pay.dropna(subset=['pembayaran'])

        # ✅ Ubah format periode di df_pay menjadi "Jan 2025"
        df_pay['periode'] = pd.to_datetime(df_pay['periode'], errors='coerce')
        df_pay['periode'] = df_pay['periode'].dt.strftime('%b %Y')

        return df_mcu, df_karyawan, df_pay

    except Exception as e:
        print("❌ ERROR LOAD MCU:", e)
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()


def clean_uang(df):
    df = df.copy()

    def fix_number(x):
        # ✅ kalau sudah angka → JANGAN DIAPA-APAIN
        if isinstance(x, (int, float)):
            return x

        x = str(x)
        x = x.replace('Rp', '')
        x = x.replace('.', '')   # hapus ribuan
        x = x.strip()

        try:
            return int(x)
        except:
            return None

    df['pembayaran'] = df['pembayaran'].apply(fix_number)

    return df.dropna(subset=['pembayaran'])

def json_safe_records(df):
    cleaned = df.copy()
    cleaned = cleaned.replace([float('inf'), float('-inf')], pd.NA)
    cleaned = cleaned.astype(object).where(pd.notna(cleaned), None)
    return cleaned.to_dict(orient='records')



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
    periode = request.args.get("periode")   # format "Jan 2026"
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

def apply_overtime_filters(df_overtime, df_absensi, df_cost):
    project = request.args.get("project")
    start = request.args.get("start")
    end = request.args.get("end")

    if project:
        df_overtime = df_overtime[df_overtime['project'] == project]
        df_cost = df_cost[df_cost['project'] == project]

    if start and end:
        start_dt = pd.to_datetime(start)
        end_dt = pd.to_datetime(end)
        df_overtime = df_overtime[(df_overtime['bulan'] >= start_dt) & (df_overtime['bulan'] <= end_dt)]
        df_absensi = df_absensi[(df_absensi['periode'] >= start_dt) & (df_absensi['periode'] <= end_dt)]
        df_cost = df_cost[(df_cost['bulan'] >= start_dt) & (df_cost['bulan'] <= end_dt)]

    return df_overtime, df_absensi, df_cost


# FILTER MCU

def apply_mcu_filters(df_mcu, df_pay):
    project = request.args.get("project")
    hasil = request.args.get("hasil")
    start = request.args.get("start")
    end = request.args.get("end")

    if project:
        df_mcu = df_mcu[df_mcu['project'] == project]
        df_pay = df_pay[df_pay['project'] == project] 

    if hasil:
        df_mcu = df_mcu[df_mcu['hasil_mcu'] == hasil.lower()]

    if start and end:
        df_mcu = df_mcu[
            (df_mcu['periode'] >= pd.to_datetime(start)) &
            (df_mcu['periode'] <= pd.to_datetime(end))
        ]

    return df_mcu, df_pay



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
def apply_bpjs_filters(df_karyawan, df_kesehatan, df_tk_permanent, df_tk_borongan):
    project = request.args.get("project")
    jenis_bpjs = request.args.get("jenis_bpjs")
    periode = request.args.get("periode")

    if project:
        df_karyawan = df_karyawan[df_karyawan['project'] == project]

    if jenis_bpjs:
        df_karyawan = df_karyawan[df_karyawan['jenis_bpjs'] == jenis_bpjs.lower()]

    if periode:
        df_kesehatan = df_kesehatan[df_kesehatan['periode'] == periode]
        df_tk_permanent = df_tk_permanent[df_tk_permanent['periode'] == periode]
        df_tk_borongan = df_tk_borongan[df_tk_borongan['periode'] == periode]

    return df_karyawan, df_kesehatan, df_tk_permanent, df_tk_borongan

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
def calculate_bpjs_kpi(df_kesehatan, df_tk_permanent, df_tk_borongan, df_karyawan):

    return {
        "bpjs_kesehatan": int(df_kesehatan['pembayaran'].sum()),
        "bpjs_tk_permanent": int(df_tk_permanent['pembayaran'].sum()),
        "bpjs_tk_borongan": int(df_tk_borongan['pembayaran'].sum()),
        "total_semua": int(
            df_kesehatan['pembayaran'].sum() +
            df_tk_permanent['pembayaran'].sum() +
            df_tk_borongan['pembayaran'].sum()
        ),
        "total_karyawan_bpjs": len(df_karyawan)
    }


# KPI MCU
def calculate_mcu_kpi(df, df_pay):
    total = len(df)

    fit = len(df[df['hasil_mcu'].str.contains('fit', na=False)])
    unfit = len(df[df['hasil_mcu'].str.contains('unfit', na=False)])

    # Gunakan df_pay yang sudah difilter (tanpa project kosong)
    total_cost = df_pay['pembayaran'].sum()

    return {
        "total_mcu": total,
        "fit": fit,
        "unfit": unfit,
        "fit_rate": round((fit / total) * 100, 2) if total > 0 else 0,
        "total_cost": int(total_cost)
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

    # Terapkan filter
    df_gaji = apply_salary_filters(df_gaji, "periode")
    df_pph = apply_salary_filters(df_pph, "periode")

    # KPI
    kpi = calculate_salary_kpi(df_gaji, df_pph)

    # Gaji per Project (bar chart)
    gaji_project = df_gaji.groupby('project')['gaji'].sum().reset_index()
    gaji_project = gaji_project.sort_values('gaji', ascending=False)

    # Tren Gaji per Bulan (line chart)
    gaji_trend = df_gaji.groupby('periode')['gaji'].sum().reset_index()
    gaji_trend['sort_date'] = pd.to_datetime(gaji_trend['periode'], format='%b %Y')
    gaji_trend = gaji_trend.sort_values('sort_date').drop('sort_date', axis=1)

    # PPh21 per Project
    pph_project = df_pph.groupby('project')['pph21'].sum().reset_index()
    pph_project = pph_project.sort_values('pph21', ascending=False)

    # Tren PPh21 per Bulan
    pph_trend = df_pph.groupby('periode')['pph21'].sum().reset_index()
    pph_trend['sort_date'] = pd.to_datetime(pph_trend['periode'], format='%b %Y')
    pph_trend = pph_trend.sort_values('sort_date').drop('sort_date', axis=1)

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
    df_karyawan, df_kesehatan, df_tk_permanent, df_tk_borongan = load_bpjs()

    df_mcu, _, _ = load_mcu()
    total_mcu = len(df_mcu)

    df_karyawan, df_kesehatan, df_tk_permanent, df_tk_borongan = apply_bpjs_filters(
        df_karyawan, df_kesehatan, df_tk_permanent, df_tk_borongan
    )

    # KPI
    kpi = calculate_bpjs_kpi(df_kesehatan, df_tk_permanent, df_tk_borongan, df_karyawan)

    # DISTRIBUSI KARYAWAN
    jenis = df_karyawan['jenis_bpjs'].value_counts().reset_index()
    jenis.columns = ['jenis_bpjs', 'jumlah']

    project = df_karyawan['project'].value_counts().reset_index()
    project.columns = ['project', 'jumlah']

    # =========================
    # TREND PER KATEGORI
    # =========================
    kesehatan_trend = df_kesehatan.groupby('periode')['pembayaran'].sum().reset_index()
    kesehatan_trend['kategori'] = 'kesehatan'

    tk_permanent_trend = df_tk_permanent.groupby('periode')['pembayaran'].sum().reset_index()
    tk_permanent_trend['kategori'] = 'tk_permanent'

    tk_borongan_trend = df_tk_borongan.groupby('periode')['pembayaran'].sum().reset_index()
    tk_borongan_trend['kategori'] = 'tk_borongan'

    trend = pd.concat([
        kesehatan_trend,
        tk_permanent_trend,
        tk_borongan_trend
    ])

    trend['pembayaran'] = trend['pembayaran'].astype(int)

    return jsonify({
        "kpi": kpi,
        "jenis": jenis.to_dict(orient='records'),
        "project": project.to_dict(orient='records'),
        "trend": trend.to_dict(orient='records')
    })


@app.route("/overtime/dashboard")
def overtime_dashboard():
    df_overtime, df_absensi, df_cost = load_overtime()

    # Terapkan filter (jika ada)
    df_overtime, df_absensi, df_cost = apply_overtime_filters(df_overtime, df_absensi, df_cost)

    # ========== KPI ==========
    avg_overtime_pct = df_overtime['overtime'].mean() * 100 if not df_overtime.empty else 0
    avg_absensi_pct = df_absensi['absensi'].mean() * 100 if not df_absensi.empty else 0
    total_cost = df_cost['total_cost'].sum()
    total_overtime_cost = df_cost['overtime_cost'].sum()

    # Growth overtime (bulan terbaru vs sebelumnya)
    monthly_ov = df_overtime.groupby(df_overtime['bulan'].dt.to_period('M'))['overtime'].mean().sort_index()
    ov_growth = 0
    if len(monthly_ov) >= 2:
        ov_growth = ((monthly_ov.iloc[-1] - monthly_ov.iloc[-2]) / monthly_ov.iloc[-2]) * 100

    # Growth absensi
    monthly_att = df_absensi.groupby(df_absensi['periode'].dt.to_period('M'))['absensi'].mean().sort_index()
    att_growth = 0
    if len(monthly_att) >= 2:
        att_growth = ((monthly_att.iloc[-1] - monthly_att.iloc[-2]) / monthly_att.iloc[-2]) * 100

    kpi = {
        "avg_overtime_percent": round(avg_overtime_pct, 2),
        "avg_absensi_percent": round(avg_absensi_pct, 2),
        "total_cost": int(total_cost),
        "total_overtime_cost": int(total_overtime_cost),
        "overtime_growth": round(ov_growth, 2),
        "absensi_growth": round(att_growth, 2)
    }

    # ========== Monthly Overtime Summary (Line Chart) ==========
    ov_summary = df_overtime.groupby(df_overtime['bulan'].dt.strftime('%b %Y'))['overtime'].mean().reset_index()
    ov_summary.columns = ['bulan', 'overtime']
    ov_summary['overtime'] = ov_summary['overtime'].round(2)  # simpan 4 digit untuk persentase akurat
    ov_summary = ov_summary.sort_values('bulan', key=lambda x: pd.to_datetime(x, format='%b %Y'))

    # ========== Monthly Attendance Summary ==========
    att_summary = df_absensi.groupby(df_absensi['periode'].dt.strftime('%b %Y'))['absensi'].mean().reset_index()
    att_summary.columns = ['bulan', 'absensi']
    att_summary['absensi'] = att_summary['absensi'].round(2)
    att_summary = att_summary.sort_values('bulan', key=lambda x: pd.to_datetime(x, format='%b %Y'))

    # ========== Overtime per Project ==========
    ov_project = df_overtime.groupby('project')['overtime'].mean().reset_index()
    ov_project['overtime_percent'] = (ov_project['overtime'] * 100).round(2)
    ov_project = ov_project.sort_values('overtime_percent', ascending=False)

    # ========== Cost per Project ==========
    cost_project = df_cost.groupby('project').agg({
        'total_cost': 'sum',
        'overtime_cost': 'sum',
        'overtime_percent': 'mean'
    }).reset_index()
    cost_project['total_cost'] = cost_project['total_cost'].round(2)
    cost_project['overtime_cost'] = cost_project['overtime_cost'].round(2)
    cost_project['overtime_percent'] = cost_project['overtime_percent'].round(2)

    # ========== Top 5 Projects by Overtime Cost ==========
    top_cost = cost_project.nlargest(5, 'overtime_cost')[['project', 'overtime_cost']].round(2)

    # ========== Monthly Overtime Cost Trend ==========
    cost_trend = df_cost.groupby(df_cost['bulan'].dt.strftime('%b %Y'))['overtime_cost'].sum().reset_index()
    cost_trend.columns = ['bulan', 'overtime_cost']
    cost_trend['overtime_cost'] = cost_trend['overtime_cost'].round(2)
    cost_trend = cost_trend.sort_values('bulan', key=lambda x: pd.to_datetime(x, format='%b %Y'))

    return jsonify({
        "kpi": kpi,
        "overtime_summary": ov_summary.to_dict(orient='records'),
        "attendance_summary": att_summary.to_dict(orient='records'),
        "overtime_by_project": ov_project.to_dict(orient='records'),
        "cost_by_project": cost_project.to_dict(orient='records'),
        "top_cost_projects": top_cost.to_dict(orient='records'),
        "overtime_cost_trend": cost_trend.to_dict(orient='records')
    })  

# DASHBOARD MCU
@app.route("/mcu/dashboard")
def mcu_dashboard():
    df_mcu, df_karyawan, df_pay = load_mcu()

    # =========================
    # MERGE DATA
    # =========================
    df_full = pd.merge(df_mcu, df_karyawan, on="nama", how="left")

    # =========================
    # FILTER
    # =========================
    df_full, df_pay = apply_mcu_filters(df_full, df_pay)

    # =========================
    # KPI
    # =========================
    kpi = calculate_mcu_kpi(df_full, df_pay)

    # =========================
    # HASIL MCU (PIE)
    # =========================
    hasil = df_full['hasil_mcu'].value_counts().reset_index()
    hasil.columns = ['hasil', 'jumlah']

    # =========================
    # DIVISI (BAR)
    # =========================
    divisi = df_full['divisi'].value_counts().reset_index()
    divisi.columns = ['divisi', 'jumlah']

    # =========================
    # PROJECT (BAR)
    # =========================
    project = df_full['project'].value_counts().reset_index()
    project.columns = ['project', 'jumlah']

    # =========================
    # TREND MCU (LINE) – SEKARANG "bulan" SUDAH FORMAT "Jan 2026"
    # =========================
    trend = df_full['bulan'].value_counts().reset_index()
    trend.columns = ['bulan', 'jumlah']
    # Sorting berdasarkan tanggal asli agar urut
    trend['sort_date'] = pd.to_datetime(trend['bulan'], format='%b %Y', errors='coerce')
    trend = trend.sort_values('sort_date').drop('sort_date', axis=1)

    # =========================
    # GENDER
    # =========================
    gender = df_full['gender'].value_counts().reset_index()
    gender.columns = ['gender', 'jumlah']

    # =========================
    # COST PER PROJECT
    # =========================
    cost_project = df_pay.groupby('project')['pembayaran'].sum().reset_index()

    # =========================
    # COST TREND – PERIODE SUDAH FORMAT "Jan 2025"
    # =========================
    cost_trend = df_pay.groupby('periode')['pembayaran'].sum().reset_index()
    # Sorting berdasarkan tanggal asli
    cost_trend['sort_date'] = pd.to_datetime(cost_trend['periode'], format='%b %Y', errors='coerce')
    cost_trend = cost_trend.sort_values('sort_date').drop('sort_date', axis=1)

    # =========================
    # TOP PROJECT (COST)
    # =========================
    top_project = cost_project.sort_values('pembayaran', ascending=False).head(5)

    return jsonify({
        "kpi": kpi,
        "hasil_mcu": hasil.to_dict(orient='records'),
        "divisi": divisi.to_dict(orient='records'),
        "project": project.to_dict(orient='records'),
        "trend": trend.to_dict(orient='records'),
        "gender": gender.to_dict(orient='records'),
        "cost_project": cost_project.to_dict(orient='records'),
        "cost_trend": cost_trend.to_dict(orient='records'),
        "top_project": top_project.to_dict(orient='records')
    })


# RUN

if __name__ == "__main__":
    app.run(debug=True)

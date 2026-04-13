from flask import Flask, jsonify, request
from flask_cors import CORS
import pandas as pd
from datetime import datetime, timedelta

app = Flask(__name__)
CORS(app)

FILE_PATH = "DATA DASHBOARD HRD PT ALDZAMA.xlsx"


# =======================================================
# HELPER
# =======================================================
def json_safe_records(df):
    cleaned = df.copy()
    cleaned = cleaned.replace([float('inf'), float('-inf')], pd.NA)
    cleaned = cleaned.astype(object).where(pd.notna(cleaned), None)
    return cleaned.to_dict(orient='records')


# =======================================================
# LOAD INTERNSHIP
# =======================================================
def load_data():
    try:
        df = pd.read_excel(FILE_PATH, sheet_name="INTERNSHIP", skiprows=2, usecols="B:M")
        df.columns = ["nama","program","institusi","kota","ket","permohonan",
                      "penempatan","durasi","status","final_project","potensi","berakhir"]
        df['institusi']  = df['institusi'].astype(str).str.strip()
        df['status']     = df['status'].astype(str).str.strip().str.lower()
        df['penempatan'] = df['penempatan'].astype(str).str.strip()
        df['permohonan'] = pd.to_datetime(df['permohonan'], errors='coerce')
        df['bulan']      = df['permohonan'].dt.strftime('%B %Y')
        df['berakhir']   = pd.to_datetime(df['berakhir'], dayfirst=True, errors='coerce')
        return df
    except Exception as e:
        print("ERROR LOAD INTERNSHIP:", e); return pd.DataFrame()


# =======================================================
# LOAD MANPOWER
# =======================================================
def load_manpower():
    try:
        df = pd.read_excel(FILE_PATH, sheet_name="MAN POWER", usecols="B:H")
        df.columns = ["nama","jenis_kelamin","jabatan","branch","group_project","ring","status_kontrak"]
        df['jabatan']        = df['jabatan'].astype(str).str.strip()
        df['branch']         = df['branch'].astype(str).str.strip()
        df['group_project']  = df['group_project'].astype(str).str.strip()
        df['status_kontrak'] = df['status_kontrak'].astype(str).str.strip().str.lower()
        return df
    except Exception as e:
        print("ERROR LOAD MANPOWER:", e); return pd.DataFrame()


# =======================================================
# LOAD RECRUITMENT
# =======================================================
def load_recruitment():
    try:
        sheets = ["RECRUITMENT","RECRUITMENT OKT","RECRUITMENT FEB 2026"]
        df_list = []
        for sheet in sheets:
            df = pd.read_excel(FILE_PATH, sheet_name=sheet, usecols="B:E")
            df.columns = ["nama","posisi","periode","status"]
            df_list.append(df)
        df = pd.concat(df_list, ignore_index=True)
        df['posisi']  = df['posisi'].astype(str).str.strip()
        df['status']  = df['status'].astype(str).str.strip().str.lower()
        df['status']  = df['status'].replace({'on process user':'on process','on process hr':'on process'})
        df['periode'] = pd.to_datetime(df['periode'], errors='coerce')
        df['bulan']   = df['periode'].dt.strftime('%B %Y')
        return df
    except Exception as e:
        print("ERROR LOAD RECRUITMENT:", e); return pd.DataFrame()


# =======================================================
# LOAD SALARY
# =======================================================
def load_salary():
    try:
        df_gaji = pd.read_excel(FILE_PATH, sheet_name="GAJI DAN PPH21", usecols="B:D", skiprows=3)
        df_gaji.columns = ["project","periode","gaji"]
        df_gaji['project']    = df_gaji['project'].astype(str).str.strip()
        df_gaji['gaji']       = pd.to_numeric(df_gaji['gaji'], errors='coerce')
        df_gaji               = df_gaji.dropna(subset=['project','gaji'])
        df_gaji['periode_dt'] = pd.to_datetime(df_gaji['periode'], errors='coerce')
        df_gaji['periode']    = df_gaji['periode_dt'].dt.strftime('%b %Y')
        df_gaji               = df_gaji.dropna(subset=['periode'])

        df_pph = pd.read_excel(FILE_PATH, sheet_name="GAJI DAN PPH21", usecols="G:I", skiprows=3)
        df_pph.columns = ["project","periode","pph21"]
        df_pph['project']    = df_pph['project'].astype(str).str.strip()
        df_pph['pph21']      = pd.to_numeric(df_pph['pph21'], errors='coerce')
        df_pph               = df_pph.dropna(subset=['project','pph21'])
        df_pph['periode_dt'] = pd.to_datetime(df_pph['periode'], errors='coerce')
        df_pph['periode']    = df_pph['periode_dt'].dt.strftime('%b %Y')
        df_pph               = df_pph.dropna(subset=['periode'])
        return df_gaji, df_pph
    except Exception as e:
        print("ERROR LOAD SALARY:", e); return pd.DataFrame(), pd.DataFrame()


# =======================================================
# LOAD TRAINING
# =======================================================
def load_training():
    try:
        df = pd.read_excel(FILE_PATH, sheet_name="INTERNAL & EXTERNAL TRAINING FE", usecols="B:H")
        df.columns = ["jenis","nama","divisi","pelatihan","lembaga","periode","status"]
        df['jenis']     = df['jenis'].astype(str).str.strip().str.lower()
        df['divisi']    = df['divisi'].astype(str).str.strip()
        df['pelatihan'] = df['pelatihan'].astype(str).str.strip()
        df['status']    = df['status'].astype(str).str.strip().str.lower()
        df['periode']   = pd.to_datetime(df['periode'], errors='coerce')
        df['bulan']     = df['periode'].dt.strftime('%B %Y')
        return df
    except Exception as e:
        print("ERROR LOAD TRAINING:", e); return pd.DataFrame()


# =======================================================
# LOAD BPJS
# =======================================================
def load_bpjs():
    try:
        df_karyawan = pd.read_excel(FILE_PATH, sheet_name="BPJS", usecols="B:D")
        df_karyawan.columns = ["nama","project","jenis_bpjs"]
        df_karyawan['project']    = df_karyawan['project'].astype(str).str.strip()
        df_karyawan['jenis_bpjs'] = df_karyawan['jenis_bpjs'].astype(str).str.strip().str.lower()
        df_karyawan               = df_karyawan.dropna(subset=['nama'])

        def clean_uang_bpjs(df):
            def fix(x):
                if isinstance(x, (int, float)): return int(x) if pd.notna(x) else None
                x = str(x).strip().replace('Rp','').replace(' ','').replace('.','')
                if ',' in x: x = x.split(',')[0]
                try: return int(x)
                except: return None
            df = df.copy(); df['pembayaran'] = df['pembayaran'].apply(fix)
            return df.dropna(subset=['pembayaran'])

        df_kes  = pd.read_excel(FILE_PATH, sheet_name="BPJS", usecols="G:H")
        df_kes.columns  = ["periode","pembayaran"]; df_kes  = clean_uang_bpjs(df_kes)
        df_kes['periode']  = pd.to_datetime(df_kes['periode'],  errors='coerce')

        df_perm = pd.read_excel(FILE_PATH, sheet_name="PAY BPJS TK", usecols="C:D")
        df_perm.columns = ["periode","pembayaran"]; df_perm = clean_uang_bpjs(df_perm)
        df_perm['periode'] = pd.to_datetime(df_perm['periode'], errors='coerce')

        df_bor  = pd.read_excel(FILE_PATH, sheet_name="PAY BPJS TK", usecols="F:G")
        df_bor.columns  = ["periode","pembayaran"]; df_bor  = clean_uang_bpjs(df_bor)
        df_bor['periode']  = pd.to_datetime(df_bor['periode'],  errors='coerce')

        return df_karyawan, df_kes, df_perm, df_bor
    except Exception as e:
        print("ERROR LOAD BPJS:", e); return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()


# =======================================================
# LOAD OVERTIME
# =======================================================
def load_overtime():
    try:
        df_overtime = pd.read_excel(FILE_PATH, sheet_name="OVERTIME DAN ABSENSI", usecols="B:D", skiprows=3)
        df_overtime.columns = ["project","bulan","overtime"]

        df_absensi = pd.read_excel(FILE_PATH, sheet_name="OVERTIME DAN ABSENSI", usecols="F:G", skiprows=3)
        df_absensi.columns = ["periode","absensi"]

        df_cost = pd.read_excel(FILE_PATH, sheet_name="OVERTIME DAN ABSENSI", usecols="I:M", skiprows=3)
        df_cost.columns = ["project","bulan","total_cost","overtime_cost","overtime_percent"]

        def clean_number(x):
            if pd.isna(x): return None
            if isinstance(x, (int, float)): return float(x)
            s = str(x).strip().replace('Rp','').replace(' ','')
            if '.' in s and ',' in s: s = s.replace('.','').replace(',','.')
            elif ',' in s:
                parts = s.split(',')
                s = s.replace(',','') if s.count(',') > 1 or len(parts[1]) > 2 else s.replace(',','.')
            elif '.' in s and s.count('.') > 1: s = s.replace('.','')
            try: return float(s)
            except: return None

        for col in ['total_cost','overtime_cost','overtime_percent']:
            df_cost[col] = df_cost[col].apply(clean_number)

        df_overtime['overtime'] = pd.to_numeric(df_overtime['overtime'], errors='coerce')
        df_absensi['absensi']   = pd.to_numeric(df_absensi['absensi'],   errors='coerce')
        df_overtime['bulan']    = pd.to_datetime(df_overtime['bulan'],   errors='coerce')
        df_absensi['periode']   = pd.to_datetime(df_absensi['periode'],  errors='coerce')
        df_cost['bulan']        = pd.to_datetime(df_cost['bulan'],       errors='coerce')

        df_overtime = df_overtime.dropna(subset=['project','bulan','overtime'])
        df_absensi  = df_absensi.dropna(subset=['periode','absensi'])
        df_cost     = df_cost.dropna(subset=['project','bulan','total_cost','overtime_cost'])
        return df_overtime, df_absensi, df_cost
    except Exception as e:
        print("ERROR LOAD OVERTIME:", e); return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()


# =======================================================
# LOAD MCU
# =======================================================
def load_mcu():
    try:
        df_mcu = pd.read_excel(FILE_PATH, sheet_name="MCU FEB 2026", usecols="B:H")
        df_mcu.columns = ["nama","divisi","project","tahun","periode","tipe_mcu","hasil_mcu"]
        df_mcu['nama']      = df_mcu['nama'].astype(str).str.strip()
        df_mcu['project']   = df_mcu['project'].astype(str).str.strip()
        df_mcu['divisi']    = df_mcu['divisi'].astype(str).str.strip()
        df_mcu['hasil_mcu'] = df_mcu['hasil_mcu'].astype(str).str.strip().str.lower()
        df_mcu['periode']   = pd.to_datetime(df_mcu['periode'], errors='coerce')
        df_mcu['bulan']     = df_mcu['periode'].dt.strftime('%b %Y')

        df_karyawan = pd.read_excel(FILE_PATH, sheet_name="MCU FEB 2026", usecols="K:O")
        df_karyawan.columns = ["nama","branch","job","ring","gender"]
        df_karyawan['nama']   = df_karyawan['nama'].astype(str).str.strip()
        df_karyawan['gender'] = df_karyawan['gender'].astype(str).str.strip().str.lower()

        df_pay = pd.read_excel(FILE_PATH, sheet_name="Pay MCU 2026", usecols="T:W")
        df_pay.columns = ["nama","project","periode","pembayaran"]
        df_pay['nama']    = df_pay['nama'].astype(str).str.strip()
        df_pay['project'] = df_pay['project'].astype(str).str.strip()
        df_pay = df_pay[(df_pay['project'].notna()) & (df_pay['project'] != '') & (df_pay['project'] != 'nan')]

        def clean_pembayaran(x):
            if isinstance(x, (int, float)): return int(x) if pd.notna(x) else None
            x = str(x).strip().replace('Rp','').replace(' ','').replace('.','')
            if ',' in x: x = x.split(',')[0]
            try: return int(x)
            except: return None

        df_pay['pembayaran'] = df_pay['pembayaran'].apply(clean_pembayaran)
        df_pay = df_pay.dropna(subset=['pembayaran'])
        df_pay['periode'] = pd.to_datetime(df_pay['periode'], errors='coerce').dt.strftime('%b %Y')
        return df_mcu, df_karyawan, df_pay
    except Exception as e:
        print("ERROR LOAD MCU:", e); return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()


# =======================================================
# LOAD KPI
# Sheet "KPI" — 2 bagian:
#   Bagian 1 (baris 3-21): Definisi 9 KPI
#     Kolom: B=NO, C=Nama KPI, D=Variable, E=Unit, F=Target desc
#     (Tidak ada lagi kolom historical bulan Jan-Des)
#   Bagian 2 (baris 26-43): Monthly 2026 — Target vs Aktual per KPI per bulan
#     Kolom: B=NO, C=Nama KPI, D=Bulan, E=Target, F=Aktual, G=Status (kosong)
# =======================================================
def load_kpi_data():
    try:
        df = pd.read_excel(FILE_PATH, sheet_name="KPI", header=None)

        def sf(v):
            if v is None or (isinstance(v, float) and pd.isna(v)): return None
            s = str(v).replace('%','').strip()
            if s in ('nan','NaN',''): return None
            try: return round(float(s), 4)
            except: return None

        # ------ BAGIAN 1: Definisi KPI (metadata saja, tidak ada historical) ------
        kpi_definitions = []
        for row in range(3, 22):
            no_v = df.iloc[row, 1]
            if pd.isna(no_v): continue
            try: no_int = int(float(str(no_v)))
            except: continue
            kpi_name = str(df.iloc[row, 2]).strip()
            if not kpi_name or kpi_name == 'nan': continue
            variable    = str(df.iloc[row, 3]).strip()
            unit        = str(df.iloc[row, 4]).strip()
            target_desc = str(df.iloc[row, 5]).strip()
            is_lower    = any(k in kpi_name.lower() for k in ['pelanggaran','turn over'])
            kpi_definitions.append({
                'no': no_int, 'kpi': kpi_name, 'variable': variable,
                'unit': unit, 'target_desc': target_desc, 'is_lower_better': is_lower
            })

        # ------ BAGIAN 2: Monthly 2026 ------
        monthly_2026 = []
        for i in range(26, 44):
            row_data = df.iloc[i]
            if pd.isna(row_data[1]) or pd.isna(row_data[2]): continue
            kpi_name = str(row_data[2]).strip()
            if not kpi_name or kpi_name == 'nan': continue
            try: no_int = int(float(str(row_data[1])))
            except: continue

            bulan_dt  = pd.to_datetime(row_data[3], errors='coerce')
            bulan_str = bulan_dt.strftime('%b %Y') if pd.notna(bulan_dt) else None
            target    = sf(row_data[4])
            actual    = sf(row_data[5])

            is_lower = any(k in kpi_name.lower() for k in ['pelanggaran','turn over'])
            status = None
            if target is not None and actual is not None:
                status = 'tercapai' if (actual <= target if is_lower else actual >= target) else 'tidak_tercapai'

            # Cari metadata unit dari kpi_definitions berdasarkan nomor urut KPI (1-9)
            kpi_no_actual = ((no_int - 1) % 9) + 1  # mapping no sekuensial ke 1-9
            meta = next((d for d in kpi_definitions if d['no'] == kpi_no_actual), None)

            monthly_2026.append({
                'no':     kpi_no_actual,
                'kpi':    kpi_name,
                'bulan':  bulan_str,
                'target': target,
                'actual': actual,
                'status': status,
                'unit':   meta['unit'] if meta else ''
            })

        return kpi_definitions, monthly_2026

    except Exception as e:
        print("ERROR LOAD KPI:", e)
        import traceback; traceback.print_exc()
        return [], []


# =======================================================
# FILTERS
# =======================================================
def apply_filters(df):
    i = request.args.get("institusi"); p = request.args.get("penempatan")
    s = request.args.get("status"); st = request.args.get("start"); en = request.args.get("end")
    if i:  df = df[df['institusi']  == i]
    if p:  df = df[df['penempatan'] == p]
    if s:  df = df[df['status']     == s.lower()]
    if st and en:
        df = df[(df['permohonan'] >= pd.to_datetime(st)) & (df['permohonan'] <= pd.to_datetime(en))]
    return df

def apply_manpower_filters(df):
    b = request.args.get("branch"); j = request.args.get("jabatan")
    sk = request.args.get("status_kontrak"); gp = request.args.get("group_project")
    if b:  df = df[df['branch']        == b]
    if j:  df = df[df['jabatan']        == j]
    if sk: df = df[df['status_kontrak'] == sk.lower()]
    if gp: df = df[df['group_project']  == gp]
    return df

def apply_recruitment_filters(df):
    s = request.args.get("status"); p = request.args.get("posisi")
    st = request.args.get("start"); en = request.args.get("end")
    if s:  df = df[df['status'] == s.lower()]
    if p:  df = df[df['posisi'] == p]
    if st and en:
        df = df[(df['periode'] >= pd.to_datetime(st)) & (df['periode'] <= pd.to_datetime(en))]
    return df

def apply_salary_filters(df, column_periode):
    p = request.args.get("periode"); pr = request.args.get("project")
    if p:  df = df[df[column_periode] == p]
    if pr: df = df[df['project'] == pr]
    return df

def apply_training_filters(df):
    j = request.args.get("jenis"); d = request.args.get("divisi")
    s = request.args.get("status"); st = request.args.get("start"); en = request.args.get("end")
    if j:  df = df[df['jenis']  == j.lower()]
    if d:  df = df[df['divisi'] == d]
    if s:  df = df[df['status'] == s.lower()]
    if st and en:
        df = df[(df['periode'] >= pd.to_datetime(st)) & (df['periode'] <= pd.to_datetime(en))]
    return df

def apply_bpjs_filters(df_k, df_kes, df_perm, df_bor):
    pr = request.args.get("project"); j = request.args.get("jenis_bpjs"); p = request.args.get("periode")
    if pr: df_k = df_k[df_k['project'] == pr]
    if j:  df_k = df_k[df_k['jenis_bpjs'] == j.lower()]
    if p:
        df_kes  = df_kes[df_kes['periode']   == p]
        df_perm = df_perm[df_perm['periode'] == p]
        df_bor  = df_bor[df_bor['periode']   == p]
    return df_k, df_kes, df_perm, df_bor

def apply_overtime_filters(df_ot, df_abs, df_cost):
    pr = request.args.get("project"); st = request.args.get("start"); en = request.args.get("end")
    if pr:
        df_ot   = df_ot[df_ot['project']   == pr]
        df_cost = df_cost[df_cost['project'] == pr]
    if st and en:
        s = pd.to_datetime(st); e = pd.to_datetime(en)
        df_ot   = df_ot[(df_ot['bulan']     >= s) & (df_ot['bulan']     <= e)]
        df_abs  = df_abs[(df_abs['periode'] >= s) & (df_abs['periode']  <= e)]
        df_cost = df_cost[(df_cost['bulan'] >= s) & (df_cost['bulan']   <= e)]
    return df_ot, df_abs, df_cost

def apply_mcu_filters(df_mcu, df_pay):
    pr = request.args.get("project"); h = request.args.get("hasil")
    st = request.args.get("start");   en = request.args.get("end")
    if pr: df_mcu = df_mcu[df_mcu['project'] == pr]; df_pay = df_pay[df_pay['project'] == pr]
    if h:  df_mcu = df_mcu[df_mcu['hasil_mcu'] == h.lower()]
    if st and en:
        df_mcu = df_mcu[(df_mcu['periode'] >= pd.to_datetime(st)) & (df_mcu['periode'] <= pd.to_datetime(en))]
    return df_mcu, df_pay

def apply_kpi_filters(monthly_2026):
    bulan  = request.args.get("bulan")   # "Jan 2026"
    kpi_no = request.args.get("no")      # "1"-"9"
    status = request.args.get("status")  # tercapai / tidak_tercapai
    result = monthly_2026
    if bulan:  result = [r for r in result if r['bulan'] == bulan]
    if kpi_no:
        try: n = int(kpi_no); result = [r for r in result if r['no'] == n]
        except: pass
    if status: result = [r for r in result if r['status'] == status]
    return result


# =======================================================
# KPI CALCULATORS
# =======================================================
def calculate_kpi(df):
    today = pd.to_datetime(datetime.today().date()); batas = today + timedelta(days=14)
    ab = df[(df['berakhir'].notna()) & (df['berakhir'] >= today) & (df['berakhir'] <= batas)]
    return {"onboard": len(df[df['status']=='onboard']),
            "butuh_surat_balasan": len(df[df['status']=='butuh surat balasan']),
            "ajukan_ulang": len(df[df['status']=='ajukan ulang']),
            "selesai": len(df[df['status']=='selesai']),
            "akan_berakhir": len(ab), "total": len(df)}

def calculate_manpower_kpi(df):
    return {"total": len(df), "permanent": len(df[df['status_kontrak']=='permanent']),
            "kontrak": len(df[df['status_kontrak']=='kontrak'])}

def calculate_recruitment_kpi(df):
    return {"total": len(df), "accepted": len(df[df['status']=='accepted']),
            "rejected": len(df[df['status']=='rejected']),
            "on_process": len(df[df['status']=='on process'])}

def calculate_salary_kpi(df_gaji, df_pph):
    return {"total_gaji": int(df_gaji['gaji'].sum()), "total_pph21": int(df_pph['pph21'].sum()),
            "jumlah_project": df_gaji['project'].nunique()}

def calculate_training_kpi(df):
    return {"total": len(df), "internal": len(df[df['jenis']=='internal']),
            "external": len(df[df['jenis']=='external']),
            "done": len(df[df['status']=='done']),
            "in_progress": len(df[df['status']=='in progress']),
            "cancel": len(df[df['status']=='cancel'])}

def calculate_bpjs_kpi(df_kes, df_perm, df_bor, df_k):
    return {"bpjs_kesehatan": int(df_kes['pembayaran'].sum()),
            "bpjs_tk_permanent": int(df_perm['pembayaran'].sum()),
            "bpjs_tk_borongan": int(df_bor['pembayaran'].sum()),
            "total_semua": int(df_kes['pembayaran'].sum()+df_perm['pembayaran'].sum()+df_bor['pembayaran'].sum()),
            "total_karyawan_bpjs": len(df_k)}

def calculate_mcu_kpi(df, df_pay):
    total = len(df)
    fit   = len(df[df['hasil_mcu'].str.contains('fit',  na=False)])
    unfit = len(df[df['hasil_mcu'].str.contains('unfit',na=False)])
    return {"total_mcu": total, "fit": fit, "unfit": unfit,
            "fit_rate": round(fit/total*100, 2) if total > 0 else 0,
            "total_cost": int(df_pay['pembayaran'].sum())}


# =======================================================
# DASHBOARD INTERNSHIP
# =======================================================
@app.route("/internship/dashboard")
def get_dashboard():
    df = load_data(); df = apply_filters(df); kpi = calculate_kpi(df)
    institusi  = df['institusi'].value_counts().reset_index(); institusi.columns  = ['institusi','jumlah']
    permohonan = df['bulan'].value_counts().sort_index().reset_index(); permohonan.columns = ['bulan','jumlah']
    penempatan = df['penempatan'].value_counts().reset_index(); penempatan.columns = ['penempatan','jumlah']
    ket        = df['ket'].value_counts().reset_index(); ket.columns = ['ket','jumlah']
    return jsonify({"kpi": kpi, "institusi": institusi.to_dict(orient='records'),
                    "permohonan": permohonan.to_dict(orient='records'),
                    "penempatan": penempatan.to_dict(orient='records'),
                    "ket": ket.to_dict(orient='records')})


# =======================================================
# DASHBOARD MANPOWER
# =======================================================
@app.route("/manpower/dashboard")
def manpower_dashboard():
    df = load_manpower(); df = apply_manpower_filters(df); kpi = calculate_manpower_kpi(df)
    jabatan       = df['jabatan'].value_counts().reset_index(); jabatan.columns       = ['jabatan','jumlah']
    branch        = df['branch'].value_counts().reset_index(); branch.columns         = ['branch','jumlah']
    group_project = df['group_project'].value_counts().reset_index(); group_project.columns = ['group_project','jumlah']
    status_kontrak= df['status_kontrak'].value_counts().reset_index(); status_kontrak.columns = ['status_kontrak','jumlah']
    return jsonify({"kpi": kpi, "jabatan": jabatan.to_dict(orient='records'),
                    "branch": branch.to_dict(orient='records'),
                    "group_project": group_project.to_dict(orient='records'),
                    "status_kontrak": status_kontrak.to_dict(orient='records')})


# =======================================================
# DASHBOARD RECRUITMENT
# =======================================================
@app.route("/recruitment/dashboard")
def recruitment_dashboard():
    df = load_recruitment(); df = apply_recruitment_filters(df); kpi = calculate_recruitment_kpi(df)
    status = df['status'].value_counts().reset_index(); status.columns = ['status','jumlah']
    posisi = df['posisi'].value_counts().reset_index(); posisi.columns = ['posisi','jumlah']
    trend  = df['bulan'].value_counts().sort_index().reset_index(); trend.columns = ['bulan','jumlah']
    funnel_order = ['on process','mcu','accepted','rejected','mengundurkan diri']
    fd = df['status'].value_counts()
    funnel = [{"status": f, "jumlah": int(fd.get(f, 0))} for f in funnel_order]
    return jsonify({"kpi": kpi, "status": status.to_dict(orient='records'),
                    "posisi": posisi.to_dict(orient='records'),
                    "trend": trend.to_dict(orient='records'), "funnel": funnel})


# =======================================================
# DASHBOARD SALARY
# =======================================================
@app.route("/salary/dashboard")
def salary_dashboard():
    df_gaji, df_pph = load_salary()
    df_gaji = apply_salary_filters(df_gaji, "periode")
    df_pph  = apply_salary_filters(df_pph,  "periode")
    kpi     = calculate_salary_kpi(df_gaji, df_pph)

    gaji_project = df_gaji.groupby('project')['gaji'].sum().reset_index().sort_values('gaji', ascending=False)
    gaji_trend   = df_gaji.groupby('periode')['gaji'].sum().reset_index()
    gaji_trend['_s'] = pd.to_datetime(gaji_trend['periode'], format='%b %Y')
    gaji_trend   = gaji_trend.sort_values('_s').drop('_s', axis=1)

    pph_project  = df_pph.groupby('project')['pph21'].sum().reset_index().sort_values('pph21', ascending=False)
    pph_trend    = df_pph.groupby('periode')['pph21'].sum().reset_index()
    pph_trend['_s'] = pd.to_datetime(pph_trend['periode'], format='%b %Y')
    pph_trend    = pph_trend.sort_values('_s').drop('_s', axis=1)

    return jsonify({"kpi": kpi, "gaji_project": gaji_project.to_dict(orient='records'),
                    "gaji_trend": gaji_trend.to_dict(orient='records'),
                    "pph_project": pph_project.to_dict(orient='records'),
                    "pph_trend": pph_trend.to_dict(orient='records')})


# =======================================================
# DASHBOARD TRAINING
# =======================================================
@app.route("/training/dashboard")
def training_dashboard():
    df = load_training(); df = apply_training_filters(df); kpi = calculate_training_kpi(df)
    jenis     = df['jenis'].value_counts().reset_index(); jenis.columns     = ['jenis','jumlah']
    divisi    = df['divisi'].value_counts().reset_index(); divisi.columns    = ['divisi','jumlah']
    pelatihan = df['pelatihan'].value_counts().reset_index(); pelatihan.columns = ['pelatihan','jumlah']
    status    = df['status'].value_counts().reset_index(); status.columns    = ['status','jumlah']
    trend     = df['bulan'].value_counts().sort_index().reset_index(); trend.columns = ['bulan','jumlah']
    return jsonify({"kpi": kpi, "jenis": jenis.to_dict(orient='records'),
                    "divisi": divisi.to_dict(orient='records'),
                    "pelatihan": pelatihan.to_dict(orient='records'),
                    "status": status.to_dict(orient='records'),
                    "trend": trend.to_dict(orient='records')})


# =======================================================
# DASHBOARD BPJS
# =======================================================
@app.route("/bpjs/dashboard")
def bpjs_dashboard():
    df_k, df_kes, df_perm, df_bor = load_bpjs()
    df_k, df_kes, df_perm, df_bor = apply_bpjs_filters(df_k, df_kes, df_perm, df_bor)
    kpi = calculate_bpjs_kpi(df_kes, df_perm, df_bor, df_k)

    jenis   = df_k['jenis_bpjs'].value_counts().reset_index(); jenis.columns   = ['jenis_bpjs','jumlah']
    project = df_k['project'].value_counts().reset_index(); project.columns     = ['project','jumlah']

    def mk_trend(df, label):
        d = df.copy(); d['ps'] = d['periode'].dt.strftime('%Y-%m')
        t = d.groupby('ps')['pembayaran'].sum().reset_index()
        t.columns = ['periode','pembayaran']; t['pembayaran'] = t['pembayaran'].astype(int)
        t['kategori'] = label; return t

    trend = pd.concat([mk_trend(df_kes,'kesehatan'), mk_trend(df_perm,'tk_permanent'), mk_trend(df_bor,'tk_borongan')])
    return jsonify({"kpi": kpi, "jenis": jenis.to_dict(orient='records'),
                    "project": project.to_dict(orient='records'),
                    "trend": trend.to_dict(orient='records')})


# =======================================================
# DASHBOARD OVERTIME
# =======================================================
@app.route("/overtime/dashboard")
def overtime_dashboard():
    df_ot, df_abs, df_cost = load_overtime()
    df_ot, df_abs, df_cost = apply_overtime_filters(df_ot, df_abs, df_cost)

    avg_ot_pct  = df_ot['overtime'].mean()  * 100 if not df_ot.empty  else 0
    avg_abs_pct = df_abs['absensi'].mean()  * 100 if not df_abs.empty else 0
    t_cost      = df_cost['total_cost'].sum()
    t_ot_cost   = df_cost['overtime_cost'].sum()

    m_ov  = df_ot.groupby(df_ot['bulan'].dt.to_period('M'))['overtime'].mean().sort_index()
    ov_g  = round(((m_ov.iloc[-1]-m_ov.iloc[-2])/m_ov.iloc[-2])*100, 2) if len(m_ov)>=2 else 0
    m_at  = df_abs.groupby(df_abs['periode'].dt.to_period('M'))['absensi'].mean().sort_index()
    at_g  = round(((m_at.iloc[-1]-m_at.iloc[-2])/m_at.iloc[-2])*100, 2) if len(m_at)>=2 else 0

    kpi = {"avg_overtime_percent": round(avg_ot_pct,2), "avg_absensi_percent": round(avg_abs_pct,2),
           "total_cost": int(t_cost), "total_overtime_cost": int(t_ot_cost),
           "overtime_growth": ov_g, "absensi_growth": at_g}

    ov_sum = df_ot.groupby(df_ot['bulan'].dt.strftime('%b %Y'))['overtime'].mean().reset_index()
    ov_sum.columns = ['bulan','overtime']; ov_sum['overtime'] = ov_sum['overtime'].round(4)
    ov_sum = ov_sum.sort_values('bulan', key=lambda x: pd.to_datetime(x, format='%b %Y'))

    at_sum = df_abs.groupby(df_abs['periode'].dt.strftime('%b %Y'))['absensi'].mean().reset_index()
    at_sum.columns = ['bulan','absensi']; at_sum['absensi'] = at_sum['absensi'].round(4)
    at_sum = at_sum.sort_values('bulan', key=lambda x: pd.to_datetime(x, format='%b %Y'))

    ov_proj = df_ot.groupby('project')['overtime'].mean().reset_index()
    ov_proj['overtime_percent'] = (ov_proj['overtime']*100).round(2)
    ov_proj = ov_proj.sort_values('overtime_percent', ascending=False)

    cp = df_cost.groupby('project').agg(total_cost=('total_cost','sum'),
                                         overtime_cost=('overtime_cost','sum'),
                                         overtime_percent=('overtime_percent','mean')).reset_index()
    for c in ['total_cost','overtime_cost','overtime_percent']: cp[c] = cp[c].round(2)

    top5 = cp.nlargest(5,'overtime_cost')[['project','overtime_cost']].round(2)

    ct = df_cost.groupby(df_cost['bulan'].dt.strftime('%b %Y'))['overtime_cost'].sum().reset_index()
    ct.columns = ['bulan','overtime_cost']; ct['overtime_cost'] = ct['overtime_cost'].round(2)
    ct = ct.sort_values('bulan', key=lambda x: pd.to_datetime(x, format='%b %Y'))

    return jsonify({"kpi": kpi, "overtime_summary": ov_sum.to_dict(orient='records'),
                    "attendance_summary": at_sum.to_dict(orient='records'),
                    "overtime_by_project": ov_proj.to_dict(orient='records'),
                    "cost_by_project": cp.to_dict(orient='records'),
                    "top_cost_projects": top5.to_dict(orient='records'),
                    "overtime_cost_trend": ct.to_dict(orient='records')})


# =======================================================
# DASHBOARD MCU
# =======================================================
@app.route("/mcu/dashboard")
def mcu_dashboard():
    df_mcu, df_karyawan, df_pay = load_mcu()
    df_full = pd.merge(df_mcu, df_karyawan, on="nama", how="left")
    df_full, df_pay = apply_mcu_filters(df_full, df_pay)
    kpi = calculate_mcu_kpi(df_full, df_pay)

    hasil        = df_full['hasil_mcu'].value_counts().reset_index(); hasil.columns   = ['hasil','jumlah']
    divisi       = df_full['divisi'].value_counts().reset_index(); divisi.columns      = ['divisi','jumlah']
    project      = df_full['project'].value_counts().reset_index(); project.columns    = ['project','jumlah']
    gender       = df_full['gender'].value_counts().reset_index(); gender.columns      = ['gender','jumlah']

    trend = df_full['bulan'].value_counts().reset_index(); trend.columns = ['bulan','jumlah']
    trend['_s'] = pd.to_datetime(trend['bulan'], format='%b %Y', errors='coerce')
    trend = trend.sort_values('_s').drop('_s', axis=1)

    cost_project = df_pay.groupby('project')['pembayaran'].sum().reset_index()
    cost_trend   = df_pay.groupby('periode')['pembayaran'].sum().reset_index()
    cost_trend['_s'] = pd.to_datetime(cost_trend['periode'], format='%b %Y', errors='coerce')
    cost_trend   = cost_trend.sort_values('_s').drop('_s', axis=1)

    top_project  = cost_project.sort_values('pembayaran', ascending=False).head(5)

    return jsonify({"kpi": kpi, "hasil_mcu": hasil.to_dict(orient='records'),
                    "divisi": divisi.to_dict(orient='records'),
                    "project": project.to_dict(orient='records'),
                    "trend": trend.to_dict(orient='records'),
                    "gender": gender.to_dict(orient='records'),
                    "cost_project": cost_project.to_dict(orient='records'),
                    "cost_trend": cost_trend.to_dict(orient='records'),
                    "top_project": top_project.to_dict(orient='records')})


@app.route("/kpi/dashboard")
def kpi_dashboard():
    kpi_definitions, monthly_2026 = load_kpi_data()
    filtered = apply_kpi_filters(monthly_2026)

    # ---- Summary (dari data yang difilter) ----
    total    = len(filtered)
    tercapai = sum(1 for r in filtered if r['status'] == 'tercapai')
    tidak    = sum(1 for r in filtered if r['status'] == 'tidak_tercapai')
    summary  = {
        "total_kpi":      total,
        "tercapai":       tercapai,
        "tidak_tercapai": tidak,
        "pct_tercapai":   round(tercapai / total * 100, 1) if total > 0 else 0
    }

    # ---- KPI List — gabung definisi + hitung achievement dari semua monthly_2026 ----
    kpi_list = []
    for defn in kpi_definitions:
        # Semua data bulan untuk KPI ini (dari seluruh monthly_2026, tidak difilter)
        rows_kpi  = [r for r in monthly_2026 if r['no'] == defn['no']]
        rows_ada  = [r for r in rows_kpi if r['actual'] is not None]
        rows_ok   = [r for r in rows_ada if r['status'] == 'tercapai']
        kpi_list.append({
            "no":              defn['no'],
            "kpi":             defn['kpi'],
            "unit":            defn['unit'],
            "target_desc":     defn['target_desc'],
            "is_lower_better": defn['is_lower_better'],
            "bulan_ada":       len(rows_ada),
            "bulan_ok":        len(rows_ok),
            "pct_tercapai":    round(len(rows_ok) / len(rows_ada) * 100, 1) if rows_ada else 0,
            "monthly":         rows_kpi    # data per bulan untuk KPI ini
        })

    # ---- Status per Bulan — dari seluruh monthly_2026 (bukan filtered) ----
    sm = {}
    for r in monthly_2026:
        b = r['bulan']
        if b not in sm:
            sm[b] = {"bulan": b, "tercapai": 0, "tidak_tercapai": 0, "total": 0}
        sm[b]['total'] += 1
        if r['status'] == 'tercapai':         sm[b]['tercapai'] += 1
        elif r['status'] == 'tidak_tercapai': sm[b]['tidak_tercapai'] += 1
    status_trend = sorted(sm.values(), key=lambda x: x['bulan'])

    # ---- Bulan tersedia ----
    bulan_options = sorted(list({r['bulan'] for r in monthly_2026 if r['bulan']}))

    return jsonify({
        "summary":       summary,
        "kpi_list":      kpi_list,
        "monthly_2026":  filtered,
        "status_trend":  status_trend,
        "bulan_options": bulan_options
    })


# =======================================================
# RUN
# =======================================================
if __name__ == "__main__":
    app.run(debug=True)
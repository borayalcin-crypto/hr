import os
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(layout="wide", page_title="İK Dashboard")

DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)

def get_existing_data_path():
    for fname in os.listdir(DATA_DIR):
        if fname.startswith("son_veri."):
            return os.path.join(DATA_DIR, fname)
    return None

UPLOAD_PASSWORD = st.secrets.get("upload_password", "ik2026") if hasattr(st, "secrets") else "ik2026"

st.markdown("""
<style>
[data-testid="stMetricValue"] { font-size: 1.0rem; }
[data-testid="stMetricLabel"] { font-size: 0.72rem; }
[data-testid="stMetricDelta"] { font-size: 0.65rem; }
[data-testid="stMetric"] { padding: 0.35rem 0.25rem; }
</style>
""", unsafe_allow_html=True)

COMPANIES = [
    'Aralık Sigorta', 'Ekim Turizm', 'Eylül Girişim',
    'Haziran Servis', 'Intercity Yatırım Holding', 'Mart Denizcilik'
]
MONTHS = ['Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz', 'Ağustos']

# ----- FORMAT YARDIMCILARI -----
def format_tl(value):
    if value is None:
        return "0TL"
    try:
        formatted = f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"{formatted}TL"
    except:
        return f"{value:.2f}TL".replace(".", ",")

def format_tl_no_decimal(value):
    if value is None:
        return "0TL"
    try:
        formatted = f"{value:,.0f}".replace(",", ".")
        return f"{formatted}TL"
    except:
        return f"{value:.0f}TL"

def format_percent(value):
    if value is None:
        return "%0"
    try:
        formatted = f"{value:.2f}".replace(".", ",")
        return f"%{formatted}"
    except:
        return f"%{value:.2f}".replace(".", ",")

def format_number(value, decimals=1):
    if value is None:
        return "0"
    try:
        formatted = f"{value:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return formatted
    except:
        return str(value)

def calc_diff(current, previous):
    if previous is None:
        return None
    return round(current - previous, 4)

def format_delta_tl(diff):
    if diff is None:
        return None
    sign = "+" if diff >= 0 else "-"
    return f"{sign}{format_tl(abs(diff))}"

def format_delta_percent(diff):
    if diff is None:
        return None
    sign = "+" if diff >= 0 else "-"
    return f"{sign}{format_percent(abs(diff))}"

def format_delta_number(diff, decimals=0):
    if diff is None:
        return None
    sign = "+" if diff >= 0 else "-"
    return f"{sign}{format_number(abs(diff), decimals)}"

# ----- NORMALİZASYON -----
def normalize_company_name(name):
    if not isinstance(name, str):
        return name
    name = name.strip()
    replacements = {
        'EKIM TURIZM': 'Ekim Turizm',
        'HAZIRAN': 'Haziran Servis',
        'Holding': 'Intercity Yatırım Holding'
    }
    return replacements.get(name, name)

def clean_columns(df):
    df.columns = [str(col).strip() for col in df.columns]
    return df

def get_month_cols(df):
    month_cols = []
    for col in df.columns:
        col_clean = str(col).strip()
        for m in MONTHS:
            if col_clean.lower() == m.lower():
                month_cols.append(col)
                break
    return month_cols

def clean_numeric_df(df):
    return df.apply(pd.to_numeric, errors='coerce').fillna(0)

def read_company_month_sheet(uploaded_file, sheet_name, total_label, agg='sum', months=None):
    if months is None:
        months = MONTHS
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=0)
    df = clean_columns(df)
    first_col = df.columns[0]
    df[first_col] = df[first_col].apply(normalize_company_name)
    df = df.set_index(first_col)
    month_cols = get_month_cols(df)

    is_total_row = df.index.astype(str).str.strip().str.upper() == total_label.strip().upper()
    total_row = df[is_total_row].iloc[0] if is_total_row.any() else None
    df_companies = df[~is_total_row].copy()
    for m in month_cols:
        df_companies[m] = pd.to_numeric(df_companies[m], errors='coerce').fillna(0)

    monthly_totals = {}
    for m in months:
        if m not in month_cols:
            monthly_totals[m] = 0
            continue
        sheet_val = total_row.get(m) if total_row is not None else None
        if sheet_val is not None and pd.notna(sheet_val):
            monthly_totals[m] = sheet_val
        else:
            series = df_companies[m]
            monthly_totals[m] = series.sum() if agg == 'sum' else series.mean()
    return df_companies, monthly_totals

def safe_read_son(uploaded_file, sheet, skip, n_months=None):
    if n_months is None:
        n_months = len(MONTHS)
    try:
        df = pd.read_excel(uploaded_file, sheet_name=sheet, skiprows=skip, header=None, nrows=1)
        if df.shape[1] >= n_months + 1:
            return df.iloc[0, 1:n_months + 1].values
        else:
            return [0] * n_months
    except:
        return [0] * n_months

@st.cache_data(show_spinner=False)
def load_data(_uploaded_file, cache_key=None):
    uploaded_file = _uploaded_file

    # ----- 1. KÜMÜLATİF GENEL TURNOVER -----
    df_gt, gt_monthly_totals = read_company_month_sheet(uploaded_file, 'genel.turnover', 'Genel Toplam', agg='sum')

    # ----- 2. KÜMÜLATİF GÖNÜLLÜ TURNOVER -----
    df_gon, gon_monthly_totals = read_company_month_sheet(uploaded_file, 'gonullu.turnover', 'Genel Toplam', agg='sum')

    # ----- 3. AYLIK GENEL TURNOVER -----
    df_aylik_gt, aylik_gt_totals = read_company_month_sheet(uploaded_file, 'aylik.turnover', 'Genel Toplam', agg='sum')

    # ----- 4. AYLIK GÖNÜLLÜ TURNOVER -----
    df_aylik_gon, aylik_gon_totals = read_company_month_sheet(uploaded_file, 'aylik.gonullu.turnover', 'Genel Toplam', agg='sum')

    # ----- 5. RAPOR ORANI -----
    df_rapor = pd.read_excel(uploaded_file, sheet_name='rapor_oran', header=0)
    df_rapor = clean_columns(df_rapor)
    month_cols = get_month_cols(df_rapor)
    df_rapor['Şirket'] = df_rapor.iloc[:, 0].apply(normalize_company_name)
    df_rapor = clean_numeric_df(df_rapor.set_index('Şirket')[month_cols])

    # ----- 6. ÇALIŞAN SAYISI -----
    df_calisan = pd.read_excel(uploaded_file, sheet_name='calisan.sayisi', header=0)
    df_calisan = clean_columns(df_calisan)
    month_cols = get_month_cols(df_calisan)
    df_calisan['Şirket'] = df_calisan.iloc[:, 0].apply(normalize_company_name)
    df_calisan = clean_numeric_df(df_calisan.set_index('Şirket')[month_cols])

    # ----- 7. NET KÖK ÜCRET -----
    df_net = pd.read_excel(uploaded_file, sheet_name='kok.ucret', header=0)
    df_net = clean_columns(df_net)
    month_cols = get_month_cols(df_net)
    df_net['Şirket'] = df_net.iloc[:, 0].apply(normalize_company_name)
    df_net = clean_numeric_df(df_net.set_index('Şirket')[month_cols])

    # ----- 8. İŞVEREN MALİYETİ -----
    df_isv = pd.read_excel(uploaded_file, sheet_name='isveren.maliyet', header=0)
    df_isv = clean_columns(df_isv)
    month_cols = get_month_cols(df_isv)
    df_isv['Şirket'] = df_isv.iloc[:, 0].apply(normalize_company_name)
    df_isv = clean_numeric_df(df_isv.set_index('Şirket')[month_cols])

    # ----- 9. FM SAAT -----
    df_fm_saat = pd.read_excel(uploaded_file, sheet_name='fm.saat', header=0)
    df_fm_saat = clean_columns(df_fm_saat)
    month_cols = get_month_cols(df_fm_saat)
    df_fm_saat['Şirket'] = df_fm_saat.iloc[:, 0].apply(normalize_company_name)
    df_fm_saat = clean_numeric_df(df_fm_saat.set_index('Şirket')[month_cols])

    # ----- 10. FM TL MALİYET -----
    df_fm_tl = pd.read_excel(uploaded_file, sheet_name='fm.maliyet', header=0)
    df_fm_tl = clean_columns(df_fm_tl)
    month_cols = get_month_cols(df_fm_tl)
    df_fm_tl['Şirket'] = df_fm_tl.iloc[:, 0].apply(normalize_company_name)
    df_fm_tl = clean_numeric_df(df_fm_tl.set_index('Şirket')[month_cols])

    # ----- 11. İZİN GÜN -----
    df_izin_gun = pd.read_excel(uploaded_file, sheet_name='izin_gun', header=0)
    df_izin_gun = clean_columns(df_izin_gun)
    month_cols = get_month_cols(df_izin_gun)
    df_izin_gun['Şirket'] = df_izin_gun.iloc[:, 0].apply(normalize_company_name)
    df_izin_gun = clean_numeric_df(df_izin_gun.set_index('Şirket')[month_cols])

    # ----- 12. İZİN ÜCRET -----
    df_izin_ucret = pd.read_excel(uploaded_file, sheet_name='izin_ucret', header=0)
    df_izin_ucret = clean_columns(df_izin_ucret)
    month_cols = get_month_cols(df_izin_ucret)
    df_izin_ucret['Şirket'] = df_izin_ucret.iloc[:, 0].apply(normalize_company_name)
    df_izin_ucret = clean_numeric_df(df_izin_ucret.set_index('Şirket')[month_cols])

    # ----- 13. KIDEM TAZMİNATI -----
    df_kidem, kidem_totals = read_company_month_sheet(uploaded_file, 'kidem.tazminati', 'TOPLAM', agg='sum')

    # ----- 14. İHBAR TAZMİNATI -----
    df_ihbar, ihbar_totals = read_company_month_sheet(uploaded_file, 'ihbar.tazminati', 'TOPLAM', agg='sum')

    # ----- 15. KİŞİ BAŞI ORTALAMA MAAŞ (aylık) - kisi.basi.ort sayfasındaki "Genel Kişi Başı Ortalama Maaş" satırı -----
    df_kisi_basi, kisi_basi_genel = read_company_month_sheet(
        uploaded_file, 'kisi.basi.ort', 'Genel Kişi Başı Ortalama Maaş', agg='mean'
    )
    # Yıllık Ortalama sütununu bul (şirket bazlı yıllık ortalama için)
    yillik_ort_col = None
    for col in df_kisi_basi.columns:
        if str(col).strip().lower() in ('yıllık ortalama', 'yillik ortalama'):
            yillik_ort_col = col
            break

    # ----- 16. gnl.kisi.basi.ort sayfasından Genel Kişi Başı Ortalama Maaş (kümülatif ortalama) -----
    try:
        df_gnl = pd.read_excel(uploaded_file, sheet_name='gnl.kisi.basi.ort', header=0)
        df_gnl = clean_columns(df_gnl)
        # "Genel Kişi Başı Ortalama Maaş" satırını bul
        genel_satir = df_gnl[df_gnl.iloc[:, 0].astype(str).str.strip().str.upper() == 'GENEL KİŞİ BAŞI ORTALAMA MAAŞ']
        if not genel_satir.empty:
            month_cols = get_month_cols(genel_satir)
            gnl_kisi_basi_vals = {}
            for m in MONTHS:
                if m in genel_satir.columns:
                    val = pd.to_numeric(genel_satir.iloc[0][m], errors='coerce')
                    gnl_kisi_basi_vals[m] = val if pd.notna(val) else 0
                else:
                    gnl_kisi_basi_vals[m] = 0
        else:
            gnl_kisi_basi_vals = {m: 0 for m in MONTHS}
    except Exception:
        gnl_kisi_basi_vals = {m: 0 for m in MONTHS}

    # ----- 17. KADIN ORANI -----
    df_kadin, kadin_genel = read_company_month_sheet(uploaded_file, 'kadin.erkek', 'Genel Kadın Oranı', agg='mean')

    # ----- 18. İLK 6 AY AYRILMA ORANI -----
    df_ilk6ay, ilk6ay_ortalama = read_company_month_sheet(uploaded_file, 'ilk.6ay', '__YOK__', agg='mean')

    # ----- 19. AY İÇİ GİRİŞ -----
    df_giris, giris_toplam = read_company_month_sheet(uploaded_file, 'aylik.giris', '__YOK__', agg='sum')

    # ----- 20. AY İÇİ ÇIKIŞ -----
    df_cikis, cikis_toplam = read_company_month_sheet(uploaded_file, 'aylik.cikis', '__YOK__', agg='sum')

    # ----- TOPLAM SATIRLARI (rapor_oran, calisan.sayisi, izin_gun, izin_ucret) -----
    genel_rapor = safe_read_son(uploaded_file, 'rapor_oran', 7)
    toplam_calisan = safe_read_son(uploaded_file, 'calisan.sayisi', 7)
    toplam_izin_gun = safe_read_son(uploaded_file, 'izin_gun', 7)
    toplam_izin_ucret = safe_read_son(uploaded_file, 'izin_ucret', 7)

    # ----- FM YAPAN LİSTESİ -----
    df_fm_yapan = pd.read_excel(uploaded_file, sheet_name='aylik.fm.yapan', header=0)
    df_fm_yapan = clean_columns(df_fm_yapan)
    if 'Şirket' in df_fm_yapan.columns:
        df_fm_yapan['Şirket'] = df_fm_yapan['Şirket'].apply(normalize_company_name)
    for m in get_month_cols(df_fm_yapan):
        df_fm_yapan[m] = pd.to_numeric(df_fm_yapan[m], errors='coerce').fillna(0)

    # ----- VERİYİ BİRLEŞTİR -----
    data = {}
    for idx, m in enumerate(MONTHS):
        comp_data = {}
        for comp in COMPANIES:
            if comp in df_calisan.index:
                comp_data[comp] = {
                    'employees': df_calisan.loc[comp, m],
                    'devamsizlik': df_rapor.loc[comp, m] * 100,
                    'turnoverKumulatif': df_gt.loc[comp, m] * 100 if comp in df_gt.index else 0,
                    'turnoverGonulluKumulatif': df_gon.loc[comp, m] * 100 if comp in df_gon.index else 0,
                    'turnoverAylik': df_aylik_gt.loc[comp, m] * 100 if comp in df_aylik_gt.index else 0,
                    'turnoverGonulluAylik': df_aylik_gon.loc[comp, m] * 100 if comp in df_aylik_gon.index else 0,
                    'netKokUcret': df_net.loc[comp, m],
                    'isverenMaliyet': df_isv.loc[comp, m],
                    'fmSaat': df_fm_saat.loc[comp, m],
                    'fmTlMaliyet': df_fm_tl.loc[comp, m],
                    'izinGun': df_izin_gun.loc[comp, m],
                    'izinUcret': df_izin_ucret.loc[comp, m],
                    'kisiBasiOrt': df_kisi_basi.loc[comp, m] if comp in df_kisi_basi.index and m in df_kisi_basi.columns else 0,
                    'yillikOrtMaas': (df_kisi_basi.loc[comp, yillik_ort_col]
                                      if yillik_ort_col is not None and comp in df_kisi_basi.index else 0),
                    'kadinOrani': df_kadin.loc[comp, m] * 100 if comp in df_kadin.index and m in df_kadin.columns else 0,
                    'ilk6ayOrani': df_ilk6ay.loc[comp, m] * 100 if comp in df_ilk6ay.index and m in df_ilk6ay.columns else 0,
                    'aySekiceGiris': df_giris.loc[comp, m] if comp in df_giris.index and m in df_giris.columns else 0,
                    'aySekiceCikis': df_cikis.loc[comp, m] if comp in df_cikis.index and m in df_cikis.columns else 0,
                }
            else:
                comp_data[comp] = {key: 0 for key in ['employees', 'devamsizlik', 'turnoverKumulatif', 'turnoverGonulluKumulatif',
                                                      'turnoverAylik', 'turnoverGonulluAylik', 'netKokUcret',
                                                      'isverenMaliyet', 'fmSaat', 'fmTlMaliyet', 'izinGun', 'izinUcret',
                                                      'kisiBasiOrt', 'yillikOrtMaas', 'kadinOrani', 'ilk6ayOrani',
                                                      'aySekiceGiris', 'aySekiceCikis']}

        # Genel toplamlar (sayfa son satırları)
        sheet_calisan = toplam_calisan[idx] if idx < len(toplam_calisan) else 0
        sheet_rapor = genel_rapor[idx] * 100 if idx < len(genel_rapor) else 0
        sheet_izin_gun = toplam_izin_gun[idx] if idx < len(toplam_izin_gun) else 0
        sheet_izin_ucret = toplam_izin_ucret[idx] if idx < len(toplam_izin_ucret) else 0

        calc_calisan = sum(v['employees'] for v in comp_data.values())
        calc_rapor = sum(v['devamsizlik'] for v in comp_data.values()) / len(COMPANIES)
        calc_izin_gun = sum(v['izinGun'] for v in comp_data.values())
        calc_izin_ucret = sum(v['izinUcret'] for v in comp_data.values())

        # Turnover genel toplamlar (kümülatif ve aylık)
        genel_kumulatif_turnover = gt_monthly_totals.get(m, 0) * 100
        genel_kumulatif_gonullu = gon_monthly_totals.get(m, 0) * 100
        genel_aylik_turnover = aylik_gt_totals.get(m, 0) * 100
        genel_aylik_gonullu = aylik_gon_totals.get(m, 0) * 100

        data[m] = {
            'companies': comp_data,
            'genelRaporOran': sheet_rapor if sheet_rapor else calc_rapor,
            'toplamCalisan': sheet_calisan if sheet_calisan else calc_calisan,
            'toplamIzinGun': sheet_izin_gun if sheet_izin_gun else calc_izin_gun,
            'toplamIzinUcret': sheet_izin_ucret if sheet_izin_ucret else calc_izin_ucret,
            'kidemTazminati': kidem_totals.get(m, 0),
            'ihbarTazminati': ihbar_totals.get(m, 0),
            # Aylık genel kişi başı ortalama maaş (kisi.basi.ort'dan)
            'kisiBasiOrtGenel': kisi_basi_genel.get(m, 0),
            # Kümülatif genel kişi başı ortalama maaş (gnl.kisi.basi.ort'dan)
            'kisiBasiOrtGenelKumulatif': gnl_kisi_basi_vals.get(m, 0),
            'kadinOraniGenel': kadin_genel.get(m, 0) * 100,
            'ilk6ayOraniGenel': ilk6ay_ortalama.get(m, 0) * 100,
            'girisToplam': giris_toplam.get(m, 0),
            'cikisToplam': cikis_toplam.get(m, 0),
            'genelKumulatifTurnover': genel_kumulatif_turnover,
            'genelKumulatifGonullu': genel_kumulatif_gonullu,
            'genelAylikTurnover': genel_aylik_turnover,
            'genelAylikGonullu': genel_aylik_gonullu,
        }

    # Şirket bazlı toplam turnover (kümülatif son sütun)
    df_gt_total = pd.read_excel(uploaded_file, sheet_name='genel.turnover', header=0)
    df_gt_total = clean_columns(df_gt_total)
    total_col = None
    for col in df_gt_total.columns:
        if 'toplam' in str(col).lower():
            total_col = col
            break
    if total_col is None:
        total_col = df_gt_total.columns[-1]
    df_gt_total['Şirket'] = df_gt_total.iloc[:, 0].apply(normalize_company_name)
    df_gt_total = df_gt_total.set_index('Şirket')
    turnover_sirket_toplam = df_gt_total[~df_gt_total.index.isna()][total_col] * 100

    df_gon_total = pd.read_excel(uploaded_file, sheet_name='gonullu.turnover', header=0)
    df_gon_total = clean_columns(df_gon_total)
    total_col_gon = None
    for col in df_gon_total.columns:
        if 'toplam' in str(col).lower():
            total_col_gon = col
            break
    if total_col_gon is None:
        total_col_gon = df_gon_total.columns[-1]
    df_gon_total['Şirket'] = df_gon_total.iloc[:, 0].apply(normalize_company_name)
    df_gon_total = df_gon_total.set_index('Şirket')
    turnover_sirket_gonullu = df_gon_total[~df_gon_total.index.isna()][total_col_gon] * 100

    turnover_sirket_bazli = pd.DataFrame({
        'Toplam': [turnover_sirket_toplam.get(c, 0) for c in COMPANIES],
        'Gönüllü': [turnover_sirket_gonullu.get(c, 0) for c in COMPANIES]
    }, index=COMPANIES)

    return {
        'by_month': data,
        'fm_yapan': df_fm_yapan,
        'turnoverSirketBazli': turnover_sirket_bazli,
    }


def main():
    st.title("📊 İK Konsolide Dashboard")

    existing_path = get_existing_data_path()
    with st.sidebar:
        with st.expander("🔒 Veriyi Güncelle", expanded=not existing_path):
            pwd = st.text_input("Şifre", type="password")
            new_file = st.file_uploader("Yeni Excel dosyası (.xlsx / .xlsb)", type=["xlsx", "xlsb"], key="admin_uploader")
            if new_file is not None:
                if pwd == UPLOAD_PASSWORD:
                    for fname in os.listdir(DATA_DIR):
                        if fname.startswith("son_veri."):
                            os.remove(os.path.join(DATA_DIR, fname))
                    ext = os.path.splitext(new_file.name)[1].lower()
                    new_path = os.path.join(DATA_DIR, f"son_veri{ext}")
                    with open(new_path, "wb") as f:
                        f.write(new_file.getbuffer())
                    st.cache_data.clear()
                    st.success("✅ Veri güncellendi.")
                    st.rerun()
                else:
                    st.error("Şifre yanlış.")

    st.markdown("---")

    data_path = get_existing_data_path()
    if not data_path:
        st.info("Henüz veri yüklenmedi. Soldaki '🔒 Veriyi Güncelle' panelinden bir Excel dosyası yükleyin.")
        return

    try:
        cache_key = f"{data_path}:{os.path.getmtime(data_path)}"
        all_data = load_data(data_path, cache_key=cache_key)
    except Exception as e:
        st.error(f"Hata oluştu: {e}")
        st.exception(e)
        st.stop()

    data = all_data['by_month']
    fm_yapan_df = all_data['fm_yapan']
    turnover_sirket_bazli = all_data['turnoverSirketBazli']

    selected_month = st.selectbox("📅 Ay Seçin", MONTHS, index=len(MONTHS) - 1)
    month_idx = MONTHS.index(selected_month)
    prev_month = MONTHS[month_idx - 1] if month_idx > 0 else None

    month_data = data[selected_month]
    prev_data = data[prev_month] if prev_month else None

    # ----- KPI KARTLARI (1. satır) -----
    col1, col2, col3, col4, col5, col6 = st.columns(6)

    cur = month_data['toplamCalisan']
    prev = prev_data['toplamCalisan'] if prev_data else None
    diff = calc_diff(cur, prev)
    col1.metric("👥 Çalışan Sayısı", format_number(cur, 0), delta=format_delta_number(diff, 0))

    cur = month_data['genelRaporOran']
    prev = prev_data['genelRaporOran'] if prev_data else None
    diff = calc_diff(cur, prev)
    col2.metric("📊 Genel Raporlu Oran", format_percent(cur), delta=format_delta_percent(diff))

    cur = round(sum(d['isverenMaliyet'] for d in month_data['companies'].values()), 2)
    prev = round(sum(d['isverenMaliyet'] for d in prev_data['companies'].values()), 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col3.metric("💼 İşveren Maliyeti", format_tl(cur), delta=format_delta_tl(diff))

    cur = round(sum(d['netKokUcret'] for d in month_data['companies'].values()), 2)
    prev = round(sum(d['netKokUcret'] for d in prev_data['companies'].values()), 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col4.metric("💰 Net Kök Ücret", format_tl(cur), delta=format_delta_tl(diff))

    cur = round(sum(d['fmSaat'] for d in month_data['companies'].values()), 2)
    prev = round(sum(d['fmSaat'] for d in prev_data['companies'].values()), 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col5.metric("⏱️ FM Saat", format_number(cur, 1), delta=format_delta_number(diff, 1))

    cur = round(sum(d['fmTlMaliyet'] for d in month_data['companies'].values()), 2)
    prev = round(sum(d['fmTlMaliyet'] for d in prev_data['companies'].values()), 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col6.metric("💸 FM (Net TL)", format_tl(cur), delta=format_delta_tl(diff))

    # ----- KPI KARTLARI (2. satır) -----
    col7, col8, col9, col10, col11 = st.columns(5)

    cur = round(sum(d['izinGun'] for d in month_data['companies'].values()), 2)
    prev = round(sum(d['izinGun'] for d in prev_data['companies'].values()), 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col7.metric("📅 İzin Gün Bakiyesi", format_number(cur, 1), delta=format_delta_number(diff, 1))

    cur = round(sum(d['izinUcret'] for d in month_data['companies'].values()), 2)
    prev = round(sum(d['izinUcret'] for d in prev_data['companies'].values()), 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col8.metric("💎 İzin Ücreti (Net TL)", format_tl(cur), delta=format_delta_tl(diff))

    cur = round(month_data['kidemTazminati'], 2)
    prev = round(prev_data['kidemTazminati'], 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col9.metric("🏷️ Kıdem Tazminatı (Net TL)", format_tl(cur), delta=format_delta_tl(diff))

    cur = round(month_data['ihbarTazminati'], 2)
    prev = round(prev_data['ihbarTazminati'], 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col10.metric("📨 İhbar Tazminatı (Net TL)", format_tl(cur), delta=format_delta_tl(diff))

    # KPI 11: Kişi Başı Ortalama Maaş - gnl.kisi.basi.ort sayfasından (kümülatif ortalama)
    cur = round(month_data['kisiBasiOrtGenelKumulatif'], 2)
    prev = round(prev_data['kisiBasiOrtGenelKumulatif'], 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col11.metric("🧮 Kişi Başı Ort. Maaş (Net TL)", format_tl(cur), delta=format_delta_tl(diff))

    # ----- KPI KARTLARI (3. satır) -----
    col12, col13, col14, col15, col16, col17 = st.columns(6)

    col12.metric("🔄 Küm. Genel Turnover", format_percent(month_data['genelKumulatifTurnover']))
    col13.metric("🚪 Küm. Gönüllü Turnover", format_percent(month_data['genelKumulatifGonullu']))
    col14.metric("📈 Aylık Genel Turnover", format_percent(month_data['genelAylikTurnover']))
    col15.metric("📉 Aylık Gönüllü Turnover", format_percent(month_data['genelAylikGonullu']))

    cur = month_data['kadinOraniGenel']
    prev = prev_data['kadinOraniGenel'] if prev_data else None
    diff = calc_diff(cur, prev)
    col16.metric("👩 Kadın Oranı", format_percent(cur), delta=format_delta_percent(diff))

    cur = month_data['ilk6ayOraniGenel']
    prev = prev_data['ilk6ayOraniGenel'] if prev_data else None
    diff = calc_diff(cur, prev)
    col17.metric("⏳ İlk 6 Ay Ayrılma Oranı", format_percent(cur), delta=format_delta_percent(diff))

    st.markdown("---")

    # ----- GRAFİKLER -----
    # 1. Satır: Çalışan Sayısı + Kümülatif Turnover
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("👥 Şirket Bazında Çalışan Sayısı")
        df_calisan_plot = pd.DataFrame({
            'Şirket': COMPANIES,
            'Çalışan Sayısı': [month_data['companies'][c]['employees'] for c in COMPANIES]
        })
        fig_calisan = px.bar(df_calisan_plot, x='Şirket', y='Çalışan Sayısı', color='Çalışan Sayısı',
                             color_continuous_scale='Greens',
                             text=df_calisan_plot['Çalışan Sayısı'].apply(lambda x: format_number(x, 0)))
        fig_calisan.update_traces(textposition='outside')
        fig_calisan.update_layout(showlegend=False, height=350, margin=dict(l=10, r=10, t=30, b=10),
                                  xaxis_title="")
        st.plotly_chart(fig_calisan, use_container_width=True)

    with col2:
        st.subheader("📈 Şirket Bazında Kümülatif Turnover")
        df_kum = pd.DataFrame({
            'Şirket': COMPANIES,
            'Küm. Genel': [month_data['companies'][c]['turnoverKumulatif'] for c in COMPANIES],
            'Küm. Gönüllü': [month_data['companies'][c]['turnoverGonulluKumulatif'] for c in COMPANIES]
        })
        fig_kum = go.Figure()
        fig_kum.add_trace(go.Bar(x=df_kum['Şirket'], y=df_kum['Küm. Genel'], name='Küm. Genel',
                                 marker_color='#f59e0b', text=df_kum['Küm. Genel'].apply(format_percent),
                                 textposition='outside'))
        fig_kum.add_trace(go.Bar(x=df_kum['Şirket'], y=df_kum['Küm. Gönüllü'], name='Küm. Gönüllü',
                                 marker_color='#ec4899', text=df_kum['Küm. Gönüllü'].apply(format_percent),
                                 textposition='outside'))
        fig_kum.update_layout(barmode='group', height=350, margin=dict(l=10, r=10, t=30, b=10),
                              legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
                              xaxis_title="")
        st.plotly_chart(fig_kum, use_container_width=True)

    # 2. Satır: Aylık Turnover + Raporlu Oran
    col3, col4 = st.columns(2)

    with col3:
        st.subheader("📊 Şirket Bazında Aylık Turnover")
        df_aylik = pd.DataFrame({
            'Şirket': COMPANIES,
            'Aylık Genel': [month_data['companies'][c]['turnoverAylik'] for c in COMPANIES],
            'Aylık Gönüllü': [month_data['companies'][c]['turnoverGonulluAylik'] for c in COMPANIES]
        })
        fig_aylik = go.Figure()
        fig_aylik.add_trace(go.Bar(x=df_aylik['Şirket'], y=df_aylik['Aylık Genel'], name='Aylık Genel',
                                   marker_color='#3b82f6', text=df_aylik['Aylık Genel'].apply(format_percent),
                                   textposition='outside'))
        fig_aylik.add_trace(go.Bar(x=df_aylik['Şirket'], y=df_aylik['Aylık Gönüllü'], name='Aylık Gönüllü',
                                   marker_color='#8b5cf6', text=df_aylik['Aylık Gönüllü'].apply(format_percent),
                                   textposition='outside'))
        fig_aylik.update_layout(barmode='group', height=350, margin=dict(l=10, r=10, t=30, b=10),
                                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
                                xaxis_title="")
        st.plotly_chart(fig_aylik, use_container_width=True)

    with col4:
        st.subheader("📊 Raporlu Oran")
        df_plot = pd.DataFrame({
            'Şirket': COMPANIES,
            'Raporlu Oran %': [month_data['companies'][c]['devamsizlik'] for c in COMPANIES]
        })
        fig_rapor = px.bar(df_plot, x='Şirket', y='Raporlu Oran %', color='Raporlu Oran %',
                           color_continuous_scale='Blues',
                           text=df_plot['Raporlu Oran %'].apply(format_percent))
        fig_rapor.update_traces(textposition='outside')
        fig_rapor.update_layout(showlegend=False, height=350, margin=dict(l=10, r=10, t=30, b=10),
                                xaxis_title="")
        st.plotly_chart(fig_rapor, use_container_width=True)

    # 3. Satır: FM Saat + İzin Gün
    col5, col6 = st.columns(2)

    with col5:
        st.subheader("⏱️ Şirket Bazında FM Saat")
        df_fm = pd.DataFrame({
            'Şirket': COMPANIES,
            'FM Saat': [month_data['companies'][c]['fmSaat'] for c in COMPANIES]
        })
        fig_fm = px.bar(df_fm, x='Şirket', y='FM Saat', color='FM Saat',
                        color_continuous_scale='Oranges',
                        text=df_fm['FM Saat'].apply(lambda x: format_number(x, 1)))
        fig_fm.update_traces(textposition='outside')
        fig_fm.update_layout(showlegend=False, height=350, margin=dict(l=10, r=10, t=30, b=10),
                             xaxis_title="")
        st.plotly_chart(fig_fm, use_container_width=True)

    with col6:
        st.subheader("📅 İzin Gün Bakiyesi")
        df_izin_gun = pd.DataFrame({
            'Şirket': COMPANIES,
            'İzin Günü': [month_data['companies'][c]['izinGun'] for c in COMPANIES]
        })
        fig_izin = px.bar(df_izin_gun, x='Şirket', y='İzin Günü', color='İzin Günü',
                          color_continuous_scale='Teal',
                          text=df_izin_gun['İzin Günü'].apply(lambda x: format_number(x, 1)))
        fig_izin.update_traces(textposition='outside')
        fig_izin.update_layout(showlegend=False, height=350, margin=dict(l=10, r=10, t=30, b=10),
                               xaxis_title="")
        st.plotly_chart(fig_izin, use_container_width=True)

    # 4. Satır: İzin Ücreti + Kadın Oranı
    col7, col8 = st.columns(2)

    with col7:
        st.subheader("💰 İzin Ücretleri (Net TL)")
        df_izin_ucret = pd.DataFrame({
            'Şirket': COMPANIES,
            'İzin Ücreti (Net TL)': [month_data['companies'][c]['izinUcret'] for c in COMPANIES]
        })
        fig_ucret = px.bar(df_izin_ucret, x='Şirket', y='İzin Ücreti (Net TL)', color='İzin Ücreti (Net TL)',
                           color_continuous_scale='Purples',
                           text=df_izin_ucret['İzin Ücreti (Net TL)'].apply(format_tl))
        fig_ucret.update_traces(textposition='outside')
        fig_ucret.update_layout(showlegend=False, height=350, margin=dict(l=10, r=10, t=30, b=10),
                                xaxis_title="")
        st.plotly_chart(fig_ucret, use_container_width=True)

    with col8:
        st.subheader("👩 Kadın Oranı")
        df_kadin_plot = pd.DataFrame({
            'Şirket': COMPANIES,
            'Kadın Oranı %': [month_data['companies'][c]['kadinOrani'] for c in COMPANIES]
        })
        fig_kadin = px.bar(df_kadin_plot, x='Şirket', y='Kadın Oranı %', color='Kadın Oranı %',
                           color_continuous_scale='Magenta',
                           text=df_kadin_plot['Kadın Oranı %'].apply(format_percent))
        fig_kadin.update_traces(textposition='outside')
        fig_kadin.update_layout(showlegend=False, height=350, margin=dict(l=10, r=10, t=30, b=10),
                                xaxis_title="")
        st.plotly_chart(fig_kadin, use_container_width=True)

    # 5. Satır: Giriş/Çıkış
    col9, col10 = st.columns(2)

    with col9:
        st.subheader("🔁 Ay İçi Giriş & Çıkış")
        df_gc = pd.DataFrame({
            'Şirket': COMPANIES,
            'Giriş': [month_data['companies'][c]['aySekiceGiris'] for c in COMPANIES],
            'Çıkış': [month_data['companies'][c]['aySekiceCikis'] for c in COMPANIES]
        })
        fig_gc = go.Figure()
        fig_gc.add_trace(go.Bar(x=df_gc['Şirket'], y=df_gc['Giriş'], name='Giriş',
                                marker_color='#22c55e',
                                text=df_gc['Giriş'].apply(lambda x: format_number(x, 0)),
                                textposition='outside'))
        fig_gc.add_trace(go.Bar(x=df_gc['Şirket'], y=df_gc['Çıkış'], name='Çıkış',
                                marker_color='#ef4444',
                                text=df_gc['Çıkış'].apply(lambda x: format_number(x, 0)),
                                textposition='outside'))
        fig_gc.update_layout(barmode='group', height=350, margin=dict(l=10, r=10, t=30, b=10),
                             legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
                             xaxis_title="")
        st.plotly_chart(fig_gc, use_container_width=True)

    with col10:
        st.empty()

    st.markdown("---")

    # ----- EN ÇOK MESAİ YAPAN 10 KİŞİ -----
    st.subheader(f"🏆 {selected_month} Ayında En Çok Mesai Yapan 10 Kişi")
    if selected_month in fm_yapan_df.columns:
        cols_needed = [c for c in ['Adı Soyadı', 'Şirket', 'Lokasyon'] if c in fm_yapan_df.columns]
        top10 = fm_yapan_df[cols_needed + [selected_month]].copy()
        top10 = top10.sort_values(by=selected_month, ascending=False).head(10)
        top10 = top10.rename(columns={selected_month: 'FM Saat'})
        top10['FM Saat'] = top10['FM Saat'].apply(lambda x: format_number(x, 1))
        top10.insert(0, 'Sıra', range(1, len(top10) + 1))
        st.dataframe(top10, use_container_width=True, hide_index=True)
    else:
        st.info("Seçilen ay için mesai verisi bulunamadı.")

    st.markdown("---")

    # ----- DETAY TABLOSU (DİP TOPLAMLI) -----
    st.subheader(f"📋 {selected_month} - Tüm Şirketlerin Detaylı Verileri")

    rows = []
    for comp in COMPANIES:
        d = month_data['companies'][comp]
        rows.append({
            'Şirket': comp,
            'Çalışan': format_number(d['employees'], 0),
            'Raporlu Oran %': format_percent(d['devamsizlik']),
            'Küm. Turnover %': format_percent(d['turnoverKumulatif']),
            'Küm. Gön. %': format_percent(d['turnoverGonulluKumulatif']),
            'Aylık Turnover %': format_percent(d['turnoverAylik']),
            'Aylık Gön. %': format_percent(d['turnoverGonulluAylik']),
            'Net Kök Ücret': format_tl(d['netKokUcret']),
            'İşveren Maliyeti': format_tl(d['isverenMaliyet']),
            'FM Saat': format_number(d['fmSaat'], 1),
            'FM (Net TL)': format_tl(d['fmTlMaliyet']),
            'İzin Gün Bakiyesi': format_number(d['izinGun'], 1),
            'İzin Ücreti (Net TL)': format_tl(d['izinUcret']),
            'Yıllık Kişi Başı Ort. Maaş (Net TL)': format_tl(d['yillikOrtMaas']),
            'Ay İçi İşe Giren': format_number(d['aySekiceGiris'], 0),
            'Ay İçi İşten Ayrılan': format_number(d['aySekiceCikis'], 0),
        })

    total_employees = sum(d['employees'] for d in month_data['companies'].values())
    total_rapor = sum(d['devamsizlik'] for d in month_data['companies'].values()) / len(COMPANIES)
    total_net = sum(d['netKokUcret'] for d in month_data['companies'].values())
    total_isveren = sum(d['isverenMaliyet'] for d in month_data['companies'].values())
    total_fm_saat = sum(d['fmSaat'] for d in month_data['companies'].values())
    total_fm_tl = sum(d['fmTlMaliyet'] for d in month_data['companies'].values())
    total_izin_gun = sum(d['izinGun'] for d in month_data['companies'].values())
    total_izin_ucret = sum(d['izinUcret'] for d in month_data['companies'].values())

    total_row = {
        'Şirket': '⭐ TOPLAM',
        'Çalışan': format_number(total_employees, 0),
        'Raporlu Oran %': format_percent(total_rapor),
        'Küm. Turnover %': format_percent(month_data['genelKumulatifTurnover']),
        'Küm. Gön. %': format_percent(month_data['genelKumulatifGonullu']),
        'Aylık Turnover %': format_percent(month_data['genelAylikTurnover']),
        'Aylık Gön. %': format_percent(month_data['genelAylikGonullu']),
        'Net Kök Ücret': format_tl(total_net),
        'İşveren Maliyeti': format_tl(total_isveren),
        'FM Saat': format_number(total_fm_saat, 1),
        'FM (Net TL)': format_tl(total_fm_tl),
        'İzin Gün Bakiyesi': format_number(total_izin_gun, 1),
        'İzin Ücreti (Net TL)': format_tl(total_izin_ucret),
        'Yıllık Kişi Başı Ort. Maaş (Net TL)': format_tl(month_data['kisiBasiOrtGenelKumulatif']),  # gnl.kisi.basi.ort değeri
        'Ay İçi İşe Giren': format_number(month_data['girisToplam'], 0),
        'Ay İçi İşten Ayrılan': format_number(month_data['cikisToplam'], 0),
    }
    rows.append(total_row)

    detail_df = pd.DataFrame(rows)
    st.dataframe(detail_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
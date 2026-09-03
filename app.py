import os
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(layout="wide", page_title="İK Dashboard")

# Yüklenen Excel dosyasının sunucuda saklanacağı kalıcı yol.
# Streamlit Community Cloud'da bu dosya, uygulama yeniden başlatılana/deploy
# edilene kadar diskte kalır; tamamen kalıcı bir depolama değildir. Uzun vadeli
# kalıcılık için GitHub'daki repo'ya commit etmek ya da bir bulut depolama
# (S3, Google Sheets vb.) kullanmak daha güvenlidir.
# Yüklenen Excel dosyasının sunucuda saklanacağı klasör. Dosya, orijinal
# uzantısıyla (xlsx ya da xlsb) saklanır ki pandas doğru okuyucu motoru
# (openpyxl / pyxlsb) otomatik seçebilsin.
# Streamlit Community Cloud'da bu klasör, uygulama yeniden başlatılana/deploy
# edilene kadar diskte kalır; tamamen kalıcı bir depolama değildir. Uzun vadeli
# kalıcılık için GitHub'daki repo'ya commit etmek ya da bir bulut depolama
# (S3, Google Sheets vb.) kullanmak daha güvenlidir.
DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)

def get_existing_data_path():
    """data/ klasöründe daha önce yüklenmiş dosyayı (uzantısı ne olursa olsun) bulur."""
    for fname in os.listdir(DATA_DIR):
        if fname.startswith("son_veri."):
            return os.path.join(DATA_DIR, fname)
    return None

# Yükleme panelini korumak için basit bir şifre. Gerçek kullanımda bunu
# st.secrets["upload_password"] üzerinden okumanız (Streamlit Cloud > Settings >
# Secrets) sabit kod yerine daha güvenli olur.
UPLOAD_PASSWORD = st.secrets.get("upload_password", "ik2026") if hasattr(st, "secrets") else "ik2026"

# KPI kartlarındaki başlık/sayı/delta fontlarını küçültmek için özel CSS
st.markdown("""
<style>
[data-testid="stMetricValue"] {
    font-size: 1.0rem;
}
[data-testid="stMetricLabel"] {
    font-size: 0.72rem;
}
[data-testid="stMetricDelta"] {
    font-size: 0.65rem;
}
[data-testid="stMetric"] {
    padding: 0.35rem 0.25rem;
}
</style>
""", unsafe_allow_html=True)

COMPANIES = [
    'Aralık Sigorta',
    'Ekim Turizm',
    'Eylül Girişim',
    'Haziran Servis',
    'Intercity Yatırım Holding',
    'Mart Denizcilik'
]
MONTHS = ['Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz', 'Ağustos']

# ----- FORMAT YARDIMCILARI -----
def format_tl(value):
    """TL formatı (boşluksuz): 17.320.728,75TL"""
    if value is None:
        return "0TL"
    try:
        formatted = f"{value:,.2f}"
        # Virgül ve noktayı yer değiştir (Türkçe format)
        formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")
        return f"{formatted}TL"
    except Exception:
        return f"{value:.2f}TL".replace(".", ",")

def format_tl_no_decimal(value):
    """TL formatı (ondalıksız, boşluksuz): 1.732.029TL"""
    if value is None:
        return "0TL"
    try:
        formatted = f"{value:,.0f}"
        formatted = formatted.replace(",", ".")
        return f"{formatted}TL"
    except Exception:
        return f"{value:.0f}TL"

def format_percent(value):
    """Yüzde formatı: %1,82"""
    if value is None:
        return "%0"
    try:
        formatted = f"{value:.2f}"
        formatted = formatted.replace(".", ",")
        return f"%{formatted}"
    except Exception:
        return f"%{value:.2f}".replace(".", ",")

def format_number(value, decimals=1):
    """Genel Türkçe sayı formatı: binlik nokta, ondalık virgül. Örn: 1.234,5"""
    if value is None:
        return "0"
    try:
        formatted = f"{value:,.{decimals}f}"
        formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")
        return formatted
    except Exception:
        return str(value)

# ----- DELTA FORMAT YARDIMCILARI (float hassasiyet gürültüsünü de temizler) -----
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

def read_sheet_flexible(uploaded_file, candidates, **kwargs):
    """Birden fazla olası sayfa adından ilk bulunanı okur (sayfa adı dosyalar arasında değişmiş olabilir)."""
    last_err = None
    for name in candidates:
        try:
            return pd.read_excel(uploaded_file, sheet_name=name, **kwargs)
        except Exception as e:
            last_err = e
    raise last_err

def safe_read_son(uploaded_file, sheet, skip, n_months=None):
    if n_months is None:
        n_months = len(MONTHS)
    try:
        df = pd.read_excel(uploaded_file, sheet_name=sheet, skiprows=skip, header=None, nrows=1)
        if df.shape[1] >= n_months + 1:
            return df.iloc[0, 1:n_months + 1].values
        else:
            return [0] * n_months
    except Exception:
        return [0] * n_months

def read_company_month_sheet(uploaded_file, sheet_name, total_label, agg='sum', months=None):
    """Şirket satırları + en altta/ayrı bir 'TOPLAM' ya da 'Genel ...' satırı olan
    sayfaları okur. Şirket bazlı DataFrame (ay sütunlu) ve ay->toplam sözlüğü döner.
    Toplam satırı sayfada yoksa ya da boşsa, şirket verilerinden hesaplanır."""
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

# ----- VERİ OKUMA -----
@st.cache_data(show_spinner=False)
def load_data(_uploaded_file, cache_key=None):
    uploaded_file = _uploaded_file
    # 1. Genel Turnover
    df_gt = pd.read_excel(uploaded_file, sheet_name='genel.turnover', header=0)
    df_gt = clean_columns(df_gt)
    month_cols = get_month_cols(df_gt)
    if not month_cols:
        raise ValueError("genel.turnover sayfasında ay sütunları bulunamadı.")
    df_gt['Şirket'] = df_gt.iloc[:, 0].apply(normalize_company_name)
    df_gt = df_gt.set_index('Şirket')[month_cols]

    # 2. Gönüllü Turnover
    df_gon = pd.read_excel(uploaded_file, sheet_name='gonullu.turnover', header=0)
    df_gon = clean_columns(df_gon)
    month_cols = get_month_cols(df_gon)
    df_gon['Şirket'] = df_gon.iloc[:, 0].apply(normalize_company_name)
    df_gon = df_gon.set_index('Şirket')[month_cols]

    # 3. Rapor Oranı
    df_rapor = pd.read_excel(uploaded_file, sheet_name='rapor_oran', header=0)
    df_rapor = clean_columns(df_rapor)
    month_cols = get_month_cols(df_rapor)
    df_rapor['Şirket'] = df_rapor.iloc[:, 0].apply(normalize_company_name)
    df_rapor = df_rapor.set_index('Şirket')[month_cols]

    # 4. Çalışan Sayısı
    df_calisan = pd.read_excel(uploaded_file, sheet_name='calisan.sayisi', header=0)
    df_calisan = clean_columns(df_calisan)
    month_cols = get_month_cols(df_calisan)
    df_calisan['Şirket'] = df_calisan.iloc[:, 0].apply(normalize_company_name)
    df_calisan = df_calisan.set_index('Şirket')[month_cols]

    # 5. Net Kök Ücret
    df_net = read_sheet_flexible(uploaded_file, ['kok.ucret', 'maliyet'], header=0)
    df_net = clean_columns(df_net)
    month_cols = get_month_cols(df_net)
    df_net['Şirket'] = df_net.iloc[:, 0].apply(normalize_company_name)
    df_net = df_net.set_index('Şirket')[month_cols]

    # 6. İşveren Maliyeti
    df_isv = pd.read_excel(uploaded_file, sheet_name='isveren.maliyet', header=0)
    df_isv = clean_columns(df_isv)
    month_cols = get_month_cols(df_isv)
    df_isv['Şirket'] = df_isv.iloc[:, 0].apply(normalize_company_name)
    df_isv = df_isv.set_index('Şirket')[month_cols]

    # 7. FM Saat
    df_fm_saat = pd.read_excel(uploaded_file, sheet_name='fm.saat', header=0)
    df_fm_saat = clean_columns(df_fm_saat)
    month_cols = get_month_cols(df_fm_saat)
    df_fm_saat['Şirket'] = df_fm_saat.iloc[:, 0].apply(normalize_company_name)
    df_fm_saat = df_fm_saat.set_index('Şirket')[month_cols]

    # 8. FM TL Maliyet
    df_fm_tl = pd.read_excel(uploaded_file, sheet_name='fm.maliyet', header=0)
    df_fm_tl = clean_columns(df_fm_tl)
    month_cols = get_month_cols(df_fm_tl)
    df_fm_tl['Şirket'] = df_fm_tl.iloc[:, 0].apply(normalize_company_name)
    df_fm_tl = df_fm_tl.set_index('Şirket')[month_cols]

    # 9. İzin Gün
    df_izin_gun = pd.read_excel(uploaded_file, sheet_name='izin_gun', header=0)
    df_izin_gun = clean_columns(df_izin_gun)
    month_cols = get_month_cols(df_izin_gun)
    df_izin_gun['Şirket'] = df_izin_gun.iloc[:, 0].apply(normalize_company_name)
    df_izin_gun = df_izin_gun.set_index('Şirket')[month_cols]

    # 10. İzin Ücreti
    df_izin_ucret = pd.read_excel(uploaded_file, sheet_name='izin_ucret', header=0)
    df_izin_ucret = clean_columns(df_izin_ucret)
    month_cols = get_month_cols(df_izin_ucret)
    df_izin_ucret['Şirket'] = df_izin_ucret.iloc[:, 0].apply(normalize_company_name)
    df_izin_ucret = df_izin_ucret.set_index('Şirket')[month_cols]

    # ----- TOPLAM SATIRLARI (skiprows 7 ile, sayfada varsa) -----
    genel_rapor = safe_read_son(uploaded_file, 'rapor_oran', 7)
    toplam_calisan = safe_read_son(uploaded_file, 'calisan.sayisi', 7)
    toplam_izin_gun = safe_read_son(uploaded_file, 'izin_gun', 7)
    toplam_izin_ucret = safe_read_son(uploaded_file, 'izin_ucret', 7)

    # 11. Aylık FM Yapan (kişi bazlı - en çok mesai yapanlar listesi için)
    df_fm_yapan = pd.read_excel(uploaded_file, sheet_name='aylik.fm.yapan', header=0)
    df_fm_yapan = clean_columns(df_fm_yapan)
    if 'Şirket' in df_fm_yapan.columns:
        df_fm_yapan['Şirket'] = df_fm_yapan['Şirket'].apply(normalize_company_name)
    for m in get_month_cols(df_fm_yapan):
        df_fm_yapan[m] = pd.to_numeric(df_fm_yapan[m], errors='coerce').fillna(0)

    # 12. Kıdem Tazminatı (şirket x ay + TOPLAM satırı)
    df_kidem, kidem_totals = read_company_month_sheet(uploaded_file, 'kidem.tazminati', 'TOPLAM', agg='sum')

    # 13. İhbar Tazminatı (şirket x ay + TOPLAM satırı)
    df_ihbar, ihbar_totals = read_company_month_sheet(uploaded_file, 'ihbar.tazminati', 'TOPLAM', agg='sum')

    # 14. Kişi Başı Ortalama Maaş (şirket x ay + "Genel Kişi Başı Ortalama Maaş" satırı)
    df_kisi_basi, kisi_basi_genel = read_company_month_sheet(
        uploaded_file, 'kisi.basi.ort', 'Genel Kişi Başı Ortalama Maaş', agg='mean'
    )
    # Sayfaya sonradan eklenen "Yıllık Ortalama" sütunu (şirket bazlı) - varsa al
    yillik_ort_col = None
    for col in df_kisi_basi.columns:
        if str(col).strip().lower() in ('yıllık ortalama', 'yillik ortalama'):
            yillik_ort_col = col
            break

    # 15. Kadın Oranı (şirket x ay + "Genel Kadın Oranı" satırı)
    df_kadin, kadin_genel = read_company_month_sheet(uploaded_file, 'kadin.erkek', 'Genel Kadın Oranı', agg='mean')

    # 16. İlk 6 Ay İşten Ayrılma Oranı (sayfada toplam/genel satırı yok, şirket ortalaması alınır)
    df_ilk6ay, ilk6ay_ortalama = read_company_month_sheet(uploaded_file, 'ilk.6ay', '__YOK__', agg='mean')

    # 17. Ay İçinde İşe Giren (sayfada toplam satırı yok, şirketler toplanır)
    df_giris, giris_toplam = read_company_month_sheet(uploaded_file, 'aylik.giris', '__YOK__', agg='sum')

    # 18. Ay İçinde İşten Ayrılan (sayfada toplam satırı yok, şirketler toplanır)
    df_cikis, cikis_toplam = read_company_month_sheet(uploaded_file, 'aylik.cikis', '__YOK__', agg='sum')

    # ----- VERİYİ OLUŞTUR -----
    data = {}
    for idx, m in enumerate(MONTHS):
        comp_data = {}
        for comp in COMPANIES:
            if comp in df_calisan.index:
                comp_data[comp] = {
                    'employees': df_calisan.loc[comp, m],
                    'devamsizlik': df_rapor.loc[comp, m] * 100,
                    'turnoverToplam': df_gt.loc[comp, m] * 100,
                    'turnoverGonullu': df_gon.loc[comp, m] * 100,
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
                comp_data[comp] = {key: 0 for key in ['employees', 'devamsizlik', 'turnoverToplam', 'turnoverGonullu',
                                                        'netKokUcret', 'isverenMaliyet', 'fmSaat', 'fmTlMaliyet',
                                                        'izinGun', 'izinUcret', 'kisiBasiOrt', 'yillikOrtMaas',
                                                        'kadinOrani', 'ilk6ayOrani', 'aySekiceGiris', 'aySekiceCikis']}

        # Sayfadaki hazır "toplam" satırı
        sheet_calisan = toplam_calisan[idx] if idx < len(toplam_calisan) else 0
        sheet_rapor = genel_rapor[idx] * 100 if idx < len(genel_rapor) else 0
        sheet_izin_gun = toplam_izin_gun[idx] if idx < len(toplam_izin_gun) else 0
        sheet_izin_ucret = toplam_izin_ucret[idx] if idx < len(toplam_izin_ucret) else 0

        # Şirket bazlı verilerden hesaplanan yedek değerler
        # (sayfadaki toplam satırı 0/okunamıyor gelirse bunlar devreye girer,
        #  örn. "çalışan sayısı görünmüyor" sorunu buradan kaynaklanıyor olabilir)
        calc_calisan = sum(v['employees'] for v in comp_data.values())
        calc_rapor = sum(v['devamsizlik'] for v in comp_data.values()) / len(COMPANIES)
        calc_izin_gun = sum(v['izinGun'] for v in comp_data.values())
        calc_izin_ucret = sum(v['izinUcret'] for v in comp_data.values())

        data[m] = {
            'companies': comp_data,
            'genelRaporOran': sheet_rapor if sheet_rapor else calc_rapor,
            'toplamCalisan': sheet_calisan if sheet_calisan else calc_calisan,
            'toplamIzinGun': sheet_izin_gun if sheet_izin_gun else calc_izin_gun,
            'toplamIzinUcret': sheet_izin_ucret if sheet_izin_ucret else calc_izin_ucret,
            'kidemTazminati': kidem_totals.get(m, 0),
            'ihbarTazminati': ihbar_totals.get(m, 0),
            'kisiBasiOrtGenel': kisi_basi_genel.get(m, 0),
            'kadinOraniGenel': kadin_genel.get(m, 0) * 100,
            'ilk6ayOraniGenel': ilk6ay_ortalama.get(m, 0) * 100,
            'girisToplam': giris_toplam.get(m, 0),
            'cikisToplam': cikis_toplam.get(m, 0),
        }

    return {
        'by_month': data,
        'fm_yapan': df_fm_yapan,
    }


# ----- ANA UYGULAMA -----
def main():
    st.title("📊 İK Konsolide Dashboard")

    # ----- GİZLİ VERİ GÜNCELLEME PANELİ -----
    # Normal ziyaretçiler bunu görmeden direkt dashboard'u görür; veriyi
    # güncellemek isteyen kişi (örn. siz) paneli açıp şifreyle yeni dosya yükler.
    existing_path = get_existing_data_path()
    with st.sidebar:
        with st.expander("🔒 Veriyi Güncelle", expanded=not existing_path):
            pwd = st.text_input("Şifre", type="password")
            new_file = st.file_uploader("Yeni Excel dosyası (.xlsx / .xlsb)", type=["xlsx", "xlsb"], key="admin_uploader")
            if new_file is not None:
                if pwd == UPLOAD_PASSWORD:
                    # Önceki dosyayı (farklı uzantıda kalmış olabilir) temizle
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

    selected_month = st.selectbox("📅 Ay Seçin", MONTHS, index=len(MONTHS) - 1)
    month_idx = MONTHS.index(selected_month)
    prev_month = MONTHS[month_idx - 1] if month_idx > 0 else None

    month_data = data[selected_month]
    prev_data = data[prev_month] if prev_month else None

    # ----- KPI KARTLARI (1. satır) -----
    col1, col2, col3, col4, col5, col6 = st.columns(6)

    # 1. Çalışan
    cur = month_data['toplamCalisan']
    prev = prev_data['toplamCalisan'] if prev_data else None
    diff = calc_diff(cur, prev)
    col1.metric("👥 Çalışan Sayısı", format_number(cur, 0), delta=format_delta_number(diff, 0))

    # 2. Genel Raporlu Oran
    cur = month_data['genelRaporOran']
    prev = prev_data['genelRaporOran'] if prev_data else None
    diff = calc_diff(cur, prev)
    col2.metric("📊 Genel Raporlu Oran", format_percent(cur), delta=format_delta_percent(diff))

    # 3. İşveren Maliyeti
    cur = round(sum(d['isverenMaliyet'] for d in month_data['companies'].values()), 2)
    prev = round(sum(d['isverenMaliyet'] for d in prev_data['companies'].values()), 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col3.metric("💼 İşveren Maliyeti", format_tl(cur), delta=format_delta_tl(diff))

    # 4. Net Kök Ücret
    cur = round(sum(d['netKokUcret'] for d in month_data['companies'].values()), 2)
    prev = round(sum(d['netKokUcret'] for d in prev_data['companies'].values()), 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col4.metric("💰 Net Kök Ücret", format_tl(cur), delta=format_delta_tl(diff))

    # 5. FM Saat
    cur = round(sum(d['fmSaat'] for d in month_data['companies'].values()), 2)
    prev = round(sum(d['fmSaat'] for d in prev_data['companies'].values()), 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col5.metric("⏱️ FM Saat", format_number(cur, 1), delta=format_delta_number(diff, 1))

    # 6. FM (Net TL)
    cur = round(sum(d['fmTlMaliyet'] for d in month_data['companies'].values()), 2)
    prev = round(sum(d['fmTlMaliyet'] for d in prev_data['companies'].values()), 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col6.metric("💸 FM (Net TL)", format_tl(cur), delta=format_delta_tl(diff))

    # ----- KPI KARTLARI (2. satır) -----
    col7, col8, col9, col10, col11 = st.columns(5)

    # 7. İzin Gün Bakiyesi
    cur = round(sum(d['izinGun'] for d in month_data['companies'].values()), 2)
    prev = round(sum(d['izinGun'] for d in prev_data['companies'].values()), 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col7.metric("📅 İzin Gün Bakiyesi", format_number(cur, 1), delta=format_delta_number(diff, 1))

    # 8. İzin Ücreti (Net TL)
    cur = round(sum(d['izinUcret'] for d in month_data['companies'].values()), 2)
    prev = round(sum(d['izinUcret'] for d in prev_data['companies'].values()), 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col8.metric("💎 İzin Ücreti (Net TL)", format_tl(cur), delta=format_delta_tl(diff))

    # 9. Kıdem Tazminatı
    cur = round(month_data['kidemTazminati'], 2)
    prev = round(prev_data['kidemTazminati'], 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col9.metric("🏷️ Kıdem Tazminatı", format_tl(cur), delta=format_delta_tl(diff))

    # 10. İhbar Tazminatı
    cur = round(month_data['ihbarTazminati'], 2)
    prev = round(prev_data['ihbarTazminati'], 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col10.metric("📨 İhbar Tazminatı", format_tl(cur), delta=format_delta_tl(diff))

    # 11. Kişi Başı Ortalama Maaş (Net TL)
    cur = round(month_data['kisiBasiOrtGenel'], 2)
    prev = round(prev_data['kisiBasiOrtGenel'], 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col11.metric("🧮 Kişi Başı Ort. Maaş (Net TL)", format_tl(cur), delta=format_delta_tl(diff))

    # ----- KPI KARTLARI (3. satır) -----
    col12, col13, col14, col15 = st.columns(4)

    # 12. Genel Kadın Oranı
    cur = month_data['kadinOraniGenel']
    prev = prev_data['kadinOraniGenel'] if prev_data else None
    diff = calc_diff(cur, prev)
    col12.metric("👩 Kadın Oranı", format_percent(cur), delta=format_delta_percent(diff))

    # 13. İlk 6 Ay İşten Ayrılma Oranı
    cur = month_data['ilk6ayOraniGenel']
    prev = prev_data['ilk6ayOraniGenel'] if prev_data else None
    diff = calc_diff(cur, prev)
    col13.metric("⏳ İlk 6 Ay Ayrılma Oranı", format_percent(cur), delta=format_delta_percent(diff))

    # 14. Ay İçinde İşe Giren
    cur = round(month_data['girisToplam'], 2)
    prev = round(prev_data['girisToplam'], 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col14.metric("⬆️ Ay İçi İşe Giren", format_number(cur, 0), delta=format_delta_number(diff, 0))

    # 15. Ay İçinde İşten Ayrılan
    cur = round(month_data['cikisToplam'], 2)
    prev = round(prev_data['cikisToplam'], 2) if prev_data else None
    diff = calc_diff(cur, prev)
    col15.metric("⬇️ Ay İçi İşten Ayrılan", format_number(cur, 0), delta=format_delta_number(diff, 0))

    st.markdown("---")

    # ----- GRAFİKLER -----
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📊 Raporlu Oran")
        df_plot = pd.DataFrame({
            'Şirket': COMPANIES,
            'Raporlu Oran %': [month_data['companies'][c]['devamsizlik'] for c in COMPANIES]
        })
        fig = px.bar(df_plot, x='Şirket', y='Raporlu Oran %', color='Raporlu Oran %',
                     color_continuous_scale='Blues',
                     text=df_plot['Raporlu Oran %'].apply(format_percent))
        fig.update_traces(textposition='outside')
        fig.update_layout(showlegend=False, height=350, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.subheader("🔄 Turnover (Toplam & Gönüllü)")
        df2 = pd.DataFrame({
            'Şirket': COMPANIES,
            'Toplam': [month_data['companies'][c]['turnoverToplam'] for c in COMPANIES],
            'Gönüllü': [month_data['companies'][c]['turnoverGonullu'] for c in COMPANIES]
        })
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(x=df2['Şirket'], y=df2['Toplam'], name='Toplam Turnover',
                               marker_color='#f59e0b',
                               text=df2['Toplam'].apply(format_percent),
                               textposition='outside'))
        fig2.add_trace(go.Bar(x=df2['Şirket'], y=df2['Gönüllü'], name='Gönüllü Turnover',
                               marker_color='#ec4899',
                               text=df2['Gönüllü'].apply(format_percent),
                               textposition='outside'))
        fig2.update_layout(barmode='group', height=350, margin=dict(l=10, r=10, t=30, b=10),
                            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1))
        st.plotly_chart(fig2, use_container_width=True)

    col3, col4 = st.columns(2)

    with col3:
        st.subheader("📅 İzin Gün Bakiyesi")
        df_izin_gun = pd.DataFrame({
            'Şirket': COMPANIES,
            'İzin Günü': [month_data['companies'][c]['izinGun'] for c in COMPANIES]
        })
        fig3 = px.bar(df_izin_gun, x='Şirket', y='İzin Günü', color='İzin Günü',
                      color_continuous_scale='Teal',
                      text=df_izin_gun['İzin Günü'].apply(lambda x: format_number(x, 1)))
        fig3.update_traces(textposition='outside')
        fig3.update_layout(showlegend=False, height=350, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        st.subheader("💰 İzin Ücretleri (Net TL)")
        df_izin_ucret = pd.DataFrame({
            'Şirket': COMPANIES,
            'İzin Ücreti (Net TL)': [month_data['companies'][c]['izinUcret'] for c in COMPANIES]
        })
        fig4 = px.bar(df_izin_ucret, x='Şirket', y='İzin Ücreti (Net TL)', color='İzin Ücreti (Net TL)',
                      color_continuous_scale='Purples',
                      text=df_izin_ucret['İzin Ücreti (Net TL)'].apply(format_tl))
        fig4.update_traces(textposition='outside')
        fig4.update_layout(showlegend=False, height=350, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig4, use_container_width=True)

    col5, col6 = st.columns(2)

    with col5:
        st.subheader("👩 Kadın Oranı")
        df_kadin_plot = pd.DataFrame({
            'Şirket': COMPANIES,
            'Kadın Oranı %': [month_data['companies'][c]['kadinOrani'] for c in COMPANIES]
        })
        fig5 = px.bar(df_kadin_plot, x='Şirket', y='Kadın Oranı %', color='Kadın Oranı %',
                      color_continuous_scale='Magenta',
                      text=df_kadin_plot['Kadın Oranı %'].apply(format_percent))
        fig5.update_traces(textposition='outside')
        fig5.update_layout(showlegend=False, height=350, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig5, use_container_width=True)

    with col6:
        st.subheader("🔁 Ay İçi Giriş & Çıkış")
        df_gc = pd.DataFrame({
            'Şirket': COMPANIES,
            'Giriş': [month_data['companies'][c]['aySekiceGiris'] for c in COMPANIES],
            'Çıkış': [month_data['companies'][c]['aySekiceCikis'] for c in COMPANIES]
        })
        fig6 = go.Figure()
        fig6.add_trace(go.Bar(x=df_gc['Şirket'], y=df_gc['Giriş'], name='Giriş',
                               marker_color='#22c55e',
                               text=df_gc['Giriş'].apply(lambda x: format_number(x, 0)),
                               textposition='outside'))
        fig6.add_trace(go.Bar(x=df_gc['Şirket'], y=df_gc['Çıkış'], name='Çıkış',
                               marker_color='#ef4444',
                               text=df_gc['Çıkış'].apply(lambda x: format_number(x, 0)),
                               textposition='outside'))
        fig6.update_layout(barmode='group', height=350, margin=dict(l=10, r=10, t=30, b=10),
                            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1))
        st.plotly_chart(fig6, use_container_width=True)

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
            'Turnover Toplam %': format_percent(d['turnoverToplam']),
            'Turnover Gönüllü %': format_percent(d['turnoverGonullu']),
            'Net Kök Ücret': format_tl(d['netKokUcret']),
            'İşveren Maliyeti': format_tl(d['isverenMaliyet']),
            'FM Saat': format_number(d['fmSaat'], 1),
            'FM (Net TL)': format_tl(d['fmTlMaliyet']),
            'İzin Gün Bakiyesi': format_number(d['izinGun'], 1),
            'İzin Ücreti (Net TL)': format_tl(d['izinUcret']),
            'Kişi Başı Ort. Maaş (Net TL)': format_tl(d['kisiBasiOrt']),
            'Yıllık Ort. Maaş (Net TL)': format_tl(d['yillikOrtMaas']),
        })

    # Toplam satırı
    total_employees = sum(d['employees'] for d in month_data['companies'].values())
    total_rapor = sum(d['devamsizlik'] for d in month_data['companies'].values()) / len(COMPANIES)
    total_turnover_toplam = sum(d['turnoverToplam'] for d in month_data['companies'].values()) / len(COMPANIES)
    total_turnover_gonullu = sum(d['turnoverGonullu'] for d in month_data['companies'].values()) / len(COMPANIES)
    total_net = sum(d['netKokUcret'] for d in month_data['companies'].values())
    total_isveren = sum(d['isverenMaliyet'] for d in month_data['companies'].values())
    total_fm_saat = sum(d['fmSaat'] for d in month_data['companies'].values())
    total_fm_tl = sum(d['fmTlMaliyet'] for d in month_data['companies'].values())
    total_izin_gun = sum(d['izinGun'] for d in month_data['companies'].values())
    total_izin_ucret = sum(d['izinUcret'] for d in month_data['companies'].values())
    total_yillik_ort = sum(d['yillikOrtMaas'] for d in month_data['companies'].values()) / len(COMPANIES)

    total_row = {
        'Şirket': '⭐ TOPLAM',
        'Çalışan': format_number(total_employees, 0),
        'Raporlu Oran %': format_percent(total_rapor),
        'Turnover Toplam %': format_percent(total_turnover_toplam),
        'Turnover Gönüllü %': format_percent(total_turnover_gonullu),
        'Net Kök Ücret': format_tl(total_net),
        'İşveren Maliyeti': format_tl(total_isveren),
        'FM Saat': format_number(total_fm_saat, 1),
        'FM (Net TL)': format_tl(total_fm_tl),
        'İzin Gün Bakiyesi': format_number(total_izin_gun, 1),
        'İzin Ücreti (Net TL)': format_tl(total_izin_ucret),
        'Kişi Başı Ort. Maaş (Net TL)': format_tl(month_data['kisiBasiOrtGenel']),
        'Yıllık Ort. Maaş (Net TL)': format_tl(total_yillik_ort),
    }
    rows.append(total_row)

    detail_df = pd.DataFrame(rows)
    st.dataframe(detail_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(layout="wide", page_title="İK Dashboard")

# Şirket listesi (Excel'deki sırayla)
COMPANIES = [
    'Aralık Sigorta',
    'Ekim Turizm',
    'Eylül Girişim',
    'Haziran Servis',
    'Intercity Yatırım Holding',
    'Mart Denizcilik'
]
MONTHS = ['Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz']

def normalize_company_name(name):
    """Şirket isimlerini standartlaştır (büyük/küçük harf ve boşlukları temizle)"""
    if not isinstance(name, str):
        return name
    # Excel'de "EKIM TURIZM" veya "Ekim Turizm" olabilir, hepsini düzelt
    name = name.strip()
    # Özel düzeltmeler
    replacements = {
        'EKIM TURIZM': 'Ekim Turizm',
        'HAZIRAN': 'Haziran Servis',
        'Holding': 'Intercity Yatırım Holding'
    }
    return replacements.get(name, name)

def clean_columns(df):
    """Sütun isimlerindeki boşlukları temizle ve ilk satırdaki gereksiz başlıkları düzelt"""
    df.columns = [str(col).strip() for col in df.columns]
    return df

def get_month_cols(df):
    """DataFrame'deki ay sütunlarını (MONTHS ile eşleşen) döndürür."""
    month_cols = []
    for col in df.columns:
        col_clean = str(col).strip()
        for m in MONTHS:
            if col_clean.lower() == m.lower():
                month_cols.append(col)
                break
    return month_cols

def load_data(uploaded_file):
    # ----- 1. GENEL TURNOVER -----
    df_gt = pd.read_excel(uploaded_file, sheet_name='genel.turnover', header=0)  # ilk satır başlık
    df_gt = clean_columns(df_gt)
    # İlk sütun şirket isimleri, son sütun toplam, ayları al
    month_cols = get_month_cols(df_gt)
    if not month_cols:
        raise ValueError("genel.turnover sayfasında ay sütunları bulunamadı.")
    # Şirket isimlerini normalize et
    df_gt['Şirket'] = df_gt.iloc[:, 0].apply(normalize_company_name)
    df_gt = df_gt.set_index('Şirket')
    df_gt = df_gt[month_cols]

    # ----- 2. GÖNÜLLÜ TURNOVER -----
    df_gon = pd.read_excel(uploaded_file, sheet_name='gonullu.turnover', header=0)
    df_gon = clean_columns(df_gon)
    month_cols = get_month_cols(df_gon)
    if not month_cols:
        raise ValueError("gonullu.turnover sayfasında ay sütunları bulunamadı.")
    df_gon['Şirket'] = df_gon.iloc[:, 0].apply(normalize_company_name)
    df_gon = df_gon.set_index('Şirket')
    df_gon = df_gon[month_cols]

    # ----- 3. RAPOR ORANI (devamsızlık) -----
    df_rapor = pd.read_excel(uploaded_file, sheet_name='rapor_oran', header=0)  # ilk satır "rapor.oran", ikinci satır aylar
    df_rapor = clean_columns(df_rapor)
    # İlk sütun şirket isimleri
    df_rapor['Şirket'] = df_rapor.iloc[:, 0].apply(normalize_company_name)
    df_rapor = df_rapor.set_index('Şirket')
    month_cols = get_month_cols(df_rapor)
    if not month_cols:
        raise ValueError("rapor_oran sayfasında ay sütunları bulunamadı.")
    df_rapor = df_rapor[month_cols]

    # ----- 4. ÇALIŞAN SAYISI -----
    df_calisan = pd.read_excel(uploaded_file, sheet_name='calisan.sayisi', header=0)  # ilk satırda "Pers.Say" başlığı var mı? Dosyada B1:H1 boş, A2 "Pers.Say", B2:H2 aylar. O yüzden header=0 ile ilk satırı sütun ismi yapıp sonra düzeltelim.
    df_calisan = clean_columns(df_calisan)
    # İlk sütun şirket isimleri
    df_calisan['Şirket'] = df_calisan.iloc[:, 0].apply(normalize_company_name)
    df_calisan = df_calisan.set_index('Şirket')
    month_cols = get_month_cols(df_calisan)
    if not month_cols:
        raise ValueError("calisan.sayisi sayfasında ay sütunları bulunamadı.")
    df_calisan = df_calisan[month_cols]

    # ----- 5. NET KÖK ÜCRET -----
    df_net = pd.read_excel(uploaded_file, sheet_name='maliyet', header=0)
    df_net = clean_columns(df_net)
    df_net['Şirket'] = df_net.iloc[:, 0].apply(normalize_company_name)
    df_net = df_net.set_index('Şirket')
    month_cols = get_month_cols(df_net)
    if not month_cols:
        raise ValueError("maliyet sayfasında ay sütunları bulunamadı.")
    df_net = df_net[month_cols]

    # ----- 6. İŞVEREN MALİYETİ -----
    df_isv = pd.read_excel(uploaded_file, sheet_name='isveren.maliyet', header=0)
    df_isv = clean_columns(df_isv)
    df_isv['Şirket'] = df_isv.iloc[:, 0].apply(normalize_company_name)
    df_isv = df_isv.set_index('Şirket')
    month_cols = get_month_cols(df_isv)
    if not month_cols:
        raise ValueError("isveren.maliyet sayfasında ay sütunları bulunamadı.")
    df_isv = df_isv[month_cols]

    # ----- 7. FM SAAT -----
    df_fm_saat = pd.read_excel(uploaded_file, sheet_name='fm.saat', header=0)
    df_fm_saat = clean_columns(df_fm_saat)
    df_fm_saat['Şirket'] = df_fm_saat.iloc[:, 0].apply(normalize_company_name)
    df_fm_saat = df_fm_saat.set_index('Şirket')
    month_cols = get_month_cols(df_fm_saat)
    if not month_cols:
        raise ValueError("fm.saat sayfasında ay sütunları bulunamadı.")
    df_fm_saat = df_fm_saat[month_cols]

    # ----- 8. FM TL MALİYET -----
    df_fm_tl = pd.read_excel(uploaded_file, sheet_name='fm.maliyet', header=0)
    df_fm_tl = clean_columns(df_fm_tl)
    df_fm_tl['Şirket'] = df_fm_tl.iloc[:, 0].apply(normalize_company_name)
    df_fm_tl = df_fm_tl.set_index('Şirket')
    month_cols = get_month_cols(df_fm_tl)
    if not month_cols:
        raise ValueError("fm.maliyet sayfasında ay sütunları bulunamadı.")
    df_fm_tl = df_fm_tl[month_cols]

    # ----- 9. İZİN GÜN -----
    df_izin_gun = pd.read_excel(uploaded_file, sheet_name='izin_gun', header=0)
    df_izin_gun = clean_columns(df_izin_gun)
    df_izin_gun['Şirket'] = df_izin_gun.iloc[:, 0].apply(normalize_company_name)
    df_izin_gun = df_izin_gun.set_index('Şirket')
    month_cols = get_month_cols(df_izin_gun)
    if not month_cols:
        raise ValueError("izin_gun sayfasında ay sütunları bulunamadı.")
    df_izin_gun = df_izin_gun[month_cols]

    # ----- 10. İZİN ÜCRET -----
    df_izin_ucret = pd.read_excel(uploaded_file, sheet_name='izin_ucret', header=0)
    df_izin_ucret = clean_columns(df_izin_ucret)
    df_izin_ucret['Şirket'] = df_izin_ucret.iloc[:, 0].apply(normalize_company_name)
    df_izin_ucret = df_izin_ucret.set_index('Şirket')
    month_cols = get_month_cols(df_izin_ucret)
    if not month_cols:
        raise ValueError("izin_ucret sayfasında ay sütunları bulunamadı.")
    df_izin_ucret = df_izin_ucret[month_cols]

    # ----- TOPLAM SATIRLARI (son satırlar) -----
    # rapor_oran - "Aylık Genel Rapor Oranı" satırı (dosyada 8. satır, indeks 7)
    df_rapor_son = pd.read_excel(uploaded_file, sheet_name='rapor_oran', skiprows=7, header=None, nrows=1)
    genel_rapor = df_rapor_son.iloc[0, 1:8].values  # ilk sütun boş

    # calisan.sayisi - "TOPLAM ÇALIŞAN SAYISI" satırı (dosyada 9. satır, indeks 8)
    df_calisan_son = pd.read_excel(uploaded_file, sheet_name='calisan.sayisi', skiprows=8, header=None, nrows=1)
    toplam_calisan = df_calisan_son.iloc[0, 1:8].values

    # izin_gun - "TOPLAM_GUN" satırı (dosyada 8. satır, indeks 7)
    df_izin_gun_son = pd.read_excel(uploaded_file, sheet_name='izin_gun', skiprows=7, header=None, nrows=1)
    toplam_izin_gun = df_izin_gun_son.iloc[0, 1:8].values

    # izin_ucret - "TOPLAM" satırı (dosyada 8. satır, indeks 7)
    df_izin_ucret_son = pd.read_excel(uploaded_file, sheet_name='izin_ucret', skiprows=7, header=None, nrows=1)
    toplam_izin_ucret = df_izin_ucret_son.iloc[0, 1:8].values

    # ----- TÜM VERİLERİ BİR SÖZLÜKTE TOPLA -----
    data = {}
    for idx, m in enumerate(MONTHS):
        # Her şirket için verileri topla
        employees = []
        devamsizlik = []
        turnover_toplam = []
        turnover_gonullu = []
        net_kok = []
        isveren_maliyet = []
        fm_saat = []
        fm_tl = []
        izin_gun = []
        izin_ucret = []

        for comp in COMPANIES:
            # Şirket indeksini kontrol et
            if comp in df_calisan.index:
                employees.append(df_calisan.loc[comp, m])
                devamsizlik.append(df_rapor.loc[comp, m] * 100)  # yüzde
                turnover_toplam.append(df_gt.loc[comp, m] * 100)
                turnover_gonullu.append(df_gon.loc[comp, m] * 100)
                net_kok.append(df_net.loc[comp, m])
                isveren_maliyet.append(df_isv.loc[comp, m])
                fm_saat.append(df_fm_saat.loc[comp, m])
                fm_tl.append(df_fm_tl.loc[comp, m])
                izin_gun.append(df_izin_gun.loc[comp, m])
                izin_ucret.append(df_izin_ucret.loc[comp, m])
            else:
                # Şirket bulunamazsa 0 ekle (hata vermesin)
                employees.append(0)
                devamsizlik.append(0)
                turnover_toplam.append(0)
                turnover_gonullu.append(0)
                net_kok.append(0)
                isveren_maliyet.append(0)
                fm_saat.append(0)
                fm_tl.append(0)
                izin_gun.append(0)
                izin_ucret.append(0)

        data[m] = {
            'employees': employees,
            'devamsizlik': devamsizlik,
            'turnoverToplam': turnover_toplam,
            'turnoverGonullu': turnover_gonullu,
            'netKokUcret': net_kok,
            'isverenMaliyet': isveren_maliyet,
            'fmSaat': fm_saat,
            'fmTlMaliyet': fm_tl,
            'izinGun': izin_gun,
            'izinUcret': izin_ucret,
            'genelRaporOran': genel_rapor[idx] * 100,
            'toplamCalisan': toplam_calisan[idx],
            'toplamIzinGun': toplam_izin_gun[idx],
            'toplamIzinUcret': toplam_izin_ucret[idx],
        }
    return data


def main():
    st.title("📊 İK Konsolide Dashboard")
    st.markdown("---")

    uploaded_file = st.file_uploader(
        "Excel dosyasını yükleyin (dashboard_27082026.xlsx)",
        type=["xlsx"]
    )

    if uploaded_file is None:
        st.info("Lütfen bir Excel dosyası yükleyin.")
        return

    try:
        data = load_data(uploaded_file)
        st.success("✅ Veri başarıyla okundu!")
    except Exception as e:
        st.error(f"Hata oluştu: {e}")
        st.stop()

    selected_month = st.selectbox("📅 Ay Seçin", MONTHS, index=MONTHS.index("Temmuz"))
    month_data = data[selected_month]

    # ----- KPI KARTLARI -----
    col1, col2, col3, col4, col5, col6, col7, col8 = st.columns(8)
    col1.metric("👥 Çalışan", f"{month_data['toplamCalisan']:,.0f}", f"{selected_month}")
    col2.metric("📊 Genel rapor oranı", f"{month_data['genelRaporOran']:.2f}%")
    col3.metric("💼 İşveren Maliyeti", f"{sum(month_data['isverenMaliyet']):,.0f} ₺")
    col4.metric("💰 Net Kök Ücret", f"{sum(month_data['netKokUcret']):,.0f} ₺")
    col5.metric("⏱️ FM_Saat", f"{sum(month_data['fmSaat']):,.1f}")
    col6.metric("💸 FM_TL Maliyet", f"{sum(month_data['fmTlMaliyet']):,.0f} ₺")
    col7.metric("📅 İzin Gün", f"{sum(month_data['izinGun']):,.1f}")
    col8.metric("💎 İzin Ücreti", f"{sum(month_data['izinUcret']):,.0f} ₺")

    st.markdown("---")

    # ----- GRAFİKLER (SATIR 1) -----
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📊 Rapor Oranı")
        fig1 = px.bar(
            x=COMPANIES,
            y=month_data['devamsizlik'],
            labels={'x': 'Şirket', 'y': 'Devamsızlık %'},
            color=month_data['devamsizlik'],
            color_continuous_scale='Blues',
            text_auto='.2f'
        )
        fig1.update_layout(showlegend=False, height=350, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig1, use_container_width=True)

    with col2:
        st.subheader("🔄 Turnover (Toplam & Gönüllü)")
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(
            x=COMPANIES,
            y=month_data['turnoverToplam'],
            name='Toplam Turnover',
            marker_color='#f59e0b',
            text=[f"{v:.2f}%" for v in month_data['turnoverToplam']],
            textposition='outside'
        ))
        fig2.add_trace(go.Bar(
            x=COMPANIES,
            y=month_data['turnoverGonullu'],
            name='Gönüllü Turnover',
            marker_color='#ec4899',
            text=[f"{v:.2f}%" for v in month_data['turnoverGonullu']],
            textposition='outside'
        ))
        fig2.update_layout(
            barmode='group',
            height=350,
            margin=dict(l=10, r=10, t=30, b=10),
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)
        )
        st.plotly_chart(fig2, use_container_width=True)

    # ----- GRAFİKLER (SATIR 2) -----
    col3, col4 = st.columns(2)

    with col3:
        st.subheader("📈 Aylık Trendler")
        trend_df = pd.DataFrame({
            'Ay': MONTHS,
            'Devamsızlık (%)': [data[m]['genelRaporOran'] for m in MONTHS],
            'Turnover (%)': [sum(data[m]['turnoverToplam']) / len(COMPANIES) for m in MONTHS]
        })
        fig3 = px.line(
            trend_df,
            x='Ay',
            y=['Devamsızlık (%)', 'Turnover (%)'],
            markers=True,
            color_discrete_map={'Devamsızlık (%)': '#3b82f6', 'Turnover (%)': '#f59e0b'}
        )
        fig3.update_layout(height=350, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        st.subheader("💰 İzin Ücreti Dağılımı")
        izin_df = pd.DataFrame({
            'Şirket': COMPANIES,
            'İzin Ücreti': month_data['izinUcret']
        })
        fig4 = px.pie(
            izin_df,
            values='İzin Ücreti',
            names='Şirket',
            hole=0.4,
            color_discrete_sequence=px.colors.qualitative.Set2
        )
        fig4.update_layout(height=350, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig4, use_container_width=True)

    st.markdown("---")

    # ----- HEATMAP -----
    st.subheader("📋 Aylık Devamsızlık Raporu (Heatmap)")
    heatmap_data = {m: data[m]['devamsizlik'] for m in MONTHS}
    df_heat = pd.DataFrame(heatmap_data, index=COMPANIES)
    st.dataframe(
        df_heat.style.background_gradient(cmap='RdYlGn_r', axis=None, vmin=0, vmax=5)
        .format("{:.2f}%"),
        use_container_width=True,
        height=300
    )

    st.markdown("---")

    # ----- DETAY TABLOSU -----
    st.subheader(f"📋 {selected_month} Detaylı Metrikler")
    detail_df = pd.DataFrame({
        'Şirket': COMPANIES,
        'Çalışan': month_data['employees'],
        'Devamsızlık %': [f"{v:.2f}%" for v in month_data['devamsizlik']],
        'Turnover % (Toplam)': [f"{v:.2f}%" for v in month_data['turnoverToplam']],
        'Turnover % (Gönüllü)': [f"{v:.2f}%" for v in month_data['turnoverGonullu']],
        'Net Kök Ücret': [f"{v:,.0f}" for v in month_data['netKokUcret']],
        'FM_Saat': month_data['fmSaat'],
        'FM_TL Maliyet': [f"{v:,.0f}" for v in month_data['fmTlMaliyet']],
        'İzin Gün': month_data['izinGun'],
        'İzin Ücreti': [f"{v:,.0f}" for v in month_data['izinUcret']],
        'İşveren Maliyeti': [f"{v:,.0f}" for v in month_data['isverenMaliyet']],
    })
    st.dataframe(detail_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
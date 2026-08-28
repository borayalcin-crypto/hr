import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(layout="wide", page_title="İK Dashboard")

# Sabitler
COMPANIES = [
    'Aralık Sigorta', 'Ekim Turizm', 'Eylül Girişim',
    'Haziran Servis', 'Intercity Yatırım Holding', 'Mart Denizcilik'
]
MONTHS = ['Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz']


def load_data(uploaded_file):
    """Excel dosyasındaki tüm sayfaları okuyarak veriyi hazırlar."""
    # 1. Genel Turnover
    df_gt = pd.read_excel(uploaded_file, sheet_name='genel.turnover', skiprows=2, nrows=6)
    df_gt = df_gt.set_index(df_gt.columns[0])[MONTHS]

    # 2. Gönüllü Turnover
    df_gon = pd.read_excel(uploaded_file, sheet_name='gonullu.turnover', skiprows=2, nrows=6)
    df_gon = df_gon.set_index(df_gon.columns[0])[MONTHS]

    # 3. Rapor Oranı (Devamsızlık)
    df_rapor = pd.read_excel(uploaded_file, sheet_name='rapor_oran', skiprows=1, nrows=6)
    df_rapor = df_rapor.set_index(df_rapor.columns[0])[MONTHS]

    # 4. Çalışan Sayısı
    df_calisan = pd.read_excel(uploaded_file, sheet_name='calisan.sayisi', skiprows=2, nrows=6)
    df_calisan = df_calisan.set_index(df_calisan.columns[0])[MONTHS]

    # 5. Net Kök Ücret
    df_net = pd.read_excel(uploaded_file, sheet_name='maliyet', skiprows=1, nrows=6)
    df_net = df_net.set_index(df_net.columns[0])[MONTHS]

    # 6. İşveren Maliyeti
    df_isv = pd.read_excel(uploaded_file, sheet_name='isveren.maliyet', skiprows=1, nrows=6)
    df_isv = df_isv.set_index(df_isv.columns[0])[MONTHS]

    # 7. FM Saat
    df_fm_saat = pd.read_excel(uploaded_file, sheet_name='fm.saat', skiprows=1, nrows=6)
    df_fm_saat = df_fm_saat.set_index(df_fm_saat.columns[0])[MONTHS]

    # 8. FM TL Maliyet
    df_fm_tl = pd.read_excel(uploaded_file, sheet_name='fm.maliyet', skiprows=1, nrows=6)
    df_fm_tl = df_fm_tl.set_index(df_fm_tl.columns[0])[MONTHS]

    # 9. İzin Gün
    df_izin_gun = pd.read_excel(uploaded_file, sheet_name='izin_gun', skiprows=1, nrows=6)
    df_izin_gun = df_izin_gun.set_index(df_izin_gun.columns[0])[MONTHS]

    # 10. İzin Ücreti
    df_izin_ucret = pd.read_excel(uploaded_file, sheet_name='izin_ucret', skiprows=1, nrows=6)
    df_izin_ucret = df_izin_ucret.set_index(df_izin_ucret.columns[0])[MONTHS]

    # Genel Rapor Oranı (son satır)
    df_rapor_son = pd.read_excel(uploaded_file, sheet_name='rapor_oran', skiprows=7, nrows=1)
    genel_rapor = df_rapor_son.iloc[0, 1:8].values  # 7 ay

    # Toplam Çalışan (son satır)
    df_calisan_son = pd.read_excel(uploaded_file, sheet_name='calisan.sayisi', skiprows=8, nrows=1)
    toplam_calisan = df_calisan_son.iloc[0, 1:8].values

    # Toplam İzin Günü (son satır)
    df_izin_gun_son = pd.read_excel(uploaded_file, sheet_name='izin_gun', skiprows=7, nrows=1)
    toplam_izin_gun = df_izin_gun_son.iloc[0, 1:8].values

    # Toplam İzin Ücreti (son satır)
    df_izin_ucret_son = pd.read_excel(uploaded_file, sheet_name='izin_ucret', skiprows=7, nrows=1)
    toplam_izin_ucret = df_izin_ucret_son.iloc[0, 1:8].values

    # Tüm ayları içeren sözlük
    data = {}
    for idx, m in enumerate(MONTHS):
        data[m] = {
            'employees': [df_calisan.loc[c, m] for c in COMPANIES],
            'devamsizlik': [df_rapor.loc[c, m] * 100 for c in COMPANIES],
            'turnoverToplam': [df_gt.loc[c, m] * 100 for c in COMPANIES],
            'turnoverGonullu': [df_gon.loc[c, m] * 100 for c in COMPANIES],
            'netKokUcret': [df_net.loc[c, m] for c in COMPANIES],
            'isverenMaliyet': [df_isv.loc[c, m] for c in COMPANIES],
            'fmSaat': [df_fm_saat.loc[c, m] for c in COMPANIES],
            'fmTlMaliyet': [df_fm_tl.loc[c, m] for c in COMPANIES],
            'izinGun': [df_izin_gun.loc[c, m] for c in COMPANIES],
            'izinUcret': [df_izin_ucret.loc[c, m] for c in COMPANIES],
            'genelRaporOran': genel_rapor[idx] * 100,
            'toplamCalisan': toplam_calisan[idx],
            'toplamIzinGun': toplam_izin_gun[idx],
            'toplamIzinUcret': toplam_izin_ucret[idx],
        }
    return data


def main():
    st.title("📊 İK Konsolide Dashboard")
    st.markdown("---")

    # Dosya yükleme
    uploaded_file = st.file_uploader(
        "Excel dosyasını yükleyin (dashboard_27082026.xlsx)",
        type=["xlsx"]
    )

    if uploaded_file is None:
        st.info("Lütfen bir Excel dosyası yükleyin.")
        return

    # Veriyi yükle
    try:
        data = load_data(uploaded_file)
        st.success("✅ Veri başarıyla okundu!")
    except Exception as e:
        st.error(f"Hata oluştu: {e}")
        return

    # Ay filtresi
    selected_month = st.selectbox("📅 Ay Seçin", MONTHS, index=MONTHS.index("Temmuz"))
    month_data = data[selected_month]

    # ----- KPI KARTLARI -----
    col1, col2, col3, col4, col5, col6, col7, col8 = st.columns(8)

    col1.metric(
        "👥 Çalışan",
        f"{month_data['toplamCalisan']:,.0f}",
        f"{selected_month}"
    )
    col2.metric(
        "📊 Genel rapor oranı",
        f"{month_data['genelRaporOran']:.2f}%"
    )
    col3.metric(
        "💼 İşveren Maliyeti",
        f"{sum(month_data['isverenMaliyet']):,.0f} ₺"
    )
    col4.metric(
        "💰 Net Kök Ücret",
        f"{sum(month_data['netKokUcret']):,.0f} ₺"
    )
    col5.metric(
        "⏱️ FM_Saat",
        f"{sum(month_data['fmSaat']):,.1f}"
    )
    col6.metric(
        "💸 FM_TL Maliyet",
        f"{sum(month_data['fmTlMaliyet']):,.0f} ₺"
    )
    col7.metric(
        "📅 İzin Gün",
        f"{sum(month_data['izinGun']):,.1f}"
    )
    col8.metric(
        "💎 İzin Ücreti",
        f"{sum(month_data['izinUcret']):,.0f} ₺"
    )

    st.markdown("---")

    # ----- GRAFİKLER (1. SATIR) -----
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

    # ----- GRAFİKLER (2. SATIR) -----
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
            color_discrete_map={
                'Devamsızlık (%)': '#3b82f6',
                'Turnover (%)': '#f59e0b'
            }
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

    # ----- HEATMAP (Devamsızlık Raporu) -----
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
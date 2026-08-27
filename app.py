import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(layout="wide", page_title="İK Konsolide Dashboard")

# ----- 1. EXCEL VERİ OKUYUCU (Sizin dosya yapınıza özel) -----
def load_data_from_excel(uploaded_file):
    companies = ['Aralık Sigorta', 'Ekim Turizm', 'Eylül Girişim', 'Haziran Servis', 'Intercity Holding', 'Mart Denizcilik']
    months = ['Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz']

    # 1. Çalışan Sayısı
    df_calisan = pd.read_excel(uploaded_file, sheet_name='calisan.sayisi', skiprows=1, nrows=6)
    df_calisan = df_calisan.set_index(df_calisan.columns[0])[months]

    # 2. Devamsızlık
    df_devamsizlik = pd.read_excel(uploaded_file, sheet_name='devamsızlık', skiprows=1, nrows=6)
    df_devamsizlik = df_devamsizlik.set_index(df_devamsizlik.columns[0])[months]

    # 3. Turnover - Toplam
    df_turnover_toplam = pd.read_excel(uploaded_file, sheet_name='turnover', skiprows=2, nrows=6)
    df_turnover_toplam = df_turnover_toplam.set_index(df_turnover_toplam.columns[0])[months]

    # 4. Turnover - Gönüllü
    df_turnover_gonullu = pd.read_excel(uploaded_file, sheet_name='turnover', skiprows=10, nrows=6)
    df_turnover_gonullu = df_turnover_gonullu.set_index(df_turnover_gonullu.columns[0])[months]

    # 5. Maliyet - Net Kök Ücret
    df_net = pd.read_excel(uploaded_file, sheet_name='maliyet', skiprows=0, nrows=6)
    df_net = df_net.set_index(df_net.columns[0])[months]

    # 6. Maliyet - İşveren (her şey dahil)
    df_isveren = pd.read_excel(uploaded_file, sheet_name='maliyet', skiprows=8, nrows=6)
    df_isveren = df_isveren.set_index(df_isveren.columns[0])[months]

    # 7. Maliyet - FM Saat
    df_fm_saat = pd.read_excel(uploaded_file, sheet_name='maliyet', skiprows=15, nrows=6)
    df_fm_saat = df_fm_saat.set_index(df_fm_saat.columns[0])[months]

    # 8. Maliyet - FM TL
    df_fm_tl = pd.read_excel(uploaded_file, sheet_name='maliyet', skiprows=22, nrows=6)
    df_fm_tl = df_fm_tl.set_index(df_fm_tl.columns[0])[months]

    # 9. İzin (sabit, kümülatif)
    df_izin = pd.read_excel(uploaded_file, sheet_name='izin_ucret', nrows=6)
    df_izin = df_izin.set_index(df_izin.columns[0])
    kalan_izin = df_izin['Toplam Kalan Izin'].to_dict()
    izin_ucreti = df_izin['Toplam izin_ucret'].to_dict()

    # Tüm verileri tek bir sözlükte topla
    month_data = {}
    for m in months:
        month_data[m] = {
            'employees': [df_calisan.loc[c, m] for c in companies],
            'devamsizlik': [df_devamsizlik.loc[c, m] * 100 for c in companies],  # yüzde
            'turnoverToplam': [df_turnover_toplam.loc[c, m] * 100 for c in companies],
            'turnoverGonullu': [df_turnover_gonullu.loc[c, m] * 100 for c in companies],
            'netKokUcret': [df_net.loc[c, m] for c in companies],
            'isverenMaliyet': [df_isveren.loc[c, m] for c in companies],
            'fmSaat': [df_fm_saat.loc[c, m] for c in companies],
            'fmTlMaliyet': [df_fm_tl.loc[c, m] for c in companies],
            'kalanIzin': [kalan_izin.get(c, 0) for c in companies],
            'izinUcreti': [izin_ucreti.get(c, 0) for c in companies],
        }
    return companies, month_data

# ----- 2. YEDEK VERİ (Excel yoksa çalışsın diye - HTML'dekiyle aynı) -----
FALLBACK_COMPANIES = ['Aralık Sigorta', 'Ekim Turizm', 'Eylül Girişim', 'Haziran Servis', 'Intercity Holding', 'Mart Denizcilik']
FALLBACK_DATA = {
    "Ocak": {
        "employees": [10, 116, 28, 274, 3, 3],
        "devamsizlik": [1.4285, 1.9293, 0.9615, 1.9090, 0, 0],
        "turnoverToplam": [0, 0, 0, 1.0948, 0, 0],
        "turnoverGonullu": [0, 0, 0, 1.0948, 0, 0],
        "netKokUcret": [1180900, 15269685.5, 2419500, 19313625, 1865000, 460000],
        "fmSaat": [0, 474, 772.5, 4081.5, 0, 0],
        "fmTlMaliyet": [0, 363027.26, 542529.83, 2732243.01, 0, 0],
        "isverenMaliyet": [1853402, 25237096, 4510226, 33607563, 2695689, 794694],
        "kalanIzin": [279, 1598, 763.5, 1650.5, 16, 91.5],
        "izinUcreti": [1621466, 9415569, 5366650, 6178472, 405833, 752333]
    },
    "Şubat": {
        "employees": [10, 115, 28, 275, 3, 3],
        "devamsizlik": [2.5, 1.0, 0.2976, 1.4393, 0, 0],
        "turnoverToplam": [0, 7.0175, 7.1428, 6.9090, 0, 0],
        "turnoverGonullu": [0, 5.2631, 7.1428, 6.1818, 0, 0],
        "netKokUcret": [1180900, 15203095.5, 2419500, 19295025, 1865000, 460000],
        "fmSaat": [0, 300, 263, 2242.5, 0, 0],
        "fmTlMaliyet": [0, 228001.19, 187100.63, 1526348.50, 0, 0],
        "isverenMaliyet": [1954196, 31365084, 5714847, 36937458, 2944608, 811952],
        "kalanIzin": [279, 1598, 763.5, 1650.5, 16, 91.5],
        "izinUcreti": [1621466, 9415569, 5366650, 6178472, 405833, 752333]
    },
    "Mart": {
        "employees": [10, 107, 27, 263, 3, 3],
        "devamsizlik": [1.4285, 1.9136, 1.1396, 1.4624, 0, 0],
        "turnoverToplam": [0, 3.7735, 0, 1.9157, 0, 0],
        "turnoverGonullu": [0, 3.7735, 0, 1.5325, 0, 0],
        "netKokUcret": [1180900, 14473495.5, 2143500, 18451830, 1865000, 460000],
        "fmSaat": [0, 138, 615, 2505, 0, 0],
        "fmTlMaliyet": [0, 147366.60, 521046.68, 1848478.93, 0, 0],
        "isverenMaliyet": [2334423, 32871909, 5192249, 44170793, 3276364, 922506],
        "kalanIzin": [279, 1598, 763.5, 1650.5, 16, 91.5],
        "izinUcreti": [1621466, 9415569, 5366650, 6178472, 405833, 752333]
    },
    "Nisan": {
        "employees": [10, 105, 27, 259, 3, 3],
        "devamsizlik": [2.7777, 1.2169, 0.4444, 1.5135, 0, 0],
        "turnoverToplam": [0, 0, 0, 2.6923, 0, 0],
        "turnoverGonullu": [0, 0, 0, 1.9230, 0, 0],
        "netKokUcret": [1180900, 13804690.5, 2143500, 18159425, 1865000, 460000],
        "fmSaat": [0, 211, 670, 3568, 0, 0],
        "fmTlMaliyet": [0, 197587.62, 572785.73, 2620720.99, 0, 0],
        "isverenMaliyet": [2057869, 25311546, 4414185, 37443196, 3326355, 850467],
        "kalanIzin": [279, 1598, 763.5, 1650.5, 16, 91.5],
        "izinUcreti": [1621466, 9415569, 5366650, 6178472, 405833, 752333]
    },
    "Mayıs": {
        "employees": [10, 105, 30, 253, 3, 4],
        "devamsizlik": [1.6666, 1.3756, 2.1212, 2.9284, 0, 0],
        "turnoverToplam": [0, 0, 3.4482, 1.5748, 0, 0],
        "turnoverGonullu": [0, 0, 0, 0.3937, 0, 0],
        "netKokUcret": [1180900, 13803985.5, 2233500, 17834525, 1865000, 610000],
        "fmSaat": [0, 205, 1314, 3162, 0, 0],
        "fmTlMaliyet": [0, 193903.19, 1129547.19, 2439366.43, 0, 0],
        "isverenMaliyet": [2384975, 28204222, 5707498, 40231469, 3401484, 1167855],
        "kalanIzin": [279, 1598, 763.5, 1650.5, 16, 91.5],
        "izinUcreti": [1621466, 9415569, 5366650, 6178472, 405833, 752333]
    },
    "Haziran": {
        "employees": [10, 104, 33, 255, 3, 5],
        "devamsizlik": [1.8181, 1.9230, 1.2820, 4.6455, 0, 0],
        "turnoverToplam": [0, 0.9708, 3.1250, 1.1764, 0, 0],
        "turnoverGonullu": [0, 0, 0, 0.7843, 0, 0],
        "netKokUcret": [1180900, 13741785.5, 2311500, 17907525, 1865000, 965000],
        "fmSaat": [0, 164, 906, 2853, 0, 0],
        "fmTlMaliyet": [0, 139987.09, 685084.68, 2217103.36, 0, 0],
        "isverenMaliyet": [2183357, 26367127, 4844198, 34300360, 3522300, 1764875],
        "kalanIzin": [279, 1598, 763.5, 1650.5, 16, 91.5],
        "izinUcreti": [1621466, 9415569, 5366650, 6178472, 405833, 752333]
    },
    "Temmuz": {
        "employees": [10, 104, 34, 254, 3, 5],
        "devamsizlik": [2.2727, 0.8741, 0.4524, 2.9830, 0, 0],
        "turnoverToplam": [0, 0, 0, 0.7874, 0, 0],
        "turnoverGonullu": [0, 0, 0, 0.7874, 0, 0],
        "netKokUcret": [1417000, 16964450, 2854500, 21283825, 2100000, 965000],
        "fmSaat": [0, 409, 795, 3420, 0, 0],
        "fmTlMaliyet": [0, 461991.08, 659068.58, 3195411.67, 0, 0],
        "isverenMaliyet": [2665486, 32241882, 5867563, 43692160, 3975122, 1860734],
        "kalanIzin": [279, 1598, 763.5, 1650.5, 16, 91.5],
        "izinUcreti": [1621466, 9415569, 5366650, 6178472, 405833, 752333]
    }
}

# ----- 3. STREAMLIT ARAYÜZÜ -----
def main():
    st.title("📊 İK Konsolide Dashboard")
    st.markdown("---")

    # Dosya yükleme
    uploaded_file = st.file_uploader("Excel dosyasını yükleyin (isteğe bağlı)", type=["xlsx"])
    
    if uploaded_file is not None:
        try:
            companies, month_data = load_data_from_excel(uploaded_file)
            st.success("✅ Excel başarıyla okundu!")
        except Exception as e:
            st.error(f"Excel okunurken hata oluştu: {e}. Örnek veriler kullanılıyor.")
            companies, month_data = FALLBACK_COMPANIES, FALLBACK_DATA
    else:
        st.info("📂 Excel yüklenmedi, örnek veriler gösteriliyor.")
        companies, month_data = FALLBACK_COMPANIES, FALLBACK_DATA

    # Ay filtresi
    months = list(month_data.keys())
    selected_month = st.selectbox("📅 Ay Seçin", months, index=months.index("Temmuz"))
    data = month_data[selected_month]

    # ----- KPI KARTLARI -----
    col1, col2, col3, col4, col5, col6, col7, col8 = st.columns(8)
    
    total_emp = sum(data['employees'])
    genel_dev = sum(data['devamsizlik']) / len(data['devamsizlik'])
    total_isveren = sum(data['isverenMaliyet'])
    total_net = sum(data['netKokUcret'])
    total_fm_saat = sum(data['fmSaat'])
    total_fm_tl = sum(data['fmTlMaliyet'])
    total_izin = sum(data['kalanIzin'])
    total_izin_ucret = sum(data['izinUcreti'])

    col1.metric("👥 Çalışan", f"{total_emp}", f"{selected_month}")
    col2.metric("📊 Genel rapor oranı", f"{genel_dev:.2f}%")
    col3.metric("💼 İşveren Maliyeti", f"{total_isveren:,.0f} ₺")
    col4.metric("💰 Net Kök Ücret", f"{total_net:,.0f} ₺")
    col5.metric("⏱️ FM_Saat", f"{total_fm_saat:.1f}")
    col6.metric("💸 FM_TL Maliyet", f"{total_fm_tl:,.0f} ₺")
    col7.metric("📅 Kalan İzin", f"{total_izin:.1f} gün")
    col8.metric("💎 İzin Ücreti Yükü", f"{total_izin_ucret:,.0f} ₺")

    st.markdown("---")

    # ----- GRAFİKLER (1. SATIR) -----
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📊 Rapor Oranı")
        fig1 = px.bar(
            x=companies, y=data['devamsizlik'],
            labels={'x': 'Şirket', 'y': 'Devamsızlık %'},
            color=data['devamsizlik'],
            color_continuous_scale='Blues',
            text_auto='.2f'
        )
        fig1.update_layout(showlegend=False, height=350, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig1, use_container_width=True)

    with col2:
        st.subheader("🔄 Seçilen Aya Göre Turnover")
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(
            x=companies, y=data['turnoverToplam'],
            name='Toplam Turnover',
            marker_color='#f59e0b',
            text=[f"{v:.2f}%" for v in data['turnoverToplam']],
            textposition='outside'
        ))
        fig2.add_trace(go.Bar(
            x=companies, y=data['turnoverGonullu'],
            name='Gönüllü Turnover',
            marker_color='#ec4899',
            text=[f"{v:.2f}%" for v in data['turnoverGonullu']],
            textposition='outside'
        ))
        fig2.update_layout(barmode='group', height=350, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig2, use_container_width=True)

    # ----- GRAFİKLER (2. SATIR) -----
    col3, col4 = st.columns(2)

    with col3:
        st.subheader("📈 Genel Devamsızlık Trendi")
        trend_df = pd.DataFrame({
            'Ay': months,
            'Devamsızlık (%)': [sum(month_data[m]['devamsizlik']) / len(month_data[m]['devamsizlik']) for m in months],
            'Turnover (%)': [sum(month_data[m]['turnoverToplam']) / len(month_data[m]['turnoverToplam']) for m in months]
        })
        fig3 = px.line(trend_df, x='Ay', y=['Devamsızlık (%)', 'Turnover (%)'],
                       markers=True, color_discrete_map={'Devamsızlık (%)': '#3b82f6', 'Turnover (%)': '#f59e0b'})
        fig3.update_layout(height=350, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        st.subheader("💰 İzin Ücreti Dağılımı")
        izin_df = pd.DataFrame({
            'Şirket': companies,
            'İzin Ücreti': data['izinUcreti']
        })
        fig4 = px.pie(izin_df, values='İzin Ücreti', names='Şirket', hole=0.4,
                      color_discrete_sequence=px.colors.qualitative.Set2)
        fig4.update_layout(height=350, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig4, use_container_width=True)

    st.markdown("---")

    # ----- HEATMAP (Devamsızlık Raporu) -----
    st.subheader("📋 Aylık Devamsızlık Raporu (Heatmap)")
    heatmap_data = {m: month_data[m]['devamsizlik'] for m in months}
    df_heat = pd.DataFrame(heatmap_data, index=companies)
    st.dataframe(
        df_heat.style.background_gradient(cmap='RdYlGn_r', axis=None, vmin=0, vmax=5)
        .format("{:.2f}%"),
        use_container_width=True, height=300
    )

    st.markdown("---")

    # ----- DETAY TABLOSU -----
    st.subheader(f"📋 {selected_month} Detaylı Metrikler")
    detail_df = pd.DataFrame({
        'Şirket': companies,
        'Çalışan': data['employees'],
        'Devamsızlık %': [f"{v:.2f}%" for v in data['devamsizlik']],
        'Turnover % (Toplam)': [f"{v:.2f}%" for v in data['turnoverToplam']],
        'Turnover % (Gönüllü)': [f"{v:.2f}%" for v in data['turnoverGonullu']],
        'Net Kök Ücret': [f"{v:,.0f}" for v in data['netKokUcret']],
        'FM_Saat': data['fmSaat'],
        'FM_TL Maliyet': [f"{v:,.0f}" for v in data['fmTlMaliyet']],
        'Kalan İzin': data['kalanIzin'],
        'İzin Ücreti': [f"{v:,.0f}" for v in data['izinUcreti']],
        'İşveren Maliyeti': [f"{v:,.0f}" for v in data['isverenMaliyet']],
    })
    st.dataframe(detail_df, use_container_width=True, hide_index=True)

if __name__ == "__main__":
    main()
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(layout="wide", page_title="İK Dashboard")

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

def load_data(uploaded_file):
    # ----- 1. GENEL TURNOVER -----
    df_gt = pd.read_excel(uploaded_file, sheet_name='genel.turnover', header=0)
    df_gt = clean_columns(df_gt)
    month_cols = get_month_cols(df_gt)
    if not month_cols:
        raise ValueError("genel.turnover sayfasında ay sütunları bulunamadı.")
    df_gt['Şirket'] = df_gt.iloc[:, 0].apply(normalize_company_name)
    df_gt = df_gt.set_index('Şirket')[month_cols]

    # ----- 2. GÖNÜLLÜ TURNOVER -----
    df_gon = pd.read_excel(uploaded_file, sheet_name='gonullu.turnover', header=0)
    df_gon = clean_columns(df_gon)
    month_cols = get_month_cols(df_gon)
    df_gon['Şirket'] = df_gon.iloc[:, 0].apply(normalize_company_name)
    df_gon = df_gon.set_index('Şirket')[month_cols]

    # ----- 3. RAPOR ORANI -----
    df_rapor = pd.read_excel(uploaded_file, sheet_name='rapor_oran', header=0)
    df_rapor = clean_columns(df_rapor)
    month_cols = get_month_cols(df_rapor)
    df_rapor['Şirket'] = df_rapor.iloc[:, 0].apply(normalize_company_name)
    df_rapor = df_rapor.set_index('Şirket')[month_cols]

    # ----- 4. ÇALIŞAN SAYISI -----
    df_calisan = pd.read_excel(uploaded_file, sheet_name='calisan.sayisi', header=0)
    df_calisan = clean_columns(df_calisan)
    month_cols = get_month_cols(df_calisan)
    df_calisan['Şirket'] = df_calisan.iloc[:, 0].apply(normalize_company_name)
    df_calisan = df_calisan.set_index('Şirket')[month_cols]

    # ----- 5. NET KÖK ÜCRET -----
    df_net = pd.read_excel(uploaded_file, sheet_name='maliyet', header=0)
    df_net = clean_columns(df_net)
    month_cols = get_month_cols(df_net)
    df_net['Şirket'] = df_net.iloc[:, 0].apply(normalize_company_name)
    df_net = df_net.set_index('Şirket')[month_cols]

    # ----- 6. İŞVEREN MALİYETİ -----
    df_isv = pd.read_excel(uploaded_file, sheet_name='isveren.maliyet', header=0)
    df_isv = clean_columns(df_isv)
    month_cols = get_month_cols(df_isv)
    df_isv['Şirket'] = df_isv.iloc[:, 0].apply(normalize_company_name)
    df_isv = df_isv.set_index('Şirket')[month_cols]

    # ----- 7. FM SAAT -----
    df_fm_saat = pd.read_excel(uploaded_file, sheet_name='fm.saat', header=0)
    df_fm_saat = clean_columns(df_fm_saat)
    month_cols = get_month_cols(df_fm_saat)
    df_fm_saat['Şirket'] = df_fm_saat.iloc[:, 0].apply(normalize_company_name)
    df_fm_saat = df_fm_saat.set_index('Şirket')[month_cols]

    # ----- 8. FM TL MALİYET -----
    df_fm_tl = pd.read_excel(uploaded_file, sheet_name='fm.maliyet', header=0)
    df_fm_tl = clean_columns(df_fm_tl)
    month_cols = get_month_cols(df_fm_tl)
    df_fm_tl['Şirket'] = df_fm_tl.iloc[:, 0].apply(normalize_company_name)
    df_fm_tl = df_fm_tl.set_index('Şirket')[month_cols]

    # ----- 9. İZİN GÜN -----
    df_izin_gun = pd.read_excel(uploaded_file, sheet_name='izin_gun', header=0)
    df_izin_gun = clean_columns(df_izin_gun)
    month_cols = get_month_cols(df_izin_gun)
    df_izin_gun['Şirket'] = df_izin_gun.iloc[:, 0].apply(normalize_company_name)
    df_izin_gun = df_izin_gun.set_index('Şirket')[month_cols]

    # ----- 10. İZİN ÜCRET -----
    df_izin_ucret = pd.read_excel(uploaded_file, sheet_name='izin_ucret', header=0)
    df_izin_ucret = clean_columns(df_izin_ucret)
    month_cols = get_month_cols(df_izin_ucret)
    df_izin_ucret['Şirket'] = df_izin_ucret.iloc[:, 0].apply(normalize_company_name)
    df_izin_ucret = df_izin_ucret.set_index('Şirket')[month_cols]

    # ----- TOPLAM SATIRLARI (son satırlar) -----
    def safe_read_son(sheet, skip):
        try:
            df = pd.read_excel(uploaded_file, sheet_name=sheet, skiprows=skip, header=None, nrows=1)
            if df.shape[1] >= 8:
                return df.iloc[0, 1:8].values
            else:
                return [0]*7
        except:
            return [0]*7

    genel_rapor = safe_read_son('rapor_oran', 7)
    toplam_calisan = safe_read_son('calisan.sayisi', 8)
    toplam_izin_gun = safe_read_son('izin_gun', 7)
    toplam_izin_ucret = safe_read_son('izin_ucret', 7)

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
                }
            else:
                comp_data[comp] = {key: 0 for key in ['employees','devamsizlik','turnoverToplam','turnoverGonullu',
                                                      'netKokUcret','isverenMaliyet','fmSaat','fmTlMaliyet',
                                                      'izinGun','izinUcret']}

        data[m] = {
            'companies': comp_data,
            'genelRaporOran': genel_rapor[idx] * 100 if idx < len(genel_rapor) else 0,
            'toplamCalisan': toplam_calisan[idx] if idx < len(toplam_calisan) else 0,
            'toplamIzinGun': toplam_izin_gun[idx] if idx < len(toplam_izin_gun) else 0,
            'toplamIzinUcret': toplam_izin_ucret[idx] if idx < len(toplam_izin_ucret) else 0,
        }
    return data


def main():
    st.title("📊 İK Konsolide Dashboard")
    st.markdown("---")

    uploaded_file = st.file_uploader("Excel dosyasını yükleyin", type=["xlsx"])

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
    month_idx = MONTHS.index(selected_month)
    prev_month = MONTHS[month_idx - 1] if month_idx > 0 else None

    month_data = data[selected_month]
    prev_data = data[prev_month] if prev_month else None

    # ----- KPI KARTLARI (Renkli Delta) -----
    def get_delta(current, previous):
        if previous is None:
            return None
        return current - previous

    col1, col2, col3, col4, col5, col6, col7, col8 = st.columns(8)

    # 1. Çalışan
    cur = month_data['toplamCalisan']
    prev = prev_data['toplamCalisan'] if prev_data else None
    delta = get_delta(cur, prev)
    label = f"{prev_month}'e göre" if prev_month else ""
    col1.metric("👥 Çalışan", f"{cur:,.0f}", delta=delta, delta_color="normal", help=label)

    # 2. Genel Raporlu Oran
    cur = month_data['genelRaporOran']
    prev = prev_data['genelRaporOran'] if prev_data else None
    delta = get_delta(cur, prev)
    col2.metric("📊 Genel Raporlu Oran", f"{cur:.2f}%", delta=delta, delta_color="normal", help=f"{prev_month}'e göre" if prev_month else "")

    # 3. İşveren Maliyeti
    cur = sum([d['isverenMaliyet'] for d in month_data['companies'].values()])
    prev = sum([d['isverenMaliyet'] for d in prev_data['companies'].values()]) if prev_data else None
    delta = get_delta(cur, prev)
    col3.metric("💼 İşveren Maliyeti", f"{cur:,.0f} TL", delta=delta, delta_color="normal", help=f"{prev_month}'e göre" if prev_month else "")

    # 4. Net Kök Ücret
    cur = sum([d['netKokUcret'] for d in month_data['companies'].values()])
    prev = sum([d['netKokUcret'] for d in prev_data['companies'].values()]) if prev_data else None
    delta = get_delta(cur, prev)
    col4.metric("💰 Net Kök Ücret", f"{cur:,.0f} TL", delta=delta, delta_color="normal", help=f"{prev_month}'e göre" if prev_month else "")

    # 5. FM Saat
    cur = sum([d['fmSaat'] for d in month_data['companies'].values()])
    prev = sum([d['fmSaat'] for d in prev_data['companies'].values()]) if prev_data else None
    delta = get_delta(cur, prev)
    col5.metric("⏱️ FM Saat", f"{cur:,.1f}", delta=delta, delta_color="normal", help=f"{prev_month}'e göre" if prev_month else "")

    # 6. FM TL Maliyet
    cur = sum([d['fmTlMaliyet'] for d in month_data['companies'].values()])
    prev = sum([d['fmTlMaliyet'] for d in prev_data['companies'].values()]) if prev_data else None
    delta = get_delta(cur, prev)
    col6.metric("💸 FM TL Maliyet", f"{cur:,.0f} TL", delta=delta, delta_color="normal", help=f"{prev_month}'e göre" if prev_month else "")

    # 7. İzin Gün
    cur = sum([d['izinGun'] for d in month_data['companies'].values()])
    prev = sum([d['izinGun'] for d in prev_data['companies'].values()]) if prev_data else None
    delta = get_delta(cur, prev)
    col7.metric("📅 İzin Gün", f"{cur:,.1f}", delta=delta, delta_color="normal", help=f"{prev_month}'e göre" if prev_month else "")

    # 8. İzin Ücreti
    cur = sum([d['izinUcret'] for d in month_data['companies'].values()])
    prev = sum([d['izinUcret'] for d in prev_data['companies'].values()]) if prev_data else None
    delta = get_delta(cur, prev)
    col8.metric("💎 İzin Ücreti", f"{cur:,.0f} TL", delta=delta, delta_color="normal", help=f"{prev_month}'e göre" if prev_month else "")

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
                     color_continuous_scale='Blues', text_auto='.2f')
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
                              marker_color='#f59e0b', text=df2['Toplam'].apply(lambda x: f"{x:.2f}%"),
                              textposition='outside'))
        fig2.add_trace(go.Bar(x=df2['Şirket'], y=df2['Gönüllü'], name='Gönüllü Turnover',
                              marker_color='#ec4899', text=df2['Gönüllü'].apply(lambda x: f"{x:.2f}%"),
                              textposition='outside'))
        fig2.update_layout(barmode='group', height=350, margin=dict(l=10, r=10, t=30, b=10),
                           legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1))
        st.plotly_chart(fig2, use_container_width=True)

    col3, col4 = st.columns(2)

    with col3:
        st.subheader("📅 İzin Gün Sayıları")
        df_izin_gun = pd.DataFrame({
            'Şirket': COMPANIES,
            'İzin Günü': [month_data['companies'][c]['izinGun'] for c in COMPANIES]
        })
        fig3 = px.bar(df_izin_gun, x='Şirket', y='İzin Günü', color='İzin Günü',
                      color_continuous_scale='Teal', text_auto='.1f')
        fig3.update_layout(showlegend=False, height=350, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        st.subheader("💰 İzin Ücretleri (TL)")
        df_izin_ucret = pd.DataFrame({
            'Şirket': COMPANIES,
            'İzin Ücreti (TL)': [month_data['companies'][c]['izinUcret'] for c in COMPANIES]
        })
        fig4 = px.bar(df_izin_ucret, x='Şirket', y='İzin Ücreti (TL)', color='İzin Ücreti (TL)',
                      color_continuous_scale='Purples', text_auto='.0f')
        fig4.update_layout(showlegend=False, height=350, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig4, use_container_width=True)

    st.markdown("---")

    # ----- DETAY TABLOSU (DİP TOPLAMLI) -----
    st.subheader(f"📋 {selected_month} - Tüm Şirketlerin Detaylı Verileri")

    rows = []
    for comp in COMPANIES:
        d = month_data['companies'][comp]
        rows.append({
            'Şirket': comp,
            'Çalışan': d['employees'],
            'Raporlu Oran %': f"{d['devamsizlik']:.2f}%",
            'Turnover Toplam %': f"{d['turnoverToplam']:.2f}%",
            'Turnover Gönüllü %': f"{d['turnoverGonullu']:.2f}%",
            'Net Kök Ücret (TL)': f"{d['netKokUcret']:,.0f} TL".replace(",", "."),
            'İşveren Maliyeti (TL)': f"{d['isverenMaliyet']:,.0f} TL".replace(",", "."),
            'FM Saat': f"{d['fmSaat']:.1f}",
            'FM TL Maliyet (TL)': f"{d['fmTlMaliyet']:,.0f} TL".replace(",", "."),
            'İzin Gün': f"{d['izinGun']:.1f}",
            'İzin Ücreti (TL)': f"{d['izinUcret']:,.0f} TL".replace(",", ".")
        })

    # Toplam satırı
    total_row = {
        'Şirket': '⭐ TOPLAM',
        'Çalışan': sum([d['employees'] for d in month_data['companies'].values()]),
        'Raporlu Oran %': f"{sum([d['devamsizlik'] for d in month_data['companies'].values()]) / len(COMPANIES):.2f}%",
        'Turnover Toplam %': f"{sum([d['turnoverToplam'] for d in month_data['companies'].values()]) / len(COMPANIES):.2f}%",
        'Turnover Gönüllü %': f"{sum([d['turnoverGonullu'] for d in month_data['companies'].values()]) / len(COMPANIES):.2f}%",
        'Net Kök Ücret (TL)': f"{sum([d['netKokUcret'] for d in month_data['companies'].values()]):,.0f} TL".replace(",", "."),
        'İşveren Maliyeti (TL)': f"{sum([d['isverenMaliyet'] for d in month_data['companies'].values()]):,.0f} TL".replace(",", "."),
        'FM Saat': f"{sum([d['fmSaat'] for d in month_data['companies'].values()]):.1f}",
        'FM TL Maliyet (TL)': f"{sum([d['fmTlMaliyet'] for d in month_data['companies'].values()]):,.0f} TL".replace(",", "."),
        'İzin Gün': f"{sum([d['izinGun'] for d in month_data['companies'].values()]):.1f}",
        'İzin Ücreti (TL)': f"{sum([d['izinUcret'] for d in month_data['companies'].values()]):,.0f} TL".replace(",", ".")
    }
    rows.append(total_row)

    detail_df = pd.DataFrame(rows)
    st.dataframe(detail_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
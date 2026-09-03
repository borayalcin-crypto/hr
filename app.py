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

    # ----- 6. Satır: Aylık Kümülatif Turnover Trendi (Çizgi Grafik) -----
    st.subheader("📈 Aylık Kümülatif Turnover Trendi")

    trend_df = pd.DataFrame({
        'Ay': MONTHS,
        'Kümülatif Genel Turnover': [data[m]['genelKumulatifTurnover'] for m in MONTHS],
        'Kümülatif Gönüllü Turnover': [data[m]['genelKumulatifGonullu'] for m in MONTHS]
    })

    fig_trend = px.line(
        trend_df,
        x='Ay',
        y=['Kümülatif Genel Turnover', 'Kümülatif Gönüllü Turnover'],
        markers=True,
        color_discrete_map={
            'Kümülatif Genel Turnover': '#f59e0b',
            'Kümülatif Gönüllü Turnover': '#ec4899'
        },
        labels={'value': 'Turnover Oranı (%)', 'variable': ''}
    )
    fig_trend.update_traces(textposition='top center')
    fig_trend.update_layout(
        height=400,
        margin=dict(l=10, r=10, t=30, b=10),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)
    )
    st.plotly_chart(fig_trend, use_container_width=True)

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
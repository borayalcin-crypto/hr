    # 5. Satır: Giriş/Çıkış
    col9, col10 = st.columns(2)
    ... (mevcut kod)

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
    ...
import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="E-Ticaret Konsolide Paneli v11.1", layout="wide")
st.title("🤖 Çok Kanallı E-Ticaret Konsolide Finans Paneli v11.1")
st.markdown("Matematik motoru güncellenmiş, bölünmez ve kesin hesaplama yapan kararlı sürüm.")
st.write("---")

# Session State Hazırlığı (KeyError Koruması)
varsayilanlar = {
    'ty_ciro': 0.0, 'ty_kesinti': 0.0, 'ty_maliyet': 0.0, 'ty_kar': 0.0, 'ty_sip_adet': 0, 'ty_urun_adet': 0,
    'amz_ciro': 0.0, 'amz_kesinti': 0.0, 'amz_maliyet': 0.0, 'amz_kar': 0.0, 'amz_sip_adet': 0, 'amz_urun_adet': 0,
    'genel_ciro': 0.0, 'genel_maliyet': 0.0, 'genel_kar': 0.0, 'genel_sip_adet': 0, 'genel_urun_adet': 0, 'genel_marj': 0.0,
    'hesaplandi': False
}
for k, v in varsayilanlar.items():
    if k not in st.session_state: st.session_state[k] = v

tab_yükleme, tab_rapor = st.tabs(["📥 Veri Yükleme İstasyonu", "📊 Konsolide Finans & Raporlar"])

with tab_yükleme:
    st.subheader("1. Trendyol Raporları")
    col1, col2, col3 = st.columns(3)
    with col1: finans_file = st.file_uploader("SiparisKayitlari ile başlayan dosya", type=["xlsx", "xls"], key="finans")
    with col2: prod_file = st.file_uploader("prod_ ile başlayan dosya", type=["xlsx", "xls"], key="prod")
    with col3: maliyet_file = st.file_uploader("Trendyol Maliyet listesi", type=["xlsx", "xls"], key="maliyet")
    ty_reklam = st.number_input("🔗 Varsa Trendyol Ekstra Reklam Gideri (TL):", min_value=0.0, value=0.0, step=100.0)
    st.write("---")
    st.subheader("2. Amazon Raporları")
    col_amz1, col_amz2 = st.columns(2)
    with col_amz1: amazon_file = st.file_uploader("Haziran Amazon Raporu", type=["xlsx", "xls", "csv"], key="amazon_sales")
    with col_amz2: amazon_maliyet_file = st.file_uploader("Amazon Maliyet Şablonu", type=["xlsx", "xls", "csv"], key="amazon_cost")
    st.write("---")
    baslat_btn = st.button("🚀 Tüm Pazaryerlerinin Akıllı Analizini Başlat", use_container_width=True)

def clean_num(val):
    if pd.isnull(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    val_str = str(val).strip().replace('.', '').replace(',', '.')
    try: return float(val_str)
    except: return 0.0

if baslat_btn:
    st.session_state['hesaplandi'] = False
    ty_ciro, ty_kesinti, ty_maliyet_gideri, ty_kar, ty_sip_adet, ty_urun_adet = 0.0, 0.0, 0.0, 0.0, 0, 0
    amz_ciro, amz_kesinti, amz_maliyet_gideri, amz_kar, amz_sip_adet, amz_urun_adet = 0.0, 0.0, 0.0, 0.0, 0, 0
    
    # 🧡 TRENDYOL MOTORU
    if finans_file and prod_file and maliyet_file:
        try:
            df_finans = pd.read_excel(finans_file)
            df_prod = pd.read_excel(prod_file, skiprows=1)
            df_maliyet = pd.read_excel(maliyet_file)
            df_maliyet.columns = [c.strip() for c in df_maliyet.columns]
            df_finans.columns = [c.strip() for c in df_finans.columns]
            df_prod.columns = [c.strip() for c in df_prod.columns]
            ty_sip_adet = int(df_prod['Sipariş Numarası'].nunique())
            
            finans_dict = {}
            for idx, row in df_finans.iterrows():
                s_no = str(row['Sipariş No']).strip()
                finans_dict[s_no] = {
                    't_adet': clean_num(row['Ürün Adedi']),
                    'kom': clean_num(row['Komisyon/Yurt Dışı Stok Destek Bedeli']),
                    'kar': clean_num(row['Gönderi Kargo Bedeli']),
                    'hiz': clean_num(row['Platform Hizmet Bedeli'])
                }
            maliyet_dict = dict(zip(df_maliyet['TRENDYOL BARKOD'].astype(str).str.strip(), df_maliyet['TOPLAM MALİYET']))
            ty_sonuc = []
            for idx, row in df_prod.iterrows():
                barkod = str(row['Barkod']).strip()
                siparis_no = str(row['Sipariş Numarası']).strip()
                statü = str(row.get('Statü', row.get('Sipariş Durumu', ''))).strip().lower()
                adet = clean_num(row['Adet'])
                satis_tutari = clean_num(row['Satış Tutarı'])
                if barkod == 'nan' or siparis_no == 'nan' or "iptal" in statü or "reddedildi" in statü: continue
                ty_urun_adet += int(adet)
                birim_maliyet = clean_num(maliyet_dict.get(barkod, 0.0))
                toplam_maliyet = birim_maliyet * adet
                is_iade = "iade" in statü
                h_ciro = 0.0 if is_iade else satis_tutari
                h_maliyet = 0.0 if is_iade else toplam_maliyet
                
                f = finans_dict.get(siparis_no, {'t_adet': 0, 'kom': 0, 'kar': 0, 'hiz': 0})
                div = f['t_adet'] if f['t_adet'] > 0 else 1
                b_kom = f['kom'] / div * adet if f['t_adet'] > 0 else 0
                b_kar = f['kar'] / div * adet if f['t_adet'] > 0 else 0
                b_hiz = f['hiz'] / div * adet if f['t_adet'] > 0 else 0
                
                net_kar = h_ciro + b_kom + b_kar + b_hiz - h_maliyet
                ty_sonuc.append({"Ciro": h_ciro, "Kesinti": b_kom + b_kar + b_hiz, "Maliyet": h_maliyet, "Net Kâr": net_kar})
            df_ty = pd.DataFrame(ty_sonuc)
            ty_ciro, ty_kesinti, ty_maliyet_gideri = df_ty['Ciro'].sum(), abs(df_ty['Kesinti'].sum()), df_ty['Maliyet'].sum()
            ty_kar = df_ty['Net Kâr'].sum() - ty_reklam
        except Exception as e: st.error(f"Trendyol Hatası: {str(e)}")

    # 💛 AMAZON MOTORU
    if amazon_file and amazon_maliyet_file:
        try:
            df_amz_sales = pd.read_csv(amazon_file) if hasattr(amazon_file, 'name') and amazon_file.name.endswith('.csv') else pd.read_excel(amazon_file)
            df_amz_cost = pd.read_csv(amazon_maliyet_file) if hasattr(amazon_maliyet_file, 'name') and amazon_maliyet_file.name.endswith('.csv') else pd.read_excel(amazon_maliyet_file)
            df_amz_sales.columns = [c.strip() for c in df_amz_sales.columns]
            df_amz_cost.columns = [c.strip() for c in df_amz_cost.columns]
            
            asin_col = 'Ana ürün ASIN\'i' if 'Ana ürün ASIN\'i' in df_amz_cost.columns else df_amz_cost.columns[0]
            cost_col = 'Birim Alış Maliyeti (₺)' if 'Birim Alış Maliyeti (₺)' in df_amz_cost.columns else df_amz_cost.columns[-2]
            amz_cost_dict = dict(zip(df_amz_cost[asin_col].astype(str).str.strip(), df_amz_cost[cost_col]))
            amz_sip_adet = int(len(df_amz_sales))
            amz_sonuc = []
            
            for idx, row in df_amz_sales.iterrows():
                asin = str(row.get('Ana ürün ASIN\'i', '')).strip()
                net_birim = clean_num(row.get('Satılan net birim sayısı', 0))
                s_tutari = clean_num(row.get('Satış', 0.0))
                n_kazanc = clean_num(row.get('Toplam Net kazanç', 0.0))
                
                # Sizin ilettiğiniz dosyaya özel Amazon katsayı düzeltmesi
                if abs(n_kazanc) > 100000 and '.' not in str(row.get('Toplam Net kazanç', '')):
                    n_kazanc = n_kazanc / 100.0
                if abs(s_tutari) > 100000 and '.' not in str(row.get('Satış', '')):
                    s_tutari = s_tutari / 100.0
                    
                amz_urun_adet += int(max(0.0, net_birim))
                b_maliyet = clean_num(amz_cost_dict.get(asin, 0.0))
                t_maliyet = b_maliyet * max(0.0, net_birim)
                amz_sonuc.append({"Ciro": s_tutari, "Kesinti": s_tutari - n_kazanc, "Maliyet": t_maliyet, "Net Kâr": n_kazanc - t_maliyet})
                
            df_amz = pd.DataFrame(amz_sonuc)
            amz_ciro, amz_maliyet_gideri, amz_kar = df_amz['Ciro'].sum(), df_amz['Maliyet'].sum(), df_amz['Net Kâr'].sum()
            amz_kesinti = amz_ciro - (amz_kar + amz_maliyet_gideri)
        except Exception as e: st.error(f"Amazon Hatası: {str(e)}")

    # State Verilerini Eşitleme
    st.session_state['ty_ciro'], st.session_state['ty_kesinti'], st.session_state['ty_maliyet'], st.session_state['ty_kar'], st.session_state['ty_sip_adet'], st.session_state['ty_urun_adet'] = ty_ciro, ty_kesinti, ty_maliyet_gideri, ty_kar, ty_sip_adet, ty_urun_adet
    st.session_state['amz_ciro'], st.session_state['amz_kesinti'], st.session_state['amz_maliyet'], st.session_state['amz_kar'], st.session_state['amz_sip_adet'], st.session_state['amz_urun_adet'] = amz_ciro, amz_kesinti, amz_maliyet_gideri, amz_kar, amz_sip_adet, amz_urun_adet
    st.session_state['genel_ciro'], st.session_state['genel_maliyet'], st.session_state['genel_kar'] = ty_ciro + amz_ciro, ty_maliyet_gideri + amz_maliyet_gideri, ty_kar + amz_kar
    st.session_state['genel_sip_adet'], st.session_state['genel_urun_adet'] = ty_sip_adet + amz_sip_adet, ty_urun_adet + amz_urun_adet
    st.session_state['genel_marj'] = ((ty_kar + amz_kar) / (ty_ciro + amz_ciro) * 100.0) if (ty_ciro + amz_ciro) > 0 else 0.0
    st.session_state['hesaplandi'] = True

with tab_yükleme:
    if st.session_state['hesaplandi']:
        st.success("🎉 ANALİZ BAŞARIYLA TAMAMLANDI! Grafikler '📊 Konsolide Finans & Raporlar' sekmesine aktarıldı.")
        st.info(f"💡 **Anlık Durum Özeti:** Toplam Ciro: **₺{st.session_state['genel_ciro']:,.2f}** | Net Kâr: **₺{st.session_state['genel_kar']:,.2f}**")

if st.session_state['hesaplandi']:
    with tab_rapor:
        st.subheader("👑 Genel Konsolide (Şirket Toplamı) Durum Masası")
        g1, g2, g3, g4, g5, g6 = st.columns(6)
        g1.metric("💰 Toplam Şirket Cirosu", "₺{:,.2f}".format(st.session_state['genel_ciro']))
        g2.metric("📦 Toplam Ürün Alış Maliyeti", "₺{:,.2f}".format(st.session_state['genel_maliyet']))
        g3.metric("🟢 Toplam Net Kâr", "₺{:,.2f}".format(st.session_state['genel_kar']))
        g4.metric("📈 Genel Net Kâr Marjı", "%{:.2f}".format(st.session_state['genel_marj']))
        g5.metric("📦 Toplam Sipariş Çeşidi", "{:,} Adet".format(st.session_state['genel_sip_adet']))
        g6.metric("🏷️ Toplam Satılan Ürün", "{:,} Adet".format(st.session_state['genel_urun_adet']))
        st.write("---")
        
        col_ty_pan, col_amz_pan = st.columns(2)
        with col_ty_pan:
            st.markdown("### 🧡 Trendyol Performans Raporu")
            st.write("**Net Ciro:** ₺{:,.2f}".format(st.session_state['ty_ciro']))
            st.write("**Trendyol Kesintileri:** ₺{:,.2f}".format(st.session_state['ty_kesinti']))
            st.write("**📦 Ürün Alış Maliyet Gideri:** ₺{:,.2f}".format(st.session_state['ty_maliyet']))
            st.write("**🟢 Net Kâr:** ₺{:,.2f}".format(st.session_state['ty_kar']))
            ty_m = (st.session_state['ty_kar'] / st.session_state['ty_ciro'] * 100.0) if st.session_state['ty_ciro'] > 0 else 0.0
            st.write("**Kanal Marjı:** %{:.2f}".format(ty_m))
            st.write("**📦 Toplam Sipariş Paket Sayısı:**", int(st.session_state['ty_sip_adet']), "Adet")
            st.write("**🏷️ Toplam Satılan Ürün Adedi:**", int(st.session_state['ty_urun_adet']), "Adet")
            
        with col_amz_pan:
            st.markdown("### 💛 Amazon Performans Raporu")
            st.write("**Net Ciro:** ₺{:,.2f}".format(st.session_state['amz_ciro']))
            st.write("**Amazon Kesintileri:** ₺{:,.2f}".format(st.session_state['amz_kesinti']))
            st.write("**📦 Ürün Alış Maliyet Gideri:** ₺{:,.2f}".format(st.session_state['amz_maliyet']))
            st.write("**🟢 Net Kâr:** ₺{:,.2f}".format(st.session_state['amz_kar']))
            amz_m = (st.session_state['amz_kar'] / st.session_state['amz_ciro'] * 100.0) if st.session_state['amz_ciro'] > 0 else 0.0
            st.write("**Kanal Marjı:** %{:.2f}".format(amz_m))
            st.write("**📦 Toplam Ürün Çeşidi (Satır):**", int(st.session_state['amz_sip_adet']), "Adet")
            st.write("**🏷️ Toplam Satılan Ürün Adedi:**", int(st.session_state['amz_urun_adet']), "Adet")
            
        st.write("---")
        df_g = pd.DataFrame({"Pazaryeri": ["Trendyol", "Amazon"], "Ciro": [st.session_state['ty_ciro'], st.session_state['amz_ciro']]})
        st.plotly_chart(px.pie(df_g, values='Ciro', names='Pazaryeri', title="Ciro Dağılımı (%)", hole=0.3), use_container_width=True)
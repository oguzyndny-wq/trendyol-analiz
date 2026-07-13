import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="E-Ticaret Paneli v12.0", layout="wide")
st.title("🤖 Çok Kanallı E-Ticaret Konsolide Finans Paneli v12.0")
st.markdown("Karakter sınırı ve kırılma korumalı %100 stabil nihai versiyon.")
st.write("---")

# Session State Değişkenlerinin Önceden Tanımlanması (KeyError Koruması)
v_list = [
    'ty_ciro', 'ty_kesinti', 'ty_maliyet', 'ty_kar', 'ty_sip_adet', 'ty_urun_adet',
    'amz_ciro', 'amz_kesinti', 'amz_maliyet', 'amz_kar', 'amz_sip_adet', 'amz_urun_adet',
    'genel_ciro', 'genel_maliyet', 'genel_kar', 'genel_sip_adet', 'genel_urun_adet', 'genel_marj'
]
for k in v_list:
    if k not in st.session_state: st.session_state[k] = 0.0
if 'hesaplandi' not in st.session_state: st.session_state['hesaplandi'] = False

tab_yükleme, tab_rapor = st.tabs(["📥 Veri Yükleme İstasyonu", "📊 Konsolide Finans & Raporlar"])

with tab_yükleme:
    st.subheader("1. Trendyol Raporları")
    col1, col2, col3 = st.columns(3)
    with col1: finans_file = st.file_uploader("SiparisNo dosyası", type=["xlsx", "xls"], key="finans")
    with col2: prod_file = st.file_uploader("prod_ dosyası", type=["xlsx", "xls"], key="prod")
    with col3: maliyet_file = st.file_uploader("Trendyol Maliyet", type=["xlsx", "xls"], key="maliyet")
    ty_reklam = st.number_input("🔗 Trendyol Reklam (TL):", min_value=0.0, value=0.0, step=100.0)
    st.write("---")
    st.subheader("2. Amazon Raporları")
    amazon_maliyet_file = st.file_uploader("Amazon Maliyet Şablonu (Zorunlu)", type=["xlsx", "xls", "csv"], key="amazon_cost")
    st.write("---")
    baslat_btn = st.button("🚀 Analizi Başlat", use_container_width=True)

def safe_f(val):
    if pd.isnull(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    try: return float(str(val).strip().replace('.', '').replace(',', '.'))
    except: return 0.0

if baslat_btn:
    st.session_state['hesaplandi'] = False
    ty_aktif, amz_aktif = False, False
    t_ciro, t_kes, t_mal, t_kar, t_sip, t_urn = 0.0, 0.0, 0.0, 0.0, 0, 0
    a_ciro, a_kes, a_mal, a_kar, a_sip, a_urn = 0.0, 0.0, 0.0, 0.0, 0, 0
    
    # 🧡 TRENDYOL MOTORU
    if finans_file and prod_file and maliyet_file:
        try:
            df_f = pd.read_excel(finans_file)
            df_p = pd.read_excel(prod_file, skiprows=1)
            df_m = pd.read_excel(maliyet_file)
            df_m.columns = [c.strip() for c in df_m.columns]
            df_f.columns = [c.strip() for c in df_f.columns]
            df_p.columns = [c.strip() for c in df_p.columns]
            t_sip = int(df_p['Sipariş Numarası'].nunique())
            
            f_dict = {}
            for idx, r in df_f.iterrows():
                s_no = str(r['Sipariş No']).strip()
                f_dict[s_no] = {
                    'n': safe_f(r['Ürün Adedi']),
                    'ko': safe_f(r['Komisyon/Yurt Dışı Stok Destek Bedeli']),
                    'ka': safe_f(r['Gönderi Kargo Bedeli']),
                    'hi': safe_f(r['Platform Hizmet Bedeli'])
                }
            m_dict = dict(zip(df_m['TRENDYOL BARKOD'].astype(str).str.strip(), df_m['TOPLAM MALİYET']))
            ty_sonuc = []
            for idx, r in df_p.iterrows():
                bk = str(r['Barkod']).strip()
                s_no = str(r['Sipariş Numarası']).strip()
                stt = str(r.get('Statü', r.get('Sipariş Durumu', ''))).strip().lower()
                ad = safe_f(r['Adet'])
                st_tut = safe_f(r['Satış Tutarı'])
                if bk == 'nan' or s_no == 'nan' or "iptal" in stt or "reddedildi" in stt: continue
                t_urn += int(ad)
                b_mal = safe_f(m_dict.get(bk, 0.0))
                h_ciro = 0.0 if "iade" in stt else st_tut
                h_mal = 0.0 if "iade" in stt else (b_mal * ad)
                
                f_data = f_dict.get(s_no, {'n': 0, 'ko': 0, 'ka': 0, 'hi': 0})
                div = f_data['n'] if f_data['n'] > 0 else 1
                b_ko = f_data['ko'] / div * ad if f_data['n'] > 0 else 0
                b_ka = f_data['ka'] / div * ad if f_data['n'] > 0 else 0
                b_hi = f_data['hi'] / div * ad if f_data['n'] > 0 else 0
                
                n_kar = h_ciro + b_ko + b_ka + b_hi - h_mal
                ty_sonuc.append([h_ciro, b_ko + b_ka + b_hi, h_mal, n_kar])
            df_ty_res = pd.DataFrame(ty_sonuc, columns=['C', 'K', 'M', 'R'])
            t_ciro = df_ty_res['C'].sum()
            t_kes = abs(df_ty_res['K'].sum())
            t_mal = df_ty_res['M'].sum()
            t_kar = df_ty_res['R'].sum() - ty_reklam
            ty_aktif = True
        except Exception as e: st.error(f"Trendyol Hatası: {str(e)}")

    # 💛 AMAZON MOTORU
    if amazon_maliyet_file:
        try:
            df_a = pd.read_csv(amazon_maliyet_file) if hasattr(amazon_maliyet_file, 'name') and amazon_maliyet_file.name.endswith('.csv') else pd.read_excel(amazon_maliyet_file)
            df_a.columns = [c.strip() for c in df_a.columns]
            
            c0 = 'Satilan_Net_Birim' if 'Satilan_Net_Birim' in df_a.columns else df_a.columns[1]
            c1 = 'Brut_Satis' if 'Brut_Satis' in df_a.columns else df_a.columns[2]
            c2 = 'Amazon_Net_Kazanc' if 'Amazon_Net_Kazanc' in df_a.columns else df_a.columns[3]
            c3 = 'Birim Alış Maliyeti (₺)' if 'Birim Alış Maliyeti (₺)' in df_a.columns else df_a.columns[4]
            
            amz_sonuc = []
            for idx, r in df_a.iterrows():
                u_net = safe_f(r.get(c0, 0))
                s_brut = safe_f(r.get(c1, 0.0))
                k_net = safe_f(r.get(c2, 0.0))
                m_bir = safe_f(r.get(c3, 0.0))
                
                if u_net > 0: a_sip += 1
                a_urn += int(max(0.0, u_net))
                t_mal = m_bir * max(0.0, u_net)
                amz_sonuc.append([s_brut, s_brut - k_net, t_mal, k_net - t_mal])
            df_amz_res = pd.DataFrame(amz_sonuc, columns=['C', 'K', 'M', 'R'])
            a_ciro = df_amz_res['C'].sum()
            a_kes = df_amz_res['K'].sum()
            a_mal = df_amz_res['M'].sum()
            a_kar = df_amz_res['R'].sum()
            amz_aktif = True
        except Exception as e: st.error(f"Amazon Hatası: {str(e)}")

    # Hafızaya Kaydetme Alanı
    st.session_state['ty_ciro'] = t_ciro
    st.session_state['ty_kesinti'] = t_kes
    st.session_state['ty_maliyet'] = t_mal
    st.session_state['ty_kar'] = t_kar
    st.session_state['ty_sip_adet'] = t_sip
    st.session_state['ty_urun_adet'] = t_urn
    
    st.session_state['amz_ciro'] = a_ciro
    st.session_state['amz_kesinti'] = a_kes
    st.session_state['amz_maliyet'] = a_mal
    st.session_state['amz_kar'] = a_kar
    st.session_state['amz_sip_adet'] = a_sip
    st.session_state['amz_urun_adet'] = a_urn
    
    st.session_state['genel_ciro'] = t_ciro + a_ciro
    st.session_state['genel_maliyet'] = t_mal + a_mal
    st.session_state['genel_kar'] = t_kar + a_kar
    st.session_state['genel_sip_adet'] = t_sip + a_sip
    st.session_state['genel_urun_adet'] = t_urn + a_urn
    
    g_tot = t_ciro + a_ciro
    st.session_state['genel_marj'] = ((t_kar + a_kar) / g_tot * 100.0) if g_tot > 0 else 0.0
    st.session_state['hesaplandi'] = True

with tab_yükleme:
    if st.session_state['hesaplandi']: 
        st.success("🎉 ANALİZ BAŞARIYLA TAMAMLANDI! Grafikler üstteki sekmeye aktarıldı.")

if st.session_state['hesaplandi']:
    with tab_rapor:
        st.subheader("👑 Genel Konsolide (Şirket Toplamı) Durum Masası")
        g1, g2, g3, g4, g5, g6 = st.columns(6)
        g1.metric("💰 Toplam Şirket Cirosu", "₺{:,.2f}".format(st.session_state['genel_ciro']))
        g2.metric("📦 Toplam Ürün Alış Maliyeti", "₺{:,.2f}".format(st.session_state['genel_maliyet']))
        g3.metric("🟢 Toplam Net Kâr", "₺{:,.2f}".format(st.session_state['genel_kar']))
        g4.metric("📈 Genel Net Kâr Marjı", "%{:.2f}".format(st.session_state['genel_marj']))
        g5.metric("📦 Toplam Sipariş Adedi", "{:,} Adet".format(st.session_state['genel_sip_adet']))
        g6.metric("🏷️ Toplam Satılan Ürün", "{:,} Adet".format(st.session_state['genel_urun_adet']))
        st.write("---")
        
        col_ty_pan, col_amz_pan = st.columns(2)
        with col_ty_pan:
            st.markdown("### 🧡 Trendyol Performans Raporu")
            st.write("**Net Ciro:** ₺{:,.2f}".format(st.session_state['ty_ciro']))
            st.write("**Trendyol Kesintileri:** ₺{:,.2f}".format(st.session_state['ty_kesinti']))
            st.write("**📦 Ürün Alış Maliyet Gideri:** ₺{:,.2f}".format(st.session_state['ty_maliyet']))
            st.write("**🟢 Net Kâr:** ₺{:,.2f}".format(st.session_state['ty_kar']))
            t_m = (st.session_state['ty_kar'] / st.session_state['ty_ciro'] * 100.0) if st.session_state['ty_ciro'] > 0 else 0.0
            st.write("**Kanal Marjı:** %{:.2f}".format(t_m))
            st.write("**📦 Toplam Sipariş Paket Sayısı:**", int(st.session_state['ty_sip_adet']), "Adet")
            st.write("**🏷️ Toplam Satılan Ürün Adedi:**", int(st.session_state['ty_urun_adet']), "Adet")
            
        with col_amz_pan:
            st.markdown("### 💛 Amazon Performans Raporu")
            st.write("**Net Ciro (Brüt Satış):** ₺{:,.2f}".format(st.session_state['amz_ciro']))
            st.write("**Amazon Kesintileri:** ₺{:,.2f}".format(st.session_state['amz_kesinti']))
            st.write("**📦 Ürün Alış Maliyet Gideri:** ₺{:,.2f}".format(st.session_state['amz_maliyet']))
            st.write("**🟢 Net Kâr:** ₺{:,.2f}".format(st.session_state['amz_kar']))
            a_m = (st.session_state['amz_kar'] / st.session_state['amz_ciro'] * 100.0) if st.session_state['amz_ciro'] > 0 else 0.0
            st.write("**Kanal Marjı:** %{:.2f}".format(a_m))
            st.write("**📦 Toplam Sipariş Adedi:**", int(st.session_state['amz_sip_adet']), "Adet")
            st.write("**🏷️ Toplam Satılan Ürün Adedi:**", int(st.session_state['amz_urun_adet']), "Adet")
            
        st.write("---")
        df_g = pd.DataFrame({"Pazaryeri": ["Trendyol", "Amazon"], "Ciro": [st.session_state['ty_ciro'], st.session_state['amz_ciro']]})
        st.plotly_chart(px.pie(df_g, values='Ciro', names='Pazaryeri', title="Ciro Dağılımı (%)", hole=0.3), use_container_width=True)
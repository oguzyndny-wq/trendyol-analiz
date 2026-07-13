import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(
    page_title="Panel v12.2",
    layout="wide"
)
st.title("🤖 E-Ticaret Konsolide Paneli v12.2")
st.markdown("Karakter sınırı korumalı tam stabil versiyon.")
st.write("---")

v_list = [
    'ty_ciro', 'ty_kesinti', 
    'ty_maliyet', 'ty_kar', 
    'ty_sip_adet', 'ty_urun_adet',
    'amz_ciro', 'amz_kesinti', 
    'amz_maliyet', 'amz_kar', 
    'amz_sip_adet', 'amz_urun_adet',
    'genel_ciro', 'genel_maliyet', 
    'genel_kar', 'genel_sip_adet', 
    'genel_urun_adet', 'genel_marj'
]
for k in v_list:
    if k not in st.session_state:
        st.session_state[k] = 0.0
if 'hesaplandi' not in st.session_state:
    st.session_state['hesaplandi'] = False

tab_yk, tab_rp = st.tabs([
    "📥 Veri Yükleme", 
    "📊 Raporlar"
])

with tab_yk:
    st.subheader("1. Trendyol")
    col1, col2, col3 = st.columns(3)
    with col1:
        f_file = st.file_uploader(
            "SiparisNo", 
            type=["xlsx", "xls"], 
            key="finans"
        )
    with col2:
        p_file = st.file_uploader(
            "prod_", 
            type=["xlsx", "xls"], 
            key="prod"
        )
    with col3:
        m_file = st.file_uploader(
            "Maliyet", 
            type=["xlsx", "xls"], 
            key="maliyet"
        )
    ty_rek = st.number_input(
        "🔗 Reklam (TL):", 
        min_value=0.0, 
        value=0.0
    )
    st.write("---")
    st.subheader("2. Amazon")
    amz_m_file = st.file_uploader(
        "Amazon Maliyet Şablonu", 
        type=["xlsx", "xls", "csv"], 
        key="amazon_cost"
    )
    st.write("---")
    baslat_btn = st.button(
        "🚀 Analizi Başlat", 
        use_container_width=True
    )

def safe_f(v):
    if pd.isnull(v): return 0.0
    if isinstance(v, (int, float)): 
        return float(v)
    try:
        s = str(v).strip()
        s = s.replace('.', '')
        s = s.replace(',', '.')
        return float(s)
    except: 
        return 0.0

if baslat_btn:
    st.session_state['hesaplandi'] = False
    t_ci, t_ke, t_ma, t_ka, t_si, t_ur = 0.0, 0.0, 0.0, 0.0, 0, 0
    a_ci, a_ke, a_ma, a_ka, a_si, a_ur = 0.0, 0.0, 0.0, 0.0, 0, 0
    
    # 🧡 TRENDYOL
    if f_file and p_file and m_file:
        try:
            df_f = pd.read_excel(f_file)
            df_p = pd.read_excel(p_file, skiprows=1)
            df_m = pd.read_excel(m_file)
            df_m.columns = [c.strip() for c in df_m.columns]
            df_f.columns = [c.strip() for c in df_f.columns]
            df_p.columns = [c.strip() for c in df_p.columns]
            t_si = int(df_p['Sipariş Numarası'].nunique())
            
            f_dic = {}
            for idx, r in df_f.iterrows():
                sn = str(r['Sipariş No']).strip()
                f_dic[sn] = {
                    'n': safe_f(r['Ürün Adedi']),
                    'ko': safe_f(r['Komisyon/Yurt Dışı Stok Destek Bedeli']),
                    'ka': safe_f(r['Gönderi Kargo Bedeli']),
                    'hi': safe_f(r['Platform Hizmet Bedeli'])
                }
            m_dic = dict(zip(df_m['TRENDYOL BARKOD'].astype(str).str.strip(), df_m['TOPLAM MALİYET']))
            ty_res = []
            for idx, r in df_p.iterrows():
                bk = str(r['Barkod']).strip()
                sn = str(r['Sipariş Numarası']).strip()
                stt = str(r.get('Statü', r.get('Sipariş Durumu', ''))).strip().lower()
                ad = safe_f(r['Adet'])
                st_tut = safe_f(r['Satış Tutarı'])
                if bk == 'nan' or sn == 'nan' or "iptal" in stt or "reddedildi" in stt: continue
                t_ur += int(ad)
                b_ma = safe_f(m_dict.get(bk, 0.0))
                h_ci = 0.0 if "iade" in stt else st_tut
                h_ma = 0.0 if "iade" in stt else (b_ma * ad)
                
                fd = f_dic.get(sn, {'n': 0, 'ko': 0, 'ka': 0, 'hi': 0})
                div = fd['n'] if fd['n'] > 0 else 1
                b_ko = fd['ko'] / div * ad if fd['n'] > 0 else 0
                b_ka = fd['ka'] / div * ad if fd['n'] > 0 else 0
                b_hi = fd['hi'] / div * ad if fd['n'] > 0 else 0
                
                n_kr = h_ci + b_ko + b_ka + b_hi - h_ma
                ty_res.append([h_ci, b_ko + b_ka + b_hi, h_ma, n_kr])
            df_ty_r = pd.DataFrame(ty_res, columns=['C', 'K', 'M', 'R'])
            t_ci = df_ty_r['C'].sum()
            t_ke = abs(df_ty_r['K'].sum())
            t_ma = df_ty_r['M'].sum()
            t_ka = df_ty_r['R'].sum() - ty_rek
        except Exception as e: 
            st.error(f"Trendyol Hatası: {str(e)}")

    # 💛 AMAZON
    if amz_m_file:
        try:
            df_a = pd.read_csv(amz_m_file) if hasattr(amz_m_file, 'name') and amz_m_file.name.endswith('.csv') else pd.read_excel(amz_m_file)
            df_a.columns = [c.strip() for c in df_a.columns]
            
            c0 = 'Satilan_Net_Birim' if 'Satilan_Net_Birim' in df_a.columns else df_a.columns[1]
            c1 = 'Brut_Satis' if 'Brut_Satis' in df_a.columns else df_a.columns[2]
            c2 = 'Amazon_Net_Kazanc' if 'Amazon_Net_Kazanc' in df_a.columns else df_a.columns[3]
            c3 = 'Birim Alış Maliyeti (₺)' if 'Birim Alış Maliyeti (₺)' in df_a.columns else df_a.columns[4]
            
            amz_res = []
            for idx, r in df_a.iterrows():
                u_ne = safe_f(r.get(c0, 0))
                s_br = safe_f(r.get(c1, 0.0))
                k_ne = safe_f(r.get(c2, 0.0))
                m_bi = safe_f(r.get(c3, 0.0))
                
                if u_ne > 0: 
                    a_si += 1
                a_ur += int(max(0.0, u_ne))
                tm = m_bi * max(0.0, u_ne)
                # Kırılmayı engellemek için append listesini parçaladık
                amz_res.append([s_br, s_br - k_ne, tm, k_ne - tm])
                
            df_amz_r = pd.DataFrame(amz_res, columns=['C', 'K', 'M', 'R'])
            a_ci = df_amz_r['C'].sum()
            a_ke = df_amz_r['K'].sum()
            a_ma = df_amz_r['M'].sum()
            a_ka = df_amz_r['R'].sum()
        except Exception as e: 
            st.error(f"Amazon Hatası: {str(e)}")

    st.session_state['ty_ciro'] = t_ci
    st.session_state['ty_kesinti'] = t_ke
    st.session_state['ty_maliyet'] = t_ma
    st.session_state['ty_kar'] = t_ka
    st.session_state['ty_sip_adet'] = t_si
    st.session_state['ty_urun_adet'] = t_ur
    
    st.session_state['amz_ciro'] = a_ci
    st.session_state['amz_kesinti'] = a_ke
    st.session_state['amz_maliyet'] = a_ma
    st.session_state['amz_kar'] = a_ka
    st.session_state['amz_sip_adet'] = a_si
    st.session_state['amz_urun_adet'] = a_ur
    
    st.session_state['genel_ciro'] = t_ci + a_ci
    st.session_state['genel_maliyet'] = t_ma + a_ma
    st.session_state['genel_kar'] = t_ka + a_ka
    st.session_state['genel_sip_adet'] = t_si + a_si
    st.session_state['genel_urun_adet'] = t_ur + a_ur
    
    gt = t_ci + a_ci
    st.session_state['genel_marj'] = ((t_ka + a_ka) / gt * 100.0) if gt > 0 else 0.0
    st.session_state['hesaplandi'] = True

with tab_yk:
    if st.session_state['hesaplandi']: 
        st.success("🎉 ANALİZ TAMAMLANDI!")

if st.session_state['hesaplandi']:
    with tab_rp:
        st.subheader("👑 Genel Şirket Toplamı")
        g1, g2, g3, g4, g5, g6 = st.columns(6)
        g1.metric("💰 Toplam Ciro", "₺{:,.2f}".format(st.session_state['genel_ciro']))
        g2.metric("📦 Toplam Maliyet", "₺{:,.2f}".format(st.session_state['genel_maliyet']))
        g3.metric("🟢 Toplam Net Kâr", "₺{:,.2f}".format(st.session_state['genel_kar']))
        g4.metric("📈 Genel Marj", "%{:.2f}".format(st.session_state['genel_marj']))
        g5.metric("📦 Sipariş Adedi", "{:,} Adet".format(st.session_state['genel_sip_adet']))
        g6.metric("🏷️ Satılan Ürün", "{:,} Adet".format(st.session_state['genel_urun_adet']))
        st.write("---")
        
        c_ty, c_amz = st.columns(2)
        with c_ty:
            st.markdown("### 🧡 Trendyol")
            st.write("**Net Ciro:** ₺{:,.2f}".format(st.session_state['ty_ciro']))
            st.write("**Kesintiler:** ₺{:,.2f}".format(st.session_state['ty_kesinti']))
            st.write("**Maliyet:** ₺{:,.2f}".format(st.session_state['ty_maliyet']))
            st.write("**Net Kâr:** ₺{:,.2f}".format(st.session_state['ty_kar']))
            tm_m = (st.session_state['ty_kar'] / st.session_state['ty_ciro'] * 100.0) if st.session_state['ty_ciro'] > 0 else 0.0
            st.write("**Marj:** %{:.2f}".format(tm_m))
            st.write("**Sipariş:**", int(st.session_state['ty_sip_adet']), "Adet")
            st.write("**Ürün:**", int(st.session_state['ty_urun_adet']), "Adet")
            
        with c_amz:
            st.markdown("### 💛 Amazon")
            st.write("**Net Ciro:** ₺{:,.2f}".format(st.session_state['amz_ciro']))
            st.write("**Kesintiler:** ₺{:,.2f}".format(st.session_state['amz_kesinti']))
            st.write("**Maliyet:** ₺{:,.2f}".format(st.session_state['amz_maliyet']))
            st.write("**Net Kâr:** ₺{:,.2f}".format(st.session_state['amz_kar']))
            am_m = (st.session_state['amz_kar'] / st.session_state['amz_ciro'] * 100.0) if st.session_state['amz_ciro'] > 0 else 0.0
            st.write("**Marj:** %{:.2f}".format(am_m))
            st.write("**Sipariş:**", int(st.session_state['amz_sip_adet']), "Adet")
            st.write("**Ürün:**", int(st.session_state['amz_urun_adet']), "Adet")
            
        st.write("---")
        df_g = pd.DataFrame({"Pazaryeri": ["Trendyol", "Amazon"], "Ciro": [st.session_state['ty_ciro'], st.session_state['amz_ciro']]})
        st.plotly_chart(px.pie(df_g, values='Ciro', names='Pazaryeri', title="Ciro Dağılımı", hole=0.3), use_container_width=True)
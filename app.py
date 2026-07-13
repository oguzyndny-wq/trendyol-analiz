import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="E-Ticaret Konsolide Paneli v12.1", layout="wide")
st.title("🤖 Çok Kanallı E-Ticaret Konsolide Finans Paneli v12.1")
st.markdown("Amazon için sadece Maliyet Şablonunu kullanan, bölünme korumalı kararlı sürüm.")
st.write("---")

v_list = [
    'ty_ciro', 'ty_kesinti', 'ty_maliyet', 'ty_kar', 'ty_sip_adet', 'ty_urun_adet',
    'amz_ciro', 'amz_kesinti', 'amz_maliyet', 'amz_kar', 'amz_sip_adet', 'amz_urun_adet',
    'genel_ciro', 'genel_maliyet', 'genel_kar', 'genel_sip_adet', 'genel_urun_adet', 'genel_marj'
]
for k in v_list:
    if k not in st.session_state: st.session_state[k] = 0.0
if 'hesaplandi' not in st.session_state: st.session_state['hesaplandi'] = False

tab_yk, tab_rp = st.tabs(["📥 Veri Yükleme İstasyonu", "📊 Konsolide Finans & Raporlar"])

with tab_yk:
    st.subheader("1. Trendyol Raporları")
    col1, col2, col3 = st.columns(3)
    with col1: finans_file = st.file_uploader("SiparisNo dosyası", type=["xlsx", "xls"], key="finans")
    with col2: prod_file = st.file_uploader("prod_ dosyası", type=["xlsx", "xls"], key="prod")
    with col3: maliyet_file = st.file_uploader("Trendyol Maliyet", type=["xlsx", "xls"], key="maliyet")
    ty_reklam = st.number_input("🔗 Trendyol Reklam (TL):", min_value=0.0, value=0.0, step=100.0)
    st.write("---")
    st.subheader("2. Amazon Raporu")
    amazon_maliyet_file = st.file_uploader("Amazon Maliyet Şablonu (Tek Dosya Yeterlidir)", type=["xlsx", "xls", "csv"], key="amazon_cost")
    st.write("---")
    baslat_btn = st.button("🚀 Analizi Başlat", use_container_width=True)

def safe_num(v):
    if pd.isnull(v): return 0.0
    if isinstance(v, (int, float)): return float(v)
    try: return float(str(v).strip().replace('.', '').replace(',', '.'))
    except: return 0.0

if baslat_btn:
    st.session_state['hesaplandi'] = False
    t_ci, t_ke, t_ma, t_ka, t_si, t_ur = 0.0, 0.0, 0.0, 0.0, 0, 0
    a_ci, a_ke, a_ma, a_ka, a_si, a_ur = 0.0, 0.0, 0.0, 0.0, 0, 0
    
    # 🧡 TRENDYOL MOTORU
    if finans_file and prod_file and maliyet_file:
        try:
            df_f = pd.read_excel(finans_file)
            df_p = pd.read_excel(prod_file, skiprows=1)
            df_m = pd.read_excel(maliyet_file)
            df_m.columns = [c.strip() for c in df_m.columns]
            df_f.columns = [c.strip() for c in df_f.columns]
            df_p.columns = [c.strip() for c in df_p.columns]
            t_si = int(df_p['Sipariş Numarası'].nunique())
            
            f_dic = {}
            for idx, r in df_f.iterrows():
                sn = str(r['Sipariş No']).strip()
                f_dic[sn] = {
                    'n': safe_num(r['Ürün Adedi']),
                    'ko': safe_num(r['Komisyon/Yurt Dışı Stok Destek Bedeli']),
                    'ka': safe_num(r['Gönderi Kargo Bedeli']),
                    'hi': safe_num(r['Platform Hizmet Bedeli'])
                }
            m_dic = dict(zip(df_m['TRENDYOL BARKOD'].astype(str).str.strip(), df_m['TOPLAM MALİYET']))
            ty_res = []
            for idx, r in df_p.iterrows():
                bk = str(r['Barkod']).strip()
                sn = str(r['Sipariş Numarası']).strip()
                stt = str(r.get('Statü', r.get('Sipariş Durumu', ''))).strip().lower()
                ad = safe_num(r['Adet'])
                st_tut = safe_num(r['Satış Tutarı'])
                if bk == 'nan' or sn == 'nan' or "iptal" in stt or "reddedildi" in stt: continue
                t_ur += int(ad)
                b_ma = safe_num(m_dic.get(bk, 0.0))
                h_ci = 0.0 if "iade" in stt else st_tut
                h_ma = 0.0 if "iade" in stt else (b_ma * ad)
                
                f_dt = f_dic.get(sn, {'n': 0, 'ko': 0, 'ka': 0, 'hi': 0})
                div = f_dt['n'] if f_dt['n'] > 0 else 1
                b_ko = f_dt['ko'] / div * ad if f_dt['n'] > 0 else 0
                b_ka = f_dt['ka'] / div * ad if f_dt['n'] > 0 else 0
                b_hi = f_dt['hi'] / div * ad if f_dt['n'] > 0 else 0
                
                n_kr = h_ci + b_ko + b_ka + b_hi - h_ma
                ty_res.append([h_ci, b_ko + b_ka + b_hi, h_ma, n_kr])
            df_ty_r = pd.DataFrame(ty_res, columns=['C', 'K', 'M', 'R'])
            t_ci = df_ty_r['C'].sum()
            t_ke = abs(df_ty_r['K'].sum())
            t_ma = df_ty_r['M'].sum()
            t_ka = df_ty_r['R'].sum() - ty_reklam
        except Exception as e: st.error(f"Trendyol Hatası: {str(e)}")

    # 💛 AMAZON MOTORU (Tek Maliyet Şablonundan Okuma)
    if amazon_maliyet_file:
        try:
            df_a = pd.read_csv(amazon_maliyet_file) if hasattr(amazon_maliyet_file, 'name') and amazon_maliyet_file.name.endswith('.csv') else pd.read_excel(amazon_maliyet_file)
            df_a.columns = [c.strip() for c in df_a.columns]
            
            c0 = 'Satilan_Net_Birim' if 'Satilan_Net_Birim' in df_a.columns else df_a.columns[1]
            c1 = 'Brut_Satis' if 'Brut_Satis' in df_a.columns else df_a.columns[2]
            c2 = 'Amazon_Net_Kazanc' if 'Amazon_Net_Kazanc' in df_a.columns else df_a.columns[3]
            c3 = 'Birim Alış Maliyeti (₺)' if 'Birim Alış Maliyeti (₺)' in df_a.columns else df_a.columns[4]
            
            amz_res = []
            for idx, r in df_a.iterrows():
                u_ne = safe_num(r.get(c0, 0))
                s_br = safe_num(r.get(c1, 0.0))
                k_ne = safe_num(r.get(c2, 0.0))
                m_bi = safe_num(r.get(c3, 0.0))
                
                if u_ne > 0: a_si += 1
                a_ur += int(max(0.0, u_ne))
                tm = m_bi * max(0.0, u_ne)
                amz_res.append([s_br, s_br - k_ne, tm, k_ne - tm])
            df_amz_r = pd.DataFrame(amz_res, columns=['C', 'K', 'M', 'R'])
            a_ci = df_amz_r['C'].
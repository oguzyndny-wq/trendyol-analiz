import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Trendyol Finans Paneli", layout="wide")
st.title("🤖 Trendyol Finansal Analiz Paneli v13.0")
st.markdown("Amazon kaldırılmıştır. Sadece Trendyol raporları için %100 kararlı sürüm.")
st.write("---")

# Hafıza Değişkenleri
v_list = ['ty_ciro', 'ty_kesinti', 'ty_maliyet', 'ty_kar', 'ty_sip_adet', 'ty_urun_adet']
for k in v_list:
    if k not in st.session_state:
        st.session_state[k] = 0.0
if 'hesaplandi' not in st.session_state:
    st.session_state['hesaplandi'] = False

# Dosya Yükleme Alanı
st.subheader("📥 Trendyol Raporlarını Yükleyin")
col1, col2, col3 = st.columns(3)
with col1:
    f_file = st.file_uploader("1. SiparisKayitlari (Finans) Dosyası", type=["xlsx", "xls"])
with col2:
    p_file = st.file_uploader("2. prod_ (Sipariş Durum) Dosyası", type=["xlsx", "xls"])
with col3:
    m_file = st.file_uploader("3. Trendyol Maliyet Listesi", type=["xlsx", "xls"])

ty_rek = st.number_input("🔗 Varsa Trendyol Ekstra Reklam Gideri (TL):", min_value=0.0, value=0.0)
st.write("---")
baslat_btn = st.button("🚀 Trendyol Analizini Başlat", use_container_width=True)

def safe_f(v):
    if pd.isnull(v): return 0.0
    if isinstance(v, (int, float)): return float(v)
    try:
        s = str(v).strip().replace('.', '').replace(',', '.')
        return float(s)
    except:
        return 0.0

if baslat_btn:
    if f_file and p_file and m_file:
        try:
            st.session_state['hesaplandi'] = False
            
            df_f = pd.read_excel(f_file)
            df_p = pd.read_excel(p_file, skiprows=1)
            df_m = pd.read_excel(m_file)
            
            df_m.columns = [c.strip() for c in df_m.columns]
            df_f.columns = [c.strip() for c in df_f.columns]
            df_p.columns = [c.strip() for c in df_p.columns]
            
            t_si = int(df_p['Sipariş Numarası'].nunique())
            t_ur = 0
            
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
                
                if bk == 'nan' or sn == 'nan' or "iptal" in stt or "reddedildi" in stt:
                    continue
                    
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
            
            st.session_state['ty_ciro'] = df_ty_r['C'].sum()
            st.session_state['ty_kesinti'] = abs(df_ty_r['K'].sum())
            st.session_state['ty_maliyet'] = df_ty_r['M'].sum()
            st.session_state['ty_kar'] = df_ty_r['R'].sum() - ty_rek
            st.session_state['ty_sip_adet'] = t_si
            st.session_state['ty_urun_adet'] = t_ur
            st.session_state['hesaplandi'] = True
            st.success("🎉 Analiz Tamamlandı! Aşağıdaki rapordan inceleyebilirsiniz.")
        except Exception as e:
            st.error(f"Hesaplama hatası meydana geldi: {str(e)}")
    else:
        st.warning("Lütfen 3 Trendyol dosyasını da eksiksiz yükleyin.")

# Raporlama Ekranı
if st.session_state['hesaplandi']:
    st.write("---")
    st.subheader("📊 Trendyol Finansal Performans Özeti")
    
    g1, g2, g3, g4, g5, g6 = st.columns(6)
    g1.metric("💰 Net Ciro", "₺{:,.2f}".format(st.session_state['ty_ciro']))
    g2.metric("❌ Toplam Kesinti", "₺{:,.2f}".format(st.session_state['ty_kesinti']))
    g3.metric("📦 Ürün Maliyeti", "₺{:,.2f}".format(st.session_state['ty_maliyet']))
    g4.metric("🟢 Net Kâr", "₺{:,.2f}".format(st.session_state['ty_kar']))
    
    marj = (st.session_state['ty_kar'] / st.session_state['ty_ciro'] * 100.0) if st.session_state['ty_ciro'] > 0 else 0.0
    g5.metric("📈 Kanal Marjı", "%{:.2f}".format(marj))
    g6.metric("📦 Sipariş / Ürün", "{}/{}".format(int(st.session_state['ty_sip_adet']), int(st.session_state['ty_urun_adet'])))
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Trendyol Finans ve Detaylı Analiz", layout="wide")
st.title("🤖 Trendyol Konsolide Finans ve Ürün Analiz Paneli v14.4")
st.markdown("Hafıza kilitlenme korumalı (Auto-Reset) ve kesin hesaplamalı güncel sürüm.")
st.write("---")

# Session State Hafızası - Her Yenilemede Sıfırlama Garantisi
v_list = ['ty_ciro', 'ty_kesinti', 'ty_maliyet', 'ty_kar', 'ty_sip_adet', 'ty_urun_adet']
for k in v_list:
    if k not in st.session_state: st.session_state[k] = 0.0
if 'hesaplandi' not in st.session_state: st.session_state['hesaplandi'] = False
if 'df_detay' not in st.session_state: st.session_state['df_detay'] = None

# Dosya Yükleme Alanı
st.subheader("📥 Trendyol Raporlarını Yükleyin")
col1, col2, col3 = st.columns(3)
with col1: finans_file = st.file_uploader("1. SiparisKayitlari (Finans) Dosyası", type=["xlsx", "xls"], key="f_up")
with col2: prod_file = st.file_uploader("2. prod_ (Sipariş Durum) Dosyası", type=["xlsx", "xls"], key="p_up")
with col3: maliyet_file = st.file_uploader("3. Trendyol Maliyet Listesi", type=["xlsx", "xls"], key="m_up")

ty_rek = st.number_input("🔗 Varsa Trendyol Ekstra Reklam Gideri (TL):", min_value=0.0, value=0.0)
st.write("---")
baslat_btn = st.button("🚀 Hafızayı Temizle ve Tüm Analizleri Başlat", use_container_width=True)

def safe_f(v):
    if pd.isnull(v): return 0.0
    if isinstance(v, (int, float)): return float(v)
    try: return float(str(v).strip().replace('.', '').replace(',', '.'))
    except: return 0.0

def color_profit_loss(val):
    color = '#2ecc71' if val >= 0 else '#e74c3c'
    return f'color: white; background-color: {color}; font-weight: bold;'

if baslat_btn:
    if finans_file and prod_file and maliyet_file:
        try:
            # Eski hatalı kayıtları zorla sıfırlıyoruz (Hafıza Kilidi Kırıcı)
            st.session_state['hesaplandi'] = False
            st.session_state['df_detay'] = None
            
            df_f = pd.read_excel(finans_file)
            df_p = pd.read_excel(prod_file, skiprows=1)
            df_m = pd.read_excel(maliyet_file)
            
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
                    'ko': abs(safe_f(r['Komisyon/Yurt Dışı Stok Destek Bedeli'])),
                    'ka': abs(safe_f(r['Gönderi Kargo Bedeli'])),
                    'hi': abs(safe_f(r['Platform Hizmet Bedeli']))
                }
            
            m_dic = dict(zip(df_m['TRENDYOL BARKOD'].astype(str).str.strip(), df_m['TOPLAM MALİYET']))
            isim_col = 'ÜRÜN ADI' if 'ÜRÜN ADI' in df_m.columns else df_m.columns[1]
            name_dic = dict(zip(df_m['TRENDYOL BARKOD'].astype(str).str.strip(), df_m[isim_col]))
            
            urun_bazli = {}
            ty_res = []
            
            for idx, r in df_p.iterrows():
                bk = str(r['Barkod']).strip()
                sn = str(r['Sipariş Numarası']).strip()
                stt = str(r.get('Statü', r.get('Sipariş Durumu', ''))).strip().lower()
                ad = safe_f(r['Adet'])
                st_tut = safe_f(r['Satış Tutarı'])
                prod_name = str(r.get('Ürün Adı', r.get('Ürün', name_dic.get(bk, 'Bilinmeyen Ürün'))))
                
                if bk == 'nan' or sn == 'nan' or "iptal" in stt or "reddedildi" in stt:
                    continue
                    
                t_ur += int(ad)
                b_ma = safe_f(m_dic.get(bk, 0.0))
                h_ci = 0.0 if "iade" in stt else st_tut
                h_ma = 0.0 if "iade" in stt else (b_ma * ad)
                
                fd = f_dic.get(sn, {'n': 0, 'ko': 0, 'ka': 0, 'hi': 0})
                div = fd['n'] if fd['n'] > 0 else 1
                
                b_ko = (fd['ko'] / div) * ad if fd['n'] > 0 else 0.0
                b_ka = (fd['ka'] / div) * ad if fd['n'] > 0 else 0.0
                b_hi = (fd['hi'] / div) * ad if fd['n'] > 0 else 0.0
                toplam_kesinti_bileseni = b_ko + b_ka + b_hi
                
                n_kr = h_ci - toplam_kesinti_bileseni - h_ma
                ty_res.append([h_ci, toplam_kesinti_bileseni, h_ma, n_kr])
                
                if bk not in urun_bazli:
                    urun_bazli[bk] = {'Ürün Adı': prod_name, 'Satılan Adet': 0, 'Ciro': 0.0, 'Kâr / Zarar': 0.0}
                
                urun_bazli[bk]['Satılan Adet'] += int(ad)
                urun_bazli[bk]['Ciro'] += h_ci
                urun_bazli[bk]['Kâr / Zarar'] += n_kr

            df_ty_r = pd.DataFrame(ty_res, columns=['C', 'K', 'M', 'R'])
            st.session_state['ty_ciro'] = df_ty_r['C'].sum()
            st.session_state['ty_kesinti'] = df_ty_r['K'].sum()
            st.session_state['ty_maliyet'] = df_ty_r['M'].sum()
            st.session_state['ty_kar'] = df_ty_r['R'].sum() - ty_rek
            st.session_state['ty_sip_adet'] = t_si
            st.session_state['ty_urun_adet'] = t_ur

            df_detay = pd.DataFrame.from_dict(urun_bazli, orient='index').reset_index()
            df_detay.columns = ['Barkod', 'Ürün Adı', 'Satılan Adet', 'Ciro', 'Kâr / Zarar']
            df_detay = df_detay.sort_values(by='Kâr / Zarar', ascending=False).reset_index(drop=True)
            
            st.session_state['df_detay'] = df_detay
            st.session_state['hesaplandi'] = True
            st.success("🎉 Hafıza Temizlendi ve Tüm Hesaplamalar Yenilendi!")
        except Exception as e:
            st.error(f"Hesaplama hatası: {str(e)}")
    else:
        st.warning("Lütfen gerekli 3 dosyayı da yükleyin.")

# RAPORLAMA EKRANI
if st.session_state['hesaplandi']:
    st.write("---")
    st.subheader("📊 1. Trendyol Genel Finansal Performans Özeti")
    
    g1, g2, g3, g4, g5, g6 = st.columns(6)
    g1.metric("💰 Net Ciro", "₺{:,.2f}".format(st.session_state['ty_ciro']))
    g2.metric("❌ Toplam Kesinti", "₺{:,.2f}".format(st.session_state['ty_kesinti']))
    g3.metric("📦 Ürün Maliyeti", "₺{:,.2f}".format(st.session_state['ty_maliyet']))
    g4.metric("🟢 Net Kâr", "₺{:,.2f}".format(st.session_state['ty_kar']))
    
    marj = (st.session_state['ty_kar'] / st.session_state['ty_ciro'] * 100.0) if st.session_state['ty_ciro'] > 0 else 0.0
    g5.metric("📈 Kanal Marjı", "%{:.2f}".format(marj))
    g6.metric("📦 Sipariş / Ürün", "{}/{}".format(int(st.session_state['ty_sip_adet']), int(st.session_state['ty_urun_adet'])))
    
    if st.session_state['df_detay'] is not None:
        st.write("---")
        st.subheader("🔍 2. Ürün Bazlı Detaylı Kârlılık Raporu")
        
        df_goster = st.session_state['df_detay'].copy()
        styled_df = df_goster.style.format({
            'Ciro': '₺{:,.2f}', 'Kâr / Zarar': '₺{:,.2f}', 'Satılan Adet': '{:,}'
        }).map(color_profit_loss, subset=['Kâr / Zarar'])
        
        st.dataframe(styled_df, use_container_width=True, height=500)
        
        st.write("---")
        st.subheader("📈 Ürün Çeşitliliği Kârlılık Durumu")
        k1, k2 = st.columns(2)
        kar_edenler = df_goster[df_goster['Kâr / Zarar'] > 0]
        zarar_edenler = df_goster[df_goster['Kâr / Zarar'] < 0]
        
        k1.metric("🟢 Kâr Eden Ürün Çeşidi", f"{len(kar_edenler)} Çeşit", f"+₺{kar_edenler['Kâr / Zarar'].sum():,.2f} Toplam Kâr")
        k2.metric("🔴 Zarar Eden Ürün Çeşidi", f"{len(zarar_edenler)} Çeşit", f"-₺{abs(zarar_edenler['Kâr / Zarar'].sum()):,.2f} Toplam Zarar", delta_color="inverse")
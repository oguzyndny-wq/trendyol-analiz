import streamlit as st
import pandas as pd

st.set_page_config(page_title="Trendyol Detaylı Analiz", layout="wide")
st.title("🤖 Trendyol Ürün Bazlı Finansal Analiz Paneli v14.1")
st.markdown("Hatalı değişken ismi düzeltilmiş, kâr-zarar renklendirmeli kesin sürüm.")
st.write("---")

# Hafıza Değişkenleri
if 'hesaplandi' not in st.session_state:
    st.session_state['hesaplandi'] = False
if 'df_detay' not in st.session_state:
    st.session_state['df_detay'] = None

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
baslat_btn = st.button("🚀 Detaylı Ürün Analizini Başlat", use_container_width=True)

def safe_f(v):
    if pd.isnull(v): return 0.0
    if isinstance(v, (int, float)): return float(v)
    try:
        return float(str(v).strip().replace('.', '').replace(',', '.'))
    except:
        return 0.0

# Renklendirme Fonksiyonu (Kâr Yeşil, Zarar Kırmızı)
def color_profit_loss(val):
    color = '#2ecc71' if val >= 0 else '#e74c3c'
    return f'color: white; background-color: {color}; font-weight: bold;'

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
            
            # Finans verilerini sözlüğe toplama
            f_dic = {}
            for idx, r in df_f.iterrows():
                sn = str(r['Sipariş No']).strip()
                f_dic[sn] = {
                    'n': safe_f(r['Ürün Adedi']),
                    'ko': safe_f(r['Komisyon/Yurt Dışı Stok Destek Bedeli']),
                    'ka': safe_f(r['Gönderi Kargo Bedeli']),
                    'hi': safe_f(r['Platform Hizmet Bedeli'])
                }
            
            # Ürün isimlerini ve maliyetlerini sözlüğe toplama
            m_dic = dict(zip(df_m['TRENDYOL BARKOD'].astype(str).str.strip(), df_m['TOPLAM MALİYET']))
            isim_col = 'ÜRÜN ADI' if 'ÜRÜN ADI' in df_m.columns else df_m.columns[1]
            name_dic = dict(zip(df_m['TRENDYOL BARKOD'].astype(str).str.strip(), df_m[isim_col]))
            
            urun_bazli = {}
            
            for idx, r in df_p.iterrows():
                bk = str(r['Barkod']).strip()
                sn = str(r['Sipariş Numarası']).strip()
                stt = str(r.get('Statü', r.get('Sipariş Durumu', ''))).strip().lower()
                ad = safe_f(r['Adet'])
                st_tut = safe_f(r['Satış Tutarı'])
                prod_name = str(r.get('Ürün Adı', r.get('Ürün', name_dic.get(bk, 'Bilinmeyen Ürün'))))
                
                if bk == 'nan' or sn == 'nan' or "iptal" in stt or "reddedildi" in stt:
                    continue
                    
                # Hata veren m_dict ismi m_dic olarak düzeltildi
                b_ma = safe_f(m_dic.get(bk, 0.0))
                h_ci = 0.0 if "iade" in stt else st_tut
                h_ma = 0.0 if "iade" in stt else (b_ma * ad)
                
                fd = f_dic.get(sn, {'n': 0, 'ko': 0, 'ka': 0, 'hi': 0})
                div = fd['n'] if fd['n'] > 0 else 1
                b_ko = fd['ko'] / div * ad if fd['n'] > 0 else 0
                b_ka = fd['ka'] / div * ad if fd['n'] > 0 else 0
                b_hi = fd['hi'] / div * ad if fd['n'] > 0 else 0
                
                n_kr = h_ci + b_ko + b_ka + b_hi - h_ma
                
                if bk not in urun_bazli:
                    urun_bazli[bk] = {'Ürün Adı': prod_name, 'Satılan Adet': 0, 'Ciro': 0.0, 'Kâr / Zarar': 0.0}
                
                urun_bazli[bk]['Satılan Adet'] += int(ad)
                urun_bazli[bk]['Ciro'] += h_ci
                urun_bazli[bk]['Kâr / Zarar'] += n_kr

            # Sözlüğü DataFrame'e çevirme
            df_detay = pd.DataFrame.from_dict(urun_bazli, orient='index').reset_index()
            df_detay.columns = ['Barkod', 'Ürün Adı', 'Satılan Adet', 'Ciro', 'Kâr / Zarar']
            
            # Kâr / Zarar durumuna göre en yüksek kârdan en düşüğe sırala
            df_detay = df_detay.sort_values(by='Kâr / Zarar', ascending=False).reset_index(drop=True)
            
            st.session_state['df_detay'] = df_detay
            st.session_state['hesaplandi'] = True
            st.success("🎉 Ürün bazlı detaylı kârlılık analizi tamamlandı!")
        except Exception as e:
            st.error(f"Hesaplama hatası: {str(e)}")
    else:
        st.warning("Lütfen analiz için gerekli 3 dosyayı da yükleyin.")

# Raporlama Ekranı
if st.session_state['hesaplandi'] and st.session_state['df_detay'] is not None:
    st.write("---")
    st.subheader("📊 Ürün Bazlı Detaylı Kârlılık Raporu")
    st.markdown("Aşağıdaki tabloda **Kâr / Zarar** sütunu kâr eden ürünler için **Yeşil**, zarar edenler için **Kırmızı** renkte boyanmıştır.")
    
    df_goster = st.session_state['df_detay'].copy()
    
    styled_df = df_goster.style.format({
        'Ciro': '₺{:,.2f}',
        'Kâr / Zarar': '₺{:,.2f}',
        'Satılan Adet': '{:,}'
    }).map(color_profit_loss, subset=['Kâr / Zarar'])
    
    st.dataframe(styled_df, use_container_width=True, height=600)
    
    # Hızlı İstatistik Kartları
    st.write("---")
    st.subheader("📈 Genel Kârlılık Durumu")
    k1, k2 = st.columns(2)
    
    kar_edenler = df_goster[df_goster['Kâr / Zarar'] > 0]
    zarar_edenler = df_goster[df_goster['Kâr / Zarar'] < 0]
    
    k1.metric("🟢 Kâr Eden Ürün Çeşidi", f"{len(kar_edenler)} Çeşit", f"+₺{kar_edenler['Kâr / Zarar'].sum():,.2f} Toplam Kâr")
    k2.metric("🔴 Zarar Eden Ürün Çeşidi", f"{len(zarar_edenler)} Çeşit", f"-₺{abs(zarar_edenler['Kâr / Zarar'].sum()):,.2f} Toplam Zarar", delta_color="inverse")
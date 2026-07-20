import streamlit as st
import pandas as pd
import datetime
import os

st.set_page_config(page_title="PRİME ENTEGRE ERP v35.2", layout="wide")

# 🖼️ KURUMSAL LOGO
LOGO_PATH = "logo.png"
if os.path.exists(LOGO_PATH):
    st.sidebar.image(LOGO_PATH, use_container_width=True)

st.title("📈 PRİME ENTEGRE E-TİCARET LTD. ŞTİ. — Finansal Denetim İstasyonu")

# Sayı Temizleme Fonksiyonu
def safe_f(v):
    if pd.isnull(v): return 0.0
    if isinstance(v, (int, float)): return float(v)
    try:
        s = str(v).strip().replace(' ', '').replace('₺', '')
        if '.' in s and ',' in s: s = s.replace('.', '').replace(',', '.')
        elif ',' in s: s = s.replace(',', '.')
        return float(s)
    except: return 0.0

def color_profit_loss(val):
    if isinstance(val, (int,float)):
        color = '#2ecc71' if val >= 0 else '#e74c3c'
        return f'color: white; background-color: {color}; font-weight: bold;'
    return ''

# 📄 EVRENSEL DOSYA OKUYUCU (Tüm Dosya İsimleri, XLSX, CSV ve TXT Destekler)
def load_data(file_obj):
    fname = file_obj.name.lower()
    if fname.endswith('.csv'):
        try:
            return pd.read_csv(file_obj)
        except Exception:
            file_obj.seek(0)
            return pd.read_csv(file_obj, sep=';')
    elif fname.endswith('.txt'):
        return pd.read_csv(file_obj, sep='\t')
    else:
        return pd.read_excel(file_obj)

# Session State Hafızaları
v_list = [
    'ty_ciro', 'ty_kesinti', 'ty_maliyet', 'ty_kar', 'ty_sip_adet', 'ty_urun_adet', 'ty_iptal_adet', 'ty_iade_adet', 'ty_hakedis',
    'amz_ciro', 'amz_kesinti', 'amz_maliyet', 'amz_kar', 'amz_sip_adet', 'amz_urun_adet', 'amz_iptal_adet', 'amz_iade_adet',
    'amz_resmi_lojistik', 'amz_resmi_komisyon', 'amz_resmi_hakedis',
    'hesaplandi_ty', 'hesaplandi_amz', 'df_detay_ty', 'df_siparisler_ty', 'df_nakit_akis_ty', 'eksik_barkodlar_ty', 'df_olu_urunler_ty',
    'df_detay_amz', 'df_siparisler_amz', 'df_olu_urunler_amz'
]
for k in v_list:
    if k not in st.session_state:
        if 'df' in k or 'list' in k or 'eksik' in k: st.session_state[k] = None
        elif 'hesaplandi' in k: st.session_state[k] = False
        else: st.session_state[k] = 0.0

# 📥 HAM VERİ GİRİŞ TERMİNALİ
st.subheader("📥 Finansal Rapor Giriş Paneli & Veri Entegrasyonu")
mapping_file = st.file_uploader("🔗 0. Çoklu Barkod Ürün Eşleştirme Kılavuzu (urun_eslestirme.xlsx)", type=["xlsx", "xls", "csv"])

tab_ty, tab_amz = st.tabs(["🟢 TRENDYOL RAPORLARI", "🟠 AMAZON RAPORLARI"])

with tab_ty:
    col1, col2, col3 = st.columns(3)
    with col1: finans_file = st.file_uploader("1. Trendyol Finans Dosyası", type=["xlsx", "xls", "csv"], key="f_ty")
    with col2: prod_file = st.file_uploader("2. Trendyol Sipariş Durum (prod_) Dosyası", type=["xlsx", "xls", "csv"], key="p_ty")
    with col3: maliyet_file = st.file_uploader("3. Trendyol Maliyet Listesi", type=["xlsx", "xls", "csv"], key="m_ty")
    ty_rek = st.number_input("🔗 Trendyol Toplam Reklam Gideri (TL):", min_value=0.0, value=0.0, step=100.0)

with tab_amz:
    col_a1, col_a2 = st.columns(2)
    with col_a1: amz_sip_file = st.file_uploader("1. Amazon Satış Kayıtları / Raporu (CSV, TXT, XLSX, XLS)", type=["csv", "txt", "xlsx", "xls"], key="s_amz")
    with col_a2: amz_mal_file = st.file_uploader("2. Amazon Ürün Maliyet / SKU Şablonu (XLSX, XLS, CSV)", type=["xlsx", "xls", "csv"], key="m_amz")
    amz_rek = st.number_input("🔗 Amazon Panel Dışı Harici Reklam Gideri (TL):", min_value=0.0, value=0.0, step=100.0)

st.write("---")
baslat_btn = st.button("🚀 TÜM MAĞAZALARI VE KONSOLİDE HESAPLAMAYI BAŞLAT", use_container_width=True)

if baslat_btn:
    # 🟠 AMAZON HESAPLAMA MOTORU (CSV & XLSX Destekli)
    if amz_sip_file and amz_mal_file:
        try:
            df_as = load_data(amz_sip_file)
            df_am = load_data(amz_mal_file)
            
            df_as.columns = [str(c).strip() for c in df_as.columns]
            df_am.columns = [str(c).strip() for c in df_am.columns]
            
            # Dinamik Sütun Tespiti
            asin_col = next((c for c in ["Ana ürün ASIN'i", "(Ana Ürün) ASIN", "ASIN", "asin", "SKU", "sku"] if c in df_am.columns), df_am.columns[0])
            birim_col = next((c for c in ["Satilan_Net_Birim", "Satılan birimler", "Sipariş edilen birimler", "Satılan net birim sayısı", "quantity"] if c in df_am.columns), None)
            brut_col = next((c for c in ["Brut_Satis", "Satış", "Sipariş edilen ürün satışları", "item-price"] if c in df_am.columns), None)
            kazanc_col = next((c for c in ["Amazon_Net_Kazanc", "Toplam Net kazanç", "Net ödeme"] if c in df_am.columns), None)
            maliyet_col = next((c for c in ["Birim Alış Maliyeti (₺)", "Birim Alış Maliyeti", "Birim Maliyet"] if c in df_am.columns), None)
            
            # Ciro ve Net Kazanç Hesaplamaları
            amz_c = df_am[brut_col].apply(safe_f).sum() if brut_col else 0.0
            amz_k_net = df_am[kazanc_col].apply(safe_f).sum() if kazanc_col else amz_c * 0.70
            
            if birim_col and maliyet_col:
                df_am['Mal_Maliyet'] = df_am[birim_col].apply(safe_f) * df_am[maliyet_col].apply(safe_f)
            else:
                df_am['Mal_Maliyet'] = 0.0
                
            amz_m = df_am['Mal_Maliyet'].sum()
            amz_kesinti = amz_c - amz_k_net
            
            st.session_state['amz_ciro'] = amz_c
            st.session_state['amz_kesinti'] = amz_kesinti
            st.session_state['amz_maliyet'] = amz_m
            st.session_state['amz_kar'] = amz_k_net - amz_m - amz_rek
            
            if "amazon-order-id" in df_as.columns:
                st.session_state['amz_sip_adet'] = df_as['amazon-order-id'].nunique()
            elif birim_col:
                st.session_state['amz_sip_adet'] = int(df_am[birim_col].apply(safe_f).sum())
            else:
                st.session_state['amz_sip_adet'] = len(df_as)
                
            st.session_state['amz_urun_adet'] = int(df_am[birim_col].apply(safe_f).sum()) if birim_col else len(df_am)
            
            # Tablo Hazırlığı
            df_am_detay = pd.DataFrame()
            df_am_detay['ASIN'] = df_am[asin_col]
            df_am_detay['Ürün Adı'] = df_am['Ürün Adı'] if 'Ürün Adı' in df_am.columns else df_am[asin_col]
            df_am_detay['Satılan Adet'] = df_am[birim_col].apply(safe_f) if birim_col else 1
            df_am_detay['Ciro'] = df_am[brut_col].apply(safe_f) if brut_col else 0.0
            df_am_detay['Amazon Net Kazanç'] = df_am[kazanc_col].apply(safe_f) if kazanc_col else df_am_detay['Ciro'] * 0.70
            df_am_detay['Toplam Ürün Maliyeti'] = df_am['Mal_Maliyet']
            df_am_detay['Kâr / Zarar'] = df_am_detay['Amazon Net Kazanç'] - df_am_detay['Toplam Ürün Maliyeti']
            
            st.session_state['df_detay_amz'] = df_am_detay.sort_values(by='Kâr / Zarar', ascending=False).reset_index(drop=True)
            st.session_state['hesaplandi_amz'] = True
            st.success("✅ Amazon Verileri Başarıyla Hesaplandı!")
        except Exception as e:
            st.error(f"Amazon Dosya Okuma Hatası: {str(e)}")

# 📊 EKRAN GÖSTERİMİ
if st.session_state['hesaplandi_amz']:
    st.write("---")
    st.subheader("📊 Amazon Mağaza Kontrol İstasyonu")
    a1, a2, a3, a4 = st.columns(4)
    a1.metric("Amazon Net Ciro", "₺{:,.2f}".format(st.session_state['amz_ciro']))
    a2.metric("Amazon Toplam Kesinti", "₺{:,.2f}".format(st.session_state['amz_kesinti']))
    a3.metric("Amazon Net Kâr", "₺{:,.2f}".format(st.session_state['amz_kar']))
    a4.metric("🛒 Toplam Ürün Adedi", f"{int(st.session_state['amz_urun_adet'])} Adet")
    
    st.write("---")
    st.dataframe(st.session_state['df_detay_amz'].style.format({
        'Ciro': '₺{:,.2f}', 
        'Amazon Net Kazanç': '₺{:,.2f}', 
        'Toplam Ürün Maliyeti': '₺{:,.2f}', 
        'Kâr / Zarar': '₺{:,.2f}'
    }).map(color_profit_loss, subset=['Kâr / Zarar']), use_container_width=True, height=400)
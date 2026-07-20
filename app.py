import streamlit as st
import pandas as pd
import datetime
import io
import os

st.set_page_config(page_title="PRİME ENTEGRE ERP v35.2 Evrensel", layout="wide")

# 🖼️ KURUMSAL LOGO ENTEGRASYON YÖNETİCİSİ
LOGO_PATH = "logo.png"

if os.path.exists(LOGO_PATH):
    st.sidebar.image(LOGO_PATH, use_container_width=True)
    st.sidebar.write("---")

col_title, col_logo = st.columns([8, 2])
with col_title:
    st.title("📈 PRİME ENTEGRE E-TİCARET LTD. ŞTİ. — Konsolide Finansal Denetim İstasyonu")
    st.markdown("Prime Entegre bünyesindeki tüm pazaryerlerinin anlık kârlılık, finansal başabaş analizi, lojistik maliyet ve holding performans göstergeleri.")

with col_logo:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=150)

st.write("---")

# Güvenli Sayı Temizleme Fonksiyonu
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

# Evrensel Dosya Okuyucu (CSV, TXT, XLSX, XLS)
def read_uploaded_file(uploaded_file):
    fname = uploaded_file.name.lower()
    if fname.endswith('.csv'):
        try:
            return pd.read_csv(uploaded_file)
        except Exception:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, sep=';')
    elif fname.endswith('.txt'):
        return pd.read_csv(uploaded_file, sep='\t')
    else:
        return pd.read_excel(uploaded_file)

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
mapping_file = st.file_uploader("🔗 0. Çoklu Barkod Ürün Eşleştirme Kılavuzu (Örn: urun_eslestirme.xlsx)", type=["xlsx", "xls"])

tab_ty, tab_amz = st.tabs(["🟢 TRENDYOL RAPORLARI", "🟠 AMAZON RAPORLARI"])

with tab_ty:
    col1, col2, col3 = st.columns(3)
    with col1: finans_file = st.file_uploader("1. Trendyol Sipariş Kayıtları (Finans) Dosyası", type=["xlsx", "xls"], key="f_ty")
    with col2: prod_file = st.file_uploader("2. Trendyol Sipariş Durum (prod_) Dosyası", type=["xlsx", "xls"], key="p_ty")
    with col3: maliyet_file = st.file_uploader("3. Trendyol Maliyet Listesi", type=["xlsx", "xls"], key="m_ty")
    ty_rek = st.number_input("🔗 Trendyol Toplam Reklam Gideri (TL):", min_value=0.0, value=0.0, step=100.0)

with tab_amz:
    col_a1, col_a2 = st.columns(2)
    with col_a1: amz_sip_file = st.file_uploader("1. Amazon Satış Kayıtları / Rapor Dosyası (CSV, TXT, XLSX, XLS)", type=["csv", "txt", "xlsx", "xls"], key="s_amz")
    with col_a2: amz_mal_file = st.file_uploader("2. Amazon Ürün Maliyet / SKU Şablonu (XLSX, XLS)", type=["xlsx", "xls"], key="m_amz")
    amz_rek = st.number_input("🔗 Amazon Panel Dışı Harici Reklam Gideri (TL):", min_value=0.0, value=0.0, step=100.0)

st.write("---")
baslat_btn = st.button("🚀 TÜM MAĞAZALARI VE KONSOLİDE HESAPLAMAYI BAŞLAT", use_container_width=True)

if baslat_btn:
    # 🟢 1. TRENDYOL HESAPLAMA MOTORU (Tam Dinamik)
    if finans_file and prod_file and maliyet_file:
        try:
            df_f = read_uploaded_file(finans_file)
            df_p = pd.read_excel(prod_file, skiprows=1) if prod_file.name.endswith(('.xlsx', '.xls')) else read_uploaded_file(prod_file)
            df_m = read_uploaded_file(maliyet_file)
            
            df_m.columns = [c.strip() for c in df_m.columns]
            df_f.columns = [c.strip() for c in df_f.columns]
            df_p.columns = [c.strip() for c in df_p.columns]
            
            f_iptal_serisi = df_f['Sipariş Statüsü'].astype(str).str.lower().str.strip() if 'Sipariş Statüsü' in df_f.columns else pd.Series([])
            t_iptal = len(df_f[f_iptal_serisi.str.contains('iptal')]) if not f_iptal_serisi.empty else 0
            t_iade = len(df_f[f_iptal_serisi.str.contains('iade')]) if not f_iptal_serisi.empty else 0
            
            f_dic = {}
            for idx, r in df_f.iterrows():
                sn = str(r['Sipariş No']).strip() if 'Sipariş No' in df_f.columns else str(idx)
                f_dic[sn] = {
                    'statuler': str(r.get('Sipariş Statüsü', 'Shipped')), 'tarih': r.get('Sipariş Tarihi', None), 'n': safe_f(r.get('Ürün Adedi', 1)),
                    'ko': abs(safe_f(r.get('Komisyon/Yurt Dışı Stok Destek Bedeli', 0.0))), 'ka': abs(safe_f(r.get('Gönderi Kargo Bedeli', 0.0))),
                    'ika': abs(safe_f(r.get('İade Kargo Bedeli', 0.0))), 'ceza': abs(safe_f(r.get('Ceza Bedeli', 0.0))),
                    'hi': abs(safe_f(r.get('Platform Hizmet Bedeli', 0.0))), 'net_tutar': safe_f(r.get('Net Tutar', 0.0))
                }
            
            barkod_col = 'TRENDYOL BARKOD' if 'TRENDYOL BARKOD' in df_m.columns else df_m.columns[0]
            maliyet_col = 'TOPLAM MALİYET' if 'TOPLAM MALİYET' in df_m.columns else df_m.columns[-1]
            m_dic = dict(zip(df_m[barkod_col].astype(str).str.strip(), df_m[maliyet_col]))
            name_dic = dict(zip(df_m[barkod_col].astype(str).str.strip(), df_m['ÜRÜN ADI' if 'ÜRÜN ADI' in df_m.columns else df_m.columns[1]]))
            
            urun_bazli = {}
            siparis_bazli = {}
            eksik_b_set = set()
            calc_ty_ciro, calc_ty_kesinti, calc_ty_maliyet = 0.0, 0.0, 0.0
            t_ur = 0
            
            for idx, r in df_p.iterrows():
                bk = str(r.get('Barkod', '')).strip()
                sn = str(r.get('Sipariş Numarası', '')).strip()
                ad = safe_f(r.get('Adet', 1))
                if not bk or bk == 'nan' or not sn or sn == 'nan': continue
                if bk not in m_dic: eksik_b_set.add(bk)
                
                fatura_tutari = safe_f(r.get('Faturalanacak Tutar', r.get('Satış Tutarı', 0.0)))
                prod_name = str(r.get('Ürün Adı', name_dic.get(bk, 'Bilinmeyen Ürün')))
                fd = f_dic.get(sn, {'statuler': 'Shipped', 'tarih': None, 'n': 1, 'ko': 0, 'ka': 0, 'ika': 0, 'ceza': 0, 'hi': 0, 'net_tutar': 0.0})
                f_durum = fd['statuler'].lower()
                
                if "iptal" in f_durum: continue
                t_ur += int(ad)
                b_ma = safe_f(m_dic.get(bk, 0.0))
                h_ma = 0.0 if "iade" in f_durum else (b_ma * ad)
                h_ci = 0.0 if "iade" in f_durum else fatura_tutari
                
                div = fd['n'] if fd['n'] > 0 else 1
                toplam_kesinti = ((fd['ko'] + fd['ka'] + fd['hi'] + fd['ika'] + fd['ceza']) / div) * ad
                
                calc_ty_ciro += h_ci
                calc_ty_kesinti += toplam_kesinti
                calc_ty_maliyet += h_ma
                
                if bk not in urun_bazli: urun_bazli[bk] = {'Ürün Adı': prod_name, 'Satılan Adet': 0, 'Ciro': 0.0, 'Kâr / Zarar': 0.0}
                urun_bazli[bk]['Satılan Adet'] += int(ad)
                urun_bazli[bk]['Ciro'] += h_ci
                urun_bazli[bk]['Kâr / Zarar'] += (h_ci - toplam_kesinti - h_ma)
                
            st.session_state['ty_ciro'] = calc_ty_ciro
            st.session_state['ty_kesinti'] = calc_ty_kesinti
            st.session_state['ty_hakedis'] = calc_ty_ciro - calc_ty_kesinti
            st.session_state['ty_maliyet'] = calc_ty_maliyet
            st.session_state['ty_kar'] = st.session_state['ty_hakedis'] - calc_ty_maliyet - ty_rek
            st.session_state['ty_sip_adet'] = len(df_p['Sipariş Numarası'].unique()) if 'Sipariş Numarası' in df_p.columns else len(df_p)
            st.session_state['ty_urun_adet'] = t_ur
            st.session_state['ty_iptal_adet'] = t_iptal
            st.session_state['ty_iade_adet'] = t_iade
            st.session_state['df_detay_ty'] = pd.DataFrame.from_dict(urun_bazli, orient='index').reset_index().rename(columns={'index': 'Barkod'})
            st.session_state['hesaplandi_ty'] = True
        except Exception as e:
            st.error(f"Trendyol Hesaplama Hatası: {str(e)}")

    # 🟠 2. AMAZON HESAPLAMA MOTORU (Tüm Dosya İsimleri ve Formatlar İçin Evrensel)
    if amz_sip_file and amz_mal_file:
        try:
            df_as = read_uploaded_file(amz_sip_file)
            df_am = read_uploaded_file(amz_mal_file)
            
            df_as.columns = [c.strip() for c in df_as.columns]
            df_am.columns = [c.strip() for c in df_am.columns]
            
            # Dinamik Sütun Yakalayıcılar
            asin_col = None
            for candidate in ["Ana ürün ASIN'i", "(Ana Ürün) ASIN", "ASIN", "asin", "sku", "SKU"]:
                if candidate in df_am.columns:
                    asin_col = candidate
                    break
            if not asin_col: asin_col = df_am.columns[0]
            
            birim_col = None
            for candidate in ["Satilan_Net_Birim", "Satılan birimler", "Sipariş edilen birimler", "Satılan net birim sayısı", "quantity"]:
                if candidate in df_am.columns:
                    birim_col = candidate
                    break
            if not birim_col: birim_col = df_am.columns[1]
            
            brut_col = None
            for candidate in ["Brut_Satis", "Satış", "Sipariş edilen ürün satışları", "item-price"]:
                if candidate in df_am.columns:
                    brut_col = candidate
                    break
            if not brut_col: brut_col = df_am.columns[2]
            
            kazanc_col = None
            for candidate in ["Amazon_Net_Kazanc", "Toplam Net kazanç", "Net ödeme"]:
                if candidate in df_am.columns:
                    kazanc_col = candidate
                    break
            
            alis_col = None
            for candidate in ["Birim Alış Maliyeti (₺)", "Birim Alış Maliyeti", "Birim Maliyet"]:
                if candidate in df_am.columns:
                    alis_col = candidate
                    break
            
            # Canlı dinamik hesaplama
            amz_c = safe_f(df_am[brut_col].apply(safe_f).sum())
            
            if kazanc_col:
                amz_k_net = safe_f(df_am[kazanc_col].apply(safe_f).sum())
            else:
                amz_k_net = amz_c * 0.70  # Standart tahmini oran
                
            if alis_col:
                df_am['Mal_Maliyet'] = df_am[birim_col].apply(safe_f) * df_am[alis_col].apply(safe_f)
            else:
                df_am['Mal_Maliyet'] = 0.0
                
            amz_m = df_am['Mal_Maliyet'].sum()
            amz_karsi_kesinti = amz_c - amz_k_net
            
            st.session_state['amz_resmi_lojistik'] = amz_karsi_kesinti * 0.35
            st.session_state['amz_resmi_komisyon'] = amz_karsi_kesinti * 0.65
            st.session_state['amz_resmi_hakedis'] = amz_k_net
            st.session_state['amz_ciro'] = amz_c
            st.session_state['amz_kesinti'] = amz_karsi_kesinti
            st.session_state['amz_maliyet'] = amz_m
            st.session_state['amz_kar'] = amz_k_net - amz_m - amz_rek
            
            # Sipariş Adedi Tespiti
            if "amazon-order-id" in df_as.columns:
                st.session_state['amz_sip_adet'] = df_as['amazon-order-id'].nunique()
            else:
                st.session_state['amz_sip_adet'] = int(df_am[birim_col].apply(safe_f).sum())
                
            st.session_state['amz_urun_adet'] = int(df_am[birim_col].apply(safe_f).sum())
            
            # Detay Tablosu Oluşturma
            df_am_detay = pd.DataFrame()
            df_am_detay['ASIN'] = df_am[asin_col]
            df_am_detay['Ürün Adı'] = df_am['Ürün Adı'] if 'Ürün Adı' in df_am.columns else df_am[asin_col]
            df_am_detay['Satılan Adet'] = df_am[birim_col].apply(safe_f)
            df_am_detay['Ciro'] = df_am[brut_col].apply(safe_f)
            df_am_detay['Amazon Net Kazanç'] = df_am[kazanc_col].apply(safe_f) if kazanc_col else df_am_detay['Ciro'] * 0.70
            df_am_detay['Toplam Ürün Maliyeti'] = df_am['Mal_Maliyet']
            df_am_detay['Kâr / Zarar'] = df_am_detay['Amazon Net Kazanç'] - df_am_detay['Toplam Ürün Maliyeti']
            
            st.session_state['df_detay_amz'] = df_am_detay.sort_values(by='Kâr / Zarar', ascending=False).reset_index(drop=True)
            st.session_state['hesaplandi_amz'] = True
        except Exception as e:
            st.error(f"Amazon Motoru Okuma Hatası: {str(e)}")

# 👑 KONSOLİDE GLOBAL FINANSAL ÖZET
if st.session_state['hesaplandi_ty'] or st.session_state['hesaplandi_amz']:
    st.write("---")
    st.subheader("💼 PRİME ENTEGRE — Konsolide Holding Finansal Özet Paneli")
    
    total_ciro = st.session_state['ty_ciro'] + st.session_state['amz_ciro']
    total_kesinti = st.session_state['ty_kesinti'] + st.session_state['amz_kesinti']
    total_maliyet = st.session_state['ty_maliyet'] + st.session_state['amz_maliyet']
    total_kar = st.session_state['ty_kar'] + st.session_state['amz_kar']
    total_sip = st.session_state['ty_sip_adet'] + st.session_state['amz_sip_adet']
    total_reklam = ty_rek + amz_rek
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("💰 Toplam Ortak Net Ciro", "₺{:,.2f}".format(total_ciro))
    c2.metric("❌ Toplam Ortak Kesinti", "₺{:,.2f}".format(total_kesinti))
    c3.metric("📦 Toplam Ürün Sermayesi", "₺{:,.2f}".format(total_maliyet))
    c4.metric("🟢 Konsolide Net Saf Kâr", "₺{:,.2f}".format(total_kar))
    
    st.write(" ")
    cc1, cc2, cc3 = st.columns(3)
    g_marj = (total_kar / total_ciro * 100.0) if total_ciro > 0 else 0.0
    cc1.metric("📊 Ortak Net Kâr Marjı (%)", "%{:.2f}".format(g_marj))
    cc2.metric("🛒 Holding Ortak Sepet Ortalaması", "₺{:.2f}".format(total_ciro / total_sip if total_sip > 0 else 0.0))
    cc3.metric("🛒 Toplam Sipariş Hacmi", f"{total_sip} Sipariş")
    
    # 👑 ÇOK KANALLI BİRLEŞİK ÜRÜN MATRİSİ
    if mapping_file is not None:
        try:
            df_map = read_uploaded_file(mapping_file)
            df_map.columns = [c.strip() for c in df_map.columns]
            code_to_sku, sku_names = {}, {}
            barkod_sutunlari = [c for c in df_map.columns if 'barkod' in c.lower() or 'Barkod' in c or 'ASIN' in c or 'asin' in c]
            
            stok_kodu_col = 'ORTAK_URUN_STOK_KODU' if 'ORTAK_URUN_STOK_KODU' in df_map.columns else df_map.columns[0]
            
            for idx, r_map in df_map.iterrows():
                main_sku = str(r_map[stok_kodu_col]).strip()
                sku_names[main_sku] = str(r_map.get('Ürün Adı', main_sku)).strip()
                for b_col in barkod_sutunlari:
                    val = str(r_map[b_col]).strip()
                    if val != 'nan' and val != "": code_to_sku[val] = main_sku
            
            sku_aggr = {}
            if st.session_state['df_detay_ty'] is not None:
                for idx, r_ty in st.session_state['df_detay_ty'].iterrows():
                    sku = code_to_sku.get(str(r_ty['Barkod']).strip(), f"Eşleşmemiş Trendyol ({r_ty['Barkod']})")
                    if sku not in sku_aggr: sku_aggr[sku] = {'Ortak Stok Kodu': sku, 'Ürün Tanımı': sku_names.get(sku, sku), 'Trendyol Satış Adet': 0, 'Trendyol Ciro': 0.0, 'Amazon Satış Adet': 0, 'Amazon Ciro': 0.0}
                    sku_aggr[sku]['Trendyol Satış Adet'] += int(r_ty['Satılan Adet'])
                    sku_aggr[sku]['Trendyol Ciro'] += safe_f(r_ty['Ciro'])
                    
            if st.session_state['df_detay_amz'] is not None:
                for idx, r_amz in st.session_state['df_detay_amz'].iterrows():
                    sku = code_to_sku.get(str(r_amz['ASIN']).strip(), f"Eşleşmemiş Amazon ({r_amz['ASIN']})")
                    if sku not in sku_aggr: sku_aggr[sku] = {'Ortak Stok Kodu': sku, 'Ürün Tanımı': sku_names.get(sku, sku), 'Trendyol Satış Adet': 0, 'Trendyol Ciro': 0.0, 'Amazon Satış Adet': 0, 'Amazon Ciro': 0.0}
                    sku_aggr[sku]['Amazon Satış Adet'] += int(r_amz['Satılan Adet'])
                    sku_aggr[sku]['Amazon Ciro'] += safe_f(r_amz['Ciro'])
            
            st.write("---")
            st.markdown("### 👑 Çok Kanallı Birleşik Ürün Kârlılık Matrisi")
            st.dataframe(pd.DataFrame(list(sku_aggr.values())), use_container_width=True)
        except Exception as e:
            st.error(f"Matris Eşleştirme Hatası: {str(e)}")

    st.write("---")
    s_ty, s_amz = st.tabs(["🟢 TRENDYOL DETAYLI ERP PANELİ", "🟠 AMAZON DETAYLI ERP PANELİ"])
    
    with s_ty:
        if st.session_state['hesaplandi_ty']:
            st.subheader("📊 Trendyol Mağaza Kontrol İstasyonu")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Net Ciro", "₺{:,.2f}".format(st.session_state['ty_ciro']))
            m2.metric("Trendyol Kesintileri", "₺{:,.2f}".format(st.session_state['ty_kesinti']))
            m3.metric("Gelecek Net Tutar", "₺{:,.2f}".format(st.session_state['ty_hakedis']))
            m4.metric("Net Kâr", "₺{:,.2f}".format(st.session_state['ty_kar']))
            st.dataframe(st.session_state['df_detay_ty'], use_container_width=True)
            
    with s_amz:
        if st.session_state['hesaplandi_amz']:
            st.subheader("📊 Amazon Mağaza Kontrol İstasyonu")
            a1, a2, a3, a4 = st.columns(4)
            a1.metric("Amazon Net Ciro", "₺{:,.2f}".format(st.session_state['amz_ciro']))
            a2.metric("Amazon Toplam Kesinti", "₺{:,.2f}".format(st.session_state['amz_kesinti']))
            a3.metric("Amazon Net Kâr", "₺{:,.2f}".format(st.session_state['amz_kar']))
            a4.metric("🛒 Toplam Ürün Adedi", f"{int(st.session_state['amz_urun_adet'])} Adet")
            st.dataframe(st.session_state['df_detay_amz'], use_container_width=True)
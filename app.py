import streamlit as st
import pandas as pd
import plotly.express as px
import io

# Sayfa Genişlik ve Başlık Ayarları
st.set_page_config(page_title="Trendyol Akıllı Yapay Zeka Paneli v5.3", layout="wide")

st.title("🤖 Trendyol Akıllı Yapay Zeka Paneli v5.3 (Tekli Barkod Destekli)")
st.markdown("Fiyat tavsiyeleri, reklam motoru ve tekli barkod yazdırma istasyonu entegre edilmiştir.")
st.write("---")

# 1. DOSYA YÜKLEME ALANLARI
col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("📊 Sipariş Kayıtları")
    finans_file = st.file_uploader("SiparisKayitlari ile başlayan dosya", type=["xlsx", "xls"], key="finans")

with col2:
    st.subheader("📦 Sipariş Listesi")
    prod_file = st.file_uploader("prod_ ile başlayan dosya", type=["xlsx", "xls"], key="prod")

with col3:
    st.subheader("🏷️ Ürün Alış Maliyetleri")
    maliyet_file = st.file_uploader("Maliyet listesi dosyası", type=["xlsx", "xls"], key="maliyet")

st.write("---")

# Reklam Gideri Giriş Alanı
reklam_gideri = st.number_input("🔗 Varsa Bu Aya Ait Toplam Reklam Giderini Giriş Yapın (TL):", min_value=0.0, value=0.0, step=100.0)

# ANALİZİ BAŞLAT BUTONU
if st.button("🚀 Akıllı Yapay Zeka Analizini Başlat", use_container_width=True):
    if finans_file and prod_file and maliyet_file:
        with st.spinner("Yapay zeka motoru finansal röntgeninizi çıkarıyor..."):
            try:
                # Verileri Oku
                df_finans = pd.read_excel(finans_file)
                df_prod = pd.read_excel(prod_file, skiprows=1)
                df_maliyet = pd.read_excel(maliyet_file)
                
                # Sütun İsimlerini Temizle
                df_maliyet.columns = [c.strip() for c in df_maliyet.columns]
                df_finans.columns = [c.strip() for c in df_finans.columns]
                df_prod.columns = [c.strip() for c in df_prod.columns]
                
                # Finans Verilerini Haritala
                finans_dict = {}
                for idx, row in df_finans.iterrows():
                    s_no = str(row['Sipariş No']).strip()
                    finans_dict[s_no] = {
                        'toplam_adet': row['Ürün Adedi'],
                        'komisyon': row['Komisyon/Yurt Dışı Stok Destek Bedeli'],
                        'kargo': row['Gönderi Kargo Bedeli'],
                        'hizmet': row['Platform Hizmet Bedeli']
                    }
                    
                # Maliyet Haritası
                maliyet_dict = dict(zip(df_maliyet['TRENDYOL BARKOD'].astype(str).str.strip(), df_maliyet['TOPLAM MALİYET']))
                
                # Hesaplama Döngüsü
                sonuc_listesi = []
                eksik_maliyetler = set()
                toplam_iade_kargo_zarari = 0
                
                for idx, row in df_prod.iterrows():
                    barkod = str(row['Barkod']).strip()
                    siparis_no = str(row['Sipariş Numarası']).strip()
                    urun_adi = row['Ürün Adı']
                    marka = row['Marka']
                    statü = str(row.get('Statü', row.get('Sipariş Durumu', ''))).strip().lower()
                    
                    adet = float(row['Adet']) if pd.notnull(row['Adet']) else 0
                    satis_tutari = float(row['Satış Tutarı']) if pd.notnull(row['Satış Tutarı']) else 0
                    
                    if barkod == 'nan' or siparis_no == 'nan':
                        continue
                    
                    if "iptal" in statü or "reddedildi" in statü:
                        continue
                    
                    if barkod in maliyet_dict:
                        birim_maliyet = maliyet_dict[barkod]
                    else:
                        birim_maliyet = 0
                        eksik_maliyetler.add((barkod, urun_adi))
                    
                    toplam_maliyet = birim_maliyet * adet
                    is_iade = "iade" in statü
                    
                    if is_iade:
                        hesaplanan_ciro = 0
                        hesaplanan_maliyet = 0
                    else:
                        hesaplanan_ciro = satis_tutari
                        hesaplanan_maliyet = toplam_maliyet
                    
                    bolunmus_komisyon = 0
                    bolunmus_kargo = 0
                    bolunmus_hizmet = 0
                    
                    if siparis_no in finans_dict:
                        f_data = finans_dict[siparis_no]
                        if f_data['toplam_adet'] > 0:
                            bolunmus_komisyon = (f_data['komisyon'] / f_data['toplam_adet']) * adet
                            bolunmus_kargo = (f_data['kargo'] / f_data['toplam_adet']) * adet
                            bolunmus_hizmet = (f_data['hizmet'] / f_data['toplam_adet']) * adet
                            
                            if is_iade:
                                toplam_iade_kargo_zarari += abs(bolunmus_kargo)
                    
                    net_kar = hesaplanan_ciro + bolunmus_komisyon + bolunmus_kargo + bolunmus_hizmet - hesaplanan_maliyet
                    
                    sonuc_listesi.append({
                        "Sipariş No": siparis_no,
                        "Barkod": barkod,
                        "Marka": marka,
                        "Ürün Adı": urun_adi,
                        "Durum": "İade" if is_iade else "Satış",
                        "Satış Adedi": adet,
                        "Net Ciro": hesaplanan_ciro,
                        "Komisyon": bolunmus_komisyon,
                        "Kargo": bolunmus_kargo,
                        "Hizmet Bedeli": bolunmus_hizmet,
                        "Ürün Maliyeti": hesaplanan_maliyet,
                        "Net Kâr": net_kar
                    })
                    
                # Sonuçları Session State'e kaydet (Arama yaparken sıfırlanmasın)
                st.session_state['df_sonuc'] = pd.DataFrame(sonuc_listesi)
                st.session_state['toplam_ciro'] = toplam_ciro
                st.session_state['genel_net_kar'] = genel_net_kar
                st.session_state['iade_orani'] = iade_orani
                st.session_state['toplam_iade_kargo_zarari'] = toplam_iade_kargo_zarari
                st.session_state['reklam_gideri'] = reklam_gideri
                st.session_state['eksik_maliyetler'] = eksik_maliyetler
                st.session_state['toplam_maliyet_gideri'] = toplam_maliyet_gideri
                st.session_state['toplam_komisyon'] = toplam_komisyon
                st.session_state['toplam_kargo'] = toplam_kargo
                st.session_state['toplam_hizmet'] = toplam_hizmet
                st.success("✅ Analiz başarıyla tamamlandı! Aşağıdaki panelleri inceleyebilirsiniz.")
                
            except Exception as e:
                st.error(f"📊 Veri İşleme Hatası: {str(e)}")

# EĞER ANALİZ YAPILDIYSA EKRANA PANELLERİ GETİR
if 'df_sonuc' in st.session_state:
    df_sonuc = st.session_state['df_sonuc']
    
    # Özet Kartları
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("💰 Net Ciro (İadeler Hariç)", f"₺{st.session_state['toplam_ciro']:,.2f}")
    m2.metric("🟢 Gerçek Net Kâr", f"₺{st.session_state['genel_net_kar']:,.2f}")
    m3.metric("📦 Genel İade Oranı", f"%{st.session_state['iade_orani']:.2f}")
    m4.metric("🚨 İade Kargo Zararı", f"₺{st.session_state['toplam_iade_kargo_zarari']:,.2f}")
    
    # 🎯 YENİ EKSTRE ÖZELLİK: TEKLİ BARKOD YAZDIRMA İSTASYONU
    st.write("---")
    st.subheader("🎯 Tekli Barkod Arama ve Yazdırma İstasyonu")
    st.markdown("Aşağıdaki kutuya yazdırılacak barkodu girin. Sistem sadece o ürünü termal boyuta getirecektir.")
    
    # Benzersiz barkod listesini alıp arama kutusu yapalım
    barkod_listesi = ["Seçiniz..."] + list(df_sonuc['Barkod'].unique())
    aranan_barkod = st.selectbox("🔎 Yazdırılacak Barkodu Seçin veya Yazın:", barkod_listesi)
    
    if aranan_barkod != "Seçiniz...":
        # Sadece o barkoda ait bilgiyi çek
        urun_bilgi = df_sonuc[df_sonuc['Barkod'] == aranan_barkod].iloc[0]
        
        # Sadece Termal Yazıcı Şablonu (Kutucuk şeklinde şık tasarım)
        st.write("### 🖨️ Yazıcı Çıktı Önizlemesi")
        
        # Bu alan tam olarak 100x100 etiket formatında görünecek
        st.markdown(f"""
        <div style="border: 3px solid black; padding:
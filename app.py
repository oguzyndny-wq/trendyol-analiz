import streamlit as st
import pandas as pd
import plotly.express as px
import io

# Sayfa Genişlik ve Başlık Ayarları
st.set_page_config(page_title="Trendyol Akıllı Yapay Zeka Paneli v6.0", layout="wide")

st.title("🤖 Trendyol Akıllı Yapay Zeka Paneli v6.0 (Gelişmiş Finans)")
st.markdown("Fiyat tavsiyeleri, reklam motoru, tekli barkod istasyonu ve detaylı finansal metrikler entegre edilmiştir.")
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
                toplam_iptal_iade_adedi = 0
                
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
                    
                    # İptal ve İade adetlerini burada sayıyoruz
                    if "iptal" in statü or "iade" in statü or "reddedildi" in statü:
                        toplam_iptal_iade_adedi += adet
                    
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
                    
                df_final = pd.DataFrame(sonuc_listesi)
                
                # METRİKLER
                toplam_ciro = df_final['Net Ciro'].sum()
                toplam_kargo = df_final['Kargo'].sum()
                toplam_komisyon = df_final['Komisyon'].sum()
                toplam_hizmet = df_final['Hizmet Bedeli'].sum()
                toplam_maliyet_gideri = df_final['Ürün Maliyeti'].sum()
                genel_net_kar = df_final['Net Kâr'].sum() - reklam_gideri
                
                # Trendyol Kesintileri (Komisyon + Kargo + Hizmet toplamı)
                trendyol_kesintileri = toplam_komisyon + toplam_kargo + toplam_hizmet
                # Trendyol'dan Gelen Para (Ciro - Kesintiler) -> Kesintiler eksi değerli geldiği için topluyoruz
                trendyoldan_gelen_para = toplam_ciro + trendyol_kesintileri
                
                # Kâr Marjı (%) -> (Net Kâr / Net Ciro) * 100
                kar_marji = (genel_net_kar / toplam_ciro) * 100 if toplam_ciro > 0 else 0
                
                # Verileri Hafızaya Al (Session State)
                st.session_state['df_sonuc'] = df_final
                st.session_state['toplam_ciro'] = toplam_ciro
                st.session_state['trendyoldan_gelen_para'] = trendyoldan_gelen_para
                st.session_state['trendyol_kesintileri'] = abs(trendyol_kesintileri)
                st.session_state['genel_net_kar'] = genel_net_kar
                st.session_state['kar_marji'] = kar_marji
                st.session_state['toplam_iptal_iade_adedi'] = toplam_iptal_iade_adedi
                
                st.session_state['reklam_gideri'] = reklam_gideri
                st.session_state['eksik_maliyetler'] = eksik_maliyetler
                st.session_state['toplam_maliyet_gideri'] = toplam_maliyet_gideri
                st.session_state['toplam_komisyon'] = toplam_komisyon
                st.session_state['toplam_kargo'] = toplam_kargo
                st.session_state['toplam_hizmet'] = toplam_hizmet
                st.success("✅ Analiz başarıyla tamamlandı! Finansal göstergeler yenilendi.")
                
            except Exception as e:
                st.error(f"📊 Veri İşleme Hatası: {str(e)}")

# EĞER HESAPLAMA YAPILDIYSA EKRANI ÇİZ
if 'df_sonuc' in st.session_state:
    df_sonuc = st.session_state['df_sonuc']
    
    # Yeni İstediğiniz Gelişmiş Finans Kartları (2 Satır Halinde Muazzam Düzen)
    st.write("### 💵 Detaylı Finansal Durum Paneli")
    row1_col1, row1_col2, row1_col3 = st.columns(3)
    row1_col1.metric("💰 Net Ciro (İadeler Hariç)", f"₺{st.session_state['toplam_ciro']:,.2f}")
    row1_col2.metric("🏦 Trendyol'dan Gelen Para (Hakediş)", f"₺{st.session_state['trendyoldan_gelen_para']:,.2f}", help="Trendyol kesintileri çıktıktan sonra banka hesabınıza giren para.")
    row1_col3.metric("✂️ Toplam Trendyol Kesintisi", f"₺{st.session_state['trendyol_kesintileri']:,.2f}", help="Komisyon, Kargo ve Hizmet bedellerinin toplamı.")
    
    st.write(" ")
    row2_col1, row2_col2, row2_col3 = st.columns(3)
    row2_col1.metric("🟢 Gerçek Net Kâr (Reklam Dahil)", f"₺{st.session_state['genel_net_kar']:,.2f}")
    row2_col2.metric("📈 Net Kâr Marjı (%)", f"%{st.session_state['kar_marji']:.2f}", help="Elde ettiğiniz cironun yüzde kaçının net kâr olarak cebinizde kaldığını gösterir.")
    row2_col3.metric("🚨 Toplam İptal / İade Adedi", f"{int(st.session_state['toplam_iptal_iade_adedi'])} Adet", help="Rapor dönemindeki toplam iptal edilen ve iade gelen ürün adedi.")
    
    # 🔎 TEKLİ BARKOD YAZDIRMA İSTASYONU
    st.write("---")
    st.subheader("🎯 Tekli Barkod Arama ve Yazdırma İstasyonu")
    
    barkod_listesi = ["Seçiniz..."] + list(df_sonuc['Barkod'].unique())
    aranan_barkod = st.selectbox("🔎 Yazdırılacak Barkodu Seçin veya Yazın:", barkod_listesi)
    
    if aranan_barkod != "Seçiniz...":
        df_filtre = df_sonuc[df_sonuc['Barkod'] == aranan_barkod]
        if not df_filtre.empty:
            urun_bilgi = df_filtre.iloc[0]
            st.write("### 🖨️ Yazıcı Çıktı Önizlemesi")
            
            m_ad = str(urun_bilgi['Marka']).upper()
            u_ad = str(urun_bilgi['Ürün Adı'])[:60]
            b_no = str(aranan_barkod)
            
            html_sablon = '<div style="border: 3px solid black; padding: 20px; width: 350px; background-color: white; color: black; font-family: Arial; border-radius: 5px;">'
            html_sablon += '<h2 style="margin: 0; padding-bottom: 5px; border-bottom: 2px solid black;">' + m_ad + '</h2>'
            html_sablon += '<p style="font-size: 14px; margin: 10px 0;"><b>Ürün:</b> ' + u_ad + '</p>'
            html_sablon += '<div style="background-color: black; color: white; text-align: center; padding: 15px; font-size: 24px; font-weight: bold; letter-spacing: 5px; margin-top: 20px;">'
            html_sablon += '|||| ' + b_no + ' ||||</div>'
            html_sablon += '<p style="text-align: center; font-size: 12px; margin: 5px 0 0 0;">Barkod No: ' + b_no + '</p></div>'
            
            st.markdown(html_sablon, unsafe_allow_html=True)
            st.info("💡 Bu tekli barkodu yazdırmak için bilgisayarınızdan CTRL + P tuşlarına basın. Yazıcı ayarlarından Yalnızca Seçimi Yazdır seçeneğini işaretleyerek doğrudan termal etiket çıkartabilirsiniz!")

    # 💡 YAPAY ZEKA ZAM TAVSİYELERİ
    st.write("---")
    st.subheader("💡 Yapay Zeka Akıllı Fiyatlandırma ve Zam Tavsiyeleri")
    
    df_urun_analiz = df_sonuc.groupby(['Barkod', 'Marka', 'Ürün Adı']).agg({
        'Net Ciro': 'sum',
        'Net Kâr': 'sum',
        'Satış Adedi': 'sum',
        'Komisyon': 'sum',
        'Kargo': 'sum'
    }).reset_index()
    
    tavsiye_listesi = []
    for idx, row in df_urun_analiz.iterrows():
        if row['Satış Adedi'] > 0:
            birim_ciro = row['Net Ciro'] / row['Satış Adedi']
            birim_kar = row['Net Kâr'] / row['Satış Adedi']
            birim_kesinti_orani = abs(row['Komisyon'] + row['Kargo']) / row['Net Ciro'] if row['Net Ciro'] > 0 else 0.30
            
            if birim_kar < 0:
                gerekli_fiyat_artisi = abs(birim_kar) * (1 + birim_kesinti_orani)
                tavsiye_listesi.append({
                    "Barkod": row['Barkod'],
                    "Marka": row['Marka'],
                    "Ürün Adı": row['Ürün Adı'],
                    "Mevcut Birim Kâr": f"₺{birim_kar:.2f}",
                    "Durum": "🔴 ZARAR EDİYOR",
                    "AI Tavsiyesi": f"Trendyol Satış Fiyatını En Az ₺{gerekli_fiyat_artisi:.2f} ARTIRMALISINIZ!"
                })
    
    if tavsiye_listesi:
        st.dataframe(pd.DataFrame(tavsiye_listesi), use_container_width=True)
    else:
        st.success("✅ Harika! Bu ay zarar eden hiçbir ürününüz bulunmuyor.")

    # 🍕 GİDER PASTASI GRAFİĞİ
    st.write("---")
    st.write("### 🍕 Toplam Cironun Gider Dağılım Röntgeni")
    gider_data = {
        "Gider Kalemi": ["Net Kâr", "Ürün Maliyetleri", "Trendyol Komisyonu", "Kargo Giderleri", "Platform Hiz & Reklam"],
        "Tutar": [
            max(0, st.session_state['genel_net_kar']),
            st.session_state['toplam_maliyet_gideri'],
            abs(st.session_state['toplam_komisyon']),
            abs(st.session_state['toplam_kargo']),
            abs(st.session_state['toplam_hizmet']) + st.session_state['reklam_gideri']
        ]
    }
    fig_pie = px.pie(pd.DataFrame(gider_data), values='Tutar', names='Gider Kalemi', title="Cironuzun Dağılımı (%)", hole=0.4)
    st.plotly_chart(fig_pie, use_container_width=True)

    # 📥 EXCEL İNDİRME BUTONU
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_sonuc.to_excel(writer, index=False, sheet_name='AI Pro Raporu')
    
    st.write("---")
    st.download_button(
        label="📥 Tüm Sonuçları Detaylı Excel Olarak İndir",
        data=output.getvalue(),
        file_name="Trendyol_AI_Analizi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
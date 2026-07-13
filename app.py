import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="E-Ticaret Konsolide Paneli v11.0", layout="wide")
st.title("🤖 Çok Kanallı E-Ticaret Konsolide Finans Paneli v11.0")
st.markdown("Uzun satırlar ve if kesilme hataları tamamen giderilmiştir. Veriler dosyalardan anlık hesaplanır.")
st.write("---")

# Session State Hazırlığı (KeyError Koruması)
varsayilanlar = {
    'ty_ciro': 0.0, 'ty_kesinti': 0.0, 'ty_maliyet': 0.0, 'ty_kar': 0.0, 'ty_sip_adet': 0, 'ty_urun_adet': 0,
    'amz_ciro': 0.0, 'amz_kesinti': 0.0, 'amz_maliyet': 0.0, 'amz_kar': 0.0, 'amz_sip_adet': 0, 'amz_urun_adet': 0,
    'genel_ciro': 0.0, 'genel_maliyet': 0.0, 'genel_kar': 0.0, 'genel_sip_adet': 0, 'genel_urun_adet': 0, 'genel_marj': 0.0,
    'hesaplandi': False
}
for k, v in varsayilanlar.items():
    if k not in st.session_state: st.session_state[k] = v

tab_yükleme, tab_rapor = st.tabs(["📥 Veri Yükleme İstasyonu", "📊 Konsolide Finans & Raporlar"])

with tab_yükleme:
    st.subheader("1. Trendyol Raporları")
    col1, col2, col3 = st.columns(3)
    with col1: finans_file = st.file_uploader("SiparisKayitlari ile başlayan dosya", type=["xlsx", "xls"], key="finans")
    with col2: prod_file = st.file_uploader("prod_ ile başlayan dosya", type=["xlsx", "xls"], key="prod")
    with col3: maliyet_file = st.file_uploader("Trendyol Maliyet listesi", type=["xlsx", "xls"], key="maliyet")
    ty_reklam = st.number_input("🔗 Varsa Trendyol Ekstra Reklam Gideri (TL):", min_value=0.0, value=0.0, step=100.0)
    st.write("---")
    st.subheader("2. Amazon Raporları")
    col_amz1, col_amz2 = st.columns(2)
    with col_amz1: amazon_file = st.file_uploader("Haziran Amazon Raporu", type=["xlsx", "xls", "csv"], key="amazon_sales")
    with col_amz2: amazon_maliyet_file = st.file_uploader("Amazon Maliyet Şablonu", type=["xlsx", "xls", "csv"], key="amazon_cost")
    st.write("---")
    baslat_btn = st.button("🚀 Tüm Pazaryerlerinin Akıllı Analizini Başlat", use_container_width=True)

def clean_number(val):
    if pd.isnull(val): return 0.0
    if isinstance(
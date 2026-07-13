import streamlit as st
import pandas as pd
import plotly.express as px
import io

# Sayfa Yapılandırması
st.set_page_config(page_title="Konsolide E-Ticaret Paneli v10.0", layout="wide")

st.title("🤖 Çok Kanallı E-Ticaret Konsolide Finans Paneli v10.0")
st.markdown("Karakter sınırı ve satır kırılma korumalı tam stabil nihai versiyon.")
st.write("---")

# SEKME SİSTEMİ
tab_yükleme, tab_rapor = st.tabs(["📥 Veri Yükleme İstasyonu", "📊 Konsolide Finans & Raporlar"])

with tab_yükleme:
    st.subheader("1. Trendyol Raporları")
    col1, col2, col3 = st.columns(3)
    with col1:
        finans_file = st.file_uploader("SiparisKayitlari ile başlayan Trendyol dosyası", type=["xlsx", "xls"], key="finans")
    with col2:
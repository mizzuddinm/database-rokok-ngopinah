import streamlit as st
import pandas as pd

# Masukkan link Google Sheets asli Anda
sheet_url = "https://docs.google.com/spreadsheets/d/ID_SPREADSHEET_ANDA/edit#gid=0"

# Ubah format link agar menjadi CSV (kunci agar bisa dibaca pandas)
csv_url = sheet_url.replace("/edit#gid=", "/export?format=csv&gid=")

st.title("Database Rokok Ngopinah")

@st.cache_data(ttl=600)
def load_data():
    return pd.read_csv(csv_url)

df = load_data()
st.dataframe(df)

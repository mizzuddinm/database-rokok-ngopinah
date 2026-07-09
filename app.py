import streamlit as st
import pandas as pd

# Link Spreadsheet Anda
sheet_url = "https://docs.google.com/spreadsheets/d/1XwYfBSF4fULdXJFo3TJsJ78JBfANmNtcw8kajWV4nK0/edit#gid=0"

# Ubah link agar bisa diunduh sebagai CSV
csv_url = sheet_url.replace("/edit#gid=", "/export?format=csv&gid=")

st.title("Database Rokok Ngopinah")

@st.cache_data(ttl=600)
def load_data():
    return pd.read_csv(csv_url)

# Menampilkan data
df = load_data()
st.write("Data Berhasil Dimuat:")
st.dataframe(df)

import streamlit as st
import pandas as pd

# Konfigurasi halaman
st.set_page_config(page_title="Database Rokok", layout="wide")

st.title("📦 Database Rokok Ngopinah")

# Link Spreadsheet Anda
SHEET_URL = "https://docs.google.com/spreadsheets/d/1XwYfBSF4fULdXJFo3TJsJ78JBfANmNtcw8kajWV4nK0/edit#gid=0"

# Fungsi untuk sinkronisasi/mengambil data
@st.cache_data(ttl=600) # Data akan diperbarui setiap 10 menit
def sync_data_from_gsheet(url):
    # Mengubah link agar menjadi format CSV agar mudah dibaca pandas
    csv_url = url.replace("/edit#gid=", "/export?format=csv&gid=")
    try:
        df = pd.read_csv(csv_url)
        return df
    except Exception as e:
        st.error(f"Gagal sinkronisasi data: {e}")
        return None

# Eksekusi sinkronisasi
df = sync_data_from_gsheet(SHEET_URL)

if df is not None:
    st.success("Data berhasil disinkronkan!")
    st

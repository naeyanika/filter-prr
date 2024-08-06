import streamlit as st
import pandas as pd
import numpy as np
import io

st.title('Aplikasi Filter Pinjaman Renovasi Rumah')
st.write("""File yang dibutuhkan pivot_simpanan.xlsx dan KDP.xlsx yang sudah di gabung dan vlookup dengan KDP_na""")

uploaded_files = st.file_uploader("Unggah file Excel", accept_multiple_files=True, type=["xlsx"])

if uploaded_files:
    dfs = {}
    for file in uploaded_files:
        df = pd.read_excel(file, engine='openpyxl')  
        dfs[file.name] = df

if 'pivot_simpanan.xlsx' in dfs:
        df_s = dfs['pivot_simpanan.xlsx']
if 'KDP.xlsx' in dfs:
        df_kdp = dfs['KDP.xlsx']

# Filter KDP
df_filter_kdp = df_kdp[df_kdp['Cr PRR']>0].copy()

st.write ("KDP Filter")
st.write(df_filter_kdp)

# Vlookup
df_s_merged = pd.merge(df_s, df_filter_kdp, on=DUMMY, suffixes=('_s','_kdp'))
df_s_merged['Pencairan Renovasi Rumah x 25%'] = df_s_merged['Cr PRR'] * 0.25
df_s_merged['Simpanan Sesuai'] = df_s_merged.apply(lambda row: row['Db Sukarela'] >= row['Pencairan Renovasi Rumah x 25%'], axis=1)

result = df_s_merged[['DUMMY', 'NAMA_s', 'CENTER_s', 'KEL_s', 'HARI_s', 'JAM_s', 'SL_s', 'TRANS. DATE_s', 'Cr PRR', 'Db Sukarela', 'Cr Sukarela', 'Pencairan Renovasi Rumah x 25%', 'Simpanan Sesuai']]


rename_dict = {
    'NAMA_s': 'NAMA',
    'CENTER_s': 'CENTER',
    'KEL_s': 'KEL',
    'HARI_s': 'HARI',
    'JAM_s': 'JAM',
    'SL_s': 'SL',
    'Cr PRR': 'Pencairan Renovasi Rumah',
    'Db Sukarela': 'Simpanan Sukarela'
}

result = result.rename(columns=rename_dict)

desired_order = [
     'NAMA','CENTER','KELOMPOK','HARI','JAM','SL','Pencairan Renovasi Rumah','Simpanan Sukarela','Simpanan Sesuai'
]

for col in desired_order:
    if col not in result.columns:
         result[col] = 0

result = result[desired_order]

st.write('Hasil')
st.write(result)

# Download links for pivot tables
for name, df in {
    'Hasil Filter.xlsx': result
}.items():
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    buffer.seek(0)
    st.download_button(
        label=f"Unduh {name}",
        data=buffer.getvalue(),
        file_name=name,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
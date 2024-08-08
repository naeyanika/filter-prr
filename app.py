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

    if 'pivot_simpanan.xlsx' in dfs and 'KDP.xlsx' in dfs:
        df_s = dfs['pivot_simpanan.xlsx']
        df_kdp = dfs['KDP.xlsx']

# Filter KDP
cr_prr_column = [col for col in df_kdp.columns if 'cr prr' in col.lower()]
if cr_prr_column:
    df_filter_kdp = df_kdp[df_kdp[cr_prr_column[0]] > 0].copy()
    st.write("KDP Filter")
    st.write(df_filter_kdp)
else:
    st.error("Column 'Cr PRR' not found in KDP.xlsx. Please check the column names.")

# Vlookup
merge_column = 'DUMMY'  # Replace with your actual merge column name
df_s_merged = pd.merge(df_s, df_filter_kdp, on=merge_column, suffixes=('_df_s','_df_kdp'))
df_s_merged['Pencairan Renovasi Rumah x 25%'] = df_s_merged['Cr PRR'] * 0.25
df_s_merged['Sukarela Sesuai'] = df_s_merged.apply(lambda row: row['Db Sukarela'] >= row['Pencairan Renovasi Rumah x 25%'], axis=1)

df_s_merged['Pencairan Renovasi Rumah x 1%'] = df_s_merged['Cr PRR'] * 0.01
df_s_merged['Wajib Sesuai'] = df_s_merged.apply(lambda row: row['Db Wajib'] < row['Pencairan Renovasi Rumah x 1%'], axis=1)

result = df_s_merged[['DUMMY', 'NAMA_df_s', 'CENTER_df_s', 'KEL_df_s', 'HARI_df_s', 'JAM_df_s', 'SL_df_s', 'TRANS. DATE_df_s', 'Cr PRR', 'Db Sukarela', 'Cr Sukarela', 'Db Wajib', 'Cr Wajib', 'Pencairan Renovasi Rumah x 1%', 'Pencairan Renovasi Rumah x 25%', 'Sukarela Sesuai', 'Wajib Sesuai']]


rename_dict = {
    'NAMA_df_s': 'NAMA',
    'CENTER_df_s': 'CENTER',
    'KEL_df_s': 'KEL',
    'HARI_df_s': 'HARI',
    'JAM_df_s': 'JAM',
    'SL_df_s': 'SL',
    'TRANS. DATE_df_s': 'TRANS. DATE',
    'Cr PRR': 'Pencairan Renovasi Rumah',
    'Db Sukarela': 'Simpanan Sukarela',
    'Db Wajib': 'Simpanan Wajib'
}

result = result.rename(columns=rename_dict)

desired_order = [
     'NAMA','CENTER','KEL','HARI','JAM','SL','TRANS. DATE','Pencairan Renovasi Rumah', 'Simpanan Wajib', 'Wajib Sesuai', 'Simpanan Sukarela','Sukarela Sesuai'
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

else:
    st.error("Pastikan Anda mengunggah kedua file: pivot_simpanan.xlsx dan KDP.xlsx")

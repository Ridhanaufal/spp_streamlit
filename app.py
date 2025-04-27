import streamlit as st
import pandas as pd
from io import BytesIO

# --- Fungsi hapus duplikat berdasarkan NIM ---
def hapus_duplikat_nama(df):
    df_sorted = df.sort_values(by=['nim', 'nama mahasiswa'])
    return df_sorted.drop_duplicates(subset='nim', keep='first')

# --- Konfigurasi Halaman ---
st.set_page_config(page_title="Sistem Pembayaran SPP", layout="wide")
st.title(":bar_chart: Sistem Pengolahan Data Pembayaran SPP Mahasiswa INSTIPER")
st.caption("Created by: Admin IT KEUANGAN")

# --- Upload File ---
with st.container():
    uploaded_file = st.file_uploader(":file_folder: Unggah File Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success(f"File **{uploaded_file.name}** berhasil diunggah!")

    with st.expander(":clipboard: Lihat Data Asli"):
        st.write(":pushpin: Kolom ditemukan:", df.columns.tolist())
        df.columns = df.columns.str.strip().str.lower()
        df["nim"] = df["nim"].astype(str).str.strip()
        df["nama mahasiswa"] = df["nama mahasiswa"].astype(str).str.strip().str.title()
        st.dataframe(df, use_container_width=True)

    # --- Pilih Tahun Akademik ---
    if "tahun akademik" in df.columns:
        tahun_tersedia = sorted(df["tahun akademik"].dropna().unique())
        tahun_dipilih = st.multiselect(":dart: Pilih Tahun Akademik", tahun_tersedia, default=tahun_tersedia)

        if tahun_dipilih:
            if st.button(":rocket: Proses Data"):
                with st.spinner('Sedang memproses data...'):
                    required_cols = ["nim", "nama mahasiswa", "jenis tagihan", "tahun akademik", "nominal"]
                    if not all(col in df.columns for col in required_cols):
                        st.error(f"Kolom wajib tidak ditemukan: {required_cols}")
                    else:
                        jenis_valid = ["spp", "spp t", "spp angsuran 2", "spp tetap", "spp c"]
                        df_spp = df[df["jenis tagihan"].str.lower().isin(jenis_valid)]
                        df_spp = df_spp[df_spp["tahun akademik"].isin(tahun_dipilih)]
                        df_spp["nominal"] = df_spp["nominal"].fillna(0)

                        # Ambil nominal terbesar per nim + tahun akademik
                        df_spp_max = df_spp.groupby(['nim', 'tahun akademik'], as_index=False)['nominal'].max()

                        # Ambil nama mahasiswa berdasarkan NIM (hapus duplikat NIM)
                        nama_mahasiswa_df = hapus_duplikat_nama(df[['nim', 'nama mahasiswa']])

                        # Gabungkan nama dengan data SPP
                        df_final = df_spp_max.merge(nama_mahasiswa_df, on='nim', how='left')

                        # Buat pivot
                        pivot = df_final.pivot_table(
                            index=["nim", "nama mahasiswa"],
                            columns="tahun akademik",
                            values="nominal",
                            aggfunc="sum"
                        ).fillna(0).reset_index()

                        tahun_cols = [col for col in pivot.columns if col not in ["nim", "nama mahasiswa"]]

                        def hitung_status(row):
                            total_bayar = row[tahun_cols].sum()
                            lunas = all(row[tahun] > 900_000 for tahun in tahun_cols)
                            return pd.Series({
                                "Jumlah Tahun Akademik": len(tahun_cols),
                                "Total Bayar": total_bayar,
                                "Status": "Sudah Membayar" if lunas else "Belum Lunas"
                            })

                        hasil_status = pivot.apply(hitung_status, axis=1)
                        df_hasil = pd.concat([pivot, hasil_status], axis=1)

                        tabs = st.tabs([":memo: Data Status", ":bar_chart: Rekapitulasi", ":inbox_tray: Unduh Data"])

                        with tabs[0]:
                            st.subheader(":bar_chart: Status Pembayaran Mahasiswa")
                            st.dataframe(df_hasil, use_container_width=True)

                        with tabs[1]:
                            if "jurusan" in df.columns:
                                df_join = df[['nim', 'jurusan']].drop_duplicates(subset='nim')
                                df_join = df_join.merge(df_hasil, on='nim', how='right')

                                rekap = df_join.groupby(['jurusan', 'Status']).size().unstack(fill_value=0)
                                rekap['Total'] = rekap.sum(axis=1)
                                rekap['% Sudah Membayar'] = (rekap.get('Sudah Membayar', 0) / rekap['Total'] * 100).round(2)
                                rekap['% Belum Lunas'] = (rekap.get('Belum Lunas', 0) / rekap['Total'] * 100).round(2)
                                rekap.reset_index(inplace=True)

                                st.subheader(":chart_with_upwards_trend: Rekap Status per Jurusan")
                                st.dataframe(rekap, use_container_width=True)
                            else:
                                st.warning("Kolom 'jurusan' tidak ditemukan!")

                        with tabs[2]:
                            df_export = df_hasil.copy()
                            for kolom_tambahan in ["jurusan", "minat", "fakultas"]:
                                if kolom_tambahan in df.columns:
                                    tambahan_df = df[["nim", kolom_tambahan]].drop_duplicates(subset='nim')
                                    df_export = df_export.merge(tambahan_df, on="nim", how="left")

                            hasil_xlsx = BytesIO()
                            with pd.ExcelWriter(hasil_xlsx, engine='xlsxwriter') as writer:
                                df_export.to_excel(writer, index=False, sheet_name='Status Mahasiswa')

                            st.download_button(
                                label=":inbox_tray: Unduh Status Mahasiswa ke Excel",
                                data=hasil_xlsx.getvalue(),
                                file_name="status_pembayaran_mahasiswa.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )

                            if "jurusan" in df.columns:
                                output_xlsx = BytesIO()
                                with pd.ExcelWriter(output_xlsx, engine='xlsxwriter') as writer:
                                    rekap.to_excel(writer, index=False, sheet_name='Rekap Jurusan')

                                st.download_button(
                                    label=":inbox_tray: Unduh Rekap Jurusan ke Excel",
                                    data=output_xlsx.getvalue(),
                                    file_name="rekap_jurusan.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )

        else:
            st.warning("⚠️ Silakan pilih minimal satu Tahun Akademik terlebih dahulu.")

else:
    st.info("⬆️ Silakan upload file Excel terlebih dahulu.")
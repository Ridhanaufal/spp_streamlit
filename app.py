import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="Aplikasi Statistik Data Pembayaran Mahasiswa INSTIPER", page_icon="📊", layout="centered")

st.markdown(
    """
    <style>
    .main {
        background-color: #f8fafc;
    }
    .stButton>button {
        color: white;
        background: #4f8bf9;
    }
    .stFileUploader>div>div {
        background: #e3eafc;
        border-radius: 8px;
        padding: 10px;
    }
    /* Border dan padding untuk dataframe */
    .stDataFrame div[data-testid="stHorizontalBlock"] {
        border: 1.5px solid #4f8bf9;
        border-radius: 10px;
        padding: 10px 5px 5px 5px;
        background-color: #fafdff;
        box-shadow: 0 2px 8px rgba(79,139,249,0.06);
    }
    /* Header tabel lebih tebal dan warna */
    .stDataFrame th {
        background-color: #e3eafc !important;
        color: #2a3f5f !important;
        font-weight: bold !important;
        border-bottom: 2px solid #4f8bf9 !important;
    }
    /* Baris tabel lebih rapi */
    .stDataFrame td {
        border-bottom: 1px solid #e3eafc !important;
        padding: 6px 8px !important;
    }
    /* Scrollbar lebih halus */
    ::-webkit-scrollbar-thumb {
        background: #b3cdfd;
        border-radius: 8px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown("<h1 style='text-align: center; color: #4f8bf9;'>📊 Aplikasi Statistik Data Pembayaran Mahasiswa INSTIPer</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>Upload file Excel Anda untuk melihat dan mengolah data pembayaran mahasiswa secara interaktif.</p>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #888;'>Create by Ridha</p>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.success("File berhasil diupload!")

    # Filter Tahun Akademik dengan multiselect jika kolom tersedia
    if 'Tahun Akademik' in df.columns:
        tahun_list = sorted(df['Tahun Akademik'].dropna().unique())
        tahun_terpilih = st.multiselect("Pilih Tahun Akademik", tahun_list, default=tahun_list)
        df = df[df['Tahun Akademik'].isin(tahun_terpilih)]

    # Membuat kolom Status SPP
    spp_types = ['SPP', 'SPP T', 'SPP Variable', 'SPP Tetap', 'SPP C']

    def cek_status_spp(sub_df):
        spp_rows = sub_df[sub_df['Jenis Tagihan'].isin(spp_types)]
        tahun_ada = spp_rows.groupby('Tahun Akademik')['Nominal'].sum() != 0
        if tahun_ada.all() and len(tahun_ada) == len(tahun_terpilih):
            return "Lunas SPP"
        else:
            return "Belum Lunas"

    if {'NIM', 'Tahun Akademik', 'Jenis Tagihan', 'Nominal'}.issubset(df.columns):
        status_spp = df.groupby('NIM').apply(cek_status_spp).reset_index()
        status_spp.columns = ['NIM', 'Status SPP']
        df = df.merge(status_spp, on='NIM', how='left')
    else:
        st.warning("Kolom NIM, Tahun Akademik, Jenis Tagihan, dan Nominal harus ada di file Excel.")

    # Pastikan Tahun Akademik bertipe string agar pivot konsisten
    df['Tahun Akademik'] = df['Tahun Akademik'].astype(str)

    # Filter hanya baris SPP
    spp_df = df[df['Jenis Tagihan'].isin(spp_types)].copy()

    # Normalisasi Nama Mahasiswa dan Jurusan untuk setiap NIM (ambil data pertama)
    identitas = spp_df.groupby('NIM')[['Nama Mahasiswa', 'Jurusan']].first().reset_index()
    spp_df = spp_df.drop(['Nama Mahasiswa', 'Jurusan'], axis=1).merge(identitas, on='NIM', how='left')

    # Untuk setiap NIM, Tahun Akademik, ambil nominal terbesar (bukan 0 jika ada)
    spp_df['Nominal'] = spp_df.groupby(['NIM', 'Tahun Akademik'])['Nominal'].transform(lambda x: x[x > 0].max() if (x > 0).any() else 0)
    spp_df = spp_df.drop_duplicates(subset=['NIM', 'Tahun Akademik'])

    spp_pivot = spp_df.pivot_table(
        index=['NIM', 'Nama Mahasiswa', 'Jurusan'],
        columns='Tahun Akademik',
        values='Nominal',
        aggfunc='max'
    ).reset_index()

    # Gabungkan dengan status SPP (pastikan hanya satu status per NIM)
    if 'Status SPP' in df.columns:
        status_spp_unique = df[['NIM', 'Status SPP']].drop_duplicates(subset=['NIM'])
        spp_pivot = spp_pivot.merge(status_spp_unique, on='NIM', how='left')

    # Drop duplikat NIM jika masih ada (safety)
    spp_pivot = spp_pivot.drop_duplicates(subset=['NIM'])

    # Ganti NaN dengan 0 (benar-benar belum bayar)
    spp_pivot = spp_pivot.fillna(0)

    # Tambahkan kolom STD Terbayar
    std_types = [
        'Tri Dharma PT', 'Tri Dharma PT 2', 'Tri Dharma PT 3',
        'Tri Dharma PT 4', 'Tri Dharma PT 5', 'Tri Dharma PT 6', 'Tri Dharma PT 7'
    ]
    std_df = df[df['Jenis Tagihan'].isin(std_types)]
    std_total = std_df.groupby('NIM')['Nominal'].sum().reset_index().rename(columns={'Nominal': 'STD Terbayar'})
    spp_pivot = spp_pivot.merge(std_total, on='NIM', how='left')
    spp_pivot['STD Terbayar'] = spp_pivot['STD Terbayar'].fillna(0)

    # Tambahkan kolom Keterangan Cuti
    cuti_df = df[df['Jenis Tagihan'].str.lower().str.contains('cuti')]
    cuti_info = (
        cuti_df.groupby('NIM')['Tahun Akademik']
        .apply(lambda x: ', '.join(sorted(set(x))))
        .reset_index()
    )
    cuti_info['Keterangan'] = 'Pernah Cuti di tahun akademik ' + cuti_info['Tahun Akademik']
    cuti_info = cuti_info[['NIM', 'Keterangan']]
    spp_pivot = spp_pivot.merge(cuti_info, on='NIM', how='left')
    spp_pivot['Keterangan'] = spp_pivot['Keterangan'].fillna('')

    # Urutkan kolom: NIM, Nama Mahasiswa, Jurusan, Tahun Akademik (pivot), STD Terbayar, Status SPP, Keterangan
    cols = (
        ['NIM', 'Nama Mahasiswa', 'Jurusan']
        + [col for col in spp_pivot.columns if col not in ['NIM', 'Nama Mahasiswa', 'Jurusan', 'STD Terbayar', 'Status SPP', 'Keterangan']]
        + ['STD Terbayar', 'Status SPP', 'Keterangan']
    )
    spp_pivot = spp_pivot[cols]

    tahun_akademik_cols = [col for col in spp_pivot.columns if col not in ['NIM', 'Nama Mahasiswa', 'Jurusan', 'STD Terbayar', 'Status SPP', 'Keterangan']]

    tab1, tab2 = st.tabs(
        [
            "📋 Preview Data SPP Mahasiswa",
            "📊 Rekapitulasi per Jurusan"
        ]
    )

    with tab1:
        st.markdown("<h3 style='color:#4f8bf9;'>Preview Data SPP Mahasiswa</h3>", unsafe_allow_html=True)
        st.dataframe(
            spp_pivot,
            use_container_width=True,
            hide_index=True,
            height=500
        )
        st.markdown(
            f"<div style='color:#4f8bf9; text-align:right; font-size:16px;'><b>Jumlah mahasiswa: {spp_pivot.shape[0]}</b></div>",
            unsafe_allow_html=True
        )
        st.markdown("<br>", unsafe_allow_html=True)

        # Tombol export ke Excel dengan tambahan kolom Fakultas dan Minat
        spp_export = spp_pivot.copy()
        # Ambil data fakultas dan minat dari df asli (pastikan kolom ada)
        if 'Fakultas' in df.columns:
            fakultas_map = df.drop_duplicates('NIM').set_index('NIM')['Fakultas']
            spp_export['Fakultas'] = spp_export['NIM'].map(fakultas_map)
        else:
            spp_export['Fakultas'] = ''
        if 'Minat' in df.columns:
            minat_map = df.drop_duplicates('NIM').set_index('NIM')['Minat']
            spp_export['Minat'] = spp_export['NIM'].map(minat_map)
        else:
            spp_export['Minat'] = ''

        buffer = io.BytesIO()
        spp_export.to_excel(buffer, index=False)
        st.download_button(
            label="⬇️ Export ke Excel",
            data=buffer,
            file_name="Preview_Data_SPP_Mahasiswa.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Download data SPP Mahasiswa dalam format Excel"
        )

    with tab2:
        st.markdown("<h3 style='color:#4f8bf9;'>Rekapitulasi per Jurusan</h3>", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        jurusan_list = sorted(spp_pivot['Jurusan'].unique())
        jurusan_terpilih = st.multiselect(
            "Pilih Jurusan",
            jurusan_list,
            default=jurusan_list,
            help="Filter rekap berdasarkan jurusan"
        )
        spp_rekap = spp_pivot[spp_pivot['Jurusan'].isin(jurusan_terpilih)]

        rekap_status = spp_rekap.groupby(['Jurusan', 'Status SPP']).agg(
            Jumlah_Mahasiswa=('NIM', 'nunique')
        ).reset_index()

        rekap_status = rekap_status.pivot_table(
            index='Jurusan',
            columns='Status SPP',
            values='Jumlah_Mahasiswa',
            fill_value=0
        ).reset_index()

        if 'Lunas SPP' not in rekap_status.columns:
            rekap_status['Lunas SPP'] = 0
        if 'Belum Lunas' not in rekap_status.columns:
            rekap_status['Belum Lunas'] = 0

        rekap_status = rekap_status.rename(columns={
            'Lunas SPP': 'Sudah Membayar SPP',
            'Belum Lunas': 'Belum Membayar SPP'
        })

        st.dataframe(
            rekap_status,
            use_container_width=True,
            hide_index=True,
            height=400
        )

    # with tab3:
    #     st.markdown("<h3 style='color:#4f8bf9;'>Grafik Pembayaran Mahasiswa</h3>", unsafe_allow_html=True)
    #     st.markdown("<br>", unsafe_allow_html=True)

    #     # Pie chart: Proporsi Lunas/Belum Lunas per jurusan
    #     pie_df = spp_pivot.groupby(['Jurusan', 'Status SPP'])['NIM'].nunique().reset_index()
    #     fig_pie = px.pie(
    #         pie_df, 
    #         names='Status SPP', 
    #         values='NIM', 
    #         color='Status SPP',
    #         facet_col='Jurusan',
    #         title="Proporsi Mahasiswa Lunas/Belum Lunas per Jurusan"
    #     )
    #     st.plotly_chart(fig_pie, use_container_width=True)

    #     # Bar chart: Total SPP dan STD Terbayar per jurusan
    #     bar_df = spp_pivot.groupby('Jurusan').agg(
    #         Total_SPP=pd.NamedAgg(column=tahun_akademik_cols, aggfunc='sum'),
    #         Total_STD_Terbayar=pd.NamedAgg(column='STD Terbayar', aggfunc='sum')
    #     ).reset_index()
    #     bar_df['Total_SPP'] = bar_df[tahun_akademik_cols].sum(axis=1)
    #     fig_bar = px.bar(
    #         bar_df, 
    #         x='Jurusan', 
    #         y=['Total_SPP', 'Total_STD_Terbayar'],
    #         barmode='group',
    #         title="Total SPP & STD Terbayar per Jurusan"
    #     )
    #     st.plotly_chart(fig_bar, use_container_width=True)

    #     # Line chart: Tren pembayaran SPP per tahun akademik
    #     spp_trend = spp_pivot.melt(
    #         id_vars=['Jurusan'],
    #         value_vars=tahun_akademik_cols,
    #         var_name='Tahun Akademik',
    #         value_name='Nominal'
    #     )
    #     spp_trend = spp_trend.groupby(['Tahun Akademik']).agg({'Nominal': 'sum'}).reset_index()
    #     fig_line = px.line(
    #         spp_trend,
    #         x='Tahun Akademik',
    #         y='Nominal',
    #         markers=True,
    #         title="Tren Total Pembayaran SPP per Tahun Akademik"
    #     )
    #     st.plotly_chart(fig_line, use_container_width=True)

    # Validasi status lunas seperti sebelumnya
    def validasi_lunas(row):
        if row['Status SPP'] == "Lunas SPP":
            for col in tahun_akademik_cols:
                if row[col] == 0:
                    val = df[(df['NIM'] == row['NIM']) & (df['Tahun Akademik'] == col) & (df['Jenis Tagihan'].isin(spp_types))]['Nominal'].max()
                    if pd.notna(val) and val > 0:
                        row[col] = val
        return row

    spp_pivot = spp_pivot.apply(validasi_lunas, axis=1)
else:
    st.info("Silakan upload file Excel untuk melihat datanya.")
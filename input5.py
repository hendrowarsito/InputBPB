"""Aplikasi Streamlit input data BPB/BTB MAPPI.

Alur kerja: unggah lembar BTB (PDF cetakan portal MAPPI, atau XLSX/CSV) ->
periksa dan koreksi hasil bacaan -> simpan ke basis data XLSX -> unduh
kembali dalam format tabel atau format panjang.

Seluruh logika baca/tulis berada di :mod:`btb_io` agar bisa diuji tanpa
menjalankan Streamlit (lihat ``test_btb_io.py``).
"""

from __future__ import annotations

import os

import pandas as pd
import streamlit as st

import btb_io as btb

# Streamlit Cloud hanya mengizinkan tulis ke /tmp; setel BTB_DATA_DIR bila
# aplikasi dijalankan di server sendiri agar data tidak ikut terhapus.
DATA_DIR = os.environ.get("BTB_DATA_DIR", "/tmp/data_btb")
EXCEL_PATH = os.path.join(DATA_DIR, "btb_data.xlsx")
os.makedirs(DATA_DIR, exist_ok=True)

st.set_page_config(page_title="Input Data BPB MAPPI", page_icon="🏗️", layout="wide")
st.title("Aplikasi Input Data BPB MAPPI")
st.caption(
    "Baca lembar Biaya Teknis Bangunan (BTB) dari PDF portal MAPPI atau XLSX/CSV, "
    "lalu simpan dan unduh kembali dalam format XLSX."
)


@st.cache_data(show_spinner=False)
def _parse(file_bytes: bytes, filename: str):
    """Baca berkas unggahan. Di-cache agar tidak diulang tiap interaksi."""
    import io

    return btb.read_any(io.BytesIO(file_bytes), filename)


def _refresh() -> pd.DataFrame:
    return btb.load_workbook_data(EXCEL_PATH)


tab1, tab2, tab3 = st.tabs(["Input Data", "Data Telah Diinput", "Download Data"])

# ---------------------------------------------------------------------------
# Tab 1 - Input Data
# ---------------------------------------------------------------------------
with tab1:
    st.write(
        "Unggah lembar BTB MAPPI. Format yang didukung: **PDF** (hasil cetak "
        "halaman *BTB Interaktif*), **XLSX/XLSM**, dan **CSV**."
    )
    uploaded = st.file_uploader("Pilih berkas", type=["pdf", "xlsx", "xlsm", "csv"])

    if uploaded is None:
        st.info("Belum ada berkas yang diunggah.")
    else:
        file_bytes = uploaded.getvalue()
        try:
            with st.spinner(
                "Membaca berkas… PDF cetakan MAPPI tidak memiliki lapisan teks "
                "sehingga dibaca sel per sel melalui OCR (± 30 detik)."
            ):
                sheets = _parse(file_bytes, uploaded.name)
        except btb.BTBError as exc:
            st.error(str(exc))
            sheets = []
        except Exception as exc:  # noqa: BLE001 - tampilkan agar bisa dilaporkan
            st.error(f"Gagal membaca berkas: {exc}")
            sheets = []

        for index, sheet in enumerate(sheets):
            if len(sheets) > 1:
                st.divider()
                st.subheader(f"Lembar {index + 1}: {sheet.kota_kabupaten or '(tanpa nama)'}")

            for note in sheet.notes:
                st.warning(note)

            col1, col2, col3 = st.columns(3)
            provinsi = col1.text_input("Provinsi", value=sheet.provinsi, key=f"prov_{index}")
            kota = col2.text_input(
                "Kota/Kabupaten", value=sheet.kota_kabupaten, key=f"kota_{index}"
            )
            tahun = col3.number_input(
                "Tahun",
                min_value=1990,
                max_value=2100,
                value=int(sheet.tahun) if sheet.tahun else 2026,
                step=1,
                key=f"tahun_{index}",
            )

            st.markdown("**Periksa dan koreksi angka bila perlu** (satuan Rp/m²):")
            # baris judul kelompok tidak memuat angka, sembunyikan dari editor
            editable = sheet.values.drop(
                index=[label for label in btb.ROW_LABELS if btb.ROW_KIND[label] == "section"],
                errors="ignore",
            ).copy()
            editable.index.name = "Elemen"
            edited = st.data_editor(
                editable,
                use_container_width=True,
                key=f"editor_{index}",
                column_config={
                    column: st.column_config.NumberColumn(column, format="%d")
                    for column in btb.BUILDING_COLUMNS
                },
            )

            final = btb.BTBSheet(
                provinsi=provinsi.strip(),
                kota_kabupaten=kota.strip(),
                tahun=int(tahun),
                values=edited.reindex(btb.ROW_LABELS),
                source=uploaded.name,
            )

            issues = btb.validate(final)
            if issues:
                with st.expander(f"⚠️ {len(issues)} ketidaksesuaian angka ditemukan", expanded=True):
                    for issue in issues:
                        st.write("- " + issue)
                    st.caption(
                        "Total dan PPN pada dokumen tidak cocok dengan hasil hitung. "
                        "Koreksi angkanya di tabel di atas, atau hitung ulang total."
                    )
                if st.checkbox(
                    "Hitung ulang total, PPN 11%, dan pembulatan dari komponen biaya",
                    key=f"recalc_{index}",
                ):
                    btb.recompute_totals(final)
                    st.info("Baris total dihitung ulang — angka di bawah ini yang akan disimpan.")
                    st.dataframe(final.values, use_container_width=True)
            else:
                st.success("Validasi aritmatika lolos: total, PPN 11%, dan pembulatan konsisten.")

            if st.button("💾 Simpan ke Excel", type="primary", key=f"save_{index}"):
                try:
                    data = btb.save_sheet(EXCEL_PATH, final)
                except btb.BTBError as exc:
                    st.error(str(exc))
                else:
                    st.success(
                        f"Data {final.kota_kabupaten} tahun {final.tahun} tersimpan "
                        f"({len(data)} baris total dalam basis data)."
                    )
                    st.download_button(
                        "⬇️ Unduh basis data XLSX",
                        data=btb.build_download(data, "wide"),
                        file_name="btb_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_after_save_{index}",
                    )

# ---------------------------------------------------------------------------
# Tab 2 - Data Telah Diinput
# ---------------------------------------------------------------------------
with tab2:
    st.write("Daftar kota/kabupaten dan tahun yang telah tersimpan:")
    data = _refresh()
    if data.empty:
        st.info("Belum ada data yang tersedia.")
    else:
        st.dataframe(btb.summary(data), use_container_width=True, hide_index=True)
        with st.expander("Lihat seluruh data mentah"):
            st.dataframe(data, use_container_width=True, hide_index=True)

        st.divider()
        keys = (
            data[["Kota/Kabupaten", "Tahun"]]
            .drop_duplicates()
            .astype({"Tahun": "Int64"})
            .apply(lambda r: f"{r['Kota/Kabupaten']} — {r['Tahun']}", axis=1)
            .tolist()
        )
        target = st.selectbox("Hapus data:", keys, index=None, placeholder="Pilih kota dan tahun")
        if target and st.button("🗑️ Hapus", key="hapus"):
            kota, tahun = target.rsplit(" — ", 1)
            keep = ~(
                (data["Kota/Kabupaten"].astype(str) == kota)
                & (pd.to_numeric(data["Tahun"], errors="coerce") == int(tahun))
            )
            btb.write_workbook(EXCEL_PATH, data[keep].reset_index(drop=True))
            st.success(f"Data {target} dihapus.")
            st.rerun()

# ---------------------------------------------------------------------------
# Tab 3 - Download Data
# ---------------------------------------------------------------------------
with tab3:
    st.write("Pilih data BPB MAPPI yang akan diunduh.")
    data = _refresh()
    if data.empty:
        st.info("Belum ada data yang tersedia untuk diunduh.")
    else:
        col1, col2 = st.columns(2)
        kota_pilihan = col1.multiselect(
            "Kota/Kabupaten",
            sorted(data["Kota/Kabupaten"].dropna().astype(str).unique()),
        )
        tahun_pilihan = col2.multiselect(
            "Tahun",
            sorted(pd.to_numeric(data["Tahun"], errors="coerce").dropna().astype(int).unique()),
        )

        filtered = data
        if kota_pilihan:
            filtered = filtered[filtered["Kota/Kabupaten"].astype(str).isin(kota_pilihan)]
        if tahun_pilihan:
            filtered = filtered[
                pd.to_numeric(filtered["Tahun"], errors="coerce").isin(tahun_pilihan)
            ]

        # bentuk transpose dijadikan pilihan pertama sekaligus default
        pilihan_format = {
            "Transpose (Provinsi di baris 1)": "transpose",
            "Tabel (seperti dokumen asli)": "wide",
            "Panjang (siap pivot)": "long",
        }
        layout = pilihan_format[
            st.radio("Format berkas", list(pilihan_format), horizontal=True, index=0)
        ]

        if layout == "transpose":
            preview = btb.wide_to_transposed(filtered).reset_index()
            st.caption(
                "Keterangan turun di kolom pertama: Provinsi baris 1, "
                "Kota/Kabupaten baris 2, Tahun baris 3, lalu No, Kelompok, "
                "Elemen, dan satu baris untuk tiap tipe bangunan."
            )
        else:
            preview = filtered if layout == "wide" else btb.wide_to_long(filtered)
        st.dataframe(preview, use_container_width=True, hide_index=True)
        st.caption(f"{len(preview):,} baris × {len(preview.columns):,} kolom siap diunduh.")

        if filtered.empty:
            st.warning("Tidak ada baris yang cocok dengan filter.")
        else:
            suffix = "_".join(kota_pilihan) if kota_pilihan else "semua"
            tahun_suffix = "_".join(map(str, tahun_pilihan)) if tahun_pilihan else "semua"
            st.download_button(
                "⬇️ Unduh XLSX",
                data=btb.build_download(filtered, layout),
                file_name=f"BTB_{suffix}_{tahun_suffix}_{layout}.xlsx".replace(" ", "-"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

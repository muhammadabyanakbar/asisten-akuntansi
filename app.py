import streamlit as st

from ai_service import tanya_materi
from document_reader import baca_dokumen

st.title("Asisten Akuntansi")

files = st.file_uploader(
    "Unggah materi (PDF atau PPTX)",
    type=["pdf", "pptx"],
    accept_multiple_files=True,
)


@st.cache_data(show_spinner=False)
def ekstrak_materi(nama: str, data: bytes) -> str:
    return baca_dokumen(nama, data)


if files:
    bagian = []
    for file in files:
        bagian.append(f"=== Materi: {file.name} ===")
        bagian.append(ekstrak_materi(file.name, file.getvalue()))
    materi = "\n".join(bagian)
    st.success(f"{len(files)} materi berhasil dibaca.")

    pertanyaan = st.text_input("Tulis pertanyaan lu:")
    if st.button("Tanya") and pertanyaan:
        with st.spinner("Sedang menjawab..."):
            jawaban = tanya_materi(materi, pertanyaan)
        st.write(jawaban)

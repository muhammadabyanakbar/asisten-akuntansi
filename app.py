import streamlit as st

from ai_service import buat_chat, kirim
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


if not files:
    st.info("Unggah materi terlebih dahulu untuk memulai.")
    st.stop()

bagian = []
for file in files:
    bagian.append(f"=== Materi: {file.name} ===")
    bagian.append(ekstrak_materi(file.name, file.getvalue()))
materi = "\n".join(bagian)

# Jika materi berubah, mulai percakapan baru.
kunci = tuple((f.name, f.size) for f in files)
if st.session_state.get("kunci_materi") != kunci:
    st.session_state.kunci_materi = kunci
    st.session_state.chat = buat_chat()
    st.session_state.riwayat = []
    st.session_state.materi_terkirim = False

st.caption(f"{len(files)} materi aktif. Ganti atau tambah file untuk memulai percakapan baru.")

# Tampilkan riwayat percakapan.
for pesan in st.session_state.riwayat:
    with st.chat_message(pesan["role"]):
        st.markdown(pesan["isi"])

# Kotak chat di bagian bawah.
if pertanyaan := st.chat_input("Tanyakan sesuatu tentang materi..."):
    st.session_state.riwayat.append({"role": "user", "isi": pertanyaan})
    with st.chat_message("user"):
        st.markdown(pertanyaan)

    materi_awal = None if st.session_state.materi_terkirim else materi
    with st.chat_message("assistant"):
        with st.spinner("Sedang menjawab..."):
            jawaban = kirim(st.session_state.chat, materi_awal, pertanyaan)
        st.markdown(jawaban)

    st.session_state.materi_terkirim = True
    st.session_state.riwayat.append({"role": "assistant", "isi": jawaban})

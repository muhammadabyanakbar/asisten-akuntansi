import streamlit as st

from ai_service import buat_chat, kirim_stream
from document_reader import baca_dokumen

st.title("Asisten Akuntansi")

files = st.file_uploader(
    "Unggah materi (PDF atau PPTX), opsional",
    type=["pdf", "pptx"],
    accept_multiple_files=True,
)


@st.cache_data(show_spinner=False)
def ekstrak_materi(nama: str, data: bytes) -> str:
    return baca_dokumen(nama, data)


materi = None
kunci = tuple((f.name, f.size) for f in files) if files else ()

if files:
    bagian = []
    for file in files:
        bagian.append(f"=== Materi: {file.name} ===")
        bagian.append(ekstrak_materi(file.name, file.getvalue()))
    materi = "\n".join(bagian)

# Mulai percakapan baru saat pertama kali dibuka, atau saat file berubah.
if st.session_state.get("kunci_materi") != kunci:
    st.session_state.kunci_materi = kunci
    st.session_state.chat = buat_chat()
    st.session_state.riwayat = []
    st.session_state.materi_terkirim = False

if files:
    st.caption(f"{len(files)} materi aktif. Mengubah file akan memulai percakapan baru.")
else:
    st.caption("Tanpa materi. Anda bisa langsung bertanya tentang akuntansi, keuangan, dan bisnis.")

for pesan in st.session_state.riwayat:
    with st.chat_message(pesan["role"]):
        st.markdown(pesan["isi"])

if pertanyaan := st.chat_input("Tanyakan sesuatu..."):
    st.session_state.riwayat.append({"role": "user", "isi": pertanyaan})
    with st.chat_message("user"):
        st.markdown(pertanyaan)

    materi_awal = None if st.session_state.materi_terkirim else materi
    with st.chat_message("assistant"):
        jawaban = st.write_stream(
            kirim_stream(st.session_state.chat, materi_awal, pertanyaan)
        )

    if materi_awal is not None:
        st.session_state.materi_terkirim = True
    st.session_state.riwayat.append({"role": "assistant", "isi": jawaban})

import os
from collections.abc import Iterator

import streamlit as st
from google import genai
from google.genai import types

MODEL = "gemini-3.5-flash-lite"

SYSTEM_INSTRUCTION = """Kamu adalah asisten AI untuk akuntansi, analisis keuangan, dan bisnis. Kamu berpikir seperti akuntan senior dan analis bisnis, dan membantu mahasiswa akuntansi serta profesional.

Aturan menjawab:
- Jika pengguna sudah mengunggah materi, jawab berdasarkan materi tersebut. Jika informasinya tidak ada di materi, katakan dengan jelas bahwa materi tidak memuat informasi itu. Boleh menjelaskan konsep umumnya, dengan menandainya sebagai pengetahuan umum. Jangan mengarang angka dari materi.
- Jika tidak ada materi, jawab sebagai asisten akuntansi umum: jelaskan konsep, bantu perhitungan, dan koreksi jurnal.
- Sebutkan nomor halaman atau slide jika relevan.
- Untuk perhitungan, tampilkan rumus, langkah, dan hasil akhirnya.
- Tulis angka dengan presisi dan satuan yang jelas.
- Gunakan bahasa Indonesia yang formal namun mudah dipahami."""


@st.cache_resource
def get_client() -> genai.Client:
    api_key = os.environ.get("GEMINI_API_KEY") or st.secrets["GEMINI_API_KEY"]
    return genai.Client(api_key=api_key)


def buat_chat():
    return get_client().chats.create(
        model=MODEL,
        config=types.GenerateContentConfig(system_instruction=SYSTEM_INSTRUCTION),
    )


def kirim_stream(chat, materi_awal: str | None, pertanyaan: str) -> Iterator[str]:
    # Materi hanya dikirim pada pesan pertama. Setelahnya, Gemini mengingat konteksnya.
    isi = f"{materi_awal}\n\nPertanyaan: {pertanyaan}" if materi_awal else pertanyaan
    for potongan in chat.send_message_stream(isi):
        if potongan.text:
            yield potongan.text

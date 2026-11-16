import json
import os
import re
from collections.abc import Iterator
from io import BytesIO

import pandas as pd
import streamlit as st
from docx import Document
from google import genai
from google.genai import types
from pptx import Presentation
from pypdf import PdfReader

MODEL_DEFAULT = "gemini-3.5-flash-lite"
BATAS_PERINGATAN = 50 * 1024 * 1024  # 50 MB

PERAN = {
    "Asisten Akuntansi": (
        "Kamu adalah nexus AI, asisten utama di bidang akuntansi, analisis keuangan, dan bisnis. "
        "Keahlianmu meliputi pencatatan dan jurnal, laporan keuangan, analisis rasio, "
        "perhitungan biaya dan penyusutan, dasar perpajakan, serta analisis tren dan "
        "kinerja usaha. Jawabanmu ringkas, akurat, dan langsung ke inti."
    ),
    "Analis Bisnis": (
        "Kamu adalah analis bisnis senior. Fokus pada implikasi keputusan, risiko, dan "
        "peluang. Sampaikan kesimpulan dulu, lalu bukti dari angka yang tersedia, dan "
        "sebutkan asumsi yang dipakai."
    ),
    "Guru Akuntansi": (
        "Kamu adalah guru akuntansi yang sabar dan teliti. Jelaskan langkah demi langkah "
        "dengan bahasa yang mudah dipahami. Setelah menjelaskan, ajukan satu atau dua "
        "pertanyaan untuk mengecek pemahaman pengguna. Jika pengguna meminta latihan, beri "
        "soal kecil dan tunggu jawabannya sebelum memberi kunci jawaban."
    ),
    "Teman Ngobrol": (
        "Kamu adalah teman ngobrol yang ramah dan hangat. Kamu boleh menanggapi percakapan "
        "santai dengan natural. Jika topiknya menyangkut keuangan, akuntansi, atau bisnis, "
        "berikan jawaban yang akurat dan jelas. Jika percakapan melenceng dari topik, "
        "tanggapi dengan ramah lalu tawarkan bantuan di bidang akuntansi atau bisnis."
    ),
}

ATURAN_BENTUK = """Pilih bentuk jawaban sesuai kebutuhan pengguna:
- Percakapan biasa, sapaan, atau pertanyaan singkat: jawab dengan prosa yang natural dan ringkas. Jangan pakai tabel atau grafik.
- Penjelasan konsep: paragraf pendek dan daftar poin jika perlu. Sertakan contoh angka yang ditandai sebagai contoh ilustratif.
- Analisis, rasio, atau perhitungan: judul singkat, langkah bernomor, rumus, dan tabel markdown untuk hasilnya.
- Jurnal: tabel markdown dengan kolom Tanggal, Akun, Debit, Kredit, dan baris total.
- Perbandingan atau data dengan beberapa baris dan kolom: tabel markdown.
- Tren waktu, perbandingan antarkategori, atau komposisi: tambahkan satu grafik di akhir jawaban dengan format berikut:
```grafik
{"judul": "Tren pendapatan 2024", "jenis": "garis", "label": ["Jan", "Feb", "Mar"], "seri": {"Pendapatan": [100, 120, 90]}}
```
"jenis" harus "garis" untuk tren waktu, atau "batang" untuk perbandingan antarkategori. Setiap daftar di "seri" harus sama panjang dengan "label", dan nilainya berupa angka tanpa satuan atau pemisah ribuan. Pakai paling banyak satu grafik per jawaban, dan hanya jika angkanya benar-benar ada di materi atau sudah dihitung."""

ATURAN_UMUM = """Aturan umum:
- Fokus utama kamu adalah akuntansi, analisis keuangan, dan bisnis. Untuk pertanyaan di luar bidang ini, jawab singkat jika wajar dan aman, lalu arahkan kembali ke bidang utama bila relevan.
- Kamu adalah AI. Jangan mengaku sebagai manusia, dan jangan mengaku punya perasaan pribadi atau hubungan khusus dengan pengguna.
- Jika pengguna sudah mengunggah materi, jawab berdasarkan materi tersebut. Jika informasinya tidak ada, katakan dengan jelas. Jangan mengarang angka dari materi.
- Jika tidak ada materi, jawab sebagai asisten akuntansi umum.
- Akui ketidakpastian. Jika tidak yakin, katakan dengan jelas dan jangan menebak angka atau aturan.
- Untuk pajak, regulasi, atau standar akuntansi, sebutkan dasar atau tanggal aturannya jika diketahui, dan ingatkan agar diverifikasi ke sumber resmi atau akuntan publik atau konsultan pajak. Jawabanmu tidak menggantikan nasihat profesional.
- Sebutkan nomor halaman, slide, atau tabel jika relevan.
- Untuk perhitungan, tampilkan rumus, langkah, dan hasil akhirnya.
- Jika pengguna tampak sangat tertekan atau cemas, tanggapi dengan empati, dan sarankan untuk berbicara dengan orang terdekat atau profesional yang tepat.
- Gunakan bahasa Indonesia yang sesuai dengan peran yang dipilih."""

ANALISIS_CEPAT = {
    "Ringkasan Eksekutif": (
        "Buat ringkasan eksekutif dari materi ini: kondisi utama, angka penting, "
        "tren, dan risiko. Sajikan angka penting dalam tabel, dan tambahkan grafik "
        "jika ada data tren."
    ),
    "Analisis Rasio Keuangan": (
        "Hitung dan jelaskan rasio keuangan utama dari materi ini (likuiditas, "
        "solvabilitas, profitabilitas, dan efisiensi) jika datanya tersedia. "
        "Sajikan hasilnya dalam tabel dengan kolom Rasio, Rumus, Perhitungan, dan Hasil. "
        "Sebutkan jika data tidak tersedia."
    ),
    "Koreksi Jurnal": (
        "Periksa jurnal atau perlakuan akuntansi dalam materi ini. Tunjukkan "
        "kesalahan jika ada, lalu sajikan jurnal koreksi dalam tabel dengan kolom "
        "Tanggal, Akun, Debit, dan Kredit, beserta penjelasannya."
    ),
}

PROMPT_VERIFIKASI = (
    "Periksa ulang jawaban terakhir Anda dengan teliti. Hitung ulang setiap angka, "
    "cek kesesuaiannya dengan materi, dan tandai dengan jelas bagian yang tidak pasti "
    "atau tidak didukung sumber. Jika ada kesalahan, tunjukkan koreksinya."
)

POLA_GRAFIK = re.compile(r"```grafik\s*(.*?)```", re.DOTALL)


# ---------- Pembacaan dokumen ----------

def baca_pdf(data: bytes) -> str:
    reader = PdfReader(BytesIO(data))
    hasil = []
    for nomor, halaman in enumerate(reader.pages, start=1):
        hasil.append(f"Halaman {nomor}:")
        hasil.append(halaman.extract_text() or "")
    return "\n".join(hasil)


def baca_pptx(data: bytes) -> str:
    prs = Presentation(BytesIO(data))
    hasil = []
    for nomor, slide in enumerate(prs.slides, start=1):
        hasil.append(f"Slide {nomor}:")
        for shape in slide.shapes:
            if shape.has_text_frame:
                hasil.append(shape.text_frame.text)
    return "\n".join(hasil)


def baca_docx(data: bytes) -> str:
    dokumen = Document(BytesIO(data))
    hasil = []
    for paragraf in dokumen.paragraphs:
        if paragraf.text.strip():
            hasil.append(paragraf.text)
    for nomor, tabel in enumerate(dokumen.tables, start=1):
        hasil.append(f"Tabel {nomor}:")
        for baris in tabel.rows:
            hasil.append(" | ".join(sel.text.strip() for sel in baris.cells))
    return "\n".join(hasil)


def baca_excel(data: bytes) -> str:
    semua_sheet = pd.read_excel(BytesIO(data), sheet_name=None)
    hasil = []
    for nama_sheet, tabel in semua_sheet.items():
        hasil.append(f"Sheet: {nama_sheet}")
        hasil.append(tabel.to_csv(index=False))
    return "\n".join(hasil)


def baca_csv(data: bytes) -> str:
    tabel = pd.read_csv(BytesIO(data))
    return tabel.to_csv(index=False)


def baca_dokumen(nama: str, data: bytes) -> str:
    nama = nama.lower()
    if nama.endswith(".pptx"):
        return baca_pptx(data)
    if nama.endswith(".docx"):
        return baca_docx(data)
    if nama.endswith(".xlsx"):
        return baca_excel(data)
    if nama.endswith(".csv"):
        return baca_csv(data)
    return baca_pdf(data)


# ---------- Data tabel untuk panel grafik ----------

def ke_angka(teks):
    # Mengubah format Indonesia seperti "Rp9.420.000" atau "(1.500)" menjadi angka.
    s = str(teks).strip().replace("Rp", "").replace(" ", "")
    negatif = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+(,\d+)?", s):
        s = s.replace(".", "").replace(",", ".")
    try:
        nilai = float(s)
    except ValueError:
        return None
    return -nilai if negatif else nilai


def rapikan_angka(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for kolom in df.columns:
        if pd.api.types.is_numeric_dtype(df[kolom]) or pd.api.types.is_datetime64_any_dtype(df[kolom]):
            continue
        konversi = df[kolom].map(ke_angka)
        if konversi.notna().mean() > 0.8:
            df[kolom] = konversi
    return df


def ambil_tabel_docx(data: bytes) -> dict[str, pd.DataFrame]:
    dokumen = Document(BytesIO(data))
    hasil = {}
    for nomor, tabel in enumerate(dokumen.tables, start=1):
        baris = [[sel.text.strip() for sel in r.cells] for r in tabel.rows]
        if len(baris) < 2:
            continue
        header = []
        for i, h in enumerate(baris[0]):
            nama_kolom = h or f"Kolom {i + 1}"
            while nama_kolom in header:
                nama_kolom += "_"
            header.append(nama_kolom)
        df = pd.DataFrame(baris[1:], columns=header)
        if not df.empty:
            hasil[f"Tabel {nomor}"] = rapikan_angka(df)
    return hasil


def ambil_tabel(nama: str, data: bytes) -> dict[str, pd.DataFrame]:
    nama = nama.lower()
    if nama.endswith(".xlsx"):
        semua = pd.read_excel(BytesIO(data), sheet_name=None)
    elif nama.endswith(".csv"):
        semua = {"data": pd.read_csv(BytesIO(data))}
    elif nama.endswith(".docx"):
        return ambil_tabel_docx(data)
    else:
        return {}
    return {sheet: rapikan_angka(df) for sheet, df in semua.items() if not df.empty}


# ---------- Layanan AI ----------

@st.cache_resource
def get_client() -> genai.Client:
    api_key = os.environ.get("GEMINI_API_KEY") or st.secrets["GEMINI_API_KEY"]
    return genai.Client(api_key=api_key)


def buat_chat(peran: str, model: str, pakai_web: bool, berpikir: bool):
    system = f"{PERAN[peran]}\n\n{ATURAN_BENTUK}\n\n{ATURAN_UMUM}"
    if pakai_web:
        system += (
            "\n\nPencarian web aktif. Gunakan hasil pencarian untuk informasi terkini, "
            "dan sebutkan sumbernya secara singkat di akhir jawaban."
        )
    konfigurasi = types.GenerateContentConfig(
        system_instruction=system,
        tools=[types.Tool(google_search=types.GoogleSearch())] if pakai_web else None,
        thinking_config=types.ThinkingConfig(thinking_budget=-1) if berpikir else None,
    )
    return get_client().chats.create(model=model, config=konfigurasi)


def kirim_stream(chat, materi_awal: str | None, pertanyaan: str) -> Iterator[str]:
    # Materi hanya dikirim pada pesan pertama. Setelahnya, Gemini mengingat konteksnya.
    isi = f"{materi_awal}\n\nPertanyaan: {pertanyaan}" if materi_awal else pertanyaan
    for potongan in chat.send_message_stream(isi):
        if potongan.text:
            yield potongan.text


# ---------- Pengolahan jawaban ----------

def pisahkan_grafik(teks: str):
    grafik = []
    for blok in POLA_GRAFIK.findall(teks):
        try:
            data = json.loads(blok)
            label = [str(x) for x in data["label"]]
            seri = {}
            for nama, nilai in data["seri"].items():
                if len(nilai) == len(label):
                    seri[str(nama)] = [float(v) for v in nilai]
            if seri:
                grafik.append({
                    "judul": str(data.get("judul", "")),
                    "jenis": data.get("jenis", "garis"),
                    "label": label,
                    "seri": seri,
                })
        except (ValueError, KeyError, TypeError):
            continue
    bersih = POLA_GRAFIK.sub("", teks).strip()
    return bersih, grafik


def sembunyikan_grafik_sementara(teks: str) -> str:
    return re.split(r"```grafik", teks)[0]


def tulis_streaming(generator: Iterator[str], wadah) -> str:
    teks = ""
    for potongan in generator:
        teks += potongan
        wadah.markdown(sembunyikan_grafik_sementara(teks) + " ▌")
    return teks


def tampilkan_grafik(grafik: dict):
    df = pd.DataFrame(grafik["seri"], index=grafik["label"])
    if grafik["judul"]:
        st.caption(grafik["judul"])
    if grafik["jenis"] == "batang":
        st.bar_chart(df)
    else:
        st.line_chart(df)


def render_pesan(pesan: dict):
    with st.chat_message(pesan["role"]):
        st.markdown(pesan["isi"])
        for g in pesan.get("grafik", []):
            tampilkan_grafik(g)


# ---------- Antarmuka ----------

@st.cache_data(show_spinner=False)
def ekstrak_materi(nama: str, data: bytes) -> str:
    return baca_dokumen(nama, data)


@st.cache_data(show_spinner=False)
def ekstrak_tabel(nama: str, data: bytes) -> dict[str, pd.DataFrame]:
    return ambil_tabel(nama, data)


with st.sidebar:
    st.header("Pengaturan")
    model = st.text_input("Model Gemini", value=MODEL_DEFAULT)
    berpikir = st.toggle("Mode berpikir mendalam", value=False)
    pakai_web = st.toggle("Cari informasi di web", value=False)
    st.caption(
        "Mode berpikir dan pencarian web membuat jawaban lebih teliti, "
        "tetapi memakai lebih banyak kuota dan waktu."
    )

st.title("Nexus AI")

files = st.file_uploader(
    "Unggah materi (PDF, PPTX, DOCX, XLSX, CSV), opsional",
    type=["pdf", "pptx", "docx", "xlsx", "csv"],
    accept_multiple_files=True,
)

peran = st.radio("Peran AI", list(PERAN.keys()), horizontal=True)

materi = None
tabel_tersedia: dict[str, pd.DataFrame] = {}
kunci = tuple((f.name, f.size) for f in files) if files else ()

if files:
    for file in files:
        if file.size > BATAS_PERINGATAN:
            st.warning(
                f"'{file.name}' berukuran lebih dari 50 MB. "
                "Pemrosesan bisa memakan waktu lebih lama."
            )
    bagian = []
    for file in files:
        bagian.append(f"=== Materi: {file.name} ===")
        bagian.append(ekstrak_materi(file.name, file.getvalue()))
        for sheet, df in ekstrak_tabel(file.name, file.getvalue()).items():
            tabel_tersedia[f"{file.name} · {sheet}"] = df
    materi = "\n".join(bagian)

# Mulai percakapan baru saat file, peran, model, atau pengaturan berubah.
kunci_sesi = (kunci, peran, model, berpikir, pakai_web)
if st.session_state.get("kunci_sesi") != kunci_sesi:
    st.session_state.kunci_sesi = kunci_sesi
    st.session_state.chat = buat_chat(peran, model, pakai_web, berpikir)
    st.session_state.riwayat = []
    st.session_state.materi_terkirim = False

if files:
    st.caption(f"{len(files)} materi aktif. Mengubah file, peran, atau pengaturan akan memulai percakapan baru.")
else:
    st.caption("Tanpa materi. Anda bisa langsung bertanya tentang akuntansi, keuangan, dan bisnis.")

# Panel grafik manual, hanya muncul jika ada tabel dari Excel, CSV, atau Word.
if tabel_tersedia:
    with st.expander("Grafik dari data tabel", expanded=False):
        nama_tabel = st.selectbox("Pilih tabel", list(tabel_tersedia.keys()))
        df = tabel_tersedia[nama_tabel]
        kolom_semua = list(df.columns)
        kolom_angka = list(df.select_dtypes("number").columns)

        if not kolom_angka:
            st.info("Tabel ini tidak memiliki kolom angka untuk digambar.")
        else:
            kolom_label = [c for c in kolom_semua if c not in kolom_angka]
            x_default = kolom_label[0] if kolom_label else kolom_semua[0]
            x = st.selectbox(
                "Sumbu X (label)",
                kolom_semua,
                index=kolom_semua.index(x_default),
            )
            pilihan_y = [c for c in kolom_angka if c != x]
            y = st.multiselect("Sumbu Y (angka)", pilihan_y, default=pilihan_y[:2])
            jenis = st.radio("Jenis grafik", ["Garis", "Batang"], horizontal=True)

            if y:
                grafik_data = df.set_index(x)[y]
                if jenis == "Garis":
                    st.line_chart(grafik_data)
                else:
                    st.bar_chart(grafik_data)
            else:
                st.info("Pilih minimal satu kolom angka untuk sumbu Y.")

for pesan in st.session_state.riwayat:
    render_pesan(pesan)

# Tombol analisis cepat dan periksa ulang.
aksi = None
if files:
    kolom = st.columns(len(ANALISIS_CEPAT))
    for k, (label, perintah) in zip(kolom, ANALISIS_CEPAT.items()):
        if k.button(label):
            aksi = (label, perintah)

if st.session_state.riwayat and st.session_state.riwayat[-1]["role"] == "assistant":
    if st.button("Periksa ulang jawaban terakhir"):
        aksi = ("Periksa ulang jawaban terakhir", PROMPT_VERIFIKASI)

masukan = st.chat_input("Tanyakan sesuatu...")
if masukan and aksi is None:
    aksi = (masukan, masukan)

if aksi:
    tampilan, perintah = aksi
    st.session_state.riwayat.append({"role": "user", "isi": tampilan})
    with st.chat_message("user"):
        st.markdown(tampilan)

    materi_awal = None if st.session_state.materi_terkirim else materi
    with st.chat_message("assistant"):
        wadah = st.empty()
        try:
            mentah = tulis_streaming(
                kirim_stream(st.session_state.chat, materi_awal, perintah),
                wadah,
            )
        except Exception as e:
            mentah = (
                "Maaf, permintaan ini gagal diproses. Periksa nama model, kuota, "
                f"atau koneksi, lalu coba lagi.\n\nDetail: {e}"
            )
        bersih, grafik = pisahkan_grafik(mentah)
        wadah.markdown(bersih)
        for g in grafik:
            tampilkan_grafik(g)

    if materi_awal is not None:
        st.session_state.materi_terkirim = True
    st.session_state.riwayat.append(
        {"role": "assistant", "isi": bersih, "grafik": grafik}
    )

# Tombol unduh percakapan.
if st.session_state.riwayat:
    teks = "\n\n".join(
        f"**{'Pengguna' if p['role'] == 'user' else 'Asisten'}:**\n\n{p['isi']}"
        for p in st.session_state.riwayat
    )
    st.download_button(
        "Unduh percakapan (.md)",
        data=teks,
        file_name="percakapan-asisten-akuntansi.md",
        mime="text/markdown",
    )

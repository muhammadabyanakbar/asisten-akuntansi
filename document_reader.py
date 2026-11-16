from io import BytesIO

import pandas as pd
from docx import Document
from pptx import Presentation
from pypdf import PdfReader


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

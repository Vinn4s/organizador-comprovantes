from __future__ import annotations

from io import BytesIO
import re

import fitz
import pandas as pd
import pytesseract
import streamlit as st
from PIL import Image


st.set_page_config(
    page_title="Organizador de Comprovantes",
    page_icon="📄",
    layout="wide",
)


def extract_text_from_image(file_bytes: bytes) -> str:
    """Extrai texto de imagens PNG, JPG e JPEG usando OCR."""
    image = Image.open(BytesIO(file_bytes)).convert("RGB")

    return pytesseract.image_to_string(
        image,
        lang="por",
        config="--psm 6",
    ).strip()


def extract_text_from_pdf(file_bytes: bytes) -> str:
    """
    Extrai texto nativo do PDF.

    Quando uma página não possui texto incorporado,
    transforma a página em imagem e usa OCR.
    """
    document = fitz.open(stream=file_bytes, filetype="pdf")
    extracted_pages: list[str] = []

    for page in document:
        page_text = page.get_text("text").strip()

        if page_text:
            extracted_pages.append(page_text)
            continue

        pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        page_image = Image.open(BytesIO(pixmap.tobytes("png")))

        ocr_text = pytesseract.image_to_string(
            page_image,
            lang="por",
            config="--psm 6",
        ).strip()

        extracted_pages.append(ocr_text)

    document.close()

    return "\n".join(extracted_pages).strip()


def extract_text(file_name: str, file_type: str, file_bytes: bytes) -> str:
    """Escolhe o extrator adequado conforme o formato enviado."""
    extension = file_name.lower().rsplit(".", maxsplit=1)[-1]

    if extension == "pdf" or file_type == "application/pdf":
        return extract_text_from_pdf(file_bytes)

    if extension in {"png", "jpg", "jpeg"}:
        return extract_text_from_image(file_bytes)

    return ""

DATE_PATTERN = re.compile(
    r"\b([0-3]?\d/[01]?\d/(?:\d{2}|\d{4}))\b"
)

PRIORITY_VALUE_PATTERN = re.compile(
    r"(?:"
    r"valor\s+final|"
    r"valor\s+pago|"
    r"total\s+pago|"
    r"valor\s+da\s+transa[cç][aã]o|"
    r"valor\s+do\s+pagamento"
    r")"
    r"\s*:?\s*R\$\s*([\d.]+,\d{2})",
    re.IGNORECASE,
)

GENERIC_VALUE_PATTERN = re.compile(
    r"R\$\s*([\d.]+,\d{2})",
    re.IGNORECASE,
)


def detect_date(text: str) -> str:
    """Encontra a primeira data no formato brasileiro."""
    match = DATE_PATTERN.search(text)

    if not match:
        return ""

    return match.group(1)


def detect_value(text: str) -> str:
    """
    Prioriza valores associados a expressões como
    'valor final' e 'valor pago'.
    """
    match = PRIORITY_VALUE_PATTERN.search(text)

    if not match:
        match = GENERIC_VALUE_PATTERN.search(text)

    if not match:
        return ""

    return f"R$ {match.group(1)}"


def detect_document_type(text: str, file_name: str) -> str:
    """Identifica o tipo do documento usando texto e nome do arquivo."""
    searchable_text = f"{text}\n{file_name}".lower()

    if "pix" in searchable_text:
        return "Pix"

    if "boleto" in searchable_text:
        return "Boleto"

    if "nota fiscal" in searchable_text:
        return "Nota fiscal"

    if re.search(r"\bnf\s*\d*", searchable_text):
        return "Nota fiscal"

    if "recibo" in searchable_text:
        return "Recibo"

    return "Comprovante"


def parse_document(text: str, file_name: str) -> dict[str, str]:
    """Transforma o texto bruto em campos estruturados."""
    return {
        "Data": detect_date(text),
        "Valor": detect_value(text),
        "Tipo": detect_document_type(text, file_name),
    }


st.title("Organizador de Comprovantes")
st.caption(
    "Envie comprovantes em PDF ou imagem para iniciar a organização."
)

uploaded_files = st.file_uploader(
    "Selecione os documentos",
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True,
)

if not uploaded_files:
    st.info("Nenhum documento enviado ainda.")
    st.stop()

if st.button("Ler documentos", type="primary"):
    documents: list[dict[str, object]] = []

    progress_bar = st.progress(0, text="Preparando leitura...")

    for index, uploaded_file in enumerate(uploaded_files):
        file_bytes = uploaded_file.getvalue()

        try:
            raw_text = extract_text(
                file_name=uploaded_file.name,
                file_type=uploaded_file.type or "",
                file_bytes=file_bytes,
            )   
            
            parsed_document = parse_document(
                text=raw_text,
                file_name=uploaded_file.name,
            )

            extraction_status = (
                "Texto encontrado"
                if raw_text
                else "Nenhum texto encontrado"
            )

        except Exception as error:
            raw_text = ""
            extraction_status = f"Erro: {error}"

        documents.append(
            {
                "Arquivo original": uploaded_file.name,
                "Formato": uploaded_file.type or "Não identificado",
                "Tamanho em KB": round(uploaded_file.size / 1024, 2),
                "Status da leitura": extraction_status,
                "Data": parsed_document["Data"],
                "Valor": parsed_document["Valor"],
                "Tipo": parsed_document["Tipo"],
                "Valor": "",
                "Tipo": "",
                "Pagador": "",
                "Recebedor": "",
                "Documento": "",
                "Descrição": "",
                "Identificador": "",
                "Possível duplicidade": "Não",
                "Observações": "",
                "Texto extraído": raw_text,
            }
        )

        progress = int(((index + 1) / len(uploaded_files)) * 100)

        progress_bar.progress(
            progress,
            text=f"Lendo {index + 1} de {len(uploaded_files)}...",
        )

    progress_bar.empty()

    dataframe = pd.DataFrame(documents)

    st.session_state["documents_dataframe"] = dataframe

if "documents_dataframe" in st.session_state:
    dataframe = st.session_state["documents_dataframe"]

    st.success(f"{len(dataframe)} documento(s) processado(s).")

    edited_dataframe = st.data_editor(
        dataframe,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key="documents_editor",
    )

    st.subheader("Prévia da organização")

    st.dataframe(
        edited_dataframe,
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("Verificação da leitura")

    selected_file = st.selectbox(
        "Escolha um documento para visualizar o texto extraído",
        options=edited_dataframe["Arquivo original"].tolist(),
    )

    selected_row = edited_dataframe[
        edited_dataframe["Arquivo original"] == selected_file
    ].iloc[0]

    st.text_area(
        "Texto identificado no documento",
        value=str(selected_row["Texto extraído"]),
        height=300,
        disabled=True,
    )
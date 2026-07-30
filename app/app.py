from __future__ import annotations

import hashlib
import re
from io import BytesIO

import fitz
import pandas as pd
import pytesseract
import streamlit as st
from openpyxl.styles import Font
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
    transforma a página em imagem e utiliza OCR.
    """
    document = fitz.open(stream=file_bytes, filetype="pdf")
    extracted_pages: list[str] = []

    try:
        for page in document:
            page_text = page.get_text("text").strip()

            if page_text:
                extracted_pages.append(page_text)
                continue

            pixmap = page.get_pixmap(
                matrix=fitz.Matrix(2, 2),
                alpha=False,
            )

            page_image = Image.open(
                BytesIO(pixmap.tobytes("png"))
            ).convert("RGB")

            ocr_text = pytesseract.image_to_string(
                page_image,
                lang="por",
                config="--psm 6",
            ).strip()

            extracted_pages.append(ocr_text)

    finally:
        document.close()

    return "\n".join(extracted_pages).strip()


def extract_text(
    file_name: str,
    file_type: str,
    file_bytes: bytes,
) -> str:
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
    r"valor\s+do\s+pagamento|"
    r"valor\s+original"
    r")"
    r"\s*:?\s*R\$\s*([\d.]+,\d{2})",
    re.IGNORECASE,
)

GENERIC_VALUE_PATTERN = re.compile(
    r"R\$\s*([\d.]+,\d{2})",
    re.IGNORECASE,
)

CPF_CNPJ_PATTERN = re.compile(
    r"(?<!\w)(?:"
    r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}|"
    r"\d{3}\.\d{3}\.\d{3}-\d{2}|"
    r"\d{14}|"
    r"\d{11}"
    r")(?!\w)"
)

RECEIVER_LABEL = (
    r"(?:"
    r"nome\s+do\s+recebedor|"
    r"quem\s+recebeu|"
    r"benefici[aá]rio|"
    r"recebedor|"
    r"favorecido"
    r")"
)

PAYER_LABEL = (
    r"(?:"
    r"nome\s+do\s+pagador|"
    r"quem\s+pagou|"
    r"pagador|"
    r"pagante"
    r")"
)

IDENTIFIER_LABEL = (
    r"(?:"
    r"identificador\s+da\s+transa[cç][aã]o|"
    r"id\s+da\s+transa[cç][aã]o|"
    r"end\s*to\s*end\s*id|"
    r"e2e\s*id|"
    r"c[oó]digo\s+da\s+transa[cç][aã]o"
    r")"
)

PAYMENT_REFERENCE_LABEL = (
    r"(?:"
    r"n[uú]mero\s+do\s+documento|"
    r"c[oó]digo\s+do\s+pagamento|"
    r"refer[eê]ncia|"
    r"identificador"
    r")"
)

RECEIVER_LABEL_PATTERN = re.compile(
    rf"^\s*{RECEIVER_LABEL}\b",
    re.IGNORECASE,
)

PAYER_LABEL_PATTERN = re.compile(
    rf"^\s*{PAYER_LABEL}\b",
    re.IGNORECASE,
)

IDENTIFIER_LABEL_PATTERN = re.compile(
    rf"^\s*{IDENTIFIER_LABEL}\b",
    re.IGNORECASE,
)

PAYMENT_REFERENCE_LABEL_PATTERN = re.compile(
    rf"^\s*{PAYMENT_REFERENCE_LABEL}\b",
    re.IGNORECASE,
)

DOCUMENT_CONTEXT_LABEL_PATTERN = re.compile(
    r"(?:"
    r"quem\s+recebeu|"
    r"recebedor|"
    r"quem\s+pagou|"
    r"pagador|"
    r"cpf\s*/\s*cnpj|"
    r"documento"
    r")\b",
    re.IGNORECASE,
)

PARTY_FIELD_BOUNDARY_PATTERN = re.compile(
    rf"^\s*(?:"
    rf"cpf\s*/\s*cnpj|"
    rf"documento|"
    rf"institui[cç][aã]o\s*(?::|$)|"
    rf"banco\s*(?::|$)|"
    rf"{IDENTIFIER_LABEL}|"
    rf"{PAYMENT_REFERENCE_LABEL}|"
    rf"{RECEIVER_LABEL}|"
    rf"{PAYER_LABEL}"
    rf")(?=\W|$)",
    re.IGNORECASE,
)

INLINE_PARTY_BOUNDARY_PATTERN = re.compile(
    rf"\s*(?:/|\|)\s*(?=(?:"
    rf"cpf\s*/\s*cnpj|"
    rf"documento|"
    rf"institui[cç][aã]o|"
    rf"banco\s*:|"
    rf"{IDENTIFIER_LABEL}|"
    rf"{PAYMENT_REFERENCE_LABEL}|"
    rf"{RECEIVER_LABEL}|"
    rf"{PAYER_LABEL}"
    rf")(?=\W|$))",
    re.IGNORECASE,
)

NAME_LABEL_PATTERN = re.compile(
    r"^\s*nome\b",
    re.IGNORECASE,
)

FIELD_PREFIX_PATTERN = re.compile(
    rf"^\s*(?:"
    rf"nome|"
    rf"cpf\s*/\s*cnpj|"
    rf"documento|"
    rf"institui[cç][aã]o|"
    rf"banco|"
    rf"data|"
    rf"valor|"
    rf"tipo|"
    rf"descri[cç][aã]o|"
    rf"{IDENTIFIER_LABEL}|"
    rf"{PAYMENT_REFERENCE_LABEL}|"
    rf"{RECEIVER_LABEL}|"
    rf"{PAYER_LABEL}"
    rf")\b",
    re.IGNORECASE,
)


def _nonempty_lines(text: str) -> list[str]:
    """Retorna as linhas úteis sem alterar seus espaços internos."""
    return [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]


def _party_value_after_label(
    line: str,
    label_match: re.Match[str],
) -> tuple[str, bool]:
    """Lê um valor inline apenas quando o rótulo é inequívoco."""
    remainder = line[label_match.end():]

    if not remainder.strip():
        return "", True

    remainder = remainder.lstrip()

    if remainder[:1] in {":", "-", "–", "—", "/", "|"}:
        return remainder[1:].strip(), True

    if not any(character.isupper() for character in remainder):
        return "", False

    return remainder.strip(), True


def _party_candidate(value: str) -> tuple[str, bool]:
    """
    Valida um possível nome.

    O segundo item indica que foi encontrada apenas a linha "Nome"
    e que o valor deve ser procurado na linha seguinte.
    """
    candidate = value.strip()

    if not candidate:
        return "", False

    candidate = INLINE_PARTY_BOUNDARY_PATTERN.split(
        candidate,
        maxsplit=1,
    )[0].strip()

    if PARTY_FIELD_BOUNDARY_PATTERN.match(candidate):
        return "", False

    name_match = NAME_LABEL_PATTERN.match(candidate)

    if name_match:
        candidate = candidate[name_match.end():].lstrip()

        if candidate[:1] in {":", "-", "–", "—", "/", "|"}:
            candidate = candidate[1:].lstrip()

        candidate = INLINE_PARTY_BOUNDARY_PATTERN.split(
            candidate,
            maxsplit=1,
        )[0].strip()

        if not candidate:
            return "", True

        if PARTY_FIELD_BOUNDARY_PATTERN.match(candidate):
            return "", False

    if not re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]", candidate):
        return "", False

    return candidate, False


def _detect_party_name(
    text: str,
    label_pattern: re.Pattern[str],
) -> str:
    """Procura um nome somente depois de um rótulo de seção claro."""
    lines = _nonempty_lines(text)

    for index, line in enumerate(lines):
        label_match = label_pattern.match(line)

        if not label_match:
            continue

        inline_value, is_clear_label = _party_value_after_label(
            line,
            label_match,
        )

        if not is_clear_label:
            continue

        candidate, expects_next_line = _party_candidate(inline_value)

        if candidate:
            return candidate

        if inline_value and not expects_next_line:
            continue

        next_index = index + 1

        if next_index >= len(lines):
            continue

        candidate, expects_next_line = _party_candidate(
            lines[next_index]
        )

        if candidate:
            return candidate

        if not expects_next_line:
            continue

        value_index = next_index + 1

        if value_index >= len(lines):
            continue

        candidate, _ = _party_candidate(lines[value_index])

        if candidate:
            return candidate

    return ""


def detect_receiver(text: str) -> str:
    """Detecta o recebedor somente em uma seção rotulada."""
    return _detect_party_name(text, RECEIVER_LABEL_PATTERN)


def detect_payer(text: str) -> str:
    """Detecta o pagador somente em uma seção rotulada."""
    return _detect_party_name(text, PAYER_LABEL_PATTERN)


def detect_document(text: str) -> str:
    """Detecta CPF/CNPJ, priorizando ocorrências próximas a rótulos."""
    document_matches = list(CPF_CNPJ_PATTERN.finditer(text))

    if not document_matches:
        return ""

    context_matches = list(
        DOCUMENT_CONTEXT_LABEL_PATTERN.finditer(text)
    )

    if not context_matches:
        return document_matches[0].group(0)

    closest_match = min(
        document_matches,
        key=lambda document_match: (
            min(
                (
                    _match_distance(
                        document_match,
                        context_match,
                    ),
                    (
                        0
                        if context_match.end()
                        <= document_match.start()
                        else 1
                    ),
                )
                for context_match in context_matches
            ),
            document_match.start(),
        ),
    )

    return closest_match.group(0)


def _match_distance(
    first_match: re.Match[str],
    second_match: re.Match[str],
) -> int:
    """Calcula a distância entre duas ocorrências no texto."""
    if first_match.end() <= second_match.start():
        return second_match.start() - first_match.end()

    if second_match.end() <= first_match.start():
        return first_match.start() - second_match.end()

    return 0


def _identifier_after_label(
    line: str,
    label_match: re.Match[str],
) -> str:
    """Remove o rótulo e seu separador, preservando o identificador."""
    identifier = line[label_match.end():].strip()

    if identifier.startswith(":"):
        identifier = identifier[1:].strip()
    elif re.match(r"^[-–—/|]\s+", identifier):
        identifier = identifier[1:].strip()

    return identifier


def detect_identifier(text: str) -> str:
    """Detecta o ID específico da transação sem mudar seu conteúdo."""
    lines = _nonempty_lines(text)

    for index, line in enumerate(lines):
        label_match = IDENTIFIER_LABEL_PATTERN.match(line)

        if not label_match:
            continue

        identifier = _identifier_after_label(line, label_match)

        if identifier:
            return identifier

        next_index = index + 1

        if next_index >= len(lines):
            continue

        identifier = lines[next_index].strip()

        if FIELD_PREFIX_PATTERN.match(identifier):
            continue

        return identifier

    return ""


def detect_payment_reference(text: str) -> str:
    """Detecta referências de cobrança sem capturar IDs de transação."""
    lines = _nonempty_lines(text)

    for index, line in enumerate(lines):
        if IDENTIFIER_LABEL_PATTERN.match(line):
            continue

        label_match = PAYMENT_REFERENCE_LABEL_PATTERN.match(line)

        if not label_match:
            continue

        reference = _identifier_after_label(line, label_match)

        if reference:
            return reference

        next_index = index + 1

        if next_index >= len(lines):
            continue

        reference = lines[next_index].strip()

        if FIELD_PREFIX_PATTERN.match(reference):
            continue

        return reference

    return ""


def detect_description(file_name: str) -> str:
    """Gera uma descrição conservadora apenas com o nome do arquivo."""
    description = file_name.rsplit(".", maxsplit=1)[0]
    description = description.replace("_", " ")
    description = re.sub(
        r"R\$\s*[\d.]+,\d{2}",
        "",
        description,
        flags=re.IGNORECASE,
    )
    description = re.sub(
        r"\b(?:comprovante|pix|pagamento|quita[cç][aã]o)\b",
        "",
        description,
        flags=re.IGNORECASE,
    )
    description = re.sub(r"\s+", " ", description).strip()
    description = re.sub(
        r"^(?:[-–—|/,:;]\s*)+|(?:\s*[-–—|/,:;])+$",
        "",
        description,
    )

    return description.strip()


def mark_possible_duplicates(
    documents: list[dict[str, object]],
    file_hashes: list[str],
) -> None:
    """Marca documentos repetidos por hash, ID da transação ou referência."""
    duplicate_indexes: set[int] = set()
    grouped_indexes: dict[str, list[int]] = {}

    for index, file_hash in enumerate(file_hashes):
        grouped_indexes.setdefault(file_hash, []).append(index)

    for indexes in grouped_indexes.values():
        if len(indexes) > 1:
            duplicate_indexes.update(indexes)

    for field in ("Identificador", "Referência"):
        grouped_indexes = {}

        for index, document in enumerate(documents):
            value = str(document.get(field, "")).strip()

            if value:
                grouped_indexes.setdefault(value, []).append(index)

        for indexes in grouped_indexes.values():
            if len(indexes) > 1:
                duplicate_indexes.update(indexes)

    for index, document in enumerate(documents):
        document["Possível duplicidade"] = (
            "Sim" if index in duplicate_indexes else "Não"
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
    'valor final', 'valor pago' e 'valor original'.
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
        "Pagador": detect_payer(text),
        "Recebedor": detect_receiver(text),
        "Documento": detect_document(text),
        "Referência": detect_payment_reference(text),
        "Descrição": detect_description(file_name),
        "Identificador": detect_identifier(text),
    }


def generate_excel(dataframe: pd.DataFrame) -> bytes:
    """
    Gera uma planilha Excel com os dados organizados
    e uma segunda aba contendo o texto bruto do OCR.
    """
    output = BytesIO()

    main_dataframe = dataframe.drop(
        columns=["Texto extraído"],
        errors="ignore",
    ).copy()

    audit_columns = [
        column
        for column in ["Arquivo original", "Texto extraído"]
        if column in dataframe.columns
    ]

    audit_dataframe = dataframe[audit_columns].copy()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        main_dataframe.to_excel(
            writer,
            index=False,
            sheet_name="Comprovantes",
        )

        audit_dataframe.to_excel(
            writer,
            index=False,
            sheet_name="Leitura OCR",
        )

        for worksheet in writer.book.worksheets:
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions

            for cell in worksheet[1]:
                cell.font = Font(bold=True)

            for column_cells in worksheet.columns:
                column_letter = column_cells[0].column_letter

                largest_content = max(
                    len(str(cell.value))
                    if cell.value is not None
                    else 0
                    for cell in column_cells
                )

                worksheet.column_dimensions[column_letter].width = min(
                    largest_content + 3,
                    50,
                )

    output.seek(0)

    return output.getvalue()


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
    file_hashes: list[str] = []

    progress_bar = st.progress(
        0,
        text="Preparando leitura...",
    )

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

            parsed_document = {
                "Data": "",
                "Valor": "",
                "Tipo": "",
                "Pagador": "",
                "Recebedor": "",
                "Documento": "",
                "Referência": "",
                "Descrição": "",
                "Identificador": "",
            }

            extraction_status = f"Erro: {error}"

        documents.append(
            {
                "Arquivo original": uploaded_file.name,
                "Formato": (
                    uploaded_file.type
                    or "Não identificado"
                ),
                "Tamanho em KB": round(
                    uploaded_file.size / 1024,
                    2,
                ),
                "Status da leitura": extraction_status,
                "Data": parsed_document["Data"],
                "Valor": parsed_document["Valor"],
                "Tipo": parsed_document["Tipo"],
                "Pagador": parsed_document["Pagador"],
                "Recebedor": parsed_document["Recebedor"],
                "Documento": parsed_document["Documento"],
                "Referência": parsed_document["Referência"],
                "Descrição": parsed_document["Descrição"],
                "Identificador": parsed_document["Identificador"],
                "Possível duplicidade": "Não",
                "Observações": "",
                "Texto extraído": raw_text,
            }
        )
        file_hashes.append(
            hashlib.sha256(file_bytes).hexdigest()
        )

        progress = int(
            ((index + 1) / len(uploaded_files)) * 100
        )

        progress_bar.progress(
            progress,
            text=(
                f"Lendo {index + 1} "
                f"de {len(uploaded_files)}..."
            ),
        )

    progress_bar.empty()

    mark_possible_duplicates(documents, file_hashes)

    dataframe = pd.DataFrame(documents)

    st.session_state["documents_dataframe"] = dataframe

if "documents_dataframe" in st.session_state:
    dataframe = st.session_state["documents_dataframe"]

    st.success(
        f"{len(dataframe)} documento(s) processado(s)."
    )

    edited_dataframe = st.data_editor(
        dataframe,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key="documents_editor",
    )

    excel_file = generate_excel(edited_dataframe)

    st.download_button(
        label="Baixar planilha Excel",
        data=excel_file,
        file_name="comprovantes_organizados.xlsx",
        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
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
        options=edited_dataframe[
            "Arquivo original"
        ].tolist(),
    )

    selected_row = edited_dataframe[
        edited_dataframe["Arquivo original"]
        == selected_file
    ].iloc[0]

    st.text_area(
        "Texto identificado no documento",
        value=str(selected_row["Texto extraído"]),
        height=300,
        disabled=True,
    )

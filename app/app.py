from __future__ import annotations

import hashlib
import json
import os
import re
import uuid
from collections.abc import MutableMapping
from decimal import Decimal, InvalidOperation
from io import BytesIO

import fitz
import pandas as pd
import pytesseract
import requests
import streamlit as st
from openpyxl.styles import Font
from PIL import Image


MERCADO_PAGO_PREFERENCES_URL = (
    "https://api.mercadopago.com/checkout/preferences"
)
MERCADO_PAGO_PAYMENTS_SEARCH_URL = (
    "https://api.mercadopago.com/v1/payments/search"
)
PAYMENT_TIMEOUT_SECONDS = 15
PAYMENT_CURRENCY = "BRL"
PAYMENT_STATE_DEFAULTS = {
    "documents_dataframe": None,
    "excel_file": None,
    "external_reference": None,
    "preference_id": None,
    "payment_url": None,
    "payment_status": None,
    "payment_approved": False,
    "expected_payment_amount": None,
}
PUBLIC_MASKED_PREVIEW_COLUMNS = [
    "Data",
    "Valor",
    "Tipo",
    "Pagador",
    "Recebedor",
    "Descrição",
    "Possível duplicidade",
]
PUBLIC_FOUND_FIELD_COLUMNS = [
    "Data",
    "Valor",
    "Tipo",
    "Pagador",
    "Recebedor",
    "Descrição",
]
_CONFIG_MISSING = object()


class PaymentServiceError(Exception):
    """Erro seguro para exibição durante a comunicação de pagamento."""


class ConfigurationError(PaymentServiceError):
    """Erro seguro para uma configuração obrigatória ou inválida."""


def get_config_value(
    name: str,
    *,
    default: object = _CONFIG_MISSING,
    required: bool = False,
) -> object:
    """Lê configuração do ambiente ou, como fallback, do Streamlit.

    Variáveis de ambiente têm prioridade para permitir o uso de secrets do
    Cloud Run, preservando o suporte a ``.streamlit/secrets.toml`` local e ao
    Streamlit Community Cloud.
    """
    value = os.getenv(name)
    if value is None:
        try:
            value = st.secrets.get(name)
        except (AttributeError, FileNotFoundError, KeyError, OSError):
            value = None

    if value is None or (isinstance(value, str) and not value.strip()):
        if required:
            raise ConfigurationError(
                "A configuração necessária para pagamentos não está disponível."
            )
        return None if default is _CONFIG_MISSING else default

    return value


def parse_boolean(value: object) -> bool:
    """Converte valores textuais de configuração em booleano com segurança."""
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        normalized_value = value.strip().lower()
        if normalized_value in {"true", "1", "yes", "on"}:
            return True
        if normalized_value in {"false", "0", "no", "off", ""}:
            return False
    if isinstance(value, int):
        if value == 1:
            return True
        if value == 0:
            return False
    raise ConfigurationError("A configuração informada é inválida.")


def is_sandbox_environment(environment: str) -> bool:
    """Indica se o ambiente pode exibir diagnósticos restritos de pagamento."""
    return environment.lower() in {"sandbox", "test", "teste"}


def redact_payment_diagnostic(value: object) -> str:
    """Remove valores de credenciais de um diagnóstico remoto limitado."""
    sensitive_key_pattern = re.compile(
        r"authorization|access[_ -]?token|api[_ -]?key|secret",
        re.IGNORECASE,
    )

    def redact(item: object) -> object:
        if isinstance(item, dict):
            return {
                str(key): redact(nested_value)
                for key, nested_value in item.items()
                if not sensitive_key_pattern.search(str(key))
            }
        if isinstance(item, list):
            return [redact(nested_value) for nested_value in item]
        if isinstance(item, str):
            item = re.sub(
                r"(?i)authorization\s*[:=]\s*[^\s,;]+(?:\s+[^\s,;]+)?",
                "[credencial ocultada]",
                item,
            )
            item = re.sub(
                r"(?i)access[_ -]?token\s*[=:]\s*[^\s,;]+",
                "[credencial ocultada]",
                item,
            )
            item = re.sub(
                r"(?i)bearer\s+[^\s,;]+",
                "[credencial ocultada]",
                item,
            )
        return item

    safe_value = redact(value)
    if isinstance(safe_value, (dict, list)):
        return json.dumps(safe_value, ensure_ascii=False)
    return str(safe_value)


def payment_preference_error_message(
    response: requests.Response | None,
    environment: str,
) -> str:
    """Monta erro de preferência seguro, detalhado somente em sandbox."""
    generic_message = (
        "Não foi possível iniciar o pagamento agora. Tente novamente."
    )
    if not is_sandbox_environment(environment) or response is None:
        return generic_message

    details = [f"HTTP {response.status_code}"]
    try:
        body = response.json()
    except ValueError:
        body = {}

    if isinstance(body, dict):
        for field in ("message", "error", "cause"):
            if field in body:
                details.append(
                    f"{field}: {redact_payment_diagnostic(body[field])}"
                )

    return f"{generic_message} Diagnóstico sandbox: {'; '.join(details)}."


def select_package(file_count: int) -> tuple[str, float]:
    """Retorna o pacote e o valor aplicáveis à quantidade de arquivos."""
    if not 1 <= file_count <= 300:
        raise ValueError("Envie entre 1 e 300 arquivos por lote.")

    if file_count <= 5:
        return "1 a 5 arquivos", 9.90
    if file_count <= 15:
        return "6 a 15 arquivos", 19.90
    if file_count <= 50:
        return "16 a 50 arquivos", 49.90
    if file_count <= 120:
        return "51 a 120 arquivos", 99.90

    return "121 a 300 arquivos", 199.90


def expected_payment_amount(file_count: int) -> Decimal:
    """Converte o preço do pacote em Decimal, sem comparação de float."""
    _, price = select_package(file_count)
    return Decimal(str(price)).quantize(Decimal("0.01"))


def payment_amount_matches(
    transaction_amount: object,
    expected_amount: Decimal,
) -> bool:
    """Compara valores monetários com precisão decimal."""
    if isinstance(transaction_amount, bool):
        return False

    try:
        amount = Decimal(str(transaction_amount))
    except (InvalidOperation, ValueError):
        return False

    return amount.is_finite() and amount == expected_amount


def _public_display_text(value: object) -> str:
    """Normaliza um valor escalar para uso apenas na amostra mascarada."""
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def mask_public_date(value: object) -> str:
    """Mantém somente o dia de uma data na prévia pública."""
    text = _public_display_text(value)
    match = re.match(r"(\d{1,2})/\d{1,2}/\d{2,4}", text)
    return f"{match.group(1)}/**/****" if match else "***"


def mask_public_value(value: object) -> str:
    """Oculta integralmente o valor monetário na prévia pública."""
    return "R$ **,**" if _public_display_text(value) else ""


def mask_public_name(value: object) -> str:
    """Mantém no máximo três caracteres de cada palavra de um nome."""
    return " ".join(
        f"{word[:3]}***"
        for word in _public_display_text(value).split()
    )


def mask_public_description(value: object) -> str:
    """Mostra somente duas palavras parcialmente mascaradas da descrição."""
    return " ".join(
        f"{word[:3]}***"
        for word in _public_display_text(value).split()[:2]
    )


def public_found_fields(dataframe: pd.DataFrame) -> list[str]:
    """Lista somente nomes dos campos que tiveram conteúdo extraído."""
    return [
        column
        for column in PUBLIC_FOUND_FIELD_COLUMNS
        if column in dataframe.columns
        and dataframe[column].map(_public_display_text).ne("").any()
    ]


def create_public_masked_preview(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Cria uma cópia segura de uma única linha para a interface pública."""
    available_columns = [
        column
        for column in PUBLIC_MASKED_PREVIEW_COLUMNS
        if column in dataframe.columns
    ]
    preview = dataframe.loc[:, available_columns].head(1).copy()

    if preview.empty:
        return preview

    if "Data" in preview:
        preview["Data"] = preview["Data"].map(mask_public_date)
    if "Valor" in preview:
        preview["Valor"] = preview["Valor"].map(mask_public_value)
    for column in ("Pagador", "Recebedor"):
        if column in preview:
            preview[column] = preview[column].map(mask_public_name)
    if "Descrição" in preview:
        preview["Descrição"] = preview["Descrição"].map(
            mask_public_description
        )

    return preview


def initialize_session_state(state: MutableMapping[str, object]) -> None:
    """Inicializa o estado usado pelos modos interno e público."""
    for key, value in PAYMENT_STATE_DEFAULTS.items():
        state.setdefault(key, value)


def reset_payment_state(
    state: MutableMapping[str, object],
    file_count: int | None = None,
) -> None:
    """Inicia o estado de pagamento de um novo lote."""
    state["external_reference"] = str(uuid.uuid4())
    state["preference_id"] = None
    state["payment_url"] = None
    state["payment_status"] = None
    state["payment_approved"] = False
    state["expected_payment_amount"] = (
        expected_payment_amount(file_count)
        if file_count is not None
        else None
    )


def create_payment_preference(
    external_reference: str,
    file_count: int,
) -> tuple[str, str]:
    """Cria uma preferência Checkout Pro sem expor detalhes sensíveis."""
    _, price = select_package(file_count)
    access_token = get_config_value(
        "MERCADO_PAGO_ACCESS_TOKEN",
        required=True,
    )
    environment = str(
        get_config_value("PAYMENT_ENVIRONMENT", default="production")
    ).lower()
    payload = {
        "items": [
            {
                "id": "organizacao-comprovantes",
                "title": "Organização de comprovantes",
                "quantity": 1,
                "currency_id": "BRL",
                "unit_price": price,
            }
        ],
        "external_reference": external_reference,
    }

    response: requests.Response | None = None
    try:
        response = requests.post(
            MERCADO_PAGO_PREFERENCES_URL,
            headers={"Authorization": f"Bearer {access_token}"},
            json=payload,
            timeout=PAYMENT_TIMEOUT_SECONDS,
        )
        response.raise_for_status()
        preference = response.json()
    except requests.HTTPError as error:
        failed_response = (
            error.response if error.response is not None else response
        )
        raise PaymentServiceError(
            payment_preference_error_message(failed_response, environment)
        ) from error
    except (requests.RequestException, ValueError) as error:
        raise PaymentServiceError(
            "Não foi possível iniciar o pagamento agora. Tente novamente."
        ) from error

    if not isinstance(preference, dict):
        raise PaymentServiceError(
            "O pagamento não pôde ser preparado. Tente novamente."
        )

    preference_id = preference.get("id")
    if environment in {"sandbox", "test", "teste"}:
        payment_url = (
            preference.get("sandbox_init_point")
            or preference.get("init_point")
        )
    else:
        payment_url = preference.get("init_point")

    if not isinstance(preference_id, str) or not isinstance(
        payment_url,
        str,
    ):
        raise PaymentServiceError(
            "O pagamento não pôde ser preparado. Tente novamente."
        )

    return preference_id, payment_url


def ensure_payment_preference(
    state: MutableMapping[str, object],
    file_count: int,
) -> tuple[str, str]:
    """Cria a preferência apenas se o lote ainda não possuir uma."""
    preference_id = state.get("preference_id")
    payment_url = state.get("payment_url")

    if not isinstance(state.get("expected_payment_amount"), Decimal):
        state["expected_payment_amount"] = expected_payment_amount(file_count)

    if isinstance(preference_id, str) and isinstance(payment_url, str):
        return preference_id, payment_url

    external_reference = state.get("external_reference")
    if not isinstance(external_reference, str):
        external_reference = str(uuid.uuid4())
        state["external_reference"] = external_reference

    preference_id, payment_url = create_payment_preference(
        external_reference,
        file_count,
    )
    state["preference_id"] = preference_id
    state["payment_url"] = payment_url
    state["payment_status"] = "pending"

    return preference_id, payment_url


def check_payment_status(
    external_reference: str,
    expected_amount: Decimal,
    expected_currency: str = PAYMENT_CURRENCY,
) -> str:
    """Libera somente pagamentos aprovados e idênticos ao pedido esperado."""
    access_token = get_config_value(
        "MERCADO_PAGO_ACCESS_TOKEN",
        required=True,
    )

    try:
        response = requests.get(
            MERCADO_PAGO_PAYMENTS_SEARCH_URL,
            headers={"Authorization": f"Bearer {access_token}"},
            params={"external_reference": external_reference},
            timeout=PAYMENT_TIMEOUT_SECONDS,
        )
        response.raise_for_status()
        result = response.json()
    except (requests.RequestException, ValueError) as error:
        raise PaymentServiceError(
            "Não foi possível consultar o pagamento agora. Tente novamente."
        ) from error

    if not isinstance(result, dict):
        raise PaymentServiceError(
            "A resposta do pagamento não pôde ser validada."
        )

    payments = result.get("results", [])
    if not isinstance(payments, list):
        raise PaymentServiceError(
            "A resposta do pagamento não pôde ser validada."
        )

    if any(
        isinstance(payment, dict)
        and payment.get("status") == "approved"
        and payment.get("external_reference") == external_reference
        and payment.get("currency_id") == expected_currency
        and payment_amount_matches(
            payment.get("transaction_amount"),
            expected_amount,
        )
        for payment in payments
    ):
        return "approved"

    return "pending" if payments else "not_found"


def refresh_payment_status(state: MutableMapping[str, object]) -> str:
    """Atualiza o estado local de liberação do download."""
    if state.get("payment_approved"):
        return "approved"

    external_reference = state.get("external_reference")
    expected_amount = state.get("expected_payment_amount")
    if not isinstance(external_reference, str) or not isinstance(
        expected_amount,
        Decimal,
    ):
        return "not_found"

    status = check_payment_status(
        external_reference,
        expected_amount,
        PAYMENT_CURRENCY,
    )
    state["payment_status"] = status
    state["payment_approved"] = status == "approved"
    return status


st.set_page_config(
    page_title="Organizador de Comprovantes em Excel | ComprovaFácil",
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


def process_documents(
    uploaded_files: list[st.runtime.uploaded_file_manager.UploadedFile],
) -> pd.DataFrame:
    """Executa a extração documental já existente para um lote enviado."""
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

    return pd.DataFrame(documents)


def download_excel_button(excel_file: bytes) -> None:
    """Exibe o download do Excel completo quando ele estiver liberado."""
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


def render_internal_mode(dataframe: pd.DataFrame) -> None:
    """Mantém a experiência integral destinada ao uso interno."""
    st.success(f"{len(dataframe)} documento(s) processado(s).")

    edited_dataframe = st.data_editor(
        dataframe,
        width="stretch",
        hide_index=True,
        num_rows="dynamic",
        key="documents_editor",
    )
    excel_file = generate_excel(edited_dataframe)
    st.session_state["excel_file"] = excel_file
    download_excel_button(excel_file)

    st.subheader("Prévia da organização")
    st.dataframe(
        edited_dataframe,
        width="stretch",
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


def render_payment_status() -> None:
    """Verifica pagamentos sem executar novamente o processamento de OCR."""
    state = st.session_state
    if state["payment_approved"]:
        return

    try:
        status = refresh_payment_status(state)
    except (KeyError, PaymentServiceError) as error:
        state["payment_status"] = "error"
        st.warning(str(error))
        return

    if status == "approved":
        st.success("Pagamento confirmado. Sua planilha completa foi liberada.")
    elif status == "pending":
        st.info("Pagamento ainda pendente de confirmação.")
    else:
        st.info("Ainda não localizamos um pagamento para este pedido.")


def render_public_mode(dataframe: pd.DataFrame) -> None:
    """Mostra somente a prévia permitida até a confirmação do pagamento."""
    file_count = len(dataframe)
    package_name, price = select_package(file_count)
    if not isinstance(st.session_state["expected_payment_amount"], Decimal):
        st.session_state["expected_payment_amount"] = (
            expected_payment_amount(file_count)
        )
    text_found_count = int(
        dataframe["Status da leitura"].eq("Texto encontrado").sum()
    )
    duplicate_count = int(
        dataframe["Possível duplicidade"].eq("Sim").sum()
    )

    st.success(f"{file_count} documento(s) processado(s).")
    metric_columns = st.columns(4)
    metric_columns[0].metric("Arquivos", file_count)
    metric_columns[1].metric("Com texto encontrado", text_found_count)
    metric_columns[2].metric("Possíveis duplicidades", duplicate_count)
    metric_columns[3].metric("Valor", f"R$ {price:.2f}".replace(".", ","))
    st.caption(f"Pacote selecionado: {package_name}")

    if st.session_state["payment_approved"]:
        st.success("Pagamento confirmado. Sua planilha completa está pronta.")
        excel_file = st.session_state["excel_file"]
        if isinstance(excel_file, bytes):
            download_excel_button(excel_file)
        return

    st.subheader("Amostra mascarada da organização")
    found_fields = public_found_fields(dataframe)
    if found_fields:
        st.caption("Campos encontrados: " + ", ".join(found_fields))
    else:
        st.caption("Nenhum campo estruturado foi encontrado.")

    masked_preview = create_public_masked_preview(dataframe)
    if not masked_preview.empty:
        st.dataframe(
            masked_preview,
            width="stretch",
            hide_index=True,
        )

    if not st.session_state["payment_url"]:
        st.info(
            "Não feche nem atualize esta página antes de concluir o "
            "pagamento e baixar sua planilha. O processamento permanece "
            "disponível somente durante esta sessão."
        )
        if st.button("Ir para pagamento", type="primary"):
            try:
                ensure_payment_preference(st.session_state, file_count)
            except KeyError:
                st.error(
                    "Não foi possível iniciar o pagamento agora. "
                    "Tente novamente."
                )
            except PaymentServiceError as error:
                st.error(str(error))
            else:
                st.rerun()

    payment_url = st.session_state["payment_url"]
    if isinstance(payment_url, str):
        st.info(
            "Não feche nem atualize esta página antes de concluir o "
            "pagamento e baixar sua planilha. O processamento permanece "
            "disponível somente durante esta sessão."
        )
        st.link_button("Abrir pagamento em nova aba", payment_url)
        st.caption(
            "Conclua o pagamento na nova aba e volte aqui para confirmar."
        )

        @st.fragment(run_every="15s")
        def payment_status_fragment() -> None:
            if st.button("Verificar pagamento agora"):
                render_payment_status()
            elif not st.session_state["payment_approved"]:
                render_payment_status()

            if st.session_state["payment_approved"]:
                st.rerun()

        payment_status_fragment()


def render_public_information() -> None:
    """Exibe orientações e avisos do fluxo público."""
    st.subheader("Como funciona")
    st.markdown(
        "1. Envie comprovantes em PDF, JPG ou PNG.\n"
        "2. Confira uma prévia protegida.\n"
        "3. Realize o pagamento pelo Mercado Pago.\n"
        "4. Baixe a planilha Excel organizada."
    )

    st.subheader("O que você recebe")
    st.write(
        "A planilha pode conter: data, valor, tipo de pagamento, pagador, "
        "recebedor, CPF ou CNPJ, identificador da transação, referência, "
        "descrição e possíveis duplicidades."
    )


def render_public_footer() -> None:
    """Exibe avisos de privacidade e suporte ao final do modo público."""
    st.divider()
    st.subheader("Privacidade")
    st.write(
        "Os arquivos são usados durante a sessão para gerar a planilha."
    )
    st.write(
        "O aplicativo não mantém um banco de dados com os documentos "
        "enviados."
    )

    st.subheader("Informações importantes")
    st.write(
        "A extração é automática e deve ser conferida antes do uso da "
        "planilha."
    )
    st.markdown(
    """
    ### Suporte

    Teve algum problema com o processamento, pagamento ou download?  
    Entre em contato pelo e-mail:

    **[comprovafacil@gmail.com](mailto:comprovafacil@gmail.com)**
    """
    )


def main() -> None:
    """Renderiza o modo interno ou público sem persistir arquivos enviados."""
    try:
        public_mode = parse_boolean(
            get_config_value("PUBLIC_MODE", default=False)
        )
    except ConfigurationError as error:
        st.error(str(error))
        st.stop()

    initialize_session_state(st.session_state)

    st.title("Organize seus comprovantes em uma planilha Excel")
    st.text(
    "Envie comprovantes Pix, recibos, PDFs, fotos ou prints e receba "
    "uma planilha Excel organizada com data, valor, pagador, recebedor, "
    "referência e possíveis duplicidades."
)

    if public_mode:
        render_public_information()

    uploaded_files = st.file_uploader(
        "Selecione os documentos",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
    )

    if not uploaded_files:
        st.info("Nenhum documento enviado ainda.")
    elif len(uploaded_files) > 300:
        st.error("O limite é de 300 arquivos por lote.")
    elif st.button("Ler documentos", type="primary"):
        dataframe = process_documents(uploaded_files)
        st.session_state["documents_dataframe"] = dataframe
        st.session_state["excel_file"] = generate_excel(dataframe)
        if public_mode:
            reset_payment_state(st.session_state, len(dataframe))
        st.rerun()

    if uploaded_files:
        dataframe = st.session_state["documents_dataframe"]
        if isinstance(dataframe, pd.DataFrame):
            if public_mode:
                render_public_mode(dataframe)
            else:
                render_internal_mode(dataframe)

    if public_mode:
        render_public_footer()


if __name__ == "__main__":
    main()

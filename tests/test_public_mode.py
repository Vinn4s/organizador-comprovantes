"""Testes locais do fluxo público sem comunicação externa real."""

from __future__ import annotations

import importlib.util
import subprocess
import unittest
from decimal import Decimal
from pathlib import Path
from unittest.mock import MagicMock, patch

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
APP_PATH = PROJECT_ROOT / "app" / "app.py"
SPEC = importlib.util.spec_from_file_location("organizador_app", APP_PATH)
APP = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(APP)


class PublicModeTests(unittest.TestCase):
    def _mocked_payment_status(
        self,
        payments: list[dict[str, object]],
        expected_amount: Decimal = Decimal("19.90"),
    ) -> str:
        response = MagicMock()
        response.json.return_value = {"results": payments}

        with (
            patch.object(
                APP.st,
                "secrets",
                {"MERCADO_PAGO_ACCESS_TOKEN": "fake-token"},
            ),
            patch.object(APP.requests, "get", return_value=response),
        ):
            return APP.check_payment_status(
                "expected-reference",
                expected_amount,
                "BRL",
            )

    @staticmethod
    def _payment(**overrides: object) -> dict[str, object]:
        payment: dict[str, object] = {
            "status": "approved",
            "external_reference": "expected-reference",
            "currency_id": "BRL",
            "transaction_amount": "19.90",
        }
        payment.update(overrides)
        return payment

    @staticmethod
    def _public_dataframe(row_count: int = 1) -> pd.DataFrame:
        return pd.DataFrame(
            {
                "Data": ["15/07/2026"] * row_count,
                "Valor": ["R$ 19,90"] * row_count,
                "Tipo": ["Pix"] * row_count,
                "Pagador": ["João da Silva"] * row_count,
                "Recebedor": ["Maria de Souza"] * row_count,
                "Descrição": ["Pagamento mensal aluguel comercial"] * row_count,
                "Possível duplicidade": ["Não"] * row_count,
                "Referência": ["referencia-real"] * row_count,
                "Identificador": ["identificador-real"] * row_count,
                "Documento": ["123.456.789-00"] * row_count,
                "Arquivo original": ["comprovante-real.pdf"] * row_count,
                "Texto extraído": ["texto confidencial completo"] * row_count,
                "Status da leitura": ["Texto encontrado"] * row_count,
            }
        )

    def test_a_package_price_boundaries(self) -> None:
        expected_packages = [
            (1, ("1 a 5 arquivos", 9.90)),
            (5, ("1 a 5 arquivos", 9.90)),
            (6, ("6 a 15 arquivos", 19.90)),
            (15, ("6 a 15 arquivos", 19.90)),
            (16, ("16 a 50 arquivos", 49.90)),
            (50, ("16 a 50 arquivos", 49.90)),
            (51, ("51 a 120 arquivos", 99.90)),
            (120, ("51 a 120 arquivos", 99.90)),
            (121, ("121 a 300 arquivos", 199.90)),
            (300, ("121 a 300 arquivos", 199.90)),
        ]

        for file_count, expected_package in expected_packages:
            with self.subTest(file_count=file_count):
                self.assertEqual(
                    APP.select_package(file_count),
                    expected_package,
                )

    def test_b_invalid_package_counts_are_rejected(self) -> None:
        for file_count in (0, 301):
            with self.subTest(file_count=file_count):
                with self.assertRaises(ValueError):
                    APP.select_package(file_count)

    def test_h_pending_payment_does_not_release_excel(self) -> None:
        state: dict[str, object] = {}
        APP.initialize_session_state(state)
        state["external_reference"] = "order-pending"
        state["expected_payment_amount"] = Decimal("19.90")

        with patch.object(APP, "check_payment_status", return_value="pending"):
            self.assertEqual(APP.refresh_payment_status(state), "pending")

        self.assertFalse(state["payment_approved"])

    def test_i_approved_payment_releases_excel(self) -> None:
        state: dict[str, object] = {}
        APP.initialize_session_state(state)
        state["external_reference"] = "order-approved"
        state["expected_payment_amount"] = Decimal("19.90")

        with patch.object(
            APP,
            "check_payment_status",
            return_value="approved",
        ) as check_status:
            self.assertEqual(APP.refresh_payment_status(state), "approved")

        self.assertTrue(state["payment_approved"])
        check_status.assert_called_once_with(
            "order-approved",
            Decimal("19.90"),
            "BRL",
        )

    def test_j_preference_is_not_recreated_on_reruns(self) -> None:
        state: dict[str, object] = {}
        APP.initialize_session_state(state)
        state["external_reference"] = "same-order"

        with patch.object(
            APP,
            "create_payment_preference",
            return_value=("preference-1", "https://checkout.example/pay"),
        ) as create_preference:
            APP.ensure_payment_preference(state, 10)
            self.assertEqual(state["preference_id"], "preference-1")
            self.assertEqual(
                state["payment_url"],
                "https://checkout.example/pay",
            )
            APP.ensure_payment_preference(state, 10)

        create_preference.assert_called_once_with("same-order", 10)
        self.assertEqual(state["preference_id"], "preference-1")
        self.assertEqual(
            state["payment_url"],
            "https://checkout.example/pay",
        )

    def test_k_new_lot_clears_previous_approval(self) -> None:
        state: dict[str, object] = {}
        APP.initialize_session_state(state)
        state.update(
            {
                "external_reference": "previous-order",
                "preference_id": "previous-preference",
                "payment_url": "https://checkout.example/previous",
                "payment_status": "approved",
                "payment_approved": True,
            }
        )

        APP.reset_payment_state(state, 16)

        self.assertNotEqual(state["external_reference"], "previous-order")
        self.assertIsNone(state["preference_id"])
        self.assertIsNone(state["payment_url"])
        self.assertFalse(state["payment_approved"])
        self.assertEqual(state["expected_payment_amount"], Decimal("49.90"))

    def test_l_internal_mode_keeps_immediate_download(self) -> None:
        dataframe = self._public_dataframe()
        fake_state: dict[str, object] = {}

        with (
            patch.object(APP.st, "session_state", fake_state),
            patch.object(APP.st, "success"),
            patch.object(APP.st, "subheader"),
            patch.object(APP.st, "dataframe"),
            patch.object(
                APP.st,
                "selectbox",
                return_value="comprovante-real.pdf",
            ),
            patch.object(APP.st, "text_area"),
            patch.object(
                APP.st,
                "data_editor",
                return_value=dataframe,
            ) as editor,
            patch.object(APP, "generate_excel", return_value=b"excel") as generate,
            patch.object(APP, "download_excel_button") as download,
        ):
            APP.render_internal_mode(dataframe)

        pd.testing.assert_frame_equal(editor.call_args.args[0], dataframe)
        generate.assert_called_once_with(dataframe)
        download.assert_called_once_with(b"excel")

    def test_single_file_public_sample_does_not_reveal_real_values(self) -> None:
        preview = APP.create_public_masked_preview(self._public_dataframe())
        rendered_preview = preview.to_json()

        for sensitive_value in (
            "15/07/2026",
            "R$ 19,90",
            "João da Silva",
            "Maria de Souza",
            "Pagamento mensal aluguel comercial",
        ):
            self.assertNotIn(sensitive_value, rendered_preview)

    def test_three_files_public_sample_has_only_one_row(self) -> None:
        preview = APP.create_public_masked_preview(self._public_dataframe(3))

        self.assertEqual(len(preview), 1)

    def test_public_sample_masks_names_dates_values_and_description(self) -> None:
        preview = APP.create_public_masked_preview(self._public_dataframe())
        sample = preview.iloc[0]

        self.assertEqual(sample["Data"], "15/**/****")
        self.assertEqual(sample["Valor"], "R$ **,**")
        self.assertEqual(sample["Pagador"], "Joã*** da*** Sil***")
        self.assertEqual(sample["Recebedor"], "Mar*** de*** Sou***")
        self.assertEqual(sample["Descrição"], "Pag*** men***")
        self.assertEqual(sample["Tipo"], "Pix")
        self.assertEqual(sample["Possível duplicidade"], "Não")

    def test_public_sample_excludes_sensitive_columns_and_keeps_original(self) -> None:
        dataframe = self._public_dataframe()
        original = dataframe.copy(deep=True)
        preview = APP.create_public_masked_preview(dataframe)

        self.assertNotIn("Referência", preview.columns)
        self.assertNotIn("Identificador", preview.columns)
        self.assertNotIn("Documento", preview.columns)
        self.assertNotIn("Arquivo original", preview.columns)
        self.assertNotIn("Texto extraído", preview.columns)
        found_fields = APP.public_found_fields(dataframe)
        self.assertNotIn("Referência", found_fields)
        self.assertNotIn("Identificador", found_fields)
        self.assertNotIn("Documento", found_fields)
        pd.testing.assert_frame_equal(dataframe, original)

    def test_public_mode_sends_only_masked_sample_to_dataframe_component(self) -> None:
        dataframe = self._public_dataframe()
        state: dict[str, object] = {}
        APP.initialize_session_state(state)
        columns = [MagicMock() for _ in range(4)]

        with (
            patch.object(APP.st, "session_state", state),
            patch.object(APP.st, "warning"),
            patch.object(APP.st, "success"),
            patch.object(APP.st, "columns", return_value=columns),
            patch.object(APP.st, "caption"),
            patch.object(APP.st, "subheader"),
            patch.object(APP.st, "button", return_value=False),
            patch.object(APP.st, "dataframe") as dataframe_component,
        ):
            APP.render_public_mode(dataframe)

        rendered_dataframe = dataframe_component.call_args.args[0]
        self.assertEqual(len(rendered_dataframe), 1)
        self.assertNotIn("R$ 19,90", rendered_dataframe.to_json())
        self.assertNotIn("Referência", rendered_dataframe.columns)
        pd.testing.assert_frame_equal(dataframe, self._public_dataframe())

    def test_m_real_secrets_file_is_not_tracked_by_git(self) -> None:
        result = subprocess.run(
            ["git", "ls-files", "--error-unmatch", ".streamlit/secrets.toml"],
            cwd=PROJECT_ROOT,
            capture_output=True,
            check=False,
            text=True,
        )
        self.assertNotEqual(result.returncode, 0)

    def test_payment_search_uses_mocked_approved_response(self) -> None:
        self.assertEqual(
            self._mocked_payment_status([self._payment()]),
            "approved",
        )

    def test_approved_payment_with_correct_details_releases_excel(self) -> None:
        self.assertEqual(
            self._mocked_payment_status([self._payment()]),
            "approved",
        )

    def test_approved_payment_with_lower_amount_does_not_release(self) -> None:
        self.assertEqual(
            self._mocked_payment_status(
                [self._payment(transaction_amount="1.00")]
            ),
            "pending",
        )

    def test_approved_payment_with_higher_amount_does_not_release(self) -> None:
        self.assertEqual(
            self._mocked_payment_status(
                [self._payment(transaction_amount="20.00")]
            ),
            "pending",
        )

    def test_approved_payment_with_different_currency_does_not_release(self) -> None:
        self.assertEqual(
            self._mocked_payment_status([self._payment(currency_id="USD")]),
            "pending",
        )

    def test_approved_payment_with_different_reference_does_not_release(self) -> None:
        self.assertEqual(
            self._mocked_payment_status(
                [self._payment(external_reference="another-reference")]
            ),
            "pending",
        )

    def test_pending_payment_with_correct_details_does_not_release(self) -> None:
        self.assertEqual(
            self._mocked_payment_status([self._payment(status="pending")]),
            "pending",
        )

    def test_one_valid_payment_among_multiple_results_releases_excel(self) -> None:
        self.assertEqual(
            self._mocked_payment_status(
                [
                    self._payment(transaction_amount="1.00"),
                    self._payment(),
                ]
            ),
            "approved",
        )

    def test_decimal_amount_19_9_matches_expected_19_90(self) -> None:
        self.assertEqual(
            self._mocked_payment_status(
                [self._payment(transaction_amount=19.9)]
            ),
            "approved",
        )

    def test_preference_uses_mocked_sandbox_checkout_url(self) -> None:
        response = MagicMock()
        response.json.return_value = {
            "id": "preference-1",
            "init_point": "https://prod.example/checkout",
            "sandbox_init_point": "https://sandbox.example/checkout",
        }

        with (
            patch.object(
                APP.st,
                "secrets",
                {
                    "MERCADO_PAGO_ACCESS_TOKEN": "fake-token",
                    "PAYMENT_ENVIRONMENT": "sandbox",
                },
            ),
            patch.object(APP.requests, "post", return_value=response) as post,
        ):
            preference_id, payment_url = APP.create_payment_preference(
                "order-1",
                10,
            )

        self.assertEqual(preference_id, "preference-1")
        self.assertEqual(payment_url, "https://sandbox.example/checkout")
        self.assertEqual(post.call_args.kwargs["timeout"], APP.PAYMENT_TIMEOUT_SECONDS)
        item = post.call_args.kwargs["json"]["items"][0]
        self.assertEqual(item["id"], "organizacao-comprovantes")
        self.assertEqual(
            set(item),
            {"id", "title", "quantity", "currency_id", "unit_price"},
        )

    def test_created_preference_renders_new_tab_link(self) -> None:
        dataframe = pd.DataFrame(
            {
                "Status da leitura": ["Texto encontrado"],
                "Possível duplicidade": ["Não"],
                "Data": ["01/01/2026"],
                "Valor": ["R$ 19,90"],
                "Tipo": ["Pix"],
                "Pagador": ["Pagador"],
                "Recebedor": ["Recebedor"],
                "Descrição": ["Teste"],
            }
        )
        state: dict[str, object] = {}
        APP.initialize_session_state(state)
        state.update(
            {
                "preference_id": "preference-1",
                "payment_url": "https://checkout.example/pay",
            }
        )
        columns = [MagicMock() for _ in range(4)]

        with (
            patch.object(APP.st, "session_state", state),
            patch.object(APP.st, "warning"),
            patch.object(APP.st, "success"),
            patch.object(APP.st, "columns", return_value=columns),
            patch.object(APP.st, "caption"),
            patch.object(APP.st, "subheader"),
            patch.object(APP.st, "dataframe"),
            patch.object(APP.st, "button", return_value=False),
            patch.object(APP.st, "link_button") as link_button,
            patch.object(
                APP.st,
                "fragment",
                side_effect=lambda **_kwargs: lambda function: function,
            ),
            patch.object(APP, "render_payment_status"),
        ):
            APP.render_public_mode(dataframe)

        link_button.assert_called_once_with(
            "Abrir pagamento em nova aba",
            "https://checkout.example/pay",
        )
        self.assertEqual(state["preference_id"], "preference-1")
        self.assertEqual(state["payment_url"], "https://checkout.example/pay")

    def test_sandbox_http_error_has_safe_diagnostic(self) -> None:
        response = MagicMock()
        response.status_code = 400
        response.json.return_value = {
            "message": "Authorization: Bearer fake-token",
            "error": "invalid_request",
            "cause": {
                "code": "invalid_item",
                "access_token": "fake-token",
            },
        }
        http_error = APP.requests.HTTPError("bad request")
        http_error.response = response
        response.raise_for_status.side_effect = http_error

        with (
            patch.object(
                APP.st,
                "secrets",
                {
                    "MERCADO_PAGO_ACCESS_TOKEN": "fake-token",
                    "PAYMENT_ENVIRONMENT": "sandbox",
                },
            ),
            patch.object(APP.requests, "post", return_value=response),
            self.assertRaises(APP.PaymentServiceError) as error_context,
        ):
            APP.create_payment_preference("order-1", 10)

        message = str(error_context.exception)
        self.assertIn("HTTP 400", message)
        self.assertIn("message:", message)
        self.assertIn("error: invalid_request", message)
        self.assertIn("cause:", message)
        self.assertNotIn("fake-token", message)
        self.assertNotIn("Authorization", message)
        self.assertNotIn("Access Token", message)

    def test_production_http_error_stays_generic(self) -> None:
        response = MagicMock()
        response.status_code = 400
        response.json.return_value = {
            "message": "detalhe interno",
            "error": "invalid_request",
        }
        http_error = APP.requests.HTTPError("bad request")
        http_error.response = response
        response.raise_for_status.side_effect = http_error

        with (
            patch.object(
                APP.st,
                "secrets",
                {
                    "MERCADO_PAGO_ACCESS_TOKEN": "fake-token",
                    "PAYMENT_ENVIRONMENT": "production",
                },
            ),
            patch.object(APP.requests, "post", return_value=response),
            self.assertRaises(APP.PaymentServiceError) as error_context,
        ):
            APP.create_payment_preference("order-1", 10)

        self.assertEqual(
            str(error_context.exception),
            "Não foi possível iniciar o pagamento agora. Tente novamente.",
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)

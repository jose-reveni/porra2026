import sys
from datetime import date
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import generate_dashboard as gd


@pytest.fixture(scope="module")
def generated_html():
    xlsx = Path(__file__).resolve().parent.parent / "Porra_Admin_v5_EN.xlsx"
    data = gd.parse_workbook(str(xlsx))
    analytics = gd.compute(data)
    return gd.render_html(analytics, date.today().strftime("%d/%m/%Y"))


class TestUserPickerSmoke:
    def test_html_contains_identity_keys(self, generated_html):
        assert "porra2026.currentUser" in generated_html
        assert "porra2026.pickerDismissed" in generated_html

    def test_html_contains_identity_functions(self, generated_html):
        assert "function loadMe(" in generated_html
        assert "function buildMeSummary(" in generated_html
        assert "function isMe(" in generated_html

    def test_html_contains_picker_ui(self, generated_html):
        assert "user-picker" in generated_html
        assert "ensureUserPicker" in generated_html
        assert "proto-user" in generated_html
        assert "data-action=\"open-user-picker\"" in generated_html

    def test_html_contains_personalization_hooks(self, generated_html):
        assert "meClass(" in generated_html
        assert "race-me-pin" in generated_html
        assert "me-summary" in generated_html
        assert "is-me" in generated_html

    def test_payload_has_participant_names(self, generated_html):
        assert '"matrix"' in generated_html
        assert '"names"' in generated_html

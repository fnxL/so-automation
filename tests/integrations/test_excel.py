import pytest
import os
from sotool.integrations.excel import ExcelClient

test_excel_path = os.path.abspath("tests/files/test_dispatch_report.xlsx")
test_macro_path = os.path.abspath("tests/files/test_macro.xlsm")


@pytest.fixture
def excel_client():
    with ExcelClient(visible=False) as excel:
        yield excel


def test_excel_connect(excel_client):
    assert excel_client.excel is not None


def test_excel_open_workbook(excel_client):
    assert excel_client.open(test_excel_path)


def test_excel_run_macro(excel_client):
    excel_client.open(test_macro_path)
    assert excel_client.run_macro("test_macro")


def test_excel_find_and_close_workbooks(excel_client):
    excel_client.open(test_excel_path)
    assert excel_client.find_and_close_workbooks(title_contains="dispatch")

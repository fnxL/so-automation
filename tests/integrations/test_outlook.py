from sotool.integrations.outook import OutlookClient
from sotool.integrations.excel import ExcelClient
import os

test_dispatch_report_path = "tests/files/test_dispatch_report.xlsx"


def test_outlook_client_connect():
    with OutlookClient() as outlook:
        assert outlook.app is not None
        assert outlook.main_window is not None


def test_create_draft_email_flow():
    # Copy test data to clipboard
    abs_path = os.path.abspath(test_dispatch_report_path)
    with ExcelClient() as excel:
        excel.open(abs_path)
        excel.copy_used_range()

    with OutlookClient() as outlook:
        outlook.create_draft_mail(
            to="test@example.com",
            cc="testcc@testcc.com",
            subject="Test Subject",
            body="Hi,\nOutlook client here, just minding my own business",
            paste_from_clipboard=True,
        )

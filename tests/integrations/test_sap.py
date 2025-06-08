import pytest
from sotool.integrations.sap_connector import SAPConnector


@pytest.fixture
def sap_session():
    with SAPConnector() as session:
        yield session


def test_sap_connect(sap_session):
    sap_session.findById("wnd[0]").maximize()
    assert sap_session.session is not None

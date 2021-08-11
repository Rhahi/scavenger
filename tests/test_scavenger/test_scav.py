import os
import pytest
import csv
from scavenger import scav

# @pytest.fixture(autouse=True)
# def LGCNS():
#     return os.path.join(scav.DOCROOT, "source", "CLIENT", "LG CNS", "LGCNS.xlsx")

def test_import_success():
    assert scav.import_success() == "Hello, world!"
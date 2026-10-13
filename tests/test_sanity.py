# tests/test_sanity.py
def test_pytest_is_working():
    assert 1 + 1 == 2

def test_can_import_project_modules():
    import models
    import io_import
    import setup
    import cycles
    assert True
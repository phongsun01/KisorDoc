import pytest
from pydantic import ValidationError
from kisorlib.engine import GenerateRequest


def test_generate_request_validation():
    # Test validation thành công
    req = GenerateRequest(
        option="Opt1",
        package_id="MS26-01",
        templates=["BaoCao", "TuTrinh"]
    )
    assert req.option == "Opt1"
    assert req.package_id == "MS26-01"
    assert req.templates == ["BaoCao", "TuTrinh"]
    assert req.dry_run is False

    # Test validation lỗi: thiếu option
    with pytest.raises(ValidationError):
        GenerateRequest(
            option="",
            package_id="MS26-01",
            templates=["BaoCao"]
        )

    # Test validation lỗi: template trống
    with pytest.raises(ValidationError):
        GenerateRequest(
            option="Opt1",
            package_id="MS26-01",
            templates=[]
        )

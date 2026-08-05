import pytest
import pandas as pd
from pathlib import Path
from kisorlib.config import AppConfig
from kisorlib.dataset import DataSet
from kisorlib.service import KisorService


class DummyDataSet(DataSet):
    def __init__(self, config: AppConfig):
        super().__init__(config)

    def _load(self):
        # Không tải file thực tế từ ổ đĩa để tránh phụ thuộc môi trường
        pass


@pytest.fixture
def dummy_service():
    config = AppConfig(
        ProjectPath="D:/Antigravity/KisorDoc",
        DataSheet="GoiThau",
        DefaultKeyId="ID",
        DefaultShow="{TT}. {GoiThau_ID}"
    )
    ds = DummyDataSet(config)
    
    # 1. Bảng GoiThau chính
    df_goi_thau = pd.DataFrame([
        {"TT": "1", "ID": "MS26-01", "GoiThau_ID": "MS26-01", "Name": "Goi thau 1", "Price": 1500000.0},
        {"TT": "2", "ID": "MS26-02", "GoiThau_ID": "MS26-02", "Name": "Goi thau 2", "Price": 2500000.0},
    ])
    ds.conn.register("GoiThau", df_goi_thau)
    ds.table_names.append("GoiThau")
    
    # 2. Bảng Options cấu hình
    df_options = pd.DataFrame([
        {"Key": "Opt1", "Value": "Option 1", "Sheet": "GoiThau", "Show": "{TT}. {Name}", "Type": "Normal", "KeyId": "ID"},
        {"Key": "Opt2", "Value": "Option 2 (Repeat)", "Sheet": "GoiThau <* TCGTTD @ ID = MemberID", "Show": "{TT}. {Name} | {MemberName}", "Type": "Repeat", "KeyId": "ID | CCCD"},
    ])
    ds.conn.register("Options", df_options)
    ds.table_names.append("Options")

    # 3. Bảng Workflow liên kết
    df_workflow = pd.DataFrame([
        {"Option": "Opt1", "Name": "TemplateA", "Condition": "Price > 1000000"},
        {"Option": "Opt2", "Name": "TemplateB", "Condition": "ALL"},
    ])
    ds.conn.register("Workflow", df_workflow)
    ds.table_names.append("Workflow")

    # 4. Bảng thành viên tổ chuyên gia gốc (dùng cho Repeat mode)
    df_tcgttd = pd.DataFrame([
        {"MemberID": "MS26-01", "CCCD": "123456", "MemberName": "Nguyen Van A", "Role": "Leader"},
        {"MemberID": "MS26-01", "CCCD": "789012", "MemberName": "Tran Thi B", "Role": "Member"},
    ])
    ds.conn.register("TCGTTD_Goc", df_tcgttd)
    ds.table_names.append("TCGTTD_Goc")

    return KisorService(config, ds)


def test_service_get_options(dummy_service):
    opts = dummy_service.get_options()
    assert len(opts) == 2
    assert any("Opt1" in o for o in opts)
    assert any("Opt2" in o for o in opts)


def test_service_get_option_config(dummy_service):
    cfg1 = dummy_service.get_option_config("Opt1")
    assert cfg1["type"] == "Normal"
    cfg2 = dummy_service.get_option_config("Opt2")
    assert cfg2["type"] == "Repeat"
    assert cfg2["key_id"] == "ID | CCCD"


def test_service_get_repeat_members(dummy_service):
    members = dummy_service.get_repeat_members("MS26-01", "Group1", "Opt2")
    assert len(members) == 2
    assert "Nguyen Van A" in members
    assert "Tran Thi B" in members


def test_service_register_temporary_tcgttd(dummy_service):
    success = dummy_service.register_temporary_tcgttd("MS26-01", ["Nguyen Van A"], "Group1", "ID | CCCD", "Opt2")
    assert success is True
    
    df_temp = dummy_service.ds.conn.execute('SELECT * FROM TCGTTD').fetchdf()
    assert len(df_temp) == 1
    assert df_temp.iloc[0]["MemberName"] == "Nguyen Van A"


def test_service_run_preview_composite_key(dummy_service):
    # Đăng ký thành viên tạm trước khi preview
    dummy_service.register_temporary_tcgttd("MS26-01", ["Nguyen Van A"], "Group1", "ID | CCCD", "Opt2")
    # Preview hoạt động chính xác với composite key
    res = dummy_service.run_preview("Opt2", "1. Goi thau 1", ["TemplateB"])
    assert res.startswith("✅ Context:")
    assert "Tables:" in res

import os
import pytest
from lxml import etree
from openpyxl import Workbook
from migrate_text_to_placeholders import (
    to_slug,
    load_mapping,
    reconstruct_paragraph_text,
    find_forbidden_spans,
    safe_replace_xml,
    NAMESPACES
)

def test_to_slug():
    assert to_slug("Họ và Tên") == "Ho_va_Ten"
    assert to_slug("Mã gói") == "Ma_goi"
    assert to_slug("Số điện thoại") == "So_dien_thoai"
    assert to_slug("") == ""
    assert to_slug(None) == ""

def test_find_forbidden_spans():
    text = "Xin chào {{ Ten_Khach_Hang }} và {{ So_Hop_Dong }}."
    spans = find_forbidden_spans(text)
    assert len(spans) == 2
    assert text[spans[0][0]:spans[0][1]] == "{{ Ten_Khach_Hang }}"
    assert text[spans[1][0]:spans[1][1]] == "{{ So_Hop_Dong }}"

def test_reconstruct_paragraph_text():
    # Build w:p XML segment
    xml_str = """
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:r><w:t>Nguyễn </w:t></w:r>
        <w:r><w:t>Văn </w:t></w:r>
        <w:r><w:t>A</w:t></w:r>
    </w:p>
    """
    p_node = etree.fromstring(xml_str)
    full_text, offset_map = reconstruct_paragraph_text(p_node)
    assert full_text == "Nguyễn Văn A"
    assert len(offset_map) == 12
    # Check that character at index 8 ("V") belongs to the second run (Văn )
    assert offset_map[8][2].text == "Văn "
    assert offset_map[8][1] == 1  # Index of 'ă' in 'Văn ' is 1

def test_safe_replace_xml_single_run():
    xml_str = """
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:r><w:t>Tên tôi là Nguyễn Văn A và tôi sống ở Hà Nội.</w:t></w:r>
    </w:p>
    """
    p_node = etree.fromstring(xml_str)
    mapping = [("Nguyễn Văn A", "{{ Ten }}")]
    file_logs = []
    
    modified = safe_replace_xml(p_node, mapping, case_insensitive=False, dense_threshold=15, file_logs=file_logs, file_path="test.docx")
    assert modified is True
    
    # Re-read text
    t_nodes = p_node.xpath('.//w:t', namespaces=NAMESPACES)
    assert len(t_nodes) == 1
    assert t_nodes[0].text == "Tên tôi là {{ Ten }} và tôi sống ở Hà Nội."
    assert len(file_logs) == 1
    assert file_logs[0]['matches'][0]['sample'] == "Nguyễn Văn A"

def test_safe_replace_xml_multi_run():
    xml_str = """
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:r><w:t>Nguyễn </w:t></w:r>
        <w:r><w:t>Văn </w:t></w:r>
        <w:r><w:t>A và Hà Nội</w:t></w:r>
    </w:p>
    """
    p_node = etree.fromstring(xml_str)
    mapping = [("Nguyễn Văn A", "{{ Ten }}")]
    file_logs = []
    
    modified = safe_replace_xml(p_node, mapping, case_insensitive=False, dense_threshold=15, file_logs=file_logs, file_path="test.docx")
    assert modified is True
    
    # Run 1: "Nguyễn " -> "{{ Ten }}"
    # Run 2: "Văn " -> ""
    # Run 3: "A và Hà Nội" -> " và Hà Nội" (A is trimmed)
    t_nodes = p_node.xpath('.//w:t', namespaces=NAMESPACES)
    assert t_nodes[0].text == "{{ Ten }}"
    assert t_nodes[1].text == ""
    assert t_nodes[2].text == " và Hà Nội"

def test_safe_replace_xml_idempotent():
    xml_str = """
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:r><w:t>Chào {{ Nguyễn Văn A }} và Nguyễn Văn A.</w:t></w:r>
    </w:p>
    """
    p_node = etree.fromstring(xml_str)
    mapping = [("Nguyễn Văn A", "{{ Ten }}")]
    file_logs = []
    
    modified = safe_replace_xml(p_node, mapping, case_insensitive=False, dense_threshold=15, file_logs=file_logs, file_path="test.docx")
    assert modified is True
    
    # Only the second "Nguyễn Văn A" should be replaced, the first is inside {{ }}
    t_nodes = p_node.xpath('.//w:t', namespaces=NAMESPACES)
    assert t_nodes[0].text == "Chào {{ Nguyễn Văn A }} và {{ Ten }}."

def test_load_mapping_mock_excel(tmp_path):
    excel_file = tmp_path / "test_data.xlsx"
    wb = Workbook()
    
    # Data sheet
    ws = wb.active
    ws.title = "Data"
    ws.append(["Họ và Tên", "Ngày sinh", "Số tiền"])
    ws.append(["Nguyễn Văn A", "01/01/2000", "1.000.000"])
    
    # Config sheet
    c_ws = wb.create_sheet(title="Config")
    c_ws.append(["Key", "Value"])
    c_ws.append(["Ten_Khach_Hang", "Họ và Tên"])
    c_ws.append(["Ngay_Sinh.Date", "Ngày sinh"])
    
    wb.save(excel_file)
    
    sorted_mapping, collisions = load_mapping(
        excel_path=str(excel_file),
        row_idx=1,
        sheet_name="Data",
        config_sheet_name="Config"
    )
    
    # Check that "Họ và Tên" maps to "{{ Ten_Khach_Hang }}"
    # Check that "Ngày sinh" maps to "{{ Ngay_Sinh_Date }}" (clean_config_key converted .Date to _Date)
    # Check that "Số tiền" fallback slug is "{{ So_tien }}"
    mapping_dict = dict(sorted_mapping)
    assert mapping_dict["Nguyễn Văn A"] == "{{ Ten_Khach_Hang }}"
    assert mapping_dict["01/01/2000"] == "{{ Ngay_Sinh_Date }}"
    assert mapping_dict["1.000.000"] == "{{ So_tien }}"

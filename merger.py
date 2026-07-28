import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from lxml import etree

from filters import filter_date, filter_date_long, filter_number

NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

PLACEHOLDER_RE = re.compile(r"<<([A-Za-z0-9_]+(?:\.[A-Za-z]+)*)>>")

MODIFIER_MAP = {
    "Date.Long": filter_date_long,
    "Date": filter_date,
    "Upper": lambda v: str(v).upper() if v else "",
    "Number": filter_number,
}

TABLE_PLACEHOLDER_KEYS = {"DanhMuc", "DanhMucKoGia"}

XML_FILES = [
    "word/document.xml",
    "word/header1.xml",
    "word/header2.xml",
    "word/header3.xml",
    "word/footer1.xml",
    "word/footer2.xml",
    "word/footer3.xml",
]


def mail_merge(template_path: Path, context: dict, output_path: Path):
    with open(template_path, "rb") as f:
        data = f.read()

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp_path = Path(tmp.name)

    try:
        with zipfile.ZipFile(template_path, "r") as zin:
            zip_contents = {name: zin.read(name) for name in zin.namelist()}

        for xml_path in XML_FILES:
            if xml_path not in zip_contents:
                continue
            xml_bytes = zip_contents[xml_path]
            root = etree.fromstring(xml_bytes)
            modified = _replace_in_paragraphs(root, context)
            if modified:
                zip_contents[xml_path] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

        temp_out = output_path.with_suffix(".tmp.docx")
        with zipfile.ZipFile(temp_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data2 in zip_contents.items():
                zout.writestr(name, data2)

        shutil.move(str(temp_out), str(output_path))

    finally:
        if tmp_path.exists():
            tmp_path.unlink()


def _replace_in_paragraphs(root, context: dict) -> bool:
    modified = False
    for para in root.iter(f"{{{NS}}}p"):
        merged = _get_paragraph_text(para)
        if not merged:
            continue
        if not PLACEHOLDER_RE.search(merged):
            continue
        if any(p in merged for p in TABLE_PLACEHOLDER_KEYS):
            continue
        new_text = _replace_placeholders(merged, context)
        if new_text != merged:
            _set_paragraph_text(para, new_text)
            modified = True
    return modified


def _get_paragraph_text(para) -> str:
    texts = []
    for t in para.iter(f"{{{NS}}}t"):
        if t.text:
            texts.append(t.text)
    return "".join(texts)


def _set_paragraph_text(para, new_text: str):
    runs = list(para.iter(f"{{{NS}}}r"))
    if not runs:
        return
    t_elements = list(runs[0].iter(f"{{{NS}}}t"))
    if t_elements:
        first_t = t_elements[0]
    else:
        first_t = etree.SubElement(runs[0], f"{{{NS}}}t")
    first_t.text = new_text
    for run in runs[1:]:
        para.remove(run)


def _replace_placeholders(text: str, context: dict) -> str:
    def replacer(m: re.Match):
        full_key = m.group(1)
        value = _resolve_value(full_key, context)
        return str(value) if value else ""

    return PLACEHOLDER_RE.sub(replacer, text)


def _resolve_value(full_key: str, context: dict):
    if full_key in context:
        return context[full_key]

    parts = full_key.split(".")
    for i in range(len(parts) - 1, 0, -1):
        base = ".".join(parts[:i])
        stripped = ".".join(parts[i:])
        if base in context:
            if stripped in MODIFIER_MAP:
                try:
                    return MODIFIER_MAP[stripped](context[base])
                except Exception:
                    return context[base]
            return context[base]

    return ""

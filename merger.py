import jinja2
from docxtpl import DocxTemplate
from pathlib import Path

from filters import filter_date, filter_date_long, filter_number

def mail_merge(template_path: Path, context: dict, output_path: Path):
    # Load template using docxtpl
    doc = DocxTemplate(str(template_path))
    
    # Initialize Jinja2 Environment and register custom filters
    jenv = jinja2.Environment(undefined=jinja2.DebugUndefined)
    jenv.filters["date"] = filter_date
    jenv.filters["date_long"] = filter_date_long
    jenv.filters["number"] = filter_number
    
    # Render template
    doc.render(context, jenv)
    
    # Save output
    doc.save(str(output_path))


def mail_merge_safe(template_path, context: dict, output_path) -> tuple[bool, str]:
    """
    Version co error handling — khong corrupt file neu Jinja2 loi.
    Returns (success: bool, error_message: str)
    """
    import tempfile, shutil
    tmp = Path(tempfile.mktemp(suffix=".docx"))
    try:
        doc = DocxTemplate(str(template_path))
        jenv = jinja2.Environment(undefined=jinja2.DebugUndefined)
        jenv.filters["date"]      = filter_date
        jenv.filters["date_long"] = filter_date_long
        jenv.filters["number"]    = filter_number
        doc.render(context, jenv)
        doc.save(str(tmp))
        shutil.move(str(tmp), str(output_path))
        return True, ""
    except Exception as e:
        if tmp.exists():
            tmp.unlink()
        return False, str(e)
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except Exception:
                pass

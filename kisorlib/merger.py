import jinja2
from docxtpl import DocxTemplate
from pathlib import Path
from datetime import datetime
from .filters import (
    filter_date, filter_date_long, filter_number, filter_num2text,
    filter_day, filter_month, filter_year, filter_add_days,
    filter_add_months, filter_date_diff, filter_quarter,
    filter_weekday, filter_date_text
)


def mail_merge(template_path: Path, context: dict, output_path: Path):
    context = dict(context)
    context.setdefault("now", datetime.now())

    # Load template using docxtpl
    doc = DocxTemplate(str(template_path))
    
    # Initialize Jinja2 Environment and register custom filters
    jenv = jinja2.Environment(undefined=jinja2.DebugUndefined)
    jenv.filters["date"] = filter_date
    jenv.filters["date_long"] = filter_date_long
    jenv.filters["number"] = filter_number
    jenv.filters["num2text"] = filter_num2text
    jenv.filters["day"] = filter_day
    jenv.filters["month"] = filter_month
    jenv.filters["year"] = filter_year
    jenv.filters["add_days"] = filter_add_days
    jenv.filters["add_months"] = filter_add_months
    jenv.filters["date_diff"] = filter_date_diff
    jenv.filters["quarter"] = filter_quarter
    jenv.filters["weekday"] = filter_weekday
    jenv.filters["date_text"] = filter_date_text
    
    # Render template
    doc.render(context, jenv)
    
    # Save output
    doc.save(str(output_path))


def mail_merge_safe(template_path, context: dict, output_path) -> tuple[bool, str]:
    """
    Version co error handling — khong corrupt file neu Jinja2 loi.
    Returns (success: bool, error_message: str)
    """
    context = dict(context)
    context.setdefault("now", datetime.now())

    import tempfile, shutil
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        tmp = Path(f.name)
    try:
        doc = DocxTemplate(str(template_path))
        jenv = jinja2.Environment(undefined=jinja2.DebugUndefined)
        jenv.filters["date"]      = filter_date
        jenv.filters["date_long"] = filter_date_long
        jenv.filters["number"]    = filter_number
        jenv.filters["num2text"]  = filter_num2text
        jenv.filters["day"] = filter_day
        jenv.filters["month"] = filter_month
        jenv.filters["year"] = filter_year
        jenv.filters["add_days"] = filter_add_days
        jenv.filters["add_months"] = filter_add_months
        jenv.filters["date_diff"] = filter_date_diff
        jenv.filters["quarter"] = filter_quarter
        jenv.filters["weekday"] = filter_weekday
        jenv.filters["date_text"] = filter_date_text
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

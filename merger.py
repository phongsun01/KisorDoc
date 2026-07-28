import jinja2
from docxtpl import DocxTemplate
from pathlib import Path

from filters import filter_date, filter_date_long, filter_number

def mail_merge(template_path: Path, context: dict, output_path: Path):
    # Load template using docxtpl
    doc = DocxTemplate(str(template_path))
    
    # Initialize Jinja2 Environment and register custom filters
    jenv = jinja2.Environment()
    jenv.filters["date"] = filter_date
    jenv.filters["date_long"] = filter_date_long
    jenv.filters["number"] = filter_number
    
    # Render template
    doc.render(context, jenv)
    
    # Save output
    doc.save(str(output_path))

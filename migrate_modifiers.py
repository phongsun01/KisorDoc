import os
import zipfile
import re
import shutil
from pathlib import Path

TEMPLATES_DIR = r"D:\Antigravity\1. Tu dong xu ly tai lieu\2. Templates"

# Regex matches placeholders starting with << or {{ and ending with >> or }}
# Example: <<NgayKy.Date>> or {{NgayKy.Date.Long}} or <<GiaTri.Number>> or {{GiaTri.Chu}}
PLACEHOLDER_RE = re.compile(r"(&lt;&lt;|\{\{)(.*?)(\&gt;&gt;|\}\})", re.DOTALL)

def clean_tags(xml_segment):
    # Remove XML tags inside placeholders
    return re.sub(r"<[^>]+>", "", xml_segment)

def map_modifier(content_str):
    val = content_str.strip()
    
    # 1. Tra cứu và chuẩn hoá modifier
    # Mẫu ngày tháng
    if val.endswith(".Date.Long") or val.endswith(".date_long") or val.endswith(".Date.long"):
        base = re.sub(r"\.Date\.Long$|\.date_long$|\.Date\.long$", "", val)
        return f"{{{{{base.strip()}|date_long}}}}"
        
    elif val.endswith(".Date") or val.endswith(".date"):
        base = re.sub(r"\.Date$|\.date$", "", val)
        return f"{{{{{base.strip()}|date}}}}"
        
    elif val.endswith(".Day") or val.endswith(".day"):
        base = re.sub(r"\.Day$|\.day$", "", val)
        return f"{{{{{base.strip()}|day}}}}"
        
    elif val.endswith(".Month") or val.endswith(".month"):
        base = re.sub(r"\.Month$|\.month$", "", val)
        return f"{{{{{base.strip()}|month}}}}"
        
    elif val.endswith(".Year") or val.endswith(".year"):
        base = re.sub(r"\.Year$|\.year$", "", val)
        return f"{{{{{base.strip()}|year}}}}"
        
    # Mẫu số và chữ
    elif val.endswith(".Number") or val.endswith(".number"):
        base = re.sub(r"\.Number$|\.number$", "", val)
        return f"{{{{{base.strip()}|number}}}}"
        
    elif val.endswith(".Chu") or val.endswith(".chu") or val.endswith(".Text") or val.endswith(".text"):
        base = re.sub(r"\.Chu$|\.chu$|\.Text$|\.text$", "", val)
        return f"{{{{{base.strip()}|num2text}}}}"
        
    elif val.endswith(".Upper") or val.endswith(".upper"):
        base = re.sub(r"\.Upper$|\.upper$", "", val)
        return f"{{{{{base.strip()}|upper}}}}"
        
    # Nếu đã có filter dạng | thì giữ nguyên hoặc chuẩn hoá
    elif "|" in val:
        parts = val.split("|")
        base = parts[0].strip()
        mod = parts[1].strip().lower()
        if mod == "date_long":
            return f"{{{{{base}|date_long}}}}"
        elif mod == "date":
            return f"{{{{{base}|date}}}}"
        elif mod == "number":
            return f"{{{{{base}|number}}}}"
        elif mod == "num2text":
            return f"{{{{{base}|num2text}}}}"
        elif mod == "day":
            return f"{{{{{base}|day}}}}"
        elif mod == "month":
            return f"{{{{{base}|month}}}}"
        elif mod == "year":
            return f"{{{{{base}|year}}}}"
        elif mod == "upper":
            return f"{{{{{base}|upper}}}}"
        return f"{{{{{val}}}}}"
        
    # Không có modifier -> convert sang double braces mặc định
    return f"{{{{{val}}}}}"

def migrate_xml(xml_content):
    def replace_placeholder(match):
        prefix = match.group(1)
        segment = match.group(2)
        suffix = match.group(3)
        
        # Làm sạch thẻ XML bị xen ngang ở giữa placeholder
        clean_text = clean_tags(segment)
        return map_modifier(clean_text)

    return PLACEHOLDER_RE.sub(replace_placeholder, xml_content)

def migrate_file(filepath):
    # Tạo thư mục bak nếu chưa có
    file_dir = os.path.dirname(filepath)
    bak_dir = os.path.join(file_dir, "bak")
    os.makedirs(bak_dir, exist_ok=True)
    
    file_name = os.path.basename(filepath)
    backup_path = os.path.join(bak_dir, file_name.replace(".docx", ".bak.docx"))
    
    # Backup nếu chưa tồn tại
    if not os.path.exists(backup_path):
        shutil.copy2(filepath, backup_path)
        print(f"Backed up: {file_name} -> bak/")

    temp_path = filepath + ".tmp"
    try:
        with zipfile.ZipFile(filepath, 'r') as yin:
            with zipfile.ZipFile(temp_path, 'w') as yout:
                for item in yin.infolist():
                    data = yin.read(item.filename)
                    if item.filename.startswith('word/') and item.filename.endswith('.xml'):
                        xml_content = data.decode('utf-8')
                        migrated = migrate_xml(xml_content)
                        yout.writestr(item.filename, migrated.encode('utf-8'))
                    else:
                        yout.writestr(item, data)
        
        os.remove(filepath)
        os.rename(temp_path, filepath)
        print(f"Migrated: {file_name}")
    except Exception as e:
        print(f"Error migrating {file_name}: {e}")
        if os.path.exists(temp_path):
            os.remove(temp_path)

def main():
    print(f"Starting template modifier migration in: {TEMPLATES_DIR}")
    migrated_count = 0
    for root, dirs, files in os.walk(TEMPLATES_DIR):
        # Tránh quét các file trong thư mục bak
        if "bak" in root.split(os.sep):
            continue
        for file in files:
            if file.endswith(".docx") and not file.endswith(".bak.docx"):
                filepath = os.path.join(root, file)
                migrate_file(filepath)
                migrated_count += 1
    print(f"Completed! Total files processed: {migrated_count}")

if __name__ == "__main__":
    main()

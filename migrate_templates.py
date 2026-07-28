import os
import zipfile
import re
import shutil

TEMPLATES_DIR = r"D:\Antigravity\1. Thanh toan nho\2. Templates"
PATTERN_LT_GT = re.compile(r"&lt;&lt;((?:(?!&gt;&gt;).)*)&gt;&gt;", re.DOTALL)
PATTERN_TABLE = re.compile(r"(?<!\{)\{((?:(?!\}).)*?)(DanhMucKoGia|DanhMuc)((?:(?!\}).)*?)\}(?!\})", re.DOTALL)

def clean_placeholder_text(xml_segment):
    return re.sub(r"<[^>]+>", "", xml_segment)

def map_placeholder(placeholder_text):
    val = placeholder_text.strip()
    if val.endswith(".Date.Long"):
        base = val[:-10]
        return f"{{{{{base}_Date|date_long}}}}"
    elif val.endswith(".Date"):
        base = val[:-5]
        return f"{{{{{base}_Date|date}}}}"
    elif val.endswith(".Upper"):
        base = val[:-6]
        return f"{{{{{base}|upper}}}}"
    elif val.endswith(".Number"):
        base = val[:-7]
        return f"{{{{{base}|number}}}}"
    else:
        return f"{{{{{val}}}}}"

def migrate_xml(xml_content):
    # 1. Migrate <<>> to {{}}
    def replace_lt_gt(match):
        segment = match.group(1)
        clean_text = clean_placeholder_text(segment)
        replacement = map_placeholder(clean_text)
        return replacement

    content = PATTERN_LT_GT.sub(replace_lt_gt, xml_content)
    
    # 2. Migrate single-brace table placeholders to double-brace
    def replace_table(match):
        placeholder = match.group(2)
        return f"{{{{{placeholder}}}}}"

    content = PATTERN_TABLE.sub(replace_table, content)
    return content

def migrate_file(filepath):
    # Ensure backup exists
    backup_path = filepath.replace(".docx", ".bak.docx")
    if not os.path.exists(backup_path):
        shutil.copy2(filepath, backup_path)
        print(f"Backed up: {os.path.basename(filepath)}")

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
        
        # Replace original with migrated
        os.remove(filepath)
        os.rename(temp_path, filepath)
        print(f"Migrated: {os.path.basename(filepath)}")
    except Exception as e:
        print(f"Error migrating {filepath}: {e}")
        if os.path.exists(temp_path):
            os.remove(temp_path)

def main():
    print("Starting template migration...")
    migrated_count = 0
    for root, dirs, files in os.walk(TEMPLATES_DIR):
        for file in files:
            if file.endswith(".docx") and not file.endswith(".bak.docx"):
                filepath = os.path.join(root, file)
                migrate_file(filepath)
                migrated_count += 1
    print(f"Completed! Total files processed: {migrated_count}")

if __name__ == "__main__":
    main()

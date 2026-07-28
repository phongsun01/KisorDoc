import os
import zipfile
import re
import shutil

TEMPLATES_DIR = r"D:\Antigravity\1. Thanh toan nho\2. Templates"

# Pattern to match single-brace {DanhMuc} and {DanhMucKoGia}
# Allowing optional XML tags between braces and word
PATTERN = re.compile(r"(?<!\{)\{((?:(?!\}).)*?)(DanhMucKoGia|DanhMuc)((?:(?!\}).)*?)\}(?!\})", re.DOTALL)

def migrate_xml(xml_content):
    def replace_match(match):
        placeholder = match.group(2)
        return f"{{{{{placeholder}}}}}"

    return PATTERN.sub(replace_match, xml_content)

def migrate_file(filepath):
    # Ensure backup exists (it should, but just in case)
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
        print(f"Migrated table placeholder: {os.path.basename(filepath)}")
    except Exception as e:
        print(f"Error migrating {filepath}: {e}")
        if os.path.exists(temp_path):
            os.remove(temp_path)

def main():
    print("Starting single-brace table placeholder migration...")
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

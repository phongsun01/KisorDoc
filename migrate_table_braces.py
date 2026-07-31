import os
import zipfile
import re
import shutil
from pathlib import Path

TEMPLATES_DIR = r"D:\Antigravity\1. Tu dong xu ly tai lieu\2. Templates"

# Regex matches single braced alphanumeric strings, e.g. {KHLCNT} or {DanhMuc}
# but avoids matching double braces like {{KHLCNT}}
SINGLE_BRACE_RE = re.compile(r"(?<!\{)\{([a-zA-Z0-9_]+)\}(?!\})")

def migrate_xml(xml_content):
    def replace_brace(match):
        placeholder = match.group(1)
        print(f"  Found single brace table placeholder: {{{placeholder}}} -> {{{{ {placeholder} }}}}")
        return f"{{{{{placeholder}}}}}"

    return SINGLE_BRACE_RE.sub(replace_brace, xml_content)

def migrate_file(filepath):
    file_dir = os.path.dirname(filepath)
    bak_dir = os.path.join(file_dir, "bak")
    os.makedirs(bak_dir, exist_ok=True)
    
    file_name = os.path.basename(filepath)
    backup_path = os.path.join(bak_dir, file_name.replace(".docx", ".bak.docx"))
    
    # Backup if not exists
    if not os.path.exists(backup_path):
        shutil.copy2(filepath, backup_path)
        print(f"Backed up: {file_name} -> bak/")

    temp_path = filepath + ".tmp"
    try:
        has_change = False
        with zipfile.ZipFile(filepath, 'r') as yin:
            with zipfile.ZipFile(temp_path, 'w') as yout:
                for item in yin.infolist():
                    data = yin.read(item.filename)
                    if item.filename.startswith('word/') and item.filename.endswith('.xml'):
                        xml_content = data.decode('utf-8')
                        
                        # Check if matches exist
                        if SINGLE_BRACE_RE.search(xml_content):
                            has_change = True
                            xml_content = migrate_xml(xml_content)
                            
                        yout.writestr(item.filename, xml_content.encode('utf-8'))
                    else:
                        yout.writestr(item, data)
        
        os.remove(filepath)
        os.rename(temp_path, filepath)
        if has_change:
            print(f"Migrated table braces: {file_name}")
        else:
            print(f"No braces to migrate in: {file_name}")
    except Exception as e:
        print(f"Error migrating {file_name}: {e}")
        if os.path.exists(temp_path):
            os.remove(temp_path)

def main():
    print(f"Starting table braces migration in: {TEMPLATES_DIR}")
    migrated_count = 0
    # Run on Opt3, Opt4, Opt5
    target_folders = ["Opt3", "Opt4", "Opt5"]
    for folder in target_folders:
        folder_path = os.path.join(TEMPLATES_DIR, folder)
        if not os.path.exists(folder_path):
            continue
        print(f"\nScanning folder: {folder}")
        for root, dirs, files in os.walk(folder_path):
            if "bak" in root.split(os.sep):
                continue
            for file in files:
                if file.endswith(".docx") and not file.endswith(".bak.docx"):
                    filepath = os.path.join(root, file)
                    migrate_file(filepath)
                    migrated_count += 1
    print(f"\nCompleted! Total files processed: {migrated_count}")

if __name__ == "__main__":
    main()

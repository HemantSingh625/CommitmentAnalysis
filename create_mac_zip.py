import os
import zipfile
import stat

def create_mac_zip():
    zip_name = "Tata_Compliance_Mac_Version.zip"
    
    # Create the Mac .app structure in memory/zip directly
    app_name = "Tata Generator.app"
    
    plist_content = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>launcher</string>
    <key>CFBundleIdentifier</key>
    <string>com.tatasteel.generator</string>
    <key>CFBundleName</key>
    <string>Tata Generator</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
</dict>
</plist>"""

    launcher_content = """#!/bin/bash
TARGET_DIR="$(cd "$(dirname "$0")/../../.." && pwd)"
osascript <<EOD
tell application "Terminal"
    activate
    do script "cd '"$TARGET_DIR"' && bash Run_on_Mac.command"
end tell
EOD
"""

    with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zf:
        def add_file(filepath, arcname, is_executable=False):
            if not os.path.exists(filepath): return
            zinfo = zipfile.ZipInfo.from_file(filepath, arcname)
            # Force unix permissions
            perms = 0o755 if is_executable else 0o644
            zinfo.external_attr = (stat.S_IFREG | perms) << 16
            with open(filepath, 'rb') as f:
                zf.writestr(zinfo, f.read())
                
        def add_dir(arcname):
            zinfo = zipfile.ZipInfo(arcname + '/')
            zinfo.external_attr = (stat.S_IFDIR | 0o755) << 16
            zf.writestr(zinfo, '')

        # Add the Mac .app wrapper
        add_dir(app_name)
        add_dir(f"{app_name}/Contents")
        add_dir(f"{app_name}/Contents/MacOS")
        
        # Add plist
        zinfo = zipfile.ZipInfo(f"{app_name}/Contents/Info.plist")
        zinfo.external_attr = (stat.S_IFREG | 0o644) << 16
        zf.writestr(zinfo, plist_content)
        
        # Add executable launcher
        zinfo = zipfile.ZipInfo(f"{app_name}/Contents/MacOS/launcher")
        zinfo.external_attr = (stat.S_IFREG | 0o755) << 16
        zf.writestr(zinfo, launcher_content)

        # Add project files
        folders_to_include = ['Codes', 'monthly_commitments', 'Production Data', 'Cleaned_Data', 'results']
        files_to_include = ['Run_on_Mac.command', 'requirements.txt']

        for file in files_to_include:
            add_file(file, file, is_executable=(file.endswith('.command')))

        for folder in folders_to_include:
            if not os.path.exists(folder): continue
            add_dir(folder)
            for root, dirs, files in os.walk(folder):
                if '__pycache__' in root: continue
                for d in dirs:
                    if '__pycache__' in d: continue
                    arc_dir = os.path.relpath(os.path.join(root, d), '.')
                    add_dir(arc_dir)
                for f in files:
                    if f.endswith('.pyc'): continue
                    filepath = os.path.join(root, f)
                    arcname = os.path.relpath(filepath, '.')
                    add_file(filepath, arcname)

    print(f"Successfully created {zip_name}")

if __name__ == '__main__':
    create_mac_zip()

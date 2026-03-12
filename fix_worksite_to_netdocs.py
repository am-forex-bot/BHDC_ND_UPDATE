#!/usr/bin/env python3
"""
Fix Word templates (.dotx/.zip) to use NetDocuments instead of WorkSite/FileSite.

This script patches the BigHand/Iphelion Outline DMS integration in
customXml/item2.xml inside Word template zip files.

Changes made:
  1. DMS question: WorkSite.dll -> NetDocuments.dll, SelectWorkSpaceViewModel -> SelectWorkspaceViewModel
  2. Save command: "Save to WorkSite" -> "Save to NetDocuments", WorkSite.SaveToDmsCommand -> NetDocuments.Commands.SaveToDmsCommand
  3. Update Author command: "Update WorkSite author" -> "Update NetDocuments author"
  4. Adds missing NetDocuments-specific Save parameters (checkForDocId, documentId, version)

Usage:
  python fix_worksite_to_netdocs.py <file_or_directory>

  If given a directory, processes all .zip files in it.
  If given a file, processes just that file.

  Creates backups with .bak suffix before modifying.
"""

import os
import re
import sys
import shutil
import zipfile
import tempfile


# The 3 extra parameters that NetDocuments Save command needs (not present in WorkSite)
EXTRA_SAVE_PARAMS = '''
        <parameter id="a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d" name="Check for Doc Id" type="System.Boolean, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" order="999" key="checkForDocId" value="True" groupOrder="-1" isGenerated="false"/>
        <parameter id="b2c3d4e5-f6a7-4b8c-9d0e-1f2a3b4c5d6e" name="Document Id" type="Iphelion.Outline.Model.Entities.ParameterFieldDescriptor, Iphelion.Outline.Model, Version=2.6.0.60, Culture=neutral, PublicKeyToken=null" order="999" key="documentId" value="b6f60e59-6c84-4d5f-8fb2-e4aaf9426131|4391e415-b6c6-4380-b620-00144a66365b" groupOrder="-1" isGenerated="false"/>
        <parameter id="c3d4e5f6-a7b8-4c9d-0e1f-2a3b4c5d6e7f" name="Version" type="Iphelion.Outline.Model.Entities.ParameterFieldDescriptor, Iphelion.Outline.Model, Version=2.6.0.60, Culture=neutral, PublicKeyToken=null" order="999" key="documentVersion" value="b6f60e59-6c84-4d5f-8fb2-e4aaf9426131|014ff85a-b1de-4839-8ed8-7d8d79cdd820" groupOrder="-1" isGenerated="false"/>'''


def fix_item2_xml(content):
    """Apply all WorkSite -> NetDocuments fixes to item2.xml content (string).
    Returns (fixed_content, changes_made_list) or (None, []) if no changes needed.
    """
    changes = []
    original = content

    # 1. DMS question: assembly + type
    old = 'assembly="Iphelion.Outline.Integration.WorkSite.dll" type="Iphelion.Outline.Integration.WorkSite.ViewModels.SelectWorkSpaceViewModel"'
    new = 'assembly="Iphelion.Outline.Integration.NetDocuments.dll" type="Iphelion.Outline.Integration.NetDocuments.ViewModels.SelectWorkspaceViewModel"'
    if old in content:
        content = content.replace(old, new)
        changes.append("DMS question: WorkSite -> NetDocuments (assembly + ViewModel)")

    # 2. Save command: name + assembly + type
    old = 'name="Save to WorkSite" assembly="Iphelion.Outline.Integration.WorkSite.dll" type="Iphelion.Outline.Integration.WorkSite.SaveToDmsCommand"'
    new = 'name="Save to NetDocuments" assembly="Iphelion.Outline.Integration.NetDocuments.dll" type="Iphelion.Outline.Integration.NetDocuments.Commands.SaveToDmsCommand"'
    if old in content:
        content = content.replace(old, new)
        changes.append("Save command: WorkSite -> NetDocuments (name + assembly + type)")

    # 3. Update Author command: name + assembly + type
    old = 'name="Update WorkSite author" assembly="Iphelion.Outline.Integration.WorkSite.dll" type="Iphelion.Outline.Integration.WorkSite.UpdateAuthorCommand"'
    new = 'name="Update NetDocuments author" assembly="Iphelion.Outline.Integration.NetDocuments.dll" type="Iphelion.Outline.Integration.NetDocuments.Commands.UpdateAuthorCommand"'
    if old in content:
        content = content.replace(old, new)
        changes.append("Update Author command: WorkSite -> NetDocuments (name + assembly + type)")

    # 4. Add extra Save command parameters if missing
    if 'checkForDocId' not in content:
        # Find the Save command's titleField parameter and insert after it
        save_match = re.search(r'(<command id="1311b0ba.*?</command>)', content, re.DOTALL)
        if save_match:
            old_save = save_match.group()
            if 'key="titleField"' in old_save:
                new_save = re.sub(
                    r'(key="titleField"[^/]*/\>)',
                    r'\1' + EXTRA_SAVE_PARAMS,
                    old_save
                )
                content = content.replace(old_save, new_save)
                changes.append("Added missing NetDocuments Save params (checkForDocId, documentId, version)")

    # Catch any remaining WorkSite references we might have missed
    remaining = re.findall(r'Iphelion\.Outline\.Integration\.WorkSite', content)
    if remaining:
        # Generic fallback: replace any remaining WorkSite assembly references
        content = content.replace(
            'Iphelion.Outline.Integration.WorkSite',
            'Iphelion.Outline.Integration.NetDocuments'
        )
        changes.append(f"Generic fix: {len(remaining)} remaining WorkSite assembly reference(s)")

    if content == original:
        return None, []
    return content, changes


def process_zip(zip_path):
    """Process a single zip file. Returns True if changes were made."""
    print(f"\nProcessing: {zip_path}")

    # Find item2.xml in the zip
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            item2_entries = [n for n in zf.namelist() if n.endswith('customXml/item2.xml')]
            if not item2_entries:
                print("  SKIP: No customXml/item2.xml found")
                return False

            item2_name = item2_entries[0]
            raw = zf.read(item2_name)
    except (zipfile.BadZipFile, Exception) as e:
        print(f"  ERROR: Cannot read zip: {e}")
        return False

    # Decode UTF-16
    try:
        if raw[:2] == b'\xff\xfe':
            content = raw[2:].decode('utf-16-le')
        elif raw[:2] == b'\xfe\xff':
            content = raw[2:].decode('utf-16-be')
        else:
            content = raw.decode('utf-16-le', errors='replace')
    except Exception as e:
        print(f"  ERROR: Cannot decode item2.xml: {e}")
        return False

    # Check if this template uses WorkSite
    if 'WorkSite' not in content:
        if 'NetDocuments' in content:
            print("  SKIP: Already using NetDocuments")
        else:
            print("  SKIP: No WorkSite or NetDocuments references found")
        return False

    # Apply fixes
    fixed_content, changes = fix_item2_xml(content)
    if not fixed_content:
        print("  SKIP: No changes needed")
        return False

    # Create backup
    backup_path = zip_path + '.bak'
    if not os.path.exists(backup_path):
        shutil.copy2(zip_path, backup_path)
        print(f"  Backup: {backup_path}")

    # Repackage zip with the fixed item2.xml
    with tempfile.NamedTemporaryFile(suffix='.zip', delete=False) as tmp:
        tmp_path = tmp.name

    try:
        with zipfile.ZipFile(zip_path, 'r') as zf_in:
            with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zf_out:
                for item in zf_in.infolist():
                    if item.filename == item2_name:
                        # Write fixed content
                        fixed_bytes = b'\xff\xfe' + fixed_content.encode('utf-16-le')
                        zf_out.writestr(item, fixed_bytes)
                    else:
                        zf_out.writestr(item, zf_in.read(item.filename))

        # Replace original with fixed version
        shutil.move(tmp_path, zip_path)
    except Exception as e:
        print(f"  ERROR: Failed to repackage: {e}")
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        return False

    # Report changes
    for change in changes:
        print(f"  FIXED: {change}")

    # Verify
    with zipfile.ZipFile(zip_path, 'r') as zf:
        data = zf.read(item2_name)
        text = data.decode('utf-16-le', errors='replace')
        ws_count = len(re.findall(r'WorkSite', text))
        nd_count = len(re.findall(r'NetDocuments', text))
        has_extra = 'checkForDocId' in text
        print(f"  Verify: WorkSite={ws_count}, NetDocuments={nd_count}, checkForDocId={has_extra}")

    return True


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    target = sys.argv[1]
    fixed_count = 0
    total_count = 0

    if os.path.isdir(target):
        # Process all .zip files in directory
        zip_files = sorted([
            os.path.join(target, f) for f in os.listdir(target)
            if f.endswith('.zip') and not f.endswith('.bak.zip')
        ])
        if not zip_files:
            print(f"No .zip files found in {target}")
            sys.exit(1)
        print(f"Found {len(zip_files)} zip file(s) in {target}")
        for zp in zip_files:
            total_count += 1
            if process_zip(zp):
                fixed_count += 1
    elif os.path.isfile(target):
        total_count = 1
        if process_zip(target):
            fixed_count = 1
    else:
        print(f"Error: {target} not found")
        sys.exit(1)

    print(f"\n{'='*50}")
    print(f"Done. Fixed {fixed_count} of {total_count} file(s).")
    if fixed_count > 0:
        print("Backups saved with .bak extension.")


if __name__ == '__main__':
    main()

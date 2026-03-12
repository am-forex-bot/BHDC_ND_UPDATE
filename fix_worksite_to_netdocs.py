#!/usr/bin/env python3
"""
Fix Word templates (.dotx) to use NetDocuments instead of WorkSite/FileSite.

This script patches the BigHand/Iphelion Outline DMS integration in
customXml/item2.xml inside Word template files.

Works DIRECTLY on .dotx files - no need to extract or re-zip manually.
(A .dotx is just a zip file with a different extension.)

Changes made:
  1. DMS question: WorkSite.dll -> NetDocuments.dll
  2. Save command: "Save to WorkSite" -> "Save to NetDocuments"
  3. Update Author command: "Update WorkSite author" -> "Update NetDocuments author"
  4. Adds missing NetDocuments-specific Save parameters (checkForDocId, documentId, version)
  5. Replaces all 16 WorkSite DMS field IDs with NetDocuments equivalents
     (Client, Matter, MatterName, DocIdFormat, DocNumber, DocVersion, etc.)
  6. Fixes docType value and Subject question active state

Usage:
  python fix_worksite_to_netdocs.py <file_or_directory>

  If given a directory, processes all .dotx files in it.
  If given a single .dotx file, processes just that file.

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

# WorkSite field ID -> NetDocuments field ID mapping
# These are the GUIDs for DMS fields exposed by each integration DLL.
# WorkSite and NetDocuments use completely different IDs for the same logical fields.
FIELD_ID_REPLACEMENTS = {
    # WorkSite ID                              NetDocuments ID                           Field Name
    "af020c1a-f826-494c-bbaa-2100b39770a7": "ab5b4913-b64e-41e6-9bee-f1c7e7c8ba41",  # Client
    "d1a0c03d-0258-47ac-bb6d-458a78e56474": "152eda4c-2831-40ac-9162-c15059d65622",  # ClientName
    "9016353d-0ab3-451f-9828-3fee96cf68ba": "6f228f23-540f-4466-ba40-14327c3f0f84",  # Connected
    "2403d342-533b-45e7-84b2-62d681290485": "c0adb0a3-a528-49ea-ae35-041d6e4fff29",  # Create new version
    "d8d8a1b7-29f2-4184-b4bb-94e86811b1dc": "c209e536-a107-40f1-bfc4-80610f2ec53b",  # DocFolderId
    "72904a47-5780-459c-be7a-448f9ad8d6b4": "cae65a0c-9dba-4e06-868e-6fc5fd082911",  # DocIdFormat
    "a1f231ea-a00f-4606-9fab-d2acd859d3ad": "4391e415-b6c6-4380-b620-00144a66365b",  # DocNumber
    "7abea0f8-46b7-4968-bb12-04a899f0d778": "f3821f3f-25f3-4793-b379-59f1f861a1b1",  # DocSubType
    "64ff0036-a6af-4b11-a4ea-402a2f273e21": "891ae94e-e0d3-4126-acdb-04d9e07b21ec",  # DocType
    "c9094b9c-52fd-4403-bb83-9bb3ab5368ad": "014ff85a-b1de-4839-8ed8-7d8d79cdd820",  # DocVersion
    "2fef3f19-232d-4142-b525-11d8a76a6e9b": "b4605801-b53b-4334-8022-11fa556809a5",  # Library
    "362ddceb-8fc2-4ead-b535-ed9e83598384": "ee374f24-04a1-482e-8118-b26f8f351cc2",  # Matter
    "a3eef514-247f-4281-b6a2-3b4d34bc68cf": "6a059c3f-9fdb-485a-8240-638fea464980",  # MatterName
    "01a5919e-9f80-47f4-93c4-a97878088c9c": "978f91f0-df54-498f-9e5c-abf7df96c941",  # Server
    "a002e78a-8e18-4375-bef7-9f687e931f65": "3743e51b-251b-402c-9f20-d842f238463a",  # Title
    "388a1e13-9978-4547-8c39-29b89a11d72a": "87d09b6b-2044-4360-b13c-604db59f9b12",  # WorkspaceId
}


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

    # 5. Replace all WorkSite DMS field IDs with NetDocuments equivalents
    field_count = 0
    for ws_id, nd_id in FIELD_ID_REPLACEMENTS.items():
        count = content.count(ws_id)
        if count > 0:
            content = content.replace(ws_id, nd_id)
            field_count += count
    if field_count > 0:
        changes.append(f"Replaced {field_count} WorkSite DMS field ID reference(s) with NetDocuments equivalents")

    # 6. Fix docType value (WorkSite uses "Letter", NetDocuments uses "LET")
    if 'key="docType" value="Letter"' in content:
        content = content.replace('key="docType" value="Letter"', 'key="docType" value="LET"')
        changes.append("Fixed docType: 'Letter' -> 'LET'")

    # 7. Enable Subject question if disabled
    old_subj = 'id="11904e11-bb39-4293-9339-71128b7bf8e7" name="Subject" assembly="Iphelion.Outline.Controls.dll" type="Iphelion.Outline.Controls.QuestionControls.ViewModels.ReferenceViewModel" order="5" active="false"'
    new_subj = 'id="11904e11-bb39-4293-9339-71128b7bf8e7" name="Subject" assembly="Iphelion.Outline.Controls.dll" type="Iphelion.Outline.Controls.QuestionControls.ViewModels.ReferenceViewModel" order="5" active="true"'
    if old_subj in content:
        content = content.replace(old_subj, new_subj)
        changes.append("Enabled Subject question (active=false -> active=true)")

    # 8. Catch any remaining WorkSite assembly references
    remaining = re.findall(r'Iphelion\.Outline\.Integration\.WorkSite', content)
    if remaining:
        content = content.replace(
            'Iphelion.Outline.Integration.WorkSite',
            'Iphelion.Outline.Integration.NetDocuments'
        )
        changes.append(f"Generic fix: {len(remaining)} remaining WorkSite assembly reference(s)")

    if content == original:
        return None, []
    return content, changes


def process_file(file_path):
    """Process a single .dotx (or .zip) file. Returns True if changes were made."""
    print(f"\nProcessing: {os.path.basename(file_path)}")

    # Find item2.xml in the zip
    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            item2_entries = [n for n in zf.namelist() if 'customXml/item2.xml' in n]
            if not item2_entries:
                print("  SKIP: No customXml/item2.xml found")
                return False

            item2_name = item2_entries[0]
            raw = zf.read(item2_name)
    except (zipfile.BadZipFile, Exception) as e:
        print(f"  ERROR: Cannot read file: {e}")
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

    # Check if this template needs fixing (has WorkSite refs OR has WorkSite field IDs)
    has_worksite = 'WorkSite' in content
    has_ws_field_ids = any(ws_id in content for ws_id in FIELD_ID_REPLACEMENTS)

    if not has_worksite and not has_ws_field_ids:
        if 'NetDocuments' in content:
            print("  SKIP: Already fully converted to NetDocuments")
        else:
            print("  SKIP: No WorkSite or NetDocuments references found")
        return False

    # Apply fixes
    fixed_content, changes = fix_item2_xml(content)
    if not fixed_content:
        print("  SKIP: No changes needed")
        return False

    # Create backup
    backup_path = file_path + '.bak'
    if not os.path.exists(backup_path):
        shutil.copy2(file_path, backup_path)
        print(f"  Backup: {os.path.basename(backup_path)}")

    # Repackage: read original zip, replace just item2.xml, write back
    with tempfile.NamedTemporaryFile(suffix=os.path.splitext(file_path)[1], delete=False) as tmp:
        tmp_path = tmp.name

    try:
        with zipfile.ZipFile(file_path, 'r') as zf_in:
            with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zf_out:
                for item in zf_in.infolist():
                    if item.filename == item2_name:
                        fixed_bytes = b'\xff\xfe' + fixed_content.encode('utf-16-le')
                        zf_out.writestr(item, fixed_bytes)
                    else:
                        zf_out.writestr(item, zf_in.read(item.filename))

        shutil.move(tmp_path, file_path)
    except Exception as e:
        print(f"  ERROR: Failed to repackage: {e}")
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        return False

    # Report changes
    for change in changes:
        print(f"  FIXED: {change}")

    # Verify
    with zipfile.ZipFile(file_path, 'r') as zf:
        data = zf.read(item2_name)
        text = data.decode('utf-16-le', errors='replace')
        ws_count = len(re.findall(r'WorkSite', text))
        nd_count = len(re.findall(r'NetDocuments', text))
        ws_fields = sum(1 for ws_id in FIELD_ID_REPLACEMENTS if ws_id in text)
        print(f"  Verify: WorkSite={ws_count}, NetDocuments={nd_count}, "
              f"remaining_WS_fields={ws_fields}, checkForDocId={'checkForDocId' in text}")

    return True


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    target = sys.argv[1]
    fixed_count = 0
    total_count = 0

    if os.path.isdir(target):
        # Process all .dotx files in directory
        files = sorted([
            os.path.join(target, f) for f in os.listdir(target)
            if f.endswith('.dotx')
            and not f.endswith('.bak.dotx')
            and not f.endswith('.bak')
        ])
        if not files:
            print(f"No .dotx files found in {target}")
            sys.exit(1)
        print(f"Found {len(files)} .dotx file(s) in {target}")
        for fp in files:
            total_count += 1
            if process_file(fp):
                fixed_count += 1
    elif os.path.isfile(target):
        total_count = 1
        if process_file(target):
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

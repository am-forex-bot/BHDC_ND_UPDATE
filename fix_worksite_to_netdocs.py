#!/usr/bin/env python3
"""
Fix Word templates (.dotx) to use NetDocuments instead of WorkSite/FileSite.

This script patches the BigHand/Iphelion Outline DMS integration in
customXml/item*.xml inside Word template files.

Works DIRECTLY on .dotx files - no need to extract or re-zip manually.
(A .dotx is just a zip file with a different extension.)

Changes made:
  1. DMS question: WorkSite.dll -> NetDocuments.dll
  2. Save command: "Save to WorkSite" -> "Save to NetDocuments"
  3. Update Author command: "Update WorkSite author" -> "Update NetDocuments author"
  4. Adds missing NetDocuments-specific Save parameters (checkForDocId, documentId, version)
  5. Auto-detects and replaces ALL wrong DMS field IDs by comparing field names
     against the known-good NetDocuments reference IDs
  6. Fixes docType value and Subject question active state

Note: The BigHand config can be in any customXml item (item1.xml, item2.xml, etc.)
      depending on the template. The script checks all of them.

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

# The correct NetDocuments field IDs, keyed by field name.
# These come from the known-working Letter.dotx template.
# Any field in a template with the same name but a different ID is a WorkSite
# leftover that needs replacing.
NETDOCS_FIELD_IDS = {
    "Client":             "ab5b4913-b64e-41e6-9bee-f1c7e7c8ba41",
    "ClientName":         "152eda4c-2831-40ac-9162-c15059d65622",
    "Connected":          "6f228f23-540f-4466-ba40-14327c3f0f84",
    "Create new version": "c0adb0a3-a528-49ea-ae35-041d6e4fff29",
    "DocFolderId":        "c209e536-a107-40f1-bfc4-80610f2ec53b",
    "DocIdFormat":        "cae65a0c-9dba-4e06-868e-6fc5fd082911",
    "DocNumber":          "4391e415-b6c6-4380-b620-00144a66365b",
    "DocSubType":         "f3821f3f-25f3-4793-b379-59f1f861a1b1",
    "DocType":            "891ae94e-e0d3-4126-acdb-04d9e07b21ec",
    "DocVersion":         "014ff85a-b1de-4839-8ed8-7d8d79cdd820",
    "Library":            "b4605801-b53b-4334-8022-11fa556809a5",
    "Matter":             "ee374f24-04a1-482e-8118-b26f8f351cc2",
    "MatterName":         "6a059c3f-9fdb-485a-8240-638fea464980",
    "Server":             "978f91f0-df54-498f-9e5c-abf7df96c941",
    "Title":              "3743e51b-251b-402c-9f20-d842f238463a",
    "WorkspaceId":        "87d09b6b-2044-4360-b13c-604db59f9b12",
}


def find_dms_question_id(content):
    """Find the DMS question's ID dynamically. Each template can have a different one.
    The DMS question is identified by its assembly referencing either WorkSite or NetDocuments.
    """
    m = re.search(
        r'<question\s[^>]*id="([^"]*)"[^>]*assembly="Iphelion\.Outline\.Integration\.'
        r'(?:NetDocuments|WorkSite)\.dll"',
        content
    )
    if m:
        return m.group(1)
    # Fallback: look for name="DMS"
    m = re.search(r'<question\s[^>]*id="([^"]*)"[^>]*name="DMS"', content)
    return m.group(1) if m else None


def auto_detect_field_replacements(content):
    """Scan the template for DMS-related <field> elements and build a replacement
    map for any that don't match the known-good NetDocuments IDs.

    Dynamically finds the DMS question ID (varies per template) then checks
    all fields belonging to it.

    Returns a dict of {wrong_id: correct_netdocs_id}.
    """
    dms_id = find_dms_question_id(content)
    if not dms_id:
        return {}

    replacements = {}

    # Find all <field> elements that belong to the DMS question (any attribute order)
    for m in re.finditer(r'<field\s+id="([^"]*)"[^>]*/?\s*>', content):
        tag = m.group()
        # Check this field belongs to the DMS question
        eid_match = re.search(r'entityId="([^"]*)"', tag)
        if not eid_match or eid_match.group(1) != dms_id:
            continue

        field_id = re.search(r'id="([^"]*)"', tag).group(1)
        name_match = re.search(r'name="([^"]*)"', tag)
        if not name_match:
            continue

        field_name = name_match.group(1)
        correct_id = NETDOCS_FIELD_IDS.get(field_name)
        if correct_id and field_id != correct_id:
            replacements[field_id] = correct_id

    return replacements


def fix_item_xml(content, new_doctype=None):
    """Apply all WorkSite -> NetDocuments fixes to a BigHand item XML content string.
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

    # 5. Auto-detect and replace wrong DMS field IDs
    field_replacements = auto_detect_field_replacements(content)
    if field_replacements:
        field_count = 0
        for wrong_id, correct_id in field_replacements.items():
            count = content.count(wrong_id)
            if count > 0:
                content = content.replace(wrong_id, correct_id)
                field_count += count
        if field_count > 0:
            changes.append(f"Replaced {field_count} wrong DMS field ID reference(s) "
                           f"({len(field_replacements)} unique fields) with NetDocuments equivalents")

    # 6. Fix docType value if a new value is specified
    if new_doctype:
        m = re.search(r'key="docType"\s+value="([^"]*)"', content)
        if m and m.group(1) != new_doctype:
            content = re.sub(
                r'(key="docType"\s+value=")[^"]*(")',
                rf'\g<1>{new_doctype}\2',
                content
            )
            changes.append(f"docType: '{m.group(1)}' -> '{new_doctype}'")

    # 7. Enable Subject question if disabled
    old_subj = 'id="11904e11-bb39-4293-9339-71128b7bf8e7" name="Subject" assembly="Iphelion.Outline.Controls.dll" type="Iphelion.Outline.Controls.QuestionControls.ViewModels.ReferenceViewModel" order="5" active="false"'
    new_subj = old_subj.replace('active="false"', 'active="true"')
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


def find_bighand_entry(zf):
    """Find the customXml item that contains the BigHand/Outline config.
    Returns (entry_name, raw_bytes) or (None, None).
    """
    candidates = sorted([n for n in zf.namelist() if re.match(r'customXml/item\d+\.xml$', n)])
    for entry in candidates:
        raw = zf.read(entry)
        if raw[:2] in (b'\xff\xfe', b'\xfe\xff'):
            try:
                text = raw[2:].decode('utf-16-le' if raw[:2] == b'\xff\xfe' else 'utf-16-be')
            except Exception:
                continue
            if '<template' in text[:200]:
                return entry, raw
        elif b'<template' in raw[:400]:
            return entry, raw
    return None, None


def decode_content(raw):
    """Decode raw bytes from a BigHand item XML to string."""
    if raw[:2] == b'\xff\xfe':
        return raw[2:].decode('utf-16-le')
    elif raw[:2] == b'\xfe\xff':
        return raw[2:].decode('utf-16-be')
    else:
        return raw.decode('utf-16-le', errors='replace')


def process_file(file_path):
    """Process a single .dotx (or .zip) file. Returns True if changes were made."""
    print(f"\nProcessing: {os.path.basename(file_path)}")

    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            entry_name, raw = find_bighand_entry(zf)
            if entry_name is None:
                print("  SKIP: No BigHand/Outline config found in any customXml item")
                return False
            print(f"  Config found in: {entry_name}")
    except (zipfile.BadZipFile, Exception) as e:
        print(f"  ERROR: Cannot read file: {e}")
        return False

    content = decode_content(raw)

    # Check if this template needs fixing
    has_worksite = 'WorkSite' in content
    has_wrong_fields = bool(auto_detect_field_replacements(content))

    if not has_worksite and not has_wrong_fields:
        if 'NetDocuments' in content:
            print("  SKIP: Already fully converted to NetDocuments")
        else:
            print("  SKIP: No WorkSite or wrong field IDs found")
        return False

    if has_wrong_fields and not has_worksite:
        print("  NOTE: DLLs already NetDocuments but field IDs still need fixing")

    # Apply fixes
    fixed_content, changes = fix_item_xml(content)
    if not fixed_content:
        print("  SKIP: No changes needed")
        return False

    # Create backup
    backup_path = file_path + '.bak'
    if not os.path.exists(backup_path):
        shutil.copy2(file_path, backup_path)
        print(f"  Backup: {os.path.basename(backup_path)}")

    # Repackage
    with tempfile.NamedTemporaryFile(suffix=os.path.splitext(file_path)[1], delete=False) as tmp:
        tmp_path = tmp.name

    try:
        with zipfile.ZipFile(file_path, 'r') as zf_in:
            with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zf_out:
                for item in zf_in.infolist():
                    if item.filename == entry_name:
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

    for change in changes:
        print(f"  FIXED: {change}")

    # Verify
    with zipfile.ZipFile(file_path, 'r') as zf:
        data = zf.read(entry_name)
        text = data.decode('utf-16-le', errors='replace')
        ws_count = len(re.findall(r'WorkSite', text))
        wrong_fields = len(auto_detect_field_replacements(text))
        nd_count = len(re.findall(r'NetDocuments', text))
        print(f"  Verify: WorkSite={ws_count}, wrong_fields={wrong_fields}, "
              f"NetDocuments={nd_count}, checkForDocId={'checkForDocId' in text}")

    return True


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    target = sys.argv[1]
    fixed_count = 0
    total_count = 0

    if os.path.isdir(target):
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

#!/usr/bin/env python3
"""
GUI tool to fix Word templates (.dotx) from WorkSite to NetDocuments.

Opens a tkinter window where you can:
  - Browse for a single .dotx file or a folder of them
  - See each file that needs fixing with its current docType
  - Choose the correct NetDocuments docType for each (LET, INV, DOC, etc.)
  - Apply all fixes with one click

Auto-detects wrong DMS field IDs by comparing field names against the
known-good NetDocuments reference. Works regardless of which customXml
item contains the BigHand config.

Creates backups with .bak suffix before modifying.
"""

import os
import re
import shutil
import zipfile
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# --- Known docType mappings (WorkSite value -> suggested NetDocuments code) ---
DOCTYPE_SUGGESTIONS = {
    "Letter":   "LET",
    "Invoice":  "INV",
    "Document": "DOC",
}

# Common NetDocuments docType codes the user can pick from
DOCTYPE_OPTIONS = ["LET", "INV", "DOC", "MEM", "FAX", "EML", "RPT", "AGR", ""]

# --- DMS fix constants ---

EXTRA_SAVE_PARAMS = '''
        <parameter id="a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d" name="Check for Doc Id" type="System.Boolean, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" order="999" key="checkForDocId" value="True" groupOrder="-1" isGenerated="false"/>
        <parameter id="b2c3d4e5-f6a7-4b8c-9d0e-1f2a3b4c5d6e" name="Document Id" type="Iphelion.Outline.Model.Entities.ParameterFieldDescriptor, Iphelion.Outline.Model, Version=2.6.0.60, Culture=neutral, PublicKeyToken=null" order="999" key="documentId" value="b6f60e59-6c84-4d5f-8fb2-e4aaf9426131|4391e415-b6c6-4380-b620-00144a66365b" groupOrder="-1" isGenerated="false"/>
        <parameter id="c3d4e5f6-a7b8-4c9d-0e1f-2a3b4c5d6e7f" name="Version" type="Iphelion.Outline.Model.Entities.ParameterFieldDescriptor, Iphelion.Outline.Model, Version=2.6.0.60, Culture=neutral, PublicKeyToken=null" order="999" key="documentVersion" value="b6f60e59-6c84-4d5f-8fb2-e4aaf9426131|014ff85a-b1de-4839-8ed8-7d8d79cdd820" groupOrder="-1" isGenerated="false"/>'''

# The DMS question ID (same across all Wallace templates)
DMS_QUESTION_ID = "b6f60e59-6c84-4d5f-8fb2-e4aaf9426131"

# The correct NetDocuments field IDs, keyed by field name.
# These come from the known-working Letter.dotx template.
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


# ---------------------------------------------------------------------------
# Core fix logic
# ---------------------------------------------------------------------------

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


def read_bighand_config(file_path):
    """Read and decode BigHand config from a .dotx.
    Returns (content_str, entry_name) or (None, None).
    """
    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            entry_name, raw = find_bighand_entry(zf)
            if entry_name is None:
                return None, None
            return decode_content(raw), entry_name
    except Exception:
        return None, None


def get_current_doctype(content):
    """Extract the current docType value from BigHand config content."""
    m = re.search(r'key="docType"\s+value="([^"]*)"', content)
    return m.group(1) if m else ""


def auto_detect_field_replacements(content):
    """Scan the template for DMS-related <field> elements and build a replacement
    map for any that don't match the known-good NetDocuments IDs.
    """
    replacements = {}
    # Match fields belonging to the DMS question (either attribute order)
    pattern = r'<field\s+id="([^"]*)"[^>]*name="([^"]*)"[^>]*entityId="' + re.escape(DMS_QUESTION_ID) + r'"'
    for m in re.finditer(pattern, content):
        field_id, field_name = m.group(1), m.group(2)
        correct_id = NETDOCS_FIELD_IDS.get(field_name)
        if correct_id and field_id != correct_id:
            replacements[field_id] = correct_id

    pattern2 = r'<field\s+id="([^"]*)"[^>]*entityId="' + re.escape(DMS_QUESTION_ID) + r'"[^>]*name="([^"]*)"'
    for m in re.finditer(pattern2, content):
        field_id, field_name = m.group(1), m.group(2)
        correct_id = NETDOCS_FIELD_IDS.get(field_name)
        if correct_id and field_id != correct_id:
            replacements[field_id] = correct_id

    return replacements


def needs_fixing(content):
    """Check if this template still has WorkSite refs or wrong field IDs."""
    if 'WorkSite' in content:
        return True
    return bool(auto_detect_field_replacements(content))


def fix_item_xml(content, new_doctype):
    """Apply all WorkSite -> NetDocuments fixes, using the given docType value."""
    changes = []
    original = content

    # DMS question assembly + type
    old = 'assembly="Iphelion.Outline.Integration.WorkSite.dll" type="Iphelion.Outline.Integration.WorkSite.ViewModels.SelectWorkSpaceViewModel"'
    new = 'assembly="Iphelion.Outline.Integration.NetDocuments.dll" type="Iphelion.Outline.Integration.NetDocuments.ViewModels.SelectWorkspaceViewModel"'
    if old in content:
        content = content.replace(old, new)
        changes.append("DMS question: WorkSite -> NetDocuments DLL")

    # Save command
    old = 'name="Save to WorkSite" assembly="Iphelion.Outline.Integration.WorkSite.dll" type="Iphelion.Outline.Integration.WorkSite.SaveToDmsCommand"'
    new = 'name="Save to NetDocuments" assembly="Iphelion.Outline.Integration.NetDocuments.dll" type="Iphelion.Outline.Integration.NetDocuments.Commands.SaveToDmsCommand"'
    if old in content:
        content = content.replace(old, new)
        changes.append("Save command: WorkSite -> NetDocuments")

    # Update Author command
    old = 'name="Update WorkSite author" assembly="Iphelion.Outline.Integration.WorkSite.dll" type="Iphelion.Outline.Integration.WorkSite.UpdateAuthorCommand"'
    new = 'name="Update NetDocuments author" assembly="Iphelion.Outline.Integration.NetDocuments.dll" type="Iphelion.Outline.Integration.NetDocuments.Commands.UpdateAuthorCommand"'
    if old in content:
        content = content.replace(old, new)
        changes.append("Update Author: WorkSite -> NetDocuments")

    # Extra Save params
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
                changes.append("Added NetDocuments Save params")

    # Auto-detect and replace wrong field IDs
    field_replacements = auto_detect_field_replacements(content)
    if field_replacements:
        field_count = 0
        for wrong_id, correct_id in field_replacements.items():
            count = content.count(wrong_id)
            if count > 0:
                content = content.replace(wrong_id, correct_id)
                field_count += count
        if field_count > 0:
            changes.append(f"Replaced {field_count} field ID ref(s) ({len(field_replacements)} fields)")

    # docType
    if new_doctype:
        current = get_current_doctype(content)
        if current != new_doctype:
            content = re.sub(
                r'(key="docType"\s+value=")[^"]*(")',
                rf'\g<1>{new_doctype}\2',
                content
            )
            changes.append(f"docType: '{current}' -> '{new_doctype}'")

    # Enable Subject question if disabled
    old_subj = 'id="11904e11-bb39-4293-9339-71128b7bf8e7" name="Subject" assembly="Iphelion.Outline.Controls.dll" type="Iphelion.Outline.Controls.QuestionControls.ViewModels.ReferenceViewModel" order="5" active="false"'
    new_subj = old_subj.replace('active="false"', 'active="true"')
    if old_subj in content:
        content = content.replace(old_subj, new_subj)
        changes.append("Enabled Subject question")

    # Catch remaining WorkSite refs
    remaining = re.findall(r'Iphelion\.Outline\.Integration\.WorkSite', content)
    if remaining:
        content = content.replace(
            'Iphelion.Outline.Integration.WorkSite',
            'Iphelion.Outline.Integration.NetDocuments'
        )
        changes.append(f"Fixed {len(remaining)} remaining WorkSite ref(s)")

    if content == original:
        return None, []
    return content, changes


def apply_fix(file_path, new_doctype):
    """Fix a single .dotx file. Returns (success, message)."""
    content, entry_name = read_bighand_config(file_path)
    if content is None:
        return False, "No BigHand config found in any customXml item"

    fixed, changes = fix_item_xml(content, new_doctype)
    if fixed is None:
        return False, "No changes needed"

    # Backup
    backup = file_path + '.bak'
    if not os.path.exists(backup):
        shutil.copy2(file_path, backup)

    # Repackage
    with tempfile.NamedTemporaryFile(suffix='.dotx', delete=False) as tmp:
        tmp_path = tmp.name

    try:
        with zipfile.ZipFile(file_path, 'r') as zf_in:
            with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zf_out:
                for item in zf_in.infolist():
                    if item.filename == entry_name:
                        zf_out.writestr(item, b'\xff\xfe' + fixed.encode('utf-16-le'))
                    else:
                        zf_out.writestr(item, zf_in.read(item.filename))
        shutil.move(tmp_path, file_path)
    except Exception as e:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        return False, f"Repackage failed: {e}"

    return True, "; ".join(changes)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class FileEntry:
    """Represents one .dotx file found."""
    def __init__(self, path, current_doctype, needs_fix):
        self.path = path
        self.filename = os.path.basename(path)
        self.current_doctype = current_doctype
        self.needs_fix = needs_fix
        self.suggested = DOCTYPE_SUGGESTIONS.get(current_doctype, current_doctype)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("WorkSite \u2192 NetDocuments Template Fixer")
        self.geometry("820x520")
        self.minsize(700, 400)

        self.file_entries = []
        self.doctype_vars = {}

        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill=tk.X)

        ttk.Button(top, text="Open File(s)...", command=self._browse_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(top, text="Open Folder...", command=self._browse_folder).pack(side=tk.LEFT, padx=(0, 5))

        self.path_label = ttk.Label(top, text="No files loaded", foreground="grey")
        self.path_label.pack(side=tk.LEFT, padx=10)

        mid = ttk.Frame(self, padding=(10, 0, 10, 0))
        mid.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(mid)
        header.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(header, text="File", font=("", 9, "bold"), width=40, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Label(header, text="Status", font=("", 9, "bold"), width=18, anchor=tk.W).pack(side=tk.LEFT, padx=5)
        ttk.Label(header, text="Current", font=("", 9, "bold"), width=10, anchor=tk.W).pack(side=tk.LEFT, padx=5)
        ttk.Label(header, text="New docType", font=("", 9, "bold"), width=12, anchor=tk.W).pack(side=tk.LEFT, padx=5)

        canvas_frame = ttk.Frame(mid)
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(canvas_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scroll_frame = ttk.Frame(self.canvas)

        self.scroll_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scroll_frame, anchor=tk.NW)
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        bot = ttk.Frame(self, padding=10)
        bot.pack(fill=tk.X)

        self.apply_btn = ttk.Button(bot, text="Apply Fixes", command=self._apply_all, state=tk.DISABLED)
        self.apply_btn.pack(side=tk.LEFT)

        self.status_label = ttk.Label(bot, text="", foreground="grey")
        self.status_label.pack(side=tk.LEFT, padx=15)

    def _browse_files(self):
        paths = filedialog.askopenfilenames(
            title="Select .dotx template(s)",
            filetypes=[("Word Templates", "*.dotx"), ("All files", "*.*")]
        )
        if paths:
            self._load_files(list(paths))

    def _browse_folder(self):
        folder = filedialog.askdirectory(title="Select folder containing .dotx templates")
        if folder:
            paths = sorted([
                os.path.join(folder, f) for f in os.listdir(folder)
                if f.lower().endswith('.dotx') and not f.endswith('.bak')
            ])
            if not paths:
                messagebox.showinfo("No files", f"No .dotx files found in:\n{folder}")
                return
            self._load_files(paths)

    def _load_files(self, paths):
        for w in self.scroll_frame.winfo_children():
            w.destroy()
        self.file_entries.clear()
        self.doctype_vars.clear()

        fixable = 0
        for path in paths:
            content, _ = read_bighand_config(path)
            if content is None:
                continue

            current_dt = get_current_doctype(content)
            fix_needed = needs_fixing(content)
            entry = FileEntry(path, current_dt, fix_needed)
            self.file_entries.append(entry)

            row = ttk.Frame(self.scroll_frame)
            row.pack(fill=tk.X, pady=1)

            ttk.Label(row, text=entry.filename, width=40, anchor=tk.W).pack(side=tk.LEFT)

            if fix_needed:
                ttk.Label(row, text="Needs fixing", foreground="red", width=18, anchor=tk.W).pack(side=tk.LEFT, padx=5)
                fixable += 1
            else:
                ttk.Label(row, text="Already converted", foreground="green", width=18, anchor=tk.W).pack(side=tk.LEFT, padx=5)

            ttk.Label(row, text=current_dt or "(empty)", width=10, anchor=tk.W).pack(side=tk.LEFT, padx=5)

            var = tk.StringVar(value=entry.suggested)
            self.doctype_vars[path] = var
            combo = ttk.Combobox(row, textvariable=var, values=DOCTYPE_OPTIONS, width=8, state="normal" if fix_needed else "disabled")
            combo.pack(side=tk.LEFT, padx=5)

        self.path_label.config(text=f"{len(self.file_entries)} file(s) loaded, {fixable} need fixing")
        self.apply_btn.config(state=tk.NORMAL if fixable > 0 else tk.DISABLED)
        self.status_label.config(text="")

    def _apply_all(self):
        fixed = 0
        errors = []

        for entry in self.file_entries:
            if not entry.needs_fix:
                continue

            new_dt = self.doctype_vars[entry.path].get().strip()
            ok, msg = apply_fix(entry.path, new_dt)
            if ok:
                fixed += 1
            else:
                errors.append(f"{entry.filename}: {msg}")

        paths = [e.path for e in self.file_entries]
        self._load_files(paths)

        if errors:
            self.status_label.config(text=f"Fixed {fixed}, {len(errors)} error(s)", foreground="orange")
            messagebox.showwarning("Errors", "\n".join(errors))
        else:
            self.status_label.config(text=f"Done! Fixed {fixed} file(s). Backups saved as .bak", foreground="green")


if __name__ == '__main__':
    App().mainloop()

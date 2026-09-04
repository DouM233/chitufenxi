from __future__ import annotations

import copy
import posixpath
import re
import shutil
import zipfile
from pathlib import Path
from tempfile import NamedTemporaryFile
from xml.etree import ElementTree as ET


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)


def _tag(name: str) -> str:
    return f"{{{MAIN_NS}}}{name}"


def _cell_ref(cell: ET.Element) -> str:
    return cell.attrib.get("r", "")


def _ref_key(ref: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]+)([0-9]+)", ref or "")
    if not match:
        return (0, 0)
    col = 0
    for char in match.group(1):
        col = col * 26 + ord(char) - 64
    return (int(match.group(2)), col)


def _shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    values: list[str] = []
    for item in root.findall(_tag("si")):
        values.append("".join(node.text or "" for node in item.iter(_tag("t"))))
    return values


def _cell_value(cell: ET.Element, shared_strings: list[str]) -> tuple[str | None, str | None, str | None]:
    """Return (type, formula, value) while keeping Excel's scalar types."""
    cell_type = cell.attrib.get("t")
    formula = cell.find(_tag("f"))
    formula_text = formula.text if formula is not None else None
    value_node = cell.find(_tag("v"))
    inline = cell.find(_tag("is"))

    if cell_type == "s" and value_node is not None:
        try:
            value = shared_strings[int(value_node.text or "0")]
        except (ValueError, IndexError):
            value = ""
        return "inlineStr", formula_text, value
    if cell_type == "inlineStr" and inline is not None:
        return "inlineStr", formula_text, "".join(node.text or "" for node in inline.iter(_tag("t")))
    if cell_type == "b" and value_node is not None:
        return "b", formula_text, value_node.text or "0"
    if cell_type == "e" and value_node is not None:
        return "e", formula_text, value_node.text or "#VALUE!"
    if value_node is not None:
        return None, formula_text, value_node.text or ""
    return None, formula_text, None


def _replace_cell_content(target: ET.Element, source: ET.Element, shared_strings: list[str]) -> None:
    cell_type, formula, value = _cell_value(source, shared_strings)
    for child in list(target):
        if child.tag in {_tag("f"), _tag("v"), _tag("is")}:
            target.remove(child)

    if cell_type is None:
        target.attrib.pop("t", None)
    else:
        target.set("t", cell_type)

    if formula is not None:
        formula_node = ET.Element(_tag("f"))
        formula_node.text = formula
        target.insert(0, formula_node)

    if value is None:
        return
    if cell_type == "inlineStr":
        inline = ET.Element(_tag("is"))
        text_node = ET.SubElement(inline, _tag("t"))
        text_node.text = value
        target.append(inline)
    else:
        value_node = ET.SubElement(target, _tag("v"))
        value_node.text = value


def _clear_cell_content(cell: ET.Element) -> None:
    for child in list(cell):
        if child.tag in {_tag("f"), _tag("v"), _tag("is")}:
            cell.remove(child)
    cell.attrib.pop("t", None)


def _copy_row_format(target_row: ET.Element, row_number: int) -> ET.Element:
    row = copy.deepcopy(target_row)
    row.attrib["r"] = str(row_number)
    for cell in row.findall(_tag("c")):
        ref = _cell_ref(cell)
        match = re.fullmatch(r"([A-Z]+)[0-9]+", ref)
        if match:
            cell.set("r", f"{match.group(1)}{row_number}")
        for child in list(cell):
            if child.tag in {_tag("f"), _tag("v"), _tag("is")}:
                cell.remove(child)
        cell.attrib.pop("t", None)
    return row


def _patch_sheet(template_xml: bytes, staging_xml: bytes, staging_shared: list[str]) -> bytes:
    target_root = ET.fromstring(template_xml)
    source_root = ET.fromstring(staging_xml)
    target_data = target_root.find(_tag("sheetData"))
    source_data = source_root.find(_tag("sheetData"))
    if target_data is None or source_data is None:
        return template_xml

    target_rows = {
        int(row.attrib["r"]): row
        for row in target_data.findall(_tag("row"))
        if row.attrib.get("r", "").isdigit()
    }
    source_rows = {
        int(row.attrib["r"]): row
        for row in source_data.findall(_tag("row"))
        if row.attrib.get("r", "").isdigit()
    }
    last_template_row = max(target_rows, default=1)

    # The mother workbook provides formatting only. Clear every old value first so
    # shorter rankings cannot expose labels left over from the template report.
    for target_row in target_rows.values():
        for target_cell in target_row.findall(_tag("c")):
            _clear_cell_content(target_cell)

    for row_number, source_row in source_rows.items():
        target_row = target_rows.get(row_number)
        if target_row is None:
            format_row = target_rows.get(last_template_row)
            target_row = (
                _copy_row_format(format_row, row_number)
                if format_row is not None
                else ET.Element(_tag("row"), {"r": str(row_number)})
            )
            target_data.append(target_row)
            target_rows[row_number] = target_row

        target_cells = {_cell_ref(cell): cell for cell in target_row.findall(_tag("c"))}
        for source_cell in source_row.findall(_tag("c")):
            ref = _cell_ref(source_cell)
            if not ref:
                continue
            target_cell = target_cells.get(ref)
            if target_cell is None:
                target_cell = ET.Element(_tag("c"), {"r": ref})
                target_row.append(target_cell)
                target_cells[ref] = target_cell
            _replace_cell_content(target_cell, source_cell, staging_shared)

    for row in list(target_data):
        if row.tag == _tag("row"):
            row[:] = sorted(list(row), key=lambda child: _ref_key(child.attrib.get("r", "")))
    target_data[:] = sorted(list(target_data), key=lambda row: int(row.attrib.get("r", "0")))
    dimension = target_root.find(_tag("dimension"))
    source_dimension = source_root.find(_tag("dimension"))
    if dimension is not None and source_dimension is not None:
        source_ref = source_dimension.attrib.get("ref")
        if source_ref:
            dimension.set("ref", source_ref)
    return ET.tostring(target_root, encoding="utf-8", xml_declaration=True)


def _archive_part_from_target(target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join("xl", target))


def _merge_extra_worksheets(
    mother: zipfile.ZipFile,
    staging: zipfile.ZipFile,
) -> tuple[dict[str, bytes], dict[str, bytes]]:
    """Append staging-only worksheets to the mother workbook package."""
    workbook_part = "xl/workbook.xml"
    rels_part = "xl/_rels/workbook.xml.rels"
    content_types_part = "[Content_Types].xml"
    mother_workbook = ET.fromstring(mother.read(workbook_part))
    staging_workbook = ET.fromstring(staging.read(workbook_part))
    mother_rels = ET.fromstring(mother.read(rels_part))
    staging_rels = ET.fromstring(staging.read(rels_part))
    mother_types = ET.fromstring(mother.read(content_types_part))
    staging_types = ET.fromstring(staging.read(content_types_part))

    mother_sheets = mother_workbook.find(_tag("sheets"))
    staging_sheets = staging_workbook.find(_tag("sheets"))
    if mother_sheets is None or staging_sheets is None:
        return {}, {}

    mother_names = {sheet.attrib.get("name", "") for sheet in mother_sheets}
    extra_sheets = [sheet for sheet in staging_sheets if sheet.attrib.get("name", "") not in mother_names]
    if not extra_sheets:
        return {}, {}

    relationship_tag = f"{{{PACKAGE_REL_NS}}}Relationship"
    override_tag = f"{{{CONTENT_TYPES_NS}}}Override"
    staging_rel_by_id = {
        rel.attrib.get("Id"): rel
        for rel in staging_rels.findall(relationship_tag)
    }
    used_rids = {rel.attrib.get("Id", "") for rel in mother_rels.findall(relationship_tag)}
    used_parts = set(mother.namelist())
    max_sheet_id = max((int(sheet.attrib.get("sheetId", "0")) for sheet in mother_sheets), default=0)
    additions: dict[str, bytes] = {}

    def next_rid() -> str:
        number = 1
        while f"rId{number}" in used_rids:
            number += 1
        value = f"rId{number}"
        used_rids.add(value)
        return value

    for source_sheet in extra_sheets:
        source_rid = source_sheet.attrib.get(f"{{{REL_NS}}}id")
        source_rel = staging_rel_by_id.get(source_rid)
        if source_rel is None:
            continue
        target = source_rel.attrib.get("Target", "")
        source_part = _archive_part_from_target(target)
        if not source_part or source_part not in staging.namelist() or source_part in used_parts:
            continue

        new_rid = next_rid()
        max_sheet_id += 1
        copied_sheet = copy.deepcopy(source_sheet)
        copied_sheet.set("sheetId", str(max_sheet_id))
        copied_sheet.set(f"{{{REL_NS}}}id", new_rid)
        mother_sheets.append(copied_sheet)

        copied_rel = copy.deepcopy(source_rel)
        copied_rel.set("Id", new_rid)
        mother_rels.append(copied_rel)
        additions[source_part] = staging.read(source_part)
        used_parts.add(source_part)

        source_rels_part = posixpath.join(
            posixpath.dirname(source_part),
            "_rels",
            posixpath.basename(source_part) + ".rels",
        )
        if source_rels_part in staging.namelist() and source_rels_part not in used_parts:
            additions[source_rels_part] = staging.read(source_rels_part)
            used_parts.add(source_rels_part)

        part_name = "/" + source_part
        if not any(item.attrib.get("PartName") == part_name for item in mother_types.findall(override_tag)):
            source_override = next(
                (item for item in staging_types.findall(override_tag) if item.attrib.get("PartName") == part_name),
                None,
            )
            if source_override is not None:
                mother_types.append(copy.deepcopy(source_override))

    replacements = {
        workbook_part: ET.tostring(mother_workbook, encoding="utf-8", xml_declaration=True),
        rels_part: ET.tostring(mother_rels, encoding="utf-8", xml_declaration=True),
        content_types_part: ET.tostring(mother_types, encoding="utf-8", xml_declaration=True),
    }
    return replacements, additions


def write_values_into_exact_template(template_path: str | Path, staging_path: str | Path, output_path: str | Path) -> None:
    """Copy the mother workbook and replace only worksheet cell contents.

    The template's drawing, chart, style, merge, and extension parts are copied byte-for-byte.
    """
    template_path = Path(template_path).resolve()
    staging_path = Path(staging_path).resolve()
    output_path = Path(output_path).resolve()
    if template_path.suffix.lower() != ".xlsx":
        raise ValueError(f"精确母版写入只支持 .xlsx：{template_path}")
    if not template_path.exists():
        raise FileNotFoundError(f"找不到 Excel 精确母版：{template_path}")
    if not staging_path.exists():
        raise FileNotFoundError(f"找不到临时数据工作簿：{staging_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(template_path, "r") as mother, zipfile.ZipFile(staging_path, "r") as staging:
        mother_names = set(mother.namelist())
        staging_names = set(staging.namelist())
        staging_strings = _shared_strings(staging)
        replacements, additions = _merge_extra_worksheets(mother, staging)
        sheet_names = sorted(
            name for name in mother_names
            if re.fullmatch(r"xl/worksheets/sheet[0-9]+\.xml", name)
        )
        with NamedTemporaryFile(prefix=f"{output_path.stem}_", suffix=".xlsx", dir=output_path.parent, delete=False) as temp:
            temp_path = Path(temp.name)
        try:
            with zipfile.ZipFile(temp_path, "w") as result:
                for item in mother.infolist():
                    data = replacements.get(item.filename, mother.read(item.filename))
                    if item.filename in sheet_names and item.filename in staging_names:
                        data = _patch_sheet(data, staging.read(item.filename), staging_strings)
                    result.writestr(item, data)
                for filename, data in additions.items():
                    result.writestr(filename, data)
            shutil.move(temp_path, output_path)
        finally:
            temp_path.unlink(missing_ok=True)

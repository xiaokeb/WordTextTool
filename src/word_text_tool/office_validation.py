"""Inspect Office containers before any desktop application opens them."""
import posixpath
import re
from pathlib import Path
from zipfile import ZipFile, BadZipFile

from lxml import etree

from word_text_tool.safe_output import OperationError

SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MAX_EXPANDED = 200 * 1024 * 1024
MAX_MEMBER = 32 * 1024 * 1024


def xml(data: bytes):
    if b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
        raise OperationError("文档含不受支持的 XML 实体声明")
    root = etree.fromstring(data, etree.XMLParser(resolve_entities=False, no_network=True, load_dtd=False))
    if root.getroottree().docinfo.doctype:
        raise OperationError("文档含不受支持的 XML 实体声明")
    return root


def inspect_package(path: Path, *, excel: bool = True) -> dict[str, bytes]:
    try:
        with ZipFile(path) as archive:
            entries = archive.infolist()
            if len(entries) > 10000 or sum(e.file_size for e in entries) > MAX_EXPANDED:
                raise OperationError("文档解压内容过大，已拒绝处理")
            parts = {}
            for entry in entries:
                name = entry.filename
                if name in parts or name.startswith("/") or "\\" in name or ".." in name.split("/"):
                    raise OperationError("文档包含重复或非法包路径")
                if entry.flag_bits & 1 or entry.file_size > MAX_MEMBER:
                    raise OperationError("文档含加密或过大部件")
                if entry.file_size > 1024 * 1024 and entry.file_size > max(1, entry.compress_size) * 300:
                    raise OperationError("文档压缩比例异常，已拒绝处理")
                lower = name.lower()
                if excel and any(s in lower for s in ("vbaproject", "macrosheet", "dialogsheet", "externallink", "connections", "querytable", "embeddings/", "activex/", "customui/")):
                    raise OperationError("工作簿含宏、外部数据或嵌入对象，不能安全自动处理")
                data = archive.read(entry)
                if lower.endswith((".xml", ".rels")):
                    root = xml(data)
                    if excel:
                        for rel in root.iter():
                            if not isinstance(rel.tag, str):
                                continue
                            relationship_type = rel.get("Type", "").lower()
                            if any(t in relationship_type for t in ("vbaproject", "oleobject", "activex", "externallink", "connections")):
                                raise OperationError("工作簿含不安全的活动内容关系")
                            if etree.QName(rel).localname in ("f", "definedName"):
                                formula = "".join(rel.itertext()).upper()
                                if (re.search(r"(?:^|[^A-Z0-9_])(?:_XLFN\.)?(?:WEBSERVICE|RTD|DDE|CALL|REGISTER|EXEC|SQL\.REQUEST|IMAGE|PY)\s*\(", formula)
                                        or ("|" in formula and "!" in formula)):
                                    raise OperationError("工作簿含可调用外部资源或代码的公式")
                            if rel.get("TargetMode") == "External":
                                if not rel.get("Type", "").endswith("/hyperlink"):
                                    raise OperationError("工作簿含外部资源关系")
                                target = rel.get("Target", "").lower()
                                if not target.startswith(("https://", "http://", "mailto:")):
                                    raise OperationError("工作簿含不安全的外部链接")
                            content = rel.get("ContentType", "").lower()
                            if any(t in content for t in ("macroenabled", "macrosheet", "activex", "oleobject")):
                                raise OperationError("工作簿包含不支持的活动内容")
                parts[name] = data
            required = "xl/workbook.xml" if excel else "word/document.xml"
            if required not in parts or "[Content_Types].xml" not in parts:
                raise OperationError("不是有效的 Office 文档包")
            return parts
    except OperationError:
        raise
    except (BadZipFile, etree.XMLSyntaxError, KeyError, RuntimeError, ValueError) as exc:
        message = "文档损坏、加密或 XML 无效" if excel else "文档损坏、已加密或不是有效的 DOCX 文件"
        raise OperationError(message) from exc


def workbook_sheets(parts: dict[str, bytes]):
    ns = {"s": SHEET_NS}
    workbook = xml(parts["xl/workbook.xml"])
    protection = workbook.find("s:workbookProtection", ns)
    if protection is not None and any(protection.get(key) in ("1", "true") for key in ("lockStructure", "lockWindows", "lockRevision")):
        raise OperationError("工作簿已保护")
    relations = {r.get("Id"): r for r in xml(parts["xl/_rels/workbook.xml.rels"])}
    result = []
    for sheet in workbook.findall("s:sheets/s:sheet", ns):
        name = sheet.get("name", "")
        if sheet.get("state", "visible") != "visible":
            raise OperationError(f"工作表“{name}”不可见，请先确认所有工作表可见有效")
        rel = relations[sheet.get(f"{{{REL_NS}}}id")]
        if not rel.get("Type", "").endswith("/worksheet"):
            raise OperationError("工作簿包含非普通工作表")
        target = rel.get("Target", "")
        path = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
        root = xml(parts[path])
        protection = root.find("s:sheetProtection", ns)
        if protection is not None and protection.get("sheet", "true") in ("true", "1"):
            raise OperationError(f"工作表“{name}”已保护")
        result.append((name, root))
    if not result:
        raise OperationError("工作簿没有可处理的工作表")
    return tuple(result)

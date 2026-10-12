"""Immutable numbering previews, including fixed-mode row insertion."""
from dataclasses import dataclass
from pathlib import Path

from word_text_tool.office_validation import SHEET_NS, inspect_package, workbook_sheets, xml
from word_text_tool.safe_output import Fingerprint, OperationError, fingerprint, require_unchanged


@dataclass(frozen=True)
class NumberSettings:
    template: str = "编号-yy"
    start: int = 1
    fixed: bool = False
    fixed_text: str = ""
    engine: str = "auto"

    def validate(self):
        if self.template.count("yy") != 1 or not 1 <= self.start <= 999999999:
            raise OperationError("编号模板须包含一个 yy，起始序号须为 1–999999999")
        if len(self.template) > 9990 or len(self.fixed_text) > 10000:
            raise OperationError("编号文字过长")
        if self.fixed and not self.fixed_text.strip():
            raise OperationError("固定编号不能为空或纯空白")
        if self.engine not in ("auto", "excel", "wps"):
            raise OperationError("无效的办公引擎")
        for text in (self.template, self.fixed_text):
            if any(ord(c) < 32 and c not in "\t\n\r" for c in text):
                raise OperationError("编号包含不支持的控制字符")

    def number(self, index):
        self.validate()
        return self.template.replace("yy", f"{self.start + index:02d}")


@dataclass(frozen=True)
class NumberPreview:
    source: Path
    fingerprint: Fingerprint
    number: str
    sheets: tuple[str, ...]
    overwrites: tuple[str, ...]


def preview_number(source: Path, settings: NumberSettings, index: int = 0, *, pdf_only=False) -> NumberPreview:
    source = Path(source).absolute()
    if source.suffix.lower() != ".xlsx":
        raise OperationError("编号与 PDF 导出仅支持 .xlsx")
    number = "" if pdf_only else settings.number(index)
    snapshot = fingerprint(source)
    parts = inspect_package(source)
    sheets = workbook_sheets(parts)
    ns = {"s": SHEET_NS}
    shared = []
    if "xl/sharedStrings.xml" in parts:
        shared = ["".join(item.itertext()) for item in xml(parts["xl/sharedStrings.xml"])]
    overwritten = []
    for name, sheet in sheets:
        # Fixed mode inserts row 1; E2:F2 then corresponds to original E1:F1.
        if pdf_only:
            continue
        for merge in sheet.findall("s:mergeCells/s:mergeCell", ns):
            ref = merge.get("ref", "")
            from openpyxl.utils.cell import range_boundaries
            c1, r1, c2, r2 = range_boundaries(ref)
            if r1 <= 1 <= r2 and c1 <= 6 and c2 >= 5 and ref != "E1:F1":
                raise OperationError(f"工作表“{name}”E1:F1 与其他合并区域冲突")
        if settings.fixed and sheet.find("s:sheetData/s:row[@r='1048576']", ns) is not None:
            raise OperationError(f"工作表“{name}”最后一行有内容，无法安全插入首行")
        for cell in sheet.findall("s:sheetData/s:row/s:c", ns):
            address = cell.get("r")
            if address not in ("E1", "F1"):
                continue
            if cell.find("s:f", ns) is not None:
                raise OperationError(f"工作表“{name}”{address} 含公式，不能覆盖")
            value = cell.findtext("s:v", "", ns)
            if cell.get("t") == "s" and value:
                value = shared[int(value)]
            elif cell.get("t") == "inlineStr":
                value = "".join(cell.find("s:is", ns).itertext())
            if address == "F1" and value:
                raise OperationError(f"工作表“{name}”F1 非空，不能安全合并")
            if address == "E1" and value and value != number:
                overwritten.append(f"{name}!{'E2' if settings.fixed else 'E1'}：{value} → {number}")
    require_unchanged(source, snapshot)
    return NumberPreview(source, snapshot, number, tuple(n for n, _ in sheets), tuple(overwritten))

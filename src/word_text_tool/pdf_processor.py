"""Export a private workbook copy and verify the resulting PDF."""
from pathlib import Path
import shutil
import tempfile

from word_text_tool.file_service import prepare_output_directory
from word_text_tool.office_backend import run_office
from word_text_tool.office_validation import inspect_package, workbook_sheets
from word_text_tool.safe_output import OperationError, fingerprint, publish, require_unchanged


def export_pdf(source, output_dir=None, engine="auto", runner=run_office, expected=None):
    from pypdf import PdfReader
    source = Path(source)
    snapshot = expected or fingerprint(source)
    require_unchanged(source, snapshot)
    workbook_sheets(inspect_package(source))
    directory = prepare_output_directory(source, output_dir)
    with tempfile.TemporaryDirectory(prefix=".wordtext-pdf-", dir=directory) as temporary:
        root = Path(temporary)
        shutil.copyfile(source, root / "input.xlsx")
        require_unchanged(source, snapshot)
        result = runner({"action": "pdf", "engine": engine}, root)
        exported = root / "export.pdf"
        with exported.open("rb") as stream:
            document = PdfReader(stream, strict=True)
            if document.is_encrypted or not document.pages:
                raise OperationError("PDF 为空或无法读取")
            for page in document.pages:
                page.get_contents()
        require_unchanged(source, snapshot)
        destination = publish(exported, directory, source.stem, ".pdf")
    return destination, result

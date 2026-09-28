import base64
import json
import os
import sys
from pathlib import Path
from datetime import datetime, timedelta

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt


def excel_date_to_datetime(value):
    """Convert an Excel serial date or ISO/string date into a datetime."""
    if isinstance(value, (int, float)):
        return datetime(1899, 12, 30) + timedelta(days=float(value))

    if isinstance(value, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y-%m-%dT%H:%M:%S"):
            try:
                return datetime.strptime(value[:19], fmt)
            except ValueError:
                pass

    raise ValueError(f"Could not interpret date: {value}")


def format_date(value):
    dt = excel_date_to_datetime(value)
    return dt.strftime("%d %B %Y").lstrip("0")


def format_currency(value):
    return f"£{float(value):,.0f}"


def add_signature_block(document):
    table = document.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    left = table.cell(0, 0)
    right = table.cell(0, 1)

    left.width = Inches(3.1)
    right.width = Inches(3.1)

    for cell in (left, right):
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

    left.text = (
        "Signed for and on behalf of\n"
        "Aurelia Technologies Ltd\n\n"
        "Name: __________________________\n\n"
        "Title: ___________________________\n\n"
        "Signature: _______________________\n\n"
        "Date: ____________________________"
    )

    right.text = (
        "Signed for and on behalf of\n"
        "the Supplier\n\n"
        "Name: __________________________\n\n"
        "Title: ___________________________\n\n"
        "Signature: _______________________\n\n"
        "Date: ____________________________"
    )


def generate_document(data):
    contract_id = data["Contract ID"]
    supplier = data["Supplier"]

    current_value = data["Current Contract Value"]
    proposed_value = data["Proposed Contract Value"]

    current_expiry = data["Current Expiry Date"]
    proposed_expiry = data["Proposed Expiry Date"]

    description = data["Amendment Description"]
    effective_date = data["effective_date"]
    filename = data["document_filename"]
    amendment_request_id = data["Amendment Request ID"]

    document = Document()

    section = document.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

    normal = document.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10.5)

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = title.add_run("AMENDMENT AGREEMENT")
    run.bold = True
    run.font.size = Pt(14)

    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.add_run(f"Contract ID: {contract_id}").bold = True

    document.add_paragraph()

    intro = document.add_paragraph()
    intro.add_run("This Amendment Agreement (the “Amendment”) is made on ")
    intro.add_run(format_date(effective_date)).bold = True
    intro.add_run(" between:")

    document.add_paragraph(
        "(1) Aurelia Technologies Ltd (“Aurelia”); and"
    )

    document.add_paragraph(
        f"(2) {supplier} (the “Supplier”)."
    )

    heading = document.add_paragraph()
    heading.add_run("BACKGROUND").bold = True

    document.add_paragraph(
        f"A. Aurelia and the Supplier are parties to the agreement identified "
        f"as {contract_id} (the “Agreement”)."
    )

    document.add_paragraph(
        "B. The parties wish to amend the Agreement on the terms set out below."
    )

    heading = document.add_paragraph()
    heading.add_run("1. AMENDMENTS").bold = True

    clause_number = 1

    if float(current_value) != float(proposed_value):
        p = document.add_paragraph()
        p.add_run(f"1.{clause_number} Contract Value. ").bold = True
        p.add_run(
            f"The contract value shall be amended from "
            f"{format_currency(current_value)} to "
            f"{format_currency(proposed_value)}."
        )
        clause_number += 1

    if format_date(current_expiry) != format_date(proposed_expiry):
        p = document.add_paragraph()
        p.add_run(f"1.{clause_number} Term. ").bold = True
        p.add_run(
            f"The expiry date of the Agreement shall be amended from "
            f"{format_date(current_expiry)} to "
            f"{format_date(proposed_expiry)}."
        )
        clause_number += 1

    p = document.add_paragraph()
    p.add_run(f"1.{clause_number} Scope. ").bold = True
    p.add_run(description)

    heading = document.add_paragraph()
    heading.add_run("2. CONTINUING EFFECT").bold = True

    document.add_paragraph(
        "Except as expressly amended by this Amendment, all remaining terms "
        "of the Agreement shall continue in full force and effect."
    )

    heading = document.add_paragraph()
    heading.add_run("3. COUNTERPARTS").bold = True

    document.add_paragraph(
        "This Amendment may be executed in counterparts, each of which shall "
        "constitute an original and all of which together shall constitute "
        "one instrument."
    )

    document.add_paragraph()
    add_signature_block(document)

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_run = footer.add_run(
        f"ContractOps | Amendment Request: {amendment_request_id}"
    )
    footer_run.font.size = Pt(8)

    default_output_dir = Path(__file__).resolve().parents[1] / "generated_documents"
    output_dir = Path(
        os.environ.get("CONTRACTOPS_OUTPUT_DIR", str(default_output_dir))
    )
    output_dir.mkdir(parents=True, exist_ok=True)

    output_path = output_dir / filename
    document.save(output_path)

    return str(output_path)


def main():
    if len(sys.argv) != 2:
        raise ValueError("Expected one Base64-encoded JSON payload.")

    encoded_payload = sys.argv[1]

    decoded_payload = base64.b64decode(encoded_payload).decode("utf-8")
    data = json.loads(decoded_payload)

    output_path = generate_document(data)

    result = {
        "success": True,
        "file_path": output_path,
        "filename": os.path.basename(output_path),
        "amendment_request_id": data["Amendment Request ID"],
        "contract_id": data["Contract ID"],
    }

    print(json.dumps(result))


if __name__ == "__main__":
    main()
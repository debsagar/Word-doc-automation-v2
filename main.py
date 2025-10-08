import sys
from src.excel_reader import read_excel
from src.word_populator import populate_word
from src.pdf_converter import convert_to_pdf

def generate_report(xlsx_path, date, template_path="input/Report Sheet (DRAFT) (1).docx"):
    # Read Excel data
    df = read_excel(xlsx_path)

    # Populate Word document
    output_docx = "output/output.docx"
    populate_word(template_path, df, date, output_docx)

    # Convert to PDF
    output_pdf = "output/output.pdf"
    convert_to_pdf(output_docx, output_pdf)

    return output_pdf

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python main.py <xlsx_path> <date>")
        print("Example: python main.py input/data.xlsx 2025-01-15")
        sys.exit(1)

    xlsx_path = sys.argv[1]
    date = sys.argv[2]

    output = generate_report(xlsx_path, date)
    print(f"Generated: {output}")

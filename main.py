import pdfplumber
import pandas as pd

def extract_cells_from_pdf(pdf_path):
    cells = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                for row in table:
                    #for cell in row:
                    #    cells.append(cell)
                    cleaned_row = [cell if cell.strip() else "" for cell in row]
                    cells.append(cleaned_row)
    return cells

def save_to_excel(cells, excel_path):
    df = pd.DataFrame(cells)
    df.to_excel(excel_path, index=False, header=False)

if __name__ == "__main__":
    pdf_path = "input.pdf"
    excel_path = "output.xlsx"

    cells = extract_cells_from_pdf(pdf_path)
    save_to_excel(cells, excel_path)
    print("Data extracted and saved to Excel successfully!")

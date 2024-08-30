import os
import glob
import pdfminer.high_level
from pdfminer.layout import LTTextContainer

def extract_text_from_pdfs(folder_path):
    pdf_files = glob.glob(os.path.join(folder_path, "*.pdf"))

    for pdf_file in pdf_files:
        print(f"Processing file: {pdf_file}")

        # Extract text and layout information
        elements = []

        with open(pdf_file, 'rb') as file:
            for page_layout in pdfminer.high_level.extract_pages(file):
                for element in page_layout:
                    if isinstance(element, LTTextContainer):
                        text = element.get_text().strip()
                        x0, y0, x1, y1 = element.bbox
                        elements.append({
                            'text': text,
                            'x0': round(x0, 2),
                            'y0': round(y0, 2),
                            'x1': round(x1, 2),
                            'y1': round(y1, 2)
                        })

        # Print extracted information
        for element in elements:
            print(f"Text: {element['text']}")
            print(f"Position: X0={element['x0']}, Y0={element['y0']}, X1={element['x1']}, Y1={element['y1']}")
            print("---")

        print("\n" + "="*50 + "\n")

# Usage
folder_path = "C:\\Users\\Keisuke_Shimizu\\Desktop\\targetFolder"
extract_text_from_pdfs(folder_path)

from pdf2docx import Converter

def convert_pdf_to_docx(pdf_path, docx_path):
    # Convert PDF to DOCX
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()

if __name__ == "__main__":
    # Specify the path to your PDF file
    pdf_file_path = "Enrollment Challenges Recruiting.pdf"

    # Specify the output DOCX file path
    docx_file_path = "Enrollment Challenges Recruiting.docx"

    # Convert PDF to DOCX
    convert_pdf_to_docx(pdf_file_path, docx_file_path)

    print(f"Conversion completed. DOCX file saved to {docx_file_path}")

import PyPDF2

pdf_path = r"c:\Ramais\Ramais.pdf"
output_path = r"c:\Ramais\pdf_content.txt"

try:
    with open(pdf_path, 'rb') as f:
        pdf = PyPDF2.PdfReader(f)
        print(f"Total pages: {len(pdf.pages)}")
        
        with open(output_path, 'w', encoding='utf-8') as out:
            for page_num, page in enumerate(pdf.pages):
                text = page.extract_text()
                out.write(f"=== PAGE {page_num + 1} ===\n")
                out.write(text)
                out.write("\n\n")
        
        print(f"Conteúdo extraído salvo em: {output_path}")
except Exception as e:
    print(f"Erro: {e}")

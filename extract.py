import pdfplumber
import json
import sys

pdf_path = r'D:\OneDrive_台北科技大學\OneDrive - 國立臺北科技大學\桌面\Test.pdf'
result = []
try:
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            page_data = {'page': i+1, 'text': '', 'tables': []}
            text = page.extract_text()
            if text:
                page_data['text'] = text
            tables = page.extract_tables()
            if tables:
                for t in tables:
                    page_data['tables'].append(t)
            result.append(page_data)

    with open('extracted_content.json', 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print('Content saved to extracted_content.json')
    print(f'Total pages: {len(result)}')
    for p in result:
        print(f"Page {p['page']}: {len(p['text'])} chars, {len(p['tables'])} tables")
except Exception as e:
    print(f"Error reading PDF: {e}")
    sys.exit(1)

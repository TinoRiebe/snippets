from docx import Document

doc = 'MOG2.docx'


def extract_table_of_contents(doc_path):
    try:
        toc = []
        document = Document(doc_path)

        for paragraph in document.paragraphs:
            print(paragraph.style.name)
            if paragraph.style.name.startswith('Heading'):
                text = paragraph.text.strip()
                toc.append(paragraph.text)
        return toc

    except Exception as e:
        print('Fehler: ', e)
        return None

toc = extract_table_of_contents(doc)
print(toc)


import os
import pdfplumber
import textract
from elasticsearch import Elasticsearch
from multiprocessing import Pool, cpu_count


# Example text extraction helper (you can customize this)
def extract_text_from_docx(path):
    from docx import Document
    doc = Document(path)
    return "\n".join(p.text for p in doc.paragraphs)


# Connect to Elasticsearch
es = Elasticsearch("http://localhost:9200")


def process_document(full_path):
    try:
        filename = os.path.basename(full_path)

        if not os.path.isfile(full_path):
            return

        if not (filename.endswith(".pdf") or filename.endswith(".docx")):
            return

        # Check if already indexed
        query = {
            "query": {
                "term": {
                    "filename.keyword": filename
                }
            }
        }
        result = es.search(index="pdf_documents", body=query)

        if result['hits']['total']['value'] > 0:
            print(f"⏭️ Skipping already indexed file: {filename}")
            return

        # Extract text
        if filename.endswith(".pdf"):
            with pdfplumber.open(full_path) as pdf:
                text = "".join([page.extract_text() or "" for page in pdf.pages])
        elif filename.endswith(".docx"):
            text = extract_text_from_docx(full_path)
        # elif filename.endswith(".doc"):
        #     try:
        #         text = textract.process(full_path).decode('utf-8')
        #     except Exception as e:
        #         print(f"❌ Failed to extract .doc file {filename}: {e}")
        #         return

        # Index the document
        doc = {
            "filename": filename,
            "content": text
        }
        res = es.index(index="pdf_documents", document=doc)
        print(f"✅ Indexed {filename} with id {res['_id']}")

    except Exception as e:
        print(f"❌ Error processing {full_path}: {e}")


def insert_documents_parallel(folder_path):
    all_files = [
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if os.path.isfile(os.path.join(folder_path, f)) and
           (f.endswith(".pdf") or f.endswith(".docx"))
    ]

    print(f"🚀 Starting parallel processing with {cpu_count()} cores...")
    with Pool(cpu_count()) as pool:
        pool.map(process_document, all_files)

    print("✅ All documents processed.")


if __name__ == "__main__":
    folder_path = r"C:\Users\isset\Downloads\hansards_2025"
    insert_documents_parallel(folder_path)

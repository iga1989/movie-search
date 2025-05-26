import csv
import os

import pdfplumber as pdfplumber
from flask import Flask, request, jsonify
from elasticsearch import Elasticsearch
from pathlib import Path
from flask_cors import CORS
import textract
import docx
import re
from datetime import datetime
from typing import List
import docx  # <-- Add this at the top with other imports

app = Flask(__name__)  # 👈 this creates the real Flask app
CORS(app)

csv_path = Path(r'C:\Users\isset\Downloads\movie_metadata\movie_metadata.csv')

# create an instance of Elasticsearch
es = Elasticsearch(hosts='http://localhost:9200')

# Test connection by checking if the server is up
if es.ping():
    print("Successfully connected to Elasticsearch!")
else:
    print("Could not connect to Elasticsearch!")

from flask import jsonify


def search_movies(query):
    # Perform search on the index
    res = es.search(index="pdf_documents", body=query)
    hits = res['hits']['hits']
    filenames_2 = list({hit['_source']['filename'] for hit in hits})
    return filenames_2



def create_pdf_index():
    index_name = "pdf_documents"

    # Check if index already exists
    if es.indices.exists(index=index_name):
        print(f"Index '{index_name}' already exists.")
        return

    # Define index settings and mappings
    body = {
        "settings": {
            "analysis": {
                "analyzer": {
                    "case_sensitive": {
                        "type": "custom",
                        "tokenizer": "standard",
                        "filter": []  # No lowercase or asciifolding
                    }
                }
            }
        },
        "mappings": {
            "properties": {
                "filename": {
                    "type": "text",
                    "fields": {
                        "keyword": {
                            "type": "keyword"
                        }
                    }
                },
                "content": {
                    "type": "text",
                    "analyzer": "case_sensitive"
                }
            }
        }
    }

    # Attempt to create the index
    try:
        response = es.indices.create(index=index_name, body=body)
        print(f"✅ Index '{index_name}' created successfully.")
        print(response)
    except Exception as e:
        print(f"❌ Failed to create index '{index_name}':")
        import traceback
        traceback.print_exc()


@app.route('/search', methods=['GET'])
def filter_movies():
    name = request.args.get('name')
    query = {
        "_source": ["filename"],
        "size": 10000,
        "query": {
            "match_phrase": {
                "content": name.upper()
            }
        }
    }
    filenames_3 = search_movies(query)
    # filenames_3 = extract_formatted_dates(filenames_3)
    return jsonify({"filenames": filenames_3})



def search_filenames_by_name(name):
    query = {
        "_source": ["filename"],
        "size": 10000,
        "query": {
            "match_phrase": {
                "content": name
            }
        }
    }
    return search_movies(query)  # This should return a list of filenames

def extract_text_from_docx(docx_path):
    try:
        doc = docx.Document(docx_path)
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        print(f"Failed to extract DOCX content from {docx_path}: {e}")
        return ""


def insert_documents(folder_path):
    try:
        for filename in os.listdir(folder_path):
            full_path = os.path.join(folder_path, filename)

            # Only process files
            if not os.path.isfile(full_path):
                continue

            if filename.endswith(".pdf") or filename.endswith(".docx") or filename.endswith(".doc"):
                # Check if the document already exists
                query = {
                    "query": {
                        "term": {
                            "filename.keyword": filename
                        }
                    }
                }
                result = es.search(index="pdf_documents", body=query)

                if result['hits']['total']['value'] > 0:
                    print(f"Skipping already indexed file: {filename}")
                    continue

                # Extract text based on file type
                if filename.endswith(".pdf"):
                    with pdfplumber.open(full_path) as pdf:
                        text = ""
                        for page in pdf.pages:
                            text += page.extract_text() or ""
                elif filename.endswith(".docx"):
                    text = extract_text_from_docx(full_path)
                elif filename.endswith(".doc"):
                    try:
                        text = textract.process(full_path).decode('utf-8')
                    except Exception as e:
                        print(f"Failed to extract .doc file {filename}: {e}")
                        continue

                # Index the document
                doc = {
                    "filename": filename,
                    "content": text
                }

                res = es.index(index="pdf_documents", document=doc)
                print(f"Indexed {filename} with id {res['_id']}")
        print("Documents inserted successfully.")
    except Exception as e:
        print(f"Error while inserting documents: {e}")


def extract_formatted_dates(filenames_4: List[str]) -> List[str]:
    """
    Extracts and formats dates from filenames like 'February01_2022.docx'
    into 'February 01, 2022', sorted in descending order without duplicates.
    """
    # Match cases like February01_2022 or September2215 or March09_2023
    pattern = re.compile(r"([A-Za-z]+)(\d{1,2})(?:_)?(\d{4})")
    date_set = set()

    for name in filenames_4:
        match = pattern.search(name)
        if match:
            month_str, day, year = match.groups()
            try:
                date_obj = datetime.strptime(f"{month_str} {day} {year}", "%B %d %Y")
                date_set.add(date_obj)
            except ValueError:
                pass  # Skip invalid dates
        else:
            pass  # Skip filenames with no date match

    # Sort the unique datetime objects in descending order
    sorted_dates = sorted(date_set, reverse=True)

    # Format dates back to strings
    formatted_dates = [date.strftime("%B %d, %Y") for date in sorted_dates]

    return formatted_dates




if __name__ == '__main__':
    # insert_movies(str(csv_path))  # insert movies first
    create_pdf_index()
    # print(es.cat.indices(format="json"))
    # print([i['index'] for i in es.cat.indices(format="json")])
    # print(es.info())
    # insert_documents(r'C:\Users\isset\Downloads\hansards_2025')
    # filenames = search_filenames_by_name("YUSUF NSIBAMBI")
    # extract_formatted_dates(filenames)
    app.run(port=5002, debug=True)  # 👈 Port 5001 here

# insert_movies('C:Users/isset/Downloads/movie_metadata.csv')

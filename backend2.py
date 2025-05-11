import csv
import os

import pdfplumber as pdfplumber
from flask import Flask, request, jsonify
from elasticsearch import Elasticsearch
from pathlib import Path
from flask_cors import CORS


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
    # Extract the hits
    hits = res['hits']['hits']
    # Use a set to remove duplicate filenames, then convert back to list
    filenames = list({hit['_source']['filename'] for hit in hits})
    # Return unique filenames along with raw hits if needed
    return jsonify({
        "filenames": filenames
    })



@app.route('/search', methods=['GET'])
def filter_movies():
    # create a bool query to filter the movies
    name = request.args.get('name')
    # actors = request.args.get('actors')
    # genre = request.args.get('genre')
    # date = request.args.get('date')
    bool_query = {
        "bool": {
            "must": []
        }
    }

    # add filters to the bool query based on the provided parameters
    if name:
        bool_query["bool"]["must"].append({"match": {"content": name}})

    query = {
        "_source": ["filename"],  # only return what you need
        "size": 1000,
        "query": bool_query
    }
    # search and return the movies
    return search_movies(query)


def insert_documents(pdf_folder):
    try:
        for filename in os.listdir(pdf_folder):
            if filename.endswith(".pdf"):
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

                pdf_path = os.path.join(pdf_folder, filename)

                # Extract text from PDF
                with pdfplumber.open(pdf_path) as pdf:
                    text = ""
                    for page in pdf.pages:
                        text += page.extract_text() or ""

                # Prepare and index the document
                doc = {
                    "filename": filename,
                    "content": text
                }

                res = es.index(index="pdf_documents", document=doc)
                print(f"Indexed {filename} with id {res['_id']}")
        print("Documents inserted successfully.")
    except UnicodeDecodeError:
        print("UnicodeDecodeError: Please check the file's encoding.")
    except PermissionError:
        print(f"Permission denied when trying to open: {filename}")
    except Exception as e:
        print(f"Error: {e}")



if __name__ == '__main__':
    # insert_movies(str(csv_path))  # insert movies first
    insert_documents(r'C:\Users\isset\Downloads\hansards_2025')
    app.run(port=5002, debug=True)  # 👈 Port 5001 here

# insert_movies('C:Users/isset/Downloads/movie_metadata.csv')


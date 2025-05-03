import csv

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


def search_movies(query):
    # search movies using the Elasticsearch `search` method
    # and return the search results
    res = es.search(index="movies", body=query)
    return jsonify(res)  # 👈 wrap it


@app.route('/search', methods=['GET'])
def filter_movies():
    # create a bool query to filter the movies
    name = request.args.get('name')
    actors = request.args.get('actors')
    genre = request.args.get('genre')
    date = request.args.get('date')
    bool_query = {
        "bool": {
            "must": []
        }
    }

    # add filters to the bool query based on the provided parameters
    if name:
        bool_query["bool"]["must"].append({"match": {"name.keyword": name}})
    if actors:
        bool_query["bool"]["must"].append({"match": {"actors": actors}})
    if genre:
        bool_query["bool"]["must"].append({"match": {"genre": genre}})
    if date:
        bool_query["bool"]["must"].append({"match": {"release_date": date}})
    # create the Elasticsearch query
    query = {
        "query": bool_query
    }
    # search and return the movies
    return search_movies(query)


def insert_movies(filename):
    try:
        with open(filename, "r", encoding="ISO-8859-1") as file:
            reader = csv.DictReader(file)
            for row in reader:
                movie = {
                    "name": row["movie_title"],
                    "actors": row["actor_1_name"],
                    "genre": row["genres"],
                    "release_date": row["title_year"]
                }
                es.index(index="movies", body=movie)
        print("Movies inserted successfully.")
    except UnicodeDecodeError:
        print("UnicodeDecodeError: Please check the file's encoding.")
    except PermissionError:
        print(f"Permission denied when trying to open: {filename}")
    except Exception as e:
        print(f"Error: {e}")



if __name__ == '__main__':
    # insert_movies(str(csv_path))  # insert movies first
    insert_movies(r'C:\Users\isset\Downloads\movie_metadata\movie_metadata.csv')
    app.run(port=5001, debug=True)  # 👈 Port 5001 here

# insert_movies('C:Users/isset/Downloads/movie_metadata.csv')


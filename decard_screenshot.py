# server.py
from flask import Flask, jsonify
import traceback
from docx_utils import insert_image_example_in_memory

app = Flask(__name__)

@app.route("/api/insert_image", methods=["POST"])
def api_insert_image():
    try:
        insert_image_example_in_memory("LYtTd3Ouio2dTKxK7KflWQK8gAh","Global")
        insert_image_example_in_memory("GnaIdsY61oxqFVxan3blfnBUgBg","sg")
        return jsonify({"message": "Image inserted successfully"})
    except Exception as e:
        tb = traceback.format_exc()
        print("[api_insert_image] Error:", tb)
        return jsonify({"error": str(e), "traceback": tb}), 500

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=18891, debug=True)

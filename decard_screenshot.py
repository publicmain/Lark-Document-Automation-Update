# server.py
import logging
import traceback
from flask import Flask, jsonify

logger = logging.getLogger("server")
logger.setLevel(logging.DEBUG)
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    filename="app.log",  
    filemode="a"         
)

app = Flask(__name__)

@app.route("/api/insert_image", methods=["POST"])
def api_insert_image():
    logger.debug("recieve /api/insert_image request")
    try:
        from docx_utils import insert_image_example_in_memory  # 延迟导入
        logger.debug("call insert_image_example_in_memory START Global")
        insert_image_example_in_memory("V0Hxd7J8PoYwYNxA1z4l7dWrg3f", "Global")
        logger.debug("call insert_image_example_in_memory START sgW")
        insert_image_example_in_memory("MCXndlSkXoyVjRxlWzilA5ktgGf", "sg")
        logger.debug("two insert_image_example_in_memory calls finished")
        return jsonify({"message": "Image inserted successfully"})
    except Exception as e:
        tb = traceback.format_exc()
        logger.error("[api_insert_image] error: %s\n%s", e, tb)
        return jsonify({"error": str(e), "traceback": tb}), 500

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=18891, debug=False)

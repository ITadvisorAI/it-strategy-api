import os
import json
import logging
import threading
from flask import Flask, request, jsonify
from process_it_strategy import process_it_strategy  # You will create this next

app = Flask(__name__)
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

BASE_DIR = "temp_sessions"
os.makedirs(BASE_DIR, exist_ok=True)

@app.route("/", methods=["GET"])
def health_check():
    return "✅ IT Strategy API is live", 200

@app.route("/start_it_strategy", methods=["POST"])
def start_it_strategy():
    try:
        data = request.get_json(force=True)
        session_id = data.get("session_id")
        email = data.get("email")
        files = data.get("files", [])
        gpt_module = data.get("gpt_module", "")
        status = data.get("status", "")

        logging.info("📥 Received request:\n%s", json.dumps(data, indent=2))

        if not all([session_id, email, files]):
            logging.error("❌ Missing required fields")
            return jsonify({"error": "Missing required fields"}), 400

        folder_name = session_id if session_id.startswith("Temp_") else f"Temp_{session_id}"
        folder_path = os.path.join(BASE_DIR, folder_name)
        os.makedirs(folder_path, exist_ok=True)

        # Launch background processing
        def runner():
            try:
                process_it_strategy(session_id, email, files, folder_path)
            except Exception as e:
                logging.exception("🔥 Error inside strategy generation thread")

        threading.Thread(target=runner, daemon=True).start()
        logging.info(f"🚀 Started IT Strategy generation for session: {session_id}")

        return jsonify({"message": "IT Strategy generation started"}), 200

    except Exception as e:
        logging.exception("🔥 Failed to start IT Strategy processing")
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 12000))
    logging.info(f"🚦 IT Strategy API starting on port {port}")
    app.run(host="0.0.0.0", port=port)

"""Backend for the AI Use Case Program tool.

Serves the static HTML app and a small REST API backed by a local MySQL
database (ai_usecase_db.use_cases), replacing the browser localStorage
persistence so multiple people on the network can share the same data.
"""
import json
import os

import pymysql
from flask import Flask, jsonify, request, send_from_directory

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
HTML_FILE = "AI_Use_Case_Program12.html"

DB_CONFIG = dict(
    host=os.environ.get("DB_HOST", "127.0.0.1"),
    port=int(os.environ.get("DB_PORT", "3306")),
    user=os.environ.get("DB_USER", "root"),
    password=os.environ.get("DB_PASSWORD", ""),
    database=os.environ.get("DB_NAME", "ai_usecase_db"),
    charset="utf8mb4",
    cursorclass=pymysql.cursors.DictCursor,
    autocommit=False,
)

app = Flask(__name__, static_folder=None)


def get_conn():
    return pymysql.connect(**DB_CONFIG)


@app.route("/")
def index():
    return send_from_directory(BASE_DIR, HTML_FILE)


@app.route("/api/usecases", methods=["GET"])
def list_usecases():
    conn = get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT payload FROM use_cases ORDER BY created_at ASC")
            rows = cur.fetchall()
        data = [json.loads(r["payload"]) if isinstance(r["payload"], str) else r["payload"] for r in rows]
        return jsonify(data)
    finally:
        conn.close()


@app.route("/api/usecases", methods=["PUT"])
def replace_usecases():
    """Replace the entire use-case list in one transaction.

    The frontend keeps the full array in memory and always saves the
    whole thing at once, so mirror that here instead of diffing rows.
    """
    items = request.get_json(force=True, silent=False)
    if not isinstance(items, list):
        return jsonify({"error": "expected a JSON array"}), 400

    conn = get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM use_cases")
            for item in items:
                if not isinstance(item, dict) or "id" not in item:
                    raise ValueError("each use case must be an object with an id")
                cur.execute(
                    "INSERT INTO use_cases (id, name, dept, payload) VALUES (%s, %s, %s, %s)",
                    (
                        item["id"],
                        item.get("name", ""),
                        item.get("dept", ""),
                        json.dumps(item, ensure_ascii=False),
                    ),
                )
        conn.commit()
        return jsonify({"ok": True, "count": len(items)})
    except Exception as exc:
        conn.rollback()
        return jsonify({"error": str(exc)}), 400
    finally:
        conn.close()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8765))
    app.run(host="0.0.0.0", port=port, debug=False)

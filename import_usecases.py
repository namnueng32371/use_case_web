"""Append use cases from a JSON file (produced from an Excel input under
input_data/) into the running AI Use Case web app, without wiping existing
data.

Usage (run on the machine where `python server.py` is already running,
i.e. where MySQL + the Flask API on port 8765 are reachable):

    python import_usecases.py input_data/new_usecases_prod_ee.json

It fetches the current list via GET /api/usecases, skips any item whose
`id` already exists, appends the rest, and PUTs the full list back —
matching the frontend's "replace whole list" contract described in
CONTEXT.md. Existing use cases are never modified or removed.
"""
import json
import sys
import urllib.request

API_URL = "http://localhost:8765/api/usecases"


def main():
    if len(sys.argv) != 2:
        print("usage: python import_usecases.py <path-to-usecases.json>")
        sys.exit(1)

    with open(sys.argv[1], "r", encoding="utf-8") as f:
        new_items = json.load(f)
    if not isinstance(new_items, list):
        print("input file must contain a JSON array of use case objects")
        sys.exit(1)

    with urllib.request.urlopen(API_URL) as resp:
        existing = json.load(resp)

    existing_ids = {item.get("id") for item in existing}
    to_add = [item for item in new_items if item.get("id") not in existing_ids]
    skipped = len(new_items) - len(to_add)

    merged = existing + to_add
    body = json.dumps(merged, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        API_URL, data=body, method="PUT",
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req) as resp:
        result = json.load(resp)

    print(f"added {len(to_add)} new use case(s), skipped {skipped} duplicate id(s)")
    print("server response:", result)


if __name__ == "__main__":
    main()

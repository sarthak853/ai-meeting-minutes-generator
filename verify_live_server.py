import time
import requests
import threading
import sys
import os

from app import app

def run_server():
    app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False)

def main():
    # Start Flask server on 127.0.0.1:5000 in background thread
    server_thread = threading.Thread(target=run_server, daemon=True)
    server_thread.start()

    base_url = "http://127.0.0.1:5000"
    print(f"Waiting for Flask server on {base_url}...")

    # Wait for server to come up
    session = requests.Session()
    up = False
    for i in range(15):
        try:
            r = session.get(f"{base_url}/", timeout=1)
            if r.status_code == 200:
                up = True
                print("Server is UP and responding!")
                break
        except Exception:
            time.sleep(0.5)

    if not up:
        print("ERROR: Server failed to start within timeout.")
        sys.exit(1)

    print("\n=======================================================")
    print("STEP 1: Verify GET / Web Dashboard")
    print("=======================================================")
    r = session.get(f"{base_url}/")
    assert r.status_code == 200, f"Expected 200, got {r.status_code}"
    assert "MinuteCraft AI" in r.text
    print("GET / passed. Web Dashboard UI rendered successfully.")

    print("\n=======================================================")
    print("STEP 2: Verify POST /generate with sample_transcript.txt")
    print("=======================================================")
    with open("sample_transcript.txt", "r", encoding="utf-8") as f:
        transcript_content = f.read()

    r = session.post(
        f"{base_url}/generate",
        json={"transcript": transcript_content},
        headers={"Content-Type": "application/json"}
    )
    assert r.status_code == 200, f"Expected 200, got {r.status_code} ({r.text})"
    minutes = r.json()
    print("JSON parsed cleanly without errors.")
    print(f"Meeting Title: {minutes.get('title')}")
    print(f"Date: {minutes.get('date')}")
    print(f"Attendees: {minutes.get('attendees')}")

    # Verify Confirmed Decisions vs Pending Discussions
    decisions = minutes.get("decisions", [])
    confirmed = [d for d in decisions if d.get("status") == "confirmed"]
    pending = [d for d in decisions if d.get("status") == "pending"]

    print("\n[Decisions Segregation Check]")
    print(f"Confirmed Decisions ({len(confirmed)}):")
    for c in confirmed:
        print(f"  - {c.get('description')}")
    print(f"Pending Discussions ({len(pending)}):")
    for p in pending:
        print(f"  - {p.get('description')}")

    assert len(confirmed) >= 1, "Must contain confirmed decisions"
    assert len(pending) >= 1, "Must contain pending discussions"

    confirmed_str = " ".join([d.get("description", "") for d in confirmed]).lower()
    pending_str = " ".join([d.get("description", "") for d in pending]).lower()

    assert "light mode" in confirmed_str, "Light mode for Q3 must be under confirmed decisions"
    assert "archiving" in pending_str, "Account archiving must be under pending discussions"
    print("Decision segregation verified: Confirmed decisions (Light mode for Q3) segregated from pending discussions (Account archiving).")

    # Verify Action Items & Deadlines Fallback
    print("\n[Action Items & Deadlines Fallback Check]")
    action_items = minutes.get("action_items", [])
    print(f"Action Items ({len(action_items)}):")
    deadlines = []
    for item in action_items:
        print(f"  - [{item.get('owner')}] {item.get('task')} (Deadline: {item.get('deadline')})")
        deadlines.append(item.get("deadline"))

    assert "Not specified" in deadlines, "Missing deadlines must fallback to 'Not specified'"
    print("Deadlines fallback verified: Missing deadlines fallback to 'Not specified'.")

    print("\n=======================================================")
    print("STEP 3: Verify POST /export/pdf (ReportLab)")
    print("=======================================================")
    r_pdf = session.post(
        f"{base_url}/export/pdf",
        json=minutes,
        headers={"Content-Type": "application/json"}
    )
    assert r_pdf.status_code == 200, f"Expected 200, got {r_pdf.status_code}"
    assert r_pdf.headers.get("Content-Type") == "application/pdf"
    pdf_bytes = r_pdf.content
    assert pdf_bytes.startswith(b"%PDF-"), f"Invalid PDF signature: {pdf_bytes[:10]}"
    print(f"PDF Export Verified: {len(pdf_bytes)} bytes streamed with signature {pdf_bytes[:7]}")

    print("\n=======================================================")
    print("STEP 4: Verify POST /export/docx (python-docx)")
    print("=======================================================")
    r_docx = session.post(
        f"{base_url}/export/docx",
        json=minutes,
        headers={"Content-Type": "application/json"}
    )
    assert r_docx.status_code == 200, f"Expected 200, got {r_docx.status_code}"
    assert "wordprocessingml.document" in r_docx.headers.get("Content-Type", "")
    docx_bytes = r_docx.content
    assert docx_bytes.startswith(b"PK\x03\x04") or docx_bytes.startswith(b"PK"), f"Invalid DOCX signature: {docx_bytes[:4]}"
    print(f"DOCX Export Verified: {len(docx_bytes)} bytes streamed with zip signature {docx_bytes[:4]}")

    print("\n=======================================================")
    print("ALL AUTOMATED VERIFICATION CHECKS PASSED SUCCESSFULLY!")
    print("=======================================================")

if __name__ == "__main__":
    main()

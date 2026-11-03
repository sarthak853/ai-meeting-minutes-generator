import os
import json
import unittest
from io import BytesIO

from app import app, clean_json_markdown, process_transcript
from docx import Document

class TestMeetingMinutesApp(unittest.TestCase):
    def setUp(self):
        self.app = app.test_client()
        self.app.testing = True

        with open("sample_transcript.txt", "r", encoding="utf-8") as f:
            self.sample_transcript = f.read()

    def test_markdown_stripping(self):
        """Test defensive stripping of markdown code fences."""
        raw = "```json\n{\"test\": 123}\n```"
        cleaned = clean_json_markdown(raw)
        self.assertEqual(cleaned, '{"test": 123}')

        raw2 = "```\n{\"test\": 456}\n```"
        cleaned2 = clean_json_markdown(raw2)
        self.assertEqual(cleaned2, '{"test": 456}')

    def test_index_route(self):
        """Test GET / returns 200 and serves HTML interface."""
        response = self.app.get("/")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"MinuteCraft AI", response.data)
        self.assertIn(b"Generate Meeting Minutes", response.data)

    def test_generate_endpoint_with_json(self):
        """Test POST /generate with sample_transcript.txt."""
        response = self.app.post(
            "/generate",
            data=json.dumps({"transcript": self.sample_transcript}),
            content_type="application/json"
        )
        self.assertEqual(response.status_code, 200)
        data = response.get_json()

        # 1. Clean JSON response without parsing errors
        self.assertIsInstance(data, dict)
        self.assertIn("title", data)
        self.assertIn("decisions", data)
        self.assertIn("action_items", data)

        # 2. Confirmed decisions segregated from pending discussions
        decisions = data.get("decisions", [])
        confirmed = [d for d in decisions if d.get("status") == "confirmed"]
        pending = [d for d in decisions if d.get("status") == "pending"]

        self.assertGreaterEqual(len(confirmed), 1, "Must have at least one confirmed decision")
        self.assertGreaterEqual(len(pending), 1, "Must have at least one pending discussion")

        confirmed_text = " ".join([d.get("description", "") for d in confirmed]).lower()
        pending_text = " ".join([d.get("description", "") for d in pending]).lower()

        self.assertIn("light mode", confirmed_text, "Confirmed decisions must include Light mode for Q3")
        self.assertIn("archiving", pending_text, "Pending discussions must include Account archiving")

        # 3. Missing deadlines fallback to 'Not specified'
        action_items = data.get("action_items", [])
        self.assertGreaterEqual(len(action_items), 1)

        deadlines = [item.get("deadline") for item in action_items]
        self.assertIn("Not specified", deadlines, "Missing deadlines must fallback to 'Not specified'")

        print("\n--- GENERATE TEST PASSED ---")
        print(f"Title: {data.get('title')}")
        print(f"Confirmed: {[d['description'] for d in confirmed]}")
        print(f"Pending: {[d['description'] for d in pending]}")
        print(f"Action Items: {action_items}")

    def test_generate_endpoint_with_docx_upload(self):
        """Test POST /generate with uploaded .docx file."""
        docx_buffer = BytesIO()
        doc = Document()
        doc.add_paragraph(self.sample_transcript)
        doc.save(docx_buffer)
        docx_buffer.seek(0)

        response = self.app.post(
            "/generate",
            data={"file": (docx_buffer, "transcript.docx")},
            content_type="multipart/form-data"
        )
        self.assertEqual(response.status_code, 200)
        data = response.get_json()
        self.assertIsInstance(data, dict)
        self.assertIn("decisions", data)

    def test_export_pdf_endpoint(self):
        """Test POST /export/pdf returns valid binary with %PDF- signature."""
        payload = {
            "title": "Product & Engineering Sync",
            "date": "October 24, 2026",
            "attendees": ["Alex", "Sarah", "David", "Priya"],
            "summary": "Team agreed to confirm Light mode for Q3 while keeping account archiving pending.",
            "decisions": [
                {"status": "confirmed", "description": "Launch Light mode for Q3 release."},
                {"status": "pending", "description": "Account archiving policy under legal review."}
            ],
            "action_items": [
                {"task": "Finalize Light mode UI", "owner": "Alex", "deadline": "August 15th"},
                {"task": "Legal review on archiving", "owner": "Sarah", "deadline": "Not specified"}
            ],
            "key_discussions": [
                {"topic": "Theme Releases", "points": ["High contrast light mode in demand."]}
            ]
        }

        response = self.app.post(
            "/export/pdf",
            data=json.dumps(payload),
            content_type="application/json"
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "application/pdf")
        # Validate %PDF- header signature
        self.assertTrue(response.data.startswith(b"%PDF-"), f"Expected %PDF- prefix, got {response.data[:10]}")
        self.assertGreater(len(response.data), 1000)
        print(f"\n--- PDF EXPORT TEST PASSED (Size: {len(response.data)} bytes, Header: {response.data[:7]}) ---")

    def test_export_docx_endpoint(self):
        """Test POST /export/docx returns valid binary with PK.. zip header."""
        payload = {
            "title": "Product & Engineering Sync",
            "date": "October 24, 2026",
            "attendees": ["Alex", "Sarah", "David", "Priya"],
            "summary": "Team agreed to confirm Light mode for Q3 while keeping account archiving pending.",
            "decisions": [
                {"status": "confirmed", "description": "Launch Light mode for Q3 release."},
                {"status": "pending", "description": "Account archiving policy under legal review."}
            ],
            "action_items": [
                {"task": "Finalize Light mode UI", "owner": "Alex", "deadline": "August 15th"},
                {"task": "Legal review on archiving", "owner": "Sarah", "deadline": "Not specified"}
            ],
            "key_discussions": [
                {"topic": "Theme Releases", "points": ["High contrast light mode in demand."]}
            ]
        }

        response = self.app.post(
            "/export/docx",
            data=json.dumps(payload),
            content_type="application/json"
        )

        self.assertEqual(response.status_code, 200)
        self.assertIn("wordprocessingml.document", response.mimetype)
        # Validate PK.. (PK\x03\x04) zip signature for docx
        self.assertTrue(response.data.startswith(b"PK\x03\x04") or response.data.startswith(b"PK"), f"Expected PK prefix, got {response.data[:4]}")
        self.assertGreater(len(response.data), 1000)
        print(f"\n--- DOCX EXPORT TEST PASSED (Size: {len(response.data)} bytes, Header: {response.data[:4]}) ---")

if __name__ == "__main__":
    unittest.main()

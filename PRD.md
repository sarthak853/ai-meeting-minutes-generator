# Product Requirements Document (PRD): AI-Powered Meeting Minutes Generator

## 1. Project Overview
The AI-Powered Meeting Minutes Generator is an enterprise-grade web application built to convert unstructured, conversational meeting transcripts into structured, actionable, and verifiable meeting minutes.

## 2. Core Directives & Free API Architecture
- **LLM Engine**: Google Gemini API utilizing `gemini-1.5-flash` model under the 100% free tier (Google AI Studio: 15 RPM, 1,500 RPD, zero credit card requirement).
- **Environment**: Python virtual environment (`python -m venv venv`), dependencies managed via `requirements.txt` (`Flask`, `python-docx`, `reportlab`, `google-generativeai`, `python-dotenv`).
- **Configuration**: Strictly load `GEMINI_API_KEY` from `.env` via `python-dotenv`.

## 3. Endpoints Specification
- `GET /`: Serves the primary web dashboard (`templates/index.html`).
- `POST /generate`:
  - Accepts raw transcript text via JSON body `{"transcript": "..."}` or multipart file uploads (`.txt`, `.docx`).
  - Calls `process_transcript()` to extract structured meeting minutes.
  - Returns strict JSON schema response.
- `POST /export/pdf`:
  - Receives live edited DOM data payload from the client.
  - Uses `ReportLab` to render and stream a well-formatted PDF document (`application/pdf`) with `%PDF-` signature.
- `POST /export/docx`:
  - Receives live edited DOM data payload from the client.
  - Uses `python-docx` to construct and stream a Word document (`application/vnd.openxmlformats-officedocument.wordprocessingml.document`) with `PK..` zip header.

## 4. Extraction & Prompt Requirements (`process_transcript()`)
- **Zero-Hallucination & Anti-Inventive System Prompt**: Extract exclusively what is explicitly mentioned in the transcript.
- **Strict JSON Schema**:
  ```json
  {
    "title": "Meeting Title / Topic",
    "date": "Meeting Date (or 'Not specified')",
    "attendees": ["Person 1", "Person 2"],
    "summary": "Executive summary paragraph",
    "decisions": [
      {
        "status": "confirmed",
        "description": "Decision that was officially agreed upon (e.g., Light mode for Q3)"
      },
      {
        "status": "pending",
        "description": "Topic discussed without final consensus (e.g., Account archiving)"
      }
    ],
    "action_items": [
      {
        "task": "Description of action item",
        "owner": "Assignee name",
        "deadline": "Due date or 'Not specified'"
      }
    ],
    "key_discussions": [
      {
        "topic": "Topic Name",
        "points": ["Discussion point 1", "Discussion point 2"]
      }
    ]
  }
  ```
- **Markdown Stripping**: Robust stripping of ````json ... ```` formatting fences before `json.loads()` parsing.
- **Segregation**: Distinctly separate confirmed decisions from pending discussions.
- **Fallback Rule**: Any missing deadline must strictly fallback to `"Not specified"`.

## 5. Frontend UI/UX Specifications (`templates/index.html`)
- Modern, clean, responsive UI with sleek gradient accents and accessible typography.
- Direct clipboard paste and drag-and-drop/file-picker support for `.txt` and `.docx` files.
- Loading transitions (spinner / progress animations) during generation.
- Interactive `contenteditable` dashboard containers for inline editing of titles, summaries, decisions, action items, and discussion points.
- Live Export triggers to `/export/pdf` and `/export/docx` passing the currently edited DOM content.

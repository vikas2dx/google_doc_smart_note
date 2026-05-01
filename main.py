import os
import traceback
from googleapiclient.discovery import build
from google.oauth2 import service_account
import google.generativeai as genai

SHEET_ID = os.environ.get("SHEET_ID")
DOC_ID = os.environ.get("DOC_ID")
GEMINI_API_KEY = os.environ.get("GEMINI_KEY")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/documents"
]

def run_agent(request):
    try:
        # ---------------- AUTH ----------------
        if not os.path.exists("credentials.json"):
            return "❌ credentials.json not found"

        creds = service_account.Credentials.from_service_account_file(
            "credentials.json", scopes=SCOPES
        )

        sheets_service = build("sheets", "v4", credentials=creds)
        docs_service = build("docs", "v1", credentials=creds)

        # ---------------- READ SHEET ----------------
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range="Sheet1!A:C"
        ).execute()

        rows = result.get("values", [])
        if not rows or len(rows) < 2:
            return "No data"

        headers = rows[0]
        data = rows[1:]

        # ---------------- FIND TASK ----------------
        task, row_num = None, None
        for i, row in enumerate(data):
            row += [""] * (3 - len(row))
            row_dict = dict(zip(headers, row))

            if row_dict.get("Status") == "Not Started":
                task = row_dict
                row_num = i + 2
                break

        if not task:
            return "🎉 No pending tasks"

        subtopic = task.get("Subtopic")

        # ---------------- GEMINI ----------------
        genai.configure(api_key=GEMINI_API_KEY)
        model = genai.GenerativeModel("gemini-2.5-flash")

        prompt = f"""
STRICT STUDY FORMAT:

TITLE: {subtopic}

SECTION: Overview
- Explain in short bullets (not paragraph)
- Keep concise

SECTION: Key Concepts

SUBSECTION: Concept Name
Definition:
- Bullet
- Bullet

Example:
- Bullet

SUBSECTION: Concept Name
Definition:
- Bullet

SECTION: Important Points
- Only bullets
- No long paragraphs

SECTION: Practical Use Case

Steps:
1. Step
2. Step
3. Step

SECTION: Quick Revision
- Bullet summary

RULES:
- No paragraphs
- No asterisks
- Always use bullets
- Keep short and crisp
"""

        response = model.generate_content(prompt)
        content = (response.text or "").replace("*", "")

        # ---------------- WRITE TO DOC ----------------
        doc = docs_service.documents().get(documentId=DOC_ID).execute()
        current_pos = doc["body"]["content"][-1]["endIndex"] - 1

        lines = content.split("\n")
        requests = []
        bullet_ranges = []
        numbered_ranges = []

        for line in lines:
            stripped = line.strip()

            if not stripped:
                requests.append({
                    "insertText": {"location": {"index": current_pos}, "text": "\n"}
                })
                current_pos += 1
                continue

            style = "NORMAL_TEXT"
            text = stripped + "\n"
            is_bullet = False
            is_numbered = False

            # ---------- STRUCTURE ----------
            if stripped.startswith("TITLE:"):
                style = "HEADING_1"
                text = stripped.replace("TITLE:", "").strip() + "\n"

            elif stripped.startswith("SECTION:"):
                style = "HEADING_2"
                text = stripped.replace("SECTION:", "").strip() + "\n"

            elif stripped.startswith("SUBSECTION:"):
                style = "HEADING_3"
                text = stripped.replace("SUBSECTION:", "").strip() + "\n"

            elif stripped.startswith("- "):
                text = stripped[2:].strip() + "\n"
                is_bullet = True

            elif stripped[:2].isdigit() and stripped[2] == ".":
                text = stripped + "\n"
                is_numbered = True

            start = current_pos
            end = current_pos + len(text)

            # Insert text
            requests.append({
                "insertText": {
                    "location": {"index": current_pos},
                    "text": text
                }
            })

            # Style
            requests.append({
                "updateParagraphStyle": {
                    "range": {"startIndex": start, "endIndex": end},
                    "paragraphStyle": {
                        "namedStyleType": style,
                        "spaceAbove": {"magnitude": 10, "unit": "PT"},
                        "spaceBelow": {"magnitude": 6, "unit": "PT"}
                    },
                    "fields": "namedStyleType,spaceAbove,spaceBelow"
                }
            })

            if is_bullet:
                bullet_ranges.append((start, end))

            if is_numbered:
                numbered_ranges.append((start, end))

            current_pos = end

        # ---------- APPLY BULLETS ----------
        for start, end in bullet_ranges:
            requests.append({
                "createParagraphBullets": {
                    "range": {"startIndex": start, "endIndex": end},
                    "bulletPreset": "BULLET_DISC_CIRCLE_SQUARE"
                }
            })

        # ---------- APPLY NUMBERED LIST ----------
        for start, end in numbered_ranges:
            requests.append({
                "createParagraphBullets": {
                    "range": {"startIndex": start, "endIndex": end},
                    "bulletPreset": "NUMBERED_DECIMAL"
                }
            })

        # ---------- SEPARATOR ----------
        separator = "\n\n====================================\n\n"
        requests.append({
            "insertText": {
                "location": {"index": current_pos},
                "text": separator
            }
        })

        # Execute
        docs_service.documents().batchUpdate(
            documentId=DOC_ID,
            body={"requests": requests}
        ).execute()

        # ---------------- UPDATE SHEET ----------------
        sheets_service.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"Sheet1!C{row_num}",
            valueInputOption="RAW",
            body={"values": [["Completed"]]}
        ).execute()

        return f"✅ Done: {subtopic}"

    except Exception as e:
        print(traceback.format_exc())
        return f"❌ Error: {str(e)}"
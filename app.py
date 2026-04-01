import fitz
import requests
import json
import re
import os

# =========================
# 🔹 OLLAMA CALL
# =========================
def ask_ollama(prompt):
    response = requests.post(
        "http://localhost:11434/api/generate",
        json={
            "model": "llama3",
            "prompt": prompt,
            "stream": False
        }
    )
    return response.json()["response"]


# =========================
# 🔹 EXTRACT WORDS → LINES
# =========================
def extract_lines(page):
    words = page.get_text("words")

    # Sort top → bottom, left → right
    words.sort(key=lambda w: (w[1], w[0]))

    lines = []
    current_line = []
    current_y = None

    for w in words:
        x0, y0, x1, y1, text, *_ = w

        if current_y is None:
            current_y = y0

        if abs(y0 - current_y) < 5:
            current_line.append((text, y0, y1))
        else:
            line_text = " ".join([t[0] for t in current_line])
            lines.append({
                "text": line_text,
                "y0": current_line[0][1],
                "y1": current_line[0][2]
            })

            current_line = [(text, y0, y1)]
            current_y = y0

    # Last line
    if current_line:
        line_text = " ".join([t[0] for t in current_line])
        lines.append({
            "text": line_text,
            "y0": current_line[0][1],
            "y1": current_line[0][2]
        })

    return lines


# =========================
# 🔹 JSON EXTRACTION
# =========================
def extract_json(text):
    match = re.search(r'\{.*\}', text, re.DOTALL)
    return match.group(0) if match else None


# =========================
# 🔹 MAP TEXT → INDICES
# =========================
def map_text_to_indices(lines, texts):
    indices = []

    for i, line in enumerate(lines):
        for t in texts:
            if t.lower() in line["text"].lower():
                indices.append(i)

    return list(set(indices))


# =========================
# 🔹 HEURISTIC DETECTION
# =========================
def detect_affiliation_lines(lines):
    keywords = ["university", "institute", "department", "@", "college", "school"]

    indices = []
    for i, line in enumerate(lines):
        text = line["text"].lower()
        if any(k in text for k in keywords):
            indices.append(i)

    return indices


# =========================
# 🔹 LLM CLASSIFICATION
# =========================
def classify_lines_with_llm(lines):
    text_input = "\n".join([f"{i}: {l['text']}" for i, l in enumerate(lines)])

    prompt = f"""
Return ONLY valid JSON. No explanation.

You are analyzing the first page of a research paper.

Tasks:
1. Identify TITLE lines (can be multiple or none)
2. Identify AUTHOR + AFFILIATION lines (can be multiple or none)

Rules:
- Title = large heading at top
- Authors = names, emails, universities
- Affiliations = institute, department, @

IMPORTANT:
- Return ONLY line numbers (integers)
- If none → return empty list

Format:
{{
  "title": [0,1],
  "authors": [2,3,4]
}}

Text:
{text_input}
"""

    response = ask_ollama(prompt)

    print("\n--- LLM RAW RESPONSE ---\n", response)

    json_text = extract_json(response)

    if not json_text:
        return {"title": [], "authors": []}

    try:
        return json.loads(json_text)
    except:
        return {"title": [], "authors": []}


# =========================
# 🔹 MAIN ANONYMIZER
# =========================
def anonymize_pdf(input_path, output_path):
    print(f"\n📄 Processing: {input_path}")

    doc = fitz.open(input_path)
    page = doc[0]

    lines = extract_lines(page)

    # Debug
    print("\n--- EXTRACTED LINES ---")
    for i, l in enumerate(lines):
        print(i, ":", l["text"])

    result = classify_lines_with_llm(lines)

    title_lines = result.get("title", [])
    author_lines = result.get("authors", [])

    # Fix if LLM returns text instead of indices
    if title_lines and isinstance(title_lines[0], str):
        title_lines = map_text_to_indices(lines, title_lines)

    if author_lines and isinstance(author_lines[0], str):
        author_lines = map_text_to_indices(lines, author_lines)

    # Heuristic fallback
    heuristic_lines = detect_affiliation_lines(lines)
    author_lines = list(set(author_lines + heuristic_lines))

    print("\nDetected title lines:", title_lines)
    print("Detected author/affiliation lines:", author_lines)

    # Skip if nothing found
    if not author_lines:
        print("⚠️ No author info found → skipping")
        doc.save(output_path)
        return

    # Redaction region
    y_top = min(lines[i]["y0"] for i in author_lines)
    y_bottom = max(lines[i]["y1"] for i in author_lines)

    rect = page.rect
    redact_area = fitz.Rect(rect.x0, y_top - 5, rect.x1, y_bottom + 5)

    print("🧹 Redacting area:", redact_area)

    page.add_redact_annot(redact_area, fill=(1, 1, 1))
    page.apply_redactions()

    # Remove metadata
    doc.set_metadata({})

    doc.save(output_path)
    doc.close()


# =========================
# 🔹 BATCH PROCESS
# =========================
def process_folder(input_dir="Input", output_dir="Output"):
    os.makedirs(output_dir, exist_ok=True)

    for folder in os.listdir(input_dir):
        submission_path = os.path.join(input_dir, folder, "Submission")

        if not os.path.exists(submission_path):
            continue

        for file in os.listdir(submission_path):
            if file.endswith(".pdf"):
                input_file = os.path.join(submission_path, file)
                output_file = os.path.join(output_dir, f"{folder}_{file}")

                try:
                    anonymize_pdf(input_file, output_file)
                except Exception as e:
                    print("❌ Error:", e)


# =========================
# 🔹 RUN
# =========================
if __name__ == "__main__":
    process_folder()
import fitz
import requests
import json
import re
import os
import win32com.client


# =========================
# 🔹 OLLAMA
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
# 🔹 WORD → PDF
# =========================
def convert_to_pdf_msword(input_file, temp_dir):
    os.makedirs(temp_dir, exist_ok=True)

    base = os.path.splitext(os.path.basename(input_file))[0]
    output_file = os.path.join(temp_dir, base + ".pdf")

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    doc = word.Documents.Open(os.path.abspath(input_file))
    doc.SaveAs(os.path.abspath(output_file), FileFormat=17)
    doc.Close()
    word.Quit()

    return output_file


# =========================
# 🔹 EXTRACT WORDS → LINES
# =========================
def extract_lines(page):
    words = page.get_text("words")
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
            lines.append({"text": line_text, "y0": current_line[0][1], "y1": current_line[0][2]})
            current_line = [(text, y0, y1)]
            current_y = y0

    if current_line:
        line_text = " ".join([t[0] for t in current_line])
        lines.append({"text": line_text, "y0": current_line[0][1], "y1": current_line[0][2]})

    return lines


# =========================
# 🔹 JSON EXTRACT
# =========================
def extract_json(text):
    match = re.search(r'\{.*\}', text, re.DOTALL)
    return match.group(0) if match else None


# =========================
# 🔹 NORMALIZE INDICES
# =========================
def normalize_indices(lst):
    result = []

    for item in lst:
        if isinstance(item, int):
            result.append(item)

        elif isinstance(item, list):
            result.extend(normalize_indices(item))

        elif isinstance(item, dict):
            for v in item.values():
                if isinstance(v, int):
                    result.append(v)

    return result


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
# 🔹 FIND TITLE BOTTOM  ✅ FIX
# =========================
def find_title_bottom(lines):
    title_candidates = lines[:8]
    title_candidates = [l for l in title_candidates if len(l["text"]) > 10]

    if not title_candidates:
        return 100

    return max(l["y1"] for l in title_candidates)


# =========================
# 🔹 FIND ABSTRACT  ✅ FIX
# =========================
def find_abstract_y(lines):
    for line in lines:
        text = line["text"].lower()

        if (
            "abstract" in text or
            text.startswith("abstract") or
            "abstract—" in text or
            "abstract:" in text
        ):
            return line["y0"]

    return None


# =========================
# 🔹 LLM CLASSIFICATION
# =========================
def classify_lines(lines):
    text_input = "\n".join([f"{i}: {l['text']}" for i, l in enumerate(lines)])

    prompt = f"""
Return ONLY JSON.

Format:
{{
  "title": [indices],
  "authors": [indices]
}}

Rules:
- Title = heading at top
- Authors = names, affiliations, emails
- If none → []

Text:
{text_input}
"""

    response = ask_ollama(prompt)

    json_text = extract_json(response)

    if not json_text:
        return {"title": [], "authors": []}

    try:
        return json.loads(json_text)
    except:
        return {"title": [], "authors": []}


# =========================
# 🔹 ANONYMIZE PDF (FIXED)
# =========================
def anonymize_pdf(input_pdf, output_pdf):
    doc = fitz.open(input_pdf)
    page = doc[0]

    lines = extract_lines(page)
    result = classify_lines(lines)

    raw_authors = result.get("authors", [])

    # Step 1: map strings if needed
    if raw_authors and any(isinstance(x, str) for x in raw_authors):
        mapped = map_text_to_indices(lines, raw_authors)
    else:
        mapped = raw_authors

    # Step 2: normalize
    author_lines = normalize_indices(mapped)

    # Step 3: validate indices
    author_lines = [i for i in author_lines if isinstance(i, int) and i < len(lines)]

    print("Detected author lines:", author_lines)

    # 🔥 FIX: Use layout boundaries instead
    title_bottom = find_title_bottom(lines)
    abstract_y = find_abstract_y(lines)

    if not abstract_y:
        print("⚠️ Abstract not found → skipping")
        doc.save(output_pdf)
        return False

    top = title_bottom + 5
    bottom = abstract_y - 5

    if bottom <= top:
        print("⚠️ Invalid region → skipping")
        doc.save(output_pdf)
        return False

    rect = page.rect
    redact_area = fitz.Rect(rect.x0, top, rect.x1, bottom)

    page.add_redact_annot(redact_area, fill=(1, 1, 1))
    page.apply_redactions()

    doc.set_metadata({})
    doc.save(output_pdf)
    doc.close()

    return True


# =========================
# 🔹 MAIN PIPELINE
# =========================
def process_all(input_dir="Input", output_dir="Output"):
    os.makedirs(output_dir, exist_ok=True)

    converted = []
    anonymized = []
    skipped = []

    temp_dir = "temp_pdf"

    for root, _, files in os.walk(input_dir):
        for file in files:
            input_path = os.path.join(root, file)

            try:
                if file.endswith((".doc", ".docx")):
                    pdf_path = convert_to_pdf_msword(input_path, temp_dir)
                    converted.append(file)
                elif file.endswith(".pdf"):
                    pdf_path = input_path
                else:
                    continue

                output_path = os.path.join(output_dir, f"anon_{os.path.basename(pdf_path)}")

                result = anonymize_pdf(pdf_path, output_path)

                if result:
                    anonymized.append(file)
                else:
                    skipped.append(file)

            except Exception as e:
                print(f"❌ Error processing {file}: {e}")
                skipped.append(file)

    print("\n========== SUMMARY ==========")

    print("\n📄 Converted Files:")
    for f in converted:
        print("-", f)

    print("\n🧹 Anonymized Files:")
    for f in anonymized:
        print("-", f)

    print("\n⚠️ Skipped Files:")
    for f in skipped:
        print("-", f)


# =========================
# 🔹 RUN
# =========================
if __name__ == "__main__":
    process_all()
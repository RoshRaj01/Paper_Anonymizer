import os
import sys
import win32com.client


def main():
    if len(sys.argv) < 2:
        print("Usage: python convert_to_pdf_msword.py <input_file> [output_file]")
        sys.exit(1)

    input_file = os.path.abspath(sys.argv[1])

    if not os.path.exists(input_file):
        print("❌ File not found:", input_file)
        sys.exit(1)

    # Output file
    if len(sys.argv) >= 3:
        output_file = os.path.abspath(sys.argv[2])
    else:
        base = os.path.splitext(input_file)[0]
        output_file = base + ".pdf"

    print("📄 Converting using Microsoft Word...")
    print("Input :", input_file)
    print("Output:", output_file)

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        doc = word.Documents.Open(input_file)

        # 17 = wdFormatPDF
        doc.SaveAs(output_file, FileFormat=17)

        doc.Close()
        word.Quit()

        print("✅ Conversion successful!")

    except Exception as e:
        print("❌ Conversion failed:", e)
        sys.exit(1)


if __name__ == "__main__":
    main()
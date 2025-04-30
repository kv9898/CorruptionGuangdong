import os
import json
from pathlib import Path
import win32com.client

VERDICTS_RAW_DIR = Path(os.getcwd()) / "verdicts" / "raw"
VERDICTS_OUTPUT_IR = Path(os.getcwd()) / "verdicts" / "docx"
FAILED_LOG_PATH = Path(os.getcwd()) / "verdicts" / "failed_doc_conversions.json"

failed_files = []


def convert_doc_to_docx(word_app, filename: str) -> str:
    try:
        input_path = VERDICTS_RAW_DIR / filename
        output_path = VERDICTS_OUTPUT_IR / filename.replace(".doc", ".docx")

        # Skip if already converted
        if output_path.exists():
            print(f"⏩ Skipping (already exists): {filename}")
            return None

        doc = word_app.Documents.Open(str(input_path))
        doc.SaveAs(str(output_path), FileFormat=16)
        doc.Close()
        print(f"✅ Converted: {filename}")
        return str(output_path)

    except Exception as e:
        print(f"❌ Failed to convert {filename}: {e}")
        failed_files.append(filename)


# Ensure output directory exists
VERDICTS_OUTPUT_IR.mkdir(parents=True, exist_ok=True)

try:
    # Start Word once
    word = win32com.client.Dispatch("Word.Application")
    word.DisplayAlerts = 0
    word.Visible = False

    # Process all .doc files
    for filename in os.listdir(VERDICTS_RAW_DIR):
        if filename.lower().endswith(".doc"):
            convert_doc_to_docx(word, filename)

except KeyboardInterrupt:
    print("\n⛔ Interrupted by user.")

finally:
    if "word" in locals():
        word.Quit()

    # Write failed files to JSON
    if failed_files:
        with open(FAILED_LOG_PATH, "w", encoding="utf-8") as f:
            json.dump(failed_files, f, ensure_ascii=False, indent=2)
        print(f"\n❗ Failed files written to: {FAILED_LOG_PATH}")
    else:
        print("\n✅ All files converted successfully.")

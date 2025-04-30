import os
import json
from pathlib import Path
import win32com.client
from tqdm import tqdm

VERDICTS_RAW_DIR = Path(os.getcwd()) / "verdicts" / "raw"
VERDICTS_OUTPUT_IR = Path(os.getcwd()) / "verdicts" / "docx"
FAILED_LOG_PATH = Path(os.getcwd()) / "verdicts" / "failed_doc_conversions.json"

failed_files = []

def convert_doc_to_docx(word_app, filename: str, index: int, total: int) -> str:
    try:
        input_path = VERDICTS_RAW_DIR / filename
        output_path = VERDICTS_OUTPUT_IR / filename.replace(".doc", ".docx")

        # Skip if already converted
        if output_path.exists():
            tqdm.write(f"⏩ Skipping: {filename} ({index + 1}/{total})")
            return None

        doc = word_app.Documents.Open(str(input_path))
        doc.SaveAs(str(output_path), FileFormat=16)
        doc.Close()
        return str(output_path)

    except Exception as e:
        tqdm.write(f"❌ Failed to convert {filename}: {e} ({index + 1}/{total})")
        failed_files.append(filename)

# Ensure output directory exists
VERDICTS_OUTPUT_IR.mkdir(parents=True, exist_ok=True)

try:
    all_docs = [f for f in os.listdir(VERDICTS_RAW_DIR) if f.lower().endswith(".doc")]
    total_files = len(all_docs)

    word = win32com.client.Dispatch("Word.Application")
    word.DisplayAlerts = 0
    word.Visible = False

    with tqdm(total=total_files, desc="Converting", unit="file") as pbar:
        for idx, filename in enumerate(all_docs):
            tqdm.write(f"🔄 Processing: {filename} ({idx + 1}/{total_files})")
            result = convert_doc_to_docx(word, filename, idx, total_files)
            pbar.update(1)

except KeyboardInterrupt:
    print("\n⛔ Interrupted by user.")

finally:
    if "word" in locals():
        word.Quit()

    if failed_files:
        with open(FAILED_LOG_PATH, "w", encoding="utf-8") as f:
            json.dump(failed_files, f, ensure_ascii=False, indent=2)
        print(f"\n❗ Failed files written to: {FAILED_LOG_PATH}")
    else:
        print("\n✅ All files converted successfully.")
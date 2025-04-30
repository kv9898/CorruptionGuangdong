import os
import json
from pathlib import Path
from docx import Document
from multiprocessing import Pool, cpu_count
from itertools import islice
from tqdm import tqdm

VERDICTS_OUTPUT_IR = Path(os.getcwd()) / "verdicts" / "docx"
KEYWORDS = ["市委书记", "市长"]
MATCHING_LOG_PATH = Path(os.getcwd()) / "matched_docx_files.json"
BATCH_SIZE = 50  # Number of files per batch

def contains_keywords(docx_path: Path) -> tuple[str, str] | None:
    try:
        doc = Document(docx_path)
        text = "\n".join(p.text for p in doc.paragraphs)
        if any(keyword in text for keyword in KEYWORDS):
            return ("match", str(docx_path.name))
    except Exception as e:
        return ("error", f"{docx_path.name}: {e}")
    return None

def batched(iterable, size):
    it = iter(iterable)
    while True:
        batch = list(islice(it, size))
        if not batch:
            break
        yield batch

if __name__ == "__main__":
    all_docx_files = [VERDICTS_OUTPUT_IR / f for f in os.listdir(VERDICTS_OUTPUT_IR) if f.lower().endswith(".docx")]

    num_processes = max(1, cpu_count() * 2 // 3)
    matching_files = []

    with tqdm(total=len(all_docx_files), desc="Filtering", unit="file") as pbar:
        for batch in batched(all_docx_files, BATCH_SIZE):
            with Pool(processes=num_processes) as pool:
                results = pool.map(contains_keywords, batch)

            for r in results:
                if r is None:
                    continue
                kind, content = r
                if kind == "match":
                    tqdm.write(f"✅ Match: {content}")
                    matching_files.append(content)
                elif kind == "error":
                    tqdm.write(f"❌ Failed: {content}")

            pbar.update(len(batch))

    if matching_files:
        with open(MATCHING_LOG_PATH, "w", encoding="utf-8") as f:
            json.dump(matching_files, f, ensure_ascii=False, indent=2)
        print(f"\n📄 Matching file list saved to: {MATCHING_LOG_PATH}")
    else:
        print("\n❌ No matching files found.")

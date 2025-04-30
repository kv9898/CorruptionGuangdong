import os
import json
from shutil import copy
from pathlib import Path
from tqdm import tqdm

VERDICTS_OUTPUT_DIR = Path(os.getcwd()) / "verdicts" / "docx"
VERDICTS_MATCHED_DIR = Path(os.getcwd()) / "verdicts" / "matched"
MATCHING_LOG_PATH = Path(os.getcwd()) / "matched_docx_files.json"

matched_files = json.load(open(MATCHING_LOG_PATH, "r", encoding="utf-8"))

VERDICTS_MATCHED_DIR.mkdir(parents=True, exist_ok=True)
with tqdm(total=len(matched_files), desc="Copying", unit="file") as pbar:
    for matched_file in matched_files:
        matched_file_path = VERDICTS_OUTPUT_DIR / matched_file
        if matched_file_path.exists():
            copy(matched_file_path, VERDICTS_MATCHED_DIR / matched_file)
        else:
            tqdm.write(f"❌ File not found: {matched_file_path}")
        pbar.update(1)
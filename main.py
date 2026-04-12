"""
main.py
Fixes color contrast accessibility issues in a single Word document.
Usage: python main.py <path_to_docx>
"""

import sys
import json
import logging
from pathlib import Path
from datetime import datetime

from word_reader import extract_issues, apply_fixes
from contrast_checker import contrast_ratio, find_accessible_color

LOG_DIR = Path("logs")
LOG_DIR.mkdir(exist_ok=True)
log_file = LOG_DIR / f"accessibility_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(log_file, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)


def process_document(doc_path: str) -> None:
    doc_path = str(Path(doc_path).resolve())
    logging.info("Processing: %s", doc_path)

    issues = extract_issues(doc_path)
    logging.info("Found %d contrast issue(s).", len(issues))

    if not issues:
        logging.info("No contrast issues found. Document is accessible.")
        return

    fixes = []
    change_log = []

    for issue in issues:
        fg = issue["fg_rgb"]
        bg = issue["bg_rgb"]

        logging.info(
            "Issue — text: '%s' | fg: %s | bg: %s | ratio: %.2f:1",
            issue["text"], fg, bg, issue["contrast"],
        )

        new_rgb = find_accessible_color(fg, bg)
        new_ratio = contrast_ratio(new_rgb, bg)
        logging.info("Fix: %s → %s (new ratio: %.2f:1)", fg, new_rgb, new_ratio)

        change_log.append({
            "text": issue["text"],
            "original_color": fg,
            "fixed_color": new_rgb,
            "original_ratio": issue["contrast"],
            "new_ratio": round(new_ratio, 2),
            "paragraph": issue["paragraph_index"],
        })

        fixes.append({
            "paragraph_index": issue["paragraph_index"],
            "run_index": issue["run_index"],
            "new_rgb": new_rgb,
        })

    change_log_path = LOG_DIR / f"changes_{Path(doc_path).stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    with open(change_log_path, "w", encoding="utf-8") as f:
        json.dump(change_log, f, indent=2)
    logging.info("Change log saved: %s", change_log_path)

    output_path = str(Path(doc_path).with_stem(Path(doc_path).stem + "_fixed"))
    apply_fixes(doc_path, output_path, fixes)
    logging.info("Fixed document saved: %s", output_path)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python main.py <path_to_docx>")
        sys.exit(1)
    process_document(sys.argv[1])

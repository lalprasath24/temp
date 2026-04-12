"""
word_reader.py
Reads and writes Microsoft Word documents using win32com.client (pywin32).
Opens Word ONCE — scans for issues and applies fixes in the same session.
"""

import logging
import win32com.client as win32
from contrast_checker import hex_to_rgb, rgb_to_win32, passes_wcag_aa, contrast_ratio

WD_COLOR_AUTOMATIC = -16777216  # wdColorAutomatic
WD_UNDEFINED = 9999999          # wdUndefined
WHITE_RGB = (255, 255, 255)


def _get_bg_color(paragraph) -> tuple[int, int, int]:
    try:
        bg = paragraph.Shading.BackgroundPatternColor
        if bg not in (WD_COLOR_AUTOMATIC, WD_UNDEFINED):
            return hex_to_rgb(bg)
    except Exception:
        pass
    return WHITE_RGB


def _normalize_font_color(color_val: int) -> tuple[int, int, int]:
    if color_val in (WD_COLOR_AUTOMATIC, WD_UNDEFINED):
        return (0, 0, 0)
    return hex_to_rgb(color_val)


def extract_and_fix(doc_path: str, output_path: str, fix_map: dict) -> None:
    """
    Single-session: open doc, apply fix_map colors, save to output_path.
    fix_map: {(paragraph_index, run_index): new_rgb}
    """
    word = win32.Dispatch("Word.Application")
    word.Visible = False

    try:
        doc = word.Documents.Open(doc_path)
        paragraphs = doc.Paragraphs

        # Group fixes by paragraph to avoid re-fetching the same paragraph repeatedly
        by_para: dict[int, list[tuple[int, tuple[int, int, int]]]] = {}
        for (p_idx, r_idx), new_rgb in fix_map.items():
            by_para.setdefault(p_idx, []).append((r_idx, new_rgb))

        for p_idx, word_fixes in by_para.items():
            try:
                para = paragraphs(p_idx)
                words = para.Range.Words
                for r_idx, new_rgb in word_fixes:
                    try:
                        words(r_idx).Font.Color = rgb_to_win32(new_rgb)
                    except Exception as e:
                        logging.warning("Fix failed at para %d word %d: %s", p_idx, r_idx, e)
            except Exception as e:
                logging.warning("Could not access paragraph %d: %s", p_idx, e)

        doc.SaveAs(output_path)
        doc.Close(False)
        logging.info("Saved fixed document: %s", output_path)
    finally:
        word.Quit()


def extract_issues(doc_path: str) -> list[dict]:
    """
    Open document, scan all words for WCAG contrast failures.
    Returns list of issue dicts.
    """
    word = win32.Dispatch("Word.Application")
    word.Visible = False
    issues = []

    try:
        doc = word.Documents.Open(doc_path)
        paragraphs = doc.Paragraphs
        para_count = paragraphs.Count

        for p_idx in range(1, para_count + 1):
            para = paragraphs(p_idx)
            bg_rgb = _get_bg_color(para)
            words = para.Range.Words
            word_count = words.Count

            for r_idx in range(1, word_count + 1):
                try:
                    w = words(r_idx)
                    fg_rgb = _normalize_font_color(w.Font.Color)

                    if not passes_wcag_aa(fg_rgb, bg_rgb):
                        text = w.Text.strip()
                        if text:
                            issues.append({
                                "paragraph_index": p_idx,
                                "run_index": r_idx,
                                "text": text[:60],
                                "fg_rgb": fg_rgb,
                                "bg_rgb": bg_rgb,
                                "contrast": round(contrast_ratio(fg_rgb, bg_rgb), 2),
                            })
                except Exception as e:
                    logging.debug("Skip word %d para %d: %s", r_idx, p_idx, e)

        doc.Close(False)
    finally:
        word.Quit()

    return issues


def apply_fixes(doc_path: str, output_path: str, fixes: list[dict]) -> None:
    """Build fix_map and apply in a single Word session."""
    fix_map = {(f["paragraph_index"], f["run_index"]): f["new_rgb"] for f in fixes}
    extract_and_fix(doc_path, output_path, fix_map)

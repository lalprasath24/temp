import os
import re
import win32com.client as win32

# Word constants
WD_LIST_BULLET = 1
WD_LIST_SIMPLE_NUMBERING = 2
WD_LIST_NO_NUMBERING = 0

BULLET_PATTERN = re.compile(r'^[\-\*\•\·\>\~]\s*')
NUMBER_PATTERN = re.compile(r'^(\d+[\.\)]\s*|[a-zA-Z][\.\)]\s+)')


def is_fake_list(para):
    """
    Returns (is_fake: bool, list_type: str)
    Fake list = visually looks like a list but ListType == 0 (no real Word list format).
    """
    text = para.Range.Text.strip()
    if not text or 'Heading' in para.Style.NameLocal:
        return False, ''
    if para.Range.ListFormat.ListType != WD_LIST_NO_NUMBERING:
        return False, ''
    if BULLET_PATTERN.match(text):
        return True, 'bullet'
    if NUMBER_PATTERN.match(text):
        return True, 'number'
    return False, ''


def strip_prefix_by_range(para):
    """
    Delete only the leading fake-list prefix characters using a character-level
    Range delete — this preserves all run formatting (font, bold, color, etc.)
    and does NOT overwrite the paragraph text.
    """
    text = para.Range.Text.strip()
    match = BULLET_PATTERN.match(text) or NUMBER_PATTERN.match(text)
    if not match:
        return

    prefix_len = match.end()

    # Build a range covering only the prefix characters
    rng = para.Range
    rng.Start = rng.Start          # paragraph start
    rng.End = rng.Start + prefix_len
    rng.Delete()                   # delete only the prefix, nothing else


def apply_list_format(para, list_type: str):
    """
    Apply proper Word list format WITHOUT changing any visual content.
    Steps:
      1. Delete the manual prefix via character-range delete (preserves formatting)
      2. Apply the Word list template
    """
    # Snapshot text before for validation
    before = para.Range.Text.strip()

    strip_prefix_by_range(para)

    lf = para.Range.ListFormat
    if list_type == 'bullet':
        lf.ApplyListTemplate(
            para.Application.ListGalleries(WD_LIST_BULLET).ListTemplates(1)
        )
    elif list_type == 'number':
        lf.ApplyListTemplate(
            para.Application.ListGalleries(WD_LIST_SIMPLE_NUMBERING).ListTemplates(1)
        )

    # Validate: content after prefix removal should be a suffix of original text
    after = para.Range.Text.strip()
    if after and after not in before:
        print(f"  [WARN] Content mismatch after fix!\n    Before: {before!r}\n    After : {after!r}")


def process_document(doc_path: str, output_path: str):
    word = win32.Dispatch('Word.Application')
    word.Visible = False

    doc = word.Documents.Open(os.path.abspath(doc_path))
    fixed = 0

    for i, para in enumerate(doc.Paragraphs):
        is_fake, list_type = is_fake_list(para)
        if is_fake:
            fixed += 1
            print(f"[Fake #{fixed}] Para {i+1} ({list_type}): {para.Range.Text.strip()[:80]!r}")
            apply_list_format(para, list_type)

    print(f"\nFixed {fixed} fake list item(s)." if fixed else "No fake lists detected.")

    doc.SaveAs(os.path.abspath(output_path))
    doc.Close()
    word.Quit()
    print(f"Saved: {os.path.abspath(output_path)}")


if __name__ == '__main__':
    input_file = 'accessibility_issues_test.docx'
    output_file = 'output_fixed.docx'

    if not os.path.exists(input_file):
        print(f"File not found: {input_file}")
    else:
        process_document(input_file, output_file)

"""
contrast_checker.py
Calculates WCAG 2.1 contrast ratios and determines accessibility compliance.
All heavy functions are cached — only 256 possible channel values exist,
and color pairs repeat heavily in real documents.
"""

from functools import lru_cache


@lru_cache(maxsize=256)
def _channel_luminance(c: int) -> float:
    """Linearize a single 0-255 channel value. Cached for all 256 values."""
    s = c / 255.0
    return s / 12.92 if s <= 0.04045 else ((s + 0.055) / 1.055) ** 2.4


@lru_cache(maxsize=1024)
def relative_luminance(rgb: tuple[int, int, int]) -> float:
    """Calculate relative luminance per WCAG 2.1 spec."""
    r, g, b = rgb
    return 0.2126 * _channel_luminance(r) + 0.7152 * _channel_luminance(g) + 0.0722 * _channel_luminance(b)


@lru_cache(maxsize=1024)
def contrast_ratio(fg: tuple[int, int, int], bg: tuple[int, int, int]) -> float:
    """Return WCAG contrast ratio between foreground and background colors."""
    l1 = relative_luminance(fg)
    l2 = relative_luminance(bg)
    lighter, darker = max(l1, l2), min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


def passes_wcag_aa(fg: tuple[int, int, int], bg: tuple[int, int, int], large_text: bool = False) -> bool:
    """Check if color pair meets WCAG AA standard (4.5:1 normal, 3:1 large text)."""
    threshold = 3.0 if large_text else 4.5
    return contrast_ratio(fg, bg) >= threshold


def hex_to_rgb(hex_color: int) -> tuple[int, int, int]:
    """Convert a Win32 BGR integer color to RGB tuple."""
    return (hex_color & 0xFF, (hex_color >> 8) & 0xFF, (hex_color >> 16) & 0xFF)


def rgb_to_win32(rgb: tuple[int, int, int]) -> int:
    """Convert RGB tuple back to Win32 BGR integer."""
    r, g, b = rgb
    return b << 16 | g << 8 | r


@lru_cache(maxsize=1024)
def find_accessible_color(
    fg: tuple[int, int, int],
    bg: tuple[int, int, int],
    threshold: float = 4.5,
) -> tuple[int, int, int]:
    """
    Binary search: blend fg toward black (light bg) or white (dark bg)
    until contrast threshold is met. Cached — same color pair computed once.
    """
    if contrast_ratio(fg, bg) >= threshold:
        return fg

    target = (0, 0, 0) if relative_luminance(bg) > 0.5 else (255, 255, 255)

    lo, hi = 0, 255
    best = target
    while lo <= hi:
        mid = (lo + hi) // 2
        t = mid / 255.0
        candidate = (
            int(fg[0] + (target[0] - fg[0]) * t),
            int(fg[1] + (target[1] - fg[1]) * t),
            int(fg[2] + (target[2] - fg[2]) * t),
        )
        if contrast_ratio(candidate, bg) >= threshold:
            best = candidate
            hi = mid - 1
        else:
            lo = mid + 1

    return best

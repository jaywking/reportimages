"""Helpers for building hyperlinks from a base URL."""
from __future__ import annotations

from urllib.parse import urljoin, urlparse


def build_hyperlink(base_url: str, filename: str) -> str | None:
    """Create a hyperlink by combining *base_url* and *filename*.

    Returns ``None`` when the base URL is missing or invalid. Trailing slashes
    on the base URL are normalized to ensure exactly one separator.
    """

    base = (base_url or "").strip()
    if not base:
        return None

    parsed = urlparse(base)
    if not parsed.scheme or not parsed.netloc:
        return None

    normalized_base = base.rstrip("/\\") + "/"
    return urljoin(normalized_base, filename)

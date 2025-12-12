"""Helpers for building hyperlinks from a base URL."""
from __future__ import annotations

from urllib.parse import quote, urlparse, urlunparse


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

    # Normalize path separators and percent-encode spaces/special chars in both
    # the base path and filename to keep SharePoint/OneDrive links valid.
    normalized_path = parsed.path.replace("\\", "/").rstrip("/")
    encoded_path = quote(normalized_path, safe="/%")
    encoded_filename = quote(filename)

    normalized_base = urlunparse(
        (
            parsed.scheme,
            parsed.netloc,
            encoded_path,
            "",  # params
            "",  # query
            "",  # fragment
        )
    ).rstrip("/") + "/"

    return normalized_base + encoded_filename

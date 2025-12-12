"""Utility functions for loading and resizing images in memory."""
from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Iterable, List

from PIL import Image


IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff"}


@dataclass(frozen=True)
class ImageEntry:
    """Represents an image file on disk and its associated hyperlink."""

    path: Path
    hyperlink: str


def discover_images(folder: Path) -> List[Path]:
    """Return a sorted list of image paths in *folder*.

    Only files with known image extensions are returned. The function does not
    recurse into subdirectories to keep the UI simple and predictable.
    """

    images = [
        path
        for path in folder.iterdir()
        if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS
    ]
    images.sort(key=lambda p: p.name.lower())
    return images


def resize_image_to_width(path: Path, target_width_px: int) -> BytesIO:
    """Resize *path* to *target_width_px* preserving aspect ratio.

    The resized image is stored in-memory and returned as a BytesIO stream.
    Original files are never modified and no temporary files are created.
    """

    with Image.open(path) as img:
        width, height = img.size
        if width <= target_width_px:
            # Do not upscale; return original bytes to preserve quality.
            output = BytesIO()
            img.save(output, format=img.format or "PNG")
            output.seek(0)
            return output

        ratio = target_width_px / float(width)
        target_height = int(height * ratio)
        resized = img.resize((target_width_px, target_height), Image.Resampling.LANCZOS)
        output = BytesIO()
        resized.save(output, format=img.format or "PNG")
        output.seek(0)
        return output

# Image → Word Link Builder

Small Tkinter utility for Windows that builds a Word document containing clickable images sourced from a OneDrive/SharePoint-synced folder.

## Features
- Browse a folder and choose one, some, or all discovered images (no subfolders).
- Resize images in-memory only (Lanczos resampling) to 2.0", 2.25" (default), or 2.5" widths.
- Each embedded image is wrapped in a hyperlink formed from the provided base OneDrive/SharePoint URL plus the filename.
- Word compression disabled and DPI set to high fidelity for maximum quality.

## Usage notes
- Enter the Base OneDrive/SharePoint URL (e.g., `https://company.sharepoint.com/sites/project/Shared%20Documents/Photos/`).
- Supported file types: `.jpg`, `.jpeg`, `.png`, `.heic` (others are silently ignored).
- Images are listed in the order reported by the operating system; no auto-sorting is applied.

## Running
```bash
python app.py
```

Select a synced folder, tick the images you want, pick a width, and click **Create Word Document** to choose where to save the new `.docx` file.

# Image → Word Link Builder

Small Tkinter utility for Windows that builds a Word document containing clickable images sourced from a OneDrive/SharePoint-synced folder.

## Features
- Browse a folder and choose one, some, or all discovered images (no subfolders).
- Resize images in-memory only (Lanczos resampling) to 2.0", 2.25" (default), or 2.5" widths.
- Each embedded image is wrapped in a hyperlink pointing at the original asset (local file URI or mapped URL).
- Word compression disabled and DPI set to high fidelity for maximum quality.

## Optional URL mapping
If the chosen folder contains a CSV named `links.csv` (or any `*links*.csv`), the app will use filename→URL pairs from that file. The CSV must have `filename` and `url` columns.

## Running
```bash
python app.py
```

Select a synced folder, tick the images you want, pick a width, and click **Create Word Document** to choose where to save the new `.docx` file.

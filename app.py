"""Image → Word Link Builder GUI application."""
from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

from doc_builder import IMAGE_WIDTH_INCHES_TO_PIXELS, build_document
from image_utils import ImageEntry, discover_images
from link_mapping import load_link_mapping, resolve_hyperlink


class ImageWordLinkBuilderApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Image → Word Link Builder")
        self.geometry("720x600")

        self.folder_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Pick a folder to begin.")
        self.selected_count_var = tk.IntVar(value=0)
        self.width_var = tk.DoubleVar(value=2.25)

        self.image_vars: dict[Path, tk.BooleanVar] = {}
        self.images: list[Path] = []
        self.link_mapping: dict[str, str] = {}

        self._build_layout()

    def _build_layout(self) -> None:
        padding = {"padx": 10, "pady": 5}

        # Folder picker row
        folder_frame = tk.Frame(self)
        folder_frame.pack(fill="x", **padding)
        tk.Label(folder_frame, text="Folder:").pack(side="left")
        folder_entry = tk.Entry(folder_frame, textvariable=self.folder_var, state="readonly")
        folder_entry.pack(side="left", fill="x", expand=True, padx=(5, 5))
        tk.Button(folder_frame, text="Browse", command=self._select_folder).pack(side="right")

        # Image list with scrollbar
        list_frame = tk.LabelFrame(self, text="Images")
        list_frame.pack(fill="both", expand=True, **padding)

        canvas = tk.Canvas(list_frame)
        scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        self.images_container = tk.Frame(canvas)

        self.images_container.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.images_container, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Width selection
        width_frame = tk.LabelFrame(self, text="Image width")
        width_frame.pack(fill="x", **padding)
        for width_inch in IMAGE_WIDTH_INCHES_TO_PIXELS:
            tk.Radiobutton(
                width_frame,
                text=f"{width_inch:.2f}\"",
                value=width_inch,
                variable=self.width_var,
            ).pack(side="left", padx=5)

        # Status and action row
        status_frame = tk.Frame(self)
        status_frame.pack(fill="x", **padding)
        self.selection_label = tk.Label(status_frame, text="0 images selected")
        self.selection_label.pack(side="left")

        tk.Button(status_frame, text="Select All", command=self._select_all).pack(side="left", padx=5)
        tk.Button(status_frame, text="Clear", command=self._clear_selection).pack(side="left")

        tk.Button(status_frame, text="Create Word Document", command=self._create_document).pack(
            side="right"
        )

        tk.Label(self, textvariable=self.status_var, anchor="w", fg="#0a5c0a").pack(
            fill="x", padx=10, pady=(0, 10)
        )

    def _select_folder(self) -> None:
        folder = filedialog.askdirectory()
        if not folder:
            return

        folder_path = Path(folder)
        self.folder_var.set(str(folder_path))
        self.status_var.set("Loading images…")
        self.update_idletasks()

        self.images = discover_images(folder_path)
        self.link_mapping, mapping_status = load_link_mapping(folder_path)

        self._populate_image_list()
        self.status_var.set(mapping_status)

    def _populate_image_list(self) -> None:
        for child in list(self.images_container.children.values()):
            child.destroy()
        self.image_vars.clear()

        for path in self.images:
            var = tk.BooleanVar(value=False)
            self.image_vars[path] = var
            cb = tk.Checkbutton(
                self.images_container,
                text=path.name,
                variable=var,
                anchor="w",
                justify="left",
                command=self._update_selection_count,
            )
            cb.pack(fill="x", anchor="w")

        self._update_selection_count()

    def _update_selection_count(self) -> None:
        selected = sum(1 for var in self.image_vars.values() if var.get())
        self.selected_count_var.set(selected)
        self.selection_label.config(text=f"{selected} images selected")

    def _select_all(self) -> None:
        for var in self.image_vars.values():
            var.set(True)
        self._update_selection_count()

    def _clear_selection(self) -> None:
        for var in self.image_vars.values():
            var.set(False)
        self._update_selection_count()

    def _create_document(self) -> None:
        if not self.images:
            messagebox.showwarning("No folder", "Please pick a folder with images first.")
            return

        selected_entries = [
            ImageEntry(path=path, hyperlink=resolve_hyperlink(path, self.link_mapping))
            for path, var in self.image_vars.items()
            if var.get()
        ]

        if not selected_entries:
            messagebox.showwarning("No images selected", "Select at least one image to continue.")
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")],
            title="Save Word document",
        )
        if not output_path:
            return

        try:
            build_document(selected_entries, self.width_var.get(), Path(output_path))
        except Exception as exc:  # noqa: BLE001 - surfaced directly to user
            messagebox.showerror("Error creating document", str(exc))
            self.status_var.set("Document creation failed.")
            return

        self.status_var.set(f"Saved Word document to {output_path}")
        messagebox.showinfo("Success", "Word document created successfully.")


def main() -> None:
    app = ImageWordLinkBuilderApp()
    app.mainloop()


if __name__ == "__main__":
    main()

"""
User interface functions
"""

import tkinter as tk
from pathlib import Path
from tkinter import filedialog


def select_pdf_files() -> tuple[str, ...]:
    """Open a file dialog to let the user select PDF files.

    :return: Tuple of selected file paths (empty if cancelled)
    :rtype: tuple[str, ...]
    """
    root = tk.Tk()
    root.withdraw()

    file_paths = filedialog.askopenfilenames(
        parent=root,
        title="Select Bills to Analyze",
        filetypes=[("PDF files", "*.pdf")],
        initialdir=str(Path.home() / "Downloads"),
    )

    # askopenfilenames returns empty string if cancelled, otherwise tuple
    if isinstance(file_paths, str):
        return ()
    return file_paths

"""
Batch XRM to TIFF/BMP Converter.

This script provides a graphical user interface (GUI) for batch converting
XRM (ZEISS proprietary microscopy data format) files to TIFF or BMP image formats.
It recursively searches a user-specified directory for XRM files, extracts
image data, and saves the converted images in the same directory as the
source XRM files.

Features:
    - User-friendly GUI for input directory selection.
    - Option to choose between TIFF and BMP output formats.
    - Progress bar to track conversion progress.
    - Error handling for missing or invalid XRM data.
    - Rescaling of image intensity using percentile-based clipping.

Dependencies:
    - tkinter: For GUI creation.
    - pathlib: For file path manipulation.
    - numpy: For numerical operations and array manipulation.
    - olefile: For reading XRM file structure.
    - struct: For unpacking binary data.
    - skimage: For image rescaling and saving.
"""

# Author: Daniel Bribiesca Sykes <daniel.bribiescasykes@glasgow.ac.uk>
# Version: 3.0.0

import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from pathlib import Path
import numpy as np
import olefile as olef
import struct
from skimage import io, exposure


def choose_directory(entry):
    """
    Open a directory selection dialog and update the input entry with the selected directory.

    Parameters
    ----------
        entry (tk.Entry): The tkinter Entry widget to update with the selected directory path.
    """
    directory = filedialog.askdirectory(title="Select XRM Directory")
    if directory:
        entry.delete(0, tk.END)
        entry.insert(0, directory)


def ole_extract(ole, stream, datatype):
    """
    Extract data from an OLE file stream.

    Parameters
    ----------
        ole (olefile.OleFileIO): The OLE file object.
        stream (str): The name of the stream to extract data from.
        datatype (str): The struct format string specifying the data type to unpack.

    Returns
    -------
        tuple: The unpacked data as a tuple, or None if the stream does not exist.
    """
    if ole.exists(stream):
        data = ole.openstream(stream).read()
        return struct.unpack(datatype, data)
    return None


def process_xrm(file_path, output_format, progress_var, progress_bar, total_files):
    """
    Process a single XRM file, extract image data, and save it as a TIFF or BMP file.

    Parameters
    ----------
        file_path (Path): The path to the XRM file.
        output_format (str): The desired output format ('tiff' or 'bmp').
        progress_var (tk.IntVar): The tkinter IntVar for tracking progress.
        progress_bar (ttk.Progressbar): The tkinter Progressbar widget for displaying progress.
        total_files (int): The total number of XRM files to process.

    Returns
    -------
        bool: True if the file was processed successfully, False otherwise.
    """
    try:
        with olef.OleFileIO(file_path) as ole:
            n_cols = ole_extract(ole, "ImageInfo/ImageWidth", "<I")[0]
            n_rows = ole_extract(ole, "ImageInfo/ImageHeight", "<I")[0]
            imgdata = ole_extract(ole, 'ImageData1/Image1', "<{}h".format(n_cols*n_rows))

            if n_cols is None or n_rows is None or imgdata is None:
                raise ValueError(f"Missing data in {file_path.name}")

            try:
                absdata = np.reshape(imgdata, (n_cols, n_rows), order="F").astype(np.uint16)
            except ValueError as e:
                raise ValueError(
                    f"Reshape error in {file_path.name}: {e}. Extracted data size does not match image dimensions."
                )

            vmin, vmax = np.percentile(absdata, (0.1, 99.9))
            rescale_img = exposure.rescale_intensity(absdata, in_range=(vmin, vmax), out_range=np.uint16)

            output_file = file_path.parent / f"{file_path.stem}.{output_format}"
            io.imsave(str(output_file), rescale_img)

            return True

    except Exception as e:
        messagebox.showerror("Error", f"Error processing {file_path.name}: {e}")
        print(f"Error processing {file_path.name}: {e}")
        return False

    finally:
        progress_var.set(progress_var.get() + 1)
        progress_bar["value"] = progress_var.get() / total_files * 100


def batch_xrm_convert(root, input_dir, output_format, progress_var, progress_bar):
    """
    Convert all XRM files in the specified directory to the chosen output format.

    Parameters
    ----------
        root (tk.Tk): The main tkinter window.
        input_dir (Path): The directory containing the XRM files.
        output_format (str): The desired output format ('tiff' or 'bmp').
        progress_var (tk.IntVar): The tkinter IntVar for tracking progress.
        progress_bar (ttk.Progressbar): The tkinter Progressbar widget for displaying progress.
    """
    files = list(input_dir.rglob("*.xrm"))
    total_files = len(files)
    progress_bar["maximum"] = total_files
    progress_var.set(0)

    for file_path in files:
        process_xrm(
            file_path, output_format, progress_var, progress_bar, total_files
        )

    root.after(10, lambda: [messagebox.showinfo("Conversion Complete", "XRM conversion finished."), root.destroy()])


def main():
    """Set up and run the main GUI for the XRM to TIFF/BMP converter."""
    root = tk.Tk()
    root.title("XRM to TIFF/BMP Converter")

    input_dir_label = tk.Label(root, text="Input Directory:")
    input_dir_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)

    input_dir_entry = tk.Entry(root, width=50)
    input_dir_entry.grid(row=0, column=1, padx=5, pady=5)

    input_dir_button = tk.Button(
        root, text="Browse", command=lambda: choose_directory(input_dir_entry)
    )
    input_dir_button.grid(row=0, column=2, padx=5, pady=5)

    output_format_label = tk.Label(root, text="Output Format:")
    output_format_label.grid(row=1, column=0, sticky="w", padx=5, pady=5)

    output_format_var = tk.StringVar(root)
    output_format_var.set("tiff")
    output_format_dropdown = tk.OptionMenu(root, output_format_var, "tiff", "bmp")
    output_format_dropdown.grid(row=1, column=1, padx=5, pady=5)

    progress_var = tk.IntVar()
    progress_bar = ttk.Progressbar(
        root,
        variable=progress_var,
        maximum=100,
        orient="horizontal",
        length=300,
        mode="determinate",
    )
    progress_bar.grid(row=2, column=0, columnspan=3, padx=5, pady=5)

    convert_button = tk.Button(
        root,
        text="Convert XRM Files",
        command=lambda: batch_xrm_convert(
            root,
            Path(input_dir_entry.get()),
            output_format_var.get(),
            progress_var,
            progress_bar,
        ),
    )
    convert_button.grid(row=3, column=1, pady=10)

    root.mainloop()


if __name__ == "__main__":
    main()

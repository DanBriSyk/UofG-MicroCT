# Batch TXM to TIFF Converter

This Python application provides a graphical user interface (GUI) for batch converting `.txm` files into TIFF image stacks or 3D TIFFs. It recursively searches a selected folder for `.txm` files and processes them in parallel using Dask for improved performance.

## Features

-   **Recursive File Search:** Automatically locates and converts all `.txm` files within a specified directory and its subdirectories.
-   **Output Format Selection:** Supports conversion to TIFF stacks or 3D TIFF files.
-   **Optional ZIP Compression:** Compresses each output image stack into a separate ZIP archive.
-   **Slice Preview:** Allows users to preview a slice of each scan before processing.
-   **8-bit Conversion:** Option to convert 16-bit images to 8-bit for compatibility.
-   **Parallel/Serial Processing:** Option to convert .txm files with Dask for parallel processing (significantly reducing conversion time for large datasets), or serially (when combined file size > available RAM)
-   **Progress Tracking:** Provides real-time progress updates via a progress bar and text output.
-   **Logging:** Logs all conversion activities and errors to a log file (`txm_converter.log`).
-   **GUI Interface:** User-friendly interface built with PyQt5 for easy operation.

## Installation

1.  **Clone the Repository:**

    ```bash
    git clone https://github.com/<<<your-github-account>>>/UofG-MicroCT.git
    cd UofG-MicroCT
    ```

2.  **Install Dependencies:**

    ```bash
    pip install PyQt5 numpy opencv-python olefile tifffile imageio dask distributed psutil
    ```

## Usage

1.  **Run the application:**

    ```bash
    python Batch_TXM_converter_UofG.py
    ```

2.  **Select Folder:**
    -   Click the "Open" button to choose the directory containing the `.txm` files.
3.  **Choose Output Format:**
    -   Select "Tiff stack" for individual TIFF files or "3D Tif" for a single 3D TIFF file.
4.  **Configure Options:**
    -   Check "Zip each image stack" to compress the output files.
    -   Check "Check each scan before processing" to preview slices.
    -   Check "Convert to 8-bit" to reduce the bit depth of the output images.
    -   Select "Parallel" or "Serial" in the processing mode menu.
5.  **Start Conversion:**
    -   Click "OK" to begin the conversion process.
6.  **Monitor Progress:**
    -   The progress bar and text output will display the conversion status.
7.  **Cancel Conversion:**
    -   Click "Cancel" to stop the conversion process.

## Code Description

-   **`ParallelWorkerThread`:** A QThread subclass that handles the parallel conversion process in a separate thread to prevent GUI freezing.
-   **`SerialWorkerThread`:** A QThread subclass that handles the serial conversion process in a separate thread to prevent GUI freezing.
-   **`convert_scans`:** Orchestrates the parallel conversion of `.txm` files using Dask.
-   **`process_txm`:** Handles the conversion of a single `.txm` file, including metadata extraction, slice loading, and saving.
-   **`_extract_metadata`, `_get_sorted_image_streams`, `_create_output_folder`, `_load_slices`, `_convert_to_8bit`, `_save_slices`, `_zip_output`:** Helper functions for various conversion tasks.
-   **`extract_number`, `ole_extract`, `close_dask_client`, `display_slice`:** Utility functions for file processing and display.
-   **`Window`:** A QDialog subclass that creates the GUI for user interaction.
-   **`Set_Batch`:** The main function that initializes and runs the GUI application.

## Dependencies

-   **PyQt5:** For the GUI.
-   **NumPy:** For numerical operations on image data.
-   **OpenCV (cv2):** For image display.
-   **olefile:** For reading `.txm` files.
-   **tifffile:** For writing TIFF files.
-   **imageio:** For writing tiff stacks.
-   **Dask:** For parallel processing.
-   **psutil:** For retrieving available memory.

## Logging

The application logs all activities and errors to `txm_converter.log`. This file can be used for troubleshooting and monitoring the conversion process.

## Author

Daniel Bribiesca Sykes (<daniel.bribiescasykes@glasgow.ac.uk>)

## Version

1.3.7

# XRM to TIFF/BMP Converter

This Python script provides a graphical user interface (GUI) for batch converting XRM files (a proprietary microscopy data format) to TIFF or BMP image formats. It recursively searches a user-specified directory for XRM files, extracts image data, and saves the converted images in the same directory as the source XRM files.

## Features

-   **User-friendly GUI:** Simple and intuitive interface for selecting input directory and output format.
-   **Batch Conversion:** Processes all XRM files within a directory and its subdirectories.
-   **Output Format Selection:** Allows users to choose between TIFF and BMP output formats.
-   **Progress Tracking:** Displays a progress bar to monitor the conversion process.
-   **Error Handling:** Robust error handling for missing or invalid XRM data.
-   **Image Rescaling:** Rescales image intensity using percentile-based clipping for improved visualization.

## Dependencies

-   Python 3.x
-   `tkinter`: For GUI creation (standard library).
-   `pathlib`: For file path manipulation (standard library).
-   `numpy`: For numerical operations and array manipulation. Install using `pip install numpy`.
-   `olefile`: For reading XRM file structure. Install using `pip install olefile`.
-   `scikit-image (skimage)`: For image processing (rescaling, saving). Install using `pip install scikit-image`.

## Installation

1.  **Clone the repository (or download the script):**

    ```bash
    git clone https://github.com/DanBriSyk/UofG-MicroCT.git
    ```

    Or download the `xrm_converter.py` file directly.

2.  **Install the required dependencies (if not already installed):**

    ```bash
    pip install numpy olefile scikit-image
    ```

## Usage

1.  **Run the script:**

    ```bash
    python xrm_converter.py
    ```

2.  **Use the GUI:**

    -   Click the "Browse" button to select the directory containing your XRM files.
    -   Choose the desired output format (TIFF or BMP) from the dropdown menu.
    -   Click the "Convert XRM Files" button to start the conversion process.
    -   A progress bar will display the conversion progress.
    -   Upon completion, a message box will confirm the conversion, and the GUI will close.
    -   The converted images will be saved in the same directories as the original XRM files.

## Example

Suppose you have a directory `C:/XRM_Data` containing several XRM files. You would:

1.  Run the script.
2.  Browse to and select `C:/XRM_Data`.
3.  Choose either TIFF or BMP as the output format.
4.  Click "Convert XRM Files".

The resulting TIFF or BMP images will be saved alongside the original XRM files within the `C:/XRM_Data` directory and any of its subdirectories.

## Author

-   Daniel Bribiesca Sykes ([daniel.bribiescasykes@glasgow.ac.uk](mailto:daniel.bribiescasykes@glasgow.ac.uk))

## Version

-   3.0.0

## License

This project is licensed under the GNU GPL-3.0 License.

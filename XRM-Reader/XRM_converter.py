"""
Batch XRM to TIFF/PNG Converter.

This script provides a graphical user interface (GUI) for batch converting
XRM (ZEISS proprietary microscopy data format) files to TIFF or PNG image formats.
It recursively searches a user-specified directory for XRM files, extracts
image data, and saves the converted images in the same directory as the
source XRM files.

Features:
    - User-friendly GUI for selecting input files or directory.
    - Option to choose between TIFF and PNG output formats.
    - Progress bar to track conversion progress.
    - Error handling for missing or invalid XRM data.
    - Rescaling of image intensity using percentile-based clipping.
    - Logging of processing information and errors to 'xrm_converter.log'.

Dependencies:
    - PyQt5: For GUI creation.
    - pathlib: For file path manipulation.
    - numpy: For numerical operations and array manipulation.
    - olefile: For reading XRM file structure.
    - struct: For unpacking binary data.
    - skimage: For image rescaling and saving.
    - logging: For logging processing information and errors.
"""

# Author: Daniel Bribiesca Sykes <daniel.bribiescasykes@glasgow.ac.uk>
# Version: 3.2.0

import sys
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QPushButton,
    QVBoxLayout,
    QFileDialog,
    QLabel,
    QProgressBar,
    QComboBox,
    QMessageBox,
    QDialog,
    QRadioButton,
    QDialogButtonBox,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from pathlib import Path
import numpy as np
import olefile as olef
import struct
from skimage import io, exposure
import logging

logging.basicConfig(filename='xrm_converter.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


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


def process_xrm(file_path, output_format):
    """
    Process a single XRM file, extract image data, and save it as a TIFF or PNG file.

    Parameters
    ----------
        file_path (Path): The path to the XRM file.
        output_format (str): The desired output format ('tiff' or 'png').

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

            logging.info(f"Successfully processed {file_path.name}")
            return True

    except Exception as e:
        logging.error(f"Error processing {file_path.name}: {e}")
        return False


class XRMConverterThread(QThread):
    """Thread for batch converting XRM files."""

    progress = pyqtSignal(int)
    finished = pyqtSignal()

    def __init__(self, file_paths, output_format):
        super().__init__()
        self.file_paths = file_paths
        self.output_format = output_format

    def run(self):
        total_files = len(self.file_paths)
        for i, file_path in enumerate(self.file_paths):
            process_xrm(Path(file_path), self.output_format)
            self.progress.emit(int((i + 1) / total_files * 100))
        self.finished.emit()


class SelectionDialog(QDialog):
    """Dialog for selecting input type (files or directory)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.initUI()

    def initUI(self):
        self.radio_files = QRadioButton("Select Files", self)
        self.radio_directory = QRadioButton("Select Directory", self)
        self.radio_files.setChecked(True)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addWidget(self.radio_files)
        layout.addWidget(self.radio_directory)
        layout.addWidget(button_box)

        self.setWindowTitle("Select Input Type")

    def getSelection(self):
        """Get the user's selection (files or directory)."""
        if self.exec_() == QDialog.Accepted:
            if self.radio_files.isChecked():
                return "files"
            else:
                return "directory"
        return None


class XRMConverter(QWidget):
    """Main GUI for the XRM to TIFF/PNG converter."""

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.select_button = QPushButton('Select Files or Directory', self)
        self.select_button.clicked.connect(self.selectFilesOrDirectory)

        self.output_format_label = QLabel('Output Format:', self)
        self.output_format_combo = QComboBox(self)
        self.output_format_combo.addItems(['tiff', 'png'])

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setAlignment(Qt.AlignCenter)

        self.convert_button = QPushButton('Convert', self)
        self.convert_button.clicked.connect(self.convert)
        self.convert_button.setEnabled(False)

        self.layout = QVBoxLayout(self)
        self.layout.addWidget(self.select_button)
        self.layout.addWidget(self.output_format_label)
        self.layout.addWidget(self.output_format_combo)
        self.layout.addWidget(self.progress_bar)
        self.layout.addWidget(self.convert_button)

        self.setWindowTitle('XRM to TIFF/PNG Converter')
        self.show()

    def selectFilesOrDirectory(self):
        """Open a dialog to select files or a directory."""
        dialog = SelectionDialog(self)
        selection = dialog.getSelection()

        if selection == "files":
            file_paths, _ = QFileDialog.getOpenFileNames(self, "Select XRM Files", "", "XRM Files (*.xrm);;All Files (*)")
            if file_paths:
                self.file_paths = file_paths
                self.convert_button.setEnabled(True)
        elif selection == "directory":
            directory = QFileDialog.getExistingDirectory(self, "Select XRM Directory")
            if directory:
                self.file_paths = list(Path(directory).rglob("*.xrm"))
                self.convert_button.setEnabled(True)

    def convert(self):
        """Start the XRM to TIFF/PNG conversion process."""
        self.progress_bar.setValue(0)
        self.convert_button.setEnabled(False)
        self.select_button.setEnabled(False)

        self.thread = XRMConverterThread(self.file_paths, self.output_format_combo.currentText())
        self.thread.progress.connect(self.progress_bar.setValue)
        self.thread.finished.connect(self.conversionFinished)
        self.thread.start()

    def conversionFinished(self):
        """Handle the completion of the conversion process."""
        QMessageBox.information(self, "Conversion Complete", "XRM conversion finished.")
        self.close()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = XRMConverter()
    sys.exit(app.exec_())

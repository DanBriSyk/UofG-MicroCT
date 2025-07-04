"""
Batch TXM to TIFF converter.

Generates a GUI for users to select a folder to locate and convert .txm files into TIFF stacks,
in a recursive and automated manner.
"""

# Author: Daniel Bribiesca Sykes <daniel.bribiescasykes@glasgow.ac.uk>
# Version: 1.3.6

from pathlib import Path
import numpy as np
import cv2
import re
from PyQt5.QtCore import QThread, pyqtSignal, QCoreApplication
from PyQt5.QtWidgets import (
    QDialog,
    QLabel,
    QDialogButtonBox,
    QCheckBox,
    QFileDialog,
    QComboBox,
    QVBoxLayout,
    QHBoxLayout,
    QMessageBox,
    QProgressBar,
    QApplication,
    QTextEdit,
    )
from zipfile import ZipFile
import olefile as olef
import struct
import tifffile
import imageio.v3 as iio
from datetime import datetime
import logging
from enum import Enum
from dask.distributed import Client, LocalCluster, as_completed
import psutil
import sys

# GUI Constants
GUI_WIDTH = 500
GUI_HEIGHT = 300
GUI_X_POS = 200
GUI_Y_POS = 200

# Processing Constants
DISPLAY_SLICE_WAIT_TIME = 10000
AVAILABLE_MEMORY = psutil.virtual_memory().available


# Output Format Constants
class OutputFormat(Enum):
    """Enumerates the supported output formats."""

    TIFF_STACK = 0
    TIFF_3D = 1


OUTPUT_FORMAT_EXTS = {
    OutputFormat.TIFF_STACK: 'tiff',
    OutputFormat.TIFF_3D: 'tif'
    }

# Logging Constants
LOG_FILE_NAME = 'txm_converter.log'
LOG_FORMAT = '%(asctime)s - %(levelname)s - %(message)s'

# Configure logging
logging.basicConfig(filename=LOG_FILE_NAME, level=logging.INFO, format=LOG_FORMAT)


class ParallelWorkerThread(QThread):
    """A separate thread for running the parallel conversion process.  This prevents the GUI from freezing."""

    progress_update = pyqtSignal(int, str)
    finished = pyqtSignal()
    error = pyqtSignal(str)
    log_message = pyqtSignal(str)
    stopped = pyqtSignal()

    def __init__(
            self,
            import_folder,
            output_base_dir,
            output_format_index,
            zip_output,
            should_display_slice,
            convert_to_8bit,
            ):
        super().__init__()
        self.import_folder = import_folder
        self.output_base_dir = output_base_dir
        self.output_format_index = output_format_index
        self.zip_output = zip_output
        self.should_display_slice = should_display_slice
        self.convert_to_8bit = convert_to_8bit
        self.txm_files = list(self.import_folder.rglob('*.txm'))
        self.total_files = len(self.txm_files)
        self.client = None
        self.is_stopped = False

    def run(self):
        """Create main method for the thread. Calls the conversion process."""
        try:
            if not self.txm_files:
                raise FileNotFoundError("No .txm files found in the selected folder.")

            # Get the Dask client from the main thread
            self.client = QCoreApplication.instance().dask_client
            if self.client is None:
                raise RuntimeError("Dask client is not initialized.")

            convert_scans(
                self.txm_files,
                self.output_base_dir,
                self.output_format_index,
                self.zip_output,
                self.should_display_slice,
                self.convert_to_8bit,
                self.progress_update,
                self.client,
                self.log_message,
                self
                )
            if not self.is_stopped:
                self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))
            logging.error(f"Error in WorkerThread: {e}")
            self.log_message.emit(f"Error: {e}")
        finally:
            if self.is_stopped:
                self.stopped.emit()

    def stop(self):
        """Stop the conversion process."""
        self.is_stopped = True
        if self.client:
            self.client.cancel(futures=None)
            self.client.close()


class SerialWorkerThread(QThread):
    """A separate thread for running the serial conversion process.  This prevents the GUI from freezing."""

    progress_update = pyqtSignal(int, str)
    finished = pyqtSignal()
    error = pyqtSignal(str)
    log_message = pyqtSignal(str)
    stopped = pyqtSignal()

    def __init__(
            self,
            import_folder,
            output_base_dir,
            output_format_index,
            zip_output,
            should_display_slice,
            convert_to_8bit
            ):
        super().__init__()
        self.import_folder = import_folder
        self.output_base_dir = output_base_dir
        self.output_format_index = output_format_index
        self.zip_output = zip_output
        self.should_display_slice = should_display_slice
        self.convert_to_8bit = convert_to_8bit
        self.is_stopped = False

    def run(self):
        """Create main method for the thread. Calls the conversion process."""
        txm_files = list(self.import_folder.rglob('*.txm'))
        startTime = datetime.now()

        for i, txm_file in enumerate(txm_files):
            if self.is_stopped:
                break
            short_output_name = f"scan_{str(i+1).zfill(2)}"
            try:
                process_txm(
                    txm_file,
                    self.output_base_dir,
                    short_output_name,
                    self.output_format_index,
                    self.zip_output,
                    self.should_display_slice,
                    self.convert_to_8bit,
                )
                self.progress_update.emit(i + 1, txm_file.name)
                self.log_message.emit(f"Processed: {txm_file.name}")
            except Exception as e:
                self.log_message.emit(f"Error processing {txm_file.name}: {e}")
                logging.error(f"Error processing {txm_file.name}: {e}")

        compTime = str((datetime.now() - startTime))[:-4]
        print(f'\nBatch conversion completed in {compTime}')
        self.log_message.emit(f'\nBatch conversion completed in {compTime}')
        logging.info(f'Batch conversion completed in {compTime}')

        if self.is_stopped:
            self.stopped.emit()
        else:
            self.finished.emit()

    def stop(self):
        """Stop the conversion process."""
        self.is_stopped = True


def convert_scans(
        txm_files,
        output_base_dir,
        output_format_index,
        zip_output,
        should_display_slice,
        convert_to_8bit,
        progress_update_signal,
        client,
        log_message_signal,
        worker_thread,
        ):
    """
    Process a list of TXM files using Dask for parallelization.

    Args
    ----
        txm_files (list): List of Path objects representing TXM files.
        output_format_index (int): Index for the output format.
        zip_output (bool): Whether to zip output.
        should_display_slice (bool): Whether to display a slice.
        convert_to_8bit (bool): Whether to convert to 8-bit.
        progress_update_signal (pyqtSignal): Signal to send progress updates.
        client (Client): The Dask client.
    """
    startTime = datetime.now()

    output_map = {txm_file: f"scan_{str(i+1).zfill(2)}" for i, txm_file in enumerate(txm_files)}

    futures = {}
    for txm_file in txm_files:
        if worker_thread.is_stopped:
            break
        # Submit the process_txm function to the Dask cluster
        future = client.submit(
            process_txm,
            txm_file,
            output_base_dir,
            output_map[txm_file],
            output_format_index,
            zip_output,
            should_display_slice,
            convert_to_8bit,
        )
        futures[future] = txm_file.stem

    log_message_signal.emit(f"Starting conversion of {len(txm_files)} scans...")

    # Process the results as they become available and update progress
    for i, future in enumerate(as_completed(futures)):
        if worker_thread.is_stopped:
            break
        try:
            future.result()
            filename = futures[future]
            progress_update_signal.emit(i + 1, filename)
            log_message_signal.emit(f"Processed: {filename}")
        except Exception as e:
            logging.error(f"Error processing file: {e}")
            filename = futures[future]
            progress_update_signal.emit(i + 1, f"Error: {filename}, Exception: {e}")
            log_message_signal.emit(f"Error processing {filename}: {e}")

    compTime = str((datetime.now() - startTime))[:-4]
    print(f'\nBatch conversion completed in {compTime}')
    log_message_signal.emit(f'\nBatch conversion completed in {compTime}')
    logging.info(f'Batch conversion completed in {compTime}')


def process_txm(txm_file, output_base_dir, short_output_name, output_format_index, zip_output, should_display_slice, convert_to_8bit):
    """
    Process a single TXM file.

    Loads image slices, converts them, and saves them.

    Args
    ----
        txm_file (Path): Path to the TXM file.
        output_format_index (int): Index for the output format.
        zip_output (bool): Whether to zip output.
        should_display_slice (bool): Whether to display a slice.
        convert_to_8bit (bool): Whether to convert to 8-bit.
    """
    print(f'Loading {txm_file.stem}', flush=True)
    logging.info(f'Loading {txm_file.stem}')

    try:
        with olef.OleFileIO(txm_file) as ole:
            n_cols, n_rows, n_images, pixel_size, f_type = _extract_metadata(ole)
            image_streams = _get_sorted_image_streams(ole)
            out_folder = _create_output_folder(txm_file, output_base_dir, short_output_name, output_format_index)

            slices = _load_slices(ole, image_streams, f_type, n_cols, n_rows)

            if convert_to_8bit:
                if f_type != 3:
                    print(f'Converting {txm_file.stem} to 8-bit', flush=True)
                    logging.info(f'Converting {txm_file.stem} to 8-bit')
                    slices_8bit = _convert_to_8bit(slices)
                    _save_slices(slices_8bit, txm_file, out_folder, pixel_size, output_format_index, should_display_slice, dtype=np.uint8)
                    del slices
                else:
                    print(f'{txm_file.stem} is already 8-bit. Skipping conversion.', flush=True)
                    logging.info(f'{txm_file.stem} is already 8-bit. Skipping conversion.')
                    _save_slices(slices, txm_file, out_folder, pixel_size, output_format_index, should_display_slice, dtype=np.uint8)
            else:
                print(f'No scaling specified for {txm_file.stem}', flush=True)
                logging.info(f'No scaling specified for {txm_file.stem}')
                _save_slices(slices, txm_file, out_folder, pixel_size, output_format_index, should_display_slice, dtype=np.uint16)

        print(f'\n{txm_file.stem} converted\n', flush=True)
        logging.info(f'{txm_file.stem} converted')

        if zip_output:
            _zip_output(txm_file, out_folder, short_output_name, output_format_index)

    except Exception as e:
        logging.error(f"Error processing {txm_file.name}: {e}")
        raise


def _extract_metadata(ole):
    """Extract metadata from the OLE file."""
    try:
        n_cols = int(ole_extract(ole, 'ImageInfo/ImageWidth', '<I')[0])
        n_rows = int(ole_extract(ole, 'ImageInfo/ImageHeight', '<I')[0])
        n_images = int(ole_extract(ole, 'ImageInfo/ImagesTaken', '<I')[0])
        pixel_size = round(ole_extract(ole, 'ImageInfo/PixelSize', '<f')[0], 2)
        f_type = int(ole_extract(ole, 'ImageInfo/DataType', '<I')[0])
        return n_cols, n_rows, n_images, pixel_size, f_type
    except Exception as e:
        raise Exception(f"Error extracting metadata: {e}")


def _get_sorted_image_streams(ole):
    """Get and sort image streams from the OLE file."""
    try:
        image_streams = [s for s in ole.listdir() if s[0].startswith('ImageData')]
        image_streams.sort(key=extract_number)
        return image_streams
    except Exception as e:
        raise Exception(f"Error getting image streams: {e}")


def _create_output_folder(txm_file, output_base_dir, short_output_name, output_format_index):
    """Create the output folder."""
    try:
        out_folder = output_base_dir / short_output_name / OUTPUT_FORMAT_EXTS[OutputFormat(output_format_index)]
        out_folder.mkdir(parents=True, exist_ok=True)
        return out_folder
    except Exception as e:
        raise Exception(f"Error creating output folder for {txm_file.name}: {e}")


def _load_slices(ole, image_streams, f_type, n_cols, n_rows):
    """Load image slices from the OLE file."""
    ftype_dic = {3: '<{}B', 5: '<{}h', 10: '<{}f'}
    try:
        slices = []
        for img_stream in image_streams:
            stream_data = ole_extract(ole, img_stream, ftype_dic[f_type].format(n_cols * n_rows))
            img_data = np.reshape(stream_data, (n_cols, n_rows), order='F')
            slices.append(img_data)
        return slices
    except Exception as e:
        raise Exception(f"Error loading slices: {e}")


def _convert_to_8bit(slices):
    """Convert slices to 8-bit."""
    try:
        arr = np.array(slices)
        global_min = np.min(arr)
        global_max = np.max(arr)
        if global_min < 0:
            arr = arr + abs(global_min)
            arr = arr * -1
            global_min = np.min(arr)
            global_max = np.max(arr)
        arr_scaled = (((arr - global_min) / (global_max - global_min)) * 255).astype(np.uint8)
        return arr_scaled
    except Exception as e:
        raise Exception(f"Error converting to 8-bit: {e}")


def _save_slices(slices, txm_file, out_folder, pixel_size, output_format_index, should_display_slice, dtype=None):
    """Save slices to the specified format (TIFF or 3D TIFF)."""
    try:
        output_format = OUTPUT_FORMAT_EXTS[OutputFormat(output_format_index)]
        if output_format_index != OutputFormat.TIFF_3D.value:
            print(f'Exporting {len(slices)} slices from {txm_file.stem} as TIFF stack', flush=True)
            logging.info(f'Exporting {len(slices)} slices from {txm_file.stem} as TIFF stack')
            for slice_index, img in enumerate(slices):
                if slice_index == round(len(slices) / 2) and should_display_slice:
                    display_slice(img.copy())
                filename = f'{out_folder}/{txm_file.stem}_{str(slice_index).zfill(4)}.{output_format}'
                metadata = {'resolution': (pixel_size, pixel_size, 'MICROMETER')}
                iio.imwrite(filename, img.astype(dtype), extension='.tiff', metadata=metadata)
        else:
            print(f'Exporting {len(slices)} slices from {txm_file.stem} as 3D TIFF', flush=True)
            logging.info(f'Exporting {len(slices)} slices from {txm_file.stem} as 3D TIFF')
            filename = f'{out_folder}/{txm_file.stem}.{output_format}'
            dpi = 25400 / pixel_size
            with tifffile.TiffWriter(filename, bigtiff=True) as tif:
                tif.write(slices, resolution=(dpi, dpi), metadata={'unit': 'um'}, dtype=dtype)
    except Exception as e:
        raise Exception(f"Error saving slices: {e}")


def _zip_output(txm_file, out_folder, output_format_index):
    """Zip the output files."""
    try:
        output_format = OUTPUT_FORMAT_EXTS[OutputFormat(output_format_index)]
        with ZipFile(f'{out_folder.parent}/{txm_file.stem}.zip', 'w') as zip_file:
            for image_file in out_folder.rglob(f'{txm_file.stem}*.{output_format}'):
                zip_file.write(image_file, image_file.name)
        print(f'{txm_file.stem} zipped\n', flush=True)
        logging.info(f'{txm_file.stem} zipped\n')
    except Exception as e:
        raise Exception(f"Error zipping output: {e}")


def extract_number(stream):
    """
    Extract the numerical part from the second element of a stream name tuple.

    This function is used to sort image streams within an OLE file based on their numerical identifiers.

    Args
    ----
        stream (tuple): A tuple where the second element (stream[1]) is a string containing a numerical identifier.
                        Example: ('ImageData', 'ImageData123')

    Returns
    -------
        int: The extracted numerical identifier as an integer.
        float('inf'): If no numerical identifier is found within the second element.
    """
    match = re.search(r'\d+', stream[1])
    return int(match.group()) if match else float('inf')


def ole_extract(ole, stream, datatype):
    """
    Extract data from a stream within an OLE file.

    Args
    ----
        ole (olefile.OleFileIO): An OleFileIO object representing the OLE file.
        stream (str): The path to the stream within the OLE file (e.g., 'ImageInfo/ImageWidth').
        datatype (str): The struct format string specifying the data type to unpack (e.g., '<I' for unsigned integer, '<f' for float).

    Returns
    -------
        tuple: A tuple containing the unpacked data from the stream, with the data type matching the 'datatype' parameter.

    Raises
    ------
        FileNotFoundError: If the specified stream does not exist in the OLE file.
    """
    if ole.exists(stream):
        data = ole.openstream(stream).read()
        nev = struct.unpack(datatype, data)
        return nev
    else:
        raise FileNotFoundError(f"Stream '{stream}' not found in OLE file.")


def close_dask_client(client):
    """
    Close the Dask client and its associated cluster to release resources.

    This function should be called when the Dask client is no longer needed to prevent resource leaks and ensure proper cleanup.

    Args
    ----
        client (Client): The Dask client object to close.
    """
    if client:
        client.close()
        client.cluster.close()


def display_slice(img):
    """
    Display a single image slice in a window using OpenCV.

    Args
    ----
        img (ndarray): The image slice as a NumPy array (e.g., a 2D array representing a grayscale image).
    """
    try:
        cv2.imshow("Slice", img)
        cv2.waitKey(DISPLAY_SLICE_WAIT_TIME)
        cv2.destroyAllWindows()

    except Exception as e:
        print(f"Error in display_slice: {e}")


class Window(QDialog):
    """Create GUI."""

    def __init__(self):
        """
        Initialise the GUI window and set default values for output variables.

        Sets the following instance variables:
            out_put (int): Index representing the selected output format. Defaults to 0.
            zip (bool): Flag indicating whether to zip output files. Defaults to False.
            slice (bool): Flag indicating whether to display a slice for preview. Defaults to False.
            Dir (str): Path to the selected directory. Defaults to an empty string.

        Calls initUI to populate the window with GUI elements.
        """
        super().__init__()
        self.worker_thread = None
        self.selected_output_index = 0
        self.should_zip_output = False
        self.should_display_slice = False
        self.convert_to_8bit = False
        self.selected_directory = ""
        self.selected_output_base_directory = None
        self.dask_client = None
        self.number_of_files = 0
        self.initUI()

    def initUI(self):
        """
        Set up the GUI layout and add widgets to the window.

        This method configures the window's geometry, title, and style.
        It adds labels, buttons, combo boxes, and radio buttons for user interaction.
        It also connects the widgets to their corresponding event handlers.
        """
        self.setGeometry(GUI_X_POS, GUI_Y_POS, GUI_WIDTH, GUI_HEIGHT)
        self.setWindowTitle("Batch Converter - Settings")
        self.setStyleSheet("""
            QLabel { color: #2c3e50; font-size: 12pt; font-weight: 600; }
            QPushButton, QDialogButtonBox QPushButton {
                background-color: #0d6efd; color: white;
                font-size: 10pt; font-weight: 600; border-radius: 6px;
                padding: 6px 12px; border: none;
            }
            QPushButton:hover, QDialogButtonBox QPushButton:hover {
                background-color: #0b5ed7;
            }
            QComboBox, QRadioButton {
                font-size: 10pt;
            }
        """)

        # Directory Selection
        self.folder_label = QLabel("Selected Input Directory:")
        self.folder_path_label = QLabel("No input directory selected")
        self.folder_path_label.setStyleSheet("font-size: 10pt; font-weight: 500; color: #555;")
        self.folder_button = QDialogButtonBox(QDialogButtonBox.Open)
        self.folder_button.clicked.connect(self.file_dialog)

        # Output Type Selection
        self.output_label = QLabel("Select Output Type:")
        self.output_combo = QComboBox()
        self.output_combo.addItems(["Tiff stack", "3D Tif"])
        self.output_combo.activated.connect(self.activated)

        # Output Folder Selection
        self.output_base_dir_label = QLabel("Output Directory:")
        self.output_base_dir_path_label = QLabel("No output directory selected")
        self.output_base_dir_path_label.setStyleSheet("font-size: 10pt; font-weight: 500; color: #555;")
        self.output_base_dir_button = QDialogButtonBox(QDialogButtonBox.Open)
        self.output_base_dir_button.clicked.connect(self.select_output_base_dir)

        # Zip Option
        self.zip_option = QCheckBox("Zip each image stack")
        self.zip_option.clicked.connect(self.check_zip)

        # Display Slice Option
        self.check_option = QCheckBox("Check each scan before processing")
        self.check_option.clicked.connect(self.check_slice)

        # Convert to 8-bit Option
        self.convert_to_8bit_option = QCheckBox("Convert to 8-bit")
        self.convert_to_8bit_option.clicked.connect(self.check_8bit_conversion)

        # Progress Bar
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setGeometry(30, 80, 440, 25)
        self.progress_bar.setVisible(False)

        # Text Output Window
        self.text_output = QTextEdit(self)
        self.text_output.setReadOnly(True)
        self.text_output.setGeometry(30, 110, 440, 100)
        self.text_output.setVisible(True)
        self.text_output.setStyleSheet("background-color: #f0f0f0; color: #2c3e50; font-size: 10pt; border: 1px solid #bdc3c7;")

        # Dialog Buttons (OK/Cancel)
        self.dialog_buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.dialog_buttons.accepted.connect(self.start_conversion)
        self.dialog_buttons.rejected.connect(self.reject)

        # Processing mode (parallel or serial)
        self.processing_mode_label = QLabel("Processing Mode:")
        self.processing_mode_combo = QComboBox()
        self.processing_mode_combo.addItems(["Parallel", "Serial"])
        self.processing_mode_combo.activated.connect(self.set_processing_mode)
        self.processing_mode = "Parallel"  # Default processing mode.

        # Layout Setup
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        layout.addWidget(self.folder_label)
        folder_layout = QHBoxLayout()
        folder_layout.addWidget(self.folder_button)
        folder_layout.addWidget(self.folder_path_label)
        layout.addLayout(folder_layout)

        layout.addWidget(self.output_base_dir_label)
        output_dir_layout = QHBoxLayout()
        output_dir_layout.addWidget(self.output_base_dir_button)
        output_dir_layout.addWidget(self.output_base_dir_path_label)
        layout.addLayout(output_dir_layout)

        layout.addWidget(self.output_label)
        layout.addWidget(self.output_combo)
        layout.addWidget(self.zip_option)
        layout.addWidget(self.check_option)
        layout.addWidget(self.convert_to_8bit_option)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.text_output)
        layout.addWidget(self.dialog_buttons)
        layout.addWidget(self.processing_mode_label)
        layout.addWidget(self.processing_mode_combo)

        self.setLayout(layout)
        self.thread = None

    def file_dialog(self):
        """
        Open a file dialog to select a directory for batch processing.

        Updates the GUI's folder path label and the instance's 'Dir' attribute
        with the selected directory.
        """
        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            self.folder_path_label.setText(folder)
            self.selected_directory = folder
            self.number_of_files = len(list(Path(self.selected_directory).rglob('*.txm')))
            # Calculate memory per worker and initialize Dask here
            memory_per_worker = AVAILABLE_MEMORY / self.number_of_files
            print(f"Number of files to process: {self.number_of_files}")
            print(f'Memory per worker: {round(memory_per_worker / 1e6, 0)} MB')
            logging.info(f"Number of files to process: {self.number_of_files}")
            logging.info(f'Memory per worker: {round(memory_per_worker / 1e6, 0)} MB')

        cluster = LocalCluster(
            processes=True,
            n_workers=self.number_of_files,
            threads_per_worker=1,
            memory_limit=memory_per_worker,
        )
        self.dask_client = Client(cluster)
        QCoreApplication.instance().dask_client = self.dask_client

    def select_output_base_dir(self):
        """Set the output directory for tiff export."""
        folder = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if folder:
            self.output_base_dir_path_label.setText(folder)
            self.selected_output_base_directory = Path(folder)
        else:
            self.selected_output_base_directory = None

    def check_zip(self):
        """Set the 'zip' instance variable based on the state of the zip option radio button."""
        self.should_zip_output = self.zip_option.isChecked()

    def check_slice(self):
        """Set the 'slice' instance variable based on the state of the check slice radio button."""
        self.should_display_slice = self.check_option.isChecked()

    def check_8bit_conversion(self):
        """Set the 'convert_to_8bit' flag."""
        self.convert_to_8bit = self.convert_to_8bit_option.isChecked()

    def set_processing_mode(self, index):
        """Set the processing mode based on the combo box selection."""
        self.processing_mode = self.processing_mode_combo.itemText(index)

    def activated(self, index):
        """Set the 'out_put' instance variable to the index of the selected output format."""
        self.selected_output_index = index

    def start_conversion(self):
        """Start the conversion process in a separate thread."""
        if not self.selected_directory:
            QMessageBox.warning(self, "Warning", "Please select a folder.")
            return

        if not self.selected_output_base_directory:
            QMessageBox.warning(self, "Warning", "Please select an output directory.")
            return

        if self.dask_client is None:
            QMessageBox.critical(self, "Error", "Dask client is not initialized. Please select a folder.")
            return

        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.dialog_buttons.button(QDialogButtonBox.Ok).setEnabled(False)

        if self.processing_mode == "Parallel":
            self.worker_thread = ParallelWorkerThread(
                Path(self.selected_directory),
                self.selected_output_base_directory,
                self.selected_output_index,
                self.should_zip_output,
                self.should_display_slice,
                self.convert_to_8bit
            )
        else:
            self.worker_thread = SerialWorkerThread(
                Path(self.selected_directory),
                self.selected_output_base_directory,
                self.selected_output_index,
                self.should_zip_output,
                self.should_display_slice,
                self.convert_to_8bit,
            )
        self.worker_thread.progress_update.connect(self.update_progress)
        self.worker_thread.finished.connect(self.conversion_finished)
        self.worker_thread.error.connect(self.conversion_error)
        self.worker_thread.log_message.connect(self.update_text_output)
        self.worker_thread.stopped.connect(self.conversion_stopped)
        self.thread = self.worker_thread
        self.worker_thread.start()

    def update_progress(self, value, filename):
        """
        Update the progress bar and optionally display the current file.

        Args
        ----
            value (int): The current progress value (number of files processed).
            filename (str): The name of the currently processed file.
        """
        progress_percentage = int((value / self.number_of_files) * 100)
        self.progress_bar.setValue(progress_percentage)
        self.setWindowTitle(f"Processing: {filename}")

    def update_text_output(self, message):
        """Update the text output with a new message."""
        self.text_output.append(message)

    def reject(self):
        """Close the dialog and print a cancellation message."""
        if self.worker_thread:
            self.worker_thread.stop()
            self.worker_thread.wait()
        print("Dialog Cancelled")
        close_dask_client(self.dask_client)
        super().reject()

    def conversion_stopped(self):
        """Call when the conversion is stopped."""
        if self.worker_thread:
            self.worker_thread.stop()
            self.worker_thread.wait()
        print("Conversion Stopped.")
        self.update_text_output("Conversion Stopped.")
        self.progress_bar.setVisible(False)
        self.dialog_buttons.setEnabled(True)
        super().reject()

    def conversion_finished(self):
        """Handle the completion of the conversion process."""
        self.setWindowTitle("Batch Converter - Settings")
        self.progress_bar.setVisible(False)
        self.dialog_buttons.setEnabled(True)
        QMessageBox.information(self, "Success", "Conversion complete.")
        close_dask_client(self.dask_client)
        if self.worker_thread:
            self.worker_thread.quit()
            self.worker_thread.wait()
        self.accept()

    def conversion_error(self, message):
        """Handle errors that occur during the conversion process."""
        self.setWindowTitle("Batch Converter - Settings")
        self.progress_bar.setVisible(False)
        self.dialog_buttons.setEnabled(True)
        QMessageBox.critical(self, "Error", f"An error occurred: {message}")
        close_dask_client(self.dask_client)
        self.worker_thread.quit()
        self.worker_thread.wait()

    def closeEvent(self, event):
        """Ensure the worker thread is stopped before the dialog is closed."""
        if self.worker_thread and self.worker_thread.isRunning():
            self.worker_thread.quit()
            self.worker_thread.wait()
        close_dask_client(self.dask_client)
        event.accept()


def Set_Batch():
    """
    Open the GUI window and initiate the batch conversion process if a directory is selected.

    Displays an error message if an exception occurs during conversion.
    Logs the completion time or any errors encountered.
    """
    app = QCoreApplication.instance()
    if app is None:
        app = QApplication(sys.argv)

    print(f'Available RAM: {round(AVAILABLE_MEMORY/1e6, 0)} MB')
    logging.info(f'Available RAM: {round(AVAILABLE_MEMORY/1e6, 0)} MB')

    win = Window()
    win.show()
    app.exec_()
    app.exit()


if __name__ == '__main__':
    Set_Batch()

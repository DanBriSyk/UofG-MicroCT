"""
TXM/RCP file metadata extractor.

Generates a GUI for users to extract metadata from a RCP/TXM/TXRM file and export as TXT/CSV file or in console.
"""

# Author: Daniel Bribiesca Sykes <daniel.bribiescasykes@glasgow.ac.uk>
# Version: 3.2.0

import olefile as olef
from pathlib import Path
import os
import struct
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication,
    QDialog,
    QLabel,
    QDialogButtonBox,
    QRadioButton,
    QGridLayout,
    QFileDialog,
    QVBoxLayout,
)
import math
from datetime import datetime
import sys


class MetadataExtractorGUI(QDialog):
    """GUI for metadata extraction."""

    def __init__(self):
        """Initialize the GUI."""
        super().__init__()
        self.out_file = 3
        self.file_path = ""
        self.init_ui()

    def init_ui(self):
        """Set up the GUI layout and widgets."""
        self.setWindowTitle("Metadata Extractor")
        layout = QVBoxLayout()

        # File selection
        file_layout = QGridLayout()
        self.file_label = QLabel("Select file:")
        self.file_path_label = QLabel("")
        self.file_button = QDialogButtonBox(QDialogButtonBox.Open)
        self.file_button.clicked.connect(self.select_file)

        file_layout.addWidget(self.file_label, 0, 0)
        file_layout.addWidget(self.file_path_label, 0, 1)
        file_layout.addWidget(self.file_button, 1, 0, 1, 2)
        layout.addLayout(file_layout)

        # Output options
        output_layout = QGridLayout()
        self.output_label = QLabel("Output format:")
        self.txt_radio = QRadioButton("Tab-delimited TXT file")
        self.csv_radio = QRadioButton("CSV file")
        self.console_radio = QRadioButton("Console only")

        self.txt_radio.clicked.connect(self.set_output_format)
        self.csv_radio.clicked.connect(self.set_output_format)
        self.console_radio.clicked.connect(self.set_output_format)
        self.console_radio.setChecked(True)

        output_layout.addWidget(self.output_label, 0, 0)
        output_layout.addWidget(self.txt_radio, 1, 0)
        output_layout.addWidget(self.csv_radio, 2, 0)
        output_layout.addWidget(self.console_radio, 3, 0)
        layout.addLayout(output_layout)

        # Action buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

        self.setLayout(layout)

    def select_file(self):
        """Open a file dialog to select a file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select file", "", "RCP/TXM/TXRM files (*.rcp *.txm *.txrm)"
        )
        if file_path:
            self.file_path = file_path
            self.file_path_label.setText(os.path.basename(file_path))

    def set_output_format(self):
        """Set the output format based on radio button selection."""
        if self.txt_radio.isChecked():
            self.out_file = 1
        elif self.csv_radio.isChecked():
            self.out_file = 2
        else:
            self.out_file = 3


def stream_unpacker(ole, stream, datatype):
    """Extract parameters from specified ole file and return."""
    try:
        header = ole.openstream(stream).read()
        if datatype == 0:
            return header
        else:
            return struct.unpack(datatype, header)
    except Exception as e:
        print(f"Error unpacking stream {stream}: {e}")
        return None


def stream_unpacker_from(ole, stream, datatype, offset):
    """Extract parameters with an offset from specified ole file and return."""
    try:
        header = ole.openstream(stream).read()
        return struct.unpack_from(datatype, header, offset)
    except Exception as e:
        print(f"Error unpacking stream {stream} with offset {offset}: {e}")
        return None


def get_versa_projections(ole, file_suffix):
    """Get the number of projections taken for Versa files."""
    if file_suffix == ".txrm":
        return stream_unpacker(ole, "ImageInfo/ImagesTaken", "<i")[0]
    elif file_suffix == ".txm":
        return stream_unpacker(ole, "AutoRecon/NumOfProjects", "<i")[0]
    return None


def get_versa_rotation(ole, file_suffix):
    """Get the rotation angle for Versa files."""
    if file_suffix == ".txm":
        return round(stream_unpacker(ole, "AutoRecon/AngleSpan", "<f")[0], 0)
    elif file_suffix == ".txrm":
        end_angle = stream_unpacker(ole, "AcquisitionSettings/EndAngle", "<f")[0]
        start_angle = stream_unpacker(ole, "AcquisitionSettings/StartAngle", "<f")[0]
        return round(abs(end_angle) + abs(start_angle))
    return None


def get_versa_exposure(ole, file_suffix):
    """Get the exposure time for Versa files."""
    if file_suffix == ".txrm":
        return round(stream_unpacker(ole, "AcquisitionSettings/ExpTime", "<f")[0], 2)
    elif file_suffix == ".txm":
        return round(stream_unpacker_from(ole, "Imageinfo/ExpTimes", "<f", 4)[0], 2)
    return None


def get_versa_filter(ole, file_suffix):
    """Get the filter name for Versa files."""
    if file_suffix == ".txrm":
        return stream_unpacker(ole, "AcquisitionSettings/SourceFilterName", 0).decode("ascii")[:-257]
    elif file_suffix == ".txm":
        return stream_unpacker(ole, "ImageInfo/SourceFilterName", 0).decode("ascii")[:-257]
    return None


def get_versa_src_dist(ole, file_suffix, src_dist):
    """Get the source-object distance for Versa files."""
    if file_suffix == ".txm":
        return round(src_dist / 1000, 2)
    elif file_suffix == ".txrm":
        return round(abs(src_dist), 2)
    return None


def get_versa_det_dist(ole, file_suffix, det_dist):
    """Get the detector-object distance for Versa files."""
    if file_suffix == ".txm":
        return round(det_dist / 1000, 2)
    elif file_suffix == ".txrm":
        return round(det_dist, 2)
    return None


def get_versa_acq_mode(ole, file_suffix):
    """Get the acquisition mode for Versa files."""
    if file_suffix == ".txrm":
        acq_mode_str = stream_unpacker(ole, "AcquisitionSettings/AcqModeString", 0).decode("ascii")[:-245]
        acq_mode_val = None
    elif file_suffix == ".txm":
        acq_mode_val = stream_unpacker(ole, "ImageInfo/AcquisitionMode", "<i")[0]
        acq_mode_str = None
    stitch_enabled = stream_unpacker(ole, "ReconSettings/StitchParams/AutoStitchSettings/Enabled", "?")[0]
    if stitch_enabled:
        segs = stream_unpacker(ole, "ReconSettings/StitchParams/AutoStitchSettings/NumSegments", "<i")[0]
    else:
        segs = 1
    acq_mode = _get_acq_mode(stitch_enabled, acq_mode_str, acq_mode_val)
    return acq_mode, segs


def _get_acq_mode(stitch_enabled, acq_mode_str, acq_mode_val):
    """
    Determine the acquisition mode based on stitching and acquisition mode values.

    This helper function calculates the acquisition mode string based on whether
    stitching is enabled and the provided acquisition mode string or value.

    Args
    ----
        stitch_enabled (bool): A boolean indicating if stitching is enabled.
        acq_mode_str (str, optional): The acquisition mode string (e.g., "Tomography Wide").
                                       Defaults to None.
        acq_mode_val (int, optional): The acquisition mode integer value (e.g., 17).
                                       Defaults to None.

    Returns
    -------
        str: The determined acquisition mode string (e.g., "Wide Stitch", "Wide", "Stitch", "Normal").
    """
    if stitch_enabled:
        if acq_mode_str == "Tomography Wide" or acq_mode_val == 17:
            acq_mode = "Wide Stitch"
        else:
            acq_mode = "Stitch"
    else:
        if acq_mode_str == "Tomography Wide" or acq_mode_val == 17:
            acq_mode = "Wide"
        else:
            acq_mode = "Normal"
    return acq_mode


def extract_metadata(file_path, out_file):
    """Extract metadata from RCP/TXM/TXRM files."""
    try:
        file_path = Path(file_path)
        with olef.OleFileIO(file_path) as ole:
            metadata = {}
            metadata["File"] = f"File:\t{file_path.name}\n"
            file_suffix = file_path.suffix

            if file_suffix == ".rcp":
                tomo_datasets = stream_unpacker(ole, "NoOfTomoDataSets", "<i")[0]
                print(f"Number of recipes:\t{tomo_datasets}\n")
                for x in range(tomo_datasets):
                    recipe_name = stream_unpacker(ole, f"RecipePoint{x}/PointName", 0).decode("ascii")[:-1]
                    metadata["Recipe"] = f"Recipe:\t{recipe_name}\n"
                    extract_recipe_data(ole, metadata, x)
                    print_or_write_metadata(metadata.copy(), out_file, file_path, recipe_name)
            else:
                extract_common_data(ole, metadata, file_suffix)
                print_or_write_metadata(metadata, out_file, file_path)
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")


def extract_common_data(ole, metadata, file_suffix):
    """Extract common metadata from TXM/TXRM files."""
    date_str = stream_unpacker(ole, "ImageInfo/Date", 0)
    date = datetime.strptime(date_str[:10].decode("ascii"), "%m/%d/%Y")
    metadata["Date"] = f"Date:\t{date:%d/%m/%Y}\n"
    time = date_str[11:19].decode("ascii")
    metadata["Time"] = f"Time:\t{time}\n"

    volts_data = stream_unpacker(ole, "ImageInfo/Voltage", "<f")
    curr_data = stream_unpacker(ole, "ImageInfo/Current", "<f")
    volts = volts_data[0]
    curr = curr_data[0]
    metadata["kV"] = f"kV:\t{volts}\n"
    metadata["uA"] = f"uA:\t{curr}\n"
    metadata["Power"] = f"Power (W):\t{round(volts * (curr / 1000), 1)}\n"

    proj = get_versa_projections(ole, file_suffix)
    metadata["imgs"] = f"Projections taken:\t{proj}\n"

    rot = get_versa_rotation(ole, file_suffix)
    metadata["rot"] = f"Rotation (deg):\t{rot}\n"

    exp = get_versa_exposure(ole, file_suffix)
    metadata["exposure"] = f"Exposure (secs):\t{exp}\n"

    mag_data = stream_unpacker(ole, "ImageInfo/ObjectiveName", 0)
    mag_list = mag_data.decode("IBM855")[:-256].partition("X")[0:2]
    metadata["objlens"] = f"Objective lens:\t{''.join(mag_list)}\n"

    filt = get_versa_filter(ole, file_suffix)
    metadata["XFilt"] = f"Filter:\t{filt}\n"

    pixel_size_data = stream_unpacker(ole, "ImageInfo/PixelSize", "<f")
    metadata["Vox_size"] = f"Voxel size (um):\t{round(pixel_size_data[0], 2)}\n"

    cone_data = stream_unpacker_from(ole, "ImageInfo/ConeAngle", "<f", 0)
    metadata["cone"] = f"Cone angle (deg):\t{round(cone_data[0], 2)}\n"

    binning_data = stream_unpacker(ole, "ImageInfo/CameraBinning", "<i")
    metadata["D_bin"] = f"Binning:\t{binning_data[0]}\n"

    frame_avg_data = stream_unpacker(ole, "ImageInfo/CameraNumberOfFramesPerImage", "<i")
    metadata["fr_avg"] = f"Frame Averaging:\t{frame_avg_data[0]}\n"

    beam_hard_data = stream_unpacker(ole, "ReconSettings/BeamHardening", "<f")
    metadata["beam_h"] = f"Beam Hardening:\t{round(beam_hard_data[0], 2)}\n"

    src_dist_data = stream_unpacker_from(ole, "ImageInfo/StoRADistance", "<f", 0)
    det_dist_data = stream_unpacker_from(ole, "ImageInfo/DtoRADistance", "<f", 0)
    src_dist = src_dist_data[0]
    det_dist = det_dist_data[0]
    metadata["Src"] = f"Src-Obj distance (mm):\t{get_versa_src_dist(ole, file_suffix, src_dist)}\n"
    metadata["Det"] = f"Det-Obj distance (mm):\t{get_versa_det_dist(ole, file_suffix, det_dist)}\n"

    x_pos_data = stream_unpacker_from(ole, "ImageInfo/XPosition", "<f", 0)
    y_pos_data = stream_unpacker_from(ole, "ImageInfo/YPosition", "<f", 0)
    z_pos_data = stream_unpacker_from(ole, "ImageInfo/ZPosition", "<f", 0)
    metadata["x_ax"] = f"X position (um):\t{round(x_pos_data[0], 2)}\n"
    metadata["y_ax"] = f"Y position (um):\t{round(y_pos_data[0], 2)}\n"
    metadata["z_ax"] = f"Z position (um):\t{round(z_pos_data[0], 2)}\n"

    acq_mode, segs = get_versa_acq_mode(ole, file_suffix)
    metadata["Acquisition mode"] = f"Acquisition mode:\t{acq_mode}\n"
    metadata["Segments"] = f"No. of segments:\t{segs}\n"

    if file_suffix == ".txrm":
        var_angle_mode = stream_unpacker(ole, "AcquisitionSettings/VariableAngleMode", "<i")
        metadata["HART"] = f"HART:\t{'Enabled' if var_angle_mode[0] == 1 else 'Disabled'}\n"
        var_exp_mode = stream_unpacker(ole, "AcquisitionSettings/VariableExposureTimeMode", "<i")
        metadata["V_exp"] = f"Variable exposure:\t{'Enabled' if var_exp_mode[0] == 1 else 'Disabled'}\n"


def extract_recipe_data(ole, metadata, x):
    """Extract metadata from RCP files."""
    date_str = stream_unpacker(ole, "TimeStamp", 0)
    if date_str:
        date = datetime.strptime(date_str[:10].decode("ascii"), "%Y-%m-%d")
        metadata["Date"] = f"Date:\t{date:%d/%m/%Y}\n"
        time = datetime.strptime(date_str[11:17].decode("ascii"), "%H%M%S")
        metadata["Time"] = f"Time:\t{time:%H:%M:%S}\n"

    volts = stream_unpacker(ole, f"RecipePoint{x}/AcquisitionSettings/SrcVoltage", "<f")[0]
    watts = stream_unpacker(ole, f"RecipePoint{x}/AcquisitionSettings/SrcPower", "<f")[0]
    curr = round((watts * 1000) / volts, 1) if volts != 0 and watts != 0 else 0.0
    metadata["kV"] = f"kV:\t{volts}\n"
    metadata["Power"] = f"Power (W):\t{watts}\n"
    metadata["uA"] = f"uA:\t{curr}\n"

    metadata["imgs"] = f"Projections taken:\t{stream_unpacker(ole, f'RecipePoint{x}/AcquisitionSettings/TotalImages', '<i')[0]}\n"
    end_angle = stream_unpacker(ole, f"RecipePoint{x}/AcquisitionSettings/EndAngle", "<f")[0]
    start_angle = stream_unpacker(ole, f"RecipePoint{x}/AcquisitionSettings/StartAngle", "<f")[0]
    metadata["rot"] = f"Rotation (deg):\t{round(abs(end_angle) + abs(start_angle))}\n"
    metadata["exposure"] = f"Exposure (secs):\t{round(stream_unpacker(ole, f'RecipePoint{x}/AcquisitionSettings/ExpTime', '<f')[0], 2)}\n"
    metadata["objlens"] = f"Objective lens:\t{stream_unpacker(ole, f'RecipePoint{x}/MagStr', 0).decode('ascii')[:-1]}\n"
    metadata["XFilt"] = f"Filter:\t{stream_unpacker(ole, f'RecipePoint{x}/AcquisitionSettings/SourceFilterName', 0).decode('ascii')[:-257]}\n"
    metadata["D_bin"] = f"Binning:\t{stream_unpacker(ole, f'RecipePoint{x}/AcquisitionSettings/Binning', '<i')[0]}\n"
    metadata["fr_avg"] = f"Frame Averaging:\t{stream_unpacker(ole, f'RecipePoint{x}/AcquisitionSettings/FramesPerImage', '<i')[0]}\n"
    metadata["beam_h"] = f"Beam Hardening:\t{round(stream_unpacker(ole, f'RecipePoint{x}/ReconSettings/BeamHardening', '<f')[0], 2)}\n"

    src_dist = stream_unpacker_from(ole, f"RecipePoint{x}/AcquisitionSettings/InitialPositions", "<f", 16)[0]
    det_dist = stream_unpacker_from(ole, f"RecipePoint{x}/AcquisitionSettings/InitialPositions", "<f", 20)[0]
    metadata["Src"] = f"Src-Obj distance (mm):\t{round(abs(src_dist), 2)}\n"
    metadata["Det"] = f"Det-Obj distance (mm):\t{round(det_dist, 2)}\n"

    ccd_size = stream_unpacker(ole, f"RecipePoint{x}/AcquisitionSettings/CCDPixelSize", "<f")[0]
    mag = stream_unpacker(ole, f"RecipePoint{x}/MagStr", 0).decode("ascii")[:-1]
    binning = stream_unpacker(ole, f"RecipePoint{x}/AcquisitionSettings/Binning", "<i")[0]
    geom_mag = (abs(src_dist) + det_dist) / abs(src_dist)
    pixel_size = (ccd_size / float(mag[:-1]) / geom_mag) * binning
    metadata["Vox_size"] = f"Voxel size (um):\t{round(pixel_size, 2)}\n"

    detrad = pixel_size * (2048 / binning)
    slant = math.sqrt(((abs(src_dist) + det_dist) ** 2) + (detrad ** 2))
    metadata["cone"] = f"Cone angle (deg):\t{round(2 * math.asin(detrad / slant), 2)}\n"

    x_pos_data = stream_unpacker_from(ole, f'RecipePoint{x}/AcquisitionSettings/InitialPositions', '<f', 0)[0]
    y_pos_data = stream_unpacker_from(ole, f'RecipePoint{x}/AcquisitionSettings/InitialPositions', '<f', 4)[0]
    z_pos_data = stream_unpacker_from(ole, f'RecipePoint{x}/AcquisitionSettings/InitialPositions', '<f', 8)[0]
    metadata["x_ax"] = f"X position (um):\t{round(x_pos_data, 2)}\n"
    metadata["y_ax"] = f"Y position (um):\t{round(y_pos_data, 2)}\n"
    metadata["z_ax"] = f"Z position (um):\t{round(z_pos_data, 2)}\n"

    acq_mode_str = stream_unpacker(ole, f"RecipePoint{x}/AcquisitionSettings/AcqModeString", 0).decode("ascii")[:-245]
    stitch_enabled = stream_unpacker(ole, f"RecipePoint{x}/AutoStitchSettings/Enabled", "?")[0]
    if stitch_enabled:
        segs = stream_unpacker(ole, f"RecipePoint{x}/AutoStitchSettings/NumSegments", "<i")[0]
    else:
        segs = 1
    acq_mode = _get_acq_mode(stitch_enabled, acq_mode_str, acq_mode_val=None)
    metadata["Acquisition mode"] = f"Acquisition mode:\t{acq_mode}\n"
    metadata["Segments"] = f"No. of segments:\t{segs}\n"

    HART_mode = stream_unpacker(ole, f'RecipePoint{x}/AcquisitionSettings/VariableAngleMode', '<i')[0]
    VE_mode = stream_unpacker(ole, f'RecipePoint{x}/AcquisitionSettings/VariableExposureTimeMode', '<i')[0]
    metadata["HART"] = f"HART:\t{'Enabled' if HART_mode == 1 else 'Disabled'}\n"
    metadata["V_exp"] = f"Variable exposure:\t{'Enabled' if VE_mode == 1 else 'Disabled'}\n"


def print_or_write_metadata(metadata, out_file, file_path, recipe_name=None):
    """Print or write metadata to a file or console."""
    if out_file == 1:  # TXT
        filename = f"{file_path.stem}_{recipe_name}_{file_path.suffix[1:]}.txt" if recipe_name else f"{file_path.stem}_{file_path.suffix[1:]}.txt"
        with open(filename, "w") as f:
            for value in metadata.values():
                f.write(value)
    elif out_file == 2:  # CSV
        filename = f"{file_path.stem}_{recipe_name}_{file_path.suffix[1:]}.csv" if recipe_name else f"{file_path.stem}_{file_path.suffix[1:]}.csv"
        with open(filename, "w") as f:
            for value in metadata.values():
                f.write(value.replace("\t", ","))
    elif out_file == 3:  # Console
        for value in metadata.values():
            print(value.strip())
        print("\n")


def main():
    """Call main function to run the application."""
    app = QApplication(sys.argv)
    gui = MetadataExtractorGUI()
    if gui.exec_() == QDialog.Accepted:
        if gui.file_path:
            extract_metadata(gui.file_path, gui.out_file)
        else:
            print("No file selected.")
    sys.exit()


if __name__ == "__main__":
    main()

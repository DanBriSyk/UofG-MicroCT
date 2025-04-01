"""
TXM/RCP file metadata extractor.

Generates a GUI for users to extract metadata from a RCP/TXM/TXRM file and export as TXT/CSV file or in console.
"""

# Author: Daniel Bribiesca Sykes <daniel.bribiescasykes@glasgow.ac.uk>
# Version: 2.3.2

import olefile as olef
from pathlib import Path
import os
import struct
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QDialog, QLabel, QDialogButtonBox, QRadioButton, QGridLayout, QFileDialog
import math
from datetime import datetime


class Window(QDialog):
    """Create GUI."""

    def __init__(self):
        """
        Initialise GUI and set output variables with default values.

        Returns
        -------
        Run initUI to populate Window

        """
        super(Window, self).__init__()
        self.out_file = 3
        self.File = ''
        self.initUI()

    def initUI(self):
        """
        Format layout and buttons in GUI.

        Returns
        -------
        None.

        """
        self.setGeometry(200, 200, 500, 200)
        self.setWindowTitle("Select a RCP/TXM/TXRM file to open")

        self.l1 = QLabel(self)
        self.l1.setText("Selected file:")
        self.l1.setStyleSheet("color: #464d55;"
                              "font-weight: 600;")

        self.l2 = QLabel(self)
        self.l2.setText("")
        self.l2.setStyleSheet("color: #464d55;"
                              "font-weight: 600;")

        self.l3 = QLabel(self)
        self.l3.setText("Output file:")
        self.l3.setStyleSheet("color: #464d55;"
                              "font-weight: 600;")

        self.b1 = QDialogButtonBox(self)
        self.b1.setStandardButtons(QDialogButtonBox.Open)
        self.b1.setStyleSheet("background-color: #0d6efd;"
                              "color: #ffffff;"
                              "font-weight: 600;"
                              "border-radius: 8px;"
                              "border: 1px solid #0d6efd;"
                              "padding: 5px 15px;"
                              "outline: 0px;")
        self.b1.clicked.connect(self.file_dialog)

        self.r1 = QRadioButton(self)
        self.r1.setText("Tab-delimited TXT file")
        self.r1.clicked.connect(self.check)

        self.r2 = QRadioButton(self)
        self.r2.setText("CSV file")
        self.r2.clicked.connect(self.check)

        self.r3 = QRadioButton(self)
        self.r3.setText("None - display only")
        self.r3.clicked.connect(self.check)

        self.bbox = QDialogButtonBox(self)
        self.bbox.setStandardButtons(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.bbox.setStyleSheet("background-color: #0d6efd;"
                                "color: #fff;"
                                "font-weight: 600;"
                                "border-radius: 8px;"
                                "border: 1px solid #0d6efd;"
                                "padding: 5px 15px;"
                                "margin-top: 10px;"
                                "outline: 0px;")
        self.bbox.accepted.connect(self.accept)
        self.bbox.rejected.connect(self.reject)

        layout = QGridLayout()
        self.setLayout(layout)
        layout.setHorizontalSpacing(15)
        layout.addWidget(self.b1, 0, 0, alignment=Qt.AlignLeft)
        layout.addWidget(self.l1, 0, 1, alignment=Qt.AlignLeft)
        layout.addWidget(self.l2, 0, 2, alignment=Qt.AlignLeft)
        layout.addWidget(self.l3, 2, 0, alignment=Qt.AlignRight)
        layout.addWidget(self.r1, 2, 1, alignment=Qt.AlignLeft)
        layout.addWidget(self.r2, 3, 1, alignment=Qt.AlignLeft)
        layout.addWidget(self.r3, 4, 1, alignment=Qt.AlignLeft)
        layout.addWidget(self.bbox, 5, 2, alignment=Qt.AlignLeft)

    def file_dialog(self):
        """
        Open dialog window to select file to read.

        Returns
        -------
        Sets File variable to Path of file selected.

        """
        File = QFileDialog.getOpenFileName(self, 'Select file', '', "RCP/TXM/TXRM files (*.rcp *.txm *.txrm)")
        self.l2.setText(File[0])
        self.l2.adjustSize()
        self.File = File[0]

    def check(self):
        """
        Check if radia buttons are checked.

        Returns
        -------
        Sets out_file variable depending on selection.

        """
        if self.r1.isChecked():
            self.out_file = 1
        elif self.r2.isChecked():
            self.out_file = 2
        else:
            self.out_file = 3


def GetFile():
    """
    Open GUI class, once closed either input file into read_files or cancel script.

    Returns
    -------
    None.

    """
    win = Window()
    result = win.exec_()
    if result == QDialog.Accepted:
        if win.File != '':
            read_files(Path(win.File), win.out_file)
        else:
            print("No file selected")
    else:
        print("User cancelled the operation")
    win.close()


def read_files(f, out):
    """
    Determine file type and assign appropriate function.

    Returns
    -------
    None.

    """
    if f.suffix == '.txrm' or f.suffix == '.txm':
        print("Reading Versa TXM/TXRM data")
        Versa_reader(f, out)
    elif f.suffix == '.rcp':
        print("Reading Recipe file")
        Recipe_reader(f, out)


def stream_unpacker(ole, stream, datatype):
    """
    Extract parameters from specified ole file and return.

    Parameters
    ----------
    ole : ole file
        Imported data file to extract info from.
    stream : olefile header
        Specific parameter to extract.
    datatype: data format
        Float/int/etc.

    Returns
    -------
    header_data : unpacked data
        Unpacked parameter from ole file.

    """
    header = ole.openstream(stream).read()
    if datatype == 0:
        header_data = header
    else:
        header_data = struct.unpack(datatype, header)
    return header_data


def stream_unpacker_from(ole, stream, datatype, offset):
    """
    Extract parameters with an offset from specified ole file and return.

    Parameters
    ----------
    File : ole file
        Imported data file to extract info from.
    stream : olefile header
        Specific parameter to extract.
    datatype: data format
        Float/int/etc.
    offset: integer
        Mumber of characters to skip in packed data.

    Returns
    -------
    header_data : unpacked data
        Unpacked parameter from ole file.

    """
    header = ole.openstream(stream).read()
    header_data = struct.unpack_from(datatype, header, offset)
    return header_data


def Versa_reader(File, out_file):
    """
    Open .txm or .txrm files and extract scan parameters.

    Parameters
    ----------
    File : Path
        TXM/TXRM files batch imported from recursive folder search from GetFile().

    Returns
    -------
    Output all scan parameters in console or .txt or .csv file.

    """
    nl = '\n'
    tab = '\t'
    param_dict = {}
    param_dict['F_name'] = f"File:{tab}{File.name}{nl}"

    with olef.OleFileIO(File) as ole:
        # print(ole.listdir())
        date = datetime.strptime(stream_unpacker(ole, 'ImageInfo/Date', 0)[:10].decode("ascii"), '%m/%d/%Y')
        param_dict['Date'] = f"Date:{tab}{date:%d/%m/%Y}{nl}"

        time = stream_unpacker(ole, 'ImageInfo/Date', 0)[11:19].decode("ascii")
        param_dict['Time'] = f"Time:{tab}{time}{nl}"

        volts = stream_unpacker(ole, 'ImageInfo/Voltage', "<f")[0]
        param_dict['kV'] = f"kV:{tab}{volts}{nl}"

        curr = stream_unpacker(ole, 'ImageInfo/Current', "<f")[0]
        param_dict['uA'] = f"uA:{tab}{curr}{nl}"

        Watts = round(volts * (curr / 1000), 1)
        param_dict['Power'] = f"Power (W):{tab}{Watts}{nl}"

        if File.suffix == '.txrm':
            proj = stream_unpacker(ole, 'ImageInfo/ImagesTaken', "<i")[0]
        elif File.suffix == '.txm':
            proj = stream_unpacker(ole, 'AutoRecon/NumOfProjects', "<i")[0]
        param_dict['imgs'] = f"Projections taken:{tab}{proj}{nl}"

        if File.suffix == '.txm':
            rot = round(stream_unpacker(ole, 'AutoRecon/AngleSpan', "<f")[0], 0)
        elif File.suffix == '.txrm':
            End_angle = stream_unpacker(ole, 'AcquisitionSettings/EndAngle', "<f")[0]
            Sta_angle = stream_unpacker(ole, 'AcquisitionSettings/StartAngle', "<f")[0]
            rot = round(abs(End_angle) + abs(Sta_angle))
        param_dict['rot'] = f"Rotation (deg):{tab}{rot}{nl}"

        if File.suffix == '.txrm':
            exp = stream_unpacker(ole, 'AcquisitionSettings/ExpTime', "<f")[0]
        elif File.suffix == '.txm':
            exp = stream_unpacker_from(ole, 'Imageinfo/ExpTimes', "<f", 4)[0]
        param_dict['exposure'] = f"Exposure (secs):{tab}{round(exp, 2)}{nl}"

        mag_list = stream_unpacker(ole, 'ImageInfo/ObjectiveName', 0).decode('IBM855')[:-256].partition('X')[0:2]
        mag = ''.join(mag_list)
        param_dict['objlens'] = f"Objective lens:{tab}{mag}{nl}"

        if File.suffix == '.txrm':
            filt = stream_unpacker(ole, 'AcquisitionSettings/SourceFilterName', 0).decode("ascii")[:-257]
        elif File.suffix == '.txm':
            filt = stream_unpacker(ole, 'ImageInfo/SourceFilterName', 0).decode("ascii")[:-257]
        param_dict['XFilt'] = f"Filter:{tab}{filt}{nl}"

        Px_size = stream_unpacker(ole, 'ImageInfo/PixelSize', "<f")[0]
        param_dict['Vox_size'] = f"Voxel size (um):{tab}{round(Px_size, 2)}{nl}"

        cone = stream_unpacker_from(ole, 'ImageInfo/ConeAngle', "<f", 0)[0]
        param_dict['cone'] = f"Cone angle (deg):{tab}{round(cone, 2)}{nl}"

        binning = stream_unpacker(ole, 'ImageInfo/CameraBinning', "<i")[0]
        param_dict['D_bin'] = f"Binning:{tab}{binning}{nl}"

        f_avg = stream_unpacker(ole, 'ImageInfo/CameraNumberOfFramesPerImage', "<i")[0]
        param_dict['fr_avg'] = f"Frame Averaging:{tab}{f_avg}{nl}"

        bh = stream_unpacker(ole, 'ReconSettings/BeamHardening', "<f")[0]
        param_dict['beam_h'] = f"Beam Hardening:{tab}{round(bh, 2)}{nl}"

        srctosample = stream_unpacker_from(ole, 'ImageInfo/StoRADistance', "<f", 0)[0]
        if File.suffix == '.txm':
            param_dict['Src'] = f"Src-Obj distance (mm):{tab}{round(srctosample/1000, 2)}{nl}"
        elif File.suffix == '.txrm':
            param_dict['Src'] = f"Src-Obj distance (mm):{tab}{round(abs(srctosample), 2)}{nl}"

        dettosample = stream_unpacker_from(ole, 'ImageInfo/DtoRADistance', "<f", 0)[0]
        if File.suffix == '.txm':
            param_dict['Det'] = f"Det-Obj distance (mm):{tab}{round(dettosample/1000, 2)}{nl}"
        elif File.suffix == '.txrm':
            param_dict['Det'] = f"Det-Obj distance (mm):{tab}{round(dettosample, 2)}{nl}"

        x_pos = stream_unpacker_from(ole, 'ImageInfo/XPosition', "<f", 0)[0]
        param_dict['x_ax'] = f"X position (um):{tab}{round(x_pos, 2)}{nl}"

        y_pos = stream_unpacker_from(ole, 'ImageInfo/YPosition', "<f", 0)[0]
        param_dict['y_ax'] = f"Y position (um):{tab}{round(y_pos, 2)}{nl}"

        z_pos = stream_unpacker_from(ole, 'ImageInfo/ZPosition', "<f", 0)[0]
        param_dict['z_ax'] = f"Z position (um):{tab}{round(z_pos, 2)}{nl}"

        if File.suffix == '.txrm':
            if stream_unpacker(ole, 'AcquisitionSettings/AcqModeString', 0).decode("ascii")[:-245] == 'Tomography Wide':
                if stream_unpacker(ole, 'ReconSettings/StitchParams/AutoStitchSettings/Enabled', "?")[0] is True:
                    param_dict['Acquisition mode'] = f"Acquisition mode:{tab}Wide Stitch{nl}"
                    segs = stream_unpacker(ole, 'ReconSettings/StitchParams/AutoStitchSettings/NumSegments', "<i")[0]
                    param_dict['Segments'] = f"No. of segments:{tab}{segs}{nl}"
                else:
                    param_dict['Acquisition mode'] = f"Acquisition mode:{tab}Wide{nl}"
                    param_dict['Segments'] = f"No. of segments:{tab}1{nl}"
            elif stream_unpacker(ole, 'ReconSettings/StitchParams/AutoStitchSettings/Enabled', "?")[0] is True:
                param_dict['Acquisition mode'] = f"Acquisition mode:{tab}Stitch{nl}"
                segs = stream_unpacker(ole, 'ReconSettings/StitchParams/AutoStitchSettings/NumSegments', "<i")[0]
                param_dict['Segments'] = f"No. of segments:{tab}{segs}{nl}"
            else:
                param_dict['Acquisition mode'] = f"Acquisition mode:{tab}Normal{nl}"
                param_dict['Segments'] = f"No. of segments:{tab}1{nl}"

            HART = stream_unpacker(ole, 'AcquisitionSettings/VariableAngleMode', "<i")[0]
            if HART == 1:
                param_dict['HART'] = f"HART:{tab}Enabled{nl}"
            else:
                param_dict['HART'] = f"HART:{tab}Disabled{nl}"

            V_exp = stream_unpacker(ole, 'AcquisitionSettings/VariableExposureTimeMode', "<i")[0]
            if V_exp == 1:
                param_dict['V_exp'] = f"Variable exposure:{tab}Enabled{nl}"
            else:
                param_dict['V_exp'] = f"Variable exposure:{tab}Disabled{nl}"

        if File.suffix == '.txm':
            if stream_unpacker(ole, 'ImageInfo/AcquisitionMode', "<i")[0] == 17:
                if stream_unpacker(ole, 'ReconSettings/StitchParams/AutoStitchSettings/Enabled', "?")[0] is True:
                    param_dict['Acquisition mode'] = f"Acquisition mode:{tab}Wide Stitch{nl}"
                    segs = stream_unpacker(ole, 'ReconSettings/StitchParams/AutoStitchSettings/NumSegments', "<i")[0]
                    param_dict['Segments'] = f"No. of segments:{tab}{segs}{nl}"
                else:
                    param_dict['Acquisition mode'] = f"Acquisition mode:{tab}Wide{nl}"
                    param_dict['Segments'] = f"No. of segments:{tab}1{nl}"
            elif stream_unpacker(ole, 'ImageInfo/AcquisitionMode', "<i")[0] == 10:
                if stream_unpacker(ole, 'ReconSettings/StitchParams/AutoStitchSettings/Enabled', "?")[0] is True:
                    param_dict['Acquisition mode'] = f"Acquisition mode:{tab}Stitch{nl}"
                    segs = stream_unpacker(ole, 'ReconSettings/StitchParams/AutoStitchSettings/NumSegments', "<i")[0]
                    param_dict['Segments'] = f"No. of segments:{tab}{segs}{nl}"
                else:
                    param_dict['Acquisition mode'] = f"Acquisition mode:{tab}Normal{nl}"
                    param_dict['Segments'] = f"No. of segments:{tab}1{nl}"

    ole.close()

    for value in param_dict.values():
        print(value.strip())

    if out_file == 1:
        with open(f'{File.stem}_{File.suffix[1:]}.txt', 'w') as filehandle:
            for value in param_dict.values():
                filehandle.write(value)

    elif out_file == 2:
        with open(f'{File.stem}_{File.suffix[1:]}.csv', 'w') as csv_write:
            for value in param_dict.values():
                csv_write.write(value.replace('\t', ','))

    elif out_file == 3:
        valid = ["Y", "N", "y", "n"]
        while True:
            try:
                end = input("Would you like to open another file? (Y/N) ")
            except ValueError:
                print("Invalid input")
                continue
            if end not in valid:
                print("Invalid input")
                continue
            elif end.upper() == "Y":
                GetFile(3)
                break
            elif end.upper() == "N":
                break


def Recipe_reader(File, out_file):
    """
    Open .rcp file and extract scan parameters from all recipes.

    Parameters
    ----------
    File : Path
        RCP file imported from GetFile().

    Returns
    -------
    Output all scan parameters in console or .txt or .csv file.

    """
    nl = '\n'
    tab = '\t'
    rp = 'RecipePoint'
    acq = 'AcquisitionSettings'
    with olef.OleFileIO(File) as ole:
        toms = stream_unpacker(ole, 'NoOfTomoDataSets', "<i")[0]
        No_Rcps = f"Number of recipes:{tab}{toms}{nl}"
        print(No_Rcps)

        for x in range(toms):
            # print(ole.listdir())
            param_dict = {}
            param_dict['F_name'] = f"File:{tab}{File.name}{nl}"

            Rname = stream_unpacker(ole, f'{rp}{x}/PointName', 0).decode("ascii")[:-1]
            param_dict['Rcp_name'] = f"Recipe:{tab}{Rname}{nl}"

            date = datetime.strptime(stream_unpacker(ole, 'TimeStamp', 0)[:10].decode("ascii"), '%Y-%m-%d')
            param_dict['Date'] = f"Date:{tab}{date:%d/%m/%Y}{nl}"

            time = datetime.strptime(stream_unpacker(ole, 'TimeStamp', 0)[11:17].decode("ascii"), '%H%M%S')
            param_dict['Time'] = f"Time:{tab}{time:%H:%M:%S}{nl}"

            volts = stream_unpacker(ole, f'{rp}{x}/{acq}/SrcVoltage', "<f")[0]
            param_dict['kV'] = f"kV:{tab}{volts}{nl}"

            Watts = stream_unpacker(ole, f'{rp}{x}/{acq}/SrcPower', "<f")[0]
            param_dict['Power'] = f"Power (W):{tab}{Watts}{nl}"

            if volts == 0 or Watts == 0:
                curr = 0.0
            else:
                curr = round((Watts * 1000) / volts, 1)
            param_dict['uA'] = f"uA:{tab}{curr}{nl}"

            proj = stream_unpacker(ole, f'{rp}{x}/{acq}/TotalImages', "<i")[0]
            param_dict['imgs'] = f"Projections taken:{tab}{proj}{nl}"

            End_angle = stream_unpacker(ole, f'{rp}{x}/{acq}/EndAngle', "<f")[0]
            Sta_angle = stream_unpacker(ole, f'{rp}{x}/{acq}/StartAngle', "<f")[0]
            param_dict['rot'] = f"Rotation (deg):{tab}{round(abs(End_angle) + abs(Sta_angle))}{nl}"

            exp = stream_unpacker(ole, f'{rp}{x}/{acq}/ExpTime', "<f")[0]
            param_dict['exposure'] = f"Exposure (secs):{tab}{round(exp, 2)}{nl}"

            mag = stream_unpacker(ole, f'{rp}{x}/MagStr', 0).decode("ascii")[:-1]
            param_dict['objlens'] = f"Objective lens:{tab}{mag}{nl}"

            filt = stream_unpacker(ole, f'{rp}{x}/{acq}/SourceFilterName', 0).decode("ascii")[:-257]
            param_dict['XFilt'] = f"Filter:{tab}{filt}{nl}"

            binning = stream_unpacker(ole, f'{rp}{x}/{acq}/Binning', "<i")[0]
            param_dict['D_bin'] = f"Binning:{tab}{binning}{nl}"

            f_avg = stream_unpacker(ole, f'{rp}{x}/{acq}/FramesPerImage', "<i")[0]
            param_dict['fr_avg'] = f"Frame Averaging:{tab}{f_avg}{nl}"

            bh = stream_unpacker(ole, f'{rp}{x}/ReconSettings/BeamHardening', "<f")[0]
            param_dict['beam_h'] = f"Beam Hardening:{tab}{round(bh, 2)}{nl}"

            srctosample = stream_unpacker_from(ole, f'{rp}{x}/{acq}/InitialPositions', "<f", 16)[0]
            param_dict['Src'] = f"Src-Obj distance (mm):{tab}{round(abs(srctosample), 2)}{nl}"

            dettosample = stream_unpacker_from(ole, f'{rp}{x}/{acq}/InitialPositions', "<f", 20)[0]
            param_dict['Det'] = f"Det-Obj distance (mm):{tab}{round(dettosample, 2)}{nl}"

            CCD_size = stream_unpacker(ole, f'{rp}{x}/{acq}/CCDPixelSize', "<f")[0]
            geom_mag = (abs(srctosample) + dettosample) / abs(srctosample)
            Px_size = (CCD_size/float(mag[:-1])/geom_mag) * binning
            param_dict['Vox_size'] = f"Voxel size (um):{tab}{round(Px_size, 2)}{nl}"

            detrad = Px_size * (2042/binning)
            slant = math.sqrt(((abs(srctosample) + dettosample)**2)+(detrad**2))
            param_dict['cone'] = f"Cone angle (deg):{tab}{round(2 * math.asin(detrad/slant), 2)}{nl}"

            x_pos = stream_unpacker_from(ole, f'{rp}{x}/{acq}/InitialPositions', "<f", 0)[0]
            param_dict['x_ax'] = f"X position (um):{tab}{round(x_pos, 2)}{nl}"

            y_pos = stream_unpacker_from(ole, f'{rp}{x}/{acq}/InitialPositions', "<f", 4)[0]
            param_dict['y_ax'] = f"Y position (um):{tab}{round(y_pos, 2)}{nl}"

            z_pos = stream_unpacker_from(ole, f'{rp}{x}/{acq}/InitialPositions', "<f", 8)[0]
            param_dict['z_ax'] = f"Z position (um):{tab}{round(z_pos, 2)}{nl}"

            if stream_unpacker(ole, f'{rp}{x}/{acq}/AcqModeString', 0).decode("ascii")[:-245] == 'Tomography Wide':
                if stream_unpacker(ole, f'{rp}{x}/AutoStitchSettings/Enabled', "?")[0] is True:
                    param_dict['Acquisition mode'] = f"Acquisition mode:{tab}Wide Stitch{nl}"
                    segs = stream_unpacker(ole, f'{rp}{x}/AutoStitchSettings/NumSegments', "<i")[0]
                    param_dict['Segments'] = f"No. of segments:{tab}{segs}{nl}"
                else:
                    param_dict['Acquisition mode'] = f"Acquisition mode:{tab}Wide{nl}"
                    param_dict['Segments'] = f"No. of segments:{tab}1{nl}"
            else:
                if stream_unpacker(ole, f'{rp}{x}/AutoStitchSettings/Enabled', "?")[0] is True:
                    param_dict['Acquisition mode'] = f"Acquisition mode:{tab}Stitch{nl}"
                    segs = stream_unpacker(ole, f'{rp}{x}/AutoStitchSettings/NumSegments', "<i")[0]
                    param_dict['Segments'] = f"No. of segments:{tab}{segs}{nl}"
                else:
                    param_dict['Acquisition mode'] = f"Acquisition mode:{tab}Normal{nl}"
                    param_dict['Segments'] = f"No. of segments:{tab}1{nl}"

            HART = stream_unpacker(ole, f'{rp}{x}/{acq}/VariableAngleMode', "<i")[0]
            if HART == 1:
                param_dict['HART'] = f"HART:{tab}Enabled{nl}"
            else:
                param_dict['HART'] = f"HART:{tab}Disabled{nl}"

            V_exp = stream_unpacker(ole, f'{rp}{x}/{acq}/VariableExposureTimeMode', "<i")[0]
            if V_exp == 1:
                param_dict['V_exp'] = f"Variable exposure:{tab}Enabled{nl}"
            else:
                param_dict['V_exp'] = f"Variable exposure:{tab}Disabled{nl}"

            for value in param_dict.values():
                print(value.strip())

            os.chdir(File.parent)
            if out_file == 1:
                no_out = 0
                with open(f'{File.stem}_{Rname}_{File.suffix[1:]}.txt', 'w') as filehandle:
                    for value in param_dict.values():
                        filehandle.write(value)

            elif out_file == 2:
                no_out = 0
                with open(f'{File.stem}_{Rname}_{File.suffix[1:]}.csv', 'w') as csv_write:
                    for value in param_dict.values():
                        csv_write.write(value.replace('\t', ','))

            elif out_file == 3:
                no_out = 1
                print('\n')

    ole.close()
    if no_out == 1:
        valid = ["Y", "N", "y", "n"]
        while True:
            try:
                end = input("Would you like to open another file? (Y/N) ")
            except ValueError:
                print("Invalid input")
                continue
            if end not in valid:
                print("Invalid input")
                continue
            elif end.upper() == "Y":
                GetFile(3)
                break
            elif end.upper() == "N":
                break


GetFile()

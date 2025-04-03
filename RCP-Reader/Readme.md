# RCP/TXM/TXRM Metadata Extractor

This Python application provides a graphical user interface (GUI) to extract metadata from RCP, TXM, and TXRM files. It supports exporting the extracted metadata as a tab-delimited TXT file, a CSV file, or displaying it directly in the console.

## Features

* **GUI-based file selection:** Easily select RCP/TXM/TXRM files through a file dialog.
* **Output format selection:** Choose between TXT, CSV, or console output.
* **Metadata extraction:** Extracts various metadata parameters, including:
    * Date and time of acquisition
    * Voltage (kV) and current (uA)
    * Power (W)
    * Number of projections taken
    * Rotation angle (degrees)
    * Exposure time (seconds)
    * Objective lens magnification
    * Filter name
    * Voxel size (um)
    * Cone angle (degrees)
    * Binning
    * Frame averaging
    * Beam hardening
    * Source-object and detector-object distances (mm)
    * X, Y, and Z positions (um)
    * Acquisition mode
    * Number of segments (for stitched acquisitions)
    * HART and variable exposure mode (for TXRM/RCP files)
    * Recipe information from RCP files.
* **Error handling:** Robust error handling for file reading and data extraction.

## Requirements

* Python 3.x
* PyQt5
* olefile

To install the required packages, you can use pip:

```bash
pip install PyQt5 olefile
```

## Usage
1. Clone or download the repository.
2. Run the RCP_Metadata_Reader.py script:
```bash
python RCP_Metadata_Reader.py
```

3. The GUI window will appear.
4. Click the "Open" button to select an RCP/TXM/TXRM file.
5. Choose the desired output format (TXT, CSV, or Console).
6. Click "OK" to extract and output the metadata.
7. If you choose TXT or CSV, the file will be saved in the same directory as the input file. If you choose console, the metadata will be printed to the terminal.
8. The application will close itself automatically.

## File Format Support
* RCP: Recipe files, can contain multiple microCT recipes.
* TXM: ZEISS MicroCT reconstructed data files.
* TXRM: ZEISS MicroCT raw data files.

## Functionality Breakdown
* MetadataExtractorGUI: Handles the GUI creation and user interaction.
* stream_unpacker, stream_unpacker_from: Helper functions for extracting data from OLE streams.
* get_versa_projections, get_versa_rotation, get_versa_exposure, get_versa_filter, get_versa_src_dist, get_versa_det_dist, get_versa_acq_mode: Functions to extract specific metadata from TXM/TXRM files.
* _get_acq_mode: Determines the acquisition mode based on stitching and acquisition mode values.
* extract_metadata: Orchestrates the metadata extraction process.
* extract_common_data, extract_recipe_data: Extracts metadata specific to TXM/TXRM and RCP files, respectively.
* print_or_write_metadata: Outputs the extracted metadata to a file or console.
* main: Entry point for the application.

## Author
Daniel Bribiesca Sykes (<daniel.bribiescasykes@glasgow.ac.uk>)

## Version
3.2.0

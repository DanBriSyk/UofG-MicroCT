"""
Batch XRM to TIFF/BMP converter.

To do:
    Improve GUI.

"""

# Author: Daniel Bribiesca Sykes <daniel.bribiescasykes@glasgow.ac.uk>
# Version: 2.1.0

from pathlib import Path
import matplotlib.pyplot as plt
import numpy as np
import olefile as olef
import struct
from skimage import io, exposure
import tkinter as tk
from tkinter.filedialog import askdirectory


def ChooseDirectory():
    """
    Get user to select directory containing XRM files.

    Returns
    -------
    directory : Path
        Directory to recursively search for XRM files.

    """
    root = tk.Tk()
    root.withdraw()
    directory = askdirectory()
    root.destroy()
    return directory


def ole_extract(ole, stream, datatype):
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
    nev : unpacked data
        Unpacked parameter from ole file.

    """
    if ole.exists(stream):
        data = ole.openstream(stream).read()
        nev = struct.unpack(datatype, data)
        return nev


def BatchXRMconv(Dir):
    """
    Iterate through directory to find XRM files, extract iamge dimensions and export to tiff.

    Parameters
    ----------
    Dir : Path
        XRM containing directory.

    Returns
    -------
    Tiff files of all XRM files found in Dir.

    """
    for f in Dir.rglob('*.xrm'):
        with olef.OleFileIO(f) as ole:
            n_cols = int(ole_extract(ole, 'ImageInfo/ImageWidth', '<I')[0])
            n_rows = int(ole_extract(ole, 'ImageInfo/ImageHeight', '<I')[0])
            absdata = np.empty((n_cols, n_rows), dtype=np.uint16)

            try:
                imgdata = ole_extract(ole, 'ImageData1/Image1', "<{}h".format(n_cols*n_rows))
                absdata[:, :] = np.reshape(imgdata, (n_cols, n_rows), order='F')

                plt.figure()
                plt.imshow(absdata, cmap='gray')
                plt.colorbar()

                vmin, vmax = np.percentile(absdata, (0.1, 99.9))
                rescale_img = exposure.rescale_intensity(absdata, in_range=(vmin, vmax), out_range=np.uint16)

                plt.figure()
                plt.imshow(rescale_img, cmap='gray')
                plt.colorbar()

                io.imsave(str(f.parent) + "/" + f.stem + ".tiff", rescale_img)

            except Exception:
                print("Can't read " + f.name)

        ole.close()


ImportFolder = Path(ChooseDirectory())
BatchXRMconv(ImportFolder)

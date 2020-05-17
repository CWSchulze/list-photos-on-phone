#!/usr/bin/env python
# A command-line script for the Windows OS to find the photos that haven't been
# copied from a connected iPhone to the local machine yet.

import os
import sys
import time
from argparse import ArgumentParser
from win32com.shell import shell, shellcon
import pywintypes
import pythoncom
from win32com import storagecon
import logging
logger = logging.getLogger("list-photos")
logging.basicConfig()

__author__ = "David Blume"
__copyright__ = "Copyright 2014, David Blume"
__license__ = "http://www.wtfpl.net/"


def process_photos(target_folder, folder, overwrite):
    """
    Copy photos, overwrite if active.
    :param target_folder: local folder where the files are copied to.
    :param folder: The PIDL of the folder to walk.
    :param overwrite: overwrite existing files
    """
    for pidl in folder.EnumObjects(0, shellcon.SHCONTF_NONFOLDERS):
        name = folder.GetDisplayNameOf(pidl, shellcon.SHGDN_FORADDRESSBAR)
        dirname = os.path.dirname(name)
        basename, ext = os.path.splitext(os.path.basename(name))
        # List only the images that are newer.
        if ext.endswith("JPG"):
            if basename.startswith('IMG_E'):
                logger.warning(f'ignoring {basename}')
            elif not os.path.isfile(os.path.join(target_folder,os.path.split(name)[1])):
                logger.info(f'copying {basename}')
                data = b''
                for chunk in stream_file_content(folder, pidl):
                    data += chunk
                open(os.path.join(target_folder,os.path.split(name)[1]),'wb').write(data)
                #photo_dict[dirname].append(name)
            else:
                logger.debug(f'{basename} is not overwritten')


def walk_dcim_folder(target_folder, dcim_pidl, parent, overwrite):
    """
    Iterates all the subfolders of the iPhone's DCIM directory, copying
    photos to the target folder.

    :param target_folder: local folder where the files are copied to.
    :param dcim_pidl: A PIDL for the iPhone's DCIM folder
    :param parent: The parent folder of the PIDL
    :param overwrite: overwrite existing files
    """
    dcim_folder = parent.BindToObject(dcim_pidl, None, shell.IID_IShellFolder)
    for pidl in dcim_folder.EnumObjects(0, shellcon.SHCONTF_FOLDERS):
        folder = dcim_folder.BindToObject(pidl, None, shell.IID_IShellFolder)
        name = folder.GetDisplayNameOf(pidl, shellcon.SHGDN_FORADDRESSBAR)
        logger.info(f'working on folder {name}')
        process_photos(target_folder, folder, overwrite)



def get_dcim_folder(device_pidl, parent):
    """
    Tries to find an iPhone by searching the pidl for the path
    "Internal Storage\DCIM".
    :param device_pidl: A candidate Windows PIDL for the iPhone
    :param parent: The parent folder of the PIDL
    """
    device_name = parent.GetDisplayNameOf(device_pidl, shellcon.SHGDN_NORMAL)
    name = None
    pidl = None

    folder = parent.BindToObject(device_pidl, None, shell.IID_IShellFolder)
    try:
        top_dir_name = ""
        for pidl in folder.EnumObjects(0, shellcon.SHCONTF_FOLDERS):
            top_dir_name = folder.GetDisplayNameOf(pidl, shellcon.SHGDN_NORMAL)
            break  # Only want to see the first folder.
        if top_dir_name != "Internal Storage":
            return None, None, device_name
    except pywintypes.com_error:
        return None, None, device_name  # No problem, must not be an iPhone

    folder = folder.BindToObject(pidl, None, shell.IID_IShellFolder)
    for pidl in folder.EnumObjects(0, shellcon.SHCONTF_FOLDERS):
        name = folder.GetDisplayNameOf(pidl, shellcon.SHGDN_NORMAL)
        break  # Only want to see the first folder.
    if name != "DCIM":
        logger.warning("%s's '%s' has '%s', not a 'DCIM' dir." %
                (device_name, top_dir_name, name))
        return None, None, device_name

    return pidl, folder, device_name


def get_computer_shellfolder():
    """
    Return the local computer's shell folder.
    """
    desktop = shell.SHGetDesktopFolder()
    for pidl in desktop.EnumObjects(0, shellcon.SHCONTF_FOLDERS):
        display_name = desktop.GetDisplayNameOf(pidl, shellcon.SHGDN_NORMAL)
        if display_name in ("Computer", "This PC", "Dieser PC"):
            return desktop.BindToObject(pidl, None, shell.IID_IShellFolder)
    return None

def stream_file_content(folder, pidl, buffer_size=8192):
    istream = folder.BindToStorage(pidl, None, pythoncom.IID_IStream)
    while True:
        contents = istream.Read(buffer_size)
        if contents:
            yield contents
        else:
            break

def main(overwrite_files):
    """
    Find a connected iPhone, and print the paths to images on it.
    :param overwrite_files: overwrite files
    """
    start_time = time.time()
    localdir = os.path.dirname(__file__)

    # Find the iPhone in the virtual folder for the local computer.
    computer_folder = get_computer_shellfolder()
    for pidl in computer_folder:
        # If this is the iPhone, get the PIDL of its DCIM folder.
        dcim_pidl, parent, iphone_name = get_dcim_folder(pidl, computer_folder)
        if dcim_pidl is not None:
            walk_dcim_folder(localdir, dcim_pidl, parent, overwrite_files)

    logger.info("Done. That took %1.2fs." % (time.time() - start_time))


if __name__ == '__main__':
    parser = ArgumentParser()
    parser.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument("-o", "--overwrite", action="store_true")
    args = parser.parse_args()
    logger.setLevel(logging.INFO if args.verbose else logging.ERROR)
    main(args.overwrite)

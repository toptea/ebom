"""
Operating System Methods
"""

from pathlib import Path
from glob import glob

import subprocess
import time
import os


def find_vault_path():
    username = os.getlogin().upper()
    for drive in 'CDEFGHIJKLMNOPQRSTUVWXYZ':
        path = Path(drive + ':/BC-Workspace/' + username + '/M-Balmoral,D-AGR/Projects')
        if os.path.exists(str(path)):
            return path
    for drive in 'CDEFGHIJKLMNOPQRSTUVWXYZ':
        path = Path(drive + ':/BC-Workspace/' + username + '/M-Bally-01,D-AGR/Projects')
        if os.path.exists(str(path)):
            return path


def find_inventor_path():
    for drive in 'CDEFGHIJKLMNOPQRSTUVWXYZ':
        path = Path(drive + ':/Program Files/Autodesk/Inventor 2016/Bin/Inventor.exe')
        if os.path.exists(str(path)):
            return path


def find_export_path():
    username = os.getlogin().upper()
    path = Path('C:/Users/' + username + '/Desktop/CAD/')
    return path


def find_path(partcode, filetype):
    """find Inventor file path

    Parameters
    ----------
    partcode : str
        AGR part number usually in 'AGR0000-000-00' format
    filetype : str
        inventor file type ('ipt', 'iam' or 'idw')

    Returns
    -------
    Path : obj
        Path object from Python pathlib module
    """
    client = '*'
    project = partcode[0:7]
    section = partcode[8:11]
    file = partcode + '.' + filetype
    paths = glob(str(find_vault_path() / client / project / section / file))

    # if len(paths) == 0:
    #     paths = glob(str(INVENTOR_DIR / '**' / file), recursive=True)
    if len(paths) == 0:
        pass
        # print('Unable to find ' + partcode)
    if len(paths) > 1:
        print('Warning - multiple files found. Use first path in list')
    if len(paths) > 0:
        return Path(paths[0])


def find_paths(partcodes, filetype):
    """find Inventor file paths

    Parameters
    ----------
    partcodes : :obj:`list` of :obj:`str`
        AGR part numbers usually in 'AGR0000-000-00' format
    filetype : str
        inventor file type ('ipt', 'iam' or 'idw')

    Returns
    -------
    Path : :obj:`list` of obj
        Path object from Python pathlib module
    """
    paths = []
    for partcode in partcodes:
        path = find_path(partcode, filetype)
        if path is not None:
            paths.append(path)
    return paths


def start_inventor():
    """Open Inventor

    Inventer must be active for the COM API to work.
    """
    if 'Inventor.exe' not in os.popen("tasklist").read():
        subprocess.Popen(find_inventor_path())
        time.sleep(10)

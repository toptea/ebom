"""
Operating System Methods
"""

from pathlib import Path
from glob import glob

import textwrap
import subprocess
import time
import os


def find_vault_path():
    username = os.getlogin().upper()
    for drive in 'CDEFGHIJKLMNOPQRSTUVWXYZ':
        arbroath_path = Path(drive + ':/BC-Workspace/' + username + '/M-Balmoral,D-AGR/Projects')
        if os.path.exists(str(arbroath_path)):
            return arbroath_path
        ireland_path = Path(drive + ':/BC-Workspace/' + username + '/M-Bally-01,D-AGR/Projects')
        if os.path.exists(str(ireland_path)):
            return ireland_path
    else:
        return Path('C:/')


def check_vault_path(path):
    if os.path.exists(str(path)):
        return True
    else:
        err = textwrap.dedent(
            """
            Unable to find Meridian Vault path location.
            -------------------------------------------------------------------
            Please make sure the path matches the following pattern:
            1) [drive]:/BC-Workspace/[username]/M-Balmoral,D-AGR/Projects
            2) [drive]:/BC-Workspace/[username]/M-Bally-01,D-AGR/Projects
            """
        )
        raise FileNotFoundError(err)


def find_inventor_path():
    for drive in 'CDEFGHIJKLMNOPQRSTUVWXYZ':
        path = Path(drive + ':/Program Files/Autodesk/Inventor 2016/Bin/Inventor.exe')
        if os.path.exists(str(path)):
            return path
    else:
        return Path('C:/')


def check_inventor_path(path):
    if os.path.exists(str(path)):
        return path
    else:
        err = textwrap.dedent(
            """
            Unable to find Autodesk Inventor path location.
            -------------------------------------------------------------------
            Please make sure the path matches the following pattern:
            1) [drive]:/Program Files/Autodesk/Inventor 2016/Bin/Inventor.exe
            """
        )
        raise FileNotFoundError(err)


def find_export_path():
    return Path(os.getcwd())

    # username = os.getlogin().upper()
    # path = Path('C:/Users/' + username + '/Desktop/')
    # if os.path.exists(str(path)):
    #    return path
    # else:
    #    return Path(os.getcwd())


def find_path(partcode, filetype, is_sub_assembly=False):
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

    if len(paths) == 0 and not is_sub_assembly:
        err = textwrap.dedent(
            """
            Unable to find '{}' path location.
            -------------------------------------------------------------------
            Please make sure the file are synced in Meridian and check for any typos.
            """
        ).format(partcode)
        raise FileNotFoundError(err)
    if len(paths) > 1:
        print("Warning - There's multiple paths. Use first path in list")
    if len(paths) == 1:
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
        path = find_path(partcode, filetype, is_sub_assembly=True)
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

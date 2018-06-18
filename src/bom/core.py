from bom.program import encompix
from bom.program import data

# from bom.program import meridian

import pandas as pd
import csv


def main(command, assembly, close_file='never', open_model=True,
         recursive=True, output_report=True, output_import=True,
         update_parent_revision=True, update_vendor_id=True):
    """Export the BOM from Inventor, Promise or from Cooperation

    The current workflow are:
    1) Inventor/Promise/Cooperation - Create bom, ibom & ebom dataframes
    2) Meridian - Update drawing revision column
    3) Encompix - Update parent revision column
    4) Encompix - Exclude sections with the same revision
    5) Encompix - Update vendor id column
    6) Python Core - Save report and import file

    Parameters
    ----------
    command : str
        Choose between 'mechanical', 'electrical' or 'cooperation
        sub-command. Extract the bom from the respective program.
    assembly : str
        The assembly number you wish to import eg AGR1338-025-00.
        If None, use the active drawing opened in Inventor instead.
    close_file : str
        Choose whether to close idw or/and iam file when finished
        (without saving). Choose between ['never', 'iam', 'idw', 'both']
    open_model : bool
        Sometime there are missing description on the parts list. To fix
        this, the assembly model needs to be opened as well.
    update_parent_revision : bool
        update assembly revision on the ebom
    recursive : bool
        Include sub assembly BOMs.
    output_report : bool
        Save Report file (partcode.xlsx)
    output_import : bool
        Save Encompix import file (partcode.csv)
    """

    # 1a) Mechanical BOMs
    if command == 'mechanical':
        app = data.Inventor(
            assembly=assembly,
            close_file=close_file,
            open_model=open_model,
            recursive=recursive
        )
        bom = app.load_bom()
        ibom = app.create_indented_bom(bom)
        ebom = app.create_ebom(bom)

    # -------------------------------------------------------------------------
    # TODO - 1b) Electrical BOMs
    # if command == 'electrical':
    #     app = data.Promise()
    #     bom = app.load_bom()
    #     ibom = app.create_indented_bom(bom)
    #     ebom = app.create_ebom(bom)
    #
    # TODO - 1c) Cooperation BOMs
    # if command == 'cooperation':
    #     app = data.Cooperation()
    #     bom = app.load_bom()
    #     ibom = app.create_indented_bom(bom)
    #     ebom = app.create_ebom(bom)
    #
    # TODO - 2) Meridian - Update drawing revision column
    # if command != 'electrical' and arg.update_drawing:
    #     drev = meridian.load_drawing_revision(ebom)
    #     ebom = meridian.update_drawing_revision(ebom, drev)
    # -------------------------------------------------------------------------

    # 3) Encompix - Update parent revision column
    if update_parent_revision:
        prev = encompix.load_parent_revision(ebom)
        ebom = encompix.update_parent_revision(ebom, prev)

    # -------------------------------------------------------------------------
    # TODO - 4) Encompix - Exclude sections with the same revision
    # if arg.exclude_same_revision:
    #     ebom = app.encompix.exlude_same_revision(ebom)
    # -------------------------------------------------------------------------

    # 5) Encompix - Update vendor id column
    if update_vendor_id:
        vendor = encompix.load_vendor_id()
        ebom = encompix.update_vendor_id(ebom, vendor)

    # 6a) Save report file
    if output_report:
        save_report_file(assembly, bom, ibom, ebom, prev)

    # 6b) Save Encompix Import csv
    if output_import:
        save_import_file(assembly, ebom)

    # -------------------------------------------------------------------------
    # TODO - 6c) Save Encompix Item csv
    # if arg.output_item:
    #    save_item_file()
    # -------------------------------------------------------------------------


def save_report_file(assembly, bom, indented_bom, ebom, prev):
    """Save Report file as partcode.xlsx

    Parameters
    ----------
    assembly : str
        assembly number
    bom : obj
        bom dataframe
    indented_bom : obj
        indented bom dataframe
    ebom : obj
        encompix import bom dataframe
    """
    writer = pd.ExcelWriter(assembly + '.xlsx', engine='xlsxwriter')
    workbook = writer.book
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': False,
        'fg_color': '#DCE6F1',
        'border': 1
    })

    _create_bom_worksheet(bom, writer, header_format)
    _create_indented_bom_worksheet(indented_bom, writer, header_format)
    _create_import_file_worksheet(ebom, writer, header_format)
    _create_new_revision_worksheet(prev, writer, header_format)

    writer.close()


def _create_bom_worksheet(df, writer, header_format):
    """Create and format bom worksheet

    Parameters
    ----------
    df : obj
        bom dataframe
    writer : obj
        xlsxwriter object
    header_format : dict
        1st row header format properties
    """
    df.to_excel(
        writer,
        sheet_name='bom',
        index=False,
        header=False,
        startrow=1
    )

    worksheet = writer.sheets['bom']
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    worksheet.set_column(0, 0, 16)  # Assembly
    worksheet.set_column(1, 1, 26)  # Assembly_Name
    worksheet.set_column(5, 5, 16)  # Dwg_No
    worksheet.set_column(6, 6, 34)  # Component


def _create_indented_bom_worksheet(df, writer, header_format):
    """Create and format indented bom worksheet

    Parameters
    ----------
    df : obj
        ibom dataframe
    writer : obj
        xlsxwriter object
    header_format : dict
        1st row header format properties
    """
    df.to_excel(
        writer,
        sheet_name='indented_bom',
        index=False,
        header=False,
        startrow=1,
    )

    worksheet = writer.sheets['indented_bom']
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    worksheet.set_column(2, 2, 21)  # Dwg_No
    worksheet.set_column(3, 3, 60)  # Component


def _create_new_revision_worksheet(df, writer, header_format):
    """Create and encompix new revision worksheet

    Parameters
    ----------
    df : obj
        prev dataframe
    writer : obj
        xlsxwriter object
    header_format : dict
        1st row header format properties
    """
    df.to_excel(
        writer,
        sheet_name='new_revision',
        index=False,
        startrow=1,
        header=False
    )

    worksheet = writer.sheets['new_revision']
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    worksheet.set_column(0, 0, 16)  # Assembly
    worksheet.set_column(1, 1, 26)  # Assembly_Name
    worksheet.set_column(2, 1, 15)  # Revision


def _create_import_file_worksheet(df, writer, header_format):
    """Create and format encompix import bom worksheet

    Parameters
    ----------
    df : obj
        ebom dataframe
    writer : obj
        xlsxwriter object
    header_format : dict
        1st row header format properties
    """
    df.to_excel(
        writer,
        sheet_name='import',
        index=False,
        startrow=1,
        header=False
    )

    worksheet = writer.sheets['import']
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

        # worksheet.set_column(0, 0, 16)  # Assembly
        # worksheet.set_column(1, 1, 26)  # Assembly_Name


def save_import_file(assembly, ebom):
    """Save Encompix import file as partcode.csv

    Parameters
    ----------
    assembly : str
        assembly number
    ebom : obj
        encompix import bom dataframe
    """
    ebom.to_csv(
        assembly + '.csv',
        index=False,
        quoting=csv.QUOTE_NONNUMERIC
    )

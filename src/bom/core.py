from bom.program import inventor

import pandas as pd
import csv


def main(command, assembly, close_file='never', open_model=True,
         recursive=True, output_report=True, output_import=True):

    if assembly is None:
        assembly = inventor.get_active_assembly_partcode()

    # 1a) Mechanical BOM
    # 2a) Indented BOM
    # 3a) Encompix Import BOM
    if command == 'mechanical':
        bom = inventor.load_bom(
            assembly=assembly,
            close_file=close_file,
            open_model=open_model,
            recursive=recursive
        )
        ibom = inventor.create_indented_bom(bom)
        ebom = inventor.create_ebom(bom)

    """
    # 1b) Electrical BOM
    # 2b) Indented BOM
    # 3b) Encompix Import BOM
    if arg.type == 'electrical':
        bom = electrical.load_bom(
            report=arg.report
        )
        ibom = electrical.create_indented_bom(bom)
        ebom = electrical.create_ebom(bom)

    # 1c) Cooperation BOM
    # 2c) Indented BOM
    # 3c) Encompix Import BOM
    if arg.type == 'cooperation':
        bom = app.promise.load_bom(
            assembly=arg.assembly,
            recursive=arg.recursive
        )
        ibom = app.promise.create_indented_bom(bom)
        ebom = app.promise.create_ebom(bom)

    # 4) Update drawing revision column
    if command != 'electrical' and arg.update_drawing:
        drev = app.meridian.load_drawing_revision(ebom)
        ebom = app.meridian.update_drawing_revision(ebom, drev)

    # 5) Update parent revision column
    if arg.update_parent_revision:
        prev = app.encompix.load_parent_revision(ebom)
        ebom = app.encompix.update_parent_revision(ebom, prev)

    if arg.exlude_same_revision:
        ebom = app.encompix.exlude_same_revision(ebom)

    # 6) Update vendor id column
    if arg.update_vendor:
        vendor = app.encompix.load_vendor_id(ebom)
        ebom = app.encompix.update_vendor_id(ebom, vendor)
    """
    # 7) Save file
    if output_report:
        save_report_file(assembly, bom, ibom, ebom)

    if output_import:
        save_import_file(assembly, ebom)


    # if arg.output_item:
    #    save_item_file()


def save_report_file(assembly, bom, indented_bom, ebom):
    writer = pd.ExcelWriter(assembly + '.xlsx', engine='xlsxwriter')
    workbook = writer.book
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': False,
        'fg_color': '#DCE6F1',
        'border': 1
    })
    # -------------------------------------------------------------------------
    bom.to_excel(
        writer,
        sheet_name='bom',
        index=False,
        header=False,
        startrow=1
    )

    worksheet = writer.sheets['bom']
    for col_num, value in enumerate(bom.columns.values):
        worksheet.write(0, col_num, value, header_format)

    worksheet.set_column(0, 0, 16)  # Assembly
    worksheet.set_column(1, 1, 26)  # Assembly_Name
    worksheet.set_column(5, 5, 16)  # Dwg_No
    worksheet.set_column(6, 6, 34)  # Component
    # -------------------------------------------------------------------------
    indented_bom.to_excel(
        writer,
        sheet_name='indented_bom',
        index=False,
        header=False,
        startrow=1,
    )

    worksheet = writer.sheets['indented_bom']
    for col_num, value in enumerate(indented_bom.columns.values):
        worksheet.write(0, col_num, value, header_format)

    worksheet.set_column(2, 2, 21)  # Dwg_No
    worksheet.set_column(3, 3, 60)  # Component
    """
    # -------------------------------------------------------------------------
    nr = load_new_revision_table(df)
    nr.to_excel(
        writer,
        sheet_name='new_revision',
        index=False,
        startrow=1,
        header=False
    )

    worksheet = writer.sheets['new_revision']
    for col_num, value in enumerate(nr.columns.values):
        worksheet.write(0, col_num, value, header_format)

    worksheet.set_column(0, 0, 16)  # Assembly
    worksheet.set_column(1, 1, 26)  # Assembly_Name
    # -------------------------------------------------------------------------
    """
    ebom.to_excel(
        writer,
        sheet_name='import',
        index=False,
        startrow=1,
        header=False
    )

    worksheet = writer.sheets['import']
    for col_num, value in enumerate(ebom.columns.values):
        worksheet.write(0, col_num, value, header_format)

    #worksheet.set_column(0, 0, 16)  # Assembly
    #worksheet.set_column(1, 1, 26)  # Assembly_Name
    # -------------------------------------------------------------------------
    writer.close()


def save_import_file(assembly, ebom):
    ebom.to_csv(
        assembly + '.csv',
        index=False,
        quoting=csv.QUOTE_NONNUMERIC
    )
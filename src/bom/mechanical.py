import pandas as pd
from utils import system
from utils import inventor
from utils import database
import numpy as np
import csv


def load_part_list(app, assembly, open_model, close_file, recursive):
    system.status('\n' + assembly)
    system.status('|_ opening idw')
    idw_path = system.find_path(assembly, 'idw')
    idw = inventor.Drawing(idw_path, app)

    df = idw.extract_part_list(lvl=1)

    if open_model and '' in df['Component'].unique():
        system.status('|_ opening iam')
        iam_path = system.find_path(assembly, 'iam')
        iam = inventor.Drawing(iam_path, app)
        idw = inventor.Drawing(idw_path, app)
        system.status('|_ extracting bom')
        df = idw.extract_part_list(lvl=1)
    else:
        system.status('|_ extracting bom')
        iam = None

    if close_file == 'idw' or close_file == 'both':
        system.status('|_ closing idw')
        idw.close()
    if (close_file == 'iam' or close_file == 'both') and iam is not None:
        system.status('|_ closing iam')
        iam.close()
    if recursive:
        df = load_sub_part_list(df, app, open_model, close_file)

    return df


def load_sub_part_list(df, app, open_model, close_file):
    lvl = 2
    while True:
        assemblies = df.loc[~df['Dwg_No'].isin(df['Assembly']), 'Dwg_No'].values
        iam_paths = system.find_paths(assemblies, 'iam')
        idw_paths = system.find_paths([path.stem for path in iam_paths], 'idw')

        if len(idw_paths) == 0:
            break

        for idw_path, iam_path in zip(idw_paths, iam_paths):
            assembly = idw_path.stem
            system.status('\n' + assembly)
            system.status('|_ opening idw')
            idw = inventor.Drawing(idw_path, app)
            rs = idw.extract_part_list(lvl)

            if open_model and '' in rs['Component'].unique():
                system.status('|_ opening iam')
                iam = inventor.Drawing(iam_path, app)
                idw = inventor.Drawing(idw_path, app)
                system.status('|_ extracting bom')
                rs = idw.extract_part_list(lvl)
            else:
                system.status('|_ extracting bom')
                iam = None
            if close_file == 'idw' or close_file == 'both':
                system.status('|_ closing idw')
                idw.close()
            if (close_file == 'iam' or close_file == 'both') and iam is not None:
                system.status('|_ closing iam')
                iam.close()
            df = df.append(rs, ignore_index=True)
        lvl += 1

    return df


def load_quantified_bom(df, is_indented=True):
    """
    Calculate the partial internal refence number for each partcode,
    then for each level, add all the assembly and the partcode irn together.
    Correct the partcode quantity by multiplying it with the assembly quantity.
    """

    dfs = {}
    df = _calc_partial_irn(df)
    for lvl in range(max(df['LVL']) + 1):
        if lvl == 0:
            dfs[0] = df.loc[df['LVL'] == 1, ['Assembly', 'Assembly_Name']]
            dfs[0] = dfs[0].drop_duplicates()
            dfs[0] = dfs[0].rename(
                columns={'Assembly': 'Dwg_No', 'Assembly_Name': 'Component'}
            )
            try:
                dfs[0]['ITEM'] = int('AGR1338-440-00'[8:11])
            except ValueError:
                dfs[0]['ITEM'] = ''
            dfs[0]['QTY'] = 1
            dfs[0]['LVL'] = 0
            dfs[0]['irn'] = 0
        else:
            df_assy = dfs[lvl - 1][['Dwg_No', 'QTY', 'irn']]
            df_assy = df_assy.rename(
                columns={'Dwg_No': 'Assembly', 'QTY': 'assy_qty', 'irn': 'assy_irn'}
            )
            dfs[lvl] = df.loc[df['LVL'] == lvl, ['ITEM', 'Assembly', 'Dwg_No', 'Component', 'QTY', 'LVL', 'irn']]
            dfs[lvl] = pd.merge(left=dfs[lvl], right=df_assy, how='inner', on='Assembly')
            dfs[lvl]['irn'] = dfs[lvl]['irn'] + dfs[lvl]['assy_irn']
            dfs[lvl]['QTY'] = dfs[lvl]['QTY'] * dfs[lvl]['assy_qty']
            dfs[lvl] = dfs[lvl][['ITEM', 'Dwg_No', 'Component', 'QTY', 'LVL', 'irn']]

    rs = pd.concat(dfs)
    rs = rs.sort_values('irn')
    rs = rs.reset_index(drop=True)

    if is_indented:
        rs['Dwg_No'] = rs.apply(_indent, column='Dwg_No', axis=1)
        rs['Component'] = rs.apply(_indent, column='Component', axis=1)

    rs = rs[['LVL', 'ITEM', 'Dwg_No', 'Component', 'QTY']]
    return rs


def _calc_partial_irn(df):
    """Partial internal reference number

    Internal Notes
    --------------
    a (int) counts the max number of digits in the item number.
    eg [101, 102, 103, ...] usually a=3.

    b (series) reverses the lvl field.
    eg 1->3, 2->2, 3->1

    c (series) calculate the number of zeros required.
    eg lvl_1 = 1 000 000, lvl_2 = 1 000, lvl_3 = 1.

    df['irn'] (series) is the partial irn derived from it's level and item number fields.
    eg 441 000 000, 441 101 000, 441 116 101

    To calculate the full irn, all the assemblies irn above the partcode needs to be added together.
    """

    a = max(np.floor(np.log10(df['ITEM']) + 1))
    b = max(df['LVL']) - df['LVL'] + 1
    c = 10 ** (a * b - a)

    df['irn'] = df['ITEM'] * c
    df['irn'] = df['irn'].astype(int)
    return df


def _indent(row, column):
    if row[column] is not np.nan:
        return (row['LVL']) * '    ' + row[column]


def load_new_revision_table(df):
    sql = (
        """
        SELECT
            "item-no" AS 'assembly',
            "item-desc" AS 'description',
            MAX("rev") AS 'Rev'
        FROM
            pub."i-bom-hdr"

        GROUP BY
            "item-no",
            "item-desc"

        ORDER BY
            "item-no"
        """
    )
    rs = database.query(sql, database_type='progress')

    df = df[['Assembly', 'Assembly_Name']].drop_duplicates()
    df = pd.merge(left=df, right=rs, how='left', left_on='Assembly', right_on='assembly')
    df['Rev'] = df['Rev'].fillna(0)
    df['Rev'] = df['Rev'].astype(int)
    df['Rev'] += df['Rev']
    df = df[['Assembly', 'Assembly_Name', 'Rev']]
    return df


def load_encompix_bom(df):
    rs = pd.DataFrame()
    rs['temp'] = df['Assembly']
    rs['Import?'] = 'YES'
    rs['Parent'] = df['Assembly']
    rs['Parent Item Description'] = df['Assembly_Name']
    rs['Parent Revision'] = 1
    rs['RoutingTemplate'] = ''
    rs['Default Oper'] = 10
    rs['Operation'] = 10
    rs['Item Number'] = df['Dwg_No']
    rs['Item Description'] = df['Component']
    rs['Quantity'] = df['QTY']
    rs['Req Type'] = df['Dwg_No'].apply(_is_manu)
    rs['Vendor ID'] = ''
    rs['Unit of Measure'] = 'EACH'
    rs['Length'] = ''
    rs['Width'] = ''
    rs['Product Code'] = 'PROJ'
    rs['Reference'] = ''
    rs['Detail'] = df['ITEM']
    rs.loc[rs['Req Type'] == 'M', 'Drawing'] = rs['Item Number']
    rs.loc[rs['Req Type'] == 'M', 'Sheet'] = 1
    rs['Notes'] = ''
    rs['Contribution '] = ''
    rs['Mfg ID'] = ''
    rs['Mfg Part#'] = ''
    rs['t-ibom.x-char1'] = ''
    rs['t-ibom.x-char2'] = ''
    rs['t-ibom.x-char3'] = ''
    rs['t-ibom.x-char4'] = ''
    rs['t-ibom.x-char5'] = ''
    rs['t-ibom.cInt1'] = ''
    rs['t-ibom.cInt2'] = ''
    rs['t-ibom.cInt3'] = ''
    rs['t-ibom.cInt4'] = ''
    rs['t-ibom.cInt5'] = ''
    rs['t-ibom.cDate1'] = ''
    rs['t-ibom.cDate2'] = ''
    rs['t-ibom.cDate3'] = ''
    rs['t-ibom.cDate4'] = ''
    rs['t-ibom.cDate5'] = ''
    rs['t-ibom.cDec1'] = ''
    rs['t-ibom.cDec2'] = ''
    rs['t-ibom.cDec3'] = ''
    rs['t-ibom.cDec4'] = ''
    rs['t-ibom.cDec5'] = ''
    rs['t-ibom.cLog1'] = ''
    rs['t-ibom.cLog2'] = ''
    rs['t-ibom.cLog3'] = ''
    rs['t-ibom.cLog4'] = ''
    rs['t-ibom.cLog5'] = ''
    del rs['temp']
    return rs


def _is_manu(row):
    """Check whether the part number is a manufacture of purchase part"""
    if len(row) >= 3:
        if all([row[i] not in '1234567890' for i in range(3)]):
            return 'M'
    else:
        return 'B'


def save_analysis(df, assembly):
    writer = pd.ExcelWriter(assembly + '.xlsx', engine='xlsxwriter')
    workbook = writer.book
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': False,
        'fg_color': '#DCE6F1',
        'border': 1
    })

    df.to_excel(
        writer,
        sheet_name='itemized_bom',
        index=False,
        header=False,
        startrow=1
    )

    worksheet = writer.sheets['itemized_bom']
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    worksheet.set_column(0, 0, 16)  # Assembly
    worksheet.set_column(1, 1, 26)  # Assembly_Name
    worksheet.set_column(5, 5, 16)  # Dwg_No
    worksheet.set_column(6, 6, 34)  # Component

    qb = load_quantified_bom(df)
    qb.to_excel(
        writer,
        sheet_name='quantified_bom',
        index=False,
        header=False,
        startrow=1,
    )

    worksheet = writer.sheets['quantified_bom']
    for col_num, value in enumerate(qb.columns.values):
        worksheet.write(0, col_num, value, header_format)

    worksheet.set_column(2, 2, 21)  # Dwg_No
    worksheet.set_column(3, 3, 60)  # Component

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
    writer.close()


def save_encompix_bom(df, assembly):
    rs = load_encompix_bom(df)
    rs.to_csv(
        assembly + '.csv',
        index=False,
        quoting=csv.QUOTE_NONNUMERIC
    )


def main(assembly, close_file, open_model=True, recursive=True,
         output_report=False, output_encompix=True):

    system.status('\nInput Parameters')
    system.status('|_ command = mechanical')
    system.status('|_ assembly =', assembly)
    system.status('|_ close_file =', close_file)
    system.status('|_ open_model =', open_model)
    system.status('|_ recursive =', recursive)
    system.status('|_ output_report =', output_report)
    system.status('|_ output_encompix =', output_encompix)

    # system.check_inventor_path(path)
    # system.check_vault_path(path)

    app = inventor.application()
    if assembly is None:
        idw = inventor.Drawing.via_active_document(app)
        iprop = idw.doc.PropertySets.Item('Inventor User Defined Properties')
        assembly = str(iprop.Item('Dwg_No')).strip()

    df = load_part_list(app, assembly, open_model, close_file, recursive)
    system.status('\nOutput')
    if output_report:
        system.status('|_ saving part list')
        save_analysis(df, assembly)
    if output_encompix:
        system.status('|_ saving encompix template')
        save_encompix_bom(df, assembly)

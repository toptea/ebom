from utils import system
from utils import inventor

import pandas as pd
import numpy as np


def get_active_assembly_partcode():
    app = inventor.application()
    idw = inventor.Drawing.via_active_document(app)
    iprop = idw.doc.PropertySets.Item('Inventor User Defined Properties')
    assembly = str(iprop.Item('Dwg_No')).strip()
    return assembly


def load_bom(assembly, close_file='never', open_model=True, recursive=True):
    """ Inventor Parts List

    Using Inventor COM API, open the assembly drawing,
    extract the parts list and return the data as a dataframe.

    Parameters
    ---------------
    assembly : str
        The asssembly number you wish to import.
        If left blank, use the active drawing instead.

    close_file : str
        - 'never' : leave everything open
        - 'idw': close the drawing when finished
        - 'iam': close the assembly model when finished
        - 'both': close everything when finished

    open_model : bool
        Sometimes the parts list have missing description.
        Fix this by opening the model as well.

    recursive : bool
        If true, include sub assembly part list as well.

    Returns
    ----------
    obj
        Pandas DataFrame

    """
    app = inventor.application()

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
        df = _load_sub_assembly_bom(df, app, open_model, close_file)

    return df


def _load_sub_assembly_bom(df, app, open_model, close_file):
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


def create_indented_bom(df, indent_format=True):
    """ Indented Bills of Material

    Show the multilevel BOM structure and quantity required for each part.

    Calculate the partial internal refence number for each partcode,
    then for each level, add all the assembly and the partcode irn together.
    Correct the partcode quantity by multiplying it with the assembly quantity.

    Parameters
    ---------------
    df : obj
        Dataframe from "bom.inventor.load_bom()" function
    indent_format : bool
        show bom structure by adding whitespace

    Returns
    ----------
    obj
        Pandas DataFrame

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

    if indent_format:

        def _indent(row, column):
            if row[column] is not np.nan:
                return (row['LVL']) * '    ' + row[column]

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


def create_ebom(df):
    """ Encompix Bills of Material

    Parameters
    ---------------
    df : obj
        Dataframe from "bom.inventor.load_bom()" function

    Returns
    ----------
    obj
        Pandas DataFrame

    """
    def _is_manu(row):
        """Check whether the part number is a manufacture of purchase part"""
        if len(row) >= 3:
            if all([row[i] not in '1234567890' for i in range(3)]):
                return 'M'
        else:
            return 'B'

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
    rs['Reference'] = '1'
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


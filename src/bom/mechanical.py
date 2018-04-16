import pandas as pd
from utils import system
from utils import inventor
import csv


def append_sub_assy_part_list(df, app, auto_close):
    lvl = 2
    while True:
        assembly_paths = _find_assembly_drawing_paths(df)
        if len(assembly_paths) == 0:
            break
        
        for assembly_path in assembly_paths:
            doc = inventor.Drawing(assembly_path, app)
            rs = doc.extract_part_list(lvl=lvl, auto_close=auto_close)
            df = df.append(rs)
        lvl += 1
    
    return df


def _find_assembly_drawing_paths(df):
    partcodes = df.loc[~df['Dwg_No'].isin(df['Assembly']),'Dwg_No']
    partcodes = partcodes.values
    iam_paths = system.find_paths(partcodes, 'iam')
    iam_partcodes = [path.stem for path in iam_paths]
    idw_paths = system.find_paths(iam_partcodes, 'idw')
    return idw_paths


def load_part_list(partcode):
    path = system.find_path(partcode, 'idw')
    print(path)
    app = inventor.application()
    doc = inventor.Drawing(path, app)

    auto_close = partcode[7:] == '-000-00'
    df = doc.extract_part_list(lvl=1, auto_close=auto_close)
    df = append_sub_assy_part_list(df, app, auto_close)
    return df


def _is_manu(row):
    """Check whether an item is a manufacture of purchase part"""

    if len(row) >= 3:
        if all([row[i] not in '1234567890' for i in range(3)]):
            return 'M'

    return 'B'


def save_csv_template(partcode):
    df = load_part_list(partcode)
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

    rs.to_csv(
        partcode + '.csv',
        index=False,
        # encoding='iso-8859-1',
        quoting=csv.QUOTE_NONNUMERIC
    )

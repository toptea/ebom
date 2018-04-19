import pandas as pd
from utils import system
from utils import inventor
import csv


def append_sub_assy_part_list(df, app, open_model, close_file):
    lvl = 2
    while True:
        partcodes = df.loc[~df['Dwg_No'].isin(df['Assembly']), 'Dwg_No'].values
        iam_paths = system.find_paths(partcodes, 'iam')
        idw_paths = system.find_paths([path.stem for path in iam_paths], 'idw')

        if len(idw_paths) == 0:
            break
        
        for idw_path, iam_path in zip(idw_paths, iam_paths):
            idw = inventor.Drawing(idw_path, app)
            rs = idw.extract_part_list(lvl)

            if open_model and '' in df['Component'].unique():
                iam = inventor.Drawing(iam_path, app)
                idw = inventor.Drawing(idw_path, app)
                rs = idw.extract_part_list(lvl)
            else:
                iam = None
            if close_file == 'idw' or close_file == 'both':
                idw.close()
            if (close_file == 'iam' or close_file == 'both') and iam is not None:
                iam.close()
            df = df.append(rs)
        lvl += 1
    
    return df


def load_part_list(app, partcode, open_model, close_file, recursive):

    idw_path = system.find_path(partcode, 'idw')
    idw = inventor.Drawing(idw_path, app)
    df = idw.extract_part_list(lvl=1)

    if open_model and '' in df['Component'].unique():
        iam_path = system.find_path(partcode, 'iam')
        iam = inventor.Drawing(iam_path, app)
        idw = inventor.Drawing(idw_path, app)
        df = idw.extract_part_list(lvl=1)
    else:
        iam = None
    if close_file == 'idw' or close_file == 'both':
        idw.close()
    if (close_file == 'iam' or close_file == 'both') and iam is not None:
        iam.close()
    if recursive:
        df = append_sub_assy_part_list(df, app, open_model, close_file)

    return df


def _is_manu(row):
    """Check whether an item is a manufacture of purchase part"""
    if len(row) >= 3 and all([row[i] not in '1234567890' for i in range(3)]):
        return 'M'
    else:
        return 'B'


def save_csv_template(partcode, close_file, open_model=True, recursive=True):
    app = inventor.application()
    if partcode is None:
        idw = inventor.Drawing.via_active_document(app)
        iprop = idw.doc.PropertySets.Item("Inventor User Defined Properties")
        partcode = str(iprop.Item('Dwg_No')).strip()

    df = load_part_list(app, partcode, open_model, close_file, recursive)
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


if __name__ == '__main__':
    pass

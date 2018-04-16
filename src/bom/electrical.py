import pandas as pd
from glob import glob
import csv


def load_part_list(path):
    df = pd.read_excel(path, skiprows=5)
    df = df.rename(columns={col: str(col).strip() for col in df.columns})
    project_number = str(df.loc[df['Unnamed: 8'] == 'Project Number', 'Unnamed: 9'].values[0])

    df = df[[
        'Item', 'Part Number', 'Qty',
        'Description', 'Manufacturer', 'Installation',
        'Location', 'Device ID', 'Page'
    ]]

    df = df[df['Part Number'].notnull()]
    group_columns = ['Part Number', 'Description', 'Manufacturer', 'Installation', 'Location']
    df = df.groupby(group_columns).agg({'Item': 'min', 'Qty': 'sum'}).reset_index()
    df['Assembly'] = project_number + '-' + df['Location'] + '-00'
    df = df.sort_values(by=['Assembly', 'Installation', 'Item'])
    return df


def save_csv_template(path):
    df = load_part_list(path)
    rs = pd.DataFrame()
    rs['temp'] = df['Assembly']
    rs['Import?'] = 'YES'
    rs['Parent'] = df['Assembly']
    rs['Parent Item Description'] = df['Installation']
    rs['Parent Revision'] = 1
    rs['RoutingTemplate'] = ''
    rs['Default Oper'] = 10
    rs['Operation'] = 10
    rs['Item Number'] = df['Part Number']
    rs['Item Description'] = df['Description']
    rs['Quantity'] = df['Qty']
    rs['Req Type'] = 'B'
    rs['Vendor ID'] = ''
    rs['Unit of Measure'] = 'EACH'
    rs['Length'] = ''
    rs['Width'] = ''
    rs['Product Code'] = 'PROJ'
    rs['Reference'] = ''
    rs['Detail'] = df['Item']
    rs['Drawing'] = ''
    rs['Sheet'] = ''
    rs['Notes'] = df['Manufacturer']
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
        path[:-5] + '.csv',
        index=False,
        # encoding='iso-8859-1',
        quoting=csv.QUOTE_NONNUMERIC
    )


def save_csv_templates():
    paths = glob('*.xlsx')
    paths = [p for p in paths if p[0] != '~']
    print('Converting the following files:')
    print(paths)
    print('')
    for path in paths:
        try:
            save_csv_template(path)
        except Exception as e:
            print('An error has occured:')
            print(e)

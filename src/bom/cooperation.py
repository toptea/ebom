import utils.database
import pandas as pd
import logging
import csv


def load_cooperation_ebom(partcode):
    """Engineering BOM"""
    validate1 = len(partcode) < 20
    validate2 = ' ' not in partcode
    validate3 = 'drop' not in partcode.lower()

    if validate1 & validate2 & validate3:
        sql = (
            """
            SELECT
                Parent_drawing,
                item,
                Drawing_no,
                description,
                UNIT_DESC,
                units,
                L,
                W,
                remarks
            FROM
                local_boms
            WHERE
                Parent_drawing = '""" + partcode + "'"
        )
        df = utils.database.query(sql, database_type='access_bom')
        df = df[~df['Parent_drawing'].isin(['DUMMY', '0', '90015410', 'jh', 'b'])]
        df = df[~df['Drawing_no'].isin(['DUMMY', '0', '90015410', 'jh', 'b'])]
        df = df.sort_values(['Parent_drawing', 'item'])
        df['Parent_drawing'] = df['Parent_drawing'].str.strip()
        df['Drawing_no'] = df['Drawing_no'].str.strip()
        return df
    else:
        print('Invalid part number')


def load_cooperation_stock():
    sql = (
        """
        SELECT DISTINCT
            pstk,
            desc1
        FROM
            stock
        WHERE
            pstk IS NOT NULL AND
            desc1 IS NOT NULL
        ORDER BY
            pstk
        """
    )
    df = utils.database.query(sql, database_type='access_bom')
    df['pstk'] = df['pstk'].str.strip()
    df['desc1'] = df['desc1'].str.strip()
    return df


def clean_unit_of_measure(row):
    if row in ['mm', 'MTRS', '1M-L', 'Lgth', '3M-L', '2M-L', '2M', '5M-L', '50ML']:
        return 'METRE'
    elif row in ['Each', '', 'LGTH', 'SqMtr', 'X 250', 'X 300']:
        return 'EACH'
    elif row == 'Box':
        return 'EACH'
    elif row == 'X 100':
        return 'x100'
    elif row == 'X 50':
        return 'x50'
    elif row == 'X 10':
        return 'x10'
    elif row == 'X 5':
        return 'x5'
    else:
        return 'EACH'


def _is_manu(row):
    """Check whether an item is a manufacture of purchase part"""

    if len(row) >= 3:
        if all([row[i] not in '1234567890' for i in range(3)]):
            return 'M'

    return 'B'


def save_csv_template(assembly):
    """Return X Pandas DataFrame"""
    logger = logging.getLogger(__name__)
    logger.info('converting ebom table...')
    df = load_cooperation_ebom(assembly)
    st = load_cooperation_stock()
    st = st.rename(columns={'desc1': 'Parent Item Description'})

    rs = pd.DataFrame()
    rs['order'] = df['item']
    rs['Import?'] = 'YES'
    rs['Parent'] = df['Parent_drawing'].str.strip()
    rs['Parent Revision'] = 1
    rs['RoutingTemplate'] = ''
    rs['Default Oper'] = 10
    rs['Operation'] = 10
    rs['Item Number'] = df['Drawing_no']
    rs['Item Description'] = df['description']
    rs['Quantity'] = df['units']
    rs['Req Type'] = df['Drawing_no'].apply(_is_manu)
    rs['Vendor ID'] = ''
    rs['Unit of Measure'] = df['UNIT_DESC'].str.strip().apply(clean_unit_of_measure)
    rs['Length'] = df['L']
    rs['Width'] = df['W']
    rs['Product Code'] = 'PROJ'
    rs['Reference'] = df['remarks']
    rs['Detail'] = df['item']
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
    rs = pd.merge(left=rs, right=st, how='left', left_on='Parent', right_on='pstk')
    rs = rs[[
        'Import?',
        'Parent',
        'Parent Item Description',
        'Parent Revision',
        'RoutingTemplate',
        'Default Oper',
        'Operation',
        'Item Number',
        'Item Description',
        'Quantity',
        'Req Type',
        'Vendor ID',
        'Unit of Measure',
        'Length',
        'Width',
        'Product Code',
        'Reference',
        'Detail',
        'Drawing',
        'Sheet',
        'Notes',
        'Contribution ',
        'Mfg ID',
        'Mfg Part#',
        't-ibom.x-char1',
        't-ibom.x-char2',
        't-ibom.x-char3',
        't-ibom.x-char4',
        't-ibom.x-char5',
        't-ibom.cInt1',
        't-ibom.cInt2',
        't-ibom.cInt3',
        't-ibom.cInt4',
        't-ibom.cInt5',
        't-ibom.cDate1',
        't-ibom.cDate2',
        't-ibom.cDate3',
        't-ibom.cDate4',
        't-ibom.cDate5',
        't-ibom.cDec1',
        't-ibom.cDec2',
        't-ibom.cDec3',
        't-ibom.cDec4',
        't-ibom.cDec5',
        't-ibom.cLog1',
        't-ibom.cLog2',
        't-ibom.cLog3',
        't-ibom.cLog4',
        't-ibom.cLog5',
    ]]

    rs.to_csv(
        assembly + '.csv',
        index=False,
        # encoding='iso-8859-1',
        quoting=csv.QUOTE_NONNUMERIC
    )
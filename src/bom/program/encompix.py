from utils import database
from utils import system
import pandas as pd


def load_parent_revision(ebom):
    system.status('\nParent Revision')
    system.status('|_ Connecting to Encompix database')
    try:
        sql = (
            """
            SELECT
                "item-no" AS 'Parent',
                "item-desc" AS 'Parent Item Description',
                MAX("rev") AS 'Parent Revision'
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
        system.status('|_ Find latest revision')
        df = ebom[['Parent', 'Parent Item Description']].drop_duplicates()
        df = pd.merge(left=df, right=rs, how='left', left_on='Parent', right_on='Parent')
        df['Parent Revision'] = df['Parent Revision'].fillna(0)
        df['Parent Revision'] = df['Parent Revision'].astype(int)
        df['Parent Revision'] += df['Parent Revision']
        df = df[['Parent', 'Parent Item Description', 'Parent Revision']]
        return df
    except:
        system.status('|_ WARNING: Database not found!')
        system.status('|_ WARNING: Set Parent Revision to 1')
        df = ebom[['Parent', 'Parent Item Description']].drop_duplicates()
        df['Parent Revision'] = 1
        df = df[['Parent', 'Parent Item Description', 'Parent Revision']]
        return df


def update_parent_revision(ebom, parent_revision):
    system.status('|_ Updating parent revision column')
    rs = pd.DataFrame()
    rs['Parent'] = ebom['Parent']
    rs = pd.merge(left=rs, right=parent_revision, how='left', on='Parent')
    ebom['Parent Revision'] = rs['Parent Revision']
    return ebom


def load_vendor_id():
    try:
        system.status('\nVendor ID')
        system.status('|_ Connecting to Encompix database')
        sql = (
            """
            SELECT
                "vendxref"."item-no",
                "vendxref"."vend-id",
                "vendxref"."quote-date" AS "last-order"
            FROM
                pub."vendxref"

            UNION

            SELECT
                "po-line"."item-no",
                "po-hdr"."vend-id",
                "po-hdr"."enter-date" AS "last-order"
            FROM
                pub."po-line"
                    LEFT JOIN pub."po-hdr"
                        ON "po-line"."po-no" = "po-hdr"."po-no"
            """
        )

        df = database.query(sql, database_type='progress')
        system.status('|_ Finding supplier')
        rs = df.groupby(['item-no'])['last-order'].max().reset_index()
        rs = pd.merge(left=rs, right=df[['item-no', 'vend-id']], how='left', on='item-no')
        rs = rs.drop_duplicates(subset='item-no', keep='last')
        return rs
    except:
        system.status('|_ WARNING: Database not found!')
        system.status('|_ WARNING: Unable to get Vendor ID')
        return None


def update_vendor_id(ebom, vendor_id):

    if vendor_id is not None:
        system.status('|_ Updating vendor ID from database')
        rs = pd.DataFrame()
        rs['Item Number'] = ebom['Item Number']
        rs = pd.merge(left=rs, right=vendor_id, how='left', left_on='Item Number', right_on='item-no')
        ebom['Vendor ID'] = rs['vend-id']

    system.status('|_ Updating vendor ID from pre-defined dict')
    vendor_dict = {
        'festo': '0191',
        'lee spring': '0304',
        'lee compression spring': '0304',
        'lee extension spring': '0304',
        'rose & krieger': '0390',
        'rs': '0413',
        'smc': '0484',
        'transmission dev': '0522',
        'transdev': '0522',
        'rittal': '0916',
        'keyence ': '1015',
        'rockwell': '1092',
        'item ': '1170',
        'balluff': '1191',
        'omron': '1248',
        'phoenix': '1248',
        'efd': '1277',
        'nu-tech': '1350',
        'montech': '1353',
        'staubli': '1428',
        'bosch': '1494',
        'hoffman': '1549',
        'wittenstein': '1580',
        'flowtech': '1590',
        'honle': '1629',
        'hiwin': '1658',
        'stemmer': '1681',
        'misumi': '1686',
        'swagelok': '1696',
        'transtecno': '1727',
        'oriental motor': '1732',
        'igus': '1735',
        'jokab': '1739',
        'legris transair': '1767',
        'edmund': '1786',
        'renishaw': '1897',

        'button cap': '1100',
        'button hd': '1100',
        'button head': '1100',
        'c/sk m': '1100',
        'c/sunk soc': '1100',
        'capscrew m': '1100',
        'dome nut': '1100',
        'dowel pin': '1100',
        'extractable dowel': '1100',
        'hex full nut': '1100',
        'hex hd': '1100',
        'hex head': '1100',
        'hex lock': '1100',
        'metric shoulder screw': '1100',
        'nylock nut': '1100',
        'plain washer': '1100',
        'shcs': '1100',
        'soc button': '1100',
        'soc cap': '1100',
        'soc countersunk': '1100',
        'soc csk': '1100',
        'soc hd': '1100',
        'soc head': '1100',
        'soc set': '1100',
        'socket button h': '1100',
        'socket cap': '1100',
        'socket countersunk': '1100',
        'socket csk': '1100',
        'socket hd': '1100',
        'socket head': '1100',
        'socket set screw': '1100',
        'washer m': '1100',
        'washer plain': '1100',
    }
    for key, value in vendor_dict.items():
        b1 = ebom['Vendor ID'].isnull()
        b2 = ebom['Item Description'].str.lower().str[0:len(key)] == key
        ebom.loc[b1 & b2, 'Vendor ID'] = value

    return ebom


def exlude_same_revision(ebom):
    pass


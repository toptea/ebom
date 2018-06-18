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
                tbl1."item-no" AS 'Parent',
                tbl1."rev" AS 'Parent Revision',
                tbl1."post-sta" AS "Post Status"

            FROM
                (
                -- main table
                    SELECT
                        "item-no",
                        "rev",
                        "post-sta"

                    FROM
                        pub."i-bom-hdr" 
                ) AS tbl1

                INNER JOIN

                (
                -- max revision table	
                    SELECT
                        "item-no",
                        MAX("rev") AS 'rev'

                    FROM 
                        pub."i-bom-hdr"

                    GROUP BY
                        "item-no"

                ) AS tbl2

                ON tbl1."item-no" = tbl2."item-no"
                AND tbl1."rev" = tbl2."rev"

            ORDER BY
                tbl1."item-no"
            """
        )

        rs = database.query(sql, database_type='progress')
        system.status('|_ Find latest revision')
        df = ebom[['Parent', 'Parent Item Description']].drop_duplicates()
        df = pd.merge(left=df, right=rs, how='left', left_on='Parent', right_on='Parent')
        df['Parent Revision'] = df['Parent Revision'].fillna(0)
        df['Parent Revision'] = df['Parent Revision'].astype(int)
        df = _is_bom_changed(ebom, df)
        df = df[['Parent', 'Parent Item Description', 'Parent Revision', 'Post Status', 'Changed']]
        return df
    except:
        system.status('|_ WARNING: Database not found!')
        system.status('|_ WARNING: Set Parent Revision to 1')
        df = ebom[['Parent', 'Parent Item Description']].drop_duplicates()
        df['Parent Revision'] = 1
        df['Post Status'] = 'Unknown'
        df['Changed'] = 'Unknown'
        df = df[['Parent', 'Parent Item Description', 'Parent Revision', 'Post Status', 'Changed']]
        return df


def _is_bom_changed(ebom, prev):
    sql = (
        """
        SELECT
            tbl1."p-item-no" AS 'Parent',
            tbl1."p-rev" AS 'Parent Revision',
            tbl1."seq-no",
            tbl1."item-no" AS 'Item Number',
            tbl1."item-desc" AS 'Item Description',
            tbl1."qty-item" AS 'Quantity'

        FROM 
            pub."i-bom" AS tbl1

            INNER JOIN

            (
            -- max revision table
                SELECT
                    "p-item-no",
                    MAX("p-rev") AS 'p_rev'
                FROM 
                    pub."i-bom"

                GROUP BY
                    "p-item-no"
            ) AS tbl2

            ON tbl1."p-item-no" = tbl2."p-item-no"
            AND tbl1."p-rev" = tbl2."p_rev"

        ORDER BY
            tbl1."p-item-no",
            tbl1."seq-no"
        """
    )

    rs = database.query(sql, database_type='progress')
    rs = rs[rs['Parent'].isin(prev['Parent'])]

    change_dict = {}
    for assembly in prev['Parent'].unique():
        rs2 = rs[rs['Parent'] == assembly]
        ebom2 = ebom[ebom['Parent'] == assembly]
        rs2 = rs2.reset_index()
        ebom2 = ebom2.reset_index()
        b1 = all(ebom2['Parent'].isin(rs['Parent']))
        b2 = all(ebom2['Item Number'].isin(rs2['Item Number']))
        b3 = all(rs2['Item Number'].isin(ebom2['Item Number']))
        b4 = sum(ebom2['Quantity'] - rs2['Quantity']) == 0
        change_dict[assembly] = ~(b1 & b2 & b3 & b4)

    change_df = pd.DataFrame(list(change_dict.items()), columns=['Parent', 'Changed'])
    prev = pd.merge(left=prev, right=change_df, how='left', on='Parent')

    b1 = prev['Changed'] == True
    b2 = prev['Post Status'] == 'Some Posted'
    b3 = prev['Post Status'] == 'All Posted'
    prev.loc[b1 & (b2 | b3), 'Parent Revision'] += 1
    return prev


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
        'keyence': '1015',
        'rockwell': '1092',
        'item': '1170',
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


def exlude_same_revision(ebom, prev):
    filter_assy = prev.loc[prev['Changed'] == False, 'Parent']
    filter_assy = filter_assy.tolist()
    ebom = ebom[~ebom['Parent'].isin(filter_assy)]
    return ebom

from utils import system
from utils import inventor
import utils.database

import pandas as pd
import numpy as np

import csv


class Base:

    def load_bom(self):
        """ return bom"""
        pass

    def create_indented_bom(self, bom):
        """ return ibom"""
        pass

    def create_ebom(self, bom):
        """ return ebom"""
        pass


class Cooperation(Base):
    """ Load Cooperation BOM"""

    def __Init__(self, assembly):
        self.assembly = assembly

    def load_bom(self):
        """Engineering BOM"""
        validate1 = len(self.assembly) < 20
        validate2 = ' ' not in self.assembly
        validate3 = 'drop' not in self.assembly.lower()

        if validate1 & validate2 & validate3:
            sql = (
                    """
                    SELECT DISTINCT
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
                        Parent_drawing = '""" + self.assembly + "'"
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

    def create_indented_bom(self, df):
        """ return ibom"""
        return None

    def create_ebom(self, df):
        """Return X Pandas DataFrame"""
        st = self._load_cooperation_stock()
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
        rs['Req Type'] = df['Drawing_no'].apply(self._is_manu)
        rs['Vendor ID'] = ''
        rs['Unit of Measure'] = df['UNIT_DESC'].str.strip().apply(self._clean_unit_of_measure)
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
            self.assembly + '.csv',
            index=False,
            # encoding='iso-8859-1',
            quoting=csv.QUOTE_NONNUMERIC
        )

    @staticmethod
    def _load_cooperation_stock():
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

    @staticmethod
    def _clean_unit_of_measure(row):
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

    @staticmethod
    def _is_manu(row):
        """Check whether an item is a manufacture of purchase part"""

        if len(row) >= 3:
            if all([row[i] not in '1234567890' for i in range(3)]):
                return 'M'

        return 'B'


class Promise(Base):
    """ Load Electrical BOM"""
    def __Init__(self, path):
        self.path = path

    def create_indented_bom(self, bom):
        """ return ibom"""
        return None

    def load_bom(self):
        df = pd.read_excel(self.path, skiprows=5)
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

    def create_ebom(self, df):
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
            self.path[:-5] + '.csv',
            index=False,
            # encoding='iso-8859-1',
            quoting=csv.QUOTE_NONNUMERIC
        )


class Inventor(Base):

    def __init__(self, assembly=None, close_file='never', open_model=True,
                 recursive=True, indent_format=True):
        """
        Load Mechanical BOM

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
        """
        self.assembly = assembly
        self.close_file = close_file
        self.open_model = open_model
        self.recursive = recursive
        self.indent_format = indent_format
        self.app = inventor.application()

        if assembly is None:
            self.assembly = self._get_active_assembly_partcode()

    def _get_active_assembly_partcode(self):
        idw = inventor.Drawing.via_active_document(self.app)
        iprop = idw.doc.PropertySets.Item('Inventor User Defined Properties')
        assembly = str(iprop.Item('Dwg_No')).strip()
        return assembly

    def load_bom(self):
        """ Inventor Parts List

        Using Inventor COM API, open the assembly drawing,
        extract the parts list and return the data as a dataframe.


        Returns
        ----------
        obj
            Pandas DataFrame

        """
        system.status('\n' + self.assembly)
        system.status('|_ opening idw')
        idw_path = system.find_path(self.assembly, 'idw')
        idw = inventor.Drawing(idw_path, self.app)

        df = idw.extract_part_list(lvl=1)

        if self.open_model and '' in df['Component'].unique():
            idw.close()
            system.status('|_ opening iam')
            iam_path = system.find_path(self.assembly, 'iam')
            iam = inventor.Drawing(iam_path, self.app)
            idw = inventor.Drawing(idw_path, self.app)
            system.status('|_ extracting bom')
            df = idw.extract_part_list(lvl=1)
        else:
            system.status('|_ extracting bom')
            iam = None

        if self.close_file == 'idw' or self.close_file == 'both':
            system.status('|_ closing idw')
            idw.close()
        if (self.close_file == 'iam' or self.close_file == 'both') and iam is not None:
            system.status('|_ closing iam')
            iam.close()
        if self.recursive:
            df = self._load_sub_assembly_bom(df)

        return df

    def _load_sub_assembly_bom(self, df):
        """ used in load_bom() method """
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
                idw = inventor.Drawing(idw_path, self.app)
                rs = idw.extract_part_list(lvl)

                if self.open_model and '' in rs['Component'].unique():
                    idw.close()
                    system.status('|_ opening iam')
                    iam = inventor.Drawing(iam_path, self.app)
                    idw = inventor.Drawing(idw_path, self.app)
                    system.status('|_ extracting bom')
                    rs = idw.extract_part_list(lvl)
                else:
                    system.status('|_ extracting bom')
                    iam = None
                if self.close_file == 'idw' or self.close_file == 'both':
                    system.status('|_ closing idw')
                    idw.close()
                if (self.close_file == 'iam' or self.close_file == 'both') and iam is not None:
                    system.status('|_ closing iam')
                    iam.close()
                df = df.append(rs, ignore_index=True)
            lvl += 1

        return df

    def create_indented_bom(self, df):
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
        df = self._calc_partial_irn(df)
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

        if self.indent_format:

            def _indent(row, column):
                if row[column] is not np.nan:
                    return (row['LVL']) * '    ' + row[column]

            rs['Dwg_No'] = rs.apply(_indent, column='Dwg_No', axis=1)
            rs['Component'] = rs.apply(_indent, column='Component', axis=1)

        rs = rs[['LVL', 'ITEM', 'Dwg_No', 'Component', 'QTY']]
        return rs

    def _calc_partial_irn(self, df):
        """Partial internal reference number
        used in create_indented_bom() method

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

    def create_ebom(self, df):
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


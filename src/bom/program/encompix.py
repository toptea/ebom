

def load_parent_revision(ebom):
    pass


def update_parent_revision(ebom, parent_revision):
    pass


def load_vendor_id(ebom):
    pass


def update_vendor_id(ebom, vendor_id):
    pass


def exlude_same_revision(ebom):
    pass


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




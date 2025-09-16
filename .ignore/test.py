import fdb

def find_cutouts(quote_nr_input):
    con = fdb.connect(
        dsn='10.0.0.108:C:/ZAWare/DB/CutMan/CUTMAN.FDB',
        user='SYSDBA',
        password='masterkey',
        charset='UTF8'
    )
    cur = con.cursor()

    quote_nr = quote_nr_input

    # Get cutlist_ids
    cur.execute("""
        SELECT CUTLIST_ID
        FROM CUTLIST
        WHERE QUOTE_NR = ?
    """, (quote_nr,))
    cutlist_ids = [row[0] for row in cur.fetchall()]
    if not cutlist_ids:
        cur.close()
        con.close()
        return []

    results = []
    for cutlist_id in cutlist_ids:
        # get list of crosscuts from CROSSCUTS table where quote_nr and cutlist_id match
        cur.execute("""
            SELECT *
            FROM CROSSCUTS
            WHERE QUOTE_NR = ? AND CUTLIST_ID = ?
        """, (quote_nr, cutlist_id))
        print(f"Cutlist ID: {cutlist_id}, Found {cur.rowcount} crosscuts")
        print(cur.fetchall())

        # cur.execute("""
        #     SELECT ITEM_ID, LENGTE, WYDTE, QTY
        #     FROM CUT_LIST_DETAIL
        #     WHERE QUOTE_NR = ? AND CUTLIST_ID = ?
        # """, (quote_nr, cutlist_id))
        # for item_id, lengte, wydte, qty in cur.fetchall():
        #     cur.execute("""
        #         SELECT CUTOUT1, CUTOUT2
        #         FROM CUTOUTS
        #         WHERE QUOTE_NR = ? AND CUTLIST_ID = ? AND ITEM_ID = ?
        #     """, (quote_nr, cutlist_id, item_id))
        #     cutout_row = cur.fetchone()
        #     if cutout_row:
        #         cutout1, cutout2 = cutout_row
        #         results.append((item_id, cutlist_id, lengte, wydte, qty, cutout1, cutout2))

    cur.close()
    con.close()
    return results

# Example usage:
cutouts = find_cutouts("26974")
# print(cutouts)

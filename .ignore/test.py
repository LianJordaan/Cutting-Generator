import fdb

def find_cutouts(customer_name, job_name):
    con = fdb.connect(
        dsn='10.0.0.108:C:/ZAWare/DB/CutMan/CUTMAN.FDB',
        user='SYSDBA',
        password='masterkey',
        charset='UTF8'
    )
    cur = con.cursor()

    # Get customer_id
    cur.execute("SELECT CUSTOMER_ID, NAME FROM CUSTOMER")
    customers = cur.fetchall()
    customer_id = None
    for customer in customers:
        if customer_name.upper() in customer[1].upper():
            customer_id = customer[0]
            break
    if not customer_id:
        cur.close()
        con.close()
        return []

    # Get quote_nr
    cur.execute("""
        SELECT QUOTE_NR
        FROM QUOTE
        WHERE CUSTOMER_ID = ?
          AND UPPER(JOB_NAME) = UPPER(?)
    """, (customer_id, job_name))
    row = cur.fetchone()
    if not row:
        cur.close()
        con.close()
        return []
    quote_nr = row[0]

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
        cur.execute("""
            SELECT ITEM_ID, LENGTE, WYDTE, QTY
            FROM CUT_LIST_DETAIL
            WHERE QUOTE_NR = ? AND CUTLIST_ID = ?
        """, (quote_nr, cutlist_id))
        for item_id, lengte, wydte, qty in cur.fetchall():
            cur.execute("""
                SELECT CUTOUT1, CUTOUT2
                FROM CUTOUTS
                WHERE QUOTE_NR = ? AND CUTLIST_ID = ? AND ITEM_ID = ?
            """, (quote_nr, cutlist_id, item_id))
            cutout_row = cur.fetchone()
            if cutout_row:
                cutout1, cutout2 = cutout_row
                results.append((item_id, cutlist_id, lengte, wydte, qty, cutout1, cutout2))

    cur.close()
    con.close()
    return results

# Example usage:
cutouts = find_cutouts("LIAN", "LIAN")
print(cutouts)

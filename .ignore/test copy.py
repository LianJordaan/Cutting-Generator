import fdb

def find_jobs_by_partial_name(partial_name):
    con = fdb.connect(
        dsn='10.0.0.108:C:/ZAWare/DB/CutMan/CUTMAN.FDB',
        user='SYSDBA',
        password='masterkey',
        charset='UTF8'
    )
    cur = con.cursor()

    like_value = f"%{partial_name}%"

    # 1️⃣ Get all CUSTOMER_IDs first — same as YOUR code
    cur.execute("SELECT CUSTOMER_ID, NAME FROM CUSTOMER")
    customers = cur.fetchall()

    if not customers:
        cur.close()
        con.close()
        return []

    results = []

    # 2️⃣ Loop through customers — same style as YOU did
    for customer_id, customer_name in customers:

        # Search QUOTE table for this customer with partial job name
        cur.execute("""
            SELECT QUOTE_NR, JOB_NAME
            FROM QUOTE
            WHERE CUSTOMER_ID = ?
              AND JOB_NAME LIKE ?
        """, (customer_id, like_value))

        rows = cur.fetchall()

        print(f"Checking customer: {customer_name} (ID {customer_id})")
        print(f"Found {len(rows)} matches")

        for r in rows:
            quote_nr, job_name = r
            results.append({
                "customer_id": customer_id,
                "customer_name": customer_name,
                "quote_nr": quote_nr,
                "job_name": job_name
            })

    cur.close()
    con.close()
    return results


# Example usage:
matches = find_jobs_by_partial_name("BRUSS")

for m in matches:
    print(m)

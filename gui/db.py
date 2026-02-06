import psycopg2

conn = psycopg2.connect(
    host="localhost",
    database="pyxcel_db",
    user="postgres",
    password="Pyxcel@123",
    port=5432
)

def log_operation(filename, operation):
    try:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO process_logs (filename, operations, status) VALUES (%s, %s, %s)",
            (filename, operation, "success")
        )
        conn.commit()
        cur.close()
        print("✅ Logged:", filename, operation)
    except Exception as e:
        print("DB error:", e)

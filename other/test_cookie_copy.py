import os

db = r"C:\Users\Meet Patel\AppData\Local\Google\Chrome\User Data\Default\Network\Cookies"
temp_db = r"automation\test_cookie.db"

try:
    with open(db, "rb") as f_in:
        data = f_in.read()
    with open(temp_db, "wb") as f_out:
        f_out.write(data)
    print("Raw read succeeded! Length:", len(data))
    if os.path.exists(temp_db):
        os.remove(temp_db)
except Exception as e:
    print("Direct read failed:", e)

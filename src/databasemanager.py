import sqlite3

database_name = "database.db"

class DatabaseManager:
    def __init__(self) -> None:
        try:
            conn = sqlite3.connect(database_name)
            cursor = conn.cursor()
            print("Database created!")

        except Exception as e:
            print("Something bad happened: ", e)
            if conn:
                conn.close()
import sqlite3

database_name = "database.db"

class DatabaseManager:
    def __init__(self) -> None:
        try:
            self.conn = sqlite3.connect(database_name)
            self.cursor = self.conn.cursor()
            print("Database created!")

        except Exception as e:
            print("Something bad happened: ", e)
            if self.conn:
                self.conn.close()

        #Opretter tabellerne.
        self.create_tables()

    def create_tables(self):
        # Create operation
        create_query = '''CREATE TABLE "tblEvents" (
                "id"	INTEGER,
                "outlook_id"	TEXT,
                "aula_id"	TEXT,
                "created"	TEXT,
                "updated"	TEXT,
                PRIMARY KEY("id" AUTOINCREMENT)
            );
        '''
        self.cursor.execute(create_query)
        print("Table created!")


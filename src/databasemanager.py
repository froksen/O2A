import sqlite3
from databaseevent import DatabaseEvent as dbEvent

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
        try:
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
        except sqlite3.OperationalError as e:
            print(e)

    def get_record(self, outlook_id):
        cursor = self.conn.cursor()
        records = cursor.execute("SELECT * FROM tblEvents WHERE outlook_id=:outlook_id",{"outlook_id":outlook_id}).fetchone()

        if records is None:
            return None

        event = dbEvent()
        event.db_id = records[0]
        event.aula_id = records[2]
        event.outlook_id = records[1]
        event.created = records[3]
        event.updated = records[4]

        return event

    def update_record(self, outlook_id, aula_id):
        pass
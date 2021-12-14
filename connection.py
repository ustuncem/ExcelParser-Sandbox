from typing import Any
import mysql.connector as db

class Database():
    """Database Connection Class
    """

    def __init__(self, username = "root", password = "", host = "localhost", database = "karekok") -> None:
        """Constructor Class for the Database

        Args:
            username (str, optional): [MySQL Database Username]. Defaults to "admin".
            password (str, optional): [MySQL Database Password]. Defaults to "".
            host (str, optional): [MySQL Database Host]. Defaults to "localhost".
            database (str, optional): [MySQL Database Host]. Defaults to "localhost".
        """

        self.username = username
        self.password = password
        self.host = host
    
    def get_conntection(self, database_name = "") -> Any:
        """ Returns the MySQL Database Connection object.

        Args:
            database_name (str): [Database Name]. Defaults to "".

        Returns:
            [object] : [MySQLConnection object.]
        """
        return db.connect( host = self.host, user = self.username, password = self.password, database = database_name) 


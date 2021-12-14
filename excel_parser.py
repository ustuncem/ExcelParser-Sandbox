import glob
from typing import Any
import openpyxl

from connection import Database

class ExcelParser():
    """ Methods to parse excel data to Collections
    """

    file_locations: list = []
    database: object = Database().get_conntection("karekok")

    def __init__(self, file_location: str) -> None:
        """[Constructor Class for the Parser]

        Args:
            file_location (str): [Give a path and leave the rest to the parser!]
        """
        self.file_location = file_location
        self._get_excel_paths()
            
    def _get_excel_paths(self) -> None:
        """Get all the paths in the string form. For security, don't expose your 
            static assets path to public as it may contain sensetive information about your customers.
        """
        self.file_locations = glob.glob("./{}".format(self.file_location))
    
    def read_and_parse(self):
        """Main Function of the Class and also the entrypoint. Here I used 4 loops because data from my client was coming
        in an excel spreadsheet with multiple worksheets. This can be optimized to reduce the process cost. O(n) notation is possible.
        TODO: Optimize algorithm to at least O(n^2)
        """
        for index in self.file_locations:
            print(index)
            sheets = openpyxl.load_workbook(index).worksheets

            for sheet in sheets:
                
                business_category = self.search_business_category(sheet.title)

                for row in sheet.iter_rows(max_col=3):
                    result = self.search((row[0].value, row[1].value, row[2].value))
                    print(result)
                    if len(result) > 0:
                        for x in result:
                            self.update(business_category, x[0])
                    else:
                        self.insert((row[2].value, business_category, row[0].value, row[1].value.upper()))
    
    def search(self, value: tuple) -> Any:
        """Search database if the certain record exists. Feel free to tweak it as you wish.

        Args:
            value (tuple): Search tuple, for unnecessary definition of parameters, use tuple instead as you may need more search parameters in the future.

        Returns:
            any: List of results that are found. Empty set if nothing is found
        """
        cursor = self.database.cursor()

        query = "SELECT business_id, business_title, business_city, business_town FROM crmbusiness WHERE business_city = '{}' AND business_town = '{}' AND business_title = '{}'".format(value[0], value[1], value[2])

        cursor.execute(query)

        return cursor.fetchall()
    
    def search_business_category(self, sheet_title: str) -> int:
        """This is a custom search function that I needed again for my client. This unnecessary in the common script.

        Args:
            sheet_title (str): [Title of the worksheet in Excel file]

        Returns:
            int: [Returns the business category id]
        """
        cursor = self.database.cursor()

        query = "SELECT business_category_id, business_category_title FROM crmbusinesscategory WHERE business_category_title LIKE '{}'".format(sheet_title)

        cursor.execute(query)

        return cursor.fetchall()[0][0]

    def insert(self, values: tuple) -> None:
        """Insert the unmatched records to the database

        Args:
            values (tuple): Parameters tuple, for unnecessary definition of parameters, use tuple as you may need more search parameters in the future.
        """
        cursor = self.database.cursor()

        query = """
            INSERT INTO crmbusiness (business_title, business_category_id, crm_business_location_address_type, business_city, business_town)
            VALUES("%s", %d, "manual", "%s", "%s")
        """%values

        cursor.execute(query)

        self.database.commit()

        print("1 record inserted, ID:", cursor.lastrowid)

    def update(self, business_id: int, business_category: int) -> None:
        """Update the matched records category id. This is also a custom function for my client. As it occurs some categories are mismatched in the live website.

        Args:
            business_id (int): [Business id of the record to be updated]
            business_category (int): [Category id that the business must be assigned to.]
        """
        cursor = self.database.cursor()

        query = "UPDATE crmbusiness SET business_category_id = %d WHERE business_id = %d"%(business_category, business_id)
        
        
        cursor.execute(query)

        self.database.commit()

        print(cursor.rowcount, "record(s) affected")

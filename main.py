from connection import Database
from excel_parser import ExcelParser

database = Database().get_conntection("karekok")
ExcelParser("./assets/*").read_and_parse()
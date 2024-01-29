from openpyxl import load_workbook
from openpyxl import Workbook
import os

class DataProcessor:
    def __init__(self):
        self.filepath = os.path.join(os.getcwd(), 'database.xlsx')
        self.sheet = load_workbook(filename=self.filepath)
        self.selected_sheet = self.sheet['Novos Safra Nov23']
    
    def add_data_to_worksheet(self, position, data):
        self.selected_sheet[position].value = data
        
        self.sheet.save(self.filepath)
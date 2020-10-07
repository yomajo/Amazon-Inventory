from amzn_parser_utils import col_to_letter
from openpyxl.styles import Alignment
import openpyxl
import os

# GLOBAL VARIABLES
SUMMARY_SHEET_NAME = 'SKU codes'
HEADERS = ['SKU', 'Quantity', 'Item']
BOLD_STYLE = openpyxl.styles.Font(bold=True, name='Calibri')


class HelperFile():
    '''accepts export data dictionary as argument, creates formatted xlsx file.
    Class does not include error handling and that should be carried out outside of this class scope

    Main method: export() - takes argument of target workbook name (path) and pushes
    export_obj accepted by class to single sheet'''
    
    def __init__(self, export_obj):
        self.export_obj = export_obj
        self.col_widths = {}

    def export(self, wb_name : str):
        '''Creates workbook, and exports self.export_obj object to single sheet, saves new workbook'''
        self.wb = openpyxl.Workbook()
        ws = self.wb.active
        ws.freeze_panes = ws['A2']
        ws.title = SUMMARY_SHEET_NAME
        self.fill_sheet(ws)
        self.wb.save(wb_name)
        self.wb.close()
    
    def fill_sheet(self, ws : object):
        '''pushes export object to workbook; adjusts column widths'''
        self.row_cursor = 1
        self.__fill_headers(ws)
        self.push_data(ws)
        self._adjust_col_widths(ws, self.col_widths)

    def __fill_headers(self, ws : object):
        '''inserts 3 headers in 1:1 row. Bold style and update col widths dict for all'''
        for col, header in enumerate(HEADERS, start=1):
            ws.cell(self.row_cursor, col).value = header
            ws.cell(self.row_cursor, col).font = BOLD_STYLE
            self.__update_col_widths(col, header, zero_indexed=False)
        self.row_cursor += 1

    def push_data(self, ws : object):
        '''unpacks self.export_obj to ws sheet'''
        for sku_data in self.export_obj:
            ws.cell(self.row_cursor, 1).value = sku_data[0]
            ws.cell(self.row_cursor, 2).value = sku_data[1]['quantity']
            ws.cell(self.row_cursor, 2).alignment = Alignment(horizontal='left')
            ws.cell(self.row_cursor, 3).value = sku_data[1]['item']
            self.__update_col_widths(1, sku_data[0], zero_indexed=False)
            self.__update_col_widths(2, str(sku_data[1]['quantity']), zero_indexed=False)
            self.__update_col_widths(3, sku_data[1]['item'], zero_indexed=False)
            self.row_cursor += 1

    def __update_col_widths(self, col : int, cell_value : str, zero_indexed=True):
        '''runs on each cell. Forms a dictionary {'A':30, 'B':15...} for max column widths in worksheet (width as length of max cell)'''
        col_letter = col_to_letter(col, zero_indexed=zero_indexed)
        if col_letter in self.col_widths:
            # check for length, update if current cell length exceeds current entry for column
            if len(cell_value) > self.col_widths[col_letter]:
                self.col_widths[col_letter] = len(cell_value)
        else:
            self.col_widths[col_letter] = len(cell_value)

    @staticmethod
    def _adjust_col_widths(ws, col_widths : dict):
        '''iterates over {'A':30, 'B':40, 'C':35...} dict to resize worksheets' column widths'''
        for col_letter in col_widths:
            adjusted_width = col_widths[col_letter] + 4
            ws.column_dimensions[col_letter].width = adjusted_width

if __name__ == "__main__":
    pass
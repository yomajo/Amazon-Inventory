from amzn_parser_utils import get_output_dir, get_last_used_row_col, col_to_letter, sort_by_quantity
from openpyxl.styles import Alignment
from shutil import copy
import logging
import openpyxl
import os


# GLOBAL VARIABLES
SUMMARY_SHEET_NAME = 'SKU codes'
SKU_MAPPING_SHEET_NAME = 'Codes Mapping'
HEADERS = ['sku', 'quantity', 'item']
BOLD_STYLE = openpyxl.styles.Font(bold=True, name='Calibri')


class HelperFileCreate():
    '''accepts export data dictionary as argument, creates formatted xlsx file.
    Class does not include error handling and that should be carried out outside of this class scope

    Main method: export() - takes argument of target workbook name (path) and pushes
    export_obj accepted by class to single sheet'''
    
    def __init__(self, export_obj:dict):
        self.sorted_export_obj = sort_by_quantity(export_obj)
        self.col_widths = {}

    def export(self, wb_name : str):
        '''Creates workbook, and exports self.sorted_export_obj object to single sheet, saves new workbook'''
        self.wb = openpyxl.Workbook()
        ws = self.wb.active
        ws.freeze_panes = ws['A2']
        ws.title = SUMMARY_SHEET_NAME
        self.fill_sheet(ws)
        self.wb.create_sheet(SKU_MAPPING_SHEET_NAME)
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
        '''unpacks self.sorted_export_obj to ws sheet'''
        for sku_data in self.sorted_export_obj:
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

class HelperFileUpdate():
    '''accepts export data dictionary as argument.
    Class includes error handling, but raises Exception to hit outside error handler to close db connection and alert VBA.

    Main method: update_workbook() - takes argument of workbook path, reads contents from SUMMARY_SHEET_NAME, cleans sheet,
    merges current contents with incoming data in export_obj and pushes updated values'''
    
    def __init__(self, export_obj:dict):
        '''Different from self.export_obj in HelperFileCreate. Stil dict of dicts at this point'''
        self.export_obj = export_obj

    def update_workbook(self, inventory_file):
        '''main cls method. Handles reading, merging of current and incoming data, pushes updated data'''
        try:
            # Backup and set workbook, worksheet objs
            wb = openpyxl.load_workbook(inventory_file)
            self.ws = wb[SUMMARY_SHEET_NAME]
            self.backup_wb(inventory_file)
            
            # Read contents to object
            data_in_wb = self.read_ws_to_sku_dict()
            updated_sku_codes_data = self.update_sku_data(data_in_wb, self.export_obj)
            
            # Sort by quantity, transform and push updated values back to ws
            sorted_updated_data = sort_by_quantity(updated_sku_codes_data)
            self.write_updated_to_ws(sorted_updated_data)

            logging.info(f'Writing updated values done. Saving, closing...')
            wb.save(inventory_file)
            wb.close()
        except Exception as e:
            logging.critical(f'Errors inside HelperFileUpdate.updateworkbook Errr: {e}. Closing wb without saving')
            wb.close()
            logging.warning(f'Raising error, to shutdown db connection, warn VBA in ParseOrders')
            raise Exception('Transition from HelperFileUpdate.updateworkbook error handling to ParseOrders.export_update_inventory_helper_file error handling')
        
    @staticmethod
    def backup_wb(inventory_file:str):
        '''Creates a backup of workbook before new edits'''
        backup_dir = get_output_dir(client_file=False)
        backup_path = os.path.join(backup_dir, 'Inventory Reduction b4lastrun.xlsx')
        copy(inventory_file, backup_path)
        logging.info(f'Backup created at: {backup_path}, before touching {inventory_file}')

    def read_ws_to_sku_dict(self):
        ws_limits = get_last_used_row_col(self.ws)
        assert ws_limits['max_col'] == 3, 'Template of helper file changed! Maximum column used in ws != 3'
        current_sku_codes = self.get_ws_data(ws_limits)
        return current_sku_codes

    def get_ws_data(self, ws_limits:dict) -> dict:
        '''iterates though data rows [2:ws.max_row] in self.ws and collects sku data to dict object:
        {sku1:{'quantity':2, 'item':'item_title1'}, sku2:{'quantity':4, 'item':'item_title2'}, ...}
        Also deletes worksheet contents (excl headers in 1:1 row)'''
        current_sku_codes = {}
        for r in range(2, ws_limits['max_row'] + 1):
            sku, quantity, item = self.get_ws_row_data(r)
            self.clean_ws_row_data(r, ws_limits)
            if sku not in current_sku_codes.keys():
                current_sku_codes[sku] = {'item':item, 'quantity':quantity}
            else:
                current_sku_codes[sku]['quantity'] += quantity
        logging.info(f'Before updating workbook has {len(current_sku_codes.keys())} distinct sku codes / data rows excl. headers')
        return current_sku_codes

    def get_ws_row_data(self, r:int):
        '''returns sku, quantity, item from columns A,B,C in self.ws (SUMMARY_SHEET_NAME) on r (arg) row'''
        sku = self.ws.cell(r, 1).value
        item = self.ws.cell(r, 3).value
        try:
            quantity = int(self.ws.cell(r, 2).value)
        except ValueError as e:
            logging.warning(f'Error converting quantity to integer, data found in wb cell: {self.ws.cell(r, 2).value}. Proceeding with string value')
            quantity = self.ws.cell(r, 2).value
        return sku, quantity, item

    def clean_ws_row_data(self, r:int, ws_limits:dict):
        '''deletes row r contents in self.ws worksheet (SUMMARY_SHEET_NAME)'''
        for col in range(1, ws_limits['max_col'] + 1):
            self.ws.cell(r, col).value = None

    def update_sku_data(self, current_sku_codes:dict, loaded_sku_data:dict) -> dict:
        '''updates current_sku_codes dict with values from loaded data, returns same form obj:
        {sku1:{'quantity':2, 'item':'item_title1'}, sku2:{'quantity':4, 'item':'item_title2'}, ...}'''
        for sku in loaded_sku_data.keys():
            if sku not in current_sku_codes.keys():
                logging.debug(f'Adding a new sku code: {sku}. Details: {loaded_sku_data[sku]}')
                current_sku_codes[sku] = {'item':loaded_sku_data[sku]['item'], 'quantity':loaded_sku_data[sku]['quantity']}
            else:
                logging.debug(f"Updating quantity for code: {sku}. Previous quantity: {current_sku_codes[sku]['quantity']}, adding: {loaded_sku_data[sku]['quantity']}")
                current_sku_codes[sku]['quantity'] += loaded_sku_data[sku]['quantity']
        return current_sku_codes

    def write_updated_to_ws(self, sorted_updated_data:list):
        '''write sorted_updated_data list of tuples to rows below header'''
        for row_cursor, sku_data in enumerate(sorted_updated_data, start=2):
            self.ws.cell(row_cursor, 1).value = sku_data[0]
            self.ws.cell(row_cursor, 2).value = sku_data[1]['quantity']
            self.ws.cell(row_cursor, 3).value = sku_data[1]['item']


if __name__ == "__main__":
    pass
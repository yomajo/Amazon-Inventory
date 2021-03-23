from amzn_parser_utils import get_output_dir, get_last_used_row_col, col_to_letter, sort_by_quantity, get_inner_quantity_and_custom_label
from constants import QUANTITY_PATTERN, VBA_ALREADY_OPEN_ERROR, SHEET_NAME, HEADERS
from openpyxl.styles import Alignment
from shutil import copy
import logging
import openpyxl
import os


# GLOBAL VARIABLES
BOLD_STYLE = openpyxl.styles.Font(bold=True, name='Calibri')
FILL_NONE = openpyxl.styles.PatternFill(fill_type=None)
FILL_HIGHLIGHT = openpyxl.styles.PatternFill(fill_type='solid', fgColor='FAFA73')


class HelperFileCreate():
    '''accepts export data, sku-custom label mapping dictionaries as args, creates formatted xlsx file.
    Class does not include error handling and that should be carried out outside of this class scope

    Main method: export() - takes argument of target workbook name (path) and pushes
    export_obj accepted by class to single sheet, highlights unmapped codes'''
    
    def __init__(self, export_obj:dict, mapping_dict:dict):
        self.sorted_export_obj = sort_by_quantity(export_obj)
        self.mapping_dict = mapping_dict
        self.col_widths = {}

    def export(self, wb_name:str):
        '''Creates workbook, and exports self.sorted_export_obj object to single sheet, saves new workbook'''
        wb = openpyxl.Workbook()
        self.ws = wb.active
        self.ws.freeze_panes = self.ws['A2']
        self.ws.title = SHEET_NAME
        self.fill_sheet()
        wb.save(wb_name)
        wb.close()
    
    def fill_sheet(self):
        '''pushes export object to workbook; adjusts column widths'''
        self.row_cursor = 1
        self.__fill_headers()
        self.push_data()
        self._adjust_col_widths(self.ws, self.col_widths)

    def __fill_headers(self):
        '''inserts 3 headers in 1:1 row. Bold style and update col widths dict for all'''
        for col, header in enumerate(HEADERS, start=1):
            self.ws.cell(self.row_cursor, col).value = header
            self.ws.cell(self.row_cursor, col).font = BOLD_STYLE
            self.__update_col_widths(col, header, zero_indexed=False)
        self.row_cursor += 1

    def push_data(self):
        '''unpacks self.sorted_export_obj to self.ws sheet'''
        for sku_data in self.sorted_export_obj:
            self.ws.cell(self.row_cursor, 1).value = sku_data[0]
            self.ws.cell(self.row_cursor, 2).value = sku_data[1]['quantity']
            self.ws.cell(self.row_cursor, 2).alignment = Alignment(horizontal='left')
            self.ws.cell(self.row_cursor, 3).value = sku_data[1]['item']
            self.__update_col_widths(1, sku_data[0], zero_indexed=False)
            self.__update_col_widths(2, str(sku_data[1]['quantity']), zero_indexed=False)
            self.__update_col_widths(3, sku_data[1]['item'], zero_indexed=False)
            # highlight row is code does not yet have a mapping
            self.__highlight_unmapped_sku_on_data_push(sku_data[0], self.row_cursor)
            self.row_cursor += 1

    def __highlight_unmapped_sku_on_data_push(self, custom_label:str, r:int):
        '''applies highlight if custom label about to be pushed to wb does not yet exist in mapping file'''
        highlight = True
        for _, mapped_custom_label in self.mapping_dict.items():
            if custom_label in mapped_custom_label:
                highlight = False
                break
        if highlight:
            self.__apply_row_highlight(r)

    def __apply_row_highlight(self, r:int, highlight_style=FILL_HIGHLIGHT):
        '''applies passed formatting style on self.ws r row, hardcoded 4 columns fill'''
        for c in range(1, 4):
            self.ws.cell(r, c).fill = highlight_style

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
    def _adjust_col_widths(ws:object, col_widths : dict):
        '''iterates over {'A':30, 'B':40, 'C':35...} dict to resize worksheets' column widths'''
        for col_letter in col_widths:
            adjusted_width = col_widths[col_letter] + 4
            ws.column_dimensions[col_letter].width = adjusted_width


class HelperFileUpdate():
    '''Input args: export data as dict, mapping_dict as dict
    Class includes error handling, but raises Exception to hit outside error handler to close db connection and alert VBA.

    Main method: update_workbook() - takes argument of workbook path, reads contents, maps sku-custom labels,
    corrects quantities, from SHEET_NAME, cleans sheet,
    merges current contents with incoming data in export_obj and pushes updated values'''
    
    def __init__(self, export_obj:dict, mapping_dict:dict):
        '''Different from self.export_obj in HelperFileCreate. Stil dict of dicts at this point'''
        self.export_obj = export_obj
        self.mapping_dict = mapping_dict

    def update_workbook(self, inventory_file):
        '''main cls method. Handles reading, cleaning, formatting, merging of current and incoming data, pushes updated data'''
        try:
            # Backup and set workbook, worksheet objs
            wb = openpyxl.load_workbook(inventory_file)
            self.ws = wb[SHEET_NAME]
            self.backup_wb(inventory_file)
            
            # Read contents to object
            mapped_data_in_wb_list = self.read_map_ws_data_to_list()
            corrected_data = self._correct_wb_data_for_inner_quantities_in_codes(mapped_data_in_wb_list)
            updated_sku_codes_data = self.update_sku_data(corrected_data, self.export_obj)    

            # Sort by quantity, transform and push updated values back to ws
            sorted_updated_data = sort_by_quantity(updated_sku_codes_data)
            self.write_updated_to_ws(sorted_updated_data)

            logging.info(f'Writing updated values done. Saving, closing...')
            wb.save(inventory_file)
            wb.close()
        except PermissionError as e:
            logging.critical(f'Workbook {inventory_file} already open. Err: {e}')
            print(VBA_ALREADY_OPEN_ERROR)
            raise Exception('Transition from HelperFileUpdate.updateworkbook error handling to ParseOrders.export_update_inventory_helper_file error handling')
        except Exception as e:
            logging.critical(f'Errors inside HelperFileUpdate.updateworkbook Errr: {e}. Closing wb without saving')
            wb.close()
            raise Exception('Transition from HelperFileUpdate.updateworkbook error handling to ParseOrders.export_update_inventory_helper_file error handling')

    def _correct_wb_data_for_inner_quantities_in_codes(self, mapped_wb_data:list) -> dict:
        '''correct workbook data for all inner quantities and custom codes.
        
        Arg input format: [sku_dict_1, sku_dict_2, ...]
        where sku_dict_n = {code1:{'quantity':2, 'item':'item_title1'}}
        
        Function returns:
        {unique_inner_code1:{'quantity':2, 'item':'item_title1'}, unique_inner_code2:{'quantity':4, 'item':'item_title2'}, ...}'''
        corrected_data = {}

        for custom_label_dict in mapped_wb_data:
            custom_label = list(custom_label_dict.keys())[0]
            inner_quantity, inner_code = get_inner_quantity_and_custom_label(custom_label, QUANTITY_PATTERN)

            # Corrected quantity=  quantity from wb 'quantity' column * extracted quantity inside custom label
            corrected_quantity = custom_label_dict[custom_label]['quantity'] * inner_quantity
            
            if inner_code not in corrected_data.keys():
                corrected_data[inner_code] = {'item':custom_label_dict[custom_label]['item'], 'quantity':corrected_quantity}
            else:
                corrected_data[inner_code]['quantity'] += corrected_quantity
        return corrected_data

    @staticmethod
    def backup_wb(inventory_file:str):
        '''Creates a backup of workbook before new edits'''
        backup_dir = get_output_dir(client_file=False)
        backup_path = os.path.join(backup_dir, 'Inventory Reduction b4lastrun.xlsx')
        copy(inventory_file, backup_path)
        logging.info(f'Backup created at: {backup_path}, before touching {inventory_file}')

    def read_map_ws_data_to_list(self) -> list:
        ws_limits = get_last_used_row_col(self.ws)
        assert ws_limits['max_col'] == 3, 'Template of helper file changed! Maximum column used in ws != 3'
        current_sku_codes = self.get_ws_data_reset_highlight(ws_limits)
        mapped_sku_custom_labels_list = self._map_sku_custom_label_codes(current_sku_codes)
        return mapped_sku_custom_labels_list

    def _map_sku_custom_label_codes(self, current_sku_codes:dict) -> list:
        '''uses self.mapping_dict to change amazon sku's to known native custom_labels in inventory management
        arg: {sku1:{'quantity':2, 'item':'item_title1'}, sku2:{'quantity':4, 'item':'item_title2'}, ...}
        
        returns: [{mapped_custom_label_1:{'quantity':2, 'item':'item_title1'}},
                    {mapped_custom_label_1:{'quantity':4, 'item':'item_title1'}}
                    {custom_label_2:{'quantity':1, 'item':'item_title2'}}, 
                    ...]'''
        # wb data sku's are unique, but after mapping, custom_labels (mapped sku) could have duplicates before merging their quantities,
        # therefore list has to be used to house potential duplicate custom_labels
        mapped_ws_data = []
        for sku in current_sku_codes.keys():
            mapped_entry_dict = {}
            if sku in self.mapping_dict.keys():
                logging.debug(f'Replacing wb sku {sku} with {self.mapping_dict[sku]}')
                mapped_entry_dict[self.mapping_dict[sku]] = current_sku_codes[sku]
                mapped_ws_data.append(mapped_entry_dict)
            else:
                mapped_entry_dict[sku] = current_sku_codes[sku]
                mapped_ws_data.append(mapped_entry_dict)
        logging.info(f'Mapped ws sku\'s with custom labels: sku count: {len(current_sku_codes.keys())} vs mapped list: {len(mapped_ws_data)} after mapping')
        return mapped_ws_data

    def get_ws_data_reset_highlight(self, ws_limits:dict) -> dict:
        '''iterates though data rows [2:ws.max_row] in self.ws and collects sku data to dict object:
        {sku1:{'quantity':2, 'item':'item_title1'}, sku2:{'quantity':4, 'item':'item_title2'}, ...}
        Also deletes worksheet contents (excl headers in 1:1 row)'''
        current_sku_codes = {}
        for r in range(2, ws_limits['max_row'] + 1):
            # Reset color formatting
            self.__apply_row_highlight(r, FILL_NONE)
            # Get data
            sku, quantity, item = self.get_ws_row_data(r)
            self.clean_ws_row_data(r, ws_limits)
            if sku not in current_sku_codes.keys():
                current_sku_codes[sku] = {'item':item, 'quantity':quantity}
            else:
                current_sku_codes[sku]['quantity'] += quantity
        logging.info(f'Before updating workbook has {len(current_sku_codes.keys())} distinct sku codes / data rows excl. headers')
        return current_sku_codes

    def get_ws_row_data(self, r:int):
        '''returns sku, quantity, item from columns A,B,C in self.ws (SHEET_NAME) on r (arg) row'''
        sku = self.ws.cell(r, 1).value
        item = self.ws.cell(r, 3).value
        try:
            quantity = int(self.ws.cell(r, 2).value)
        except ValueError as e:
            logging.warning(f'Error converting quantity to integer, data found in wb cell: {self.ws.cell(r, 2).value}. Proceeding with string value')
            quantity = self.ws.cell(r, 2).value
        return sku, quantity, item

    def clean_ws_row_data(self, r:int, ws_limits:dict):
        '''deletes row r contents in self.ws worksheet (SHEET_NAME)'''
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
            self.__highlight_unmapped_sku_on_data_push(sku_data[0], row_cursor)

    def __highlight_unmapped_sku_on_data_push(self, custom_label:str, r:int):
        '''applies highlight if custom label about to be pushed to wb does not yet exist in mapping file'''
        highlight = True
        for _, mapped_custom_label in self.mapping_dict.items():
            if str(custom_label) in mapped_custom_label:
                highlight = False
                break
        if highlight:
            self.__apply_row_highlight(r)

    def __apply_row_highlight(self, r:int, highlight_style=FILL_HIGHLIGHT):
        '''applies passed formatting style on self.ws r row, hardcoded 4 columns fill'''
        for c in range(1, 4):
            self.ws.cell(r, c).fill = highlight_style

if __name__ == "__main__":
    pass
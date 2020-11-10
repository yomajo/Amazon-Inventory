import openpyxl
import os
from shutil import copy
from amzn_parser_utils import get_output_dir

from parse_orders import EXPORT_FILE
from helper_file import SUMMARY_SHEET_NAME
from pprint import pprint

# fresh export object formed from program:
'''returns sorted object:
export_obj = {'sku1': {
                'item': 'item_name1',
                'quantity: 2},
            'sku3': {
                'item': 'item_name2',
                'quantity: 5}, ...}'''

TEST_EXPORT_OBJ = {'EX200': {
                        'item': 'GOLDEN UNIVERSAL TAROT DECK [Karten] [2013] Lo Scarabeo',
                        'quantity': 15},
                    'EX198': {
                        'item': 'As Above Deck: Book of Shadows Tarot, Volume 1 [Karten] [2013] Lo Scarabeo',
                        'quantity': 5}
                    }


def backup_wb():
    output_dir = get_output_dir(client_file=False)
    target_path = os.path.join(output_dir, 'Inventory Reduction b4lastrun.xlsx')
    copy(EXPORT_FILE, target_path)
    print(f'Backup created at: {target_path}')


def read_ws(ws:object):
    try:
        max_rows = ws.max_row
        max_cols = ws.max_column
        headers = get_headers(ws, max_cols)
        current_sku_codes = get_ws_data(ws, headers, max_rows)
        inspect_obj(current_sku_codes, headers)
    except Exception as e:
        print(f'Error handling workbook. Error: {e}')
        print('Saving, closing workbook...')


def get_ws_data(ws:object, headers:list, max_rows:int):
    ''''''
    current_sku_codes = {}
    # Collecting data to dict of dicts
    try:
        for r in range(2, max_rows + 1):
            sku, quantity, item = get_ws_row_data(ws, r)
            if sku not in current_sku_codes.keys():
                current_sku_codes[sku] = {'item':item, 'quantity':quantity}
            else:
                current_sku_codes[sku]['quantity'] += quantity
        return current_sku_codes
    except Exception as e:
        print(f'Error collecting data from ws. Error: {e}')
        print('Saving, closing workbook...')

def get_ws_row_data(ws:object, r:int):
    sku = ws.cell(r, 1).value
    try:
        quantity = int(ws.cell(r, 2).value)
    except ValueError as e:
        print(f'Error converting quantity to integer, data found in wb cell: {ws.cell(r, 2).value}. Proceeding with string value')
        quantity = ws.cell(r, 2).value
    item = ws.cell(r, 3).value
    return sku, quantity, item



def inspect_obj(current_sku_codes, headers):
    if isinstance(current_sku_codes, list):
        print(f'First data row: {current_sku_codes[0]}')
        print(f'Last row data: {current_sku_codes[-1]}')
        print(f'In total this obj has: {len(current_sku_codes)} sku codes. Should be > > 165 < <.')
        print('trying to access headers of first data row')
        first_data_dict = current_sku_codes[0]
        for header in headers:
            print(f'Header: {header} has value: {first_data_dict[header]}')
    else:
        print(f'Collected dict has: {len(list(current_sku_codes.keys()))} keys in total')
        pprint(current_sku_codes)
    print('Done inspecting')

def get_headers(ws:object, max_cols:int):
    '''gets a list of column headers in 1:1 row of ws'''
    headers = []
    for c in range(1, max_cols + 1):
        header = ws.cell(1, c).value
        headers.append(header)
    return headers


def run():
    backup_wb()

    # Open WB
    wb = openpyxl.load_workbook(EXPORT_FILE)
    ws = wb[SUMMARY_SHEET_NAME]
    
    # Read contents to object
    read_ws(ws)

    # Close WB
    wb.save(EXPORT_FILE)
    wb.close()


if __name__ == "__main__":
    run()
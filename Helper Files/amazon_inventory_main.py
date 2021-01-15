from constants import VBA_ERROR_ALERT, VBA_KEYERROR_ALERT, VBA_OK, SKU_MAPPING_WB_NAME
from amzn_parser_utils import get_output_dir, is_windows_machine
from parse_orders import ParseOrders
from orders_db import OrdersDB
from datetime import datetime
import logging
import sys
import csv
import os

# Logging config:
log_path = os.path.join(get_output_dir(client_file=False), 'amazon_inventory.log')
logging.basicConfig(handlers=[logging.FileHandler(log_path, 'a', 'utf-8')], level=logging.INFO)

# GLOBAL VARIABLES
TESTING = False
EXPECTED_SYS_ARGS = 2

if is_windows_machine():
    TEST_AMZN_EXPORT_TXT = r'C:\Coding\Ebay\Working\Backups\Amazon exports\Collected exports\export 2020.06.16.txt'
else:
    TEST_AMZN_EXPORT_TXT = r'/home/devyo/Coding/Git/Amazon Inventory/Amazon exports/run1.txt'

# Avoiding circular reference on imports
from sku_mapping import SKUMapping
SKU_MAPPING_WB_PATH = os.path.join(get_output_dir(client_file=False), SKU_MAPPING_WB_NAME)
SKU_MAPPING_DICT = SKUMapping(SKU_MAPPING_WB_PATH).read_sku_mapping_to_dict()


def get_cleaned_orders(source_file:str) -> list:
    '''returns cleaned orders (as cleaned in clean_orders func) from source_file arg path'''
    raw_orders = get_raw_orders(source_file)
    cleaned_orders = clean_orders(raw_orders)
    return cleaned_orders

def get_raw_orders(source_file:str) -> list:
    '''returns raw orders as list of dicts for each order in txt source_file'''
    with open(source_file, 'r', encoding='utf-8') as f:
        source_contents = csv.DictReader(f, delimiter='\t')
        raw_orders = [{header : value for header, value in row.items()} for row in source_contents]
    return raw_orders

def clean_orders(orders:list) -> list:
    '''replaces plus sign in phone numbers with 00'''
    for order in orders:
        try:
            order['buyer-phone-number'] = order['buyer-phone-number'].replace('+', '00')
        except KeyError as e:
            logging.warning(f'New header in source file. Alert VBA on new header. Error: {e}')
            print(VBA_KEYERROR_ALERT)
    return orders

def parse_export_orders(testing:bool, parse_orders:list, loaded_txt:str):
    '''interacts with classes (ParseOrders, OrdersDB) to filter new orders, export desired files and push new orders to db'''
    db_client = OrdersDB(parse_orders, loaded_txt)
    new_orders = db_client.get_new_orders_only()
    logging.info(f'After checking with database, further processing: {len(new_orders)} new orders')
    ParseOrders(new_orders, db_client, SKU_MAPPING_DICT).export_orders(testing)

def parse_args():
    '''accepts txt_path from as command line argument'''
    if len(sys.argv) == EXPECTED_SYS_ARGS:
        txt_path = sys.argv[1]
        logging.info(f'Accepted sys args on launch: txt_path: {txt_path}')
        return txt_path
    else:
        print(VBA_ERROR_ALERT)
        logging.critical(f'Error parsing arguments on script initialization in cmd. Arguments provided: {len(sys.argv)} Expected: {EXPECTED_SYS_ARGS}')
        sys.exit()


def main(testing, amazon_export_txt_path):
    '''Main function executing parsing of provided txt file and exporting labels summary file'''
    logging.info(f'\n NEW RUN STARTING: {datetime.today().strftime("%Y.%m.%d %H:%M")}')    
    if not testing:
        txt_path = parse_args()
    else:
        print('RUNNING IN TESTING MODE')
        txt_path = amazon_export_txt_path
    if os.path.exists(txt_path):
        cleaned_source_orders = get_cleaned_orders(txt_path)
        parse_export_orders(testing, cleaned_source_orders, txt_path)
        print(VBA_OK)
    else:
        logging.critical(f'Provided file {txt_path} does not exist.')
        print(VBA_ERROR_ALERT)
        sys.exit()
    logging.info(f'\nRUN ENDED: {datetime.today().strftime("%Y.%m.%d %H:%M")}\n\n')


if __name__ == "__main__":
    main(testing=TESTING, amazon_export_txt_path=TEST_AMZN_EXPORT_TXT)  
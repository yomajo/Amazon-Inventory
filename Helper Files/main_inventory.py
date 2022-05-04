import sqlalchemy.sql.default_comparator    #neccessary for executable packing
from constants import VBA_ERROR_ALERT, VBA_KEYERROR_ALERT, VBA_OK, EXPECTED_SALES_CHANNELS, AMAZON_KEYS, ETSY_KEYS
from amzn_parser_utils import get_output_dir, is_windows_machine, split_sku, get_country_code, dump_to_json, delete_file
from parse_orders import ParseOrders
from database import SQLAlchemyOrdersDB
from datetime import datetime
import logging
import sys
import csv
import os

# Logging config:
log_path = os.path.join(get_output_dir(client_file=False), 'inventory.log')
logging.basicConfig(handlers=[logging.FileHandler(log_path, 'a', 'utf-8')], level=logging.INFO)

# GLOBAL VARIABLES
TESTING = False
SALES_CHANNEL = 'Amazon'
EXPECTED_SYS_ARGS = 3

if is_windows_machine():
    # ORDERS_SOURCE_FILE = r'C:\Coding\Ebay\Working\Backups\Etsy\EtsySoldOrders2022-4 (2).csv'    
    ORDERS_SOURCE_FILE = r'C:\Coding\Ebay\Working\Backups\Amazon exports\EU 2022.05.04.txt'
    # ORDERS_SOURCE_FILE = r'C:\Coding\Ebay\Working\Backups\Amazon exports\COM 2021.12.07.txt'
else:
    ORDERS_SOURCE_FILE = r'/home/devyo/Coding/Git/Amazon Inventory/Amazon exports/run1.txt'


def get_cleaned_orders(source_file:str, sales_channel:str, proxy_keys:dict) -> list:
    '''returns cleaned orders (as cleaned in clean_orders func) from source_file arg path'''
    delimiter = ',' if sales_channel == 'Etsy' else '\t'
    raw_orders = get_raw_orders(source_file, delimiter)
    if TESTING:
        replace_old_testing_json(raw_orders, 'DEBUG_raw_orders.json')
    cleaned_orders = clean_orders(raw_orders, sales_channel, proxy_keys)
    return cleaned_orders

def get_raw_orders(source_file:str, delimiter:str) -> list:
    '''returns raw orders as list of dicts for each order in txt source_file'''
    with open(source_file, 'r', encoding='utf-8') as f:
        source_contents = csv.DictReader(f, delimiter=delimiter)
        raw_orders = [{header : value for header, value in row.items()} for row in source_contents]
    return raw_orders

def replace_old_testing_json(raw_orders, json_fname:str):
    '''deletes old json, exports raw orders to json file'''
    output_dir = get_output_dir(client_file=False)
    json_path = os.path.join(output_dir, json_fname)
    delete_file(json_path)
    dump_to_json(raw_orders, json_fname)

def clean_orders(orders:list, sales_channel:str, proxy_keys:dict) -> list:
    '''performs universal data cleaning for amazon and etsy raw orders data'''
    for order in orders:
        try:
            # split sku for each order without replacing original keys. sku str value replaced by list of skus
            order[proxy_keys['sku']] = split_sku(order[proxy_keys['sku']], sales_channel)
            if sales_channel == 'Etsy':
                # transform etsy country (Lithuania) to country code (LT)
                country = order[proxy_keys['ship-country']]
                order[proxy_keys['ship-country']] = get_country_code(country)
        except KeyError as e:
            logging.critical(f'Failed while cleaning loaded orders. Last order: {order} Err: {e}')
            print(VBA_KEYERROR_ALERT)
            sys.exit()
    return orders


def parse_args():
    '''returns arguments passed from VBA or hardcoded test environment'''
    if TESTING:
        print('--- RUNNING IN TESTING MODE. Using hardcoded args---')
        logging.warning('--- RUNNING IN TESTING MODE. Using hardcoded args---')
        assert SALES_CHANNEL in EXPECTED_SALES_CHANNELS, f'Unexpected sales_channel value passed from VBA side: {SALES_CHANNEL}'
        return ORDERS_SOURCE_FILE, SALES_CHANNEL
    try:
        assert len(sys.argv) == EXPECTED_SYS_ARGS, 'Unexpected number of sys.args passed. Check TESTING mode'
        source_fpath = sys.argv[1]
        sales_channel = sys.argv[2]
        logging.info(f'Accepted sys args on launch: source_fpath: {source_fpath}; sales_channel: {sales_channel}. Whole sys.argv: {list(sys.argv)}')
        assert sales_channel in EXPECTED_SALES_CHANNELS, f'Unexpected sales_channel value passed from VBA side: {sales_channel}'
        return source_fpath, sales_channel
    except Exception as e:
        print(VBA_ERROR_ALERT)
        logging.critical(f'Error parsing arguments on script initialization in cmd. Arguments provided: {list(sys.argv)} Number Expected: {EXPECTED_SYS_ARGS}. Err: {e}')
        sys.exit()

def main():
    '''Main function executing parsing of provided txt file and exporting labels summary file'''
    logging.info(f'\n NEW RUN STARTING: {datetime.today().strftime("%Y.%m.%d %H:%M")}')        
    source_fpath, sales_channel = parse_args()
    proxy_keys = ETSY_KEYS if sales_channel == 'Etsy' else AMAZON_KEYS
    
    # Get cleaned source orders
    cleaned_source_orders = get_cleaned_orders(source_fpath, sales_channel, proxy_keys)
    
    db_client = SQLAlchemyOrdersDB(cleaned_source_orders, source_fpath, sales_channel, proxy_keys, testing=TESTING)
    new_orders = db_client.get_new_orders_only()
    logging.info(f'Loaded file contains: {len(cleaned_source_orders)}. Further processing: {len(new_orders)} orders')

    # Parse orders, export target files
    ParseOrders(new_orders, db_client, sales_channel, proxy_keys).export_orders(TESTING)

    print(VBA_OK)
    logging.info(f'\nRUN ENDED: {datetime.today().strftime("%Y.%m.%d %H:%M")}\n\n')


if __name__ == "__main__":
    main()
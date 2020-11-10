from amzn_parser_utils import get_output_dir
from helper_file import HelperFile
from collections import defaultdict
from datetime import datetime
import logging
import sys
import csv
import os


# GLOBAL VARIABLES
EXPORT_FILE = 'Amazon Inventory Reduction.xlsx'
SHEET_NAME = 'SKU codes'
VBA_ERROR_ALERT = 'ERROR_CALL_DADDY'
VBA_NO_NEW_JOB = 'NO NEW JOB'
VBA_KEYERROR_ALERT = 'ERROR_IN_SOURCE_HEADERS'

class ParseOrders():
    '''Input: orders as list of dicts, parses orders, forms sorted export obj;
    exports helper txt file with custom labels and corresponding quantities
    Interacts with database client instance; main method:
    
    export_orders(testing=False)    NOTE: check behaviour when testing flag is True in export_orders'''
    
    def __init__(self, all_orders : list, db_client : object):
        self.all_orders = all_orders
        self.db_client = db_client
    
    def _prepare_filepath(self):
        '''constructs cls variable of output abs file path'''
        output_dir = get_output_dir()
        self.inventory_file = os.path.join(output_dir, EXPORT_FILE)
    
    def get_sorted_export_obj(self, orders:list):
        '''returns sorted object: export_obj = {'sku1': {
                                                    'item': 'item_name1',
                                                    'quantity: 2},
                                                'sku3': {
                                                    'item': 'item_name2',
                                                    'quantity: 5}, ...}'''
        export_obj = {}
        for order in orders:
            try:
                # Retrieving order values of interest:
                item = order['product-name']
                sku = order['sku']
                quantity = self.get_order_quantity(order)
            except KeyError:
                logging.exception(f'Could not find \'sku\'or \'product-name\' in order keys. Order: {order}\nClosing connection to database, alerting VBA, exiting...')
                self.db_client.close_connection()
                print(VBA_KEYERROR_ALERT)
                sys.exit()

            if sku not in export_obj.keys():
                export_obj[sku] = {'item':item, 'quantity':quantity}
            else:
                export_obj[sku]['quantity'] += quantity
        self.exit_no_new_orders(export_obj)
        return self._sort_labels(export_obj)

    @staticmethod
    def _sort_labels(labels:dict) -> list:
        '''sorts default dict by descending quantities. Returns list of tuples'''
        return sorted(labels.items(), key=lambda sku_dict: sku_dict[1]['quantity'], reverse=True)

    def get_order_quantity(self, order:dict) -> int:
        '''returns 'quantity-purchased' order key value in integer form'''
        try:
            return int(order['quantity-purchased'])
        except ValueError:
            logging.exception(f'Could convert quantity to integer. Order: {order}\nClosing connection to database, alerting VBA, exiting...')
            self.db_client.close_connection()
            print(VBA_ERROR_ALERT)
            sys.exit()
        except KeyError:
            logging.exception(f'No such key \'quantity-purchased\'. Order: {order}\nClosing connection to database, alerting VBA, exiting...')
            self.db_client.close_connection()
            print(VBA_KEYERROR_ALERT)
            sys.exit()

    def exit_no_new_orders(self, export_obj):
        '''Suspend program, warn VBA if no new orders were found'''
        if len(export_obj) == 0:
            logging.info(f'No new orders found. Terminating, closing database connection, alerting VBA.')
            self.db_client.close_connection()
            print(VBA_NO_NEW_JOB)
            sys.exit()

    def export_inventory_helper_file(self):
        '''creates HelperFile instance, and exports data in xlsx format'''
        export_obj = self.get_sorted_export_obj(self.all_orders)
        try:
            HelperFile(export_obj).export(self.inventory_file)
            os.startfile(self.inventory_file)
            logging.info(f'Helper file {os.path.basename(self.inventory_file)} successfully created.')
        except:
            logging.exception(f'Unexpected error creating helper file. Closing database connection, alerting VBA, exiting...')
            self.db_client.close_connection()
            print(VBA_ERROR_ALERT)
            sys.exit()
        
    def push_orders_to_db(self):
        '''adds all orders in this class to orders table in db'''
        count_added_to_db = self.db_client.add_orders_to_db()
        logging.info(f'Total of {count_added_to_db} new orders have been added to database, after exports were completed')

    def export_orders(self, testing=False):
        '''Summing up tasks inside ParseOrders class'''
        self._prepare_filepath()
        if testing:
            logging.info(f'Testing mode: {testing}. Change behaviour in export_orders method in ParseOrders class')
            print(f'Testing mode: {testing}. Change behaviour in export_orders method in ParseOrders class')
            print('ENABLED REPORT EXPORT WHILE TESTING')
            self.export_inventory_helper_file()
            # self.push_orders_to_db()
            return
        self.export_inventory_helper_file()
        self.push_orders_to_db()

if __name__ == "__main__":
    pass
from amzn_parser_utils import get_output_dir
from collections import defaultdict
from datetime import datetime
import logging
import sys
import csv
import os


# GLOBAL VARIABLES
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
        date_stamp = datetime.today().strftime("%Y.%m.%d %H.%M")
        self.inventory_file = os.path.join(output_dir, f'Amazon Inventory Reduction {date_stamp}.txt')
    
    def get_orders_labels_obj(self) -> dict:
        '''returns a dictionary/counter of unique codes in form: {'code1':5, 'code2':3, ... 'coden':2}'''    
        self.labels = defaultdict(int)
        for order in self.all_orders:
            try:
                self.labels[order['sku']] += self.get_order_quantity(order)
            except KeyError:
                logging.exception(f'Could not find \'sku\' in order keys. Order: {order}\nClosing connection to database, alerting VBA, exiting...')
                self.db_client.close_connection()
                print(VBA_KEYERROR_ALERT)
                sys.exit()
        self.exit_no_new_orders()
        return self._sort_labels(self.labels)
    
    @staticmethod
    def _sort_labels(labels:dict) -> list:
        '''sorts default dict by descending quantities. Returns list of tuples'''
        return sorted(labels.items(), key=lambda lab_qty_tuple: lab_qty_tuple[1], reverse=True)

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

    def exit_no_new_orders(self):
        '''Suspend program, warn VBA if no new orders were found'''
        logging.info(f'INSIDE EXIT_NO_NEW_ORDERS. Length of self.labels: {len(self.labels)}. type validation: {type(self.labels)}')
        if len(self.labels) == 0:
            logging.info(f'No new orders found. Terminating, closing database connection, alerting VBA.')
            self.db_client.close_connection()
            print(VBA_NO_NEW_JOB)
            sys.exit()

    def export_inventory_helper_file(self):
        '''exports txt file with sorted custom labels (sku's) by quantity'''
        export_labels = self.get_orders_labels_obj()
        with open(self.inventory_file, 'w', encoding='utf-8') as f:
            for label, qty in export_labels:
                line = f'{label:<20} {qty:>4}\n'
                f.write(line)
        logging.info(f'Export completed! TXT inventory helper file {os.path.basename(self.inventory_file)} successfully created.')
        
    def push_orders_to_db(self):
        '''adds all orders in this class to orders table in db'''
        count_added_to_db = self.db_client.add_orders_to_db()
        logging.info(f'Total of {count_added_to_db} new orders have been added to database, after exports were completed')

    def export_orders(self, testing=False):
        '''Summing up tasks inside ParseOrders class'''
        self._prepare_filepath()
        if testing:
            logging.info(f'Due to flag testing value: {testing}. Order export and adding to database suspended. Change behaviour in export_orders method in ParseOrders class')
            print(f'Due to flag testing value: {testing}. Order export and adding to database suspended. Change behaviour in export_orders method in ParseOrders class')
            print('ENABLED REPORT EXPORT WHILE TESTING')
            self.export_inventory_helper_file()
            self.push_orders_to_db()
            return
        self.export_inventory_helper_file()
        self.push_orders_to_db()

if __name__ == "__main__":
    pass
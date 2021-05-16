from .constants import QUANTITY_PATTERN, EXPORT_FILE, SHEET_NAME, VBA_ERROR_ALERT, VBA_NO_NEW_JOB, VBA_KEYERROR_ALERT
from .amzn_parser_utils import get_output_dir, orders_column_to_file, get_inner_quantity_and_custom_label
from .helper_file import HelperFileCreate, HelperFileUpdate
from collections import defaultdict
from datetime import datetime
import logging
import sys
import csv
import os


class ParseOrders():
    '''Input args: orders as list of dicts, db_client instance, sku_mapping as dict.
    
    Parses orders, cleans incoming orders, maps sku-custom labels; corrects quantities; forms export obj;
    exports helper txt file with custom labels and corresponding quantities. Interacts with database client instance; main method:
    
    export_orders(testing=False)    NOTE: check behaviour when testing flag is True in export_orders'''
    
    def __init__(self, all_orders:list, db_client:object, mapping_dict:dict):
        self.all_orders = self.__get_mapped_cleaned_orders(all_orders, mapping_dict)
        self.db_client = db_client
        self.mapping_dict = mapping_dict

    def __get_mapped_cleaned_orders(self, all_orders:list, mapping_dict:dict) -> list:
        '''for self.all_orders: converts quantity to int, maps sku codes to custom labels, extracts and calculates real quantity, correcting custom_label'''
        logging.debug(f'Before mapping and cleaning incoming orders. Count: {len(all_orders)} in ParseOrders __get_mapped_cleaned_orders')
        for order in all_orders:
            quantity_purchased = self.__get_order_quantity(order)
            sku = order['sku']
            # change sku if in mapping keys
            if sku in mapping_dict.keys():
                inner_quantity, inner_code = get_inner_quantity_and_custom_label(mapping_dict[sku], QUANTITY_PATTERN)                
                logging.debug(f'Entered mapping. Old sku: {sku} New sku: {inner_code}. Original q-ty: {quantity_purchased}, Recog. q-ty: {inner_quantity}')
            else:
                # No mapping sku-custom_label pair, but attempt to correct quantities for quantity hidden in sku/custom_label
                inner_quantity, inner_code = get_inner_quantity_and_custom_label(sku, QUANTITY_PATTERN)
                logging.debug(f'No mapping. Inner code: {inner_code}, Recog. q-ty: {inner_quantity}')

            # Corrected quantity = 'quantity-purchased' in order dict (source txt column) * extracted quantity inside custom label
            corrected_quantity = quantity_purchased * inner_quantity

            order['sku'] = inner_code
            order['quantity-purchased'] = corrected_quantity
            logging.debug(f'Correcting to quantity: {corrected_quantity}')

        logging.info(f'Finished mapping and correcting q-ties for incoming orders. Returning total {len(all_orders)} orders')
        return all_orders

    def __get_order_quantity(self, order:dict) -> int:
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

    def _prepare_filepath(self):
        '''constructs cls variable of output abs file path'''
        output_dir = get_output_dir()
        self.inventory_file = os.path.join(output_dir, EXPORT_FILE)
    
    def get_export_obj(self, orders:list) -> dict:
        '''returns export object: export_obj = {'sku1': {
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
                quantity = order['quantity-purchased']
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
        return export_obj

    def exit_no_new_orders(self, export_obj):
        '''Suspend program, warn VBA if no new orders were found'''
        if len(export_obj) == 0:
            logging.info(f'No new orders found. Terminating, closing database connection, alerting VBA.')
            self.db_client.close_connection()
            print(VBA_NO_NEW_JOB)
            sys.exit()

    def export_update_inventory_helper_file(self):
        '''Depending on file existence CREATES or UPDATES helper file via different functions'''
        
        export_obj = self.get_export_obj(self.all_orders)

        if os.path.exists(self.inventory_file):
            logging.debug(f'{self.inventory_file} found. Updating...')
            self.update_inventory_file(export_obj)
        else:
            logging.debug(f'{self.inventory_file} not found. Creating file from scratch...')
            self.create_inventory_file(export_obj)
    
    def update_inventory_file(self, export_obj:dict):
        '''creates HelperFileUpdate instance, and updates data in self.inventory_file xlsx file'''
        try:
            HelperFileUpdate(export_obj, self.mapping_dict).update_workbook(self.inventory_file)
            logging.info(f'Helper file {os.path.basename(self.inventory_file)} successfully updated, opening....')
            os.startfile(self.inventory_file)
        except Exception as e:
            logging.exception(f'Unexpected error UPDATING helper file. Closing database connection, alerting VBA, exiting... Last error: {e}')
            self.db_client.close_connection()
            print(VBA_ERROR_ALERT)
            sys.exit()

    def create_inventory_file(self, export_obj:dict):
        '''creates HelperFileCreate instance, and exports data in xlsx format'''
        try:
            HelperFileCreate(export_obj, self.mapping_dict).export(self.inventory_file)
            logging.info(f'Helper file {os.path.basename(self.inventory_file)} successfully created, opening...')
            os.startfile(self.inventory_file)
        except Exception as e:
            logging.exception(f'Unexpected error CREATING helper file. Closing database connection, alerting VBA, exiting... Last error: {e}')
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
            self.export_update_inventory_helper_file()
            # self.push_orders_to_db()
            return
        self.export_update_inventory_helper_file()
        self.push_orders_to_db()

if __name__ == "__main__":
    pass
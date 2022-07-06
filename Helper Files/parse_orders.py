import logging
import sys
import os
from datetime import datetime
from utils import get_output_dir, get_inner_qty_sku, get_order_quantity, dump_to_json
from utils import delete_file, export_invalid_order_ids
from helper_file import HelperFileCreate, HelperFileUpdate
from sku_mapping import SKUMapping
from constants import QUANTITY_PATTERN, EXPORT_FILE, SKU_MAPPING_WB_NAME
from constants import VBA_ERROR_ALERT, VBA_NO_NEW_JOB, VBA_KEYERROR_ALERT


class ParseOrders():
    '''Parses and prepares orders with Helper File generation / update.
    
    Args:
    - orders:list - list of orders
    - db_client:object - instance of database class
    - sales_channel:str - 'Etsy' / 'Amazon' / 'Amazon Warehouse'
    - proxy_keys:dict - keys mapping specific to sales channel
    
    Main method:

    - export_orders(testing=False)
    
    sorts orders to valid/invalid, exports invalid as separate text file, 
    
    NOTE: check behaviour when testing flag is True in export_orders'''
    
    def __init__(self, orders:list, db_client:object, sales_channel:str, proxy_keys:dict):
        self.orders = orders
        self.db_client = db_client
        self.sales_channel = sales_channel
        self.proxy_keys = proxy_keys

        self.quantity_pattern = QUANTITY_PATTERN[sales_channel]
        self.__get_fpaths()

    def __get_fpaths(self):
        '''initiates filepaths used to refer to or create files'''
        timestamp = datetime.today().strftime("%Y.%m.%d %H.%M")
        invalid_orders_fname = f'{self.sales_channel} invalid_orders {timestamp}.txt'
        
        self.sku_mapping_fpath = os.path.join(get_output_dir(client_file=False), SKU_MAPPING_WB_NAME)
        self.inventory_file = os.path.join(get_output_dir(client_file=True), EXPORT_FILE)
        self.invalid_orders_fpath = os.path.join(get_output_dir(client_file=True), invalid_orders_fname)

    def export_orders(self, testing=False):
        '''Summing up tasks inside ParseOrders class'''
        if testing:
            self.__delete_debug_jsons()
            dump_to_json(self.orders, 'DEBUG_new_unparsed.json')

        valid_orders, invalid_orders = self._parse_based_on_sales_channel()
        logging.info(f'Orders inside valid: {len(valid_orders)}; invalid: {len(invalid_orders)}')
        self._exit_no_new_valid_orders(valid_orders, invalid_orders)
        export_obj = self.get_export_obj(valid_orders)
        
        if testing:
            # CHANGE BEHAVIOR WHEN TESTING HERE
            logging.info(f'Testing mode: {testing}. Change behaviour in export_orders method in ParseOrders class')
            print(f'Testing mode: {testing}. Change behaviour in export_orders method in ParseOrders class')
            dump_to_json(valid_orders, 'DEBUG_valid_parsed.json')
            dump_to_json(invalid_orders, 'DEBUG_invalid_orders.json')
            self.export_update_inventory_helper_file(export_obj)
            self.push_orders_to_db()
            return

        self.export_update_inventory_helper_file(export_obj)
        self.push_orders_to_db()

    def __delete_debug_jsons(self):
        '''deletes three json files from previous program run in testing mode'''
        output_dir = get_output_dir(client_file=False)
        files_to_delete = ['DEBUG_new_unparsed.json', 'DEBUG_valid_parsed.json', 'DEBUG_invalid_orders.json']
        for json_f in files_to_delete:
            json_path = os.path.join(output_dir, json_f)
            delete_file(json_path)
        logging.debug(f'Old json files deleted. Ready for debugging')

    def _parse_based_on_sales_channel(self):
        '''parses orders based on sales channel'''
        valid_orders, invalid_orders = [], []
        if self.sales_channel == 'Etsy':
            return self._parse_etsy_orders(valid_orders, invalid_orders)
        else:
            # AmazonCOM / AmazonEU / Amazon Warehouse
            return self._parse_amazon_orders(valid_orders, invalid_orders)

    def _parse_etsy_orders(self, valid_orders:list, invalid_orders:list):
        '''returns valid_orders and invalid_orders lists based on ability to correctly calculate etsy sku quantities'''
        for order in self.orders:
            qty_purchased = get_order_quantity(order, self.proxy_keys)
            skus = order[self.proxy_keys['sku']]
            if len(skus) > 1 and qty_purchased > 1 and qty_purchased != len(skus):
                logging.info(f'Etsy order q-ty and skus may yield various combinations. Qty: {qty_purchased}, skus: {skus}. Ordr being added to invalid list')
                invalid_orders.append(order)
            else:
                if len(skus) == qty_purchased:
                    # order having 7 skus in order will have qty_purchased 7. In reality 7 items were purchased w/ individual quantity = 1 
                    qty_purchased = 1

                parsed_order = self._parse_etsy_order_qty_skus(order, qty_purchased, skus)
                valid_orders.append(parsed_order)
        return valid_orders, invalid_orders

    def _parse_etsy_order_qty_skus(self, order:dict, qty_purchased:int, skus:list):
        '''returns order with added key: 'sku_quantities', value: dict for each sku and matching real parsed quantity'''
        sku_qties = {}
        for sku in skus:
            inner_qty, inner_sku = get_inner_qty_sku(sku, self.quantity_pattern)
            real_sku_qty = inner_qty * qty_purchased
            sku_qties[inner_sku] = real_sku_qty
        order['sku_quantities'] = sku_qties
        return order
    
    def _parse_amazon_orders(self, valid_orders:list, invalid_orders:list):
        '''returns valid_orders and invalid_orders lists for Amazon orders. In unlikely error when parsing, adds order to invalid_orders list'''
        sku_mapping = SKUMapping(self.sku_mapping_fpath).read_sku_mapping_to_dict()
        for order in self.orders:
            qty_purchased = get_order_quantity(order, self.proxy_keys)
            skus = order[self.proxy_keys['sku']]
            try:
                parsed_order = self._parse_amazon_order_qty_skus(order, qty_purchased, skus, sku_mapping)
                valid_orders.append(parsed_order)
            except Exception as e:
                logging.critical(f'Unexpected error while parsing amazon order: {order} Err: {e}. Adding to invalid orders list')
                invalid_orders.append(order)
        return valid_orders, invalid_orders

    def _parse_amazon_order_qty_skus(self, order:dict, qty_purchased:int, skus:list, sku_mapping:dict):
        '''returns order with added key: 'sku_quantities', value: dict for each sku and matching real parsed quantity.
        Attempts to use mapped sku if found in sku_mapping'''
        sku_qties = {}
        for sku in skus:
            # attempt to find matching Shop4Top sku for every amazon original sku before parsing
            if sku in sku_mapping:
                logging.debug(f'Mapping match found for code: {sku}, match: {sku_mapping[sku]}')
                sku = sku_mapping[sku]
            
            inner_qty, inner_sku = get_inner_qty_sku(sku, self.quantity_pattern)
            real_sku_qty = inner_qty * qty_purchased
            sku_qties[inner_sku] = real_sku_qty
        order['sku_quantities'] = sku_qties
        return order

    def _exit_no_new_valid_orders(self, valid_orders:list, invalid_orders:list):
        '''Suspend program, warn VBA if no new orders were found'''
        if not valid_orders and not invalid_orders:
            logging.info(f'No new orders found. Terminating, closing database connection, alerting VBA.')
            self.db_client.session.close()
            print(VBA_NO_NEW_JOB)
            sys.exit()
        elif not valid_orders:
            # invalid orders present
            self._export_invalid_orders_start_file(invalid_orders)
            print(VBA_NO_NEW_JOB)
        else:
            # 1+ valid orders present and 0 or more invalid
            self._export_invalid_orders_start_file(invalid_orders)

    def _export_invalid_orders_start_file(self, invalid_orders:list):
        '''exports invalid order IDs to txt file and opens it'''
        if invalid_orders:
            export_invalid_order_ids(invalid_orders, self.proxy_keys, self.invalid_orders_fpath)
            os.startfile(self.invalid_orders_fpath)
            logging.info(f'Invalid orders exported at {self.invalid_orders_fpath} and opened.')

    def get_export_obj(self, orders:list) -> dict:
        '''returns export object: export_obj = {'sku1': qty1, 'sku2': qty2, ...}'''
        export_obj = {}
        for order in orders:
            try:
                sku_qties = order['sku_quantities']
                for sku in sku_qties:
                    # Add new sku or add its quantity to existing sku
                    if sku not in export_obj:
                        export_obj[sku] = sku_qties[sku]
                    else:
                        export_obj[sku] += sku_qties[sku]
            except Exception as e:
                logging.exception(f'Err: {e} creating export_obj. Last order: {order}\nClosing connection to database, alerting VBA, exiting...')
                self.db_client.session.close()
                print(VBA_KEYERROR_ALERT)
                sys.exit()
        return export_obj

    def export_update_inventory_helper_file(self, export_obj:dict):
        '''Depending on file existence CREATES or UPDATES helper file via different functions'''
        if export_obj:
            if os.path.exists(self.inventory_file):
                logging.debug(f'{self.inventory_file} found. Updating...')
                self.update_inventory_file(export_obj)
            else:
                logging.debug(f'{self.inventory_file} not found. Creating file from scratch...')
                self.create_inventory_file(export_obj)
        else:
            logging.info(f'Formed export_obj is empty. Helper File Creation / Update bypassed.')
    
    def update_inventory_file(self, export_obj:dict):
        '''creates HelperFileUpdate instance, and updates data in self.inventory_file xlsx file'''
        try:
            HelperFileUpdate(export_obj).update_workbook(self.inventory_file)
            logging.info(f'Helper file {os.path.basename(self.inventory_file)} successfully updated, opening....')
            os.startfile(self.inventory_file)
        except Exception as e:
            logging.exception(f'Unexpected error UPDATING helper file. Closing database connection, alerting VBA, exiting... Last error: {e}')
            self.db_client.session.close()
            print(VBA_ERROR_ALERT)
            sys.exit()

    def create_inventory_file(self, export_obj:dict):
        '''creates HelperFileCreate instance, and exports data in xlsx format'''
        try:
            HelperFileCreate(export_obj).export(self.inventory_file)
            logging.info(f'Helper file {os.path.basename(self.inventory_file)} successfully created, opening...')
            os.startfile(self.inventory_file)
        except Exception as e:
            logging.exception(f'Unexpected error CREATING helper file. Closing database connection, alerting VBA, exiting... Last error: {e}')
            self.db_client.session.close()
            print(VBA_ERROR_ALERT)
            sys.exit()
        
    def push_orders_to_db(self):
        '''adds all orders in this class to orders table in db'''
        count_added_to_db = self.db_client.add_orders_to_db()
        logging.info(f'Total of {count_added_to_db} new orders have been added to database, after exports were completed, closing connection to DB')
        self.db_client.session.close()



if __name__ == "__main__":
    pass
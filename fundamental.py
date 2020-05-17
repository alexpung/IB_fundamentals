import datetime
import logging
import os.path
import time

import get_jp_from_excel as excel
from ibapi import wrapper
from ibapi.client import EClient
from ibapi.common import *  # @UnusedWildImport
from ibapi.contract import *  # @UnusedWildImport

# CONSTANTS ##################
DIRECTORY = './data/'
REPORT_TYPE = 'ReportSnapshot'


##################


def create_japan_stock(symbol: str) -> Contract:
    contract = Contract()
    contract.symbol = symbol
    contract.secType = "STK"
    contract.currency = "JPY"
    contract.exchange = "TSEJ"
    return contract


class TestApp(wrapper.EWrapper, EClient):
    def __init__(self, ipaddress, portid, clientid, stock_list):
        wrapper.EWrapper.__init__(self)
        EClient.__init__(self, self)
        self.connect(ipaddress, portid, clientid)
        self.target = None
        self.stock_list = stock_list
        self.reqId = 1

    def nextValidId(self, orderId: int):
        app.request_fundamental()

    def request_fundamental(self):
        if self.stock_list:
            time.sleep(10)
            self.target = self.stock_list.pop()
            logging.debug(f'Requesting data for {self.target}')
            self.reqFundamentalData(self.reqId, create_japan_stock(self.target), REPORT_TYPE, [])
            self.reqId += 1
        else:
            app.disconnect()

    def fundamentalData(self, req_id, data):
        super().fundamentalData(req_id, data)
        date_stamp = datetime.date.today()
        filename = f'{self.target}_{date_stamp}_{REPORT_TYPE}.xml'
        file_path = os.path.join(DIRECTORY, filename)
        if not os.path.isdir(DIRECTORY):
            os.mkdir(DIRECTORY)
        with open(file_path, 'w') as f:
            f.write(data)
        self.request_fundamental()

    def error(self, reqId: TickerId, errorCode: int, errorString: str):
        super().error(reqId, errorCode, errorString)
        logging.error(f"Error Id: {reqId}, Code: {errorCode}, Msg:{errorString}")
        if errorCode == 430 and self.stock_list:
            # We are sorry, but fundamentals data for the security specified is not available.failed to fetch
            app.request_fundamental()
        if errorCode == 200 and self.stock_list:  # No security definition has been found for the request
            app.request_fundamental()


if __name__ == '__main__':
    # Check that the port is the same as on the Gateway
    # ipaddress is 127.0.0.1 if one same machine, clientid is arbitrary
    logging.basicConfig(level=logging.DEBUG, filename='app.log', filemode='w',
                        format='%(asctime)s - %(message)s', datefmt='%d-%b-%y %H:%M:%S')
    logging.getLogger().addHandler(logging.StreamHandler(sys.stdout))
    app = TestApp("127.0.0.1", 7496, 1, excel.get_jp_list('data_jp.xls'))
    app.run()

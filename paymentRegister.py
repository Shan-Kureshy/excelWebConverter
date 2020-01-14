import sys
import re
from openpyxl import Workbook
import os.path
from report import Report

class Item(object):
    def __init__(self, line):
        self.ref = None
        self.type = None
        self.enty = None
        self.account = None
        self.subAccount = None
        self.cashAmt = None
        self.discountAmt = None
        self.ArAmt = None
        self.unassignedAmt = None
        self.nonArAmt = None

class Account(object):
    def __init__(self):
        pass

class Batch(object):
    def __init__(self):
        self.num = None

class PaymentRegister(Report):
    def __init__(self, filePath):
        super().__init__(filePath)
        self.cleanReport()

    def cleanReport(self):
        toRemove = [0,1]
        for i in range(2,len(self.data)):
            if i not in toRemove:

                if re.findall(r'End of Report', self.data[i]):
                    toRemove+= [x for x in range(i, i+18)]
                    break
                elif re.findall(r'Payment Register', self.data[i]):
                    toRemove+= [x for x in range(i, i+4)]
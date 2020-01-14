import sys
import re
from openpyxl import Workbook
import os.path
from report import Report

class Vendor(object):
    def __init__(self, line1, line2, line3=None):
        self.info = {
            'addressLine1':None,
            'addressLine2':None,
            'vendorNum':None,
            'name':None,
            'attention':None,
            'city':None,
            'state':None,
            'country':None,
            'postal':None,
            'phone':None,
            'ext':None
        }

        self.getData(line1, line2, line3)

    # if address doesnt have number or po box, add to name
    # address line 1 is attention, line 2 & 3 is 
    def getData(self, line1, line2, line3):
        try:
            self.info['vendorNum'] = line1[:8]
            self.info['name'] = line1[9:37]
            self.info['addressLine1'] =line1[38:66]
            self.info['city'] = line1[67:91]
            self.info['state'] = line1[92:96]
            self.info['postal'] = line1[97:107]
            self.info['phone'] = line1[108:124]
            self.info['ext'] = line1[125:130]

            self.info['attention'] = line2[14:37]
            self.info['addressLine2'] = line2[38:66]
            self.info['country'] = line2[67:91]
        except IndexError:
            pass

        # remove trailing spaces from all values
        for key in self.info:
            self.info[key] = self.info[key].rstrip(' \n')

        # Check if address is correctly formatted
        if not re.findall(r'[0-9]|[Pp]\.? ?[Oo]', self.info['addressLine1']):
            self.info['name'] += self.info['addressLine1']
            self.info['addressLine1'] = self.info['addressLine2']
        if line3:
            self.info['addressLine2'] = line3[38:66]

    def getLine(self):
        return ['', self.info['name'], self.info['vendorNum'], self.info['name'], self.info['name'],'','','','',
                self.info['attention'], self.info['addressLine1'], self.info['addressLine2'], '', self.info['city'],
                self.info['state'], self.info['postal'], self.info['country'], '','', self.info['phone']]


class VendorList(Report):
    def __init__(self, filePath):
        super().__init__(filePath)
        self.vendors = []
        self.cleanReport()
        self.getVendors()

    def cleanReport(self):
        toRemove = []
        for i in range(0, len(self.data)):
            if i not in toRemove:
                if re.findall(r'Supplier Address Report', self.data[i]):
                    toRemove += [x for x in range(i, i+4)]
                elif re.findall(r'End of Report', self.data[i]):
                    toRemove += [x for x in range(i, i+9)]
        for i in reversed(toRemove):
            self.data.pop(i)
        pass

    def getVendors(self):
        firstLine = self.data[0]
        secondLine = None
        thirdLine = None

        for line in self.data[1:]:
            if re.match(r'^[0-9]{8}.*$', line):
                # add the last vendor
                if firstLine and secondLine:
                    if thirdLine:
                        self.vendors.append(Vendor(firstLine, secondLine, thirdLine))
                    else:
                        self.vendors.append(Vendor(firstLine, secondLine))
                    secondLine = None
                    thirdLine = None

                firstLine = line
            elif re.match(r'^ {1,10}Attn:.*$', line):

                secondLine = line
            else:
                thirdLine = line


        pass

    def createSpreadsheet(self, spreadsheetName):
        global currentLine
        currentLine = 1
        wb = Workbook()
        ws = wb.active

        def addRow(data=[''], isDataItem=False):
            global currentLine
            maxCol = len(data)
            for j in range(0, maxCol):
                ws.cell(row=currentLine, column=j + 1, value=data[j])
            currentLine += 1

        for vendor in self.vendors:
            addRow(vendor.getLine())

        wb.save(spreadsheetName +'.xlsx')


def _runTest():
    test = VendorList(r'C:\Users\Shan\Desktop\Work\vendor list\QAD vendor list')
    test.createSpreadsheet('vendorTest.xlsx')

if __name__ == '__main__':
    _runTest()
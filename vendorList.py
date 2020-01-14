import sys
import re
from openpyxl import Workbook
import os.path
from report import Report


class Vendor(object):
    def __init__(self, line1, line2, line3=None, bounds=None):
        self.info = {
            'addressLine1': None,
            'addressLine2': None,
            'vendorNum': None,
            'name': None,
            'attention': None,
            'city': None,
            'state': None,
            'country': None,
            'postal': None,
            'phone': None,
            'ext': None
        }

        self.getData(line1, line2, line3, bounds)

    # if address doesnt have number or po box, add to name
    # address line 1 is attention
    def getData(self, line1, line2, line3, bounds):
        try:

            self.info['vendorNum'] = line1[bounds['vendorNum'][0]:bounds['vendorNum'][1]]
            self.info['name'] = line1[bounds['name'][0]:bounds['name'][1]]
            self.info['addressLine1'] = line1[bounds['addressLine1'][0]:bounds['addressLine1'][1]]
            self.info['city'] = line1[bounds['city'][0]:bounds['city'][1]]
            self.info['state'] = line1[bounds['state'][0]:bounds['state'][1]]
            self.info['postal'] = line1[bounds['postal'][0]:bounds['postal'][1]]
            self.info['phone'] = line1[bounds['phone'][0]:bounds['phone'][1]]
            self.info['ext'] = line1[bounds['ext'][0]:bounds['ext'][1]]
            # shares column with name(+5 to remove 'attn:' at beginning of string)
            self.info['attention'] = line2[bounds['attention'][0]:bounds['attention'][1]]
            # shares column with address line 1
            self.info['addressLine2'] = line2[bounds['addressLine2'][0]:bounds['addressLine2'][1]]
            # shares column with city
            self.info['country'] = line2[bounds['country'][0]:bounds['country'][1]]
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
        return ['', self.info['name'], self.info['vendorNum'], self.info['name'], self.info['name'], '', '', '', '',
                self.info['attention'], self.info['addressLine1'], self.info['addressLine2'], '', self.info['city'],
                self.info['state'], self.info['postal'], self.info['country'], '', '', self.info['phone']]


class VendorList(Report):
    def __init__(self, filePath):
        super().__init__(filePath)
        self.vendors = []
        self.getVendorBounds()
        self.cleanReport()
        self.getVendors()

    def getVendorBounds(self):
        boundLine = self.data[3]

        def findnth(n, needle=' '):
            parts = boundLine.split(needle, n + 1)
            if len(parts) <= n + 1:
                return -1
            return len(boundLine) - len(parts[-1]) - len(needle)

        self.bounds = {'vendorNum': None,
                       'name': None,
                       'addressLine1': None,
                       'city': None,
                       'state': None,
                       'postal': None,
                       'phone': None,
                       'ext': None}

        for key, i in zip(self.bounds, range(0, 8)):
            if i is 0:
                self.bounds[key] = (0, findnth(i))
            elif i is 7:
                self.bounds[key] = (findnth(i - 1) + 1, len(boundLine) - 1)
            else:
                self.bounds[key] = (findnth(i - 1) + 1, findnth(i))

        self.bounds['addressLine2'] = self.bounds['addressLine1']
        self.bounds['country'] = self.bounds['city']
        self.bounds['attention'] = (self.bounds['name'][0] + 5, self.bounds['name'][1])

    # remove page headers from report
    def cleanReport(self):
        toRemove = []
        for i in range(0, len(self.data)):
            if i not in toRemove:
                if re.findall(r'Supplier Address Report', self.data[i]):
                    toRemove += [x for x in range(i, i + 4)]
                elif re.findall(r'End of Report', self.data[i]):
                    toRemove += [x for x in range(i, i + 9)]
        for i in reversed(toRemove):
            self.data.pop(i)
        pass

    # pull all vendors from report and add to vendors list
    def getVendors(self):
        firstLine = self.data[0]
        secondLine = None
        thirdLine = None

        for line in self.data[1:]:
            if re.match(r'^[0-9]{8}.*$', line):
                # add the last vendor
                if firstLine and secondLine:
                    if thirdLine:
                        self.vendors.append(Vendor(firstLine, secondLine, thirdLine, bounds=self.bounds))
                    else:
                        self.vendors.append(Vendor(firstLine, secondLine, bounds=self.bounds))
                    secondLine = None
                    thirdLine = None

                firstLine = line
            elif re.match(r'^ {1,10}Attn:.*$', line):

                secondLine = line
            else:
                thirdLine = line

        pass

    # save spreadsheet with vendor data
    def createSpreadsheet(self, spreadsheetName):
        global currentLine
        currentLine = 1
        wb = Workbook()
        ws = wb.active

        def addRow(data=None, isDataItem=False):
            if data is None:
                data = ['']

            global currentLine
            maxCol = len(data)
            for j in range(0, maxCol):
                ws.cell(row=currentLine, column=j + 1, value=data[j])
            currentLine += 1

        for vendor in self.vendors:
            addRow(vendor.getLine())

        wb.save(spreadsheetName + '.xlsx')


def _runTest():
    test = VendorList(r'C:\Users\Shan\Desktop\Work\vendor list\QAD vendor list')
    test.createSpreadsheet('vendorTest')


if __name__ == '__main__':
    _runTest()

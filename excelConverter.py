import os
import trialBalance
import accountDetail
import vendorList
from APAgingReport import *
from ARAgingReport import *

'''
To do next:
add regex to string reading in report, account, and item classes
begin excel functions
'''
class ReportError(Exception):
    print("Unexpected formatting. report failed to convert")

def checkReportType(filePath):
    with open(filePath) as file:
        data = file.readline()
        if re.match(r'^.*AP Aging.*$', data):
            return 'AP'
        elif re.match(r'^.*AR Aging.*$', data):
            return 'AR'
        elif re.match(r'^.*(Account|Acct).*$', data):
            return 'ABD'
        elif re.match(r'^.*Trial.*$', data):
            return 'TBS'
        elif re.match(r'^.*Supplier Address Report.*$', data):
            return 'SAR'
        else:
            return None

def runAPreport(filePath, spreadsheetName):
    report = APreport(filePath)
    report.createSpreadsheet(spreadsheetName)

def runARreport(filePath, spreadsheetName):
    report = ARreport(filePath)
    report.createSpreadsheet(spreadsheetName)

def runTrialBalance(filePath, spreadsheetName):
    trialBalance.createspreadsheet(filePath, spreadsheetName)

def runAccountDetail(filePath, spreadsheetName):
    accountDetail.createSpreadsheet(filePath, spreadsheetName)

def noReportError(filePath=None, spreadsheetName=None):
    raise ReportError()

def runSupplierAddress(filePath, spreadsheetName):
    report = vendorList.VendorList(filePath)
    report.createSpreadsheet(spreadsheetName)

def getSpreadsheet(filePath, spreadsheetName):
    reports = {'AP': runAPreport,
               'AR': runARreport,
               'ABD': runAccountDetail,
               'TBS': runTrialBalance,
               'SAR': runSupplierAddress,
               None: noReportError}
    reportType = checkReportType(filePath)
    try:
        reports[reportType](filePath, spreadsheetName)
    except:
        noReportError()


if __name__ =='__main__':
    print('Excel Converter 0.0.3\n'
          'shank@autogenomics.com\n'
          'Currently accepts: AP/AR aging, Trial Balance Summary, Account Balance Detail\n'
          'How To: download desired report from QAD, save it into the same folder as this program, then\n'
          'type the name of the report and the name you want to give the spreadsheet into the program.\n'
          'If run successfully, the spreadsheet should appear in the same folder as the program\n\n')
    while True:
        filePath = input('Enter a QAD file name, or press enter to exit: ')
        if filePath is '':
            break

        while not os.path.isfile(filePath):

            if os.path.isfile(filePath + '.txt'):
                filePath += '.txt'
            else:
                filePath = input(
                    'Error: File not found. Please make sure file is in the correct folder, and name is being entered correctly.'
                    '\nEnter a QAD file name: ')

        spreadsheetName = input('Enter a name to save spreadsheet as: ')

        getSpreadsheet(filePath, spreadsheetName)
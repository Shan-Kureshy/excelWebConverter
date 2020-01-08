from agingReport import *

class ARdataItem(DataItem):
    def __init__(self, line, accountInfo=None):
        super().__init__()
        if accountInfo:
            self.accountNum = accountInfo[0]
            self.accountName = accountInfo[1]

        # check if line has data
        if line == ('' or '\n') or re.match(r'^[- ]+$', line):
            self.blank = True
            raise IndexError('line is blank')

        # split lines into lists of strings
        self.line = line
        self.parsedLine = line.split(' ')
        while '' in self.parsedLine:
            self.parsedLine.remove('')

        # check if data item is the 'account total' line, which requires different formatting
        if re.findall('Totals:', self.parsedLine[2]):
            self.base = True
            self.under30 = self.parsedLine[3]
            self.over30 = self.parsedLine[4]
            self.over60 = self.parsedLine[5]
            self.over90 = self.parsedLine[6]
            self.total = self.parsedLine[7]

        else:
            # if item contains credit terms
            self.base = False
            self.voucherNum = self.parsedLine[0]
            self.invoiceNum = self.parsedLine[1]

            self.effectiveDate = datetime.strptime(self.parsedLine[2], '%m/%d/%y')
            self.dueDate = datetime.strptime(self.parsedLine[3], '%m/%d/%y')
            self.invoiceDate = datetime.strptime(self.parsedLine[4], '%m/%d/%y')
            if len(self.parsedLine) == 11:
                self.crTerms = self.parsedLine[5]
                self.under30 = self.parsedLine[6]
                self.over30 = self.parsedLine[7]
                self.over60 = self.parsedLine[8]
                self.over90 = self.parsedLine[9]
                self.total = self.parsedLine[10]

            else:  # if item doesnt contain credit terms
                self.crTerms = ''
                self.under30 = self.parsedLine[5]
                self.over30 = self.parsedLine[6]
                self.over60 = self.parsedLine[7]
                self.over90 = self.parsedLine[8]
                self.total = self.parsedLine[9]

        # convert values to right type
        self.under30 = toNum(self.under30)
        self.over30 = toNum(self.over30)
        self.over60 = toNum(self.over60)
        self.over90 = toNum(self.over90)
        self.total = toNum(self.total)


class ARaccount(Account):
    def setAccountHeader(self, line):

        # split account details into list of words
        self.parsedLine = line.split(' ')
        while '' in self.parsedLine:
            self.parsedLine.remove('')
        self.accountNum = self.parsedLine[0]
        self.accountName = ''

        # since multi-word strings get broken up from split, this puts them back together
        for i in range(1, self.parsedLine.index('State:')):
            self.accountName += self.parsedLine[i] + ' '

    def addDataItem(self, line):
        try:
            data = ARdataItem(line, self.getAccountHeader())
            self.dataItems.append(data)
        except IndexError:
            pass


class ARreport(Report):
    def __init__(self, filePath):
        super().__init__(filePath)
        self.date = None
        self.getDate()

        self.totalAging = None
        self.totalAgingExchange = None
        self.invoices = None
        self.memos = None
        self.charges = None
        self.payments = None
        self.drafts = None
        self.getTotalAging()

        self.cleanReport()
        self.getAccounts()

    def checkType(self):
        return 'AR'

    def getTotalAging(self):
        for line in self.data:
            try:
                parsedLine = line.split('  ')

                while '' in parsedLine:
                    parsedLine.remove('')
                parsedLine[0].rstrip(' ')

                if re.match(r'^.*Invoices.*$', parsedLine[0]):
                    self.invoices = parsedLine[1]
                elif re.match(r'^.*Memos.*$', parsedLine[0]):
                    self.memos = parsedLine[1]
                elif re.match(r'^.*Charges.*$', parsedLine[0]):
                    self.charges = parsedLine[1]
                elif re.match(r'^.*Unapplied Payments.*$', parsedLine[0]):
                    self.payments = parsedLine[1]
                elif re.match(r'^.*Drafts.*$', parsedLine[0]):
                    self.drafts = parsedLine[1]
                elif re.match(r'^.*Total USD Aging.*$', parsedLine[0]):
                    self.totalAging = parsedLine[1]
                elif re.match(r'^.*Exchange Rate.*$', parsedLine[0]):
                    self.totalAgingExchange = parsedLine[1]

            except IndexError:
                continue

    # removes page headers from report
    def cleanReport(self):
        # mark which list items to remove
        itemsToRemove = []
        for i in range(0, len(self.data)):
            if "AR Aging as of Effective Date" in self.data[i]:
                for j in range(i, i + 7):  # after removing empty lines, page headers are 7 lines long
                    itemsToRemove.append(j)

            elif "End of Report" in self.data[i]:
                for j in range(i, i + 19):

                    # check if line is already marked to be removed
                    if j not in itemsToRemove:
                        itemsToRemove.append(j)
                break

        # remove marked items
        for index in reversed(itemsToRemove):
            self.data.pop(index)

        # remove empty lines
        while '\n' in self.data:
            self.data.remove('\n')

    # scan through data to retrieve all accounts
    def getAccounts(self):
        self.accounts = []
        newAccount = True
        currentAccount = None
        for line in self.data:
            if newAccount:
                try:
                    currentAccount = ARaccount(line)
                    newAccount = False
                except ValueError:
                    break

            # line is last line of the report, break since there's no more data to be added
            elif re.match(r'^.*USD +Report +Totals:.*$', line):
                break

            else:  # line is second line of data item
                currentAccount.addDataItem(line)

            if currentAccount.isComplete():
                self.accounts.append(currentAccount)
                newAccount = True

if __name__ == "__main__":
    x = ARreport(r'C:\Users\Shan\Desktop\Work\AR aging\AR aging QAD')
    x.createSpreadsheet('ARtest')
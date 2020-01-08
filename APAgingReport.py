from agingReport import *


class APdataItem(DataItem):
    def __init__(self, line1, line2, accountInfo):
        super().__init__()
        self.accountNum = accountInfo[0]
        self.accountName = accountInfo[1]

        # split lines into lists of strings
        self.line1 = line1
        self.parsedLine1 = line1.split(' ')
        while '' in self.parsedLine1:
            self.parsedLine1.remove('')

        self.line2 = line2
        self.parsedLine2 = line2.split(' ')
        while '' in self.parsedLine2:
            self.parsedLine2.remove('')

        # check if data item is the 'account total' line, which requires different formatting
        if self.parsedLine2[0] == 'Base':
            self.base = True
            self.under30 = self.parsedLine2[3]
            self.over30 = self.parsedLine2[4]
            self.over60 = self.parsedLine2[5]
            self.over90 = self.parsedLine2[6]
            self.total = self.parsedLine2[7]

        else:
            if len(self.parsedLine1) == 9:  # if item contains credit terms
                self.base = False
                self.voucherNum = self.parsedLine1[0]
                self.invoiceDate = datetime.strptime(self.parsedLine1[1], '%m/%d/%y')
                self.effectiveDate = datetime.strptime(self.parsedLine1[2], '%m/%d/%y')
                self.crTerms = self.parsedLine1[3]
                self.under30 = self.parsedLine1[4]
                self.over30 = self.parsedLine1[5]
                self.over60 = self.parsedLine1[6]
                self.over90 = self.parsedLine1[7]
                self.total = self.parsedLine1[8]

            else:  # if item doesnt contain credit terms
                self.base = False
                self.voucherNum = self.parsedLine1[0]
                self.invoiceDate = datetime.strptime(self.parsedLine1[1], '%m/%d/%y')
                self.effectiveDate = datetime.strptime(self.parsedLine1[2], '%m/%d/%y')
                self.crTerms = ''
                self.under30 = self.parsedLine1[3]
                self.over30 = self.parsedLine1[4]
                self.over60 = self.parsedLine1[5]
                self.over90 = self.parsedLine1[6]
                self.total = self.parsedLine1[7]

            # use regex to parse 2nd line
            line2match = re.match(r'\s*(?P<invoice>(?:[^ ]+\s)+)\s*(?P<date>[0-9]{2}\/[0-9]{2}\/[0-9]{2})', self.line2)
            self.invoiceNum = line2match.group('invoice')
            self.dueDate = datetime.strptime(line2match.group('date'), '%m/%d/%y')

        # convert values to right type
        self.under30 = toNum(self.under30)
        self.over30 = toNum(self.over30)
        self.over60 = toNum(self.over60)
        self.over90 = toNum(self.over90)
        self.total = toNum(self.total)


class APaccount(Account):
    def setAccountHeader(self, line):
        self.accountAttention = ''
        self.accountPhoneNum = ''
        self.accountName = ''
        # split account details into list of words
        self.parsedLine = line.split(' ')
        while '' in self.parsedLine:
            self.parsedLine.remove('')
        self.accountNum = self.parsedLine[0]

        # , and phone number can be found in indexes right after respective labels
        attentionIndex = self.parsedLine.index('Attention:')
        phoneIndex = self.parsedLine.index('Telephone:')

        # since multi-word strings get broken up from split, this puts them back together
        for i in range(1, attentionIndex):
            self.accountName += self.parsedLine[i] + " "
        for i in range(attentionIndex + 1, phoneIndex):
            self.accountAttention += self.parsedLine[i] + " "
        for i in range(phoneIndex + 1, len(self.parsedLine)):
            self.accountPhoneNum += self.parsedLine[i] + " "

    def addDataItem(self, line1, line2):
        self.dataItems.append(APdataItem(line1, line2, self.getAccountHeader()))



class APreport(Report):
    def __init__(self, filePath):
        super().__init__(filePath)
        self.totalAgingExchange = None
        self.variance = None
        self.getTotalAging()
        self.cleanReport()
        self.getAccounts()

    def checkType(self):
        return 'AP'

    def getTotalAging(self):
        for line in self.data:
            try:
                parsedLine = line.split('  ')

                while '' in parsedLine:
                    parsedLine.remove('')
                parsedLine[0].rstrip(' ')

                if re.match(r'^.*Total USD Aging.*$', parsedLine[0]):
                    self.totalAging = parsedLine[1]
                elif re.match(r'^.*Aging at Exchange Rates.*$', parsedLine[0]):
                    self.totalAgingExchange = parsedLine[1]
                elif re.match(r'^.*Variance.*$', parsedLine[0]):
                    self.variance = parsedLine[1]
            except IndexError:
                continue

    # removes page headers from report
    def cleanReport(self):
        # mark which list items to remove
        itemsToRemove = []
        for i in range(0, len(self.data)):
            if "End of Report" in self.data[i]:
                for j in range(i, i + 16):
                    itemsToRemove.append(j)
                break
            elif "AP Aging as of Effective Date" in self.data[i]:
                for j in range(i, i + 7):
                    itemsToRemove.append(j)

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
        isFirstLine = True
        firstLine = None
        for i in range(0, len(self.data)):
            if newAccount:
                try:
                    currentAccount = APaccount(self.data[i])
                    newAccount = False
                except ValueError:
                    break

            elif isFirstLine:
                firstLine = self.data[i]
                isFirstLine = False

            elif self.data[i][0] == 'Total':  # line is last line of the report, no more data to be added
                break

            else:  # line is second line of data item
                currentAccount.addDataItem(firstLine, self.data[i])
                isFirstLine = True

            if currentAccount.isComplete():
                self.accounts.append(currentAccount)
                newAccount = True

if __name__ == "__main__":
    y = APreport(r'C:\Users\Shan\Desktop\Work\AP aging\AP')
    y.createSpreadsheet('APtest')
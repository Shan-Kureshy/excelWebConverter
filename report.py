

class Report(object):
    def __init__(self, filePath):
        self.filePath = filePath
        with open(self.filePath) as file:
            self.data = []
            for line in file:
                self.data.append(line.rstrip('\n'))
            while '' in self.data:
                self.data.remove('')
        file.close()
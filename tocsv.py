import threading

class tocsv:
    def __init__(self):
        self.lock = threading.Lock()
        self.file = ''

    def append(self,data):
        with self.lock:
            for entry in data:
                self.file = self.file + entry
                if entry is data[-1]:
                    self.file = self.file + "\n"
                else:
                    self.file = self.file+ ","

    def fromlist(self, listoflists):
        for roll, lists in listoflists:
            self.append(lists)

    def getcsv(self):
        return self.file

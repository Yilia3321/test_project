class Table:
    tableName = ''
    colName = ['Parameter', 'Length (Byte)', 'Range/Format', 'Default']
    body = [[]]

    def __init__(self):
        self.deleted = False

    def tostring(self):
        print("\033[0;35;0m                " + self.tableName + "\033[0m")
        print("===========================================")
        for col in self.colName:
            print("\033[0;32;0m" + col + "\033[0m", end='  |  ')
        print()
        print("-------------------------------------------")
        for row in (self.body or []):
            for item in (row or []):
                print(item, end='  |  ')
            print()
            print("-------------------------------------------")

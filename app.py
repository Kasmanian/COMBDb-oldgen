from db import Database

class App:
    def __init__(self):
        # show splash screen (ss)
        # open user.dat file
        # ss++
        # load GUI
        # ss++
        self.database = Database()
        self.database.connect('<.accdb-file>')
        # ss++
        # show GUI
        # destroy ss
        pass
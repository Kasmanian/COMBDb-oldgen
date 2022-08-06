import json, pyodbc, bcrypt
from datetime import date

class Model:
  def __init__(self):
    pass

  def __usesCursor(func):
    def wrap(self, *args, **kwargs):
        try:
          cursor = self.db.cursor()
          result = func(self, cursor, *args, **kwargs)
          self.db.commit()
          return result
        except (Exception, pyodbc.Error) as e:
          print(f'Error in connection: {e}')
          return e
        finally:
          if cursor: cursor.close()
    return wrap
  
  def connect(self):
    try:
      f = open('COMBDb\local.json')
      PATH = json.load(f)['DBQ']
      CONSTR = r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+PATH
      f.close()
      self.db = pyodbc.connect(CONSTR)
      return True
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False

  def close(self):
    if self.db: self.db.close()

  @__usesCursor
  def addTech(self, cursor, first, middle, last, username, password):
    query = ('INSERT INTO Techs(First, Middle, Last, Username, Password, Active) VALUES(?, ?, ?, ?, ?, ?)')
    cursor.execute(query, first, middle, last, username, self.encrypt(password).decode('utf-8'), 'Yes')
    return True

  @__usesCursor 
  def toggleTech(self, cursor, entry, active):
    query = ('UPDATE Techs SET Active=? WHERE Entry=?')
    cursor.execute(query, active, entry)
    return True

  @__usesCursor
  def addClinician(self, cursor, prefix, first, last, designation, phone, fax, email, addr1, addr2, city, state, zip, enrolled, inactive, comments):
    cursor = self.db.cursor()
    query = ('INSERT INTO Clinicians(Prefix, First, Last, Designation, Phone, Fax, Email, [Address 1], [Address 2], City, State, Zip, Comments) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)')
    cursor.execute(query, prefix, first, last, designation, phone, fax, email, addr1, addr2, city, state, zip, comments)

  @__usesCursor
  def addPrefixes(self, cursor, type, prefix, word):
    query = ('INSERT INTO Prefixes(Type, Prefix, Word) VALUES(?, ?, ?)')
    cursor.execute(query, type, prefix, word)

  @__usesCursor
  def genSampleID(self, cursor):
    tables = ['Cultures', 'CATs', 'Waterlines']
    yy = self.date.year-2000
    count = 0
    for table in tables:
      query = (f'SELECT COUNT(*) FROM {table} WHERE SampleID >= {yy}0000 AND SampleID < {yy+1}0000')
      cursor.execute(query)
      catch = cursor.fetchone()
      count += catch[0] if catch is not None else 0
    return (yy*10000)+count+1

  @__usesCursor
  def addPatientOrder(self, cursor, table, chartID, clinician, first, last, collected, received, type, comments, notes):
    sampleID = self.genSampleID()
    query = (f'INSERT INTO {table}(SampleID, ChartID, Clinician, First, Last, Collected, Received, Type, Comments, Notes) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)')
    cursor.execute(query, sampleID, chartID, clinician, first, last, self.fQtDate(collected), self.fQtDate(received), type, comments, notes)
    return sampleID

  @__usesCursor
  def addCATResult(self, cursor, sampleID, clinician, first, last, reported, type, volume, time, flow, pH, bc, sm, lb, comments, notes):
    query = ('UPDATE CATs SET [Clinician]=?, [First]=?, [Last]=?, [Tech]=?, [Reported]=?, [Type]=?, [Volume (ml)]=?, [Time (min)]=?, [Flow Rate (ml/min)]=?, [Initial (pH)]=?, [Buffering Capacity (pH)]=?, [Strep Mutans (CFU/ml)]=?, [Lactobacillus (CFU/ml)]=?, [Comments]=?, [Notes]=? WHERE [SampleID]=?')
    cursor.execute(query, clinician, first, last, self.tech[0], self.fQtDate(reported), type, volume, time, flow, pH, bc, sm, lb, comments, notes, sampleID)
    return True

  @__usesCursor
  def addCultureResult(self, cursor, sampleID, chartID, clinician, first, last, tech, reported, type, smear, aerobic, anaerobic, comments, notes):
    query = ('UPDATE Cultures SET [ChartID]=?, [Clinician]=?, [First]=?, [Last]=?, [Tech]=?, [Reported]=?, [Type]=?, [Direct Smear]=?, [Aerobic Results]=?, [Anaerobic Results]=?, [Comments]=?, [Notes]=? WHERE [SampleID]=?')
    cursor.execute(query, chartID, clinician, first, last, tech, self.fQtDate(reported), type, smear, aerobic, anaerobic, comments, notes, sampleID)
    return True
  
  @__usesCursor
  def addWaterlineOrder(self, cursor, clinician, shipped, comments, notes):
    sampleID = self.genSampleID()
    query = ('INSERT INTO Waterlines(SampleID, Clinician, Shipped, Comments, Notes) VALUES(?, ?, ?, ?, ?)')
    cursor.execute(query, sampleID, clinician, self.fQtDate(shipped), comments, notes)
    return sampleID

  @__usesCursor
  def addWaterlineReceiving(self, cursor, sampleID, operatoryID, clinician, collected, received, product, procedure, comments, notes):
    query = ('UPDATE Waterlines SET [OperatoryID]=?, [Clinician]=?, [Collected]=?, [Received]=?, [Product]=?, [Procedure]=?, [Comments]=?, [Notes]=? WHERE [SampleID]=?')
    cursor.execute(query, operatoryID, clinician, self.fQtDate(collected), self.fQtDate(received), product, procedure, comments, notes, sampleID)
    return True

  @__usesCursor
  def addWaterlineResult(self, cursor, sampleID, clinician, reported, count, cdcada, comments, notes):
    query = ('UPDATE Waterlines SET [Clinician]=?, [Reported]=?, [Bacterial Count]=?, [CDC/ADA]=?, [Comments]=?, [Notes]=? WHERE SampleID=?')
    cursor.execute(query, clinician, self.fQtDate(reported), count, cdcada, comments, notes, sampleID)
    return True

  @__usesCursor
  def findSample(self, cursor, table, sampleID, columns):
    query = f'SELECT {columns} FROM {table} WHERE SampleID=?'
    cursor.execute(query, sampleID)
    return cursor.fetchone()

  @__usesCursor
  def findClinician(self, cursor, entry):
    query = 'SELECT Prefix, First, Last, Designation, [Address 1], [City], [State], [Zip] FROM Clinicians WHERE Entry=?'
    cursor.execute(query, entry)
    return cursor.fetchone()

  @__usesCursor
  def findTech(self, cursor, entry, columns):
    query = f'SELECT {columns} FROM Techs WHERE Entry=?'
    cursor.execute(query, entry)
    return cursor.fetchone()

  @__usesCursor
  def findPrefix(self, cursor, prefix, columns):
    query = f'SELECT {columns} FROM Prefixes WHERE Prefix=?'
    cursor.execute(query, prefix)
    return cursor.fetchone()

  @__usesCursor
  def selectClinicians(self, cursor, columns):
    query = f'SELECT {columns} FROM Clinicians'
    cursor.execute(query)
    return cursor.fetchall()

  @__usesCursor
  def selectTechs(self, cursor, columns):
    query = f'SELECT {columns} FROM Techs'
    cursor.execute(query)
    return cursor.fetchall()

  @__usesCursor
  def selectPrefixes(self, cursor, type, columns):
    query = f'SELECT {columns} FROM Prefixes WHERE Type=?'
    cursor.execute(query, type)
    return cursor.fetchall()

  @__usesCursor
  def updateTech(self, cursor, entry, first, middle, last, username, password):
    query = f'UPDATE Techs SET [First]=?, [Middle]=?, [Last]=?, [Username]=?, [Password]=? WHERE Entry=?'
    cursor.execute(query, first, middle, last, username, self.encrypt(password).decode('utf-8'), entry)
    return True

  @__usesCursor
  def updatePrefixes(self, cursor, entry, type, prefix, word):
    query = f'UPDATE Prefixes SET [Type]=?, [Prefix]=?, [Word]=? WHERE Entry=?'
    cursor.execute(query, type, prefix, word, entry)
    return True

  @__usesCursor
  def techLogin(self, cursor, username, password):
    query = 'SELECT Entry, First, Middle, Last, Username, Password, Active FROM Techs WHERE username=?'
    cursor.execute(query, username)
    for tech in cursor.fetchall():
      if bcrypt.checkpw(password.encode('utf-8'), tech[5].encode('utf-8')) and tech[6] == 'Yes':
        self.date = date.today()
        self.tech = tech
        return True
      return False

  def encrypt(self, token):
    bsalt = bcrypt.gensalt()
    return bcrypt.hashpw(token.encode('utf-8'), bsalt)

  def fQtDate(self, qtDate):
    return date(qtDate.year(), qtDate.month(), qtDate.day())
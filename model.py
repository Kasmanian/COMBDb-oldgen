import bcrypt
from datetime import date
import pyodbc, json

class Model:
  def __init__(self):
    pass
  
  def connect(self):
    try:
      f = open('COMBDb\local.json')
      PATH = json.load(f)['DBQ']
      CONSTR = r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+PATH
      f.close()
      self.db = pyodbc.connect(CONSTR)
      #techs = self.selectTechs('*')
      #print(techs)
      return True
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False

  def close(self):
    self.db.close()

  def addTech(self, first, middle, last, username, password):
    try:
      cursor = self.db.cursor()
      query = (
        'INSERT INTO Techs(First, Middle, Last, Username, Password, Active) VALUES(?, ?, ?, ?, ?, ?)'
      )
      cursor.execute(query, first, middle, last, username, self.encrypt(password).decode('utf-8'), 'Yes')
      self.db.commit()
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False
    finally:
      cursor.close()

  def toggleTech(self, entry, active):
    try:
      cursor = self.db.cursor()
      query = (
        'UPDATE Techs '
        f'SET Active=? '
        f'WHERE Entry=?'
      )
      cursor.execute(query, active, entry)
      self.db.commit()
      return True
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False
    finally:
      cursor.close()

  def addGuest(self, pw, ls):
    pass

  def addClinician(self, prefix, first, last, designation, phone, fax, email, addr1, addr2, city, state, zip, enrolled, inactive, comments):
    try:
      cursor = self.db.cursor()
      query = (
        'INSERT INTO Clinicians(Prefix, First, Last, Designation, Phone, Fax, Email, '
        '[Address 1], [Address 2], City, State, Zip, Comments) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'
      )
      cursor.execute(query, prefix, first, last, designation, phone, fax, email, addr1, addr2, city, state, zip, comments)
      self.db.commit()
    except Exception as e:
      print(f'Error in connection: {e}')
      return False
    finally:
      cursor.close()

  def genSampleID(self):
    tables = ['Cultures', 'CATs', 'Waterlines']
    yy = self.date.year-2000
    cursor = self.db.cursor()
    count = 0
    for table in tables:
      query = (
        f'SELECT COUNT(*) FROM {table} WHERE SampleID >= {yy}0000 AND SampleID < {yy+1}0000'
      )
      cursor.execute(query)
      catch = cursor.fetchone()
      count += catch[0] if catch is not None else 0
    return (yy*10000)+count+1

  def addPatientOrder(self, table, chartID, clinician, first, last, collected, received, type, comments, notes):
    try:
      cursor = self.db.cursor()
      sampleID = self.genSampleID()
      query = (
        f'INSERT INTO {table}(SampleID, ChartID, Clinician, First, Last, Collected, Received, Type, Comments, Notes) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'
      )
      cursor.execute(query, sampleID, chartID, clinician, first, last, self.fQtDate(collected), self.fQtDate(received), type, comments, notes)
      self.db.commit()
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      sampleID = False
    finally:
      cursor.close()
      return sampleID

  def addCATResult(self, sampleID, clinician, first, last, reported, type, volume, time, flow, pH, bc, sm, lb, comments, notes):
    try:
      cursor = self.db.cursor()
      query = (
        'UPDATE CATs '
        f'SET [Clinician]=?, [First]=?, [Last]=?, [Tech]=?, [Reported]=?, [Type]=?, [Volume (ml)]=?, [Time (min)]=?, [Flow Rate (ml/min)]=?, [Initial (pH)]=?, '
        f'[Buffering Capacity (pH)]=?, [Strep Mutans (CFU/ml)]=?, [Lactobacillus (CFU/ml)]=?, [Comments]=?, [Notes]=? WHERE [SampleID]=?'
      )
      cursor.execute(query, clinician, first, last, self.tech[0], self.fQtDate(reported), type, volume, time, flow, pH, bc, sm, lb, comments, notes, sampleID)
      self.db.commit()
      return True
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False
    finally:
      cursor.close()

  def addCultureResult(self, sampleID, chartID, clinician, first, last, reported, aerobic, anaerobic, comments, notes):
    try:
      cursor = self.db.cursor()
      query = (
        'UPDATE Cultures SET [ChartID]=?, [Clinician]=?, [First]=?, [Last]=?, [Tech]=?, [Reported]=?, [Aerobic Results]=?, [Anaerobic Results]=?, [Comments]=?, [Notes]=? WHERE [SampleID]=?'
      )
      cursor.execute(query, chartID, clinician, first, last, self.tech[0], self.fQtDate(reported), aerobic, anaerobic, comments, notes, sampleID)
      self.db.commit()
      return True
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False
    finally:
      cursor.close()
  
  def addWaterlineOrder(self, clinician, shipped, comments, notes):
    try:
      cursor = self.db.cursor()
      sampleID = self.genSampleID()
      query = (
        'INSERT INTO Waterlines(SampleID, Clinician, Shipped, Comments, Notes) VALUES(?, ?, ?, ?, ?)'
      )
      cursor.execute(query, sampleID, clinician, self.fQtDate(shipped), comments, notes)
      self.db.commit()
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      sampleID = False
    finally:
      cursor.close()
      return sampleID

  def addWaterlineReceiving(self, sampleID, operatoryID, clinician, collected, received, product, procedure, comments, notes):
    try:
      ret = False
      cursor = self.db.cursor()
      query = (
        'UPDATE Waterlines SET [OperatoryID]=?, [Clinician]=?, [Collected]=?, [Received]=?, [Product]=?, [Procedure]=?, [Comments]=?, [Notes]=? WHERE [SampleID]=?'
      )
      cursor.execute(query, operatoryID, clinician, self.fQtDate(collected), self.fQtDate(received), product, procedure, comments, notes, sampleID)
      self.db.commit()
      ret = True
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      ret = False
    finally:
      cursor.close()
      return ret

  def addWaterlineResult(self, sampleID, clinician, reported, count, cdcada, comments, notes):
    try:
      cursor = self.db.cursor()
      query = (
        'UPDATE Waterlines SET [Clinician]=?, [Reported]=?, [Bacterial Count]=?, [CDC/ADA]=?, [Comments]=?, [Notes]=? WHERE SampleID=?'
      )
      cursor.execute(query, clinician, self.fQtDate(reported), count, cdcada, comments, notes, sampleID)
      self.db.commit()
      return True
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False
    finally:
      cursor.close()

  def findSample(self, table, sampleID, columns):
    try:
      query = f'SELECT {columns} FROM {table} WHERE SampleID=?'
      cursor = self.db.cursor()
      cursor.execute(query, sampleID)
      return cursor.fetchone()
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return None
    finally:
      cursor.close()

  def findClinician(self, entry):
    try:
      cursor = self.db.cursor()
      query = 'SELECT Prefix, First, Last, Designation, [Address 1], [City], [State], [Zip] FROM Clinicians WHERE Entry=?'
      cursor.execute(query, entry)
      return cursor.fetchone()
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return None
    finally:
      cursor.close()

  def findTech(self, entry, columns):
    try:
      cursor = self.db.cursor()
      query = f'SELECT {columns} FROM Techs WHERE Entry=?'
      cursor.execute(query, entry)
      return cursor.fetchone()
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return None
    finally:
      cursor.close()

  def selectClinicians(self, columns):
    try:
      cursor = self.db.cursor()
      query = f'SELECT {columns} FROM Clinicians'
      cursor.execute(query)
      return cursor.fetchall()
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return None
    finally:
      cursor.close()

  def selectTechs(self, columns):
    try:
      cursor = self.db.cursor()
      query = f'SELECT {columns} FROM Techs'
      cursor.execute(query)
      return cursor.fetchall()
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return None
    finally:
      cursor.close()

  def updateTech(self, entry, first, middle, last, username, password):
    try:
      cursor = self.db.cursor()
      query = f'UPDATE Techs SET [First]=?, [Middle]=?, [Last]=?, [Username]=?, [Password]=? WHERE Entry=?'
      cursor.execute(query, first, middle, last, username, self.encrypt(password).decode('utf-8'), entry)
      return True
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False
    finally:
      cursor.close()

  def techLogin(self, username, password):
    try:
      cursor = self.db.cursor()
      query = 'SELECT Entry, First, Middle, Last, Username, Password, Active FROM Techs WHERE username=?'
      cursor.execute(query, username)
      for tech in cursor.fetchall():
        if bcrypt.checkpw(password.encode('utf-8'), tech[5].encode('utf-8')) and tech[6] == 'Yes':
          self.date = date.today()
          self.tech = tech
          return True
        return False
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False
    finally:
      cursor.close()

  def encrypt(self, token):
    bsalt = bcrypt.gensalt()
    return bcrypt.hashpw(token.encode('utf-8'), bsalt)

  def fQtDate(self, qtDate):
    return date(qtDate.year(), qtDate.month(), qtDate.day())
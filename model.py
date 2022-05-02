import bcrypt, math
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
      return True
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False

  def close(self):
    self.db.close()

  def addTech(self, first, middle, last, username, password):
    # Insert new Admin into the database using a given username and password
    try:
      cursor = self.db.cursor()
      # cursor.execute(f'INSERT INTO Techs(First, Last, Username, Password) VALUES(?, ?, ?, ?, ?, ?))', first, last, middle, username, self.encrypt(password).decode('utf-8'))
      # query = (
      #   'INSERT INTO Techs(First, Middle, Last, Username, Password, Active)'
      #   f'VALUES({first}, {middle}, {last}, {username}, {self.encrypt(password).decode("utf-8")}, Yes)'
      # )
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
    # Insert new Guest into the database using:
    # a randomly generated key (pw), a timestamp (ts, hours since epoch to the nearest hundredth), and a lifespan (ls, hours)
    # self.jdrv['Guest'].insert_one({'password': self.encrypt(pw), 'timestmp': math.floor(time.time()/36)/100, 'lifespan': ls})
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

  def addPatientOrder(self, table, chartID, clinician, first, last, collected, received, comments):
    try:
      yy = self.date.year-2000
      cursor = self.db.cursor()
      query = (
        f'SELECT COUNT(*) FROM {table} WHERE SampleID >= {yy}0000 AND SampleID < {yy+1}0000'
      )
      cursor.execute(query)
      lastPatientOrder = cursor.fetchone()
      sampleID = (yy*10000)+lastPatientOrder[0]+1 if lastPatientOrder is not None else (yy*10000)+1
      query = (
        f'INSERT INTO {table}(SampleID, ChartID, Clinician, First, Last, Collected, Received, Comments) VALUES(?, ?, ?, ?, ?, ?, ?, ?)'
      )
      #  f'VALUES({sampleID}, {chartID}, {clinician}, {first}, {last}, {collected}, {received}, {comments})'
      cursor.execute(query, sampleID, chartID, clinician, first, last, self.fQtDate(collected), self.fQtDate(received), comments)
      self.db.commit()
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False
    finally:
      cursor.close()
      return sampleID

  def addCATResult(self, sampleID, chartID, clinician, first, last, reported, volume, time, flow, pH, bc, sm, lb, comments):
    try:
      cursor = self.db.cursor()
      # query = (
      #   'UPDATE CATs '
      #   f'SET ChartID={chartID}, Clinician={clinician}, First={first}, Last={last}, Tech={self.tech[0]}, Reported={reported}, ' 
      #   f'[Volume (ml)]={volume}, [Time (min)]={time}, [Flow Rate (ml/min)]={flow}, pH={pH}, '
      #   f'[Buffering Capacity (pH)]={bc}, [Strep Mutans (CFU/ml)]={sm}, [Lactobacillus (CFU/ml)]={lb}, Comments={comments} '
      #   f'WHERE SampleID={sampleID}'
      # )
      query = (
        'UPDATE CATs '
        f'SET ChartID=?, Clinician=?, First=?, Last=?, Tech=?, Reported=?, [Volume (ml)]=?, [Time (min)]=?, [Flow Rate (ml/min)]=?, pH=?, '
        f'[Buffering Capacity (pH)]=?, [Strep Mutans (CFU/ml)]=?, [Lactobacillus (CFU/ml)]=?, Comments=? WHERE SampleID=?'
      )
      cursor.execute(query, chartID, clinician, first, last, self.tech[0], reported, volume, time, flow, pH, bc, sm, lb, comments, sampleID)
      self.db.commit()
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False
    finally:
      cursor.close()

  def addCultureResult(self, sampleID, chartID, clinician, first, last, reported, results, comments):
    try:
      cursor = self.db.cursor()
      # query = (
      #   'UPDATE Cultures '
      #   f'SET ChartID={chartID}, Clinician={clinician}, First={first}, Last={last}, Tech={self.tech[0]}, Reported={reported}, Results={results}, Comments={comments} '
      #   f'WHERE SampleID={sampleID}'
      # )
      query = (
        'UPDATE Cultures SET ChartID=?, Clinician=?, First=?, Last=?, Tech=?, Reported=?, Results=?, Comments=? WHERE SampleID=?'
      )
      cursor.execute(query, chartID, clinician, first, last, self.tech[0], reported, results, comments, sampleID)
      self.db.commit()
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False
    finally:
      cursor.close()
  
  def addWaterlineOrder(self, clinician, shipped, comments):
    try:
      yy = int(self.date[-1:-2])
      cursor = self.db.cursor()
      query = (
        f'SELECT COUNT(*) FROM Waterlines WHERE SampleID >= {yy}0000 AND SampleID < {yy+1}0000'
      )
      cursor.execute(query)
      sampleID = (yy*10000)+cursor.fetchone()[0]+1
      # query = (
      #   'INSERT INTO Waterlines(SampleID, Clinician, Shipped, Comments)'
      #   f'VALUES({sampleID}, {clinician}, {shipped}, {comments})'
      # )
      query = (
        'INSERT INTO Waterlines(SampleID, Clinician, Shipped, Comments) VALUES(?, ?, ?, ?)'
      )
      cursor.execute(query, sampleID, clinician, shipped, comments)
      self.db.commit()
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False
    finally:
      cursor.close()

  def addWaterlineReceiving(self, sampleID, operatoryID, clinician, collected, received, product, procedure, comments):
    try:
      cursor = self.db.cursor()
      # query = (
      #   'UPDATE Waterlines'
      #   f'SET OperatoryID={operatoryID}, Clinician={clinician}, Collected={collected}, Received={received}, Product={product}, Procedure={procedure}, Comments={comments}'
      #   f'WHERE SampleID={sampleID}'
      # )
      query = (
        'UPDATE Waterlines SET OperatoryID=?, Clinician=?, Collected=?, Received=?, Product=?, Procedure=?, Comments=? WHERE SampleID=?'
      )
      cursor.execute(query, operatoryID, clinician, collected, received, product, procedure, comments, sampleID)
      self.db.commit()
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False
    finally:
      cursor.close()

  def addWaterlineResult(self, sampleID, clinician, reported, count, cdcada, comments):
    try:
      cursor = self.db.cursor()
      # query = (
      #   'UPDATE Waterlines'
      #   f'SET Clinician={clinician}, Reported={reported}, [Bacterial Count]={count}, [CDC/ADA]={cdcada}, Comments={comments}'
      #   f'WHERE SampleID={sampleID}'
      # )
      query = (
        'UPDATE Waterlines SET Clinician=?, Reported=?, [Bacterial Count]=?, [CDC/ADA]=?, Comments=? WHERE SampleID=?'
      )
      cursor.execute(query, clinician, reported, count, cdcada, comments, sampleID)
      self.db.commit()
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False
    finally:
      cursor.close()

  def findSample(self, sampleType, sampleID):
    query = ' WHERE SampleID=?'
    if sampleType == 'CAT':
      query = 'SELECT * FROM CATs' + query
    elif sampleType == 'Culture':
      query = 'SELECT * FROM Cultures' + query
    elif sampleType == 'DUWL':
      query = 'SELECT * FROM Waterlines' + query
    else:
      return None
    try:
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
      query = 'SELECT Prefix, First, Last FROM Clinicians WHERE Entry=?'
      cursor.execute(query, entry)
      return cursor.fetchone()
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return None
    finally:
      cursor.close()

  def findSample(self, table, sampleID):
    try:
      cursor = self.db.cursor()
      extraf = ' Shipped,' if table=='Waterlines' else ''
      query = f'SELECT chartID, Clinician, First, Last,{extraf} Collected, Received, Comments FROM {table} WHERE sampleID=?'
      cursor.execute(query, sampleID)
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

  def techLogin(self, username, password):
    # Pull from Techs table matching & validating user input
    try:
      cursor = self.db.cursor()
      query = 'SELECT * FROM Techs WHERE username=?'
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

  def guestLogin(self, pw):
    # Create Guest object modeled in schemas.py from the results of fetching a Guest using password (pw)
    # guest = Guest(self.jdrv['Guest'].find_one({'password': self.encrypt(pw)}))
    # If Guest's data is None, there is no such account; bcrypt will validate the password and check expiry if the account exists
    # return guest.data is not None and bcrypt.checkpw(pw, guest.data['password']) and math.floor(time.time()/36)/100-guest.data['timestamp']>guest.data['lifespan']
    pass

  def encrypt(self, token):
    # Salt and hash token (tk)
    bsalt = bcrypt.gensalt()
    return bcrypt.hashpw(token.encode('utf-8'), bsalt)

  def fQtDate(self, qtDate):
    # Convert QtDate to MS Access Data
    # return f'#{date.month()}/{date.day()}/{date.year()}#'
    return date(qtDate.year(), qtDate.month(), qtDate.day())
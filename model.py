from pymongo import MongoClient
from schemas import *
import bcrypt, math, time
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

  def addTech(self, first, last, middle, username, password):
    # Insert new Admin into the database using a given username and password
    try:
      cursor = self.db.cursor()
      # cursor.execute(f'INSERT INTO Techs(First, Last, Username, Password) VALUES(?, ?, ?, ?, ?, ?))', first, last, middle, username, self.encrypt(password).decode('utf-8'))
      query = (
        'INSERT INTO Techs(First, Middle, Last, Username, Password, Active)'
        f'VALUES({first}, {middle}, {last}, {username}, {self.encrypt(password).decode("utf-8")}, Yes)'
      )
      cursor.execute(query)
      self.db.commit()
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False
    finally:
      cursor.close()

  def addGuest(self, pw, ls):
    # Insert new Guest into the database using:
    # a randomly generated key (pw), a timestamp (ts, hours since epoch to the nearest hundredth), and a lifespan (ls, hours)
    self.jdrv['Guest'].insert_one({'password': self.encrypt(pw), 'timestmp': math.floor(time.time()/36)/100, 'lifespan': ls})

  def addPatientOrder(self, table, chartID, clinician, first, last, collected, received, comments):
    try:
      year = 22
      cursor = self.db.cursor()
      query = (
        f'SELECT COUNT(*) FROM {table} WHERE SampleID >= {year}0000 AND SampleID < {year+1}0000'
      )
      cursor.execute(query)
      sampleID = (year * 10000) + cursor.fetchone()[0]+1
      query = (
        f'INSERT INTO {table}(SampleID, ChartID, Clinician, First, Last, Collected, Received, Comments)'
        f'VALUES({sampleID}, {chartID}, {clinician}, {first}, {last}, {collected}, {received}, {comments})'
      )
      cursor.execute(query)
      self.db.commit()
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False
    finally:
      cursor.close()

  def addCATResult(self):
    pass

  def addCultureResult():
    pass
  
  def addWaterlineOrder(self, clinician, shipped, comments):
    try:
      year = 22
      cursor = self.db.cursor()
      query = (
        f'SELECT COUNT(*) FROM Waterlines WHERE SampleID >= {year}0000 AND SampleID < {year+1}0000'
      )
      cursor.execute(query)
      sampleID = (year * 10000) + cursor.fetchone()[0]+1
      query = (
        'INSERT INTO Waterlines(SampleID, Clinician, Shipped, Comments)'
        f'VALUES({sampleID}, {clinician}, {shipped}, {comments})'
      )
      cursor.execute(query)
      self.db.commit()
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False
    finally:
      cursor.close()

  def addWaterlineReceiving(self, sampleID, operatoryID, clinician, collected, received, product, procedure, comments):
    try:
      cursor = self.db.cursor()
      query = (
        'UPDATE Waterlines'
        f'SET OperatoryID={operatoryID}, Clinician={clinician}, Collected={collected}, Received={received}, Product={product}, Procedure={procedure}, Comments={comments}'
        f'WHERE SampleID={sampleID}'
      )
      cursor.execute(query)
      self.db.commit()
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False
    finally:
      cursor.close()

  def addWaterlineResult():
    pass

  def findSample(sampleType, sampleID):
    pass

  def techLogin(self, username, password):
    # Pull from Techs table matching & validating user input
    try:
      cursor = self.db.cursor()
      query = (
        f'SELECT * FROM Techs WHERE username = {username}'
      )
      cursor.execute(query)
      for tech in cursor.fetchall():
        if bcrypt.checkpw(password.encode('utf-8'), tech[4].encode('utf-8')):
          return True
        return False
    except (Exception, pyodbc.Error) as e:
      print(f'Error in connection: {e}')
      return False
    finally:
      cursor.close()

  def guestLogin(self, pw):
    # Create Guest object modeled in schemas.py from the results of fetching a Guest using password (pw)
    guest = Guest(self.jdrv['Guest'].find_one({'password': self.encrypt(pw)}))
    # If Guest's data is None, there is no such account; bcrypt will validate the password and check expiry if the account exists
    return guest.data is not None and bcrypt.checkpw(pw, guest.data['password']) and math.floor(time.time()/36)/100-guest.data['timestamp']>guest.data['lifespan']

  def encrypt(self, token):
    # Salt and hash token (tk)
    bsalt = bcrypt.gensalt()
    return bcrypt.hashpw(token.encode('utf-8'), bsalt)
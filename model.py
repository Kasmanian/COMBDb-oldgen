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
      f = open('local.json')
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

  # @__usesCursor
  # def genSampleID(self, cursor):
  #   tables = ['Cultures', 'CATs', 'Waterlines']
  #   yy = self.date.year-2000
  #   count = 0
  #   for table in tables:
  #     query = (f'SELECT COUNT(*) FROM {table} WHERE SampleID >= {yy}0000 AND SampleID < {yy+1}0000')
  #     cursor.execute(query)
  #     catch = cursor.fetchone()
  #     count += catch[0] if catch is not None else 0
  #   return (yy*10000)+count+1

  @__usesCursor
  def genSampleID(self, cursor):
    yy = self.date.year-2000
    query = (f'SELECT ID FROM SampleID')
    cursor.execute(query)
    fetchID = cursor.fetchone()[0]
    catchID = (yy*10000)+1 if yy-(fetchID//10000)>0 else fetchID
    query = (f'UPDATE SampleID SET ID=? WHERE ID=?')
    cursor.execute(query, catchID+1, fetchID)
    return catchID

  @__usesCursor
  def addPatientOrder(self, cursor, table, chartID, clinician, first, last, collected, received, type, tech, comments, notes):
    sampleID = self.genSampleID()
    query = (f'INSERT INTO {table}(SampleID, ChartID, Clinician, First, Last, Collected, Received, Type, Tech, Comments, Notes) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)')
    cursor.execute(query, sampleID, chartID, clinician, first, last, self.fQtDate(collected), self.fQtDate(received), type, tech, comments, notes)
    return sampleID

  @__usesCursor
  def addCATResult(self, cursor, sampleID, clinician, first, last, tech, reported, type, volume, time, flow, pH, bc, sm, lb, comments, notes, rejectionDate, rejectionReason):
    query = ('UPDATE CATs SET [Clinician]=?, [First]=?, [Last]=?, [Tech]=?, [Reported]=?, [Type]=?, [Volume (ml)]=?, [Time (min)]=?, [Flow Rate (ml/min)]=?, [Initial (pH)]=?, [Buffering Capacity (pH)]=?, [Strep Mutans (CFU/ml)]=?, [Lactobacillus (CFU/ml)]=?, [Comments]=?, [Notes]=?, [Rejection Date]=?, [Rejection Reason]=? WHERE [SampleID]=?')
    cursor.execute(query, clinician, first, last, tech, self.fQtDate(reported), type, volume, time, flow, pH, bc, sm, lb, comments, notes, self.fQtDate(rejectionDate) if rejectionDate != None else None, rejectionReason, sampleID)
    return True

  @__usesCursor
  def addCultureResult(self, cursor, sampleID, chartID, clinician, first, last, tech, reported, type, smear, aerobic, anaerobic, comments, notes, rejectionDate, rejectionReason):
    query = ('UPDATE Cultures SET [ChartID]=?, [Clinician]=?, [First]=?, [Last]=?, [Tech]=?, [Reported]=?, [Type]=?, [Direct Smear]=?, [Aerobic Results]=?, [Anaerobic Results]=?, [Comments]=?, [Notes]=?, [Rejection Date]=?, [Rejection Reason]=? WHERE [SampleID]=?')
    cursor.execute(query, chartID, clinician, first, last, tech, self.fQtDate(reported), type, smear, aerobic, anaerobic, comments, notes, self.fQtDate(rejectionDate) if rejectionDate != None else None, rejectionReason, sampleID)
    return True
  
  @__usesCursor
  def addWaterlineOrder(self, cursor, clinician, shipped, comments, notes, tech):
    sampleID = self.genSampleID()
    query = ('INSERT INTO Waterlines(SampleID, Clinician, Shipped, Comments, Notes, Tech) VALUES(?, ?, ?, ?, ?, ?)')
    cursor.execute(query, sampleID, clinician, self.fQtDate(shipped), comments, notes, tech)
    return sampleID

  @__usesCursor
  def addWaterlineReceiving(self, cursor, sampleID, operatoryID, clinician, collected, received, product, procedure, comments, notes, rejectionDate, rejectionReason, tech):
    query = ('UPDATE Waterlines SET [OperatoryID]=?, [Clinician]=?, [Collected]=?, [Received]=?, [Product]=?, [Procedure]=?, [Comments]=?, [Notes]=?, [Rejection Date]=?, [Rejection Reason]=?, [Tech]=? WHERE [SampleID]=?')
    cursor.execute(query, operatoryID, clinician, self.fQtDate(collected), self.fQtDate(received), product, procedure, comments, notes, self.fQtDate(rejectionDate) if rejectionDate != None else None, rejectionReason, tech, sampleID)
    return True

  @__usesCursor
  def addWaterlineResult(self, cursor, sampleID, clinician, reported, count, cdcada, comments, notes, rejectionDate, rejectionReason, tech):
    query = ('UPDATE Waterlines SET [Clinician]=?, [Reported]=?, [Bacterial Count]=?, [CDC/ADA]=?, [Comments]=?, [Notes]=?, [Rejection Date]=?, [Rejection Reason]=?, [Tech]=? WHERE SampleID=?')
    cursor.execute(query, clinician, self.fQtDate(reported), count, cdcada, comments, notes, self.fQtDate(rejectionDate) if rejectionDate != None else None, rejectionReason, tech, sampleID)
    return True

  @__usesCursor
  def findRejections(self, cursor, table, columns):
    query = f'SELECT {columns} FROM {table} WHERE [Rejection Date] IS NOT NULL AND [Rejection Reason] IS NOT NULL'
    cursor.execute(query)
    return cursor.fetchall()

  @__usesCursor
  def findSamplesQA(self, cursor, table, columns, fromDate, toDate):
    query = f'SELECT {columns} FROM {table} WHERE [Received] BETWEEN #{fromDate}# AND #{toDate}#'
    cursor.execute(query)
    return cursor.fetchall()

  @__usesCursor
  def findSample(self, cursor, table, sampleID, columns):
    query = f'SELECT {columns} FROM {table} WHERE SampleID=?'
    cursor.execute(query, sampleID)
    return cursor.fetchone()

  @__usesCursor
  def findSamples(self, cursor, table, inputs, columns):
    dynamicString = "WHERE "
    counter = 0
    for key, value in inputs.items():
      if counter > 0:
        dynamicString += " AND " if value is not None else ""
      if key != "Clinician" and key != "SampleID":
        dynamicString += f"{key}='{value}'" if value is not None else ""
      else:
        dynamicString += f"{key}={value}" if value is not None else ""
      counter += 1 if value is not None else 0
    query = f'SELECT {columns} FROM {table} {dynamicString}'
    cursor.execute(query)
    return cursor.fetchall()

  @__usesCursor
  def findSampleNumbers(self, cursor, table, columns):
    query = f'SELECT {columns} FROM {table}'
    cursor.execute(query)
    return cursor.fetchall()

  @__usesCursor
  def findClinician(self, cursor, entry):
    query = 'SELECT [Prefix], [First], [Last], [Designation], [Address 1], [City], [State], [Zip] FROM Clinicians WHERE Entry=?'
    cursor.execute(query, entry)
    return cursor.fetchone()

  @__usesCursor
  def findClinicianFull(self, cursor, entry):
    query = 'SELECT [Prefix], [First], [Last], [Phone], [Fax], [Designation], [Address 1], [Address 2], [City], [State], [Zip], [Email], [Enrolled], [Comments] FROM Clinicians WHERE Entry=?'
    cursor.execute(query, entry)
    return cursor.fetchone() 

  @__usesCursor
  def findTech(self, cursor, entry, columns):
    query = f'SELECT {columns} FROM Techs WHERE Entry=?'
    cursor.execute(query, entry)
    return cursor.fetchone()

  @__usesCursor
  def findTechUsername(self, cursor, username):
    query = f'SELECT Entry FROM Techs WHERE Username=?'
    cursor.execute(query, username)
    return cursor.fetchone()

  @__usesCursor
  def currentTech(self, cursor, username, columns):
    query = f'SELECT {columns} FROM Techs WHERE Username=?'
    cursor.execute(query, username)
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
  def updateCultureOrder(self, cursor, table, sampleID, chartID, clinician, first, last, collected, received, type, tech, comments, notes, rejectionDate, rejectionReason):
    query = f'UPDATE {table} SET [ChartID]=?, [Clinician]=?, [First]=?, [Last]=?, [Collected]=?, [Received]=?, [Type]=?, [Tech]=?, [Comments]=?, [Notes]=?, [Rejection Date]=?, [Rejection Reason]=? WHERE [SampleID]=?'
    cursor.execute(query, chartID, clinician, first, last, self.fQtDate(collected), self.fQtDate(received), type, tech, comments, notes, self.fQtDate(rejectionDate) if rejectionDate != None else None, rejectionReason, sampleID)
    return True

  @__usesCursor
  def updateWaterlineOrder(self, cursor, sampleID, clinician, shipped, comments, notes, rejectionDate, rejectionReason, tech):
    query = f'UPDATE Waterlines SET [Clinician]=?, [Shipped]=?, [Comments]=?, [Notes]=?, [Rejection Date]=?, [Rejection Reason]=?, [Tech]=? WHERE [SampleID]=?'
    cursor.execute(query, clinician, self.fQtDate(shipped), comments, notes, self.fQtDate(rejectionDate) if rejectionDate != None else None, rejectionReason, tech, sampleID)
    return True

  @__usesCursor
  def updateClinician(self, cursor, entry, prefix, first, last, designation, phone, fax, email, addr1, addr2, city, state, zip, inactive, comments):
    cursor = self.db.cursor()
    query = f'UPDATE Clinicians SET [Prefix]=?, [First]=?, [Last]=?, [Designation]=?, [Phone]=?, [Fax]=?, [Email]=?, [Address 1]=?, [Address 2]=?, [City]=?, [State]=?, [Zip]=?, [Comments]=? WHERE Entry=?'
    cursor.execute(query, prefix, first, last, designation, phone, fax, email, addr1, addr2, city, state, zip, comments, entry)
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
  
  @__usesCursor
  def auditor(self, cursor, tech, action, type, form, timestamp):
    query = ('INSERT INTO AuditLog([Tech], [Action], [Type], [Form], [Timestamp]) VALUES (?, ?, ?, ?, ?)')
    cursor.execute(query, tech, action, type, form, timestamp)
    return True
    

  def encrypt(self, token):
    bsalt = bcrypt.gensalt()
    return bcrypt.hashpw(token.encode('utf-8'), bsalt)

  def fQtDate(self, qtDate):
    return date(qtDate.year(), qtDate.month(), qtDate.day())
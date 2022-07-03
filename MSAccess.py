import pyodbc, json, bcrypt

class MSAccess:
  def connect(self):
    try:
      with open('COMBDb\local.json') as f:
        self.db = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+json.load(f)['DBQ'])
      return True
    except (Exception, pyodbc.Error) as e:
      self.Error = e
      return False

  def close(self):
    self.db.close()

  def insert(self, table: str, fields: tuple, *args: any):
    try:
      cursor = self.db.cursor()
      params = ','.join(fields)
      values = '?'
      for _ in range(1, len(fields)):
        values += ',?'
      query = f'INSERT INTO {table}({params}) VALUES({values})'
      cursor.execute(query, *args)
    except (Exception, pyodbc.Error) as e:
      self.Error = e
      return False
    finally:
      if cursor: cursor.close()

  def update(self, table: str, fields: tuple, *args: any):
    try:
      cursor = self.db.cursor()
      params = '=? '.join(fields[1:])
      query = f'UPDATE {table} SET {params}WHERE {fields[0]}=?'
      cursor.execute(query, *args)
    except (Exception, pyodbc.Error) as e:
      self.Error = e
      return False
    finally:
      if cursor: cursor.close()

  def select(self, table: str, fields: tuple):
    try:
      cursor = self.db.cursor()
      params = ', '.join(fields[1:])
      query = f'SELECT {params} FROM {table} WHERE {fields[0]}=?'
      cursor.execute(query)
    except (Exception, pyodbc.Error) as e:
      self.Error = e
      return False
    finally:
      if cursor: cursor.close()
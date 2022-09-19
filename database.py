import pyodbc, json

class Database:
  def connect(self):
    try:
      with open('local.json') as f:
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

  def update(self, table: str, fields: tuple, reqs: str, *args: any):
    try:
      cursor = self.db.cursor()
      params = '=? '.join(fields)
      query = f'UPDATE {table} SET {params}WHERE {reqs}'
      cursor.execute(query, *args)
    except (Exception, pyodbc.Error) as e:
      self.Error = e
      return False
    finally:
      if cursor: cursor.close()

  def select(self, table: str, fields: tuple, reqs: str, count: int):
    try:
      cursor = self.db.cursor()
      params = ', '.join(fields)
      reqs = f' WHERE {reqs}' if reqs is not None else None
      query = f'SELECT {params} FROM {table}{reqs}'
      cursor.execute(query)
      if count == 1:
        return cursor.fetchone()
      elif count > 1 and count < float('inf'):
        return cursor.fetchmany(count)
      else:
        return cursor.fetchall()
    except (Exception, pyodbc.Error) as e:
      self.Error = e
      return False
    finally:
      if cursor: cursor.close()

  def sample(self, year: int):
    tables = ['Cultures', 'CATs', 'Waterlines']
    cursor = self.db.cursor()
    count = 0
    for table in tables:
      query = (
        f'SELECT COUNT(*) FROM {table} WHERE SampleID >= {year}0000 AND SampleID < {year+1}0000'
      )
      cursor.execute(query)
      catch = cursor.fetchone()
      count += catch[0] if catch is not None else 0
    return (year*10000)+count+1
    
  def sample(self):
    cursor = self.db.cursor()
    query = f'SELECT [SampleID] FROM SampleID WHERE Entry=1'
    cursor.execute(query)
    sampleID = cursor.fetchone()
    query = f'UPDATE SampleID SET [SampleID]={sampleID+1} WHERE Entry=1'
    return sampleID
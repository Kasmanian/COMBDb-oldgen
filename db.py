import pyodbc, json

class Database:
  def __query(func):
    def query(self, *args, **kwargs):
        try:
          cursor = self.db.cursor()
          result = func(self, cursor, *args, **kwargs)
          self.db.commit()
          return result
        except (Exception, pyodbc.Error) as e:
          self.Error = e
          print(f'Error in connection: {e}')
        finally:
          if cursor: cursor.close()
    return query

  def connect(self):
    try:
      with open('COMBDb\local.json') as f:
        self.db = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+json.load(f)['DBQ'])
      return True
    except (Exception, pyodbc.Error) as e:
      self.Error = e
      return False

  def close(self):
    if self.db: self.db.close()

  @__query
  def insert(self, cursor: pyodbc.Cursor, table: str, fields: tuple, *args: any):
    params = ','.join(fields)
    values = '?'+(len(fields)-1)*',?'
    query = f'INSERT INTO {table}({params}) VALUES({values})'
    cursor.execute(query, *args)
    return True

  @__query
  def update(self, cursor: pyodbc.Cursor, table: str, fields: tuple, reqs: str, *args: any):
    params = '=? '.join(fields)
    query = f'UPDATE {table} SET {params}WHERE {reqs}'
    cursor.execute(query, *args)
    return True

  @__query
  def select(self, cursor: pyodbc.Cursor, table: str, fields: tuple, reqs: str, count: int):
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
    
  def sample(self, year):
    yearlyID = (year-2000)*10000
    sampleID = self.select('SampleID', ('[SampleID]'), 'Entry=1', 1)
    sampleID = yearlyID if yearlyID-sampleID>0 else sampleID
    self.update('SampleID', ('[SampleID]'), 'Entry=1', sampleID+1)
    return sampleID
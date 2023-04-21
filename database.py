import pyodbc, json

class Database:
  def __cursor(func):
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
      with open('local.json') as f:
        self.db = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+json.load(f)['DBQ'])
      return True
    except (Exception, pyodbc.Error) as e:
      self.Error = e
      return False

  def close(self):
    self.db.close()

  @__cursor
  def insert(self, cursor, table: str, fields: tuple, *args: any):
    params = ','.join(fields)
    values = '?'
    for _ in range(1, len(fields)):
      values += ',?'
    query = f'INSERT INTO {table}({params}) VALUES({values})'
    cursor.execute(query, *args)

  @__cursor
  def update(self, cursor, table: str, fields: tuple, reqs: str, *args: any):
    params = '=? '.join(fields)
    query = f'UPDATE {table} SET {params}WHERE {reqs}'
    cursor.execute(query, *args)

  @__cursor
  def select(self, cursor, table: str, fields: tuple, reqs: str, count: int):
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

  @__cursor 
  def sample(self, cursor):
    yy = self.date.year-2000
    query = (f'SELECT [ID] FROM [SampleID]')
    cursor.execute(query)
    fetchID = cursor.fetchone()[0]
    catchID = (yy*10000)+1 if yy-(fetchID//10000)>0 else fetchID
    query = (f'UPDATE [SampleID] SET [ID]=? WHERE [ID]=?')
    cursor.execute(query, catchID+1, fetchID)
    return catchID
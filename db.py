import pyodbc

class Database:


  def __query(func):
    '''#Decorator for grabbing/closing a cursor and catching database errors.'''
    def query(self, *args, **kwargs):
        try:
          cursor = self.db.cursor()
          result = func(self, cursor, *args, **kwargs)
          self.db.commit()
          return result
        except (Exception, pyodbc.Error) as e:
          self.error = e
        finally:
          if cursor: cursor.close()
    return query


  def connect(self, accdbPath: str):
    '''#Uses given path to accdb file and establishes a connection.

    Args:
      accdbPath: the full path to the accdb file
    '''
    try:
      self.db = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+accdbPath)
      return True
    except (Exception, pyodbc.Error) as e:
      self.error = e
      return False


  def close(self):
    '''#Closes the connection to the database if one exists.'''
    if self.db: self.db.close()


  @__query
  def insert(self, cursor: pyodbc.Cursor, table: str, fields: tuple, *args: any):
    '''#Stylized INSERT for COMBDb's database file.
    
    Args:
      cursor: the cursor given by the query wrapper (IGNORE LIKE "self")
      table: the name of the specific table to query
      fields: a tuple of strings listing the table's chosen fields
      *args: values to be inserted into the table's chosen fields
    '''
    params = ','.join(fields)
    values = '?'+(len(fields)-1)*',?'
    query = f'INSERT INTO {table}({params}) VALUES({values})'
    cursor.execute(query, *args)
    return True


  @__query
  def update(self, cursor: pyodbc.Cursor, table: str, fields: tuple, reqs: str, *args: any):
    '''#Stylized UPDATE for COMBDb's database file.
    
    Args:
      cursor: the cursor given by the query wrapper (IGNORE LIKE "self")
      table: the name of the specific table to query
      fields: a tuple of strings listing the table's chosen fields
      reqs: string specifying how to match the row(s) to be updated
      *args: values to be inserted into the table's chosen fields
    '''
    params = '=? '.join(fields)
    query = f'UPDATE {table} SET {params}WHERE {reqs}'
    cursor.execute(query, *args)
    return True


  @__query
  def select(self, cursor: pyodbc.Cursor, table: str, fields: tuple, reqs: str, count: int):
    '''#Stylized SELECT for COMBDb's database file.
    
    Args:
      cursor: the cursor given by the query wrapper (IGNORE LIKE "self")
      table: the name of the specific table to query
      fields: a tuple of strings listing the table's chosen fields
      reqs: string specifying how to match the row(s) to be selected
      count: the number of rows to select
    '''
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


  def generateSampleID(self, yy: int):
    '''#Generates a new non-colliding sample ID sourced from an index in the database.
    
    Args:
      yy: the current year (in YY format)
    '''
    fetchID = int(self.select('[IDKeys]', ('[ID]'), '[Type]="Sample"', 1))
    catchID = (yy*10000)+1 if yy-(fetchID//10000)>0 else fetchID
    self.update('[IDkeys]', ('[ID]'), '[Type]="Sample"', str(catchID+1))
    return str(catchID)


  def generateHexID(self, type: str):
    '''#Generates a new non-colliding hex ID sourced from an index in the database.
    
    Args:
      type: the name of the field using the hex ID
    '''
    fetchID = int(self.select('[IDKeys]', ('[ID]'), f'[Type]="{type}"', 1), 16)
    hexSize = 3 if type == '[Antibiotic]' or type == '[Bacteria]' else 1
    catchID = hex(fetchID+1)[2:].zfill(hexSize)
    self.update('[IDkeys]', ('[ID]'), f'[Type]="{type}"', catchID)
    return hex(fetchID)[2:].zfill(hexSize)
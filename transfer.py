import pyodbc, bcrypt, re

def addTech(first, middle, last, username, password):
    CONSTR = r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\simmsk\Desktop\COMBDb_Test.accdb'
    db = pyodbc.connect(CONSTR)
    cursor = db.cursor()
    cursor.execute('INSERT INTO Techs(First, Middle, Last, Username, Password, Active) VALUES(?, ?, ?, ?, ?, ?)', first, middle, last, username, encrypt(password).decode('utf-8'), 'Yes')
    # self.jdrv['Admin'].insert_one({'username': un, 'password': self.encrypt(pw)})
    db.commit()

def convertPractices():
    CONSTR1 = r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\simmsk\Desktop\COPY.CDSNew.mdb'
    db1 = pyodbc.connect(CONSTR1)
    cursor1 = db1.cursor()
    cursor1.execute('SELECT * FROM Dentist')
    CONSTR2 = r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\simmsk\Desktop\COMBDb_Test.accdb'
    db2 = pyodbc.connect(CONSTR2)
    cursor2 = db2.cursor()
    practices = cursor1.fetchall()
    for row in practices:
        title = row[3].lower().replace('.', '').capitalize() if row[3] is not None else None
        comment_fax = None
        if row[17] is not None:
            row[17] = re.sub('[^0-9a-zA-Z]+', '', row[17]).lower()
            if 'fax' in row[17]:
                comment_fax = re.sub('[^0-9]+', '', row[17])
        fax = row[13] if row[13] is not None else comment_fax
        designation = None
        if row[3] is None:
            if True:
                pass
        cursor2.execute('INSERT INTO Clinicians(Prefix, First, Last, )')
    cursor1.close()
    cursor2.close()
    db1.close()
    db2.close()

def convertData():
    CONSTR1 = r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\simmsk\Desktop\COPY.CDSNew.mdb'
    db1 = pyodbc.connect(CONSTR1)
    cursor1 = db1.cursor()
    cursor1.execute('SELECT * FROM [Data Table]')
    CONSTR2 = r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\simmsk\Desktop\COMBDb_Test.accdb'
    db2 = pyodbc.connect(CONSTR2)
    cursor2 = db2.cursor()
    for row in cursor1.fetchall():
        if row[0] < 20000:
            continue
        if row[4].lower() == 'duwl':
            # add to Waterlines table
            cursor2.execute('INSERT INTO Waterlines()')
        elif row[4].lower() == 'cat':
            # add to CATs table
            continue
        else:
            # add to Cultures table
            continue
        print(row)

def countRows(year):
    try:
        CONSTR = r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\simmsk\Desktop\COPY.CDSNew.mdb'
        db = pyodbc.connect(CONSTR)
        cursor = db.cursor()
        cursor.execute(f'SELECT COUNT(*) FROM [Data Table] WHERE SampleNum >= {year}0000 AND SampleNum < {year+1}0000')
        count = cursor.fetchone()[0]
        print(count)
    except (Exception, pyodbc.Error) as e:
        print(f'Error in connection: {e}')
    finally:
        cursor.close()
        db.close()

def encrypt(token):
    bsalt = bcrypt.gensalt()
    return bcrypt.hashpw(token.encode('utf-8'), bsalt)

if __name__=="__main__":
    # addTech('Eric', 'V.', 'Simmons', 'eric_simmons', 'Password')
    # addTech('Shannon', 'M.', 'Reisdorf', 'reisdorf', 'Password')
    countRows(20)
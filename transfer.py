import pyodbc, bcrypt, re, datetime

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
        query = (
            'INSERT INTO Clinicians(Prefix, First, Last, Phone, Fax, Email, [Address 1], [Address 2], City, State, Zip, Enrolled, Inactive, Comments, OldID) '
            f'VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'
        )
        # enrolled = datetime.datetime(int(row[15][6:]), int(row[15][0:2]), int(row[15][3:5]), 0, 0) if row[15] != '' else None
        # inactive = datetime.datetime(int(row[16][6:]), int(row[16][0:2]), int(row[16][3:5]), 0, 0) if row[16] != '' else None
        enrolled = None
        if row[15] != None and row[15] != '':
            d1 = row[15].replace('-', '/').replace('*', '').split('/')
            if len(d1) == 3:
                if d1[2] != '':
                    d1[2] = 2000+int(d1[2]) if int(d1[2])<1900 else int(d1[2])
                    enrolled = datetime.datetime(d1[2], int(d1[0]), int(d1[1]))
        inactive = None
        if row[16] != None and row[16] != '':
            d1 = row[16].replace('-', '/').replace('*', '').split('/')
            if len(d1) == 3:
                if d1[2] != '':
                    d1[2] = 2000+int(d1[2]) if int(d1[2])<1900 else int(d1[2])
                    enrolled = datetime.datetime(d1[2], int(d1[0]), int(d1[1]))
        cursor2.execute(query, row[3], row[2], row[1], row[12], row[13], row[14], row[7], row[8], row[9], row[10], row[11], enrolled, inactive, row[17], row[0])
        db2.commit()
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
        elif row[4] == None:
            continue
        elif row[4] == '':
            continue
        elif row[4].lower() == 'duwl':
            # add to Waterlines table
            clinician = -1
            if row[1] != None and row[1] != 0 and row[1] != '':
                cursor2.execute(f'SELECT * FROM Clinicians WHERE OldID={row[1]}')
                clinician = cursor2.fetchone()[0]
            query = (
                'INSERT INTO Waterlines(SampleID, OperatoryID, Clinician, Tech, Shipped, Collected, Received, )'
            )
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

def selectRow(ID):
    try:
        CONSTR = r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\simmsk\Desktop\COPY.CDSNew.mdb'
        db = pyodbc.connect(CONSTR)
        cursor = db.cursor()
        cursor.execute(f'SELECT * FROM [Data Table] WHERE SampleNum={ID}')
        row = cursor.fetchone()
        print(row)
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
    convertData()
from pymongo import MongoClient
from schemas import *
import bcrypt, math, time

class Model:
  def __init__(self):
    self.jdrv = MongoClient('mongodb://localhost:27017/')['OMCDb']

  def addAdmin(self, un, pw):
    # Insert new Admin into the database using:
    # an given username (un) and password (pw)
    self.jdrv['Admin'].insert_one({'username': un, 'password': self.encrypt(pw)})

  def addGuest(self, pw, ls):
    # Insert new Guest into the database using:
    # a randomly generated key (pw), a timestamp (ts, hours since epoch to the nearest hundredth), and a lifespan (ls, hours)
    self.jdrv['Guest'].insert_one({'password': self.encrypt(pw), 'timestmp': math.floor(time.time()/36)/100, 'lifespan': ls})


  def adminLogin(self, un, pw):
    # Create Admin object modeled in schemas.py from the results of fetching an Admin using username (un)
    admin = Admin(self.jdrv['Admin'].find_one({'username': un}))
    # If Admin's data is None, there is no such account; bcrypt will validate the password if the account exists
    return admin.data is not None and bcrypt.checkpw(pw, admin.data['password'])

  def guestLogin(self, pw):
    # Create Guest object modeled in schemas.py from the results of fetching a Guest using password (pw)
    guest = Guest(self.jdrv['Guest'].find_one({'password': self.encrypt(pw)}))
    # If Guest's data is None, there is no such account; bcrypt will validate the password and check expiry if the account exists
    return guest.data is not None and bcrypt.checkpw(pw, guest.data['password']) and math.floor(time.time()/36)/100-guest.data['timestamp']>guest.data['lifespan']

  def encrypt(self, tk):
    # Salt and hash token (tk)
    bsalt = bcrypt.gensalt()
    return bcrypt.hashpw(tk, bsalt)
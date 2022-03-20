from pymongo import MongoClient
from schemas import *
import bcrypt
import datetime

class Model:
  def __init__(self, view):
    self.view = view
    self.jdrv = MongoClient('mongodb://localhost:27017/')['OMCDb']

  def addAdmin(self, un, pw):
    bsalt = bcrypt.gensalt()
    hshpw = bcrypt.hashpw(pw, bsalt)
    self.jdrv['Admin'].insert_one({'username': un, 'password': hshpw})

  def addGuest(self, pw, ls):
    bsalt = bcrypt.gensalt()
    hshpw = bcrypt.hashpw(pw, bsalt)
    ts = str(datetime.datetime.now())
    ts = int(datetime.datetime(int(ts[0:4]), int(ts[5:7]), int(ts[8:10])).timestamp()/3600)
    self.jdrv['Guest'].insert_one({'password': hshpw, 'timestmp': ts, 'lifespan': ls})


  def adminLogin(self, un, pw):
    bsalt = bcrypt.gensalt()
    hshpw = bcrypt.hashpw(pw, bsalt)
    admin = Admin(self.jdrv['Admin'].find_one({'username': un}))
    #print(admin)
    if (admin.data['password']!=hshpw):
      print('Not logged in')
      return False
    print('Success!')
    return True

  def guestLogin(self, pw):
    bsalt = bcrypt.gensalt()
    hshpw = bcrypt.hashpw(pw, bsalt)
    guest = Guest(self.jdrv['Guest'].find_one({'password': hshpw}))
    if (guest.data!=None):
      if (int(datetime.datetime.now().timestamp()/3600)-int(guest.data.timestmp)>int(guest.data.lifespan)):
        return False
    else:
      return False
    return True

  def encrypt():
    return True

  def decrypt():
    return True
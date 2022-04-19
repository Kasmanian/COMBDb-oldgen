import sys
from pymongo import MongoClient
from model import Model
from view import View
from controller import Controller
from schemas import *
import webbrowser

def main():
    model = Model()
    model.connect()
    # model.addAdmin('John', 'Doe', 'Admin', 'Password')
    # model.adminLogin('rmaik', 'password1')
    view = View(model)
    # controller = Controller(model)

if __name__=="__main__":
    main()
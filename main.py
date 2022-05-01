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
    # try:
    #     model.addTech(None, None, 'Doe', 'Admin', 'Password')
    # except Exception as e:
    #     print(e)
    # model.adminLogin('rmaik', 'password1')
    view = View(model)
    # controller = Controller(model)

if __name__=="__main__":
    main()
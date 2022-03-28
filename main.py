import sys
from pymongo import MongoClient
from model import Model
from view import View
from controller import Controller
from schemas import *
import webbrowser

def main():
    # myclient = MongoClient("mongodb://localhost:27017/")
    # mydb = myclient["OMCDb"]
    # print(myclient.list_database_names())
    # mycol = mydb["Clinician"]
    # mydict = Clinician('Jane', 'Doe', 'Address 1', 'Address 2', 'NC', 'Chapel Hill', 28390, '9103334444', '9102225555', 'email@url.com')
    # x = mycol.insert_one(mydict.data)
    model = Model()
    # 
    # model.addAdmin('admin2', 'password')
    view = View(model)
    controller = Controller(model)
    #view.init()

if __name__=="__main__":
    main()
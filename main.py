import sys
from pymongo import MongoClient
from model import Model
from view import View
from controller import Controller
from schemas import *

def main():
#    myclient = MongoClient("mongodb://localhost:27017/")
#    mydb = myclient["OMCDb"]
#    print(myclient.list_database_names())
#    mycol = mydb["Clinician"]
#    mydict = Clinician('John', 'Doe', 'Address 1', 'Address 2', 'NC', 'Chapel Hill', 28390, '9103334444', '9102225555', 'email@url.com')
#    x = mycol.insert_one(mydict.data)
    view = View()
    model = Model(view)
    controller = Controller(model)
    #view.init()

if __name__=="__main__":
    main()
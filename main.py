import sys
from pymongo import MongoClient
from model import Model
from view import View
from controller import Controller

def main():
    view = View(sys)
    model = Model(view) 
    controller = Controller(model)
    view.display()

if __name__=="__main__":
    main()
import PyPDF2
from PyPDF2.pdf import PdfFileReader
from docx import Document
from docx.shared import Cm, Pt
import os
import datetime
from datetime import date, datetime as DateTime

class document:
    def __init__(self,doc_type,name,date_string,path_string):
        if doc_type == 1 or doc_type == "1":
            self.doc_type = "Pleadings"
        elif doc_type == 2:
            self.doc_type = "ET_Correspondence"
        elif doc_type == 3:
            self.doc_type = "Documents_Correspondence"
        elif doc_type == 4:
            self.doc_type = "Payslips"
        else:
            print("Error assigning doc_type!")
            print("."+str(doc_type)+".")
            print(str(type(doc_type)))
            self.doc_type = "Null"
        self.name = name
        self.date = datetime.date.fromisoformat(date_string)
        self.path_string = path_string
        print(path_string)

def returnDate(doc_arg):
    return doc_arg.date

def main():
    print("Initialising PDF handling...")

    dir = input("Please enter the directory to bundle from: ")

    print(dir)

    docs = os.listdir(dir)
    doc_list = []
    for doc in docs:
        print(doc)
        doc_proc = doc.strip(".pdf").split(";")
        doc_class = document(int(doc_proc[0]), doc_proc[1], doc_proc[2], (dir + "/" + str(doc)))
        doc_list.append(doc_class)
    
    print(str(doc_list))

    key_1_list = [doc for doc in doc_list if doc.doc_type == "Pleadings"]
    key_2_list = [doc for doc in doc_list if doc.doc_type == "ET_Correspondence"]
    key_3_list = [doc for doc in doc_list if doc.doc_type == "Documents_Correspondence"]
    key_4_list = [doc for doc in doc_list if doc.doc_type == "Payslips"]

    key_1_list.sort(key= lambda x: returnDate(x))
    key_2_list.sort(key= lambda x: returnDate(x))
    key_3_list.sort(key= lambda x: returnDate(x))
    key_4_list.sort(key= lambda x: returnDate(x))

    master_list = key_1_list + key_2_list + key_3_list + key_4_list

    for doc in master_list:
        print(doc.name + ": " + str(doc.date))

    with open(dir + "/ListOfDocuements_" + str(date.today()) + ".txt", "w+") as x:
        index_count = 1
        start_page = 1
        for doc in master_list:
            pdf_read = PdfFileReader(doc.path_string)
            end_page = start_page + int(pdf_read.getNumPages())
            x.write(str(index_count) + ": " + doc.name + " : " + str(doc.date) + " : " + str(start_page) + "-" + str(end_page) + "\n")
            index_count += 1
            start_page = end_page

    input("Press enter to exit")

if __name__ == "__main__":
    main()
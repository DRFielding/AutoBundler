# Bundling automation proof of concept

import PyPDF2
from PyPDF2.merger import PdfFileMerger
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
        self.date = datetime.date.fromisoformat(date_string.strip())
        self.path_string = path_string
        print(path_string)

def returnDate(doc_arg):
    return doc_arg.date

def main():
    print("%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%##%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%*  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%#  *%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%%%%%%%    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%%%%%    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%%%%    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%%(    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%#    %%%%%%%%%%     %%%%%       *%%%%%%%%%        %%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%     %%%%%%%%%%     %%%%%        %%%%%%%%         %%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%(    %%%%%%%%%%%     %%%%%         %%%%%%          %%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%     %%%%%%%%%%%     %%%%%    #     %%%%(    %     %%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%     %%%%%    #%     %%%    %%     %%%%%%%%%%%%,,%%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%     %%%%%    #%%     %    %%%     %%%%%%%%%%%    %%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%     %%%%%    #%%%        /%%%     %%%%%%%%%%(    %%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%     %%%%%    #%%%%       %%%%     %%%%%%%%%%     %%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%     %%%%%    #%%%%/     %%%%%     %%%%%%%%%%    %%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%    %%%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%    %%%%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%#   %%%%%%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%   %%%%%%%%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%   %%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%  %%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%( (%%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%,/%%%%%%%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%#%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%\n" +
    "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%\n")

    dir = input("Please enter the directory to bundle from: ")

    program_start = DateTime.now()

    if os.path.isdir(dir):
        pass
    else:
        print("Given directory not found!  Please try another directory")
        main()

    docs = os.listdir(dir)
    if docs:
        pass
    else:
        print("No files found in given directory!  Please try another folder")
        main()

    print("Finding docs...")
    doc_list = []
    for doc in docs:
        if doc.endswith(".pdf") != True:
            continue
        elif doc == "bundle.pdf": #ignore any previous bundles
            continue
        doc_proc = doc.strip(".pdf").split(";")
        doc_class = document(int(doc_proc[0]), doc_proc[1], doc_proc[2], (dir + "/" + str(doc)))
        if doc_class in doc_list:
            print("Warning!  Duplicate detected!  Document = " + str(doc_class.name))
        doc_list.append(doc_class)
    print("Done")

    print("Ordering...")
    key_1_list = [doc for doc in doc_list if doc.doc_type == "Pleadings"]
    key_2_list = [doc for doc in doc_list if doc.doc_type == "ET_Correspondence"]
    key_3_list = [doc for doc in doc_list if doc.doc_type == "Documents_Correspondence"]
    key_4_list = [doc for doc in doc_list if doc.doc_type == "Payslips"]

    key_1_list.sort(key= lambda x: returnDate(x))
    key_2_list.sort(key= lambda x: returnDate(x))
    key_3_list.sort(key= lambda x: returnDate(x))
    key_4_list.sort(key= lambda x: returnDate(x))

    master_list = key_1_list + key_2_list + key_3_list + key_4_list
    print("Done")

    with open(dir + "/ListOfDocuements_" + str(date.today()) + ".txt", "w+") as x:
        index_count = 1
        start_page = 1
        for doc in master_list:
            pdf_read = PdfFileReader(doc.path_string)
            doc_pages = int(pdf_read.getNumPages())
            if doc_pages == 1:
                end_page = start_page
            else:
                end_page = start_page + int(pdf_read.getNumPages())
            if start_page == end_page:
                x.write(str(index_count) + ": " + doc.name + " : " + str(doc.date) + " : " + str(start_page) + "\n")
            else:
                x.write(str(index_count) + ": " + doc.name + " : " + str(doc.date) + " : " + str(start_page) + "-" + str(end_page) + "\n")
            index_count += 1
            start_page = end_page + 1

    start_page = 1
    end_page = 0
    index_count = 0

    print("List completed, see text file in given directory!")

    print("Merging files and bundling...")
    pdf_merge = PdfFileMerger()
    for doc in master_list:
        pdf_merge.append(doc.path_string)
    pdf_merge.write(dir + "/bundle.pdf")
    print("Done")

    print("Bundle is completed, see bundle.pdf!")

    print("Creating table of documents...")
    table_doc = Document()
    table = table_doc.add_table(0,0)
    table.style = 'TableGrid'
    first_column_width = 5
    second_column_with = 10
    third_column_width = 10
    fourth_column_width = 10

    table.add_column(Cm(first_column_width))
    table.add_column(Cm(second_column_with))
    table.add_column(Cm(third_column_width))
    table.add_column(Cm(fourth_column_width))

    for index,doc in enumerate(master_list):
        pdf_read = PdfFileReader(doc.path_string)
        doc_pages = int(pdf_read.getNumPages())
        if doc_pages == 1:
            end_page = start_page
        else:
            end_page = start_page + int(pdf_read.getNumPages())
        if start_page == end_page:
            pages = str(start_page)
        else:
           pages = str(start_page) + "-" + str(end_page)

        index_count += 1
        start_page = end_page + 1
            
        table.add_row()
        row = table.rows[index]
        row.cells[0].text = str(index_count)
        row.cells[1].text = str(doc.name)
        row.cells[2].text = str(doc.date)
        row.cells[3].text = str(pages)

    table_doc.add_page_break()

    table_doc.save(dir + "\index.docx")

    program_end = DateTime.now()
    duration = program_end - program_start

    print("Done")

    print("Finished in " + str(duration.total_seconds()) + " seconds")

    input("Press enter to exit")

if __name__ == "__main__":
    main()
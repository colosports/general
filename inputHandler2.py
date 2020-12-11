#!/usr/bin/python3

# import library if needed
import datetime
# read input from file. 

# validate input

# Reading an excel file using Python
import pylightxl as xl

def readInput():
    # readxl returns a pylightxl database that holds all worksheets and its data
    db = xl.readxl(fn='records.xlsx')
    # read only selective sheetnames
    db = xl.readxl(fn='records.xlsx', ws=('Sheet1'))    
    # return all sheetnames
    # print(db.ws_names)
    # get individual cell
    print(db.ws(ws='Sheet1').address(address='A2'))
    print(db.ws(ws='Sheet1').address(address='B2'))

    #convert date create
    temp = datetime.date(1899, 12, 30)
    xldate = (db.ws(ws='Sheet1').address(address='C2'))
    dateCreated = temp+datetime.timedelta(days=xldate)
    print(dateCreated)

    #convert date create
    temp = datetime.date(1899, 12, 30)
    xldate = (db.ws(ws='Sheet1').address(address='E2'))
    dueCreated = temp+datetime.timedelta(days=xldate)
    print(dueCreated)

    #print(db.ws(ws='Sheet1').address(address='D2'))
    # get entire row
    headersValues = db.ws(ws='Sheet1').row(row=1)
    print (headersValues)
    # get all rows
    #for row in db.ws(ws='Sheet1').rows:
     #   print(row)
      #  print("newline")

    #for i in range(db.ws(ws='Sheet1').row(row=(i+1))))
    # check if data is valid
    # handle invalid data


def main():
    print ("Starting script")
    readInput()

if __name__ == '__main__':
    main()

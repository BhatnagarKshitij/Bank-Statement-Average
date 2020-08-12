import datetime, xlrd
from dateutil.relativedelta import relativedelta

#Global Variable Declaration 
sheet=totalRows=totalCols=Date=globalAverage=globalTotal=globalDays=endDate=None
globalTotal=0.0
globalDays=0

def init(filename="datasheet.xlsx"):
    global sheet
    fileLocation=filename
    #OpenWorkBook
    try:
        wb=xlrd.open_workbook(fileLocation)
    except:
        print("FILE NOT FOUND")
        exit()
    #GetSheet
    try:
        sheet=wb.sheet_by_index(1)
    except:
        print("NO SHEET FOUND")
        exit()
        
    #GetRowsAndCols
    global totalRows, totalCols
    totalRows=sheet.nrows
    totalCols=sheet.ncols

    if totalRows == 0 and totalCols == 0:
        print("BLANK OR CORRUPTED DOCUMENT: ")
        exit()

    if totalCols > 2:
        print("INVALID DOCUMENT")


    print("Total Rows: " +str(totalRows))
    print("Total Columns " +str(totalCols))

#--------------------------------------------------------------------------------------------#
init(filename="datasheet_2020.xlsx")
#--------------------------------------------------------------------------------------------#
def setInitDate():
    global Date,endDate
    getFirstDateFromStatement=sheet.cell_value(0,0)
    date=int(getFirstDateFromStatement[:2])
    month=int(getFirstDateFromStatement[3:5])
    year=int("20"+str(getFirstDateFromStatement[6:]))
    Date=datetime.datetime(year,month,date)    
    print("First Date: "+str(Date))
#--------------------------------------------------------------------------------------------#
setInitDate()
#--------------------------------------------------------------------------------------------#
# def dayChanged(currentDate,NextDate):
#     if(currentDate==NextDate):
#         return False
#     else:
#         return True
def formattedDate(Date):
    #FORMAT NORMAL TEXT TO DATE OBJECT
    date=Date[:2]
    month=Date[3:5]
    year=Date[6:]
    dateObj=datetime.datetime(int(str("20")+str(year)),int(month),int(date))
    return dateObj

def findAverage(months=6):
    global globalAverage,globalTotal,globalDays,endDate,Date
    endDate=Date+relativedelta(months=+months)
    print("End Date: "+str(endDate))
    rows=0
    while Date != endDate:
        if rows == totalRows-1:
            break
        #COMPARING TODAY DATE WITH NEXT DATE
        nextRowDate=sheet.cell_value(rows+1,0)
        formattedNextRowDate=formattedDate(nextRowDate)
        
        if formattedNextRowDate == Date:
            rows+=1
        else:
            globalTotal+=float(sheet.cell_value(rows,1))
            globalDays+=1
            Date += datetime.timedelta(days=1)

    # LAST DATE TO END DATE (IF TRANSATION END BEFORE THE MONTH ACTUALLY ENDS)       
    if Date != endDate:
        lastValue=sheet.cell_value(totalRows-1,1)
        while Date != endDate:
            globalTotal+=float(lastValue)
            globalDays+=1
            Date += datetime.timedelta(days=1)
            
#--------------------------------------------------------------------------------------------#
findAverage(6) # Total Month Average, DEFAULT 6
#--------------------------------------------------------------------------------------------#

print("Total: "+str(globalTotal))
print("Days: "+str(globalDays))
print("Total Average: "+ str(float(globalTotal/globalDays)))
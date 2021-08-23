import xlwings as xw
from id_validator import validator
from xlwings.main import Range

def GetSheet(bookname):
    app = xw.App(visible=False,add_book=False)
    wb = app.books.open(bookname)
    list = []
    num = len(wb.sheets)
    for i in range(0,num):
        sht = wb.sheets[i]
        list.append(sht.name)
    wb.close()
    app.kill()
    return list
    

def is_number(number):
     number = str(number)
     local = number.find(".")
     if local != -1:
         number = number[0:local]
     if len(number) == 11:
         return True
     return False


def input_error(sht):
    sht.color = 255,0,0


def verify(book,sheet,column,select):
    app = xw.App(visible=True,add_book=False)
    wb = app.books.open(book)
    sht = wb.sheets[sheet]
    CountRow = sht.range("{}1".format(column)).expand("table").rows.count
    StartRow = column + "1"
    EndRow = column + str(CountRow)
    RangeDate = sht.range(StartRow + ":" + EndRow).value
    num = 0
    none = 0
    if select:
        while True:
            num += 1
            table = sht.range(column + str(num))
            if table.value == None or table.value == "":
                none += 1
                if none != 10:
                    continue
                else:
                    break
            table.color = 255,255,0
            if not validator.is_valid(str(table.value)):
                input_error(table)
    else:
        while True:
            num += 1
            table = sht.range(column + str(num))
            if table.value == None or table.value == "":
                none += 1
                if none != 10:
                    continue
                else:
                    break
            table.color = 255,255,0
            if not is_number(str(table.value)):
                input_error(table)


def main():
    #is_id_validator(r"E:\root\Desktop\新建 XLSX 工作表.xlsx","test","B")
    return

if __name__ == "__main__":
    main()
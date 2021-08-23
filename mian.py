from subprocess import run
import sys,Ui_FormUi,Excel,threading
from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow


def get_file():
    name = QFileDialog.getOpenFileName(None,"选择文件", "/", "xlsx files (*.xlsx);;xls files (*.xls);;all files (*)")
    if name[0] != "":
        global path
        path = name[0]
        ui.filelabel.setText(path)
        global sheet_list
        sheet_list = Excel.GetSheet(path)
        ui.sheetlist.clear()
        for i in sheet_list:
            ui.sheetlist.addItem(i)
        ui.sheetlist.setCurrentRow(0)
        ui.runbutton.setEnabled(True)
        ui.rangecombom.setEnabled(True)

def run():
    column = ui.rangecombom.itemText(ui.rangecombom.currentIndex())
    selnum = ui.sheetlist.currentRow()
    t = threading.Thread(target=Excel.verify,args=(path,sheet_list[selnum],column[0],ui.idbutton.isChecked(),),daemon = True)
    t.start()

def main():
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    global ui
    ui = Ui_FormUi.Ui_MainWindow()
    ui.setupUi(MainWindow)
    ui.getfilebutton.clicked.connect(get_file)
    ui.runbutton.clicked.connect(run)
    ui.exitbutton.clicked.connect(sys.exit)
    ui.runbutton.setEnabled(False)
    ui.rangecombom.setEnabled(False)
    ui.idbutton.setChecked(True)
    for asc in range(65,90 + 1):
        ui.rangecombom.addItem("{}列".format(chr(asc)))
    MainWindow.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
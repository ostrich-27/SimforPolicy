import sys
import json
import ctypes
from PyQt5 import QtGui, QtWidgets, QtCore
from UI_Browser import Ui_AsaBrowser

class interface_browser(QtWidgets.QDialog,Ui_AsaBrowser):
    def __init__(self,parent=None):
        super(interface_browser, self).__init__(parent)
        self.setupUi(self)
        #DIY
        self.SourceData={}
        self.pushButton_path.clicked.connect(self.__getSource)
        self.pushButton_query.clicked.connect(self.__query)

    def __getSource(self):
        InputPath, i = QtWidgets.QFileDialog.getOpenFileName(self, "请选择需要查看的文件...")
        if not InputPath=="":
            with open(InputPath, 'r+', encoding="utf-8") as f:
                js = f.read()
                if not js == "":
                    self.SourceData = json.loads(js)
        if len(self.SourceData)>0:
            for k in self.SourceData.keys():
                self.comboBox_eng.addItem(k)
            for t in ["资产负债表","利润表","现金流量表"]:
                self.comboBox_type.addItem(t)
    def __query(self):
        eng=self.comboBox_eng.currentText()
        report=self.comboBox_type.currentText()
        CONTENT = self.SourceData[eng][report]
        HEADER=["Item","New_Item","Value"]
        r_count = len(CONTENT['Item'])
        c_count = 3
        model = QtGui.QStandardItemModel(r_count, c_count)
        model.setHorizontalHeaderLabels(HEADER)
        for r in range(r_count):
            for c in range(len(HEADER)):
                colName=HEADER[c]
                item = QtGui.QStandardItem(str(CONTENT[colName][r]))
                model.setItem(r,c, item)
        self.tableView.setModel(model)
        self.tableView.resizeColumnToContents(1)

if __name__=='__main__':
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("myappid")
    myApp = QtWidgets.QApplication(sys.argv)
    main1=interface_browser()
    main1.show()
    sys.exit(myApp.exec_())
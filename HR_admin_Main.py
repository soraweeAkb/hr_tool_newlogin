# coding:utf-8

from PyQt5 import QtGui,QtCore,QtWidgets,QtSql
from PyQt5.QtWidgets import QApplication
import sys
import pymysql

'''
    州的先生 - 在PyQt5中使用数据库
'''

class MainUi(QtWidgets.QMainWindow):

    def __init__(self):
        super().__init__()
        self.initUi()
        self.connect_db()

    # 初始化UI界面
    def initUi(self):
        # 设置窗口标题
        self.setWindowTitle("州的先生 - 在PyQt5中使用数据库")
        # 设置窗口大小
        self.resize(600,400)

        # 创建一个窗口部件
        self.widget = QtWidgets.QWidget()
        # 创建一个网格布局
        self.grid_layout = QtWidgets.QGridLayout()
        # 设置窗口部件的布局为网格布局
        self.widget.setLayout(self.grid_layout)

        # 创建一个按钮组
        self.group_box = QtWidgets.QGroupBox('数据库按钮')
        self.group_box_layout = QtWidgets.QVBoxLayout()
        self.group_box.setLayout(self.group_box_layout)
        # 创建一个表格部件
        self.table_widget = QtWidgets.QTableView()
        # 将上述两个部件添加到网格布局中
        self.grid_layout.addWidget(self.group_box,0,0)
        self.grid_layout.addWidget(self.table_widget,0,1)

        # 创建按钮组的按钮
        self.b_create_db = QtWidgets.QPushButton("创建数据库")
        self.b_create_db.clicked.connect(self.create_db)
        self.b_view_data = QtWidgets.QPushButton("浏览数据")
        self.b_view_data.clicked.connect(self.view_data)
        self.b_add_row = QtWidgets.QPushButton("添加一行")
        self.b_delete_row = QtWidgets.QPushButton("删除一行")
        self.b_close = QtWidgets.QPushButton("退出")
        self.b_close.clicked.connect(self.close)
        # 添加按钮到按钮组中
        self.group_box_layout.addWidget(self.b_create_db)
        self.group_box_layout.addWidget(self.b_view_data)
        self.group_box_layout.addWidget(self.b_add_row)
        self.group_box_layout.addWidget(self.b_delete_row)
        self.group_box_layout.addWidget(self.b_close)

        # 设置UI界面的核心部件
        self.setCentralWidget(self.widget)

    def connect_db(self):
        db=QtSql.QSqlDatabase.addDatabase('QMYSQL')
        db.setHostName('210.1.31.3')
        db.setUserName('hr')
        db.setPort(3306)
        db.setPassword('gwP6xTsA')
        db.setDatabaseName('TEST_AKT1')
        ok=db.open()
        if not ok:
            print(db.lastError().text())
        else: print("connected")


    def show_db(self):
        pass

    def view_data(self):
        # 实例化一个可编辑数据模型
        self.model = QtSql.QSqlTableModel()


        self.model.setTable('time_card')  # 设置数据模型的数据表
        self.model.setEditStrategy(QtSql.QSqlTableModel.OnFieldChange)  # 允许字段更改
        self.model.select()  # 查询所有数据
        # 设置表格头
        self.model.setHeaderData(0, QtCore.Qt.Horizontal, 'SERIAL')
        self.model.setHeaderData(1, QtCore.Qt.Horizontal, 'USER_ID')
        self.model.setHeaderData(2, QtCore.Qt.Horizontal, 'CLOCK_IN')

        self.table_widget.setModel(self.model)
        print('OKOK')

    def create_db(self):
        pass



if __name__ == '__main__':
    app = QApplication(sys.argv)
    MainUi=MainUi()
    MainUi.show()

    sys.exit(app.exec_())
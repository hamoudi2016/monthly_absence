
import  datetime
from PySide6 import QtCore, QtWidgets 
from load_constants import  absence_motif , categorys , months  

MyStyleSheet = """

QWidget#container {
    background-color: yellow;
    border: none;
    border-radius: 10px;
}
QWidget {
    background-image: #1f95ef; 
    color: #1f95ef;
    font-size: 16px;
    font-family: Arial, Helvetica, sans-serif;
}
QWidget#icon_name_widget {
    color: white;
    background-color: rgb(31, 149, 239);
}
QPushButton{
    background-color: #1f95ef;
	color : #f5fafe;
	font-weight : bold;
    border : none;
    border-radius : 10px;
    padding-right : 10px;
}
QPushButton:hover{
	background-color: #AB00FF;
	color : #e2e7eb;	
	font-weight : bold;
}
QPushButton:pressed{
    background-color: #AB00FF;
	color : #e2e7eb;
    font-weight : bold;
}

QComboBox{
    background-color: #1f95ef;
	color : #f5fafe;
	font-weight : bold;
	border : none;
	border-radius : 10px;
	padding-right : 10px;
}

QComboBox:hover{
	background-color: #AB00FF;
	color : #e2e7eb;
	font-weight : bold;
}
QComboBox::drop-down{
	background-color :#1f95ef;
	color : #f5fafe;
}
QSpinBox{
    background-color: #1f95ef;
	color : #f5fafe;
	font-weight : bold;
	border : none;
	border-radius : 10px;
	padding : 5px;
}

QSpinBox:hover{
	background-color: #AB00FF;
	color : #e2e7eb;
}
QDateEdit{
    background-color: #1f95ef;
	color : #f5fafe;
	font-weight : bold;
	border : none;
	border-radius : 10px;
	padding-right : 5px;
}
QDateEdit:hover{
	background-color: #AB00FF;
	color : #e2e7eb;
}
QTableWidget {
    background-color: yellow;
    color: #1f95ef;
    font-size: 16px;
    font-family: Arial, Helvetica, sans-serif;
}

QTableWidget QHeaderView {
    background-color: #1f95ef;
    color: #f5fafe;
    font-weight: bold;
    border: none;
    padding-right: 10px;
}

QTableWidget QHeaderView::section {
    background-color: #1f95ef;
    color: #f5fafe;
    font-weight: bold;
    border: none;
    padding-right: 10px;
}

QTableWidget::item {
    background-color: yellow;
    color: #1f95ef;
    border: none;
    padding-right: 10px;
}

QTableWidget::item:selected {
    background-color: #AB00FF;
    color: #e2e7eb;
    font-weight: bold;
}
"""
   
# **************************************************************************************************************************
# **************************************************************************************************************************
#                             class myGui
# **************************************************************************************************************************
#***************************************************************************************************************************
print("Class MyGui is called")
class MyGui(object):
    
    def setupUi(self, MainWindow):
        print("setupUi function of MyGui is called")
        MainWindow.setObjectName("MainWindow")
        print("Gui instance created")
        MainWindow.setLocale(QtCore.QLocale(QtCore.QLocale.Language.Arabic, QtCore.QLocale.Country.Algeria))
        MainWindow.setWindowTitle("Monthly Absences")
        MainWindow.resize(800,400)
        MainWindow.setLayoutDirection(QtCore.Qt.LayoutDirection.RightToLeft)
        self.centralwidget = QtWidgets.QWidget(parent=MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        MainWindow.setCentralWidget(self.centralwidget)
        MainWindow.setStyleSheet(MyStyleSheet)
        MainWindow.setWindowFlags(QtCore.Qt.WindowType.FramelessWindowHint)
        MainWindow.setAttribute(QtCore.Qt.WidgetAttribute.WA_TranslucentBackground)
        
        # إضافة حاوية جديدة
        self.container = QtWidgets.QWidget(parent=self.centralwidget)
        self.container.setObjectName("container")
        self.container.setGeometry(QtCore.QRect(0, 0, 800, 400))
        
        # Create labels

        self.label_titel = QtWidgets.QLabel(parent=self.centralwidget)
        self.label_category = QtWidgets.QLabel(parent=self.centralwidget)
        self.label_month = QtWidgets.QLabel(parent=self.centralwidget)
        self.label_employee = QtWidgets.QLabel(parent=self.centralwidget)
        self.label_days = QtWidgets.QLabel(parent=self.centralwidget)
        self.label_start_day = QtWidgets.QLabel(parent=self.centralwidget)
        self.label_end_day = QtWidgets.QLabel(parent=self.centralwidget)
        self.label_motif = QtWidgets.QLabel(parent=self.centralwidget)
        self.label_year = QtWidgets.QLabel(parent=self.centralwidget)

        
        self.label_category.setObjectName("label_category")
        self.label_month.setObjectName("label_month")
        self.label_employee.setObjectName("label_employee")
        self.label_days.setObjectName("label_days")
        self.label_start_day.setObjectName("label_start_day")
        self.label_end_day.setObjectName("label_end_day")
        self.label_motif.setObjectName("label_motif")
        self.label_year.setObjectName("label_year")
 

        self.label_titel.setText("الحضور والغياب")
        self.label_category.setText("الفئة")
        self.label_month.setText("الشهر")
        self.label_employee.setText("الغائب")
        self.label_days.setText("عدد الأيام")
        self.label_start_day.setText("من")
        self.label_motif.setText("السبب")
        self.label_year.setText("السنة")
        
        self.label_titel.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.label_category.setAlignment(QtCore.Qt.AlignmentFlag.AlignJustify)
        self.label_month.setAlignment(QtCore.Qt.AlignmentFlag.AlignJustify)
        self.label_employee.setAlignment(QtCore.Qt.AlignmentFlag.AlignJustify)
        self.label_days.setAlignment(QtCore.Qt.AlignmentFlag.AlignJustify)
        self.label_start_day.setAlignment(QtCore.Qt.AlignmentFlag.AlignJustify)
        self.label_end_day.setAlignment(QtCore.Qt.AlignmentFlag.AlignJustify)
        self.label_motif.setAlignment(QtCore.Qt.AlignmentFlag.AlignJustify)
        self.label_year.setAlignment(QtCore.Qt.AlignmentFlag.AlignJustify)
        
        
        # Create comboboxs
        self.comboBox_employees = QtWidgets.QComboBox(parent=self.centralwidget)
        self.comboBox_employees.setObjectName("comboBox_employees")
        
        self.comboBox_category = QtWidgets.QComboBox(parent=self.centralwidget)
        self.comboBox_category.setObjectName("comboBox_category")
        self.comboBox_category.addItems(categorys)
        
        self.comboBox_absence_motif = QtWidgets.QComboBox(parent=self.centralwidget)
        self.comboBox_absence_motif.setObjectName("comboBox_absence_motif")
        self.comboBox_absence_motif.addItems(absence_motif)
                
        self.comboBox_month = QtWidgets.QComboBox(parent=self.centralwidget)
        self.comboBox_month.setObjectName("comboBox_month")
        self.comboBox_month.addItems(months)
        self.comboBox_month.setCurrentIndex(datetime.datetime.now().month-1)

        # Create spinboxs
        
        self.spinbox_year = QtWidgets.QSpinBox(parent=self.centralwidget)
        self.spinbox_year.setObjectName("spinbox_year")
        self.spinbox_year.setMinimum(2024)
        self.spinbox_year.setMaximum(2030)
        self.spinbox_year.setValue(datetime.datetime.now().year)

        self.spinBox_nbr_jour = QtWidgets.QSpinBox(parent=self.centralwidget)
        self.spinBox_nbr_jour.setObjectName("spinBox_nbr_jour")
        self.spinBox_nbr_jour.setMinimum(1)
        self.spinBox_nbr_jour.setMaximum(98)
        
        # create dateedits
        self.dateEdit_StartDay = QtWidgets.QDateEdit(parent=self.centralwidget)
        self.dateEdit_StartDay.setMaximumSize(QtCore.QSize(300, 30))
        self.dateEdit_StartDay.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.dateEdit_StartDay.setMaximumDate(QtCore.QDate(2026, 12, 31))
        self.dateEdit_StartDay.setMinimumDate(QtCore.QDate(2022, 12, 1))
        self.dateEdit_StartDay.setDate(QtCore.QDate(datetime.datetime.now().year, datetime.datetime.now().month, 1))
        self.dateEdit_StartDay.setCalendarPopup(True)
        self.dateEdit_StartDay.setObjectName("dateEdit_StartDay")

        # Create pushbuttons
        self.pushButton_exit = QtWidgets.QPushButton(parent=self.centralwidget)
        self.pushButton_exit.setObjectName("pushButton_exit")
        self.pushButton_exit.setText("خروج")
        
        self.pushButton_insert_absence = QtWidgets.QPushButton(parent=self.centralwidget)
        self.pushButton_insert_absence.setObjectName("pushButton_insert_absence")
        self.pushButton_insert_absence.setText("إضافة")
        
        self.pushButton_delete = QtWidgets.QPushButton(parent=self.centralwidget)
        self.pushButton_delete.setObjectName("pushButton_delete")
        self.pushButton_delete.setText("حذف")
        
        self.pushButton_save = QtWidgets.QPushButton(parent=self.centralwidget)
        self.pushButton_save.setObjectName("pushButton_save")
        self.pushButton_save.setText("حفظ")
        
        self.pushButton_notices = QtWidgets.QPushButton(parent=self.centralwidget)
        self.pushButton_notices.setObjectName("pushButton_notices")
        self.pushButton_notices.setText("الاشعارات")
        
        # Create grid layout
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")

        # # Add widgets to the grid layout
        self.gridLayout.addWidget(self.label_titel, 0, 0, 1, 4)
        self.gridLayout.addWidget(self.pushButton_exit, 0, 5)
        
        self.gridLayout.addWidget(self.label_year, 2, 0)
        self.gridLayout.addWidget(self.spinbox_year, 2, 1)
        
        self.gridLayout.addWidget(self.label_month, 3, 0)
        self.gridLayout.addWidget(self.comboBox_month, 3, 1)
        
        self.gridLayout.addWidget(self.label_category, 2, 2)
        self.gridLayout.addWidget(self.comboBox_category, 2, 3)
        
        self.gridLayout.addWidget(self.label_employee, 3, 2)
        self.gridLayout.addWidget(self.comboBox_employees, 3, 3)
        
        self.gridLayout.addWidget(self.label_days, 1,4 )
        self.gridLayout.addWidget(self.spinBox_nbr_jour, 1, 5)
        
        self.gridLayout.addWidget(self.label_start_day, 2, 4)
        self.gridLayout.addWidget(self.dateEdit_StartDay, 2, 5)
        
        self.gridLayout.addWidget(self.label_motif, 3, 4)
        self.gridLayout.addWidget(self.comboBox_absence_motif, 3, 5)
        
        self.gridLayout.addWidget(self.pushButton_insert_absence, 4, 0)
        self.gridLayout.addWidget(self.pushButton_delete, 4, 2)
        self.gridLayout.addWidget(self.pushButton_save, 4, 4)
        self.gridLayout.addWidget(self.pushButton_notices, 4, 5)

        # Create a table widget
        self.tableWidget = QtWidgets.QTableWidget(parent=self.centralwidget)
        self.tableWidget.setObjectName("tableWidget")

        # Prevent editing rows and cells
        self.tableWidget.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)  
        self.tableWidget.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)  
        self.tableWidget.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)  

        # Create a layout for the table widget
        self.tableLayout = QtWidgets.QHBoxLayout()
        self.tableLayout.addWidget(self.tableWidget)
        self.gridLayout.addLayout(self.tableLayout, 5, 0, 1, 6)
        
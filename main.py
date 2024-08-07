import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QMessageBox, QTableWidgetItem , QHeaderView
from load_constants import  categorys, header_Title
from database_operation import (
    create_database,
    get_absence_records,
    get_names,
    get_info,
    add_absence_record,
    delete_absence_record,
)
from generate_doc import generate_absence_doc , generate_notices
from gui import MyGui


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = MyGui()
        self.ui.setupUi(self)
        self.setContentsMargins(0, 0, 0, 0)
        self.hide_column()
        self.load_tableWidget()
        self.ui.comboBox_category.currentIndexChanged.connect(self.load_tableWidget)
        self.ui.comboBox_month.currentIndexChanged.connect(self.load_tableWidget)
        self.ui.spinbox_year.valueChanged.connect(self.load_tableWidget)
        self.ui.pushButton_insert_absence.clicked.connect(self.add_absence)
        self.ui.pushButton_delete.clicked.connect(self.delete_absence)
        self.ui.pushButton_save.clicked.connect(self.generate_absence_montheul)
        self.ui.pushButton_notices.clicked.connect(self.genrate_notices)
        self.ui.pushButton_exit.clicked.connect(self.close)
    
    
    
    def hide_column(self):
        self.ui.tableWidget.clear()
        self.ui.tableWidget.setColumnCount(11)
        self.ui.tableWidget.setHorizontalHeaderLabels(header_Title)
        self.ui.tableWidget.verticalHeader().setVisible(False)
        self.ui.tableWidget.setColumnHidden(0, True)
        self.ui.tableWidget.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.ui.tableWidget.setColumnHidden(2, True)
        self.ui.tableWidget.setColumnHidden(3, True)
        self.ui.tableWidget.setColumnWidth(3, 150)
        self.ui.tableWidget.setColumnWidth(4, 150)
        self.ui.tableWidget.setColumnWidth(5, 150)
        self.ui.tableWidget.setColumnWidth(6, 50)
        self.ui.tableWidget.setColumnWidth(7, 150)
        self.ui.tableWidget.setColumnHidden(8, True)
        self.ui.tableWidget.setColumnHidden(9, True)
        self.ui.tableWidget.setColumnHidden(10, True)
        
    def load_tableWidget(self):
        try:
            category: int = self.ui.comboBox_category.currentIndex()
            names: list[str] = get_names(category)
            # استرجاع الفئة والشهر والسنة المحددة
            month: int = self.ui.comboBox_month.currentIndex()
            year: int = self.ui.spinbox_year.value()
            records: list[dict] = get_absence_records(category, month, year)
            # تحديث قائمة الموظفين
            self.ui.comboBox_employees.clear()
            if names is not None:
                self.ui.comboBox_employees.addItems(names)
            
            # تحديث الجدول
            self.hide_column()
            if records is not None:
                self.ui.tableWidget.setRowCount(len(records))
                for row_index, record in enumerate(records):
                    for column_index, key in enumerate(record.keys()):
                        item = QTableWidgetItem(str(record[key]))
                        self.ui.tableWidget.setItem(row_index, column_index, item)
        except Exception as e:
            QMessageBox.warning(None, "خطأ في تحديث الجدول", str(e))
            
    def add_absence(self) -> None:
        try:
            employee_name: str = self.ui.comboBox_employees.currentText()
            employee_ccp: str = get_info(employee_name)["ccp"]
            employee_grade: str = get_info(employee_name)["grade"]
            dayes: int = self.ui.spinBox_nbr_jour.value()
            start_date: str = self.ui.dateEdit_StartDay.date().toString("yyyy-MM-dd")
            end_date: str = self.ui.dateEdit_StartDay.date().addDays(dayes - 1).toString("yyyy-MM-dd")

            new_record: dict = {
                "name": employee_name,
                "ccp": employee_ccp,
                "grade": employee_grade,
                "start_date": start_date,
                "end_date": end_date,
                "days": dayes,
                "motif": self.ui.comboBox_absence_motif.currentText(),
                "category": self.ui.comboBox_category.currentIndex(),
                "month": self.ui.comboBox_month.currentIndex(),
                "year": self.ui.spinbox_year.value(),
            }

            add_absence_record(new_record)
            self.load_tableWidget()
        except Exception as e:
            QMessageBox.warning(None, "خطأ في إضافة غياب", str(e))
            

    def delete_absence(self):
        try:
            selected_row = self.ui.tableWidget.currentRow()
            selected_id = self.ui.tableWidget.item(selected_row, 0).text()
            id: int = int(selected_id)
            delete_absence_record(id)
            self.load_tableWidget()
        except Exception as e:
            QMessageBox.warning(None, "خطأ في حذف غياب", str(e))
                
    def generate_absence_montheul(self):
        year: int = self.ui.spinbox_year.value()
        month: int = self.ui.comboBox_month.currentIndex()
        word_file: str = f"غيابات {month+1}-{year}.docx"
        
        employees_absences = {
           categorys[0]: get_absence_records(0, month, year),
           categorys[1]: get_absence_records(1, month, year),
           categorys[2]: get_absence_records(2, month, year),
           categorys[3]: get_absence_records(3, month, year),
        }
        
        generate_absence_doc( word_file, month, year, employees_absences)

    def genrate_notices(self):
        year: int = self.ui.spinbox_year.value()
        month: int = self.ui.comboBox_month.currentIndex()
        category: int = self.ui.comboBox_category.currentIndex()
        employees_absences = get_absence_records(category, month, year)    
        output_docx = f"إشعارات {categorys[category]} {month + 1}-{year}.docx"
        generate_notices( output_docx, month, year, employees_absences)

                
if __name__ == "__main__":
    create_database()
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec())

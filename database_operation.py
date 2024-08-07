
import pandas as pd
import sqlite3
from load_constants import excel_file, columns_to_read 
from PySide6.QtWidgets import QMessageBox
from typing import Optional, Union, List, Dict


# **************************** READ AND CLEAN excel file  ****************************
def read_and_clean_data(excel_file: str, columns: list[str]) -> pd.DataFrame:
    """
    Reads data from an Excel file and cleans it by removing duplicates and filling NaN values.

    Args:
    - excel_file (str): The path to the Excel file.
    - columns (list[str]): The columns to read from the Excel file.

    Returns:
    - pd.DataFrame: The cleaned data as a DataFrame.
    """
    try:
        data: pd.DataFrame = pd.read_excel(excel_file, usecols=columns)
        data.drop_duplicates(inplace=True)
        data.fillna('', inplace=True)
        return data
    except Exception as e:
        print(f"Error: {str(e)}")
        return None



# ********************* CREATE DATABASE AND employees Table and absence_records_tbl TABLE ***************************************
def create_database() -> None:
    """
    Creates the 'employees_tbl'  and 'absence_records_tbl' in the 'MyDB.db' database.
    """
    try:
        with sqlite3.connect('MyDB.db') as connection:
            data: pd.DataFrame = read_and_clean_data(excel_file, columns_to_read)
            if data is not None:
                data.to_sql('employees_tbl', connection, if_exists='replace', index=False)      # create table employees
                connection.execute(
                    '''
                    CREATE TABLE IF NOT EXISTS absence_records_tbl (                     
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        name TEXT NOT NULL,
                        ccp TEXT NOT NULL,
                        grade TEXT NOT NULL,
                        start_date DATE NOT NULL,
                        end_date DATE NOT NULL,
                        days INTEGER NOT NULL,
                        motif TEXT,
                        category INTEGER NOT NULL,
                        month INTEGER NOT NULL,
                        year INTEGER NOT NULL)
                    '''
                )
                connection.execute(
                    '''
                    CREATE INDEX IF NOT EXISTS category_month_year_index
                    ON absence_records_tbl (category, month, year)
                    '''
                )
                connection.commit()
                connection.cursor()
                print("database created and table created in database")
    except sqlite3.Error as e:
        print(f"Error: {e}")
        return None

#*************************** Employee names ********************************* 

def get_names(category: int) -> list[str]:
    try:
        with sqlite3.connect('MyDB.db') as connection:
            cursor = connection.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='employees_tbl';")
            result: sqlite3.Cursor = connection.execute("SELECT name FROM employees_tbl WHERE category = ?", (category,))
            names = [row[0] for row in result.fetchall()]
            return names
    except sqlite3.Error as e:
        print(f"Error: {e}")
        return None
    
    
    
    
# ************************** Functions to get data from the database **************************
def get_absence_records(category: int, month: int, year: int) -> List[Dict[str, Optional[Union[int, str]]]]:
    try:
        query: str = '''
            SELECT id, name, ccp, grade, start_date, end_date, days, motif
            FROM absence_records_tbl
            WHERE category = ? AND month = ? AND year = ?
            ORDER BY name
        '''
        with sqlite3.connect('MyDB.db') as connection:
            
            df = pd.read_sql_query(query, connection, params=(category, month, year))
            return df.to_dict(orient='records')
    except sqlite3.Error as e:
        print(f"Error: {e}")
        return None


# ************************** Employee Info ccp and grade   *********************************
def get_info(name: str) -> Optional[dict[str, str]]:
    """
    Retrieves employee information based on the name.

    Args:
    - name (str): The name of the employee.

    Returns:
    - A dictionary containing the employee's CCP and grade,
      with keys "ccp" and "grade", respectively.
      The values are of type str.
      Returns None if the employee is not found or an error occurs.
    """
    try:
        with sqlite3.connect('MyDB.db') as connection:
            query: str = "SELECT ccp, grade FROM employees_tbl WHERE name = ?"
            result: sqlite3.Cursor = connection.execute(query, (name,))
            employee_data: Optional[tuple[str, str]] = result.fetchone()
            if employee_data:
                return {
                    "ccp": employee_data[0],
                    "grade": employee_data[1],
                }
            else: return None
    except sqlite3.Error as e:
        return None

# ************************ function to add record absence ****************************************************
def add_absence_record(record: dict) -> None:
    """
    Adds an absence record to the 'absence_records_tbl' table in the 'MyDB.db' database.

    Args:
    - record (dict): The absence record to add, containing the following fields:
        - name (str): The name of the employee.
        - ccp (str): The CCP of the employee.
        - grade (str): The grade of the employee.
        - start_date (str): The start date of the absence in 'YYYY-MM-DD' format.
        - end_date (str): The end date of the absence in 'YYYY-MM-DD' format.
        - days (int): The number of days the employee was absent.
        - motif (str): The motif of the absence.
        - category (str): the category of the employee.
        - month (int): The month of the absence.
        - year (int): The year of the absence.
    """
    try:
        with sqlite3.connect('MyDB.db') as connection:
            query = '''
                INSERT INTO absence_records_tbl (name, ccp, grade, start_date, end_date, days, motif,
                category, month, year)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            '''
            connection.execute(
                query,
                (
                    record['name'],
                    record['ccp'],
                    record['grade'],
                    record['start_date'],
                    record['end_date'],
                    record['days'],
                    record['motif'],
                    record['category'],
                    record['month'],
                    record['year']
                )
            )
            connection.commit()
        
    except sqlite3.Error as e:
        QMessageBox.warning(None, "Error", f"حدث خطأ في اضافة البيانات: {e}")

# ************************** Functions to delete data from the database **************************
def delete_absence_record(record_id: int) -> None:
    """
    Deletes an absence record from the database.

    Args:
    - record_id (int): The ID of the absence record to delete.

    Returns:
    - None
    """
    try:
        with sqlite3.connect('MyDB.db') as connection:
            connection.execute('DELETE FROM absence_records_tbl WHERE id = ?', (record_id,))
            connection.commit()
    except sqlite3.Error as e:
        QMessageBox.warning(None, "Error", f"حدث خطأ في حذف البيانات: {e}")

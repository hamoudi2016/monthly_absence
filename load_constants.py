
# import os
# from datetime import date


# # Define constants
# absence_motif: list = list(["", "عطلة مرضية","عطلة أمومة"])
# header_Title: list =list(["id", "Name", "CCP", "Grade", "Start Date","End Date", "Days", "Motif", "Category", "Month", "Anne"])
# categorys: list = list(["الأساتذة","الأساتذة المتعاقدين","الإداريين","العمال المهنيين"])
# months: list = list(["جانفي", "فيفري", "مارس", "أفريل", "ماي", "جوان","جويلية", "أوت", "سبتمبر", "اكتوبر", "نوفمبر", "ديسمبر"])
# myfont: str = "ManaraDocs Amatti Font"  # Use a font that supports Arabic characters
# ecole: str = 'متوسطة بزيط محمد'
# city: str = 'أولاد جلال'

# current_date: str = date.today().strftime("%d-%m-%Y")
# excel_file = os.path.join(os.path.dirname(__file__), "mydata.xlsx")
# columns_to_read: list = list(["name", "grade", "category","ccp"])

import json

def load_constants(file_path='constants.json'):
    with open(file_path, 'r', encoding='utf-8') as f:
        constants = json.load(f)
    return constants

constants = load_constants()

months = constants["months"]
categorys = constants["categorys"]
header_Title = constants["header_Title"]
excel_file = constants["excel_file"] 
columns_to_read = constants["columns_to_read"]
city = constants["city"]
ecole = constants["ecole"]
absence_motif = constants["absence_motif"]
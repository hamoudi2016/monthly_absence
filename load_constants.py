
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

from win32com import client
import re

# input("Нажми Enter для подключения к Excel.Application")


column = 1
excel = client.Dispatch("Excel.Application")
sheet = excel.ActiveSheet


def get_selection_row_numbers():
    # get first end last selected row
    selection = excel.Application.Selection
    start_row = selection.Cells(1).Row
    stop_row = selection.Cells(selection.Cells.Count).Row
    return start_row, stop_row

def get_change_patterns(start_row, stop_row):
    # create dict from selected aria in excel document
    changePatterns = {}
    current_row = start_row
    print(start_row, stop_row)
    while current_row <= stop_row:
        changePatterns[sheet.Rows(current_row).Cells(column).Value] = sheet.Rows(current_row).Cells(column+1).Value
        current_row += 1 
    return changePatterns
    
change_patterns = get_change_patterns(*get_selection_row_numbers())

print(f'Количество замен в файле Excel - {len(change_patterns)}')
print(change_patterns)

# prefix = r"\b("
# sufix = r"{1})((,)|($){2})"
# tempAddSufix = r"\*\(changed\)"

# print(f'Принято для замены {selection_set.Count} элемент!')

# for cnt in range(selection_set.Count):
#     item = selection_set.Item(cnt)
#     result_string = item.TextString
#     for key, value in changePatterns.items():
#         pattern = prefix + key + sufix
#         result_string = re.sub(pattern, value + "*(changed)" + r'\2', result_string)   # r'\2' it is part two of suffix and result value is end of text "$" or ","
#     item.TextString = re.sub(tempAddSufix, "", result_string)


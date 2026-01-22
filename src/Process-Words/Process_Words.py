import os
from config import data_info_original, get_folder_paths, get_file_names

# Workbook to open
from Remove_Excludes import remove_excludes

folder_paths = get_folder_paths()
file_names = get_file_names()
excel_file = os.path.join(folder_paths['result'], file_names['result_xlsx'])
# excel_file = '../Excel/words.xlsx'

p_w = remove_excludes(excel_file)

p_w.process_words()
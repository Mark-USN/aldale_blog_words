import os
from config import data_info_original, get_folder_paths, get_file_names

from Create_Sheets import Create_Sheets

# # MMH 20240525 Taken from create_topic.py
# def get_folder_paths():
#     folder_paths = {}
#     for folder_info in data_info_original:
#         folder_path = folder_info['folder_path']
#         if 'data_folder_result' in folder_path:
#             folder_paths['result'] = folder_path
#         elif 'data_folder_result' in folder_path:
#             folder_paths['result'] = folder_path
#     return folder_paths

folder_paths = get_folder_paths()
file_names = get_file_names()

if folder_paths:
    example = Create_Sheets(folder_paths, file_names) 
else:
    raise("Required folders not configured.")

# # Files used to create excel spreadsheets and location of the workbook
# # Workbook to create
# excel_file = '../Excel/words.xlsx'

# Load text, remove punctuation and blank lines and split
# into a word list
example.words_sheet()

# Create a list of case insensitive unique words and the
# number of times each appears in the text
# example.word_count_sheet()

# example.print_words_with_counts()
# print()

# Import the exclude 'always' and 'also' files and merge 
# them into a dataframe
example.excludes_sheet()

# example.print_excludes()
# print()

# Put the resultant excel workbook in the results directory
# excel_file = '../Excel/words.xlsx'

excel_file = os.path.join(folder_paths['result'], file_names['result_xlsx'])

# Use the three lists above to create a workbook with three 
# sheets. One the original words. The next of unique words
# and the number of times each appears.  And the final sheet
# containing two rows, one of words that are always excluded 
# and the other of words we want excluded for this example only
example.create_excel(excel_file)

print(f"Excel workbook created successfully at '{excel_file}'")






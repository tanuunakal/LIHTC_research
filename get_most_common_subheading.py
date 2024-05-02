from tqdm import tqdm
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import urllib
import warnings
import json
import re
from setup_functions import load_workbook_from_url, get_year_by_name, folder_list, construct_drive_url

def list_to_dict_with_count(lst):
    counts = {}
    for item in lst:
        counts[item] = counts.get(item, 0) + 1
    return counts


def get_sheet_name_funds(year):
  if int(year) > 2020:
    return "Part II-Sources of Funds"
  else:
    return "Part III-Sources of Funds"

def get_all_files_bolded(drive_dir_id):
  data = []
  for file_info in tqdm(folder_list (drive_dir_id)):
    year_str = get_year_by_name(file_info[1])
    print(year_str)
    yearly_data = get_data_for_year(file_info[0], get_sheet_name_funds(year_str))
    data = data + yearly_data
  return data


#get data for the year for all the fields in the section
def get_data_for_year(drive_dir_id, sheet_name):
  data = []
  for file_info in tqdm(folder_list (drive_dir_id)):
    bolded_data = get_bolded_by_drive_id(file_info[0], sheet_name)
    data = data + bolded_data
  return data

def get_bolded(wb, sheet_name):
  ws = wb[sheet_name]
  headings = []
  for row_num, row in enumerate(ws.iter_rows(values_only=True, max_col=5),
                                start=1):
      # Get the entity from the cell next to user_selection that has a blue fill
      for cell in ws[row_num]:
        if cell.font.bold and cell.value:
          headings.append(cell.value)
  return headings

def get_bolded_by_drive_id(drive_id, sheet_name):
  wb = load_workbook_from_url (construct_drive_url(drive_id))
  data = get_bolded (wb, sheet_name)
  print(data)
  return data
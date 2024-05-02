from tqdm import tqdm
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import urllib
import warnings
import json
import re
from setup_functions import load_workbook_from_url, get_year_by_name, folder_list, get_sheet_name, construct_drive_url

def write_sections_by_year(drive_dir_id):
  sections = {}
  for file_info in tqdm(folder_list (drive_dir_id)):
    year = int(get_year_by_name(file_info[1]))
    yearly_section = get_section_for_year(file_info[0], get_sheet_name(year))
    sections[year] = yearly_section
  with open(sections_file, 'w') as file:
    json.dump(sections, file)

def get_section_for_year(drive_id, sheet_name):
  file_of_year_info = folder_list (drive_id)[0]
  file_id = file_of_year_info[0]
  wb = load_workbook_from_url (construct_drive_url (file_id))
  ws = wb[sheet_name]
  return get_sections_for_sheet(ws)

def get_sections_for_sheet(ws):
  sections = {}
  is_search_section = False

  for row_num, row in enumerate(ws.iter_rows(values_only=True),
                                start=1):
    if is_start_of_contact_section(row):
      sections["Applicant Contact"] = [row_num]
    if is_end_of_contact_section(row):
      sections["Applicant Contact"].append(row_num)
    if is_start_of_project_location_section(row):
      sections["Project Location"] = [row_num]
    if is_end_of_project_location_section(row):
      sections["Project Location"].append(row_num)

  return sections

def is_start_of_contact_section(row):
  return ("APPLICANT CONTACT FOR APPLICATION SUBMISSION AND REVIEW" in row) \
  or ("APPLICANT CONTACT FOR APPLICATION REVIEW" in row)

def is_end_of_contact_section(row):
  return "PROJECT LOCATION" in row

def is_start_of_project_location_section(row):
  return "PROJECT LOCATION" in row

def is_end_of_project_location_section(row):
  return ("PROJECT DESCRIPTION" in row) \
  or ("WAIVERS AND/OR PRE-APPROVALS " in row)
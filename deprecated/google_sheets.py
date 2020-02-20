import io
import traceback

import gspread
from oauth2client.service_account import ServiceAccountCredentials

FILE_NAME = "FDD - RFC Final 1.0"
SHEET_NAME = "URA 33663553 VDN 3553"
SKILLS_ROW = 8
MAIN_SECTION_NAME = 'OPM'
CONFIG_FILE_NAME = "{}.cfg"
DEFAULT_CONFIG_FILE = 'default_options.cfg'
EXCLUDED_SHEETS = [
    'Decreto Aluguel e Garantia',
    'Decreto Portocap',
    'Porto Aluguel Produção',
]


def create_copy_from_default(copy_name):
    # Create a copy from a default parameters file
    with io.open(DEFAULT_CONFIG_FILE, 'r', encoding='utf-8') as f:
        with io.open(copy_name, "w", encoding='utf-8') as new_file:
            for line in f:
                new_file.write(line)
            new_file.close()


def get_client():
    # use creds to create a client to interact with the Google Drive API
    scope = ['https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
    return gspread.authorize(creds)


def normalize_value(option_value):
    if option_value.lower() == 'não':
        return 'false'
    elif option_value.lower() == 'sim':
        return 'true'
    return option_value


def get_skills(sheet):
    # Extract skills from skills row
    skills_list = sheet.row_values(SKILLS_ROW)
    skills_dict = []
    for index, skill in enumerate(skills_list):
        if skill != '' and 'TIPOSERVICO' not in skill:
            skills_dict.append({'skill': skill, 'col': index + 1})
    return skills_dict


def get_params(sheet):
    # Extract parameters from rows
    parameters_list = sheet.col_values(1)
    params_dict = []
    for index, param in enumerate(parameters_list):
        if param != '':
            params_dict.append({'param': param, 'row': index + 1})  # todo maybe save by coordinates
    return params_dict


def get_dates_pattern(days, hours):
    days_pattern = []
    if hours is not None:
        hours = hours.replace(' ', '')
        if 'seg' in days:
            days_pattern.append({'pattern1': 'W//1/{}\n'.format(hours)})
        if 'ter' in days:
            days_pattern.append({'pattern2': 'W//2/{}\n'.format(hours)})
        if 'qua' in days:
            days_pattern.append({'pattern3': 'W//3/{}\n'.format(hours)})
        if 'qui' in days:
            days_pattern.append({'pattern4': 'W//4/{}\n'.format(hours)})
        if 'sex' in days:
            days_pattern.append({'pattern5': 'W//5/{}\n'.format(hours)})
        if 'sab' in days:
            days_pattern.append({'pattern6': 'W//6/{}\n'.format(hours)})
        if 'dom' in days:
            days_pattern.append({'pattern7': 'W//0/{}\n'.format(hours)})
    return days_pattern


# Open a workbook by name
# Make sure you use the right name here.
workbook = get_client().open(FILE_NAME)
# Open worksheet by name from workbook.
worksheet_list = workbook.worksheets()

for worksheet in worksheet_list:
    if worksheet.title not in EXCLUDED_SHEETS:
        print('Working on this worksheet: {}'.format(worksheet.title))
        parameters_dict = get_params(worksheet)
        skills_dict = get_skills(worksheet)

        for skill in skills_dict:
            config_file = None
            try:
                # Create a copy from a default parameters file
                create_copy_from_default(CONFIG_FILE_NAME.format(skill['skill']))
                config_file = io.open(CONFIG_FILE_NAME.format(skill['skill']), 'r+', encoding='utf-8')

                # Get all keys/values from file and transform into dict variable
                params_dict = {}
                for line in config_file:
                    name, var = line.partition("=")[::2]
                    params_dict[name.strip()] = '{}\n'.format(var.strip())

                config_file.seek(0)  # To bring the pointer back to the starting of the file
                config_file.truncate()  # Clear file contents

                # Get all values from current column
                col = skill['col']
                col_values = worksheet.col_values(col)

                days = None
                hours = None
                # Loop through all params in google sheet
                for param in parameters_dict:
                    value = normalize_value(col_values[param['row']-1])
                    values = param['param'].split(';')

                    if param['param'] == 'days':
                        days = value
                        continue

                    if param['param'] == 'hours':
                        hours = value
                        continue

                    # Checks if it is has multiple sub parameters inside this parameter
                    if len(values) > 1:
                        for index, val in enumerate(values):
                            splited_value = worksheet.cell(param['row'], col + index)
                            final_value = normalize_value(splited_value.value)

                            # Replace value at params_dict if value is set in google sheet
                            if val != '':
                                params_dict[val] = '{}\n'.format(final_value)
                    else:
                        params_dict[param['param']] = '{}\n'.format(value)

                # Writes all changed options to final options file
                days = get_dates_pattern(days, hours)
                for day in days:
                    for key in day:
                        params_dict[key] = day[key]

                for key in params_dict:

                    line = '{} = {}'.format(key, params_dict[key])

                    # Skip not needed lines
                    if key in ['days', 'hours']:
                        continue

                    # Checks if is a blank line
                    if key == '':
                        line = '\n'

                    # Checks if it is a section row e.g [OPM]
                    elif '[' in key:
                        line = '{}\n'.format(key)
                    config_file.write(line)
            except IndexError as ie:
                traceback.print_exc()
            except ValueError as ve:
                traceback.print_exc()
            finally:
                config_file.close()
    else:
        print('Excluded this worksheet: {}'.format(worksheet.title))

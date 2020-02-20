import csv
import io
import os
import traceback

import xlrd

from base import create_copy_from_default, normalize_value, get_dates_pattern

FILE_NAME = "sheets/fdd_final_2.1.xlsx"
SHEET_NAME = "URA 33663553 VDN 3553"
SKILLS_ROW = 7
MAIN_SECTION_NAME = 'OPM'
GENERATED_DIR = 'generated'
CONFIG_FILE_NAME = os.path.join(GENERATED_DIR, "options", "{}.cfg")
DEFAULT_CONFIG_FILE = 'default_options.cfg'
VQ = "vq"
VQ_DB_DC = "vq_db_dc"
VAG = "vag"
VQG = "vqg"
SKILLS = "skills"
TYPE_DICT = {
    VQ: {"fields": ["Number", "Type", "State", "Alias", "Register", "Route Type", "Switch Specific Type", "Association"]},
    VAG: {"fields": ["Agent Group Name", "State"]},
    VQG: {"fields": ["DN Group Name", "Group Type", "State"]},
    SKILLS: {"fields": ["Skill Name"]}
}
VQ_LIST = ['vq_principal']
VQ_DB_DC_LIST = ['vq_db', 'vq_dc']
VAG_LIST = ['vag_oper', 'vag_n1', 'vag_all']
VQG_LIST = ['vqg_area', 'vqg_oper', 'vqg_n1']
SKILLS_LIST = ['basica_skill1', 'adicional1_skill1', 'adicional2_skill1', 'adicional3_skill1']
EXCLUDED_SHEETS = [
    'Áudios',
    'Agentes',
    'MiniURA',
    'Apoio RFC'
]

writers = {}


def get_writer(type):
    if type in writers:
        return writers[type]
    else:
        return create_csv_writer(type)


def get_skills(sheet):
    # Extract skills from skills row
    skills_list = sheet.row_values(SKILLS_ROW)
    skills_dict = []
    for index, skill in enumerate(skills_list):
        if skill != '' and 'TIPOSERVICO' not in skill:
            skills_dict.append({'skill': skill, 'col': index})
    return skills_dict


def get_params(sheet):
    # Extract parameters from rows
    parameters_list = sheet.col_values(0)
    params_dict = []
    for index, param in enumerate(parameters_list):
        if param != '':
            params_dict.append({'param': param, 'row': index})  # todo maybe save by coordinates
    return params_dict


def create_csv_writer(type):
    type_dict = TYPE_DICT[VQ]
    if type in VAG_LIST:
        type_dict = TYPE_DICT[VAG]
    elif type in VQG_LIST:
        type_dict = TYPE_DICT[VQG]
    elif type in SKILLS_LIST:
        type_dict = TYPE_DICT[SKILLS]
    file_dir = os.path.join(GENERATED_DIR, "objects", "{}.csv".format(type))
    csv_file = open(file_dir, 'w', newline='')
    fieldnames = type_dict["fields"]
    writer = csv.DictWriter(csv_file, fieldnames=fieldnames, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
    writer.writeheader()
    writers[type] = writer
    return writer


def write_vq_row(key, value):
    suffix = '_sw_sips_rtr_dc'
    writer = get_writer(key)
    try:
        if key.split('_')[1] == 'db':
            suffix = '_sw_sips_rtr_db'
    except IndexError:
        traceback.print_exc()
        pass
    value = value.strip()
    row_dict = {
        'Number': value,
        'Type': 'Virtual Queue',
        'State': 'Enabled',
        'Alias': '{}{}'.format(value, suffix),
        'Register': 'True',
        'Route Type': 'Default',
        'Switch Specific Type': '1',
        'Association': '',
    }
    writer.writerow(row_dict)


def write_vqg_row(key, value):
    writer = get_writer(key)
    row_dict = {
        "DN Group Name": value.strip(),
        "Group Type": "ACD Queues",
        "State": "Enabled",
    }
    writer.writerow(row_dict)


def write_vag_row(key, value):
    writer = get_writer(key)
    row_dict = {
        'Agent Group Name': value.strip(),
        'State': 'Enabled',
    }
    writer.writerow(row_dict)


def write_skill_row(key, value):
    writer = get_writer(key)
    row_dict = {
        'Skill Name': value.strip(),
    }
    writer.writerow(row_dict)


# Open a workbook by name
# Use pandas package to open Excel file
workbook = xlrd.open_workbook(FILE_NAME)

# Open worksheet by name from workbook.
worksheet_list = workbook.sheet_names()

for sheet_name in worksheet_list:
    if sheet_name not in EXCLUDED_SHEETS:
        worksheet = workbook.sheet_by_name(sheet_name)
        print('Working on this worksheet: {}'.format(sheet_name))

        # Create folder for current sheet
        config_dir = os.path.join(GENERATED_DIR, "options", sheet_name)
        # if not os.path.exists(config_dir):
        #     os.makedirs(config_dir)

        parameters_dict = get_params(worksheet)
        skills_dict = get_skills(worksheet)

        for skill in skills_dict:
            config_file = None
            try:
                # Create a copy from a default parameters file
                filename = CONFIG_FILE_NAME.format(skill['skill'])
                create_copy_from_default(DEFAULT_CONFIG_FILE, filename)
                config_file = io.open(filename, 'r+', encoding='utf-8')

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
                    value = normalize_value(col_values[param['row']])
                    values = param['param'].split(';')
                    col_value = col_values[param['row']]

                    key = param['param']
                    if key == 'days':
                        days = value
                        continue

                    if key == 'hours':
                        hours = value
                        continue

                    # Add VQs to .csv VQs file (to import in Genesys Administrator)
                    if key in VQ_LIST:
                        write_vq_row(key, col_value)
                        continue

                    # Add VQG to .csv VQGs file (to import in Genesys Administrator)
                    if key in VQG_LIST:
                        write_vqg_row(key, col_value)
                        continue

                    # Add VAG to .csv VAGs file (to import in Genesys Administrator)
                    if key in VAG_LIST:
                        write_vag_row(key, col_value)
                        continue

                    if key in VQ_DB_DC_LIST:
                        write_vq_row(key, value)
                        if key == 'vq_db':
                            suffix = '_sw_sips_rtr_db'
                        else:
                            suffix = '_sw_sips_rtr_dc'
                        value = '{}{}'.format(value, suffix)
                    try:
                        if key.split(';')[0] in SKILLS_LIST:
                            write_skill_row(key.split(';')[0], col_value)
                    except IndexError as e:
                        traceback.print_exc()
                        pass

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
                    if key in ['days', 'hours'] or key in VQ_LIST or key in VAG_LIST or key in VQG_LIST:
                        continue

                    # Checks blank lines
                    if key == '':
                        line = '\n'

                    # Checks section rows. Example: [OPM]
                    elif '[' in key:
                        line = '{}\n'.format(key)

                    config_file.write(line)
            except Exception as e:
                traceback.print_exc()
            finally:
                config_file.close()
    else:
        print('Excluded this worksheet: {}'.format(sheet_name))

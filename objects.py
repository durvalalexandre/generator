import csv
import io
import os
import traceback

import xlrd

from base import create_copy_from_default, normalize_value, get_dates_pattern

####FOCO AQUI
PRODUCTION = False
FILE_NAME = "20200113_Estrutura de Roteamento_CorretoresPotenciais_FDD.xlsx"
#################

INPUT_FILE_NAME = False
ROUTING_SHEET_NAME = "Estrategia_Roteamento"
VQS_SHEET_NAME = "Fila Virtual_VQ"
SKILLS_ROW = 1
PARAMS_ROW = 1
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
    VAG: {"fields": ["Agent Group Name", "State", "Section Name", "Option Name", "Option Value"]},
    VQG: {"fields": ["DN Group Name", "Group Type", "State"]},
    SKILLS: {"fields": ["Skill Name"]}
}
VQ_LIST = ['vq_principal']
VQ_DB_DC_LIST = ['vq_db', 'vq_dc']
VAG_LIST = ['vag_oper', 'vag_n1', 'vag_all', 'vag_out']
VQG_LIST = ['vqg_area', 'vqg_oper', 'vqg_n1']
SKILLS_LIST = ['basica_skill1', 'adicional1_skill1', 'adicional2_skill1', 'adicional3_skill1']

FRASES_HOMOLOGACAO = {
    'HORARIO_SEG_SEX_0815_1730_EXCETO_FERIADOS': '1764',
    'HORARIO_SEG_SEX_0815_1830_EXCETO_FERIADOS': '1060',
    'HORARIO_SEG_SEX_0815_1800_EXCETO_FERIADOS': '1596',
    'HORARIO_SEG_SEX_0800_1830_EXCETO_FERIADOS': '1692',
    'HORARIO_SEG_SEX_0800_2000_EXCETO_FERIADOS': '1161',
    'HORARIO_SEG_SEX_0900_1800_EXCETO_FERIADOS': '1073',
    'HORARIO_SEG_SEX_1000_1900_EXCETO_FERIADOS': '1759',
    'HORARIO_SEG_SEX_0700_1900_EXCETO_FERIADOS': '1648',
    'HORARIO_SEG_SEX_0830_1730_EXCETO_FERIADOS':'1813',
    'HORARIO_SEG_SEX_0800_1800_SAB_0800_1500': '1771',
    'HORARIO_SEG_SEX_0800_1800': '1772',
    'SUSSURRO_ESPEC_NUCLEO_APOIO': '1760',
    'SUSSURRO_ESPEC_CANAL_BANCO_SAC': '1761',
    'SUSSURRO_ESPEC_PRIVATE': '1762',
    'SUSSURRO_ESPEC_DIRETORES': '1763',
    'SUSSURRO_RE_SINISTRO_IMOBILIARIA': '1765',
    'ANUNCIO_49801_AGUARDAR': '1683',
    'MSG_CARNAVAL_QUA_13H': '1067',
}

writers = {}


def get_writer(writer_type):
    if writer_type in writers:
        return writers[writer_type]
    else:
        return create_csv_writer(writer_type)


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
    parameters_list = sheet.row_values(PARAMS_ROW)
    params_dict = []
    for index, param in enumerate(parameters_list):
        if param != '':
            params_dict.append({'param': param, 'col': index})
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
    suffixes = ['dc', 'db']
    suffixes = suffixes[:1] if not PRODUCTION else suffixes
    for suffix in suffixes:
        writer = get_writer('{}_{}'.format(key, suffix))
        alias = '{}_sw_sips_rtr_{}'.format(value, suffix)
        row_dict = {
            'Number': value,
            'Type': 'Virtual Queue',
            'State': 'Enabled',
            'Alias': alias,
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
        "State": "Enabled"
    }
    writer.writerow(row_dict)


def write_vag_row(key, value):
    writer = get_writer(key)
    row_dict = {
        'Agent Group Name': value.strip(),
        'State': 'Enabled',
    }
    skill_name = None
    comparator = None
    if key == "vag_all":
        skill_name = "sk_" + value[4:].lower()
        comparator = '> 1'
    if key == "vag_n1":
        skill_name = "sk_" + value[7:].lower()
        comparator = '= 10'
    if skill_name is not None and comparator is not None:
        row_dict['Section Name'] = 'virtual'
        row_dict['Option Name'] = 'script'
        row_dict['Option Value'] = 'Skill("{skill}") {comparator}'.format(skill=skill_name, comparator=comparator)

    writer.writerow(row_dict)


def write_skill_row(key, value):
    writer = get_writer(key)
    row_dict = {
        'Skill Name': value.strip(),
    }
    writer.writerow(row_dict)


valid_filename = False
workbook_filename = FILE_NAME
if INPUT_FILE_NAME:
    while not valid_filename:
        workbook_filename = input("Insira o nome do arquivo .xslx: ")
        try:
            if workbook_filename.split('.')[-1:][0] == 'xlsx':
                valid_filename = True
            else:
                raise IndexError
        except IndexError:
            print('Nome de arquivo inválido.')
            continue


# Open a excel workbook by name
workbook = xlrd.open_workbook(os.path.join("sheets", workbook_filename))

options_worksheet = workbook.sheet_by_name(ROUTING_SHEET_NAME)
vqs_worksheet = workbook.sheet_by_name(VQS_SHEET_NAME)
print('Processing file: {}'.format(workbook_filename))

for i, row in enumerate(options_worksheet.get_rows()):
    if i > 1:
        config_file = None

        # Create file for current sheet
        config_dir = os.path.join(GENERATED_DIR, "options", options_worksheet.name)

        parameters_dict = get_params(options_worksheet)
        # skills_dict = get_skills(worksheet)

        try:
            skill = row[5].value
            # Create a copy from a default parameters file
            filename = CONFIG_FILE_NAME.format(skill)
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

            days = None
            hours = None
            special_days = None
            # Loop through all params in selected sheet
            for param in parameters_dict:
                key = param['param']
                value = normalize_value(row[param['col']].value)
                values = []
                if ';' in key:
                    keys = param['param'].split(";")
                    operator = value[:2]
                    grade = value[2:]
                    params_dict[keys[0]] = '{}\n'.format(operator)
                    params_dict[keys[1]] = '{}\n'.format(grade)
                else:
                    if key == 'days':
                        days = value.lower()
                        continue

                    if key == 'special_day':
                        special_days = value.lower()

                        def get_formatted_day(raw_day):
                            try:
                                splitted = raw_day.split('/')
                                return 'D//*-{month}-{day}/'.format(month=splitted[1], day=splitted[0])
                            except IndexError:
                                return None

                        special_days = list(map(get_formatted_day, special_days.split(' - ')))
                        params_dict[key] = 'OPM_special_day\n' if special_days[0] is not None else params_dict[key]
                        continue

                    if key == 'hours':
                        hours = value
                        continue

                    if key in ['frase_atendimento', 'frase_feriado', 'frase_sussurro', 'anuncio_fila1', 'anuncio_fila2']:
                        try:
                            value = value if PRODUCTION else FRASES_HOMOLOGACAO[value]
                        except KeyError:
                            pass

                    if key == 'vq':
                        alias = 'VQ_{}_sw_sips_rtr_{}'
                        vq_db_suffix = 'db' if PRODUCTION else 'dc'
                        params_dict['vq_db'] = '{}\n'.format(alias.format(value.upper(), vq_db_suffix))
                        params_dict['vq_dc'] = '{}\n'.format(alias.format(value.upper(), 'dc'))

                    if value != "":
                        params_dict[key] = '{}\n'.format(value)

            # Writes all changed options to final options file
            days = get_dates_pattern(days, hours)
            for day in days:
                for key in day:
                    params_dict[key] = day[key]

            for key in params_dict:
                line = '{} = {}'.format(key, params_dict[key])

                # Skip not needed lines
                if key in ['days', 'hours', 'tiposervico', 'vq'] or key in VQ_LIST or key in VAG_LIST or key in VQG_LIST:
                    continue

                # Checks blank lines
                if key == '':
                    line = '\n'

                # Checks section rows. Example: [OPM]
                elif '[' in key:
                    line = '{}\n'.format(key)

                config_file.write(line)

            for x, day in enumerate(special_days):
                if day is not None:
                    if x == 0:
                        config_file.write('\n[OPM_special_day]\n')
                    config_file.write(f'pattern{x + 1} = {day}\n')

        except IndexError as e:
            traceback.print_exc()
        finally:
            config_file.close()

config_file = None
print(f"Options for file '{workbook_filename}' processed.")

# Create file for current sheet
config_dir = os.path.join(GENERATED_DIR, "objects", vqs_worksheet.name)
skills_dict = get_skills(vqs_worksheet)

try:
    for param in skills_dict:
        vq = param['skill']
        filename = CONFIG_FILE_NAME.format(vq)
        col_values = vqs_worksheet.col_values(param['col'])
        for j, val in enumerate(col_values):
            if j > 1 and val != "":
                # Add VQs to .csv VQs file (to import in Genesys Administrator)
                if vq in VQ_LIST:
                    write_vq_row(vq, val)
                    continue

                # Add VQG to .csv VQGs file (to import in Genesys Administrator)
                if vq in VQG_LIST:
                    write_vqg_row(vq, val)
                    continue

                # Add VAG to .csv VAGs file (to import in Genesys Administrator)
                if vq in VAG_LIST:
                    write_vag_row(vq, val)
                    continue

                if vq in VQ_DB_DC_LIST:
                    write_vq_row(vq, val)
except IndexError as e:
    traceback.print_exc()
print(f"Objects for file '{workbook_filename}' processed.")
print(f"Processing of file '{workbook_filename}' has been completed.")

import csv
import os
import traceback

import xlrd

FILE_NAME = "fdd_rfc_2.1_prod.xlsx"
SHEET_NAME = "Agentes"
HEADER_ROW = 1
AGENT_LOGIN_HEADER = [
    "Action"
    "Code",
    "Switch",
    "Switch-specific Type",
    "Enabled"
]

USERNAME = "User Name"
EMPLOYEE_ID = "EmployeeID"
EXTERNAL_ID = "External ID"

additional_agent_headers = {
    # "Action": "UPDATE",
    "Agent Flag": "True",
    "State": "Enabled",
    "Password change": "N",
    EMPLOYEE_ID: "",
    EXTERNAL_ID: ""
}

SKILLS_HEADERS = {
    "Skill Added": "",
    "Skill Level": "",
}

SKILLS_COUNT = 3

GENERATED_DIR = os.path.join('generated', 'agents')
writers = {}
FIELDS = []
AGENTS = "agents"
LOGINS = "logins"


def get_header_row(sheet):
    # Extract header row attributes from selected sheet
    header_row_values = sheet.row_values(HEADER_ROW)
    header_row = []
    for i, header in enumerate(header_row_values):
        if header != '':
            header_row.append({"key": header, "val": i})
    for header in additional_agent_headers:
        header_row.append({"key": header, "val": 0})
    return header_row


def get_row_as_dict(row):
    row_dict = []
    for a in row:
        row_dict.append(a["key"])
    return row_dict


def create_csv_writer(category, header_row):
    file_dir = os.path.join(GENERATED_DIR, "{}.csv".format(category))
    csv_file = open(file_dir, 'w', newline='', encoding='utf-8')
    writer = csv.DictWriter(csv_file, fieldnames=get_row_as_dict(header_row), delimiter=',', quoting=csv.QUOTE_NONE)
    writer.writeheader()
    writers[category] = writer
    return writer


def write_agent_row(agent_row):
    writer = writers[AGENTS]
    writer.writerow(agent_row)


def add_single_attribute(header, attribute):
    return header


def get_agent_login(header, attribute):
    return header


def get_skill(header, attribute):
    return header


def get_agent_group(header, attribute):
    return header


def get_access_group(header, attribute):
    return header


def add_to_header(header, index, attribute):
    if attribute not in header:
        header[attribute] = index
    return header


# Open a workbook by name
workbook = xlrd.open_workbook(FILE_NAME)

# Open worksheet by name from workbook.
worksheet_list = workbook.sheet_names()
worksheet = workbook.sheet_by_name(SHEET_NAME)
print('Working on this worksheet: {}'.format(SHEET_NAME))

# Create folder for current sheet
header_row = get_header_row(worksheet)

if AGENTS in writers:
    agent_writer = writers[AGENTS]
else:
    agent_writer = create_csv_writer(AGENTS, header_row)

for index, row in enumerate(worksheet.get_rows()):
    try:
        if index > 1:
            agent_row = header_row.copy()
            for header in agent_row:
                header["val"] = row[header["val"]].value
            # write_agent_row(agent_row)
            print(agent_row)
    except Exception as e:
        traceback.print_exc()


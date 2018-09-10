import os
import pathlib
from collections import OrderedDict

import pandas as pd

# Utility constants
file_suffix = ['short', 'full']
float_precision = 3

# dictionaries to convert dict names to xl naming conventions
schemas = [
    {  # short
        'B3F_id': 'B3F ID',
        'name': 'Name',
        'type': 'Type',
        'desc': 'Description',
        'loc': 'Location Path',
        'cls': 'Classe di Resistenza CLS',
        'status': 'Status',
        'n_issues': '# Issues',
        'n_open_issues': '# Open Issues',
        'n_checklists': '# Checklists',
        'n_open_checklists': '# Open Checklists',
        'date_created': 'Date Created',
        'contractor': 'Appaltatore',
        'completion_percentage': 'Percentuale di Completamento',
        'pillar_number': 'n° pilasto',
        'superficial_quality': 'Qualità superficiale getto',
        'phase': 'Phase',
        'temperature': 'Flow Temperature',
        'moisture': 'Flow Moisture',
        'pressure': 'Flow Pressure',
        'record_timestamp': 'Timestamp',
        'BIM_id': 'BIM Object ID'
    }, {  # full
        'BIM_id': 'BIM Object ID',
        'temperature': 'Flow Temperature',
        'moisture': 'Flow Moisture',
        'pressure': 'Flow Pressure',
        'phase': 'Phase',
        'status': 'Status',
        'begin_timestamp': 'Begin Timestamp',
        'end_timestamp': 'End Timestamp'
    }]


def convert_dict_keys(old_dict, conversion_table):
    """Converts a dictionary to an equal dictionary,
    changing the keys according to the given conversion table"""
    converted_dict = OrderedDict()

    for key, value in old_dict.items():
        try:
            converted_key = conversion_table[key]
            converted_dict[converted_key] = value
        except KeyError:
            # if the key is not contained in the conversion table, drop that element
            continue
    return converted_dict


def read_data_from_spreadsheet(file_name):
    """Reads data from the given spreadsheet"""
    sheet_path = pathlib.Path.cwd().joinpath('res', file_name)

    # loads data in a dataframe with the structure "moisture", "pressure", "temperature"
    xl_df = pd.read_excel(open(os.fspath(sheet_path), 'rb'), sheet_name='Sheet1')
    return xl_df


def append_row_to_spreadsheet(log, file_detail):
    """Generates and writes the summarized line on the Excel spreadsheet"""
    # Setting the right path to the spreadsheet, according to the type of summary
    sheet_path = pathlib.Path.cwd().joinpath('res', 'report-' + file_detail + '-test.xlsx')
    # print("Updating file at: " + os.fspath(sheet_path))

    # Loading previous Excel data into dataframe
    xl_df = pd.read_excel(open(os.fspath(sheet_path), 'rb'), sheet_name='Sheet1')

    # Key conversion of the dictionary in order to match with the schema on the destination file
    schema_index = file_suffix.index(file_detail)
    converted_log = convert_dict_keys(log, conversion_table=schemas[schema_index])
    # pprint("Converted log: " + str(converted_log))

    # Update previous dataframe with the new row
    new_row = pd.DataFrame.from_records([converted_log])
    new_df = pd.concat([xl_df, new_row], ignore_index=True, sort=True)

    # Avoid sorting columns when concatenating dataframes
    new_df = new_df.reindex_axis(xl_df.columns, axis=1)

    # Write update dataframe to Excel spreadsheet
    try:
        writer = pd.ExcelWriter(os.fspath(sheet_path))
        new_df.to_excel(writer, 'Sheet1', index=False)
        writer.save()
    except PermissionError:
        print("The file is currently in use! Try closing it and sending the data again")
        return False
    return True

# -*- coding: utf-8 -*-
import xlrd
import argparse
from collections import defaultdict
from datetime import datetime, timedelta

def load_excel(file_path):
    try:
        workbook = xlrd.open_workbook(file_path)
        sheet = workbook.sheet_by_index(0)  # Assuming the first sheet
        header = [cell.value for cell in sheet.row(0)]
        data = [sheet.row_values(i) for i in range(1, sheet.nrows)]
        return header, data
    except Exception as e:
        print(u"Error loading Excel file '{}': {}".format(file_path, e))
        return None, None

def excel_serial_to_date(serial):
    # Excel's serial date starts from 1900-01-01
    base_date = datetime(1899, 12, 30)
    return (base_date + timedelta(days=serial)).strftime('%Y-%m-%d')

def convert_hours(hours_str):
    try:
        if isinstance(hours_str, float):
            hours_str = str(hours_str)  # Convert float to string
        if u',' in hours_str:
            hours, minutes = map(int, hours_str.split(u','))
            return hours + minutes / 60.0
        else:
            return float(hours_str)
    except (ValueError, AttributeError):
        return 0.0

def aggregate_excel_file(file_path, header_mapping):
    header, data = load_excel(file_path)

    if data is None:
        return

    # Map header names to indices for the given file
    name_index = header.index(header_mapping[u'Full name'])
    hours_index = header.index(header_mapping[u'Hours'])
    date_index = header.index(header_mapping[u'Date created'])

    # Create a dictionary to aggregate hours by name and date
    aggregate_data = defaultdict(lambda: defaultdict(float))

    for row in data:
        name = row[name_index]
        hours = convert_hours(row[hours_index])
        date = row[date_index]

        if isinstance(date, float):
            date_str = excel_serial_to_date(date)  # Convert Excel serial to date string
        else:
            date_str = date

        # Konverter navnet til en Unicode-streng
        name = name.encode('utf-8').decode('utf-8', 'ignore')

        aggregate_data[name][date_str] += hours

    return aggregate_data

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=u"Aggregate 'Hours' by 'Full name' on the same date from different files.")
    parser.add_argument("file1", help=u"Path to the first Excel file")
    parser.add_argument("file2", help=u"Path to the second Excel file")

    args = parser.parse_args()

    file1_header_mapping = {
        u'Full name': u'Full name',
        u'Hours': u'Hours',
        u'Date created': u'Date created'
    }

    file2_header_mapping = {
        u'Full name': u'Full name',
        u'Hours': u'Hours',
        u'Date created': u'Work date'
    }

    aggregate_data_file1 = aggregate_excel_file(args.file1, file1_header_mapping)
    aggregate_data_file2 = aggregate_excel_file(args.file2, file2_header_mapping)

    combined_data = defaultdict(lambda: defaultdict(float))

    for name, date_hours in aggregate_data_file1.items():
        combined_data[name].update(date_hours)

    for name, date_hours in aggregate_data_file2.items():
        combined_data[name].update(date_hours)

    for name, date_hours in sorted(combined_data.items()):
        name = name.encode('utf-8').decode('utf-8', 'ignore')  # Ignorer ikke-ASCII-tegn
        print(u"Full name: {}".format(name))
        for date, hours in sorted(date_hours.items()):
            total_hours_file1 = aggregate_data_file1[name].get(date, 0.0)
            total_hours_file2 = aggregate_data_file2[name].get(date, 0.0)
            total_hours_combined = total_hours_file1 + total_hours_file2
            print(u"  Date: {}, Total Hours: {:.2f} (File 1: {:.2f}, File 2: {:.2f})".format(date, total_hours_combined, total_hours_file1, total_hours_file2))

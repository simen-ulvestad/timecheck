from flask import Flask, request, render_template
import pandas as pd
from collections import defaultdict
from datetime import datetime, timedelta
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Set the directory where uploaded files will be stored temporarily
app.config['UPLOAD_FOLDER'] = 'uploads'
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

# Function to load an Excel file and return the header and data
def load_excel(file_path):
    try:
        data = pd.read_excel(file_path)
        return data
    except Exception as e:
        print("Error loading Excel file '{}': {}".format(file_path, e))
        return None

# Function to convert hours in various formats to a float
def convert_hours(hours_str):
    try:
        if isinstance(hours_str, float):
            hours_str = str(hours_str)
        if ',' in hours_str:
            hours, minutes = map(int, hours_str.split(','))
            return hours + minutes / 60.0
        else:
            return float(hours_str)
    except (ValueError, AttributeError):
        return 0.0

# Function to aggregate data from an Excel file based on header mapping
def aggregate_excel_file(file_path, header_mapping):
    data = load_excel(file_path)

    if data is None:
        return

    aggregate_data = defaultdict(lambda: defaultdict(float))

    for _, row in data.iterrows():
        name = row[header_mapping['Full name']]
        hours = convert_hours(row[header_mapping['Hours']])
        date = row[header_mapping['Date created']]

        if isinstance(date, (float, int)):  # Check if date is a valid numeric type
            date_str = excel_serial_to_date(date)
        else:
            try:
                date_str = date.strftime('%Y-%m-%d')
            except AttributeError:
                date_str = str(date)

        # Parse the time from the "Work date"
        if ' ' in date_str:
            date_str, time_str = date_str.split(' ')
            date_hours = datetime.strptime(date_str, '%Y-%m-%d')
            time_hours = datetime.strptime(time_str, '%H:%M')
            date_hours += timedelta(hours=time_hours.hour, minutes=time_hours.minute)
        else:
            date_hours = datetime.strptime(date_str, '%Y-%m-%d')

        name_str = name

        aggregate_data[name_str][date_hours.strftime('%Y-%m-%d')] += hours

    return aggregate_data

# Function to convert Excel serial date to a formatted date string
def excel_serial_to_date(serial):
    base_date = datetime(1899, 12, 30)
    return (base_date + timedelta(days=serial)).strftime('%Y-%m-%d')

# Route to the home page
@app.route('/')
def home():
    return render_template('index.html')

# Route to upload and process files
@app.route('/process', methods=['POST'])
def upload_files():
    file1 = request.files['file1']
    file3 = request.files['file3']

    # Ensure uploaded files are secure and have allowed extensions
    allowed_extensions = {'.xls', '.xlsx', 'xlsx', 'xls'}
    uploaded_files = []

    for uploaded_file in [file1, file3]:
        if '.' in uploaded_file.filename:
            file_extension = uploaded_file.filename.rsplit('.', 1)[1].lower()
            if file_extension in allowed_extensions:
                filename = secure_filename(uploaded_file.filename)
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                uploaded_file.save(file_path)
                uploaded_files.append(file_path)
            else:
                app.logger.error(f"Invalid file format: {file_extension}")
                return "Invalid file format. Allowed formats are .xls and .xlsx."

    # Perform aggregation on file1 
    file1_header_mapping = {
        'Full name': 'Full name',
        'Hours': 'Hours',
        'Date created': 'Work date'
    }

    aggregate_data_file1 = aggregate_excel_file(uploaded_files[0], file1_header_mapping)

    # Perform aggregation on file3
    file3_header_mapping = {
        'Full name': 'Description',
        'Hours': 'Quantity',
        'Date created': 'Posting Date'
    }
    aggregate_data_file3 = None  # Set initial value to None
    if len(uploaded_files) > 1:
        aggregate_data_file3 = aggregate_excel_file(uploaded_files[1], file3_header_mapping)

    # Combine data from file1
    combined_data = defaultdict(lambda: defaultdict(float))

    if aggregate_data_file1 is not None:
        for name, date_hours in aggregate_data_file1.items():
            for date, hours in date_hours.items():
                combined_data[name][date] += hours

    # Compare with data from file3
    compared_data = defaultdict(lambda: defaultdict(float))

    if aggregate_data_file3 is not None:
        for name, date_hours in aggregate_data_file3.items():
            for date, hours in date_hours.items():
                if name in combined_data and date in combined_data[name]:
                    date_info = {}
                    date_info['diff'] = combined_data[name][date] - hours
                    date_info['file1_hours'] = aggregate_data_file1[name].get(date, 0)
                    date_info['date_hours'] = combined_data[name][date]
                    date_info['file3_hours'] = hours

                    compared_data[name][date] = date_info

    # Sort the compared_data by date
    for name, date_hours in compared_data.items():
        compared_data[name] = dict(sorted(date_hours.items()))

    return render_template('results.html', data=compared_data)

# Ensure the upload folder exists
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

if __name__ == "__main__":
      app.run(host='0.0.0.0', port=3005, debug=True)

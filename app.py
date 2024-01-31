import os
import re
import datetime
import gspread
import pandas as pd
from flask import Flask, render_template,  request, jsonify
from google.oauth2.service_account import Credentials
from utils import drop_top_rows, find_free_room, generate_timetable, get_timeslots, match_timeslot, \
    order_files, preprocess, remove_xlsx_files


app = Flask(__name__)
app.config.from_pyfile('config.py')

timetable_path = app.config['TIMETABLE_PATH']
spreadsheet_url = app.config['URL']
credentials_path = app.config['CREDENTIALS_PATH']
output_folder = timetable_path

classes, labs = preprocess()

@app.route("/get_modification_time", methods=["GET"])
def modification_time():
    '''
        get time when the timetable was last modified or downloaded
    '''
    modification_date = app.config['TIME']

    print(f"Last modification date: {modification_date}")
    
    return jsonify(modification_date)

@app.route("/all-subjects", methods=["GET"])
def all_subjects():
    '''
        this route returns list of all subjects avaialable in the timetable
    '''
    subjects = []
    # Regular expression pattern to match time values like "1:30-2:50"
    time_pattern = r'\d+:\d+-\d+:\d+'

    # Get a list of all files in the folder
    all_files = os.listdir(timetable_path)

    # Filter for files with .xlsx extension
    xlsx_files = [file for file in all_files if file.endswith(".xlsx")]
    # xlsx_files.remove('Welcome.xlsx')

    # # Loop through each Excel file and read its contents
    for excel_file in xlsx_files:
        # Construct the full path to the Excel file within the subfolder
        excel_file_path = os.path.join(timetable_path, excel_file)

        temp = pd.read_excel(excel_file_path)
        # Remove top rows
        temp = drop_top_rows(temp)

        # Remove the first column
        temp = temp.iloc[:, 1:]

        # Remove time values in the format "1:30-2:50" using regular expressions
        temp = temp.applymap(lambda cell: re.sub(time_pattern, '', str(cell)))

        # Flatten and extend subjects
        subjects.extend(temp.values.flatten())

    # Remove empty strings and strip whitespace
    subjects = [subject.strip() for subject in subjects if subject.strip()]

    # Remove duplicates and sort
    subjects = list(set(subjects))

    # remove nan values
    subjects = [subject for subject in subjects if str(subject) != 'nan']
    subjects.sort() # I wonder why sort function doesn't work on this list
    return jsonify(subjects)

# Define the /time-table route to return the timetable
@app.route("/time-table", methods=["POST"])
def get_time_table():
    data = request.get_json()
    subjects = data.get("subjects", [])
    
    print('--------------------------')
    print('Selected Subjects')
    print('--------------------------')
    print(subjects)
    print('type of subjects variable : ', type(subjects))
    print('--------------------------')
    
    df = generate_timetable(subjects, classes, labs)

    print('--------------------------')
    print('Generated Timetable')
    print('--------------------------')
    print(df)
    print('--------------------------')

    # Convert the timetable DataFrame to JSON format
    timetable_json = df.to_json(orient="records")

    print(timetable_json)

    # Return the JSON response
    return timetable_json


@app.route("/update-timetable", methods=["GET"])
def download_sheet_as_excel():
    # Load credentials from the JSON file with the specified scopes
    scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/drive']
    credentials = Credentials.from_service_account_file(credentials_path, scopes=scopes)

    # Authenticate with the Google Sheets API using the updated credentials
    gc = gspread.authorize(credentials)

    try:
        # Open the Google Spreadsheet using its URL
        spreadsheet = gc.open_by_url(spreadsheet_url)
        print(spreadsheet.title)

        # Iterate through each sheet in the spreadsheet
        for worksheet in spreadsheet.worksheets():
            # Get all values from the worksheet
            values = worksheet.get_all_values()

            # Convert the values to a Pandas DataFrame
            df = pd.DataFrame(values[1:], columns=values[0])

            # Check if the folder already exists
            if not os.path.exists(output_folder):
                # Create the folder if it doesn't exist
                os.makedirs(output_folder)

            # Specify the path to save the Excel file
            excel_file_path = f'{output_folder}/{worksheet.title}.xlsx'

            # Save the DataFrame to an Excel file
            df.to_excel(excel_file_path, index=False)

            print(f'Sheet "{worksheet.title}" downloaded and saved as {excel_file_path}')

            # Update modification time
            app.config['TIME'] = datetime.datetime.now()
            
            # remove extra unnecessary files
            remove_xlsx_files()
            classes, labs = preprocess()

        # Return a success response
        return jsonify({"message": "Timetable update successful", "code": 200}), 200

    except gspread.exceptions.APIError as e:
        # Return an error response
        error_message = f"Error accessing Google Sheets API: {e}"
        print(error_message)
        return jsonify({"error": error_message, "code": 500}), 500

@app.route("/submit-rating", methods=["POST"])
def submit_rating():
    """
    Route to receive and store a rating submitted by a client.
    """
    data = request.get_json()
    rating = data.get("rating")

    # Check if the rating is valid (an integer between 1 and 5)
    if not isinstance(rating, int) or rating < 1 or rating > 5:
        return jsonify({"error": "Invalid rating. Please provide an integer between 1 and 5."}), 400

    # Here you can store the rating however you want, e.g., in a database
    # For demonstration purposes, let's just print the rating
    print("Received rating:", rating)
    
    # retrieve variables from file
    total_ratings = app.config['TOTAL_RATINGS']
    sum_rating = app.config['SUM_RATING']
    average = app.config['AVERAGE']
    
    # calculate
    total_ratings += 1
    sum_rating += rating
    average = sum_rating / total_ratings
    
    print(f'Calculations : total_ratings{total_ratings}, sum_rating{sum_rating}, average{average}')
    
    # write to file
    app.config['TOTAL_RATINGS'] = total_ratings
    app.config['SUM_RATING'] = sum_rating
    app.config['AVERAGE'] = average

    return jsonify({"message": "Rating submitted successfully.", "rating": rating}), 200

# Route to fetch the current rating
@app.route("/current-rating", methods=["GET"])
def get_current_rating():
    average = app.config['AVERAGE']
    total_ratings = app.config['TOTAL_RATINGS']
    
    print(average, total_ratings)
    return jsonify({"rating": average, "total votes": total_ratings})

# Route to fetch available files
@app.route("/get_files", methods=["GET"])
def get_files():
    path = 'timetable'
    files = os.listdir(path)
    files = sorted(files, key=order_files)
    files = [file.split('.')[0] for file in files]
    
    return jsonify(files), 200

@app.route("/selected-file", methods=["POST"])
def selected_file():
    """
    Route to receive and store a file selected by the client
    """
    data = request.get_json()
    file = data.get("file") + ".xlsx"
    file_path = os.path.join('timetable', file)
    stype = data.get("selection_type")
    
    print(f'Client Selected File: {file} and asked to search in {stype}')
    
    try:
        timeslots, classes_df, lab_df = get_timeslots(file_path, stype)
        
        # Dump this data in the config.py file for later use
        app.config['CLASS_DF'] = classes_df
        app.config['LAB_DF'] = lab_df
        app.config['SELECTION_TYPE'] = stype
        
        return jsonify(timeslots), 200
    
    except Exception as e:
        print(e)
        return jsonify({"error": str(e)}), 500  # Handle errors gracefully


@app.route("/get-free-room", methods=["POST"])
def get_free_room():
    """
        Route to receive selected timeslot and return free rooms
    """
    print('user called get free room route')
    try:
        data = request.get_json()
        selected_timeslot = data.get("timeslot")
        print('user selected timeslot : ', selected_timeslot)
        stype = app.config['SELECTION_TYPE']
    
        classes_df = app.config['CLASS_DF']
        lab_df = app.config['LAB_DF']
        result = find_free_room(selected_timeslot, stype, classes_df, lab_df)
        print('result : ', result)
        return jsonify(result), 200
    except Exception as e:
        print(e)
        return jsonify({"error": str(e)}), 500
    
    
@app.route('/now-empty', methods=["GET", "POST"])
def now_empty():
    try:
        data = request.get_json()
        # print(data)
        
        current_day = data.get('current-day')
        current_time = data.get('current-time')
        # print(current_time)
        
        # select file of current day
        file = current_day + ".xlsx"
        file_path = os.path.join('timetable', file)
        
        timeslots1, classes_df, _ = get_timeslots(file_path, stype='Room')
        timeslots2, _, lab_df = get_timeslots(file_path, stype='Lab')
        # print(timeslots1, timeslots2)
        
        selected_timeslot1 = match_timeslot(current_time, timeslots1, stype='Room')
        selected_timeslot2 = match_timeslot(current_time, timeslots2, stype='Lab')
        # print(selected_timeslot1, selected_timeslot2)
        
        if selected_timeslot1 is None and selected_timeslot2 is None:
            return jsonify({'message': "University Closed. Try Again later in office timings."}), 201
        
        result1 = find_free_room(selected_timeslot1, 'Room', classes_df, lab_df)
        result2 = find_free_room(selected_timeslot2, 'Lab', classes_df, lab_df)        
        
        return jsonify({'result1': result1, 'result2': result2}), 200
    
    except Exception as e:
        print(e)
        return jsonify({'error': str(e)}), 500
    
@app.route('/subscribe-email', methods=['POST'])
def subscribe_email():
    """
    Route to receive and store subscription email
    """
    try:
        data = request.get_json()
        path = 'Subscribed Emails'
        
        if not os.path.exists(path):
            os.mkdir(path)
            
        # Open the file in append mode to write new emails into it
        with open(f"{path}/emails.txt", "a") as file:
            # Write new emails into the file
            file.write(f"{data.get('email')}\n")        
        
        return jsonify({"success": "email subscribed successfully"}) ,200
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    
@app.route('/post-feedback', methods=['POST'])
def get_feedback():
    try:
        data = request.get_json()

        # Convert data into a DataFrame
        feedback_df = pd.DataFrame([data])

        # Check if the CSV file exists
        csv_file = 'feedback_data.csv'
        if os.path.exists(csv_file):
            # Append the DataFrame to the existing CSV file
            existing_df = pd.read_csv(csv_file)
            updated_df = pd.concat([existing_df, feedback_df], ignore_index=True)
        else:
            updated_df = feedback_df

        # Write the updated DataFrame to CSV
        updated_df.to_csv(csv_file, index=False)

        return jsonify({'message': 'Feedback submitted successfully'}), 200

    except Exception as e:
        return jsonify({'message': f'There was an error in storing the feedack {e}'}), 500
    
    
@app.route('/')
def index():
    return render_template('index.html')

if __name__ == "__main__":
    app.run()
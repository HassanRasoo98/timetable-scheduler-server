import os
import re
import datetime
from fastapi.responses import HTMLResponse
import gspread
import pandas as pd
from fastapi import FastAPI, Request
from fastapi.middleware import Middleware
from fastapi.middleware.cors import CORSMiddleware
from google.oauth2.service_account import Credentials
from utils import *

from dotenv import load_dotenv

middleware = [
    Middleware(
        CORSMiddleware,
        allow_origins=["*"],  # Allows all origins
        allow_credentials=True,
        allow_methods=["*"],  # Allows all methods (GET, POST, PUT, DELETE, etc.)
        allow_headers=["*"]   # Allows all headers
    )
]

app = FastAPI(middleware=middleware)

load_dotenv()
timetable_path = os.getenv('TIMETABLE_PATH')
spreadsheet_url = os.getenv('URL')
credentials_path = os.getenv('CREDENTIALS_PATH')
output_folder = timetable_path

classes, labs = preprocess()

@app.get("/get_modification_time")
def modification_time():
    '''
        get time when the timetable was last modified or downloaded
    '''
    modification_date = os.getenv('TIME')

    print(f"Last modification date: {modification_date}")
    
    return modification_date

@app.get("/all-subjects")
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
    return subjects

# Define the /time-table route to return the timetable
@app.post("/time-table")
def get_time_table(subjects: list):
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


@app.get("/update-timetable")
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
        return "Timetable update successful"

    except gspread.exceptions.APIError as e:
        # Return an error response
        error_message = f"Error accessing Google Sheets API: {e}"
        print(error_message)
        return {"error": error_message, "code": 500}

@app.post("/submit-rating")
def submit_rating(rating: int):
    """
    Route to receive and store a rating submitted by a client.
    """

    # Check if the rating is valid (an integer between 1 and 5)
    if not isinstance(rating, int) or rating < 1 or rating > 5:
        return {"error": "Invalid rating. Please provide an integer between 1 and 5"}

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

    return {"message": "Rating submitted successfully.", "rating": rating}

# Route to fetch the current rating
@app.get("/current-rating")
def get_current_rating():
    average = app.config['AVERAGE']
    total_ratings = app.config['TOTAL_RATINGS']
    
    print(average, total_ratings)
    return {"rating": average, "total votes": total_ratings}

# Route to fetch available files
@app.get("/get_files")
def get_files():
    path = 'timetable'
    files = os.listdir(path)
    files = sorted(files, key=order_files)
    files = [file.split('.')[0] for file in files]
    
    return files

@app.post("/selected-file")
def selected_file(file: str, selection_type: str):
    """
    Route to receive and store a file selected by the client
    """
    file = file + ".xlsx"
    file_path = os.path.join('timetable', file)
    
    print(f'Client Selected File: {file} and asked to search in {selection_type}')
    
    try:
        timeslots, classes_df, lab_df = get_timeslots(file_path, selection_type)
        
        # Dump this data in the config.py file for later use
        os.environ['CLASS_DF'] = classes_df
        os.environ['LAB_DF'] = lab_df
        os.environ['SELECTION_TYPE'] = selection_type
        
        return timeslots
    
    except Exception as e:
        print(e)
        return f"Error : {e}"


@app.post("/get-free-room")
def get_free_room(selected_timeslot):
    """
        Route to receive selected timeslot and return free rooms
    """
    print('user called get free room route')
    try:
        print('user selected timeslot : ', selected_timeslot)
        stype = os.getenv('SELECTION_TYPE')
    
        classes_df = os.getenv('CLASS_DF')
        lab_df = os.getenv('LAB_DF')
        result = find_free_room(selected_timeslot, stype, classes_df, lab_df)
        print('result : ', result)
        return result
    except Exception as e:
        print(e)
        return f"error: {str(e)}"
    
    
@app.post('/now-empty')
def now_empty(data: dict):
    try:
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
        
        return {'result1': result1, 'result2': result2}
    
    except Exception as e:
        print(e)
        return f"Error: {str(e)}"
    
@app.post('/subscribe-email')
def subscribe_email(data: dict):
    """
    Route to receive and store subscription email
    """
    try:
        path = 'Subscribed Emails'
        
        if not os.path.exists(path):
            os.mkdir(path)
            
        # Open the file in append mode to write new emails into it
        with open(f"{path}/emails.txt", "a") as file:
            # Write new emails into the file
            file.write(f"{data.get('email')}\n")        
        
        return "email subscribed successfully"
    except Exception as e:
        return f"Error : {str(e)}"
    
@app.post('/post-feedback')
def get_feedback(data: dict):
    try:
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

        return "Feedback submitted successfully"

    except Exception as e:
        return "Error in Storing the Feedback"
    
    
@app.get('/', response_class=HTMLResponse)
def index(request: Request):
    url = str(request.base_url) + 'docs'

    return f"""
<HTML>

<h1>Timetable Scheduler Backend</h1>
<a href="{url}">Documentation</a>

</HTML>
"""


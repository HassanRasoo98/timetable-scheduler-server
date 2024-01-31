import os
import pandas as pd

# defining variables and base paths
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1feZLJJN4NDjAnqA8J5vHnVGrl9R91-NFGOqAW0gU5h4/edit?usp=sharing'
base_path = os.getcwd()
output_folder = os.path.join(base_path, 'timetable')
# spreadsheet_title = 'TimeTable, FSC, Spring-2024'
credentials_path = os.path.join(base_path, 'timetable-api-412213-75a336ca8f77.json')
timetable = os.path.join(base_path, 'timetable')
# Get a list of all files in the folder
all_files = os.listdir(timetable)

def drop_top_rows(df):    
    '''
        this function will be called in a loop and unnecessary rows of all days dataframes will be dropped
    '''
    df.drop([0, 1, 2], axis=0, inplace=True)
    df.reset_index(drop=True, inplace=True)
    df.columns = df.iloc[0] # make time schedule the column name/header

    schedule = list(df.columns)
    cleaned_schedule = [value for value in schedule if not pd.isna(value)]

    df = df[cleaned_schedule]
    df.drop(0, axis=0, inplace=True)

    df.reset_index(drop=True, inplace=True)
    # df.head()
    
    return df

def separate_labs_and_classes(df, classes, labs):    
    '''
        finds the splitting point of labs and classes
        stores lab df to labs list and class df to class list
    '''
    lab_df = df[df.eq('Lab').any(axis=1)]
    ind = list(lab_df.index)[0]

    classes.append(df[:ind])

    lab_df = df[ind:]
    lab_df.reset_index(drop=True, inplace=True)
    lab_df.columns = lab_df.iloc[0]

    schedule = list(lab_df.columns)
    cleaned_schedule = [value for value in schedule if not pd.isna(value)]

    lab_df = lab_df[cleaned_schedule]
    lab_df.drop(0, axis=0, inplace=True)

    lab_df.reset_index(drop=True, inplace=True)
    labs.append(lab_df)
    # lab_df.head()
    
def order_files(file):
    order = {
        "Monday.xlsx": 0,
        "Tuesday.xlsx": 1,
        "Wednesday.xlsx": 2,
        "Thursday.xlsx": 3,
        "Friday.xlsx": 4,
        "Saturday.xlsx": 5,
        "Sunday.xlsx": 6
    }
    return order.get(file, 0)
    
def preprocess():
    '''
        prepare the classes and labs list
        store dfs of all days in these 2 lists
    '''
    classes=[] # list to store dataframe of all classes of all days of the week
    labs=[] # list to store dataframe of all labs of all days of the week
    
    # Get a list of all files in the folder
    all_files = os.listdir(timetable)

    # Filter for files with .xlsx extension
    xlsx_files = [file for file in all_files if file.endswith(".xlsx")]
    # order the files
    xlsx_files = sorted(xlsx_files, key=order_files)

    # Loop through each Excel file and read its contents
    for excel_file in xlsx_files:
        # Construct the full path to the Excel file within the subfolder
        excel_file_path = os.path.join(timetable, excel_file)
        df = pd.read_excel(excel_file_path)
        
        drop_top_rows(df)
        separate_labs_and_classes(df, classes, labs)
        
    return classes, labs
        
def generate_timetable(subjects, classes, labs):
    '''
        this function takes a list of subjects and
        generates a timetable for it
    '''

    # print('------------------in utils.py--------------------')
    # print('------------------labs list----------------------')    
    # print(classes)
    # print(labs)
    # print('------------------in utils.py--------------------')
    # print('------------------labs list----------------------')
    
    result_day = []
    result_time = []
    result_class = []
    result_subject = []
    # Filter for files with .xlsx extension
    xlsx_files = [file for file in all_files if file.endswith(".xlsx")]
    xlsx_files = sorted(xlsx_files, key=order_files)
    
    for desired_value in subjects:
        # Loop through each Excel file and read its contents
        file_number = -1
        for df in classes:   
            file_number += 1             
            # Check if the value is present in any cell across all columns
            result_rows = df[df.isin([desired_value]).any(axis=1)]
            room = result_rows['Room'].tolist()
            
            if len(result_rows) != 0:
                row_index = result_rows.index.tolist()[0]
            
                # Get the row at the specified index
                row_values = df.iloc[row_index]

                # Find columns containing the specific value in the row
                matching_columns = [col for col, value in row_values.items() if value == desired_value]
                
            # store result if found
            if room != []:
                if len(matching_columns) > 1:
                    for time in matching_columns:
                        result_day.append(xlsx_files[file_number])
                        result_time.append(time)
                        result_class.append(room)
                        result_subject.append(desired_value)                        
                else:
                    # print('match found')
                    result_day.append(xlsx_files[file_number])
                    result_time.append(matching_columns)
                    result_class.append(room)
                    result_subject.append(desired_value)
                
        # if result not found in a class dataframe then search for it in the lab dataframe
        file_number = -1
        for df in labs:   
            file_number += 1   
            # Check if the value is present in any cell across all columns
            result_rows = df[df.isin([desired_value]).any(axis=1)]
            room = result_rows['Lab'].tolist()

            if len(result_rows) != 0:
                row_index = result_rows.index.tolist()[0]
            
                # Get the row at the specified index
                row_values = df.iloc[row_index]

                # Find columns containing the specific value in the row
                matching_columns = [col for col, value in row_values.items() if value == desired_value]
                
            # store result if found
            if room != []:
                if len(matching_columns) > 1:
                    for time in matching_columns:
                        result_day.append(xlsx_files[file_number])
                        result_time.append(time)
                        result_class.append(room)
                        result_subject.append(desired_value)                        
                else:
                    # print('match found')
                    result_day.append(xlsx_files[file_number])
                    result_time.append(matching_columns)
                    result_class.append(room)
                    result_subject.append(desired_value)

    # Create a DataFrame
    data = {'Day': result_day, 'Time': result_time, 'Class': result_class, 'Subject': result_subject}
    
    # print('------------------in utils.py--------------------')
    # print('------------------labs list----------------------')
    # print(result_day)
    # print(result_class)
    # print(result_time)
    # print('------------------in utils.py--------------------')
    # print('------------------labs list----------------------')
    
    result = pd.DataFrame(data)
    
    # print(result)

    def format(value):
        if isinstance(value, list):
            return value[0]
        else:
            return value
    
    def start_time(time):
        return time.split('-')[0]

    def end_time(time):
        return time.split('-')[1]

    result['Time'] = result['Time'].apply(format)
    result['Class'] = result['Class'].apply(format)
    
    # print(result)

    # Extract start and end times from the 'Time' column
    result['Start_Time'] = result['Time'].apply(start_time)
    result['End_Time'] = result['Time'].apply(end_time)
    
    return result
    
def remove_xlsx_files():
    # List all files in the current directory
    all_files = os.listdir(output_folder)

    # Filter files with .xlsx extension
    xlsx_files = [file for file in all_files if file.endswith(".xlsx")]
    allowed_files = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
    allowed_files = [os.path.join(output_folder, f"{filename}.xlsx") for filename in allowed_files]    
    
    # Remove each .xlsx file
    for xlsx_file in xlsx_files:
        file_path = os.path.join(output_folder, xlsx_file)
        
        if file_path not in allowed_files:
            os.remove(file_path)
            print(f"Removed file: {file_path}")
            
def get_timeslots(file, stype='Room'):
    df = pd.read_excel(file)
    df = drop_top_rows(df)
    
    lab_df = df[df.eq('Lab').any(axis=1)]
    ind = list(lab_df.index)[0]

    classes_df = df[:ind]
    lab_df = df[ind:]

    if stype=='Room':    
        timeslots = classes_df.columns.tolist()
        
        
    else:
        lab_df.reset_index(drop=True, inplace=True)
        lab_df.columns = lab_df.iloc[0]

        schedule = list(lab_df.columns)
        cleaned_schedule = [value for value in schedule if not pd.isna(value)]

        lab_df = lab_df[cleaned_schedule]
        lab_df.drop(0, axis=0, inplace=True)

        lab_df.reset_index(drop=True, inplace=True)

        timeslots = lab_df.columns.tolist()
        
    timeslots.remove(stype)
    return timeslots, classes_df, lab_df

def find_free_room(selected_timeslot, stype, classes_df, lab_df):
    result = None
    if stype == 'Room':
        # filter null values and cancelled classes
        filtered_df = classes_df[(classes_df[selected_timeslot].isnull()) | (classes_df[selected_timeslot].str.contains('cancel'))]
        result = filtered_df['Room'].tolist()

    else:
        # filter null values and cancelled labs
        filtered_df = lab_df[(lab_df[selected_timeslot].isnull()) | (lab_df[selected_timeslot].str.contains('cancel'))]
        result = filtered_df['Lab'].tolist()
        
    return result

from datetime import datetime

def update_time_ranges(time_ranges):
    updated_time_ranges = []

    for time_range in time_ranges:
        # '05:20 - 08:05 (inc. 10 min. break)  '
        if '(' in time_range:
            time_range = time_range.split('(')[0].strip()
            
        time_range = time_range.strip()
        start, end = time_range.split('-')
        start = start.strip()
        end = end.strip()
        start_hour = int(start.split(':')[0])
        # start_minute = int(start.split(':')[1])
        end_hour = int(end.split(':')[0])
        # end_minute = int(start.split(':')[1])
        
        if start_hour >= 8 and start_hour < 12:
            updated_start = f'{start}AM'
        else:
            updated_start = f'{start}PM'

        if end_hour >= 8 and end_hour < 12:
            updated_end = f'{end}AM'
        else:
            updated_end = f'{end}PM'

        updated_time_ranges.append(f'{updated_start}-{updated_end}')

    # print(updated_time_ranges)
    return updated_time_ranges


# Function to convert time string to datetime object
def time_str_to_datetime(time_str):
    return datetime.strptime(time_str, '%I:%M%p')

# Function to compare times considering only hours and minutes
def compare_times(time1, time2):
    return time1.replace(second=0, microsecond=0) == time2.replace(second=0, microsecond=0)

def format_time(time):
    res = ''
    for t in time:
        if t.isnumeric() or t==':' or t=='-':
            res += t
            
    return res

def match_timeslot(given_time, time_intervals, stype):
    time_intervals = update_time_ranges(time_intervals)
        
    print(time_intervals)
    given_time = time_str_to_datetime(given_time)
    # Iterate through time intervals and check if the given time falls within any interval
    interval_found = False
    for interval in time_intervals:
        start_str, end_str = interval.split('-')
        start_time = time_str_to_datetime(start_str)
        end_time = time_str_to_datetime(end_str)
        
        # print(start_time, given_time, end_time)
        
        if start_time <= given_time <= end_time:
            interval_found = True
            print(f"The given time {given_time.strftime('%Y-%m-%d %I:%M:%S %p')} falls within the interval: {interval}")
            break
        elif compare_times(given_time, end_time):
            interval_found = True
            print(f"The given time {given_time.strftime('%Y-%m-%d %I:%M:%S %p')} falls within the interval: {interval}")
            break

    if not interval_found:
        print("The given time does not fall within any interval.")
        return None
        
    return format_time(interval)


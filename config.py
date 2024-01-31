# config.py
import os
from datetime import datetime
    
TOTAL_RATINGS = 0
SUM_RATING = 0
AVERAGE = 0

CLASS_DF = None
LAB_DF = None
SELECTION_TYPE = None

CREDENTIALS_PATH = 'timetable-api-412213-75a336ca8f77.json'
TIMETABLE_PATH = os.path.join(os.getcwd(), 'timetable')
URL = 'https://docs.google.com/spreadsheets/d/1feZLJJN4NDjAnqA8J5vHnVGrl9R91-NFGOqAW0gU5h4/edit?usp=sharing'
TIME = datetime.now()
SECRET_KEY = 'my_secret_key'

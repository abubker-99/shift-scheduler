# Author: Abubaker Haider
# Description:
# This script generates a weekly shift schedule for customer support agents
# considering gender, language, leave status, and shift eligibility.
# It assigns random shifts, ensures 2 consecutive off-days, and exports to Excel.

import random
import pandas as pd
from datetime import timedelta, datetime

def schedule(agents, start_date_str):
    """
    Generate a 7-day schedule for customer support agents based on rules.

    Parameters:
    - agents: List of dictionaries containing agent data
    - start_date_str: Start date string in format "MM.DD.YYYY"
    """
    
    # Step 1: Create a list of 7 consecutive dates
    start_date = datetime.strptime(start_date_str, "%m.%d.%Y")
    excel_dates = [(start_date + timedelta(days=i)) for i in range(7)]

    # Step 2: Define available shift timings
    schedules = [
        "6:00-15:00", "7:00-16:00", "8:00-17:00", "9:00-18:00",
        "10:00-19:00", "12:00-21:00", "13:00-22:00", "14:00-23:00"
    ]

    # Step 3: Assign shifts and working dates based on gender/language/type
    for agent in agents:
        if agent["leave"]:
            continue  # Skip agents who are on leave

        # Female Bilingual agents
        if agent["gender"] == "female" and agent["language"] == "Both":
            agent["schedule"] = random.choices(schedules[1:2], k=5)
            agent["date"] = excel_dates[2:7] if "ty" in agent else excel_dates[0:5]

        # Female Arabic-speaking agents
        elif agent["gender"] == "female" and agent["language"] == "Ar":
            agent["schedule"] = random.choices(schedules[1:3], k=5)
            agent["date"] = excel_dates[2:7] if "ty" in agent else excel_dates[0:5]

        # Female English-speaking agents
        elif agent["gender"] == "female" and agent["language"] == "En":
            agent["schedule"] = random.choices(schedules[0:6], k=5)
            agent["date"] = excel_dates[0:5]

        # Male Bilingual agents
        elif agent["gender"] == "male" and agent["language"] == "Both":
            agent["schedule"] = random.choices(schedules[6:], k=5)
            agent["date"] = excel_dates[2:7] if "ty" in agent else excel_dates[0:5]

        # Male English-speaking agents
        elif agent["gender"] == "male" and agent["language"] == "En":
            agent["schedule"] = random.choices(schedules[6:], k=5)
            agent["date"] = excel_dates[2:7]

    # Step 4: Compile schedule records for Excel export
    rows = []
    for agent in agents:
        if agent["leave"]:
            continue
        for i in range(5):
            rows.append({
                "Agent name": agent["name"],
                "Date": agent["date"][i],
                "schedule": agent["schedule"][i]
            })

    # Step 5: Create a DataFrame and format Date column
    df = pd.DataFrame(rows)
    df['Date'] = pd.to_datetime(df['Date'])

    # Step 6: Export the schedule to Excel with proper date formatting
    with pd.ExcelWriter("x.xlsx", engine="xlsxwriter", datetime_format='mm-dd-yyyy') as writer:
        df.to_excel(writer, index=False)

# Example usage:
# agents = [
#     {"name": "Ali", "gender": "male", "language": "En", "leave": False},
#     {"name": "Sara", "gender": "female", "language": "Both", "leave": False, "ty": "x"},
# ]
# schedule(agents, "06.17.2025")
    
    
    
# Example usage


agents = [
  {"name": "Ahmed", "gender": "male", "leave": True, "language": "Both", 'ty':"x"},
  {"name": "Abdallah", "gender": "male","leave": True, "language": "Both", 'ty':"x"},
  {"name": "Abubker", "gender": "male", "leave": True, "language": "Both", 'ty':"x"  },  
  {"name": "Khalid", "gender": "male", "leave": False, "language": "Both" },
  {"name": "Basheer", "gender": "male", "leave": False, "language": "Both", 'ty':"x"},
  {"name": "Fatima", "gender": "female", "leave": False, "language": "Both"},  
  {"name": "Shaimaa-H", "gender": "female", "leave": False, "language": "Ar"},
  {"name": "Shaimaa-R", "gender": "female", "leave": False, "language": "Both", 'ty':"x"},
  {"name": "Salsabeel", "gender": "female", "leave": False, "language": "Both", 'ty':"x"},  
  {"name": "Reem", "gender": "female", "leave": False ,"language": "Ar", 'ty':"x"},
  {"name": "Abeir", "gender": "female", "leave": False, "language": "Both"},
  {"name": "Hiba", "gender": "female", "leave": True , "language": "Both"},  
  {"name": "Yassin", "gender": "male", "leave": True, "language": "Both"},
  {"name": "Jalal", "gender": "male", "leave": False, "language": "Both"},
  {"name": "Sohaila", "gender": "female", "leave": False, "language": "Both"},  
  {"name": "Meriem", "gender": "female", "leave": False, "language": "Both", 'ty':"x"},
  {"name": "John", "gender": "male", "leave": False, "language": "En", 'ty':"x"},
  {"name": "Raees", "gender": "male", "leave": False, "language": "En", 'ty':"x"},  
  {"name": "Trixie", "gender": "female", "leave": False, "language": "En"},  
]

schedule(agents, "6.15.2025")
# Author: Abubaker Haider
# Description:
# This script generates a weekly shift schedule for customer support agents
# considering gender, language, leave status, and shift eligibility.
# It assigns random shifts, ensures 2 consecutive off-days, and exports to Excel.

import json
import pandas as pd
from datetime import timedelta, datetime


with open("custom_schedules.json", "r", encoding="utf-8") as f:
   custom_schedules = json.load(f)

def schedule(agents, start_date_str):
    
    # Step 1: Create a list of 7 consecutive dates
    start_date = datetime.strptime(start_date_str, "%m.%d.%Y")
    excel_dates = [(start_date + timedelta(days=i)) for i in range(7)]

    # Step 2: Define available shift timings
    schedules = [
        "6:00-15:00", "7:00-16:00", "8:00-17:00", "9:00-18:00",
        "10:00-19:00","11:00-20:00", "12:00-21:00", "13:00-22:00", "14:00-23:00", "22:00-07:00"
    ]

    friday_sat_off = excel_dates[0:5]  # Sunday to Thursday
    sun_mon_off = excel_dates[2:7]  # Tuesday to Saturday
    tue_wed_off = excel_dates[0:2] + excel_dates[4:]
 

    # Step 3: Assign shifts and working dates based on gender/language/type
    for agent in agents:
        if agent["leave"]:
            continue  # Skip agents who are on leave
        # Female Bilingual agents
        if agent['name'] in custom_schedules:
            agent['schedule'] = custom_schedules[agent["name"]]
            if agent["week_start"] == "sun":
                agent["date"] = excel_dates[0:len(agent["schedule"])]
            else:
                if agent["name"] == 'Meriem':
                    agent["date"] =  tue_wed_off 
                else:
                    agent["date"] = excel_dates[2:2+len(agent["schedule"])]

        # Female Arabic-speaking agents

       
        elif agent["gender"] == "female" and agent["language"] == "Ar":
            if agent['week_start'] == "sun":
               agent["date"] = friday_sat_off
       
               agent["schedule"] = [schedules[1]] * 5

               if agent['name'] == "Fatima":
                  agent["schedule"] = [schedules[0]] * 5   
            else :
                agent["schedule"] = [schedules[0], schedules[0],schedules[1],schedules[2],schedules[3]]
                agent["date"] = excel_dates[2:2+len(agent["schedule"])]
            
           

        # Female English-speaking agents
        elif agent["gender"] == "female" and agent["language"] == "En":
            agent["schedule"] = [schedules[0]] * 5 
            agent["date"] = friday_sat_off
        

        # Male Bilingual agents
        elif agent["gender"] == "male" and agent["language"] == "Both":
            if agent['week_start'] == "tue":
             agent["date"] = sun_mon_off 
             agent["schedule"] = [schedules[8],schedules[8],schedules[8],schedules[8],schedules[8]]
            elif agent["night"] == True:
                agent["date"] = friday_sat_off
                agent["schedule"]= [schedules[9]] * 5

            else:
                agent["date"] = friday_sat_off  
                agent["schedule"] = [schedules[8],schedules[8],schedules[7],schedules[7],schedules[7]]
            
           

        # Male English-speaking agents
        elif agent["gender"] == "male" and agent["language"] == "En":
            if agent["night"] == True:
                agent["date"] = friday_sat_off
                agent["schedule"] = [schedules[9]] * 5
            else:
             agent["date"] = sun_mon_off
             agent["schedule"] = [schedules[8]] * 5
            
        
    
       
    # Step 4: Compile schedule records for Excel export
    rows = []
    for agent in agents:
        if agent["leave"]:
            continue
        for i in range(len(agent["schedule"])):
            rows.append({
                "Agent name": agent["name"],
                "Date": agent["date"][i],
                "schedule": agent["schedule"][i]
            })
        
    
    # Step 5: Create a DataFrame and format Date column
    df = pd.DataFrame(rows)

    # Step 6: Export the schedule to Excel with proper date formatting
    
    df.to_excel('x.xlsx', index=False)
# Example usage
#     Json DATA
#     {
#         "name": "agent",
#         "gender": "female",
#         "leave": false,
#         "language": "Both",
#         "week_start": "sun"
#     },
#     {
#         "name": "agent",
#         "gender": "female",
#         "leave": false,
#         "language": "Both",
#         "week_start": "tue"
#     },
#     {
#         "name": "agent",
#         "gender": "male",
#         "leave": false,
#         "language": "En",
#         "week_start": "sun",
#         "night": false
#     },
#     {
#         "name": "agent",
#         "gender": "male",
#         "leave": true,
#         "language": "En",
#         "week_start": "sun",
#         "night": true
#     },
#     {
#         "name": "agent",
#         "gender": "female",
#         "leave": false,
#         "language": "En",
#         "week_start": "sun"
#     }
# ]

schedule(agents, "6.15.2025")

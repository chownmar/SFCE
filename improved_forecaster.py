#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon May 17 15:27:42 2021

@author: martinchown
"""


import matplotlib.pyplot as plt
import pandas as pd
import requests
import json
from datetime import datetime, timedelta
import xlwt
from xlwt import Workbook

import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook


## This function accesses the climaCell API and retrieves the updated weather
## prediction data the forecast will run on. It returns the data in json format
def get_weather_data(weather_type):
    url = "https://api.climacell.co/v3/weather/forecast/hourly"
    
    querystring = {"fields":weather_type,"unit_system":"si","lat":"44.77","lon":"-85.58"}
    
    headers = {
    'apikey':"CVrWUUalbBEiK0PinVEqHmVqksAp8omv"
    }
    
    response = requests.request("GET", url, headers=headers, params=querystring)
    
    rawdata = response.text
    
    sun_json = json.loads(rawdata)
    
    return sun_json


## This function just retrieves the solar insolation forecast for a specified
## amount of hours in the future
def get_solar_radiation_from_hour(sun_json_data, hours_in_future):
    value = sun_json_data[hours_in_future]['surface_shortwave_radiation']['value']
    return value


def get_solar_radiation_from_time(sun_json_data, day_of_month, hour):
    
    ## This function takes in a json data table with radiation values,
    ## a day of a month (like 17), and a time (like 16), and returns
    ## the solar radiation value on that day and hour
    
    ## First convert day_of_month and hour to strings if they are ints
    day_of_month = str(day_of_month)
    hour = str(hour)
    
    ## This should loop through json data to find the matching day and hour
    for i in range(109):
        date_string = sun_json_data[i]['observation_time']['value']
        
        day = date_string[8] + date_string[9]
        ## The below if statement accounts for single digit dates like "01"
        if date_string[8] == '0':
            day = date_string[9]
            
        time_hour = date_string[11] + date_string[12]
        if date_string[11] == '0':
            time_hour = date_string[12]
        if day == day_of_month and time_hour == hour:
            return sun_json_data[i]['surface_shortwave_radiation']['value']
        
    
    print("Oh shucks this did not work")
    
    
def get_temp_from_hour(temp_json_data, hours_in_future):
    value = temp_json_data[hours_in_future]['temp']['value']
    return value


def get_temp_from_time(temp_json_data, day_of_month, hour):
    
    ## This function takes in a json data table with temp values,
    ## a day of a month (like 17), and a time (like 16), and returns
    ## the temp value on that day and hour
    
    ## First convert day_of_month and hour to strings if they are ints
    day_of_month = str(day_of_month)
    hour = str(hour)
    
    
    ## This should loop through json data to find the matching day and hour
    for i in range(109):
        date_string = temp_json_data[i]['observation_time']['value']
        
        day = date_string[8] + date_string[9]
        ## The below if statement accounts for single digit dates like "01"
        if date_string[8] == '0':
            day = date_string[9]
            
        time_hour = date_string[11] + date_string[12]
        if date_string[11] == '0':
            time_hour = date_string[12]
        if day == day_of_month and time_hour == hour:
            return temp_json_data[i]['temp']['value']
        
    
    print("Oh shucks this did not work")
    


## This function estimates the power produced by the array based on the temp 
## and solar insolation. The surface area of a 10kW array is 60 m^2
def get_power(radiation, temp):
    efficiency = 0.22
    surface_area = 60 ##m^2
    
    power = efficiency * surface_area * radiation * (1 - (0.05 * (temp - 25)))
    return power
  

## This function introduces the forecaster and retrieves and returns the data
def prepare_to_run_forecast():
    print("Hello! Welcome to Central High School Solar Forecaster")
   
    print("This sim is based on Solar Irradiance and Temperature data")
    
    print("To begin, we will load in weather data from the ClimaCell API")
    
    solar_data = get_weather_data('surface_shortwave_radiation')
    
    temp_data = get_weather_data('temp')
    
    print("Data successfully loaded!")
    
    today_dt = datetime.today()
    today = str(today_dt)
    
    month = today[5] + today[6]
    day = today[8] + today[9]
    print("Today, the month is ", month, "and the day is ", day)
    td = timedelta(days = 1)
    tmrw_dt = today_dt + td
    tmrw = str(tmrw_dt)
    tmrw_month = tmrw[5] + tmrw[6]
    tmrw_day = tmrw[8] + tmrw[9]
    print("The forecast will start tomorrow, month ", tmrw_month, " and day ", tmrw_day)
    
    return temp_data, solar_data

## This is the main function that uses the data and call the functions that
## calculate the expected energy each day based on the data
def run_forecast(sun_data, tmp_data):
    today = datetime.today()
    td = timedelta(days = 1)
    day_one = today + td
    day_two = day_one + td
    day_three = day_two + td
    day_four = day_three + td
    day_five = day_four + td
    day_six = day_five + td
    
    day_one_str = str(day_one)
    day_two_str = str(day_two)
    day_three_str = str(day_three)
    day_four_str = str(day_four)
    day_five_str = str(day_five)
    day_six_str = str(day_six)

    
    power_day_one = calc_energy_in_day(tmp_data, sun_data, day_one_str[8] + day_one_str[9])
    power_day_two = calc_energy_in_day(tmp_data, sun_data, day_two_str[8] + day_two_str[9])
    power_day_three = calc_energy_in_day(tmp_data, sun_data, day_three_str[8] + day_three_str[9])
    power_day_four = calc_energy_in_day(tmp_data, sun_data, day_four_str[8] + day_four_str[9])
    power_day_five = calc_energy_in_day(tmp_data, sun_data, day_five_str[8] + day_five_str[9])
    power_day_six = calc_energy_in_day(tmp_data, sun_data, day_six_str[8] + day_six_str[9])
    
    print("On day one, the array will produce", power_day_one, " kWh")
    print("On day two, the array will produce", power_day_two, " kWh")
    print("On day three, the array will produce", power_day_three, " kWh")
    print("On day four, the array will produce", power_day_four, " kWh")
    print("On day five, the array will produce", power_day_five, " kWh")
    print("On day six, the array will produce", power_day_six, " kWh")
    
    return power_day_one, power_day_two, power_day_three, power_day_four, power_day_five, power_day_six

    
    
    
## This function loops through the forecast data and if the day in the data
## matches the day it wants to calculate, then it will sume the power produced on that day
def calc_energy_in_day(temp_data, sun_data, frcst_day):
    energy_total = 0
    frcst_day = str(frcst_day)
    
    for i in range(109):
        date_string = temp_data[i]['observation_time']['value']
        
        day = date_string[8] + date_string[9]
      
        if day == frcst_day:
            sun = get_solar_radiation_from_hour(sun_data, i)
            temp = get_temp_from_hour(temp_data, i)
            energy_total = energy_total + get_power(sun, temp)*0.001
            
    return energy_total


## The below functions are for writing to excel and this one returns a cell to edit
def get_cell_to_edit(row_int, days_out):
    letter = 'A'
    if days_out == 1:
        letter = 'C'
        row_int = row_int + 1
    if days_out == 2:
        letter = 'D'
        row_int = row_int + 2
    if days_out == 3:
        letter = 'E'
        row_int = row_int + 3
    if days_out == 4:
        letter = 'F'
        row_int = row_int + 4
    if days_out == 5:
        letter = 'G'
        row_int = row_int + 5
    if days_out == 6:
        letter = 'H'
        row_int = row_int + 6
    cell = letter + str(row_int + 2)
    
    return cell

## This is useful because the difference in days from the current date today
## and the origin date (5/17/21) that is the first row in the excel file 
## determines how many rows down the program needs to write
def find_day_diff(origin_date, today_date):
    diff = today_date - origin_date
    diff_str = str(diff)
    num_str = diff_str[0]
    num = int(num_str)
    return num
    

## This function takes in the forecasted power for the next 6 days and then
## calls the functions that will add the forecast to the correct spot in the excel file. 
## FOR EXAMPLE: on June 2, the "p2" power refers to the forecasted power for June 4 so 
## the program will put the p2 power in the June 4 row and the 2 day out column
def write_to_excel(p1, p2, p3, p4, p5, p6):
    wb = load_workbook(filename = '/Users/martinchown/Downloads/Forecast_Data.xlsx')
    ws1 = wb.active
    ws1.title = "Data"
    today = datetime.today()
    origin = datetime(2021, 5, 17, datetime.today().hour, 0, 0,)
    
    row_num = find_day_diff(origin, today)

    ws1[get_cell_to_edit(row_num, 0)] = datetime.today()
    ws1[get_cell_to_edit(row_num, 1)] = p1
    ws1[get_cell_to_edit(row_num, 2)] = p2
    ws1[get_cell_to_edit(row_num, 3)] = p3
    ws1[get_cell_to_edit(row_num, 4)] = p4
    ws1[get_cell_to_edit(row_num, 5)] = p5
    ws1[get_cell_to_edit(row_num, 6)] = p6
    
    wb.save('/Users/martinchown/Downloads/Forecast_Data.xlsx')
    
    
    

    

temp_data, solar_data = prepare_to_run_forecast()
power1, power2, power3, power4, power5, power6 = run_forecast(solar_data, temp_data)
write_to_excel(power1, power2, power3, power4, power5, power6)



   


    
    


    


    

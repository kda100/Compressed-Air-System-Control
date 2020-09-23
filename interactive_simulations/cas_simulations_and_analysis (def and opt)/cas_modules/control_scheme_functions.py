# -*- coding: utf-8 -*-
"""
Created on Mon Jul 27 22:47:35 2020
Functions used to produce results for compressed air system simulations in the following scripts:
    * 'defined_simulation.py'
    * 'optimised_simulation.py'
    * 'defined_optimised_simulations_and_analysis.py'
@author: kda_1
"""

from sys import exit
import cas_modules.compressor_models2 as cm2
import openpyxl

def get_compressor_capacity(input_statement):
    '''
    Takes in input_statement argument and returns maximum air capacity for compressor object
    '''
    while True:
        try:
            compressor_capacity = input(input_statement) #saves input for user
            compressor_capacity = int(compressor_capacity) #trys to convert number to integer
            if compressor_capacity == 0: #breaks from while loop if user types 0
                break
        except:
            try:
                compressor_capacity = float(compressor_capacity) #trys to convert n to float if it is not an integer
            except:
                print("\nPlease enter a number") #if an error occurs
                continue
            else:
                break
        else:
            break
    if compressor_capacity == 0:
        exit('Program has ended') #exits from entire program, if user types 0
    return compressor_capacity
            
    
def get_compressor_power(input_statement):
    '''
    Takes in input_statement argument and returns maximum power capacity for compressor object
    '''
    #while True:
    while True:
        try:
            compressor_power = input(input_statement) #saves input for user
            compressor_power = int(compressor_power) #trys to convert number to integer
            if compressor_power == 0: #breaks from while loop if user types 0
                break
        except:
            try:
                compressor_power = float(compressor_power) #trys to convert n to float if it is not an integer
            except:
                print("\nPlease enter a number") #if an error occurs
                continue
            else:
                break
        else:
            break
    if compressor_power == 0:
        exit('Program has ended') #exits from entire program, if user types 0
    return compressor_power   

def GetCompressor1(compressor_num, maximum_power, maximum_flow):
    '''
    Takes in arguments: compressor_num, maximum_power and minimum_flow to return compressor object type, from user defined input.
    '''
    COMPRESSOR_TYPE_DICT = {'OnOff': cm2.OnOff(maximum_power, maximum_flow),
                            'LoadUnload':cm2.LoadUnload(maximum_power, maximum_flow),
                            'InletModulation':cm2.InletModulation(maximum_power, maximum_flow),
                            'VariableSpeed':cm2.VariableSpeed(maximum_power, maximum_flow)} #dictionary with keys as compressor type strings and the compressor objects as values.
    while True:
        try:
            print('\nCompressor types available:\n\tOnOff\n\tLoadUnLoad\n\tInletModulation\n\tVariableSpeed')
            C = input(f'Choose a compressor type for C{compressor_num} which has a maximum power of {maximum_power}kW and a maximum flow of {maximum_flow}m3/hr: ') #User inputs compressor type and string is returned
            if C == '0': #breaks from while loop if user types 0
                break
            C = COMPRESSOR_TYPE_DICT[C] #assigns compressor object from compressor_type_dict, user must input string thats in compressor_type_dict
        except:
            print("\nInvalid compressor type")
            continue
        else:
            break
    if C == '0':
        exit('Program has ended') #exits from entire program, if user types 0
    return C
        
def GetCompressor2(compressor, maximum_power, maximum_flow):
    '''
    Takes in arguments: compressor_num, maximum_power and minimum_flow to return compressor object type, from user defined input.
    '''
    COMPRESSOR_TYPE_DICT = {'OnOff': cm2.OnOff(maximum_power, maximum_flow),
                            'LoadUnload':cm2.LoadUnload(maximum_power, maximum_flow),
                            'InletModulation':cm2.InletModulation(maximum_power, maximum_flow),
                            'VariableSpeed':cm2.VariableSpeed(maximum_power, maximum_flow)} #dictionary with keys as compressor type strings and the compressor objects as values.
    
    return COMPRESSOR_TYPE_DICT[compressor]
    
    
def get_trim(compressor_dict):
    '''
    Takes in a compressor_dict dictionary to return a compressor object to act as trim compressor, from user input.
    '''
    while True:
        try:
            trim = input("Choose the trim compressor: ") #User input compressor name and string is returned
            if trim == '0': #breaks from while loop if user types 0
                break
            trim = compressor_dict[trim] #string returned must be an compressor name in the compressor_dict to give the compressor object.
        except:
            length = len(compressor_dict)
            print(f'\nChoose compressor C1-C{length}')
            continue
        else:
            break
    if trim == '0':
        exit('Program has ended') #exits from entire program, if user types 0 
    return trim

def get_control_type(input_statement):
    '''
    Takes in a input statement argument and returns control types 'opt' or 'def' from user input
    '''
    CONTROL_TYPES = ['opt','def']
    while True:
        try:
            control_scheme_type = input(input_statement) #User inputs control scheme type and string is returned
            if control_scheme_type == '0': #breaks from while loop if user types 0
                break
            elif control_scheme_type not in CONTROL_TYPES:
                continue #user input not one of control types then ask for user input again
        except:
            print("\nPlease type 'opt' for optimisation or 'def' for defined")
            continue
        else:
            break
    if control_scheme_type == '0':
        exit('Program has ended') #exits from entire program, if user types 0 
    return control_scheme_type

def get_setpoints(input_statement):
    '''
    Take in input_statement argument and returns number of setpoints to be used in defined control sequence
    '''
    while True:
        try:
            setpoints = input(input_statement) #saves input for user
            setpoints = int(setpoints) #trys to convert number to integer
            if setpoints == 0: #breaks from while loop if user types 0
                break
        except:
            try:
                setpoints = float(setpoints) #trys to convert n to float if it is not an integer
            except:
                print("\nPlease enter a number") #if an error occurs
                continue
            else:
                break
        else:
            break
    if setpoints == 0:
        exit('Program has ended') #exits from entire program, if user types 0
    return setpoints

def get_power_array(compressor, trim):
    '''
    Takes in a compressor object and the trim compressor to returns either the discrete_power_output function or a list of the maximum and minmum power outputs attributes of the compressor object.
    '''
    if trim == compressor:
        return compressor.discrete_power_output() #trim is the same as the compressor object
    else:
        return [compressor.maximum_power, compressor.minimum_power] #trim is not the same as the compressor object
    
def get_air_array(compressor, trim):
    '''
    Takes in a compressor object and the trim compressor to returns either the discrete_air_output function or a list of the maximum and minmum power outputs attributes of the compressor object.
    '''
    if trim == compressor:
        return compressor.discrete_air_output() #trim is the same as the compressor object
    else:
        return [compressor.maximum_flow, compressor.minimum_flow] #trim is not the same as the compressor object

def get_workbook(input_statement):
    '''
    Takes in a input statement argument and returns an excel workbook loaded with openpyxl
    '''
    while True:
        try:
            workbook = input(input_statement) #user inputs control scheme type and string is returned
            if workbook == '0': #breaks from while loop if user types 0
                break
            print('\nUploading workbook...')
            workbook = openpyxl.load_workbook(workbook +'.xlsx') #workbook loaded with openpyxl
            print('\nWorkbook uploaded')
        except PermissionError: #workbook not able to be opened
            print('\nEnsure the document is closed\n')
            continue
        except:
            print('\nInput error, input: C:\path\excelworkbookname')
            continue
        else:
            break    
    if workbook == '0':
        exit('Program has ended') #exits from entire program, if user types 0
    return workbook
        
def get_sheet(input_statement, workbook):
    '''
    Takes in a input statement argument and returns an loaded excel workbook on user inputted sheetname
    '''
    while True:  
        try:
            sheet = input(input_statement) #user inputs sheet name and string is returned 
            if sheet == '0': #breaks from while loop if user types 0
                break
            print('\nUploading sheet...')
            sheet = workbook[sheet] #workbook loaded on given sheet name
            print('\nSheet uploaded.')
        except:
             print('\nInput error, sheet not contained in workbook')
             continue
        else:
            break
    if sheet == '0':
        exit('Program has ended') #exits from entire program, if user types 0
    return sheet

def get_active_compressors(boundaries, air_demand, setpoints, contri_compressors, trim):
    '''
    Takes in boundaries (list of tuples), air_demand (integer), setpoints (list), contri_compressor(list of lists of compressors) and trim (compressor object), then returns a list of active compressors.
    '''
    for minimum, maximum in boundaries: #tuple unpacking of minimum and maximum setpoints in boundaries.
        if minimum<=air_demand<maximum: #checking if air demand is between the minimum and maximum values.
            active_compressors = contri_compressors[boundaries.index((minimum, maximum))] #grabs list of compressors from contri_compressors based on index position of (minimum, maximum) tuple in boundaires.
            active_compressors[-1] = trim #changes the last compressor object in active_compressors to trim compressor.
    return active_compressors

def update_power_output(compressor, active_compressors, power_output_list, air_output_list, air_demand):
    '''
    Takes in compressor object, active compressors list, power_output_list and air_demand values and appends the compressor.power_output(air_demand, active_compressors) of the compressor object to the power_output_list.
    '''
    if type(compressor) == cm2.OnOff or type(compressor) == cm2.VariableSpeed:
        power_output_list.append(compressor.power_output(air_demand, active_compressors)) #if compressor type is variablespeed or onoff, append power output as normal
    else:
        if len(air_output_list) == 0: #at the start of simulation 
            if compressor in active_compressors: #compressor part of active compressors
                power_output_list.append(compressor.power_output(air_demand, active_compressors)) #turn on compressor
            else:
                power_output_list.append(0) #otherwise levae compressor off
        
        elif len(air_output_list) > 0: #after simulation has started
            if compressor in active_compressors: #compressor part of active compressors 
                power_output_list.append(compressor.power_output(air_demand, active_compressors)) #turn compressor on or leave it on
            elif air_output_list[-1]>0: #if compressor produced air previously then leave compressor on 
                power_output_list.append(compressor.power_output(air_demand, active_compressors))
            elif air_output_list[-1]==0: #if compressor did not produce air previously then turnoff
                power_output_list.append(0)
        
def update_air_output(compressor, active_compressors, air_output_list, air_demand):
    '''
    Takes in compressor object, active compressors list, air_output_list and air_demand values and appends the compressor.pair_output(air_demand, active_compressors) of the compressor object to the power_air_list.
    '''
    air_output_list.append(compressor.air_output(air_demand, active_compressors))    
 
def get_coordinate(input_statement, sheet):
    '''
    Takes in an input statement and an openpyxl excel workbook sheet object to return a coordinate in the sheet
    '''
    while True:
        try:
            coordinate = input(input_statement) #user inputs coordinate and string is returned 
            if coordinate == '0': #breaks from while loop if user types 0
                break
            coordinate = sheet[coordinate] #checks if coordinate is a valid coordinate in excel sheet object
        except:
            print("\nInvalid cell, try again")
            continue
        else:
            break
    if coordinate == '0':
        exit('Program has ended') #exits from entire program, if user types 0
    return coordinate

def save_and_close_results(input_statement, workbook):
    '''
    Takes in an input statement and an openpyxl excel workbook object to return a coordinate in the sheet, then 
    takes in a user input workbook name, then saves excel workbook as a user input and closes it.
    '''
    while True:  
        try:
            workbook_name = input(input_statement) #user inputs coordinate and string is returned 
            if workbook_name == '0': #breaks from while loop if user types 0
                break
            if len(workbook_name.split('.')) > 1: #checks if workbook name is valid
                raise NameError
            workbook.save(workbook_name +'.xlsx') #saves excel workbook with results
        except:
            print("\nEnter a filename with the format 'excelworkbook'")
        else:
            print('\nWorkbook saved')
            workbook.close() #closes workbook
            break
    if workbook_name == '0':
        exit('Program has ended') #exits from entire program, if user types 0
    
    
    
    
    
    
    
# -*- coding: utf-8 -*-
"""
Created on Mon Aug  3 11:28:15 2020
Functions used to produce figures for analysis in the following scripts:
    * 'generate_simulation_figures.py'
    * 'generate_multi_simulation_figures.py'
    * 'defined_optimised_simulations_and_analysis.py'
    
@author: kda_1
"""
from sys import exit

def get_num_compressors(input_statement):
    '''
    Takes in input statement arguemnt and returns integer number of compressors used in simulation.
    '''
    while True:
        try:
            num_of_comp = int(input(input_statement)) #user input must be a number
            if num_of_comp == 0: #breaks from while loop if user types 0
                break
        except:
            print("\nPlease enter an integer")
            continue
        else:
            break
    if num_of_comp == 0:
        exit('Program has ended') #exits from entire program, if user types 0
    return num_of_comp

def get_num_simulations(input_statement):
    '''
    Takes in an input statement argument and returns integer number of simualtions to be analysed.
    '''
    while True:
        try:
            num_of_simulations = int(input(input_statement)) #user input must be a number
            if num_of_simulations == 0: #breaks from while loop if user types 0
                break
        except:
            print("\nPlease enter an integer")
            continue
        else:
            break
    if num_of_simulations == 0:
        exit('Program has ended') #exits from entire program, if user types 0
    return num_of_simulations

def get_num_control_sequences(input_statement):
    '''
    Takes in an input statement argument and returns integer number of control sequences to be simulated and analysed.
    '''
    while True:
        try:
            num_control_sequences = int(input(input_statement)) #user input must be a number
            if num_control_sequences == 0: #breaks from while loop if user types 0
                break
        except:
            print("\nPlease enter an integer")
            continue
        else:
            break
    if num_control_sequences == 0:
        exit('Program has ended') #exits from entire program, if user types 0
    return num_control_sequences

def get_num_operation_modes(input_statement):
    '''
    Takes in an input statement arugument and returns integer number of operation modes to be simulated and analysed.
    '''
    while True:
        try:
            num_operation_modes = int(input(input_statement)) #user input must be a number
            if num_operation_modes == 0: #breaks from while loop if user types 0
                break
        except:
            print("\nPlease enter an integer")
            continue
        else:
            break
    if num_operation_modes == 0:
        exit('Program has ended') #exits from entire program, if user types 0
    return num_operation_modes

def get_excel_workbook(input_statement):
    '''
    Takes in an input statement argument and returns string excel workbook name
    '''
    while True:
        try:
            excel_workbook = input(input_statement) #user inputs a path/string_name of excel workbook that exists
            if excel_workbook == '0': #breaks from while loop if user inputs 0
                break
            if len(excel_workbook.split('.')) > 1: #checks if name of string is a valid excel workbook name 'C:\path\excelworkbookname.xlsx'
                raise NameError
        except:
            print('\nInput error, input: C:\path\excelworkbookname')
            continue
        else:
            break
    if excel_workbook == '0':
        exit('Program has ended')# exits from entire program, if user types 0
    return excel_workbook

def get_sheet_name(input_statement):
    '''
    Takes in an input statement arguement and returns a string excel sheetname
    '''
    while True:
        try:
            sheet_name = input(input_statement) #user inputs a string_name for excel sheetname that exists
            if sheet_name == '0': #breaks from while loop if user inputs 0
                break
        except:
            print('\nInvalid sheet name')
            continue
        else:
            break
    if sheet_name == '0':
        exit('Program has ended') #exits from entire program, if user types 0
    return sheet_name

def get_simulation_name(input_statement):
    '''
    Takes in a input statement argument and returns the string name for simualtion name
    '''
    while True:
        try:
            simulation_name = input(input_statement) #user inputs a string_name for operation mode combination
            if simulation_name == '0': #breaks from while loop if user inputs 0
                break
        except:
            print('\nInvalid operation mode name')
            continue
        else:
            break
    if simulation_name == '0':
        exit('Program has ended') #exits from entire program, if user types 0
    return simulation_name

def get_operation_mode_name(input_statement):
    '''
    Takes in a input statement argument and returns the string name for operation mode combination
    '''
    while True:
        try:
            operation_mode = input(input_statement) #user inputs a string_name for operation mode combination
            if operation_mode == '0': #breaks from while loop if user inputs 0
                break
        except:
            print('\nInvalid operation mode name')
            continue
        else:
            break
    if operation_mode == '0':
        exit('Program has ended') #exits from entire program, if user types 0
    return operation_mode

def get_control_sequence_name(input_statement):
    '''
    Takes in a input statement argument and returns a string name for control sequence combination
    '''
    while True:
        try:
            name = input(input_statement) #user inputs a string_name for operation mode combination
            if name == '0': #breaks from while loop if user inputs 0
                break
        except:
            print('\nInvalid control scheme name')
            continue
        else:
            break
    if name == '0':
        exit('Program has ended') #exits from entire program, if user types 0
    return name

def get_proportion(series1, series2):
    '''
    Takes in two series arguments and returns proportion of series1 in series1 and series2
    '''
    return round((series1/(series1+series2))*100) 

def get_figures_name(input_statement):
    '''
    Takes in a input statement argument and returns the name to be given to figures in results for simulation performed using one control sequence and one operation mode combination.
    '''
    while True:
        try:
            name = input(input_statement) #user inputs a string_name for operation mode combination
            if name == '0': #breaks from while loop if user inputs 0
                break
        except:
            print('\nInvalid figures name')
            continue
        else:
            break
    if name == '0':
        exit('Program has ended') #exits from entire program, if user types 0
    return name

def save_results(input_statement, fig):
    '''
    Takes in a input statement argument and asks user for a 'C:/path/filename' and saves results from analysis as a pdf
    '''
    while True:  
        try:
            figure_name = input(input_statement) #user inputs a string_name for operation mode combination
            if figure_name == '0': #breaks from while loop if user inputs 0
                break
            if len(figure_name.split('.')) > 1: #checks if filename of string is a valid name
                raise NameError
            fig.savefig(figure_name+'.pdf') #saves results as a pdf
        except:
            print("\nInvalid filename")
        else:
            print('Results saved')
            break
    if figure_name == '0':
        exit('Program has ended') #exits from entire program, if user types 0

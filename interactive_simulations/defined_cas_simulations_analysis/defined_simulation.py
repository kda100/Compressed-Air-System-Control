# -*- coding: utf-8 -*-
"""
Created on Sun Jul 26 22:22:45 2020
**************************************************************************************************************************************************
COMPRESSORS:

Compressor    Maximum Power (kW)    Maximum Flow (m3/hr)
1             125                   947
2             125                   947
3             125                   947
4             177                   1326
5             364                   2604
6             274                   1902

**************************************************************************************************************************************************
INTRODUCTION:
  
This Program simulates a compressed air system comprising of six compressors. It takes input as air demand readings (in m3/hr) from an excel 
workbook to determine the power output and air output of the compressors based on the operation mode of each compressor and a control 
sequencing logic for the compressors defined by the user. 

There are four types of operation modes that can be chosen for the compressors, defined and given in the 'compressor_models2' module in the 'cas_modules' 
directory: OnOff, LoadUnload, Inlet Modulation and VariableSpeed. The user-defined control sequencing logic is based on a chosen number of air demand 
setpoints.

**************************************************************************************************************************************************
DESCRIPTION:

First the program will ask the user to input the type of operation modes for each compressor and to input the trim compressor, after the program will 
ask the user for the number of setpoints, then ask which compressor's maximum air capacities will be used to determine each setpoint to develop a 
sequencing logic for the control of compressors for a given air demand.

The program will then ask the user to upload an excel 'workbook' then select the 'sheetname' that contains the hourly air demand data, the air demand must 
be in one single continuous column with the first row of the column having 'CA_READINGS' written inside it. Using the air demand data the program 
will then run the compressed air system simulation with the operation modes of the compressors and user-defined setpoints. 

The program will then ask the user to create an excel workbook to open by passing in the C:\path\excelworkbook.xlsx or open an existing workbook 
by also passingin C:\path\excelworkbook.xlsx. Then the program will ask the user for the sheetname and the coordinate (e.g. 'A1') of the excel 
workbook sheet to write the results into. The program will finish by asking the user the path and filename to save the results as, in the format 
C:\path\excelworkbook.xlsx.

**************************************************************************************************************************************************
@author: kda_1
"""

import cas_modules.control_scheme_functions as csf
from sys import exit

print("\nAt any point to exit the program type 0")


C1 = csf.GetCompressor1(1, 125, 947) #asks user for compressor 1, maximum_power = 125 kW, maximum_flow = 947
C2 = csf.GetCompressor1(2, 125, 947) #asks user for compressor 2, maximum_power = 125 kW, maximum_flow = 947
C3 = csf.GetCompressor1(3, 125, 947) #asks user for compressor 3, maximum_power = 125 kW, maximum_flow = 947
C4 = csf.GetCompressor1(4, 177, 1326) #asks user for compressor 4, maximum_power = 177 kW, maximum_flow = 1326
C5 = csf.GetCompressor1(5, 364, 2604) #asks user for compressor 5, maximum_power = 364 kW, maximum_flow = 2604
C6 = csf.GetCompressor1(6, 274, 1902) #asks user for compressor 6, maximum_power = 274 kW, maximum_flow = 1902

compressor_dict = {'C1':C1, 'C2':C2, 'C3':C3, 'C4':C4, 'C5':C5, 'C6':C6} #dictionary to store compressor objects using string names for compressors.
trim = csf.get_trim(compressor_dict) #asks user for trim compressor.

C1_power_output = [] #list object to store power output for compressor C1
C1_air_output = [] #list object to store air output for compressor C1
C2_power_output = [] #list object to store power output for compressor C2
C2_air_output = [] #list object to store air output for compressor C2
C3_power_output = [] #list object to store power output for compressor C3
C3_air_output = [] #list object to store air output for compressor C3
C4_power_output = [] #list object to store power output for compressor C4
C4_air_output = [] #list object to store air output for compressor C4
C5_power_output = [] #list object to store power output for compressor C5
C5_air_output = [] #list object to store air output for compressor C5
C6_power_output = [] #list object to store power output for compressor C6
C6_air_output = [] #list object to store air output for compressor C6

compressor_power_dict = {'C1':['C1_power_output', C1_power_output], 'C2':['C2_power_output', C2_power_output],
                         'C3':['C3_power_output', C3_power_output], 'C4':['C4_power_output', C4_power_output],
                         'C5':['C5_power_output', C5_power_output], 'C6':['C6_power_output', C6_power_output]} #dictionary to store a list of string names for compressor power outputs and compressor power output list, using the compressor string name. 

compressor_air_dict = {'C1':['C1_air_output', C1_air_output], 'C2':['C2_air_output', C2_air_output],
                       'C3':['C3_air_output', C3_air_output], 'C4':['C4_air_output', C4_air_output],
                       'C5':['C5_air_output', C5_air_output], 'C6':['C6_air_output', C6_air_output]} #dictionary to store a list of string names for compressor air outputs and compressor air output list, using the compressor string name.

print("Now asking for number of setpoints in control scheme")

setpoints = [0] #list for storing the integer setpoints
contri_compressors = [] #list to store compressor objects contributing to each setpoint
boundaries = [] #list to store the integer boundaries for the air_demand values

input_statement = 'How many setpoints in control scheme: '
n = csf.get_setpoints(input_statement) #function to get number of setpoints

print("\nNow asking what compressor(s) maximum capacities are contributing to the setpoints")
num = 1
while num<=n: #iterates over all the setpoints
    try:
        contri = input(f"Which compressor(s) are contributing to setpoint {num} (s{num}): ") #asks user for the compressor(s) contributing to each setpoint in the control sequenc logic.
        if contri == '0': #breaks from while loop if user types 0
            break
        lst_com = [] #list to store the contributing compressors for each setpoint
        setpoint = 0 #variable to store setpoint
        contri = [x.strip() for x in contri.upper().split(',')] #splits the contri input based on commas separating strings e.g. C1, C2, C3 => ['C1', 'C2', 'C3']
        for compressor in contri: 
            if len(compressor)>2: #checks if each compressor in contri is a valid input length input
                exit()
            if len(contri) != len(set(contri)):
                exit()
        for compressor in contri:
            lst_com.append(compressor_dict[compressor]) #checks if input is in compressor_dict and if it is adds to lst_com 
        contri_compressors.append(lst_com) #add lst_com to contri_compressors
        for compressor in lst_com:
            setpoint = setpoint + compressor.maximum_flow #add maximum flow of compressors in lst_com to setpoint
        setpoints.append(setpoint) #adds setpoint to setpoints list.
    except:
        print('\nInvalid compressor object(s) chosen\n')
        print('\nCompressor(s) chosen must be at least one of the compressors C1 - C6,\nmultiple compressors should be written with a "," between them')
        continue
    else:
        num = num + 1
        continue

if contri == '0': 
    exit('Program has ended') #exits from entire program, if user types 0
    
contri_compressors = [x for _,x in sorted(zip(setpoints[1:], contri_compressors))] #sorts contri_compressors list based on setpoints list
            
setpoints.sort() #sorts setpoints list

for i in range(len(setpoints)-1):
    boundaries.append((setpoints[i],setpoints[i+1])) #uses setpoints list to create a list of tuples corresponding to the min and max boundaries (min, max).

checking = True
while checking == True:
    input_statement = "Name and/or path of excel workbook with Compressed Air Production Data: " 
    compressed_air_data = csf.get_workbook(input_statement) #gets excel workbook object
    input_statement = "Name of sheet with stacked compressed air volume data: "
    air_volume_stacked = csf.get_sheet(input_statement, compressed_air_data) #gets excel sheet object from workbook
    
    for row in air_volume_stacked.iter_rows(min_row = 1, max_row = 1, min_col = 1, max_col = air_volume_stacked.max_column): #iterates over first row of air_volume_stacked, excel sheet object
        col = 1
        for cell in row: #iterates over each cell in row
            col = col + 1
            if cell.value == 'CA_READINGS': #checks for cell in first row that contains 'CA_READINGS'
                coordinate = cell.coordinate #stores the coordinate of cell with 'CA_READINGS'
                checking = False # breaks out of while loop if cell is found
            elif col == air_volume_stacked.max_column:
                print("\nEnsure 'CA_reading' is in first row of column with compressed air volume data\n")
                continue #asks user for excel workbook and sheetname again if no 'CA_READINGS' is found in first row.

print('Performing Simulation...')

for cell in air_volume_stacked[coordinate[0]]: #iterates over column with air_volume data in m3/hr
    if type(cell.value) == int or type(cell.value) == float: #checks if value in cell is an int or float
        air_demand = round(cell.value) #stores data as air demand
        active_compressors = csf.get_active_compressors(boundaries, air_demand, setpoints, contri_compressors, trim) #returns a list of active compressors
        
        csf.update_power_output(C1, active_compressors, C1_power_output, air_demand) #appends result of C1.power_output to C1_power_output
        csf.update_power_output(C2, active_compressors, C2_power_output, air_demand) #appends result of C2.power_output to C2_power_output
        csf.update_power_output(C3, active_compressors, C3_power_output, air_demand) #appends result of C3.power_output to C3_power_output
        csf.update_power_output(C4, active_compressors, C4_power_output, air_demand) #appends result of C4.power_output to C4_power_output
        csf.update_power_output(C5, active_compressors, C5_power_output, air_demand) #appends result of C5.power_output to C5_power_output
        csf.update_power_output(C6, active_compressors, C6_power_output, air_demand) #appends result of C6.power_output to C6_power_output
        
        csf.update_air_output(C1, active_compressors, C1_air_output, air_demand) #appends result of C1.air_output to C1_air_output
        csf.update_air_output(C2, active_compressors, C2_air_output, air_demand) #appends result of C2.air_output to C2_air_output
        csf.update_air_output(C3, active_compressors, C3_air_output, air_demand) #appends result of C3.air_output to C3_air_output
        csf.update_air_output(C4, active_compressors, C4_air_output, air_demand) #appends result of C4.air_output to C4_air_output
        csf.update_air_output(C5, active_compressors, C5_air_output, air_demand) #appends result of C5.air_output to C5_air_output
        csf.update_air_output(C6, active_compressors, C6_air_output, air_demand) #appends result of C6.air_output to C6_air_output

print('\nSimulation complete')

print("\nCreate an excel document where the compressor power_output and air_output for each compressor will be saved")

input_statement = "Name and/or path of excel document where power_output and air_output will be copied to: "
results = csf.get_workbook(input_statement) #gets excel workbook object

input_statement = "Name of sheet where compressor's power output will be saved: "
power_output = csf.get_sheet(input_statement, results) #gets excel sheet object from workbook

input_statement = "What cell would you like compressor's power output saved: " 
coor = csf.get_coordinate(input_statement, power_output) #gets excel coordinate object from sheet 

for row in power_output.iter_rows(min_row = coor.row, max_row = coor.row, min_col = coor.col_idx, max_col = coor.col_idx + len(compressor_dict)-1): #iterates over row attribute with coordinate object 
    comp = 1
    for cell in row:
        cell = compressor_power_dict['C{}'.format(comp)][0] #writes each cell object columns with the names in first index of compressor_power_dict key values e.g C1_Power_Output, C2_Power_Output... 
        comp = comp + 1

comp = 1
for col in power_output.iter_cols(min_row = coor.row+1, max_row = coor.row + len(compressor_power_dict['C{}'.format(comp)][1]), min_col = coor.col_idx, max_col = coor.col_idx + len(compressor_dict)-1): #iterates over excel sheet object
    i = 0
    for cell in col:
        cell = compressor_power_dict['C{}'.format(comp)][1][i] #writes to each cell object the data contained in the second index of the compressor_power_dict key values e.g. C1_Power_Output list,C1_Power_Output list...
        i = i + 1
    comp = comp + 1
    
input_statement = "Name of sheet where you would like compressor's air output copied: "
air_output = csf.get_sheet(input_statement, results) #gets excel sheet object from workbook

input_statement = "What cell would you like compressor's air output saved: " 
coor = csf.get_coordinate(input_statement, power_output) #gets excel coordinate object from sheet

for row in air_output.iter_rows(min_row = coor.row, max_row = coor.row, min_col = coor.col_idx, max_col = coor.col_idx + len(compressor_dict)-1): #iterates over row attribute with coordinate object 
    comp = 1
    for cell in row:
        cell = compressor_air_dict['C{}'.format(comp)][0] #writes each cell object columns with the names in first index of compressor_power_dict key values e.g C1_Air_Output, C2_Air_Output... 
        comp + comp + 1

comp = 1
for col in air_output.iter_cols(min_row = coor.row+1, max_row = coor.row + len(compressor_air_dict['C{}'.format(comp)][1]), min_col = coor.col_idx, max_col = coor.col_idx + len(compressor_dict)-1): #iterates over excel sheet object
    i = 0
    for cell in col:
        cell = compressor_power_dict['C{}'.format(comp)][1][i] #writes to each cell object the data contained in the second index of the compressor_power_dict key values e.g. C1_Power_Output list,C1_Power_Output list...
        i = i + 1
    comp = comp + 1

input_statement = "Workbook name the results should be saved as: "
csf.save_and_close_results(input_statement, results) #saves and closes results.

print('\nProgram finished')
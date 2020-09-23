# -*- coding: utf-8 -*-
"""
Created on Thu Mar 26 20:56:37 2020
@author: kda_1
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
  
This Program simulates a compressed air system comprising of six compressors. It takes input as hourly air demand readings (in m3/hr) from an excel 
workbook to determine the power output and air output of the compressors based on the operation mode of each compressor and an optimised control 
sequencing logic for the compressors. 

There are four types of operation modes that can be chosen for the compressors, defined and given in the 'compressor_models2' module in the 'cas_modules' 
directory: OnOff, LoadUnload, Inlet Modulationand VariableSpeed. The optimised control sequencing logic is based on minimising the energy used to 
produce the air demanded.

**************************************************************************************************************************************************
DESCRIPTION:

First the program will ask the user to input the type of operation modes for each compressor and to input the trim compressor.

The program will then ask the user to upload an excel 'workbook' then select the 'sheetname' that contains the hourly air demand data, the air demand 
data must be in one single continuous column with the first row of the column having 'CA_READINGS' written inside it. Using the air demand data, the 
program will then run the compressed air system simulation with the operation modes of the compressors, using the optimised control sequencing logic. 

The program will then ask the user to create an excel workbook to open by passing in the C:\path\excelworkbook.xlsx or open an existing workbook 
by also passingin C:\path\excelworkbook.xlsx. Then the program will ask the user for the sheetname and the coordinate (e.g. 'A1') of the excel 
workbook sheet to write the results into. The program will finish by asking the user the path and filename to save the results as, in the format 
C:\path\excelworkbook.xlsx.

**************************************************************************************************************************************************
"""

import cas_modules.control_scheme_functions as csf
import numpy as np

print("\nAt any point to exit the program type 0")
print('\nChoose the compressor type for compressors C1 - C6')

C1 = csf.GetCompressor1(1, 125, 947) #asks user for compressor 1, maximum_power = 125 kW, maximum_flow = 947
C2 = csf.GetCompressor1(2, 125, 947) #asks user for compressor 2, maximum_power = 125 kW, maximum_flow = 947
C3 = csf.GetCompressor1(3, 125, 947) #asks user for compressor 3, maximum_power = 125 kW, maximum_flow = 947
C4 = csf.GetCompressor1(4, 177, 1326) #asks user for compressor 4, maximum_power = 177 kW, maximum_flow = 1326
C5 = csf.GetCompressor1(5, 364, 2604) #asks user for compressor 5, maximum_power = 364 kW, maximum_flow = 2604
C6 = csf.GetCompressor1(6, 274, 1902) #asks user for compressor 6, maximum_power = 274 kW, maximum_flow = 1902

compressor_dict = {'C1':C1, 'C2':C2, 'C3':C3, 'C4':C4, 'C5':C5, 'C6':C6} #dictionary to store compressor objects using string names for compressors.
trim = csf.get_trim(compressor_dict) #asks user for trim compressor.

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

print('\nPerforming simulation...')

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
                         'C5':['C5_power_output', C5_power_output], 'C6':['C6_power_output', C6_power_output]} #dictionary to store a list of string names for compressor power outputs and compressor power output list, using the compressor string name
        

compressor_air_dict = {'C1':['C1_air_output', C1_air_output], 'C2':['C2_air_output', C2_air_output],
                       'C3':['C3_air_output', C3_air_output], 'C4':['C4_air_output', C4_air_output],
                       'C5':['C5_air_output', C5_air_output], 'C6':['C6_air_output', C6_air_output]} #dictionary to store a list of string names for compressor air outputs and compressor air output list, using the compressor string name.


c1 = np.array([csf.get_power_array(C1, trim), csf.get_air_array(C1, trim)]) #2D array object of power_output and air outputs for compressor C1
c2 = np.array([csf.get_power_array(C2, trim), csf.get_air_array(C2, trim)]) #2D array object of power_output and air outputs for compressor C2
c3 = np.array([csf.get_power_array(C3, trim), csf.get_air_array(C3, trim)]) #2D array object of power_output and air outputs for compressor C3
c4 = np.array([csf.get_power_array(C4, trim), csf.get_air_array(C4, trim)]) #2D array object of power_output and air outputs for compressor C4
c5 = np.array([csf.get_power_array(C5, trim), csf.get_air_array(C5, trim)]) #2D array object of power_output and air outputs for compressor C5
c6 = np.array([csf.get_power_array(C6, trim), csf.get_air_array(C6, trim)]) #2D array object of power_output and air outputs for compressor C6

c1_air_output = np.array([]) #air_output 1D array object for compressor C1
c2_air_output = np.array([]) #air_output 1D array object for compressor C2
c3_air_output = np.array([]) #air_output 1D array object for compressor C3
c4_air_output = np.array([]) #air_output 1D array object for compressor C4
c5_air_output = np.array([]) #air_output 1D array object for compressor C5
c6_air_output = np.array([]) #air_output 1D array object for compressor C6
total_air_output = np.array([]) #air_output 1D array object for all compressors in compressed air system

for air_output1 in c1[1]: #iterates over air_output1 for compressor C1
    for air_output2 in c2[1]: #iterates over air_output1 for compressor C2
        for air_output3 in c3[1]: #iterates over air_output1 for compressor C3
            for air_output4 in c4[1]: #iterates over air_output1 for compressor C4
                for air_output5 in c5[1]: #iterates over air_output1 for compressor C5
                    for air_output6 in c6[1]: #iterates over air_output1 for compressor C6
                            c1_air_output = np.append(c1_air_output, air_output1) #appends air_output1 to c1_air_output array
                            c2_air_output = np.append(c2_air_output, air_output2) #appends air_output2 to c2_air_output array
                            c3_air_output = np.append(c3_air_output, air_output3) #appends air_output3 to c3_air_output array
                            c4_air_output = np.append(c4_air_output, air_output4) #appends air_output4 to c4_air_output array
                            c5_air_output = np.append(c5_air_output, air_output5) #appends air_output5 to c5_air_output array
                            c6_air_output = np.append(c6_air_output, air_output6) #appends air_output6 to c6_air_output array
                            total_air_output = np.append(total_air_output, (air_output1 + air_output2 + air_output3 + air_output4 + air_output5 + air_output6)) #appends total compressor air_outputs to total_air_output array

air_output_array = np.vstack((c1_air_output, c2_air_output, c3_air_output, c4_air_output, c5_air_output, c6_air_output, total_air_output)) #creates a stack array of arrays

c1_power_output = np.array([]) #power_output 1D array object for compressor C1
c2_power_output = np.array([]) #power_output 1D array object for compressor C2
c3_power_output = np.array([]) #power_output 1D array object for compressor C3
c4_power_output = np.array([]) #power_output 1D array object for compressor C4
c5_power_output = np.array([]) #power_output 1D array object for compressor C5
c6_power_output = np.array([]) #power_output 1D array object for compressor C6
total_power_output = np.array([]) #power_output 1D array object for all compressors in compressed air system

for power_output1 in c1[0]: #iterates over power_output1 for compressor C1
    for power_output2 in c2[0]: #iterates over power_output2 for compressor C2
        for power_output3 in c3[0]: #iterates over power_output3 for compressor C3
            for power_output4 in c4[0]: #iterates over power_output4 for compressor C4
                for power_output5 in c5[0]: #iterates over power_output5 for compressor C5
                    for power_output6 in c6[0]: #iterates over power_output6 for compressor C6
                            c1_power_output = np.append(c1_power_output, power_output1) #appends power_output1 to c1_power_output array
                            c2_power_output = np.append(c2_power_output, power_output2) #appends power_output2 to c2_power_output array
                            c3_power_output = np.append(c3_power_output, power_output3) #appends power_output3 to c3_power_output array
                            c4_power_output = np.append(c4_power_output, power_output4) #appends power_output4 to c4_power_output array
                            c5_power_output = np.append(c5_power_output, power_output5) #appends power_output5 to c5_power_output array
                            c6_power_output = np.append(c6_power_output, power_output6) #appends power_output6 to c6_power_output array
                            total_power_output = np.append(total_power_output, (power_output1 + power_output2 + power_output3 + power_output4 + power_output5 + power_output6)) #appends total compressor power_outputs to total_power_output array

power_output_array = np.vstack((c1_power_output, c2_power_output, c3_power_output, c4_power_output, c5_power_output, c6_power_output, total_power_output)) #creates a stack array of arrays                       

optimisation_array = np.vstack((total_power_output, total_air_output)) #creates a stack array of arrays        

for cell in air_volume_stacked[coordinate[0]]: #iterates over column with air_volume data in m3/hr
    if type(cell.value) == int or type(cell.value) == float: #checks if value in cell is an int or float
        air_demand = round(cell.value) #stores data as air demand
        optimised_indexes = [] #list for indexes
        optimised_power_output = [] #list for optimised power outputs
        for index in np.where(optimisation_array[1] == air_demand)[0]: #checks air outputs array in optimisation arrays for indices where air_demand is the same
            optimised_indexes.append(index) #appends index to optimised_indexes list.
        for index in optimised_indexes:
            optimised_power_output.append(optimisation_array[0][index]) #checks power output array in optimisation array for power output values where the index is the same as in optimised index
        minimum_power_index = optimised_power_output.index(min(optimised_power_output)) #stores index of optimised power index, where the optimised_power_output is minimum
        optimised_index = optimised_indexes[minimum_power_index] #uses the minimum power index to store the an optimised index in the optimisation array where the power_output is minimum for the given air demand
        
        C1_power_output.append(power_output_array[0][optimised_index]) #appends power_output value in power_output_array for C1 to C1_power_output
        C1_air_output.append(air_output_array[0][optimised_index]) #appends air_output value in air_output_array for C1 to C1_air_output
        C2_power_output.append(power_output_array[1][optimised_index]) #appends power_output value in power_output_array for C2 to C2_power_output
        C2_air_output.append(air_output_array[1][optimised_index]) #appends air_output value in air_output_array for C1 to C1_air_output
        C3_power_output.append(power_output_array[2][optimised_index]) #appends power_output value in power_output_array for C3 to C3_power_output
        C3_air_output.append(air_output_array[2][optimised_index]) #appends air_output value in air_output_array for C3 to C3_air_output
        C4_power_output.append(power_output_array[3][optimised_index]) #appends power_output value in power_output_array for C4 to C4_power_output
        C4_air_output.append(air_output_array[3][optimised_index]) #appends air_output value in air_output_array for C4 to C4_air_output
        C5_power_output.append(power_output_array[4][optimised_index]) #appends power_output value in power_output_array for C4 to C4_power_output
        C5_air_output.append(air_output_array[4][optimised_index]) #appends air_output value in air_output_array for C5 to C5_air_output
        C6_power_output.append(power_output_array[5][optimised_index]) #appends power_output value in power_output_array for C6 to C6_power_output
        C6_air_output.append(air_output_array[5][optimised_index]) #appends air_output value in air_output_array for C6 to C6_air_output
        
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
        cell = compressor_power_dict['C{}'.format(comp)][1][i]#writes to each cell object the data contained in the second index of the compressor_power_dict key values e.g. C1_Power_Output list,C1_Power_Output list...
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
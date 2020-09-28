# -*- coding: utf-8 -*-
"""
Created on Thu Mar 26 20:56:37 2020
@author: kda_1

***************************************************************************************************************************************
This program loads an excel workbook from the same directory the script is contained in, called 'compressed_air_data_12_bar.xlsx'
which contains the 12 bar compressed air data from a leading car manufacturer's 12 bar compressed air system. The program then loads
the sheet 'air_volume_stacked_2019' and iterates over the column that contains the 12 bar compressed air data for 2019. For each
air demand value contained in the sheet the program produces and stores air output and power output values for each compressor 
based on a compressor optimised control sequencing logic that uses the compressors based on minimising the energy needed to meet 
the air demand. The program then writes the power output and air output values into two separate sheets called 'compressors_power_output' 
and 'compressors_air_output' in an excel workbook called 'opt_op1_2019_results.xlsx' located in the same directory as the script.

***************************************************************************************************************************************
2019
Maximum Flow of System - FAD (m3/hr) = 8672
Maximum Air Demand of Plant - FAD (m3/hr) = 5340

C1 = LoadUnload: max_power = 125, max_flow = 947
C2 = LoadUnload: max_power = 125, max_flow = 947
C3 = LoadUnload: max_power = 125, max_flow = 947
C4 = LoadUnload: max_power = 177, max_flow = 1326
C5 = InletModulation max_power = 384, max_flow = 2604
C6 = LoadUnload: max_power = 274, max_flow = 1902

***************************************************************************************************************************************
"""

import openpyxl
import cas_modules.compressor_models1 as cm1
import numpy as np

C1 = cm1.LoadUnload(125, 947) #create LoadUnload compressor object for C1
C2 = cm1.LoadUnload(125, 947) #create LoadUnload compressor object for C2
C3 = cm1.LoadUnload(125, 947) #create LoadUnload compressor object for C3
C4 = cm1.LoadUnload(177, 1326) #create LoadUnload compressor object for C4
C5 = cm1.InletModulation(384, 2604) #create InletModulation compressor object for C5
C6 = cm1.LoadUnload(274, 1902) #create LoadUnload compressor object for C6

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

c1 = np.array([get_power_array(C1, trim),get_air_array(C1, trim)]) #2D array object of power_output and air outputs for compressor C1
c2 = np.array([get_power_array(C2, trim),get_air_array(C2, trim)]) #2D array object of power_output and air outputs for compressor C2
c3 = np.array([get_power_array(C3, trim),get_air_array(C3, trim)]) #2D array object of power_output and air outputs for compressor C3
c4 = np.array([get_power_array(C4, trim),get_air_array(C4, trim)]) #2D array object of power_output and air outputs for compressor C4
c5 = np.array([get_power_array(C5, trim),get_air_array(C5, trim)]) #2D array object of power_output and air outputs for compressor C5
c6 = np.array([get_power_array(C6, trim),get_air_array(C6, trim)]) #2D array object of power_output and air outputs for compressor C6

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

Compressed_Air_Data_12_Bar = openpyxl.load_workbook('compressed_air_data_12_bar.xlsx') #open compressed air system data workbook in current directory
Air_Volume_Stacked_2019 = Compressed_Air_Data_12_Bar['air_volume_stacked_2019'] #calls sheet 'air_volume_stacked_2018' in workbook
for row in range(1,8089):
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
    
Control_Scheme_Results = openpyxl.load_workbook('opt_op4_2019_results.xlsx') #opens results workbook in current directory
Compressors_Power_Output = Control_Scheme_Results['compressors_power_output'] #calls sheet 'compressors_power_output' in workbook
Compressors_Power_Output['D1'].value = 'C1_power_output' #names cell in workbook
for i in range(len(C1_power_output)):
    Compressors_Power_Output['D' + str(i + 2)].value = C1_power_output[i] #writes data in C1_power_output to sheet
Compressors_Power_Output['E1'].value = 'C2_power_output' #names cell in workbook
for i in range(len(C2_power_output)):
    Compressors_Power_Output['E' + str(i + 2)].value = C2_power_output[i] #writes data in C2_power_output to sheet
Compressors_Power_Output['F1'].value = 'C3_power_output' #names cell in workbook
for i in range(len(C3_power_output)):
    Compressors_Power_Output['F' + str(i + 2)].value = C3_power_output[i] #writes data in C3_power_output to sheet
Compressors_Power_Output['G1'].value = 'C4_power_output' #names cell in workbook
for i in range(len(C4_power_output)):
    Compressors_Power_Output['G' + str(i + 2)].value = C4_power_output[i] #writes data in C4_power_output to sheet
Compressors_Power_Output['H1'].value = 'C5_power_output' #names cell in workbook
for i in range(len(C5_power_output)):
    Compressors_Power_Output['H' + str(i + 2)].value = C5_power_output[i] #writes data in C5_power_output to sheet
Compressors_Power_Output['I1'].value = 'C6_power_output' #names cell in workbook
for i in range(len(C6_power_output)):
    Compressors_Power_Output['I' + str(i + 2)].value = C6_power_output[i] #writes data in C6_power_output to sheet

Compressors_Air_Output = Control_Scheme_Results['compressors_air_output'] #calls sheet 'compressors_air_output' in workbook
Compressors_Air_Output['D1'].value = 'C1_air_output' #names cell in workbook
for i in range(len(C1_air_output)):
    Compressors_Air_Output['D' + str(i + 2)].value = C1_air_output[i] #writes data in C1_air_output to sheet
Compressors_Air_Output['E1'].value = 'C2_air_output' #names cell in workbook
for i in range(len(C2_air_output)):
    Compressors_Air_Output['E' + str(i + 2)].value = C2_air_output[i] #writes data in C2_air_output to sheet
Compressors_Air_Output['F1'].value = 'C3_air_output' #names cell in workbook
for i in range(len(C3_air_output)):
    Compressors_Air_Output['F' + str(i + 2)].value = C3_air_output[i] #writes data in C3_air_output to sheet
Compressors_Air_Output['G1'].value = 'C4_air_output' #names cell in workbook
for i in range(len(C4_air_output)):
    4Compressors_Air_Output['G' + str(i + 2)].value = C4_air_output[i] #writes data in C4_air_output to sheet
Compressors_Air_Output['H1'].value = 'C5_air_output' #names cell in workbook
for i in range(len(C5_air_output)):
    Compressors_Air_Output['H' + str(i + 2)].value = C5_air_output[i] #writes data in C5_air_output to sheet
Compressors_Air_Output['I1'].value = 'C6_air_output' #names cell in workbook
for i in range(len(C6_air_output)):
    Compressors_Air_Output['I' + str(i + 2)].value = C6_air_output[i] #writes data in C6_air_output to sheet
   
Control_Scheme_Results.save('opt_op4_2019_results.xlsx') #saves results workbook
Control_Scheme_Results.close() #closes results workbook
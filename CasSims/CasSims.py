# -*- coding: utf-8 -*-
"""
Created on Sun Jul 26 22:22:45 2020

**************************************************************************************************************************************************
INTRODUCTION:
  
This Program simulates the performance of the compressors of a compressed air system where the user defines the number of compressors as well as 
the maximum air volume capacities, maximum power consumptions, operations modes and the control sequencing logic of the compressors. It takes as 
input air demand readings from an excel workbook to determine the power output and air output of the compressors.

There are four types of operation modes that can be chosen, defined and given in the 'compressor_models2' module in the 'cas_modules' directory: OnOff, 
LoadUnload, Inlet Modulation and VariableSpeed. The user-defined control sequencing logic is based on a chosen number of air demand setpoints and the 
optimised control sequencing logic minimises the amount of energy needed to produce a given air demand. 

**************************************************************************************************************************************************
DESCRIPTION:
    
First the program will ask the user for the number of simulations to be analysed and number of compressors that will be in the simulation. 
Then the user will be asked to input the maximum power, maximum air flows and the types of operation modes for each compressor. After the program will
ask the user for the compressor that will be used as the trim compressor and the type of control sequencing logic that will be used in the simulation.

If the sequence is defined the program will ask the user for the number of setpoints, then ask which compressor's maximum air capacities will be 
used to determine each setpoint and develop a sequencing logic for the compressors and If the control sequence is optimised then the program will use 
the optimised compressor sequencing logic. 

The program will then ask the user to upload an excel 'workbook' then select the 'sheetname' that contains the air demand data, the air demand must 
be in one single continuous column with the first row of the column having 'CA_READINGS' written inside it. The program will then run simulations with 
combinination of compressor operation modes and the control sequence 

The program will then ask the user to name the results of the simulation before  (e.g. operation mode 1 to op1) and each control sequencing logic 
before producing and displaying graphs of the performance metrics for the compressors in the simulation. 

************************************************************************************************************************************************** 
@author: kda_1
"""
#add comments, make application executable, look into threading when uploading workbook, sheets and performing simulations (allow to close program)

import cas_modules.control_scheme_functions as csf
import pandas as pd
import cas_modules.analysis_functions as af
import cas_modules.simulation_guis as gui
import matplotlib.pyplot as plt
import matplotlib
import numpy as np
from itertools import product
import tkinter as tk
import threading
import cas_modules.compressor_models2 as cm2

def get_simulation_details():
    '''
    Creates user interface that asks user for the details on the compressors and the control sequence of the
    simulation they would like to perform.
    '''
    
    num_of_simulations = int(main.num_of_simulations) #gets number of simulations from MainInterface
    
    simulation_user_inputs_interface = tk.Toplevel()
    
    global simulation_details #assigns simulation_details as global variable
    simulation_details = gui.SimulationUserInputsInterface(master=simulation_user_inputs_interface, 
                           num_of_simulations=num_of_simulations, 
                           main=main, 
                           main_function=prepare_simulation_parameters) 
    #calls instance of SimulationUserInputsInterface in simulation_guis on simulation_user_inputs_interface.

def prepare_simulation_parameters(): #second main function takes in parent window.
    
    '''
    Prepares global simulation parameters and objects needed to run simulations and store results.
    '''
    
    num_of_compressors = simulation_details.num_of_compressors #gets number of compressors from user input
    
    global compressors_dict 
    global compressors_power_dict
    global compressors_air_dict 
    
    compressors_dict = {} #creates dictionary to store compressor objects in compressor_models2 module from user inputs
    compressors_power_dict = {} #creates dictionary to store compressor's power output results from simulations
    compressors_air_dict = {} #creates dictionary to store compressor's air output results from simulations
    for i in range(num_of_compressors):
        compressors_dict[f'C{i+1}'] = csf.GetCompressor2(simulation_details.operation_mode_entries[f'C{i+1}'][-1], simulation_details.maximum_power_entries[f'C{i+1}'][-1], simulation_details.maximum_air_flow_entries[f'C{i+1}'][-1]) #gets the users compressor objects from inputs in SimulationUserInputsInterface 
        compressors_power_dict[f'C{i+1}'] = [] #creates empty list to store power output results for each compressor in simulation 
        compressors_air_dict[f'C{i+1}'] = [] #creates empty list to store air output results for each compressor in simulation 
        
    global trim
    trim = compressors_dict[simulation_details.trim] #stores trim compressor from user input
    
    global control_sequence_type
    control_sequence_type = simulation_details.control_sequence_type #stores control sequence type from user input
    
    if control_sequence_type == 'def': #if control sequence type is defined
        
        global setpoints
        global contributing_compressors
        global boundaries
        
        num_of_setpoints = simulation_details.num_of_setpoints
        
        if len(compressors_dict) == 1: #if only one compressor is in simulation
            setpoints = [0, compressors_dict['C1'].maximum_flow] #list for storing setpoints
            contributing_compressors = [[compressors_dict['C1']]] #list to store compressor objects contributing to each setpoint
            boundaries = [(setpoints[0],setpoints[1])] #list to store the integer boundaries for the air_demand values
        else:
            setpoints = [0] #list for storing the integer setpoints
            contributing_compressors = [] #list to store compressor objects contributing to each setpoint
            boundaries = [] #list to store the integer boundaries for the air_demand values
           
            for i in range(num_of_setpoints):
                contri = simulation_details.setpoint_entries[f's{i+1}'][-1]#asks user for the compressor(s) contributing to each setpoint in the control sequenc logic.
                lst_com = [] #list to store the contributing compressors for each setpoint
                setpoint = 0 #variable to store setpoint
                for compressor in contri:
                    lst_com.append(compressors_dict[compressor]) #checks if input is in compressor_dict and if it is adds to lst_com 
                contributing_compressors.append(lst_com) #add lst_com to contri_compressors
                for compressor in lst_com:
                    setpoint = setpoint + compressor.maximum_flow #add maximum flow of compressors in lst_com to setpoint
                setpoints.append(setpoint) #adds setpoint to setpoints list.
            
            contributing_compressors = [x for _,x in sorted(zip(setpoints[1:], contributing_compressors))] #sorts contri_compressors list based on setpoints list
                
            setpoints.sort() #sorts setpoints list
                
            for i in range(len(setpoints)-1):
                boundaries.append((setpoints[i],setpoints[i+1])) #uses setpoints list to create a list of tuples corresponding to the min and max boundaries (min, max).
    
    simulation_thread = threading.Thread(target=run_simulation) #thread for simulation
    simulation_thread.start()
    
    
def run_simulation():
    
    '''
    Runs the simulation and stores the compressors power output results in the compressors_power_dict and the compressors 
    air output results in the compressors_air_dict.
    '''

    air_demand_data_sheet = simulation_details.air_demand_data_sheet #stores excel sheet with air demand data from user input
    air_demand_data_column = simulation_details.air_demand_data_column #stores column name in sheet with air demand data from user input
    
    if control_sequence_type == 'opt': #if control scheme type is optimisation run optimisation simulation
        
        compressors_power_array_dict = {} 
        compressors_air_array_dict = {}
        compressors_air_output_opt_dict = {}
        compressors_power_output_opt_dict = {}
        for i, compressor in enumerate(compressors_dict):
            compressors_power_array_dict[f'c{i+1}'] = np.array(csf.get_power_array(compressors_dict[f'C{i+1}'], trim)) #1D array object of power_output possibilities for each compressor
            compressors_air_array_dict[f'c{i+1}'] = np.array(csf.get_air_array(compressors_dict[f'C{i+1}'], trim)) #1D array object of air_output possibilities for each compressor
            compressors_air_output_opt_dict[f'c{i+1}_air_output'] = np.array([]) #empty air_output 1D array object for each compressor
            compressors_power_output_opt_dict[f'c{i+1}_power_output'] = np.array([]) #empty power_output 1D array object for each compressor
        
        total_air_output = np.array([]) #empty air_output 1D array object for total air outputs of compressors in compressed air system
        total_power_output = np.array([]) #empty power_output 1D array object for total power outputs of compressors in compressed air system
        
        power_lst = [compressors_power_array_dict[f'c{x+1}'] for x in range(len(compressors_dict))] #list of array objects with power_output possibilities for each compressor
        air_lst = [compressors_air_array_dict[f'c{x+1}'] for x in range(len(compressors_dict))] #list of array objects with air_output possibilities for each compressor
        prod_power = list(product(*power_lst)) #returns all possible combinations of power_outputs for each compressor as a tuple
        prod_air = list(product(*air_lst)) #returns all possible combinations of air_outputs for each compressor as a tuple
        
        comp_num = 1
        while comp_num<=len(compressors_power_output_opt_dict):
            for power_tup in prod_power: #iterates over power output tuples
                compressors_power_output_opt_dict[f'c{comp_num}_power_output'] = np.append(compressors_power_output_opt_dict[f'c{comp_num}_power_output'], power_tup[comp_num-1]) #adds individual power output possibility to compressor array in 'compressors_power_output_opt_dict'
            comp_num = comp_num + 1
            
        comp_num = 1
        while comp_num<=len(compressors_air_output_opt_dict):
            for air_tup in prod_air: #iterates over power output tuples
                compressors_air_output_opt_dict[f'c{comp_num}_air_output'] = np.append(compressors_air_output_opt_dict[f'c{comp_num}_air_output'], air_tup[comp_num-1]) #adds individual air output possibility to compressor array in 'compressors_air_output_opt_dict'
            comp_num = comp_num + 1
            
        for power_tup in prod_power: #iterates over tuples
            total_power_output = np.append(total_power_output, sum(power_tup)) #adds total power output possibility to 'total_power_output' array
            
        for air_tup in prod_air: #iterates over tuples
            total_air_output = np.append(total_air_output, sum(air_tup)) #adds total air output possibility to 'total_air_output' array
        
        power_output_array = np.vstack(tuple(compressors_power_output_opt_dict[f'c{i+1}_power_output'] for i in range(len(compressors_power_output_opt_dict)))) #creates a stack array of arrays                       
        air_output_array = np.vstack(tuple(compressors_air_output_opt_dict[f'c{i+1}_air_output'] for i in range(len(compressors_air_output_opt_dict)))) #creates a stack array of arrays 
        
        optimisation_array = np.vstack((total_power_output, total_air_output)) #creates a stack array of arrays        
        
        for cell in air_demand_data_sheet[air_demand_data_column]: #iterates over column with air_volume data in m3/hr
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
                
                for i in range(len(compressors_dict)): #iterates over all compressors
                    
                    if type(compressors_dict[f'C{i+1}']) == cm2.OnOff or type(compressors_dict[f'C{i+1}']) == cm2.VariableSpeed:
                        compressors_power_dict[f'C{i+1}'].append(power_output_array[i][optimised_index]) #if compressor type is OnOff or VariableSpeed append air_output or power_output as normal
                        compressors_air_dict[f'C{i+1}'].append(air_output_array[i][optimised_index])
                    
                    else: #if compressor type is inlet modulation or LoadUnload
                        if len(compressors_air_dict[f'C{i+1}']) == 0: #at the start of the simulation when no data exists
                            
                            if air_output_array[i][optimised_index]==0: #compressor is producing no air
                                compressors_power_dict[f'C{i+1}'].append(0) #do not turn on compressor
                                 
                            else:
                                compressors_power_dict[f'C{i+1}'].append(power_output_array[i][optimised_index]) #otherwise turn on compressor
                                
                        elif len(compressors_air_dict[f'C{i+1}']) > 0: #for the rest of the simulation
                            
                            if air_output_array[i][optimised_index]>0: #if compressor is suppose to produce air
                                compressors_power_dict[f'C{i+1}'].append(power_output_array[i][optimised_index]) #then keep compressor on or turn on.
                            
                            elif compressors_air_dict[f'C{i+1}'][-1]==0: #if last air output compressor did not produce air
                                compressors_power_dict[f'C{i+1}'].append(0) #do not turn on compressor
                            
                            elif compressors_air_dict[f'C{i+1}'][-1]>0: #if last air output compressor did produce air
                                compressors_power_dict[f'C{i+1}'].append(power_output_array[i][optimised_index]) #keep compressor on
                                
                        compressors_air_dict[f'C{i+1}'].append(air_output_array[i][optimised_index]) #add air_output
                            
    
    if control_sequence_type == 'def': #control_sequence type is defined run defined simulation
        
        for cell in air_demand_data_sheet[air_demand_data_column]: #iterates over column with air_volume data
            if type(cell.value) == int or type(cell.value) == float: #checks if value in cell is an int or float
                air_demand = round(cell.value) #stores data as air demand
                _active_compressors = csf.get_active_compressors(boundaries, air_demand, setpoints, contributing_compressors, trim) #returns a list of active compressors
                
                for i in range(len(compressors_dict)): #iterates over all 'C1_power_output' and 'C1_air_output' lists
                    csf.update_power_output(compressors_dict[f'C{i+1}'], _active_compressors, compressors_power_dict[f'C{i+1}'], compressors_air_dict[f'C{i+1}'], air_demand) #appends power output of each compressor to compressor list
                    csf.update_air_output(compressors_dict[f'C{i+1}'], _active_compressors, compressors_air_dict[f'C{i+1}'], air_demand) #appends air output of each compressor to compressor list

    get_results_name()


def get_results_name():
    
    '''
    Gets the name that will be displayed on the results created.
    '''
    
    results_name_interface = tk.Toplevel() #creates toplevel windows for results name input
    global results_name #assigns results_name to be global 
    results_name = gui.NameResultsInterface(master=results_name_interface,
                                            main=main, 
                                            main_function = create_results) #creates a instance of NameResultsInterface from simulations gui on results_name_interface
    
    
def create_results(): 
    
    '''
    Creates figures to present the results from the simulation using the compressor_power_dict and the compressor_air_dict
    '''
    
    name = results_name.results_name #takes name attribute from root_class_3
    
    comp_power_outputs = pd.DataFrame(compressors_power_dict) #loads power_output results in compressor_power_dict into a dataframe
    comp_air_outputs = pd.DataFrame(compressors_air_dict) #loads air_output results in compressor_power_dict into a dataframe
    
    average_power = round(comp_power_outputs.mean()) #applies mean aggregate function to comp_power_outputs dataframe

    average_air = round(comp_air_outputs.mean()) #applies mean aggregate function to comp_air_outputs dataframe
    
    average_power_air = pd.concat([average_power, average_air], axis=1) #concatenates the average_power and average_air dataframes
    average_power_air.rename(columns = {0:'Average_Power_Output', 1:'Average_Air_Output'}, inplace=True) #renames the columns in the concatenated dataframes
    
    total_power = comp_power_outputs.sum() #applies sum aggregate function to comp_power_outputs dataframe
    
    total_air = comp_air_outputs.sum() #applies sum aggregate function to comp_air_outputs dataframe
    
    total_power_air = pd.concat([total_power, total_air], axis=1) #concatenates the total_power and total_air dataframes
    total_power_air.rename(columns = {0:'Total_Power_Output', 1:'Total_Air_Output'}, inplace=True) #renames the columns in the concatenated dataframes
    
    average_efficiency = (comp_power_outputs.mean())/(comp_air_outputs.mean()) #applies mean aggregate function to comp_power_outputs dataframe and mean aggregate function to comp_air_outputs dataframe then divides them.
    
    time_loaded_unloaded = pd.concat([comp_air_outputs[comp_air_outputs>0].count(), comp_air_outputs[comp_air_outputs==0].count()], axis=1) #concatenate two series where one contains the aggregated count data in comp_air_outputs that is greater than 0 and the other contains aggregated count data that is equal to 0
    time_loaded_unloaded.rename(columns = {0:'Time_Loaded', 1:'Time_UnLoaded'}, inplace=True) #renames the columns in the concatenated dataframes
    
    perc_time_loaded_unloaded = pd.concat([comp_air_outputs[comp_air_outputs>0].count(), comp_air_outputs[comp_air_outputs==0].count()], axis=1) #concatenate two series where one contains the aggregated count data in comp_air_outputs that is greater than 0 and the other contains aggregated count data that is equal to 0
    perc_time_loaded_unloaded['%_Time_Loaded'] = af.get_proportion(perc_time_loaded_unloaded[0], perc_time_loaded_unloaded[1]) #creates column with values calculated using get proportion function
    perc_time_loaded_unloaded['%_Time_UnLoaded'] = af.get_proportion(perc_time_loaded_unloaded[1], perc_time_loaded_unloaded[0]) #creates column with values calculated using get proportion function
    perc_time_loaded_unloaded.drop(columns = [0,1], inplace=True) #drops two columns
    
    power_loaded_unloaded = pd.concat([comp_power_outputs[comp_air_outputs>0].sum(), comp_power_outputs[comp_air_outputs==0].sum()], axis=1) #grabs the aggregated sum data in comp_power_outputs that are greater than 0 and equal to 0
    power_loaded_unloaded.rename(columns = {0:'Power_Loaded', 1:'Power_UnLoaded'}, inplace=True) #renames the columns in the concatenated dataframes
    
    perc_power_loaded_unloaded = pd.concat([comp_power_outputs[comp_air_outputs>0].sum(), comp_power_outputs[comp_air_outputs==0].sum()], axis=1) #concatenate two series where one contains the aggregated sum data in comp_power_outputs that is greater than 0 and the other contains aggregated sum data that is equal to 0
    perc_power_loaded_unloaded['%_Power_Loaded'] = af.get_proportion(perc_power_loaded_unloaded[0], perc_power_loaded_unloaded[1]) #creates column with values calculated using get proportion function
    perc_power_loaded_unloaded['%_Power_UnLoaded'] = af.get_proportion(perc_power_loaded_unloaded[1], perc_power_loaded_unloaded[0]) #creates column with values calculated using get proportion function
    perc_power_loaded_unloaded.drop(columns = [0,1], inplace=True) #drops two columns
    
    global fig, axes
    fig ,axes = plt.subplots(nrows=3, ncols=3, figsize=(15,60)) #creates canvas for subplots
    
    ax1=axes[0][0].twinx()
    average_power_air['Average_Air_Output'].plot(ax=ax1, kind='line', lw=1.5, color='darkblue', marker='s', ls='-',
                     markerfacecolor='darkblue', markersize=6, ylim=(0, max(average_power_air['Average_Air_Output'])*1.10))
    average_power_air['Average_Power_Output'].plot(ax=axes[0][0], kind='bar', edgecolor='black', lw=1, color='orange', ylim=(0, max(average_power_air['Average_Power_Output'])*1.10))
    axes[0][0].set_ylabel('kW')
    ax1.set_ylabel('m3/hr')
    axes[0][0].set_title(f'Average Power and Air Output ({name})', fontsize=18, fontweight='bold')
    axes[0][0].set_xlabel("")
    axes[0][0].get_yaxis().set_major_formatter(matplotlib.ticker.FuncFormatter(lambda y, p: format(int(y), ',')))
    ax1.get_yaxis().set_major_formatter(matplotlib.ticker.FuncFormatter(lambda y, p: format(int(y), ',')))
    axes[0][0].legend(loc=1)
    ax1.legend(loc=2)
    #creates subplot on the first axes with a barplot of the average power of each compressor and a line plot of the average air output for each compressor.
    
    ax2=axes[0][2].twinx()
    total_power_air['Total_Air_Output'].plot(ax=ax2, kind='line', lw=1.5, color='darkblue', marker='s', ls='-',
                     markerfacecolor='darkblue', markersize=6, ylim=(0, max(total_power_air['Total_Air_Output'])*1.10))
    total_power_air['Total_Power_Output'].plot(ax=axes[0][2], kind='bar', edgecolor='black', lw=1, color='orange', ylim=(0, max(total_power_air['Total_Power_Output'])*1.10))
    axes[0][2].set_ylabel('kW')
    ax2.set_ylabel('m3/hr')
    axes[0][2].set_title(f'Total Power and Air Output ({name})', fontsize=18, fontweight='bold')
    axes[0][2].set_xlabel("")
    axes[0][2].get_yaxis().set_major_formatter(matplotlib.ticker.FuncFormatter(lambda y, p: format(int(y), ',')))
    ax2.get_yaxis().set_major_formatter(matplotlib.ticker.FuncFormatter(lambda y, p: format(int(y), ',')))
    axes[0][2].legend(loc=1)
    ax2.legend(loc=2)
    #creates subplot on the second axes with a barplot of the total power of each compressor and a line plot of the total air output for each compressor.
    
    average_efficiency.plot(ax=axes[1][1], kind = 'bar', edgecolor='black', lw=1, color='purple')
    axes[1][1].set_ylabel('kW/m3/hr')
    axes[1][1].set_title(f'Average Efficiency ({name})', fontsize=18, fontweight='bold')
    axes[1][1].set_xlabel("")
    axes[1][1].set_xticklabels(list(average_efficiency.index),rotation=0)
    #creates subplot on the third axes with a barplot of the average efficiency of each compressor
    
    time_loaded_unloaded.plot(ax=axes[2][0], kind='bar', lw=1, color=['green', 'red'], stacked=True, edgecolor='black')
    axes[2][0].set_ylabel('hr')
    axes[2][0].set_title(f'Time Unload and Loaded ({name})', fontsize=18, fontweight='bold')
    axes[2][0].set_xlabel("")
    axes[2][0].set_xticklabels(list(time_loaded_unloaded.index),rotation=0)
    axes[2][0].get_yaxis().set_major_formatter(matplotlib.ticker.FuncFormatter(lambda y, p: format(int(y), ',')))
    #creates a subplot on the fourth axes with a barplot that is stacked of the time load and time unload for each compressor
    
    perc_time_loaded_unloaded.plot(ax=axes[1][0], kind='bar', lw=1, color=['green', 'red'], stacked=True, edgecolor='black')
    axes[1][0].set_ylabel('%')
    axes[1][0].set_title(f'%_Time Unload and Loaded ({name})', fontsize=18, fontweight='bold')
    axes[1][0].set_xlabel("")
    axes[1][0].set_xticklabels(list(perc_time_loaded_unloaded.index),rotation=0)
    #creates a subplot on the fifth axes with a barplot that is stacked of the percentage of time loaded and percentage of the time unload for each compressor
    
    power_loaded_unloaded.plot(ax=axes[2][2], kind='bar', lw=1, color=['green', 'red'], stacked=True, edgecolor='black')
    axes[2][2].set_ylabel('kW')
    axes[2][2].set_title(f'Power Unload and Loaded ({name})', fontsize=18, fontweight='bold')
    axes[2][2].set_xlabel("")
    axes[2][2].set_xticklabels(list(power_loaded_unloaded.index),rotation=0)
    axes[2][2].get_yaxis().set_major_formatter(matplotlib.ticker.FuncFormatter(lambda y, p: format(int(y), ',')))
    #creates a subplot on the sixth axes with a barplot that is stacked of the power loaded and power unload for each compressor
    
    perc_power_loaded_unloaded.plot(ax=axes[1][2], kind='bar', lw=1, color=['green', 'red'], stacked=True, edgecolor='black', rot=0)
    axes[1][2].set_ylabel('%')
    axes[1][2].set_title(f'% Power Unload and Loaded ({name})', fontsize=18, fontweight='bold')
    axes[1][2].set_xlabel("")
    axes[1][2].set_xticklabels(list(perc_power_loaded_unloaded.index),rotation=0)
    #creates a subplot on the seventh axes with a barplot that is stacked of the percentage of power loaded and percentage of the power unload for each compressor

    axes[2][1].axis('off') #removes axes from display
    axes[0][1].axis('off') #removes axes from display

    display_results()
    
def display_results():
    
    '''
    displays the results produced.
    '''
    global fig
    results_interface = tk.Toplevel() #creates an instance of a toplevel windows object
    gui.ShowResultsProgram(master=results_interface, fig=fig, main=main) #creates an instance ShowResultsProgram class on root_4 to display results of simulation
    
    global results_name
    global simulation_details
    del simulation_details
    del results_name
    
if __name__ == '__main__':
    main_root = tk.Tk() #creates main interface for program
    main = gui.MainInterface(master=main_root,
                             main_function = get_simulation_details) #creates instances of MainInterface for user to input number of simulations to be performed
    main_root.mainloop() #puts main_interface in the mainloop

'''
else: #num_of_simulations > 1
    
    simulations_info = [] #creates empty list to store simulation info
    for i in range(num_of_simulations):
        simulations_info.append({}) #adds a dictionary to store simulation info for each simulation
    
    simulation_num = 1
    while simulation_num<=num_of_simulations: #iterates over all simulations
        
        input_statement = f'How many compressor(s) are going to be in simulation {simulation_num}: ' 
        num_of_comp = af.get_num_comp(input_statement) #gets number of compressors into simulation to be analysed
        simulations_info[simulation_num-1]['num_of_comps'] = num_of_comp
        
        compressors_power_output_dict = {} #dictionary to stores lists of each compressor's power output results of simulation
        compressors_air_output_dict = {} #dictionary to stores lists of each compressor's air output results of simulation
        for j in range(num_of_comp):
            compressors_power_output_dict[f'C{j+1}_power_output'] = [] #adds list for each compressors power output to 'compressors_power_output_dict'
            compressors_air_output_dict[f'C{j+1}_air_output'] = [] #adds list for each compressors air output to 'compressors_air_output_dict'
        
        simulations_info[simulation_num-1]['compressors_power_output_dict'] = compressors_power_output_dict  #stores 'compressors_power_output_dict' for simulation i
        simulations_info[simulation_num-1]['compressors_air_output_dict'] = compressors_air_output_dict #stores 'compressors_air_output_dict' for simulation i
        
        print(f'\nChoose the maximum power(s), maximum air flow(s) and operation mode(s) for the compressor(s) in simulation {simulation_num}')
        
        compressors = [] #list to store compressor objects for simulation
        for j in range(num_of_comp): #iterates over num_of_comp
            input_statement = f"Maximum power capacity of compressor C{j+1}: "
            maximum_power = csf.get_compressor_power(input_statement) #gets the maximum power of a individual compressor
            input_statement = f"Maximum flow capacity of compressor C{j+1}: "
            maximum_flow =csf.get_compressor_capacity(input_statement) #gets the maximum air flows of a individual compressor
            C = csf.get_compressor(j+1, maximum_power, maximum_flow) #gets the type of compressor
            compressors.append(C) #appends to compressor list
            
        simulations_info[simulation_num-1]['compressors'] = compressors #stores 'compressors' for simulation i
        
        compressors_dict = {} #empty compressor dictionary
        compressors_power_dict = {} #empty compressor power output results dictionary
        compressors_air_dict = {} #empyty compressor air output results dictionary
        for j, compressor in enumerate(compressors): #iterates over compressor objects in compressor list
            compressors_dict[f'C{j+1}'] = compressor #dictionary to store compressor objects using string names for compressors.
            compressors_power_dict[f'C{j+1}'] = compressors_power_output_dict[f'C{j+1}_power_output'] #dictionary to store a list of string names for compressor power outputs and compressor power output list, using the compressor string name. 
            compressors_air_dict[f'C{j+1}'] = compressors_air_output_dict[f'C{j+1}_air_output'] #dictionary to store a list of string names for compressor air outputs and compressor air output list, using the compressor string name.
            
        simulations_info[simulation_num-1]['compressors_dict'] = compressors_dict #stores 'compressors_dict' for simulation i
        simulations_info[simulation_num-1]['compressors_power_dict'] = compressors_power_dict #stores 'compressors_power_dict' for simulation i
        simulations_info[simulation_num-1]['compressors_air_dict'] = compressors_air_dict #stores 'compressors_air_dict' for simulation i
            
        if len(compressors_dict) == 1: #if only one compressor is used in simulation
            trim = compressors_dict['C1'] #trim becomes that only compressor
        else:
            trim = csf.get_trim(compressors_dict)#asks user for trim compressor.
        
        simulations_info[simulation_num-1]['trim'] = trim #stores 'trim' for simulation i
            
        input_statement = f"Please choose whether the control scheme is optimised('opt') or defined('def') for simulation {simulation_num}: "
        control_sequence_type = csf.get_control_type(input_statement) #gets type of control sequence ('def' or 'opt')
        
        simulations_info[simulation_num-1]['control_sequence_type'] = control_sequence_type #stores 'control_sequence_type' for simulation i
        
        checking = True
        while checking == True:
            input_statement = f"Name and/or path of excel workbook with Compressed Air Production Data for simulation {simulation_num}: "
            compressed_air_data = csf.get_workbook(input_statement) #gets excel workbook object
            input_statement = f"Name of sheet with stacked compressed air volume data for simulation {simulation_num}: "
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
        
        simulations_info[simulation_num-1]['air_volume_stacked'] = air_volume_stacked #stores 'air_volume_stacked' for simulation i
        simulations_info[simulation_num-1]['coordinate'] = coordinate #stores 'coordinate' for simulation i
        
        simulation_num = simulation_num + 1
        continue
     
        
        #compressor_dict = {'C1':C1, 'C2':C2, 'C3':C3, 'C4':C4, 'C5':C5, 'C6':C6}
        
        #compressor_power_dict = {'C1':C1_power_output, 'C2':C2_power_output,
                                # 'C3':C3_power_output, 'C4':C4_power_output,
                                # 'C5':C5_power_output, 'C6':C6_power_output} #dictionary to store a list of string names for compressor power outputs and compressor power output list, using the compressor string name. 
    
        #compressor_air_dict = {'C1':C1_air_output, 'C2':C2_air_output,
                               #'C3':C3_air_output, 'C4':C4_air_output,
                               #'C5':C5_air_output, 'C6':C6_air_output} #dictionary to store a list of string names for compressor air outputs and compressor air output list, using the compressor string name.
        
    
    for i, info in enumerate(simulations_info): #iterates over simualtion info
        if info['control_sequence_type'] == 'def': #checks if simulation type is defined
            
            if len(info['compressors_dict']) == 1: #if only one compressor is in simulation
                setpoints = [0, info['compressors_dict']['C1'].maximum_flow] #list for storing setpoints
                contri_compressors = [[info['compressors_dict']['C1']]] #list to store compressor objects contributing to each setpoint
                boundaries = [(setpoints[0],setpoints[1])] #list to store the integer boundaries for the air_demand values
            else:
                print(f"\nNow asking for number of setpoints in simulation's {i+1} control scheme")
        
                setpoints = [0] #list for storing the integer setpoints
                contri_compressors = [] #list to store compressor objects contributing to each setpoint
                boundaries = [] #list to store the integer boundaries for the air_demand values
                
                input_statement = f"How many setpoints simulation's {i+1} in the control scheme: "
                n = csf.get_setpoints(input_statement) #function to get number of setpoints
                
                print(f"\nNow asking what compressor(s) maximum capacities are contributing to each setpoint in simulation's {i+1} control scheme")
                num = 1
                while num<=n:
                    try:
                        contri = input(f"Which compressor(s) are contributing to setpoint {num} (s{num}) in simulation's {i+1} control scheme: ") #asks user for the compressor(s) contributing to each setpoint in the control sequenc logic.
                        if contri == '0': #breaks from while loop if user types 0
                            break
                        lst_com = [] #list to store the contributing compressors for each setpoint
                        setpoint = 0 #variable to store setpoint
                        contri = [x.strip() for x in contri.upper().split(',')] #splits the contri input based on commas separating strings e.g. C1, C2, C3 => ['C1', 'C2', 'C3']
                        for compressor in contri: 
                            if len(compressor)>2: #checks if each compressor in contri is a valid input length input
                                exit()
                        for compressor in contri:
                            lst_com.append(compressors_dict[compressor]) #checks if input is in compressor_dict and if it is adds to lst_com 
                        contri_compressors.append(lst_com) #add lst_com to contri_compressors
                        for compressor in lst_com:
                            setpoint = setpoint + compressor.maximum_flow #add maximum flow of compressors in lst_com to setpoint
                        setpoints.append(setpoint) #adds setpoint to setpoints list.
                    except:
                        comp_num = info['num_of_comps'] #stores num_of_compressors from simulation as a variable
                        print('\nInvalid compressor object(s) chosen')
                        print(f'\nCompressor(s) chosen must be at least one of the compressors C1 - C{comp_num},\nmultiple compressors should be written with a "," between them\n')
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
        
            simulations_info[i]['setpoints'] = setpoints #stores 'setpoints' for simulation i
            simulations_info[i]['contri_compressors'] = contri_compressors #stores 'contri_compressors' for simulation i
            simulations_info[i]['boundaries'] = boundaries #stores 'boundaries' for simulation i
            
    print('\nPerforming simulations...')
    
    for i, info in enumerate(simulations_info):
        if info['control_sequence_type'] == 'opt': #checks if control_sequence type is optimisation
            
            compressors = simulations_info[i]['compressors'] #assigns comprssors variable from simulation info dictionary 
            trim = simulations_info[i]['trim'] #assigns trim variable from simulation info dictionary 
            coordinate = simulations_info[i]['coordinate'] #assigns coordinate variable from simulation info dictionary 
            air_volume_stacked = simulations_info[i]['air_volume_stacked'] #assigns air_volume_stacked variable from simulation info dictionary 
            
            compressors_power_array_dict = {}
            compressors_air_array_dict = {}
            compressors_air_output_opt_dict = {}
            compressors_power_output_opt_dict = {}
            for j, compressor in enumerate(compressors):
                compressors_power_array_dict[f'c{j+1}'] = np.array(csf.get_power_array(compressor, trim)) #1D array object of power_output possibilities for each compressor
                compressors_air_array_dict[f'c{j+1}'] = np.array(csf.get_air_array(compressor, trim)) #1D array object of air_output possibilities for each compressor
                compressors_air_output_opt_dict[f'c{j+1}_air_output'] = np.array([]) #empty air_output 1D array object for each compressor
                compressors_power_output_opt_dict[f'c{j+1}_power_output'] = np.array([]) #empty power_output 1D array object for each compressor
            
            total_air_output = np.array([]) #empty air_output 1D array object for total air outputs of compressors in compressed air system
            total_power_output = np.array([]) #empty power_output 1D array object for total power outputs of compressors in compressed air system
            
            power_lst = [compressors_power_array_dict[f'c{x+1}'] for x in range(len(compressors))] #list of array objects with power_output possibilities for each compressor
            air_lst = [compressors_air_array_dict[f'c{x+1}'] for x in range(len(compressors))] #list of array objects with air_output possibilities for each compressor
            prod_power = list(product(*power_lst)) #returns all possible combinations of power_outputs for each compressor as a tuple
            prod_air = list(product(*air_lst)) #returns all possible combinations of air_outputs for each compressor as a tuple
            
            comp_num = 1
            while comp_num<=len(compressors_power_output_opt_dict):
                for power_tup in prod_power: #iterates over power output tuples
                    compressors_power_output_opt_dict[f'c{comp_num}_power_output'] = np.append(compressors_power_output_opt_dict[f'c{comp_num}_power_output'], power_tup[comp_num-1]) #adds individual power output possibility to compressor array in 'compressors_power_output_opt_dict'
                comp_num = comp_num + 1
                
            comp_num = 1
            while comp_num<=len(compressors_air_output_opt_dict):
                for air_tup in prod_air: #iterates over power output tuples
                    compressors_air_output_opt_dict[f'c{comp_num}_air_output'] = np.append(compressors_air_output_opt_dict[f'c{comp_num}_air_output'], air_tup[comp_num-1]) #adds individual air output possibility to compressor array in 'compressors_air_output_opt_dict'
                comp_num = comp_num + 1
                
            for power_tup in prod_power: #iterates over tuples
                total_power_output = np.append(total_power_output, sum(power_tup)) #adds total power output possibility to 'total_power_output' array
                
            for air_tup in prod_air: #iterates over tuples
                total_air_output = np.append(total_air_output, sum(air_tup)) #adds total air output possibility to 'total_air_output' array
            
            power_output_array = np.vstack(tuple(compressors_power_output_opt_dict[f'c{x+1}_power_output'] for x in range(len(compressors_power_output_opt_dict)))) #creates a stack array of arrays                       
            air_output_array = np.vstack(tuple(compressors_air_output_opt_dict[f'c{x+1}_air_output'] for x in range(len(compressors_air_output_opt_dict)))) #creates a stack array of arrays 
            
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
                    
                    for j in range(len(simulations_info[i]['compressors_dict'])): #iterates over all 'C1_power_output' and 'C1_air_output' lists
                        simulations_info[i]['compressors_power_dict'][f'C{j+1}'].append(power_output_array[j][optimised_index]) #appends power output of each compressor to compressor list
                        simulations_info[i]['compressors_air_dict'][f'C{j+1}'].append(air_output_array[j][optimised_index]) #appends air output of each compressor to compressor list
    
        else:
            compressors = simulations_info[i]['compressors'] #assigns compressors variable from simulation info dictionary 
            trim = simulations_info[i]['trim'] #assigns trim variable from simulation info dictionary 
            coordinate = simulations_info[i]['coordinate'] #assigns coordinate variable from simulation info dictionary 
            air_volume_stacked = simulations_info[i]['air_volume_stacked'] #assigns air_volume_stacked variable from simulation info dictionary 
            boundaries = simulations_info[i]['boundaries'] #assigns boundaries variable from simulation info dictionary 
            setpoints = simulations_info[i]['setpoints'] #assigns setpoints variable from simulation info dictionary 
            contri_compressors = simulations_info[i]['contri_compressors'] #assigns contri_compressors variable from simulation info dictionary 
            
            for cell in air_volume_stacked[coordinate[0]]: #iterates over column with air_volume data in m3/hr
                if type(cell.value) == int or type(cell.value) == float: #checks if value in cell is an int or float
                    air_demand = round(cell.value) #stores data as air demand
                    active_compressors = csf.get_active_compressors(boundaries, air_demand, setpoints, contri_compressors, trim) #returns a list of active compressors
                    
                    for j in range(len(compressors_dict)): #iterates over all 'C1_power_output' and 'C1_air_output' lists
                        csf.update_power_output(simulations_info[i]['compressors_dict'][f'C{j+1}'], active_compressors, simulations_info[i]['compressors_power_dict'][f'C{j+1}'], air_demand) #appends power output of each compressor to compressor list
                        csf.update_air_output(simulations_info[i]['compressors_dict'][f'C{j+1}'], active_compressors, simulations_info[i]['compressors_air_dict'][f'C{j+1}'], air_demand) #appends air output of each compressor to compressor list


    df_power_outputs = [] #list for power output dataframes
    df_air_outputs = [] #list for air_output dataframes
    
    for info in simulations_info:
        df_power_outputs.append(pd.DataFrame(info['compressors_power_dict'])) #creates power_outputs dataframe from simulation info 
        df_air_outputs.append(pd.DataFrame(info['compressors_air_dict'])) #creates air_outputs dataframe from simulation info 
                
    
    for i, df in enumerate(df_power_outputs):
        df['Total'] = df[list(df.columns)].sum(axis=1) #creates a aggregate sum column for all compressors in all dataframes in df_power_outputs
        df.drop([f'C{n}' for n in range(1, len(simulations_info[i]['compressors_dict'])+1)], axis=1, inplace=True) #drops data for individual compressors in all dataframes in df_power_outputs
            

    for i, df in enumerate(df_air_outputs):
        df['Total'] = df[list(df.columns)].sum(axis=1) #creates a aggregate sum column for all compressors in all dataframes in df_air_outputs
        df.drop([f'C{n}' for n in range(1, len(simulations_info[i]['compressors_dict'])+1)], axis=1, inplace=True) #drops data for individual compressors in all dataframes in df_air_outputs
    
    simulation_names = [] #empty list for operation mode names
    simulation_num = 1
    while simulation_num<=num_of_simulations: #iterates over all operation mode combinations
        input_statement=f"Name to be given for simulation {simulation_num}'s results: "
        name = af.get_op_name(input_statement) #get name for each operation mode combination
        simulation_names.append(name) #add name to op_name list
        df_power_outputs[simulation_num-1].rename(columns={'Total':name}, inplace=True) #renames column name to operation mode combination name in each dataframe
        df_air_outputs[simulation_num-1].rename(columns={'Total':name}, inplace=True) #renames column name to operation mode combination name in each dataframe
        simulation_num = simulation_num+1
        continue   
    
    print("\nGenerating results...")    
    
    concat_power_outputs = pd.concat(df_power_outputs, axis=1) #concatenates the control sequences dfs in the df_power_outputs together
    concat_air_outputs = pd.concat(df_air_outputs, axis=1) #concatenates the control sequences dfs in the df_power_outputs together
    
    average_power = round(concat_power_outputs.mean()) #applies mean aggregate function to comp_power_outputs dataframe
    
    average_air = round(concat_air_outputs.mean()) #applies mean aggregate function to comp_air_outputs dataframe
    
    average_power_air = pd.concat([average_power, average_air], axis=1) #concatenates the average_power and average_air dataframes
    average_power_air.rename(columns = {0:'Average_Power_Output', 1:'Average_Air_Output'}, inplace=True) #renames the columns in the concatenated dataframes
    
    total_power = concat_power_outputs.sum() #applies sum aggregate function to comp_power_outputs dataframe
    
    total_air = concat_air_outputs.sum() #applies sum aggregate function to comp_air_outputs dataframe
    
    total_power_air = pd.concat([total_power, total_air], axis=1) #concatenates the total_power and total_air dataframes
    total_power_air.rename(columns = {0:'Total_Power_Output', 1:'Total_Air_Output'}, inplace=True) #renames the columns in the concatenated dataframes
    
    average_efficiency = (concat_power_outputs.mean())/(concat_air_outputs.mean()) #applies mean aggregate function to comp_power_outputs dataframe and mean aggregate function to comp_air_outputs dataframe then divides them.
    
    time_loaded_unloaded = pd.concat([concat_air_outputs[concat_air_outputs>0].count(), concat_air_outputs[concat_air_outputs==0].count()], axis=1) #concatenate two series where one contains the aggregated count data in comp_air_outputs that is greater than 0 and the other contains aggregated count data that is equal to 0
    time_loaded_unloaded.rename(columns = {0:'Time_Loaded', 1:'Time_UnLoaded'}, inplace=True) #renames the columns in the concatenated dataframes
    
    perc_time_loaded_unloaded = pd.concat([concat_air_outputs[concat_air_outputs>0].count(), concat_air_outputs[concat_air_outputs==0].count()], axis=1) #concatenate two series where one contains the aggregated count data in comp_air_outputs that is greater than 0 and the other contains aggregated count data that is equal to 0
    perc_time_loaded_unloaded['%_Time_Loaded'] = af.get_proportion(perc_time_loaded_unloaded[0], perc_time_loaded_unloaded[1]) #creates column with values calculated using get proportion function
    perc_time_loaded_unloaded['%_Time_UnLoaded'] = af.get_proportion(perc_time_loaded_unloaded[1], perc_time_loaded_unloaded[0]) #creates column with values calculated using get proportion function
    perc_time_loaded_unloaded.drop(columns = [0,1], inplace=True) #drops two columns
    
    power_loaded_unloaded = pd.concat([concat_power_outputs[concat_air_outputs>0].sum(), concat_power_outputs[concat_air_outputs==0].sum()], axis=1) #grabs the aggregated sum data in comp_power_outputs that are greater than 0 and equal to 0
    power_loaded_unloaded.rename(columns = {0:'Power_Loaded', 1:'Power_UnLoaded'}, inplace=True) #renames the columns in the concatenated dataframes
    
    perc_power_loaded_unloaded = pd.concat([concat_power_outputs[concat_air_outputs>0].sum(), concat_power_outputs[concat_air_outputs==0].sum()], axis=1) #concatenate two series where one contains the aggregated sum data in comp_power_outputs that is greater than 0 and the other contains aggregated sum data that is equal to 0
    perc_power_loaded_unloaded['%_Power_Loaded'] = af.get_proportion(perc_power_loaded_unloaded[0], perc_power_loaded_unloaded[1]) #creates column with values calculated using get proportion function
    perc_power_loaded_unloaded['%_Power_UnLoaded'] = af.get_proportion(perc_power_loaded_unloaded[1], perc_power_loaded_unloaded[0]) #creates column with values calculated using get proportion function
    perc_power_loaded_unloaded.drop(columns = [0,1], inplace=True) #drops two columns
    
    fig ,axes = plt.subplots(nrows=7, ncols=1, figsize=(15,60))
    
    ax1=axes[0].twinx()
    average_power_air['Average_Air_Output'].plot(ax=ax1, kind='line', lw=1.5, color='darkblue', marker='s', ls='-',
                     markerfacecolor='darkblue', markersize=6, ylim=(0, max(average_power_air['Average_Air_Output'])*1.10))
    average_power_air['Average_Power_Output'].plot(ax=axes[0], kind='bar', edgecolor='black', lw=1, color='orange', ylim=(0, max(average_power_air['Average_Power_Output'])*1.10))
    axes[0].set_ylabel('kW')
    ax1.set_ylabel('m3/hr')
    axes[0].set_title(f'Average Power and Air Output ({name})', fontsize=18, fontweight='bold')
    axes[0].set_xlabel("")
    axes[0].get_yaxis().set_major_formatter(matplotlib.ticker.FuncFormatter(lambda y, p: format(int(y), ',')))
    ax1.get_yaxis().set_major_formatter(matplotlib.ticker.FuncFormatter(lambda y, p: format(int(y), ',')))
    axes[0].legend(loc=1)
    ax1.legend(loc=2)
    #creates subplot on the first axes with a barplot of the average power of each compressor and a line plot of the average air output for each compressor.
    
    ax2=axes[1].twinx()
    total_power_air['Total_Air_Output'].plot(ax=ax2, kind='line', lw=1.5, color='darkblue', marker='s', ls='-',
                     markerfacecolor='darkblue', markersize=6, ylim=(0, max(total_power_air['Total_Air_Output'])*1.10))
    total_power_air['Total_Power_Output'].plot(ax=axes[1], kind='bar', edgecolor='black', lw=1, color='orange', ylim=(0, max(total_power_air['Total_Power_Output'])*1.10))
    axes[1].set_ylabel('kW')
    ax2.set_ylabel('m3/hr')
    axes[1].set_title(f'Total Power and Air Output ({name})', fontsize=18, fontweight='bold')
    axes[1].set_xlabel("")
    axes[1].get_yaxis().set_major_formatter(matplotlib.ticker.FuncFormatter(lambda y, p: format(int(y), ',')))
    ax2.get_yaxis().set_major_formatter(matplotlib.ticker.FuncFormatter(lambda y, p: format(int(y), ',')))
    axes[1].legend(loc=1)
    ax2.legend(loc=2)
    #creates subplot on the second axes with a barplot of the total power of each compressor and a line plot of the total air output for each compressor.
    
    average_efficiency.plot(ax=axes[2], kind = 'bar', edgecolor='black', lw=1, color='purple')
    axes[2].set_ylabel('kW/m3/hr')
    axes[2].set_title(f'Average Efficiency ({name})', fontsize=18, fontweight='bold')
    axes[2].set_xlabel("")
    axes[2].set_xticklabels(list(average_efficiency.index),rotation=0)
    axes[2].get_yaxis().set_major_formatter(matplotlib.ticker.FuncFormatter(lambda y, p: format(int(y), ',')))
    #creates subplot on the third axes with a barplot of the average efficiency of each compressor
    
    time_loaded_unloaded.plot(ax=axes[3], kind='bar', lw=1, color=['green', 'red'], stacked=True, edgecolor='black')
    axes[3].set_ylabel('hr')
    axes[3].set_title(f'Time Unload and Loaded ({name})', fontsize=18, fontweight='bold')
    axes[3].set_xlabel("")
    axes[3].set_xticklabels(list(time_loaded_unloaded.index),rotation=0)
    axes[3].get_yaxis().set_major_formatter(matplotlib.ticker.FuncFormatter(lambda y, p: format(int(y), ',')))
    #creates a subplot on the fourth axes with a barplot that is stacked of the time load and time unload for each compressor
    
    perc_time_loaded_unloaded.plot(ax=axes[4], kind='bar', lw=1, color=['green', 'red'], stacked=True, edgecolor='black')
    axes[4].set_ylabel('%')
    axes[4].set_title(f'%_Time Unload and Loaded ({name})', fontsize=18, fontweight='bold')
    axes[4].set_xlabel("")
    axes[4].set_xticklabels(list(perc_time_loaded_unloaded.index),rotation=0)
    #creates a subplot on the fifth axes with a barplot that is stacked of the percentage of time loaded and percentage of the time unload for each compressor
    
    power_loaded_unloaded.plot(ax=axes[5], kind='bar', lw=1, color=['green', 'red'], stacked=True, edgecolor='black')
    axes[5].set_ylabel('kW')
    axes[5].set_title(f'Power Unload and Loaded ({name})', fontsize=18, fontweight='bold')
    axes[5].set_xlabel("")
    axes[5].set_xticklabels(list(power_loaded_unloaded.index),rotation=0)
    axes[5].get_yaxis().set_major_formatter(matplotlib.ticker.FuncFormatter(lambda y, p: format(int(y), ',')))
    #creates a subplot on the sixth axes with a barplot that is stacked of the power loaded and power unload for each compressor
    
    perc_power_loaded_unloaded.plot(ax=axes[6], kind='bar', lw=1, color=['green', 'red'], stacked=True, edgecolor='black', rot=0)
    axes[6].set_ylabel('%')
    axes[6].set_title(f'% Power Unload and Loaded ({name})', fontsize=18, fontweight='bold')
    axes[6].set_xlabel("")
    axes[6].set_xticklabels(list(perc_power_loaded_unloaded.index),rotation=0)
    #creates a subplot on the seventh axes with a barplot that is stacked of the percentage of power loaded and percentage of the power unload for each compressor
    
    fig.tight_layout()
    
    print('\nResults generated...')
    
    input_statement = 'Name of filename to save figures as: '
    af.save_results(input_statement) #saves results as a pdf
    
    print('\nProgram finished')
'''
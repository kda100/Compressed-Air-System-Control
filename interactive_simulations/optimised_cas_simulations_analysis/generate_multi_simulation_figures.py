# -*- coding: utf-8 -*-
"""
Created on Mon Aug  3 11:10:15 2020
**************************************************************************************************************************************************
INTRODUCTION:
  
This program takes in multiple excel documents containing the power_output and air_output results from compressed air system simulations then delivers 
figures of performance metrics, such as the Average Power, Average Air Output, Average Efficiency and more, of the compressors in the simulations, 
that can be used to compare and analyse the compressor operation mode combinations and the compressor control sequencing logics used in the simulations 

**************************************************************************************************************************************************
DESCRIPTION:

First the program will ask the user for the number of compressors that were used in the compressed air system program (for the results in the 
'defined_simulation.py' and 'optimised_simulation.py' scripts, this will be 6). Then the program will ask the user for the of simulations to be analysed.

Then for each simulation (operation mode 1 and control sequence 1) the program will ask the user for the 'C:\path\excelworkbook.xlsx' that contains 
the results for the power outputs and air outputs of the compressors, then ask the user for the sheet name that contains the power outputs and sheet 
name that contains the air outputs of the compressors and load it into separate dataframes. Ensure the column names are kept in the same format the 
'defined_simulation.py' or the 'optimised_simulation.py' simulations produces them in and that the column titles ('C1_Power_Output...' and 
'C1_Air_Output...') are kept the same.

The program will then do some formatting on the dataframes to produce a series of subplots by aggregating the results of the compressors in each 
simulation uploaded from  the excel workbooks. The program will then ask the user to name the simulations (simulation 1 to cs1, op1), then
barplots of the performance metrics are created for analysis and comparisons of the simulations. Finally the program will ask the user to name the 
path/name in the format of 'C:\path\filename' to save the subplots as and save the them as a pdf.

**************************************************************************************************************************************************
@author: kda_1
"""
import pandas as pd
import cas_modules.analysis_functions as af
import matplotlib.pyplot as plt
import matplotlib

print("\nAt any point to exit the program type 0")

input_statement = 'How many compressors were in the compressed air system results to be analysed: '
num_of_comp = af.get_num_compressors(input_statement) #gets number of compressors into simulation to be analysed

input_statement = 'Number of simulation results to be analysed: '
num_of_simulations = af.get_num_simulations(input_statement) #gets the number of simulations for analysis

df_power_outputs = [] #list for power output dataframes
df_air_outputs = [] #list for air_output dataframes

simulations_num = 1
while simulations_num<=num_of_simulations: #iterates over all simulations
    input_statement = f"Name and/or path of excel workbook {simulations_num} with power output and air_output results: "
    excel_workbook = af.get_excel_workbook(input_statement) #gets excel workbook that contains power_output and air_output data to be loaded in pd.Dataframe
    input_statement = f"Name of sheet in excel workbook {simulations_num} with power output results: "
    sheet_name1 = af.get_sheet_name(input_statement) #gets excel sheet name for power_output data to be loaded into pd.Dataframe
    input_statement = f"Name of sheet in excel workbook {simulations_num} with air output results: "
    sheet_name2 = af.get_sheet_name(input_statement) #gets excel sheet name for air_output data to be loaded into pd.Dataframe
    try:
        print(f'\nUploading data in excelworkbook {simulations_num}...')
        df_power_outputs.append(pd.read_excel(excel_workbook +'.xlsx', sheet_name = sheet_name1)) #appends power output simulation dataframe to the df_power_outputs list 
        df_air_outputs.append(pd.read_excel(excel_workbook +'.xlsx', sheet_name = sheet_name2)) #appends air output simulation dataframe to the df_air_outputs list
        print(f'\nData in excelworkbook {simulations_num} uploaded')
    except PermissionError:
        print('\nEnsure the document is closed\n')
        continue
    except:
        print('Invalid excel workbook or sheetname(s)')
        continue
    else:
        simulations_num = simulations_num+1 #goes to next simulation
        continue

for index, simulation in enumerate(df_power_outputs):
    df_power_outputs[index] = simulation[[f'C{n}_power_output' for n in range(1, num_of_comp)]] #grabs column names 'C1_power_output, C2_power_output...' in comp_power_outputs dataframe 
    df_power_outputs[index]['Total'] = simulation[list(simulation.columns)].sum(axis=1) #creates new column by summing all other columns
    df_power_outputs[index].drop([f'C{n}_power_output' for n in range(1, num_of_comp)], axis=1, inplace=True) #drops columns with individual compressor data

for index,simulation in enumerate(df_air_outputs):
    df_air_outputs[index] = simulation[[f'C{n}_air_output' for n in range(1, num_of_comp)]] #grabs column names 'C1_air_output, C2_air_output...' in comp_air_outputs dataframe
    df_air_outputs[index]['Total'] = simulation[list(simulation.columns)].sum(axis=1) #creates new column by summing all other columns
    df_air_outputs[index].drop(labels=[f'C{n}_air_output' for n in range(1, num_of_comp)], axis=1, inplace=True) #drops columns with individual compressor data

sim_names = [] #creates list of simulation names
simulations_num = 1
while simulations_num<=num_of_simulations: #iterates over all simulations
        input_statement=f"Name to be given for simulation {simulations_num}'s results for each control scheme: "
        name = af.get_simulation_name(input_statement) #gets names for each operation mode
        sim_names.append(name) #appends name to list
        for simulation in range(num_of_simulations):
            df_power_outputs[simulation].rename(columns={'Total':name}, inplace=True) #renames 'Total' column in df_power_output for the given operation mode in each control sequence to name given
            df_air_outputs[simulation].rename(columns={'Total':name}, inplace=True) #renames 'Total' column in df_air_output for the given operation mode in each control sequence to name given
        simulations_num = simulations_num+1 #goes to next simulation
        continue

print('\nGenerating results...')

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

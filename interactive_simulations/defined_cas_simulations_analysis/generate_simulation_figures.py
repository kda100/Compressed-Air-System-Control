# -*- coding: utf-8 -*-
"""
Created on Mon Aug  3 11:10:15 2020
**************************************************************************************************************************************************
INTRODUCTION:
  
This program takes in an excel document containing the power_output and air_output results from a compressed air system simulation then delivers 
figures of performance metrics of compressors in the simulation that can be used to analyse the results such as the Average Power, Average Air 
Output, Average Efficiency and more.

**************************************************************************************************************************************************
DESCRIPTION:

First the program will ask the user for the number of compressors that were used in the compressed air system simulations (for the results in the 
'defined_simulation.py' and the 'optimised_simulation.py' scripts, this will be 6).

Then the program will ask the user for the 'C:\path\excelworkbook.xlsx' that contains the results of the power outputs and the air outputs for the 
compressors in the simulation. Then, the program will ask the user for the sheet name that contains the power outputs and the sheet name that contains 
the air outputs of the compressors and load it into separate dataframes. Ensure the column names are kept in the same format the 'defined_simulation.py' 
or the 'optimised_simulation.py' simulations produces them in and that the column titles ('C1_Power_Output...' and 'C1_Air_Output...') are kept the same. 

Finally, the program will do some work on the dataframes and produce a series of subplots on the performance metrics of the individual compressors in 
the simulation and ask the user to name the graphs in the subplots (e.g. op1, cs1), before asking the user the path/name in the format 
of 'C:\path\filename' to save the subplots as and save the them as a pdf.

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

while True:
    input_statement = f"Name and/or path of excel workbook with power output and air_output results: "
    excel_workbook = af.get_excel_workbook(input_statement) #gets excel workbook that contains power_output and air_output data to be loaded in pd.Dataframe
    input_statement = f"Name of sheet in excel workbook with power output results: "
    sheet_name1 = af.get_sheet_name(input_statement) #gets excel sheet name for power_output data to be loaded into pd.Dataframe
    input_statement = f"Name of sheet in excel workbook with air output results: "
    sheet_name2 = af.get_sheet_name(input_statement) #gets excel sheet name for air_output data to be loaded into pd.Dataframe
    try:
        print('\nUploading data...')
        comp_power_outputs = pd.read_excel(excel_workbook +'.xlsx', sheet_name = sheet_name1) #loads the power_output data in sheet_name1 and excel workbook to a pandas dataframe
        comp_air_outputs = pd.read_excel(excel_workbook +'.xlsx', sheet_name = sheet_name2) #loads the air_output data in sheet_name2 and excel workbook to a pandas dataframe
        print('\nData uploaded')
    except PermissionError:
        print('\nEnsure the document is closed\n')
        continue
    except:
        print('Invalid excel workbook or sheetname(s)')
        continue
    else:
        break #breaks out of while loop

input_statement = "Name the figures to be produced: "
name = af.get_figures_name(input_statement) #gets name of figures

print('\nGenerating results...')

comp_power_outputs = comp_power_outputs[[f'C{n}_power_output' for n in range(1, num_of_comp+1)]] #grabs column names 'C1_power_output, C2_power_output...' in comp_power_outputs dataframe 

comp_air_outputs = comp_air_outputs[[f'C{n}_air_output' for n in range(1, num_of_comp+1)]] #grabs column names 'C1_air_output, C2_air_output...' in comp_air_outputs dataframe

comp_power_outputs.columns = [f'C{n}' for n in range(1, num_of_comp+1)] #changes columns names in comp_power_outputs dataframes to 'C1, C2...'
comp_air_outputs.columns = [f'C{n}' for n in range(1, num_of_comp+1)] #changes columns names in comp_air_outputs dataframes to 'C1, C2...'

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

fig ,axes = plt.subplots(nrows=7, ncols=1, figsize=(15,60)) #creates canvas for subplots

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

print('\nResults generated')

input_statement = 'Name of filename to save figures as: '
af.save_results(input_statement) #saves results as a pdf

print('\nProgram finished')

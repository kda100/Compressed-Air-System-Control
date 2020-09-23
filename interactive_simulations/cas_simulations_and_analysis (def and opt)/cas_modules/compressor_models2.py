# -*- coding: utf-8 -*-
"""
Created on Tue Feb 18 16:17:11 2020

This script contains the compressor class objects used in the compressed air system simulations, they describe the 
control of four different types of compressors. There is a 'Compressor' class object that contains the attributes: maximum 
power, minimum power, maximum flow and minimum flow of the 'Compressor' object. There are two sub-classes of the 'Compressor' 
object: 'FixedControl' and 'VariableControl' that inherit the attributes of the 'Compressor' class object and contains 
'power_output' and 'air_output' function calls that are used to produce a power output and air output reading:

'FixedControl' object - has two-states of control, which can have the compressor either operating at its maximum power output producing 
its maximum air output or operating at its minimum power output producing its minimum air output.

'VariableControl' object - is a continuous mode of control; it can vary its air output between its maximum air output and its minimum air output between its maximum power
output and its minimum power output.

****************************************************************************************************************************************
FUNCTIONS:
power_output - takes in a float/integer air_demand and list of active_compressors arguments that are determined by the compressor control sequencing logic.

air_output - takes in a float/integer air_demand and list of active_compressors arguments that are determined by the compressor control sequencing logic.


****************************************************************************************************************************************

There are a further two sub-classes for the 'FixedControl' object and two sub-classes for the 'VariableControl' which are the different types 
of operation modes used in the simulations:

'OnOff' object - Inherits the 'FixedControl' subclass. Compressor objects with this operation mode operates at its maximum 
power output when producing anything between its maximum air output or less and operates at 0% of its maximum power output 
when not producing any air output (0 m3/hr).

'LoadUnload' object - Inherits the 'FixedControl' subclass. Compressor objects with this operation mode operates at its maximum power 
output when producing anything between its maximum air output or less and operates at 50% of its maximum power output when 
not producing any air output (0 m3/hr).

'InletModulation' object - Inherits the 'VariableControl' subclass. Compressor objects with this operation mode can vary its 
air output anywhere between its maximum air output and minimum air output with the power output varying linearly 
between its maximum power output and 70% of its maximum power output.

'VariableSpeed' object - Inherits the 'VariableControl' subclass. Compressor objects with this operation mode can vary its air output 
anywhere between its maximum air output and minimum air output with the power output varying linearly 
between its maximum power output and 0% of its maximum power output.
"""    

#create Compressor class object
class Compressor():
    def __init__(self, maximum_power, minimum_power, maximum_flow, minimum_flow):
        self.maximum_power = maximum_power #maximum power attribute
        self.minimum_power = minimum_power #minimum power attribute
        self.maximum_flow = maximum_flow #maximum power attribute
        self.minimum_flow = minimum_flow #maximum power attribute

#create FixedControl subclass
class FixedControl(Compressor): #two-states of control

    #function to produce power output reading for FixedControl
    def power_output(self, air_demand, active_compressors):
        '''
        Takes in two arguments: air demand and active_compressors and returns the resulting power output for the compressor object
        '''
        if self in active_compressors: 
            return self.maximum_power #compressor in active compressors list, then it is producing compressed air and the return its maximum power
        else:
            return self.minimum_power # otherwise compressor is producing minmum compressed air and operating at minimum power
    
    #function to produce air output reading for FixedControl
    def air_output(self, air_demand, active_compressors):
        '''
        Takes in two arguments: air demand and active_compressors and returns the resulting air output for the compressor object
        '''
        if self in active_compressors: #compressor in active compressors list, then it is producing compressed air. 
            if self == active_compressors[0]: #compressor is first compressor in list
                if air_demand<self.maximum_flow: #air demand is less than maximum flow of first compressor (only one compressor in active compressors)
                    return air_demand #compressor delivers the air demanded 
                else:
                    return self.maximum_flow #air demqnd greater then maximum flow of compressor return (active compressors contain more than one compressor)
            elif self == active_compressors[-1]: #compresssor is last compressor in active compressors
                air_supply = 0
                for compressor in active_compressors:
                    if compressor in active_compressors[:active_compressors.index(self)]:
                        air_supply = air_supply + compressor.maximum_flow #sum of the maximum flow of all compressors before last compressor in active compressor
                return air_demand - air_supply #last compressor delivers difference between the sum and air demand
            else:
                return self.maximum_flow #all compressor between first and last compressor in list return their maximum flow
        else:
            return self.minimum_flow #other compressors not in list return their minimum flow

    def discrete_air_output(self):
        '''
        Returns a list of individual air output values between the compressor objects minimum and maximum air output incremented by 1
        '''
        discrete_air = []
        for air_flow in range(0, round(self.maximum_flow)+1):
            discrete_air.append(air_flow) 
        return discrete_air
    
    def discrete_power_output(self):
        '''
        Returns a list of individual power output values for each air output value between the compressor objects minimum and maximum power
        output based on a 2-state mode of operation.
        '''
        discrete_power = [0]
        for air_flow in range(1, round(self.maximum_flow)+1):
            discrete_power.append(self.maximum_power) #for each air_flow value the compressor produces the compressor operates at its maximum power output.
        return discrete_power
    
#create VariableControl subclass
class VariableControl(Compressor):
    
    #function to produce power output reading for VariableControl
    def power_output(self, air_demand, active_compressors):
        '''
        Takes in two arguments: air demand and active_compressors and returns the resulting power output for the compressor object
        '''
        if self in active_compressors: #compressor in active compressors list, then it is producing compressed air. 
            if active_compressors[0] == self: #compressor is first compressor in list
                if air_demand < self.maximum_flow: #air demand is less than maximum flow of first compressor (only one compressor in active compressors)
                    return (((self.maximum_power)-(self.minimum_power))/((self.maximum_flow)-(self.minimum_flow))*air_demand)+self.minimum_power #compressor deliver the air demanded and the power output is determined by a linear relationship between the air output and the power output. 
                else:
                    return self.maximum_power #otherwise the air demand is the maximum_flow and the compressor operates at its maximum air output.
            elif active_compressors[-1] == self: #compresssor is last compressor in active compressors
                air_supply = 0
                for compressor in active_compressors:
                    if compressor in active_compressors[:active_compressors.index(self)]:
                        air_supply = air_supply + compressor.maximum_flow #sum of the maximum flow of all compressors before last compressor in active compressor
                return ((((self.maximum_power)-(self.minimum_power))/((self.maximum_flow)-(self.minimum_flow)))*(air_demand-air_supply))+self.minimum_power #last compressor delivers difference between the sum and air demand and power output is determined by the linear relationship.
            else:
                return self.maximum_power #all compressor between first and last compressor in list return their maximum power
        else:
        	return self.minimum_power #other compressors not in list return their minimum power
    
    #function to produce power output reading for VariableControl    
    def air_output(self, air_demand, active_compressors):
        '''
        Takes in two arguments: air demand and active_compressors and returns the resulting air output for the compressor object
        '''
        if self in active_compressors: #compressor in active compressors list, then it is producing compressed air. 
            if self == active_compressors[0]: #compressor is first compressor in list
                if air_demand<self.maximum_flow: #air demand is less than maximum flow of first compressor (only one compressor in active compressors)
                    return air_demand #compressor delivers the air demanded 
                else:
                    return self.maximum_flow #air demqnd greater then maximum flow of compressor return (active compressors contain more than one compressor)
            elif self == active_compressors[-1]: #compresssor is last compressor in active compressors
                air_supply = 0
                for compressor in active_compressors:
                    if compressor in active_compressors[:active_compressors.index(self)]:
                        air_supply = air_supply + compressor.maximum_flow #sum of the maximum flow of all compressors before last compressor in active compressor
                return air_demand - air_supply #last compressor delivers difference between the sum and air demand
            else:
                return self.maximum_flow #all compressor between first and last compressor in list return their maximum flow
        else:
            return self.minimum_flow #other compressors not in list return their minimum flow
    
    def discrete_air_output(self):
        '''
        Returns a list of individual air output values between the compressor objects minimum and maximum air output incremented by 1
        '''
        discrete_air = []
        for air_flow in range(0, round(self.maximum_flow)+1):
            discrete_air.append(air_flow)
        return discrete_air
    
    def discrete_power_output(self):
        '''
        Returns a list of individual power output values for each air output value between the compressor objects minimum and maximum power
        output based on a continuous mode of operation.
        '''
        discrete_power = []
        for air_flow in range(0, round(self.maximum_flow)+1):
            discrete_power.append(round(((((self.maximum_power)-(self.minimum_power))/((self.maximum_flow)-(self.minimum_flow)))*air_flow)+self.minimum_flow)) #for each air_flow value the compressor produces the power_output is determined by the linear relationship.
        return discrete_power

#create OnOff subclass
class OnOff(FixedControl): #2-state mode of control
    def __init__(self, maximum_power, maximum_flow):
        FixedControl.__init__(self, maximum_power, maximum_power*0, maximum_flow, maximum_flow*0) #minimum_power and minimum_flow is 0.

#create LoadUnload subclass
class LoadUnload(FixedControl): #2-state mode of control
    def __init__(self, maximum_power, maximum_flow):
        FixedControl.__init__(self, maximum_power, maximum_power*0.5, maximum_flow, maximum_flow*0) #minimum_power is 50% of maximum power and minimum_flow is 0.

#create InletModulation subclass              
class InletModulation(VariableControl): #continuous mode of control
    def __init__(self, maximum_power, maximum_flow):
        VariableControl.__init__(self, maximum_power, maximum_power*0.7, maximum_flow, maximum_flow*0) #minimum_power is 50% of maximum power and minimum_flow is 0.

#create VariableSpeed subclass
class VariableSpeed(VariableControl): #continuous mode of control
    def __init__(self, maximum_power, maximum_flow):
        VariableControl.__init__(self, maximum_power, maximum_power*0, maximum_flow, maximum_flow*0)#minimum_power and minimum_flow is 0.
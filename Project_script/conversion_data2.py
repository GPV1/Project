import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.styles import Font
import pandas as pd
import numpy as np
import pprint
import math
import sqlite3
from collections import defaultdict
from pandas import ExcelWriter

# todo check at the end if all the above imports are needed.


# import datafiles and sheetname. Necessary: database files PEG and Species order and input data
dfBiovolume = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/PEG_BVOL2019_PJ.xlsx', sheet_name='Biovolume file')
dfSpeciesOrder = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/Species_order_GPV.xlsx', sheet_name='Sheet1')
InputData = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/Data_2018_GPV.xlsx', sheet_name='Sheet1')
test_data = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/test_data.xlsx', sheet_name='data')


''' preperation of database by forming dictionaries for taxonomy groups, creating averages of multiple biovolume
 values per species and usage of dictionaries in functions'''

#  defines the columns in the database for the taxonomy groups
Divisions = dfBiovolume['Division']
Classes = dfBiovolume['Class']
Order = dfBiovolume['Order'].str.title()
Genus = dfBiovolume['Genus']
Species = dfBiovolume['Species']
Species_other = dfSpeciesOrder['spec_name'].str.replace('_', ' ')
Species_other_order = dfSpeciesOrder['order']
Species_other_group = dfSpeciesOrder['group']
All_Species = pd.concat([pd.Series(Species), pd.Series(Species_other)])

InputData['spec_name'] = InputData['spec_name'].str.replace('.', '')
InputData['spec_name'] = InputData['spec_name'].str.replace('<', '')
InputData['spec_name'] = InputData['spec_name'].str.replace('>', '')
InputData['spec_name'] = InputData['spec_name'].str.replace('~', '')

#todo below needed?

#  defines dictionaries of different levels in taxonomy groups based on the PEG database
# SpGe = dict(zip(Species, Genus)) # species on genus level
# SpOr = dict(zip(Species, Order)) # species on order level
#
'----------------------------------------------------------------------------------------------------------------------'
''' Overview of all dictonaries and their functions:

- 1. species_Biovolume_dict:                    Species from the PEG database with their corresponding biovolume values 
                                                (multiple values per species). 
                                                * {Keys: Species, values: biovolume}.
                                     
- 2. species_avg_Biovolume_dict:                Average biovolume for species from the PEG database. 
                                                * {Keys: Species, values: average biovolume per species}.

- 3. Order_Species_avg_Biovolume_dict:          The average biovolume per species is nested in the corresponding order. 
                                                In this dictionary also the species from the database SpeciesOrder are 
                                                added (other_species_dict) since these species are not found in the PEG
                                                database (The biovolume for these species can only be found on order 
                                                level).
                                                * {Keys: Order {Keys: Species, values: biovolume}}.
                                    
- 4. Order_avg_Biovolume_dict:                  For each order the average biovolume is based on the species values 
                                                inside this order.
                                                * { Keys: order, values: average biovolume per order}
                                    
- 5. Class_Species_avg_Biovolume:               Averaged biovolume per species nested in the corresponding taxanomy 
                                                group 'class'.
                                                * {Keys: Class {Keys: species, values: biovolume per species}}
                                                
- 6. Class_order_avg_Biovolume:                 Averaged biovolume per order nested in the corresponding taxanomy group
                                                'class'.
                                                * {Keys: Class {Keys: Order, values: biovolume per Order}}

- 7. Class_avg_Biovolume:                       Averaged biovolume per class based on the corresponding species.
                                                * {keys: Class, value: biovolume per class}
                                        
- 8. Group_Order_Species_avg_Biovolume_dict:    The functional groups (diatom, dionflagellates, flagellates and other) 
                                                are added to the different taxonomy order.
                                                * {Keys: group {keys: order, values: biovolume per species}}

'''
'----------------------------------------------------------------------------------------------------------------------'

species_Biovolume_dict = {}
''' 1. Creates a dictionary from the PEG database based on the species and the corresponding biovolume (µm3/L). This 
    dictionary will later be used to average the biovolume per species.
    
    - Defines the variable searchFor (the species name) and the corresponding data row (row_info).
    - If the unit of the measured biovolume is 'cell' the biovolume is added to the corresponding species.
    - Since there are multiple biovolume values for one species, all other values are appended to the species avoid
      doubling.
    - {Keys: Species, values: biovolume}.
    
    '''
for index, row in dfBiovolume.iterrows():
    searchFor = row['Species']
    row_info = row.values
    if row_info[16] == 'cell':
        if searchFor not in species_Biovolume_dict: # add the species to the dictionary
            species_Biovolume_dict[searchFor] = [row_info[25]] # row info column 25 is the species value for biovolume in µm3/L
        else: # if the species name is already written in the dictionary, add the found biovolume values
            species_Biovolume_dict[searchFor].append(row_info[25])


'----------------------------------------------------------------------------------------------------------------------'

species_avg_Biovolume_dict = {}
''' 2. Creates a dictionary in which the biovolume is averaged per species. This dictionary will be later nested in the 
    taxonomy group order and class level.
    
    - The species to be searchFor is checked on its biovolume value in the dictionary. The average biovolume value per 
      species is added after which divided by the amount of values, rounded by 4 decimals.
    - {Keys: Species, values: average biovolume per species}.
     
    '''
for searchFor, values_list in species_Biovolume_dict.items(): # goes trough the species_BV_dict made in previous step
    avg_BV = round(sum(values_list) / len(values_list), 4)
    species_avg_Biovolume_dict[searchFor] = avg_BV


'----------------------------------------------------------------------------------------------------------------------'

Order_Species_avg_Biovolume_dict = {}
''' 3. Creates a dictionary in which the averaged biovolume per species is nested in the taxonomy group 'order'. This 
    dictionary will later be used to calculate the average biovolume per order. This dict is created  since not all 
    species are found in the database at species level
    
    - Species and order are zipped to be able to call the variables spec and ord. Since the dictionary is empty the 
      order values are added as key with the species biovolume (average) dict as value. all other species values for 
      the order are added as additional value.
    - Some species are not singular cells (form filaments, colonies or coenobium). These species will not have a biovol
      value represented by 0 in the dictionary. Species that form filaments, colonies or coenobium is not covered in the
      scope of this script.
    - Species with their corresponding order from the Species Order database(other_species_dict are added to this 
      dicitonary and nested if the order is already found or added as a new order.
    - {Keys: Order {Keys: Species, values: biovolume}}.
    
    '''
Order_Species_avg_Biovolume_dict_unadapted = {}

for spec, ord in zip(Species, Order):
    if ord not in Order_Species_avg_Biovolume_dict_unadapted:
        Order_Species_avg_Biovolume_dict_unadapted[ord] = {spec: species_avg_Biovolume_dict.get(spec, int(0))}
    else:
        Order_Species_avg_Biovolume_dict_unadapted[ord][spec] = species_avg_Biovolume_dict.get(spec, int(0))

other_species_dict = dict(zip(Species_other, Species_other_order)) # creates a dictionary based on the Species_order
                                                                   # database.

for sp, ord in other_species_dict.items():
    if ord in Order_Species_avg_Biovolume_dict_unadapted:
        Order_Species_avg_Biovolume_dict_unadapted[ord][sp] = other_species_dict.get(ord, 0)
    else:
        Order_Species_avg_Biovolume_dict_unadapted[ord] = {sp: other_species_dict.get(ord, 0)}

for order, biovolume in Order_Species_avg_Biovolume_dict_unadapted.items():
    order_key = str(order)
    if order_key == 'nan':
        order_key = 'Undefined order'
    elif order_key == ' ':
        order_key = 'Undefined order'
    else:
        order_key = order
    Order_Species_avg_Biovolume_dict[order_key] = biovolume

# pprint.pprint(Order_Species_avg_Biovolume_dict)


'----------------------------------------------------------------------------------------------------------------------'

Order_avg_Biovolume_dict = {}
''' 4. This dictionary creates an average biovolume for each order, based on the nested species biovolume values.

    - the keys order and species are called from the previous made dictionary.
    - A list is made for the biovolume values of the species biovolume values.
    - The average biovolume value per order is added and divided by the amount of values, rounded by 4 decimals.
    - {Keys: order, values: average biovolume per order}
    
    '''
for order, species in Order_Species_avg_Biovolume_dict.items():
   values_list = species.values() # create a list of biovolume values for all species in the order
   avg_ord_BV = round(sum(values_list)/len(values_list), 4) # calculate the average biovolume for the order
   Order_avg_Biovolume_dict[order] = avg_ord_BV # store the average biovolume in the new dictionary


# pprint.pprint(Order_avg_Biovolume_dict)

'----------------------------------------------------------------------------------------------------------------------'

Class_Species_avg_Biovolume = {}
''' 5. Creates a dictionary in which the biovolume per corresponding species is nested in the taxonomy group 'class'. 

    - From the order species average biovolume dictionary the order and species are called after which the biovolume 
      value for the species could be defined.
    - All species with their corresponding classes are zipped in a new dictionary (SpeciesClass) to define the class per
      species.
    - If species in the dictionaries are equal the corresponding class is added to a new dictionary. The biovolume of 
      the species is added after which the new dictionary is added to the class species dictionary.
    - {Keys: Class {Keys: species, values: biovolume per species}}  

    '''
SpeciesClass = dict(zip(Species, Classes)) # species on class level

for order, spec in Order_Species_avg_Biovolume_dict.items():
    for sp, value in spec.items():
        for species, Cl in SpeciesClass.items():
            if species == sp:
                class_dict = Class_Species_avg_Biovolume.get(Cl, {})
                class_dict[sp] = value
                Class_Species_avg_Biovolume[Cl] = class_dict

# pprint.pprint(Class_Species_avg_Biovolume)

'----------------------------------------------------------------------------------------------------------------------'
Class_Order_dict = {}
''' 6. Dictionary in which the average order biovolume is nested in taxonomy group 'Class'. 

    - A new dictionary is made from the Species Order database in which the order is attached to the corresponding
      classes (OrderClass)
    - Since not all orders are defined in the Species Order database a try except approach is used to avoid key errors.
    - Classes are defined based on orders of the Order averaged biovolume in OrderClass.
    - classes, orders and the corresponding biovolume are then appended to the Class order dictionary.
      * {Keys: Class {Keys: order, values: biovolume}}.
    
    '''


OrderClass = dict(zip(Order, Classes))

for order, biovolume in Order_avg_Biovolume_dict.items():
    try:
        for ord in OrderClass:
            classes = OrderClass[order]
            if classes in Class_Order_dict:
                Class_Order_dict[classes][order] = biovolume
            else:
                Class_Order_dict[classes] = {order: biovolume}
    except KeyError:
        continue

# pprint.pprint(Class_Order_dict)

'----------------------------------------------------------------------------------------------------------------------'
Class_avg_Biovolume = {}
''' 7. Creates a dictionary in which the biovolume is averaged per class.

    - A new list is made from the biovolume values of the order from the class order dictionary, after which the average
      is calculated and rounded by 4 decimals.
    - The average biovolume is added per class. 
    - {keys: Class, value: biovolume per class}

    '''
for clss, orders in Class_Order_dict.items():
    biovolumes = []
    for values in orders.values():
        biovolumes.append(values)
    avg_biovolume = round(sum(biovolumes) / len(biovolumes), 4)
    Class_avg_Biovolume[clss] = avg_biovolume

# pprint.pprint(Class_avg_Biovolume)
'----------------------------------------------------------------------------------------------------------------------'

Group_Order_Species_avg_Biovolume_dict = {}
''' 8. Creates a dictionary for functional groups (diatoms, dinoflagellates, flagellates, other) with the corresponding 
    order and their species nested in it. This dictionary will provide an addition in conversion to order level and in 
    later stages to group the data in functional groups.
    
    - a new dictionary is made from the Species Order database, based on the order with their corresponding classes 
      (GroupOrder), so that all the species inside these orders can be appended to the Group Order Dictionary.
    - If orders from the GroupOrder and the Order averaged biovolume dictionary are equal the group and the order with 
      the corresponding biovolume are added. 
    - In the second step the species corresponding to the order are added. 
    - {Keys: group {keys: order, values: biovolume per species}}
    '''
GroupOrder = dict(zip(Species_other_order, Species_other_group))

for order, biovolume in Order_avg_Biovolume_dict.items():
    for orderdiff, group in GroupOrder.items():
        if orderdiff == order:
            if group in Group_Order_Species_avg_Biovolume_dict:
                Group_Order_Species_avg_Biovolume_dict[group][order] = biovolume
            else:
                Group_Order_Species_avg_Biovolume_dict[group] = {order: biovolume}

# Nest the species corresponding to the order in the SpOr_merge_BV_orderlevel_Group dictionary
for group, order in Group_Order_Species_avg_Biovolume_dict.items():
    for ord, spec in Order_Species_avg_Biovolume_dict.items():
        if ord in order:
            Group_Order_Species_avg_Biovolume_dict[group][ord] = spec

# pprint.pprint(Group_Order_Species_avg_Biovolume_dict)
'----------------------------------------------------------------------------------------------------------------------'
''' (To avoid scrolling back up) Overview of all dictonaries and their functions:

- 1. species_Biovolume_dict:                    Species from the PEG database with their corresponding biovolume values 
                                                (multiple values per species). 
                                                * {Keys: Species, values: biovolume}.
                                     
- 2. species_avg_Biovolume_dict:                Average biovolume for species from the PEG database. 
                                                * {Keys: Species, values: average biovolume per species}.

- 3. Order_Species_avg_Biovolume_dict:          The average biovolume per species is nested in the corresponding order. 
                                                In this dictionary also the species from the database SpeciesOrder are 
                                                added (other_species_dict) since these species are not found in the PEG
                                                database (The biovolume for these species can only be found on order 
                                                level).
                                                * {Keys: Order {Keys: Species, values: biovolume}}.
                                    
- 4. Order_avg_Biovolume_dict:                  For each order the average biovolume is based on the species values 
                                                inside this order.
                                                * { Keys: order, values: average biovolume per order}
                                    
- 5. Class_Species_avg_Biovolume:               Averaged biovolume per species nested in the corresponding taxanomy 
                                                group 'class'.
                                                * {Keys: Class {Keys: species, values: biovolume per species}}
                                                
- 6. Class_order_avg_Biovolume:                 Averaged biovolume per order nested in the corresponding taxanomy group
                                                'class'.
                                                * {Keys: Class {Keys: Order, values: biovolume per Order}}

- 7. Class_avg_Biovolume:                       Averaged biovolume per class based on the corresponding species.
                                                * {keys: Class, value: biovolume per class}
                                        
- 8. Group_Order_Species_avg_Biovolume_dict:    The functional groups (diatom, dionflagellates, flagellates and other) 
                                                are added to the different taxonomy order.
                                                * {Keys: group {keys: order, values: biovolume per species}}

'''
'----------------------------------------------------------------------------------------------------------------------'
#todo tomorrow: check what is necassary inside (especially the defenitions and stuff). Clean up whole file and check
# variable names etc. Add bar graph to excel and see if you can smarten up the layout things with openpyxl.



def BiovolumeSpeciesLevel (ID_col_name, ID_data, Species_col_name, Species_data, Conc_col_name, Conc_data):
    inputdata_dataframe = pd.DataFrame({ID_col_name: ID_data, Species_col_name: Species_data, Conc_col_name: Conc_data})
    sample_ID = []
    species = []
    result_spec = []
    Concentration = []
    Group = []
    for index, row in inputdata_dataframe.iterrows():
        ID = row[ID_col_name]
        searchFor = row[Species_col_name]
        Conc = row[Conc_col_name]
        if searchFor in species_avg_Biovolume_dict:
            for group, order in Group_Order_Species_avg_Biovolume_dict.items():
                for ord, spec in order.items():
                    if searchFor in spec:
                        BV = round(species_avg_Biovolume_dict[searchFor] * Conc/1e+9, 5)
                        if searchFor not in species:
                            sample_ID.append(ID)
                            species.append(searchFor)
                            Concentration.append(round(Conc, 2))
                            print(f"{searchFor} found in PEG database. The biovolume for {searchFor} is {BV} ml")
                            result_spec.append(BV)
                            Group.append(group)
        else:
            print(f"{searchFor} not found in PEG database at any taxonomy level")
    if len(result_spec) > 0.0:
        df = pd.DataFrame({
            ID_col_name: sample_ID,
            Species_col_name: species,
            'Group': Group,
            Conc_col_name: Concentration,
            'Biovolume (ml)': result_spec
             })
        return df

df = BiovolumeSpeciesLevel( # define the column name and data of the dataset (InputData) to be converted below:

    ID_col_name='ID',  # If there is an ID for the sample, name this column for in the output.
    ID_data=InputData['sample_ID'],  # define in the input data in which column (str(header)) the ID data can be found.
    Species_col_name='Species',  # name the column header for species data in the output.
    Species_data=InputData['spec_name'],  # define in the input data in which column the species data can be found.
    Conc_col_name='Concentration (cells L-1)',  # name the column header for concentration data in the output.
    Conc_data=InputData['conc_cells_per_L']  # define in the input data in which column the conc. data can be found.
    )
#
# '----------------------------------------------------------------------------------------------------------------------'
# #
def BiovolumeOrderLevel(ID_col_name, ID_data, Species_col_name, Species_data, Conc_col_name, Conc_data):
    inputdata_dataframe = pd.DataFrame({ID_col_name: ID_data, Species_col_name: Species_data, Conc_col_name: Conc_data})
    sample_ID = []
    species = []
    result_ord = []
    Concentration = []
    Group = []
    for index, row in inputdata_dataframe.iterrows():
        ID = row[ID_col_name]
        searchFor = row[Species_col_name]
        Conc = row[Conc_col_name]
        if searchFor not in species_avg_Biovolume_dict:
            for ord, species_dict in Order_Species_avg_Biovolume_dict.items():
                if searchFor in species_dict:
                    searchFor_order = ord
                    BV_order = round(Order_avg_Biovolume_dict[searchFor_order] * Conc / 1e+9, 5)
                    # print(f"Biovolume of {searchFor} found at order level ({ord}): {BV_order} ml")
                    if BV_order > 0.0:
                        sample_ID.append(ID)
                        species.append(searchFor)
                        Concentration.append(round(Conc, 2))
                        result_ord.append(BV_order)
                        for ord in Group_Order_Species_avg_Biovolume_dict.values():
                            if searchFor_order in ord:
                                Group.append(group)
                    elif BV_order == 0.0:
                        continue
                        # print(f"See for biovolume for species {searchFor} class level")
                    break
                    # else:
                    #      print(f"Species {searchFor} biovolume not found at species and order level")
    if len(result_ord) > 0:
        df2 = pd.DataFrame({
            ID_col_name: sample_ID,
            Species_col_name: species,
            'Group' : Group,
            Conc_col_name: Concentration,
            'Biovolume (ml)': result_ord,
             })
        return df2

df2 = BiovolumeOrderLevel(# define the column name and data of the dataset (InputData) to be converted below:

    ID_col_name='ID',  # If there is an ID for the sample, name this column for in the output.
    ID_data=InputData['sample_ID'],  # define in the input data in which column (str(header)) the ID data can be found.
    Species_col_name='Species',  # name the column header for species data in the output.
    Species_data=InputData['spec_name'],  # define in the input data in which column the species data can be found.
    Conc_col_name='Concentration (cells L-1)',  # name the column header for concentration data in the output.
    Conc_data=InputData['conc_cells_per_L']  # define in the input data in which column the conc. data can be found.
    )
'----------------------------------------------------------------------------------------------------------------------'
def BiovolumeClassLevel(ID_col_name, ID_data, Species_col_name, Species_data, Conc_col_name, Conc_data):
    inputdata_dataframe = pd.DataFrame({ID_col_name: ID_data, Species_col_name: Species_data, Conc_col_name: Conc_data})
    sample_ID = []
    species = []
    Concentration = []
    result_Class = []
    Group = []
    for index, row in inputdata_dataframe.iterrows():
        ID = row[ID_col_name]
        searchFor = row[Species_col_name]
        Conc = row[Conc_col_name]
        for ord, species_dict in Order_Species_avg_Biovolume_dict.items():
            for spec in species_dict.keys():
                if searchFor == spec:
                    searchFor_order = ord
                    BV_order = round(Order_avg_Biovolume_dict[searchFor_order] * Conc/1e+9, 5)
                    if BV_order == 0:
                        for cls in Class_avg_Biovolume:
                            sample_ID.append(ID)
                            Concentration.append(round(Conc, 2))
                            species.append(searchFor)
                            searchFor_Class = cls
                            BV_Class = round(Class_avg_Biovolume[searchFor_Class] * Conc/1e+9, 5)
                            # print(f"Biovolume of {searchFor} found at class level ({cls}): {BV_Class} ml")
                            result_Class.append(BV_Class)
                            for group, Or in Group_Order_Species_avg_Biovolume_dict.items():
                                for o in Or.keys():
                                    if o == ord:
                                        Group.append(group)
                            break
                        break
                    break
                break
                    # else:
                    #      print(f"Species {searchFor} biovolume not found at species and order level")
    if len(result_Class) > 0:
        df3 = pd.DataFrame({
            ID_col_name: sample_ID,
            Species_col_name: species,
            'Group': Group,
            Conc_col_name: Concentration,
            'Biovolume (ml)': result_Class
             })
        return df3


df3 = BiovolumeClassLevel(# define the column name and data of the dataset (InputData) to be converted below:

    ID_col_name='ID',  # If there is an ID for the sample, name this column for in the output.
    ID_data=InputData['sample_ID'],  # define in the input data in which column (str(header)) the ID data can be found.
    Species_col_name='Species',  # name the column header for species data in the output.
    Species_data=InputData['spec_name'],  # define in the input data in which column the species data can be found.
    Conc_col_name='Concentration (cells L-1)',  # name the column header for concentration data in the output.
    Conc_data=InputData['conc_cells_per_L']  # define in the input data in which column the conc. data can be found.
    )

# if df is not None:
#     print(df)
# elif df2 is not None:
#     print(df2)
# elif df3 is not None:
#     print(df3)
# else:
#     print('No data found')
'----------------------------------------------------------------------------------------------------------------------'

#  creates an excel file with the data from function 1.

with pd.ExcelWriter('Biovolume_test.xlsx') as writer:
    df.to_excel(writer,
                sheet_name = 'Biovolume Species',
                index = False)
    df2.to_excel(writer,
                 sheet_name = 'Biovolume Species',
                 startrow = len(df)+1,
                 index = False,
                 header = False)
    df3.to_excel(writer,
                 sheet_name = 'Biovolume Species',
                 startrow = len(df)+len(df2)+1,
                 index = False,
                 header = False)
#
'----------------------------------------------------------------------------------------------------------------------'

# function outputs to df (choose own path and file name!):
Output_dataframe = pd.read_excel('C:/GIS_Course/EGM722/Project/Project_script/Biovolume_test.xlsx',
                                 sheet_name='Biovolume Species')

def BiovolumeToCarbon (biovolume, cells):
    Carbon = []
    for index, row in Output_dataframe.iterrows():
        Carbon_diatoms = (0.228 * row[biovolume] ** 0.811) * row[cells]
        Carbon_other = (0.216 * row[biovolume] ** 0.939) * row[cells]
        Group = row['Group']
        if Group == 'Diatom':
            Carbon_output = round(Carbon_diatoms, 2)
            # print(f" Carbon for {Group} is {Carbon_output}")
            Carbon.append(Carbon_output)
        else:
            Carbon_output = round(Carbon_other, 2)
            # print(f"Carbon for {Group} is {Carbon_output}")
            Carbon.append(Carbon_output)
    if len(Carbon) > 0:
        df4 = pd.DataFrame({
            'Carbon (pg C)': Carbon
             })
        return df4


df4 = BiovolumeToCarbon(
    biovolume = 'Biovolume (ml)',
    cells= 'Concentration (cells L-1)')

'----------------------------------------------------------------------------------------------------------------------'
with pd.ExcelWriter('Biovolume_test.xlsx') as writer:
    df.to_excel(writer,
                sheet_name='Biovolume Species',
                index=False)
    df2.to_excel(writer,
                 sheet_name='Biovolume Species',
                 startrow=len(df) + 1,
                 index=False,
                 header=False)
    df3.to_excel(writer,
                 sheet_name='Biovolume Species',
                 startrow=len(df) + len(df2) + 1,
                 index=False,
                 header=False)
    df4.to_excel(writer,
                sheet_name = 'Biovolume Species',
                startcol= 5,
                index = False)

Output_dataframe_Carbon = pd.read_excel('C:/GIS_Course/EGM722/Project/Project_script/Biovolume_test.xlsx',
                                 sheet_name='Biovolume Species')

'----------------------------------------------------------------------------------------------------------------------'
def Sum (Concentration, Biovolume, Carbon):
    Con = []
    BV = []
    Carb = []
    Sum_Con = round(sum(Concentration), 2)
    Con.append(Sum_Con)
    Sum_BV = round(sum(Biovolume), 2)
    BV.append(Sum_BV)
    Sum_Car = round(sum(Carbon), 2)
    Carb.append(Sum_Car)
    df5 = pd.DataFrame({
            'Sum_Concentration': Con,
            'Sum_Biovolume': BV,
            'Sum_Carbon': Carb
             })
    return df5

df5 = Sum(
    Concentration = Output_dataframe_Carbon['Concentration (cells L-1)'],
    Biovolume = Output_dataframe_Carbon['Biovolume (ml)'],
    Carbon = Output_dataframe_Carbon['Carbon (pg C)'])

'----------------------------------------------------------------------------------------------------------------------'
Output_dataframe_Carbon = pd.read_excel('C:/GIS_Course/EGM722/Project/Project_script/Biovolume_test.xlsx',
                                 sheet_name='Biovolume Species')

#todo: check function variables below

def AveragesFunctionGroups (Concentration, Biovolume, Carbon):
    Diatom_values = []
    Flagellates_values = []
    Dinoflagellates_values = []
    Other_values = []
    Sum_Carbon_Diatom = []
    Sum_Carbon_Flag = []
    Sum_Carbon_Dino = []
    Sum_Other = []
    for index, row in Output_dataframe_Carbon.iterrows():
        Carbon = row['Carbon (pg C)']
        Group = row['Group']
        if Group == 'Diatom':
            Diatom_values.append(Carbon)
            Sum_Carbon_Diatom = round(sum(Diatom_values), 2)
        elif Group == 'Flagellates':
            Flagellates_values.append(Carbon)
            Sum_Carbon_Flag = round(sum(Flagellates_values), 2)
        elif Group == 'Dinoflagellates':
            Dinoflagellates_values.append(Carbon)
            Sum_Carbon_Dino = round(sum(Dinoflagellates_values), 2)
        else:
            Other_values.append(Carbon)
            Sum_Other = round(sum(Other_values), 2)
    df6 = pd.DataFrame({
        'Carbon Diatoms (pg C)': Sum_Carbon_Diatom,
        'Carbon Flagellates (pg C)': Sum_Carbon_Flag,
        'Carbon Dinoflagellates (pg C)': Sum_Carbon_Dino,
        'Carbon Other (pg C)': Sum_Other},
        index=[0])
    return df6

df6 = AveragesFunctionGroups(
        Concentration=Output_dataframe_Carbon['Concentration (cells L-1)'],
        Biovolume=Output_dataframe_Carbon['Biovolume (ml)'],
        Carbon=Output_dataframe_Carbon['Carbon (pg C)'])





with pd.ExcelWriter('Output_data_conversion_to_carbon.xlsx') as writer:
    df.to_excel(writer,
                sheet_name='Biovolume Species',
                index=False)
    df2.to_excel(writer,
                 sheet_name='Biovolume Species',
                 startrow=len(df) + 1,
                 index=False,
                 header=False)
    df3.to_excel(writer,
                 sheet_name='Biovolume Species',
                 startrow=len(df) + len(df2) + 1,
                 index=False,
                 header=False)
    df4.to_excel(writer,
                sheet_name = 'Biovolume Species',
                startcol= 5,
                index = False)
    df5.to_excel(writer,
                 sheet_name= 'Biovolume Species',
                 startrow = len(df) + len(df2) + len(df3) + 1,
                 startcol= 3,
                 index = False,
                 header = False)
    df6.to_excel(writer,
                 sheet_name= 'Biovolume Species',
                 startcol= 7,
                 index= True,
                 header= True)


# Make workbook look nicer, create averages and graphs

Final_Output = openpyxl.load_workbook('C:/GIS_Course/EGM722/Project/Project_script/Output_data_conversion_to_carbon.xlsx')
sheet = Final_Output['Biovolume Species']
data_range = sheet['A1': get_column_letter(sheet.max_column)+str(sheet.max_row)]

# freeze headers and adjust column width and sort data
sheet.freeze_panes = 'A2'

sheet.column_dimensions['A'].width = 20
sheet.column_dimensions['B'].width = 33
sheet.column_dimensions['C'].width = 20
sheet.column_dimensions['D'].width = 25
sheet.column_dimensions['E'].width = 20
sheet.column_dimensions['F'].width = 15

# changing font
fontobj1 = Font(name='Calibri', size=12, bold=True, italic=False)
sheet['C'+str(sheet.max_row)] = 'Sum:'
sheet['C'+str(sheet.max_row)].font = fontobj1
sheet['D'+str(sheet.max_row)].font = fontobj1
sheet['E'+str(sheet.max_row)].font = fontobj1
sheet['F'+str(sheet.max_row)].font = fontobj1

# sort_data = sorted(data_range, key=lambda row: row[0].value)
#
# for index, row in enumerate(sort_data, start=2):
#     for cell in row:
#         col_letter = coordinate_from_string(cell.coordinate)[0]
#         col_index = column_index_from_string(col_letter)
#         sheet.cell(row=index, column=col_index).value = cell.value
# sheet.delete_rows(sheet.max_row)


# make averages for concentration, biovolume and carbon


Final_Output.save('Output_data_conversion_to_carbon.xlsx')


#todo: eat, write sum in col2, row 67. Make sum and averages for functional groups and show data in bar graphs.














'''-----------------------Below are trials and old setup of functions and codestrings-------------------------------'''
'''-------------------------------------------Just to be sure-------------------------------------------------------'''
# todo clean at the end!

#elif searchFor != row['Species']:
#    print('Biovolume for {} cant be calculated based on species level'.format(searchFor))
            #elif searchFor != row['Species']:

#wbSpecData = openpyxl.load_workbook('C:/GIS_Course/EGM722/Project/Data_files/Data_2018_GPV.xlsx')
#sheetSpecData = wbSpecData['Sheet1']
#print(sheetSpecData.max_row) # 1721



# def BiovolumeGenLevel(dataset, datasetColSpec, dataTobBeConverted, database, databaseColumn): # define the correct files and
#     input data (see instructions below print(Biovolume)
#    datasetColSpec = dataset[datasetColSpec]
#    for i in range(len(datasetColSpec)): # iterate over the species names in dataset
#        searchFor = datasetColSpec[i]
#        try:
#            GenusSearchFor = SpGe[searchFor]
#        except KeyError:
#            print('Species {} not in database'.format(searchFor))
#            continue
#         for index, row in database.iterrows(): # iterate over the species name in the database
#            row_info = row.values  # gives row information to define the unit
#            if GenusSearchFor == row[(databaseColumn)]: # check if the species name is inside the database
#                if row_info[16] == 'cell':
#                    BV = round(row_info[25] * dataset[dataTobBeConverted][0]/1e+9, 4) # *1000/1e+12 to convert µm3/L,
#                    # to ml with 4 decimals
#                    print('Biovolume for {} Genus is {} ml'.format(GenusSearchFor, BV))


#def BiovolumeSpecLevel(dataset, datasetColSpec, dataTobBeConverted, database, databaseColumn): # define the correct files and
# #input data (see instructions below print(Biovolume)
#    datasetColSpec = dataset[datasetColSpec]
#    for i in range(len(datasetColSpec)): # iterate over the species names in dataset
#        searchFor = datasetColSpec[i]
#        try: # check if the species name is inside the database
#            searchFor == database[databaseColumn]
#        except:
#            searchFor != database[databaseColumn]
#            print('Biovolume for {} species is not in database'.format(searchFor))
#            continue
#        for index, row in database.iterrows():  # iterate over the species name in the database
#            if searchFor == database[databaseColumn][index]:
#                row_info = row.values  # gives row information to define the unit
#                if row_info[16] == 'cell':
#                     BV = round(row_info[25] * dataset[dataTobBeConverted][0] / 1e+9, 4)  # *1000/1e+12 to convert µm3/L,
#                     # to ml with 4 decimals
#                     print('Biovolume for {} species is {} ml'.format(searchFor, BV))

# result = dict(zip(np.array(searchFor), np.array(BV)))
# df = pd.DataFrame(result)
# df.to_excel('Biovolume_test.xlxs', index=True)
# for rowNum in range(2, 100):
# BiovolumeWbSheet.cell(row=rowNum, column=2).value = df
# BiovolumeWbSheet.cell(row=rowNum, column=3).value = searchFor




# def BiovolumeGenLevel(dataset, datasetColSpec, dataTobBeConverted, database, databaseColumn): # define the correct files and
#     # input data (see instructions below print(Biovolume)
#     datasetColSpec = dataset[datasetColSpec]
#     for i in range(len(datasetColSpec)): # iterate over the species names in dataset
#         searchFor = datasetColSpec[i]
#         try:
#             GenusSearchFor = SpGe[searchFor]
#         except KeyError:
#             print('Species {} not in database'.format(searchFor))
#             continue
#         for index, row in database.iterrows(): # iterate over the species name in the database
#             row_info = row.values  # gives row information to define the unit
#             if GenusSearchFor == row[(databaseColumn)]: # check if the species name is inside the database
#                 if row_info[16] == 'cell':
#                     BV = round(row_info[25] * dataset[dataTobBeConverted][0]/1e+9, 4) # *1000/1e+12 to convert µm3/L,
#                     # to ml with 4 decimals
#                     print('Biovolume for {} Genus is {} ml'.format(GenusSearchFor, BV))


# define here the dataset and databases for function 2:
# BiovolumeGenLevel(
#     dataset=SpecData, # define the dataset (the workbook sheet of the data to be converted)
#     datasetColSpec='spec_name', # Define the column in the sheet where the species name data can be found (string!)
#     dataTobBeConverted='conc_cells_per_L', # define the column in the sheet where the data to be converted to biovolume
#     # can be found (string!)
#     database=dfBV, # define which database must be used for biovolume conversion
#     databaseColumn='Genus') # Define the column where the database Genus name can be found (string!)

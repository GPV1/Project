import openpyxl
from openpyxl.styles import Font
import pandas as pd
import numpy as np
import pprint
import math
import sqlite3

from pandas import ExcelWriter

# todo check at the end if all the above imports are needed.


# import datafiles and sheetname. Necessary: database files PEG and Species order and input data
dfBiovolume = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/PEG_BVOL2019_PJ.xlsx', sheet_name='Biovolume file')
dfSpeciesOrder = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/Species_order_GPV.xlsx', sheet_name='Sheet1')
InputData = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/Data_2018_GPV.xlsx', sheet_name='Sheet2')
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

- 6. Class_avg_Biovolume:                       Averaged biovolume per class based on the corresponding species.
                                                * {keys: Class, value: biovolume per class}

- 7. Group_Order_Species_avg_Biovolume_dict:    The functional groups (diatom, dionflagellates, flagellates and other) 
                                                are added to the different taxonomy order.
                                                * {Keys: group {keys: order, values: biovolume per order}}

'''
'----------------------------------------------------------------------------------------------------------------------'

species_Biovolume_dict = {}
''' 1. Creates a dictionary from the PEG database based on the species and the corresponding biovolume (µm3/L). This 
    dictionary will later be used to average the biovolume per species.
    
    - Defines the variable searchFor (the species name) and the corresponding data row (row_info).
    - If the unit of the measured biovolume is 'cell' the biovolume is added to the corresponding species.
    - Since there are multiple biovolume values for one species, all other values are appended to the species avoid
      doubling.
    
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
''' 2. Creates a dictionary in which the biovolume is averaged per species. This dictionary will be nested at the 
    taxonomy group order and class level.
    
    - The species to be searchFor is checked on its biovolume value in the dictionary.
    - The average biovolume value per species is added and divided by the amount of values, rounded by 4 decimals.
     
    '''
for searchFor, values_list in species_Biovolume_dict.items(): # goes trough the species_BV_dict made in previous step
    avg_BV = round(sum(values_list) / len(values_list), 4)
    species_avg_Biovolume_dict[searchFor] = avg_BV

'----------------------------------------------------------------------------------------------------------------------'

Order_Species_avg_Biovolume_dict = {}
''' 3. Creates a dictionary in which the averaged biovolume per species is nested in the taxonomy group 'order'. This 
    dictionary will later be used to calculate the average biovolume per order. This dict is created  since not all 
    species are found in the database at species level
    
    - Species and order are zipped to be able to call the variables spec and ord.
    - Since the dictionary is empty the order values are added as key with the species biovolume (average) dict as value
    - all other species values for the order are added as additional value.
    - Some species are not singular cells (form filaments, colonies or coenobium). These species will not have a biovol
      value represented by 0 in the dictionary. Species that form filaments, colonies or coenobium is not covered in the
      scope of this script.
    
    '''
Order_Species_avg_Biovolume_dict_unadapted = {}

for spec, ord in zip(Species, Order):
    # print(f"Processing species: {spec}, order: {ord}")
    if ord not in Order_Species_avg_Biovolume_dict_unadapted:
        Order_Species_avg_Biovolume_dict_unadapted[ord] = {spec: species_avg_Biovolume_dict.get(spec, int(0))}
        # print(f"Adding new order: {ord}")
    else:
        Order_Species_avg_Biovolume_dict_unadapted[ord][spec] = species_avg_Biovolume_dict.get(spec, int(0))
        # print(f"Adding species to existing order: {ord}")

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
    - A variable is made for the biovolume values in of the species biovolume values.
    - The average biovolume value per order is added and divided by the amount of values, rounded by 4 decimals.
    
    '''
for order, species in Order_Species_avg_Biovolume_dict.items():
   values_list = species.values() # create a list of biovolume values for all species in the order
   avg_ord_BV = round(sum(values_list)/len(values_list), 4) # calculate the average biovolume for the order
   Order_avg_Biovolume_dict[order] = avg_ord_BV # store the average biovolume in the new dictionary


# pprint.pprint(Order_avg_Biovolume_dict)

'----------------------------------------------------------------------------------------------------------------------'

Class_Species_avg_Biovolume = {}
''' 5. Creates a ditionary in which the biovolume per corresponding species is nested in the taxonomy group 'class'.

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
''' Dictionary in which the average order biovolume is nested in taxonomy group 'Class' 

    '''
#todo: fix dictionary below, then adapt info and function 3.

OrderClass = dict(zip(Order, Classes))

Class_Order_dict_unadapted = {}
for order, clss in OrderClass.items():
    for ord, values in Order_avg_Biovolume_dict.items():
        if clss in Class_Order_dict_unadapted:
            Class_Order_dict_unadapted[clss][order] = values
        else:
            Class_Order_dict_unadapted[clss] = {ord: values}

pprint.pprint(OrderClass)

for clss, order in Class_Order_dict_unadapted.items():
    order_key = str(order)
    if order_key == 'nan':
        order_key = 'Undefined order'
    elif order_key == ' ':
        order_key = 'Undefined order'
    else:
        order_key = order
    Class_Order_dict[clss][order_key] = biovolume

pprint.pprint(Class_Order_dict)


'----------------------------------------------------------------------------------------------------------------------'
Class_avg_Biovolume = {}
''' 6. Creates a dictionary where the biovolume is averaged per class.

    '''
for clss, orders in Class_Order_dict.items():
    # print(clss)
    pprint.pprint(orders)
    biovolumes = []
    for order in orders.values():
        biovolumes.append(order)
    avg_biovolume = round(sum(biovolumes) / len(biovolumes), 4)
    Class_avg_Biovolume[clss] = avg_biovolume

# pprint.pprint(Class_avg_Biovolume)
'----------------------------------------------------------------------------------------------------------------------'

Group_Order_Species_avg_Biovolume_dict = {}
''' 7. Creates a dictionary for functional groups (diatoms, dinoflagellates, flagellates, other) with the corresponding 
    order and their species nested in it. This dictionary will provide an addition in conversion to order level and in 
    later stages to group the data in functional groups.
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
                                    
- 4. Order_avg_Biovolume_dict_adapted:          For each order the average biovolume is based on the species values 
                                                inside this order.
                                                * { Keys: order, values: average biovolume per order}
                                    
- 5. Class_Species_avg_Biovolume:               Averaged biovolume per species nested in the corresponding taxanomy 
                                                group 'class'.
                                                * {Keys: Class {Keys: species, values: biovolume per species}}

- 6. Class_avg_Biovolume:                       Averaged biovolume per class based on the corresponding species.
                                                * {keys: Class, value: biovolume per class}
                                        
- 7. Group_Order_Species_avg_Biovolume_dict:    The functional groups (diatom, dionflagellates, flagellates and other) 
                                                are added to the different taxonomy order.
                                                * {Keys: group {keys: order, values: biovolume per order}}

'''
'----------------------------------------------------------------------------------------------------------------------'
#todo- tomorrow: fix the group in the functions (class level, spec and order level are fixed).
# Add the carbon calculations to the outcomes (new function) and print it in the same excel.


def BiovolumeSpeciesLevel (dataset, datasetColSpec, dataTobBeConverted):
    datasetColSpec = dataset[datasetColSpec]
    species = []
    result_spec = []
    Group = []
    for searchFor in datasetColSpec:
        if searchFor in species_avg_Biovolume_dict:
            # print(searchFor)
            for group, order in Group_Order_Species_avg_Biovolume_dict.items():
                for ord, spec in order.items():
                    if searchFor in spec:
                        species.append(searchFor)
                        BV = round(species_avg_Biovolume_dict[searchFor] * dataset[dataTobBeConverted][0] / 1e+9, 4)
                        print(f"{searchFor} found in PEG database. The biovolume for {searchFor} is {BV} ml")
                        result_spec.append(BV)
                        Group.append(group)
        # else:
        #     print(f"{searchFor} not found in PEG database at any taxonomy level")
    if len(result_spec) > 0.0:
        df = pd.DataFrame({
            'Species': species,
            'Group': Group,
            'Biovolume (ml)': result_spec
             })
        return df
#
df = BiovolumeSpeciesLevel(
        dataset = InputData,  # define the dataset (the workbook sheet of the data to be converted)
        datasetColSpec = 'spec_name',  # Define the col in the sheet where the species name data can be found (string!)
        dataTobBeConverted = 'conc_cells_per_L')  # define the column in the sheet where the data to be converted to
        # biovolume can be found (string!)


#
def BiovolumeOrderLevel(dataset, datasetColSpec, dataTobBeConverted):
    datasetColSpec = dataset[datasetColSpec]
    species = []
    order = []
    result_ord = []
    Group = []
    for searchFor in datasetColSpec:
        if searchFor not in species_avg_Biovolume_dict:
            for ord, species_dict in Order_Species_avg_Biovolume_dict.items():
                if searchFor in species_dict:
                    searchFor_order = ord
                    BV_order = round(Order_avg_Biovolume_dict[searchFor_order] * dataset[dataTobBeConverted][0] / 1e+9, 4)
                    print(f"Biovolume of {searchFor} found at order level ({ord}): {BV_order} ml")
                    if BV_order > 0.0:
                        species.append(searchFor)
                        order.append(searchFor_order)
                        result_ord.append(BV_order)
                        for ord in Group_Order_Species_avg_Biovolume_dict.values():
                            if searchFor_order in ord:
                                # print(searchFor_order)
                                Group.append(group)
                                # print(group)
                    elif BV_order == 0.0:
                        print(f"See for biovolume for species {searchFor} class level")
                    break
                    # else:
                    #      print(f"Species {searchFor} biovolume not found at species and order level")
    if len(result_ord) > 0:
        df2 = pd.DataFrame({
            'Species': species,
            # 'Order': order,
            'Group' : Group,
            'Biovolume (ml)': result_ord,
             })
        return df2

df2 = BiovolumeOrderLevel(
        dataset = InputData,  # define the dataset (the workbook sheet of the data to be converted)
        datasetColSpec = 'spec_name',  # Define the col in the sheet where the species name data can be found (string!)
        dataTobBeConverted = 'conc_cells_per_L')  # define the column in the sheet where the data to be converted to
        # biovolume can be found (string!)
#
#
def BiovolumeClassLevel(dataset, datasetColSpec, dataTobBeConverted):
    datasetColSpec = dataset[datasetColSpec]
    species = []
    Class = []
    result_Class = []
    for searchFor in datasetColSpec:
        for ord, species_dict in Order_Species_avg_Biovolume_dict.items():
            if searchFor in species_dict:
                searchFor_order = ord
                BV_order = round(Order_avg_Biovolume_dict[searchFor_order] * dataset[dataTobBeConverted][0] / 1e+9, 4)
                if BV_order == 0:
                    print(f"{searchFor} biovolume cannot be calculated on order {searchFor_order} level")
                    for cls in Class_avg_Biovolume:
                        species.append(searchFor)
                        searchFor_Class = cls
                        Class.append(searchFor_Class)
                        BV_Class = round(Class_avg_Biovolume[searchFor_Class] * dataset[dataTobBeConverted][0] / 1e+9, 10)
                        print(f"Biovolume of {searchFor} found at class level ({cls}): {BV_Class} ml")
                        result_Class.append(BV_Class)
                        break
                    # else:
                    #      print(f"Species {searchFor} biovolume not found at species and order level")
    if len(result_Class) > 0:
        df3 = pd.DataFrame({
            'Species': species,
            # 'Class': Class,
            'Biovolume (ml)': result_Class,
             })
        return df3

df3 = BiovolumeClassLevel(
        dataset = InputData,  # define the dataset (the workbook sheet of the data to be converted)
        datasetColSpec = 'spec_name',  # Define the col in the sheet where the species name data can be found (string!)
        dataTobBeConverted = 'conc_cells_per_L')  # define the column in the sheet where the data to be converted to
        # biovolume can be found (string!)
#
#
# # if df is not None:
# #     print(df)
# # elif d2 is not None:
# #     print(df2)
# # elif df3 is not None:
# #     print(df3)
# # else:
# #     print('No data found')
#
# # print(df, df2, df3)
#
# # creates an excel file with the data from function 1.
#
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

import openpyxl
from openpyxl.styles import Font
import pandas as pd
import numpy as np
import pprint
import sqlite3
# todo check at the end if all the above imports are needed.


# import datafiles and sheetname. Necessary: database files PEG and Species order and input data
dfBV = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/PEG_BVOL2019_PJ.xlsx',sheet_name='Biovolume file')
dfSpN = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/Species_order_GPV.xlsx',sheet_name='Sheet1')
SpecData = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/Data_2018_GPV.xlsx', sheet_name='Sheet2')
test_data = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/test_data.xlsx', sheet_name='data')


# check the column headers to choose the right columns for conversion functions
# print(SpecData.columns,
# dfBV.columns,
# test_data.columns)


''' preperation of database by forming dictionaries for taxonomy groups, creating averages of multiple biovolume
 values per species and usage of dictionaries in functions'''

#  defines the columns in the database for the taxonomy groups
divisions = dfBV['Division']
Classes = dfBV['Class']
Order = dfBV['Order'].str.title()
Genus = dfBV['Genus']
Species = dfBV['Species']
print(Species)


#  defines dictionaries of different levels in taxonomy groups based on the PEG database
SpGe = dict(zip(Species, Genus)) # species on genus level
SpOr = dict(zip(Species, Order)) # species on order level
SpCl = dict(zip(Species, Classes)) # species on class level






species_BV_dict = {}
''' Creates a dictionary from the PEG database based on the species and the corresponding biovolume (µm3/L). This 
    dictionary will later be used to average the biovolume per species.
    
    - Defines the variable searchFor (the species name) and the corresponding data row (row_info).
    - If the unit of the measured biovolume is 'cell' the biovolume is added to the corresponding species.
    - Since there are multiple biovolume values for one species, all other values are appended to the species avoid
      doubling.
    
    '''
for index, row in dfBV.iterrows():
    searchFor = row['Species']
    row_info = row.values
    if row_info[16] == 'cell':
        if searchFor not in species_BV_dict: # add the species to the dictionary
            species_BV_dict[searchFor] = [row_info[25]] # row info column 25 is the species value for biovolume in µm3/L
        else: # if the species name is already written in the dictionary, add the found biovolume values
            species_BV_dict[searchFor].append(row_info[25])







species_avg_BV_specieslevel = {}
''' Creates a dictionary in which the biovolume is averaged per species. This dictionary will be nested at the taxonomy
    group genus and order level.
    
    - The species to be searchFor is checked on its biovolume value in the dictionary.
    - The average biovolume value per species is added and divided by the amount of values, rounded by 4 decimals.
     
    '''
for searchFor, values_list in species_BV_dict.items(): # goes trough the species_BV_dict made in previous step
    avg_BV = round(sum(values_list) / len(values_list), 4)
    species_avg_BV_specieslevel[searchFor] = avg_BV




SpOr_merge = {}
''' Creates another dictionary in which the averaged biovolume per species is nested in the taxonomy group order. This 
    dictionary will later be used to calculate the average biovolume per order. This is done since not all species are
    found in the database at species level
    
    - Species and order are zipped to be able to call the variables spec and ord.
    - Since the dictionary is empty the order values are added as key with the species biovolume (average) dict as value
    - all other species values for the order are added as additional value.
    - Some species are not singular cells (form filaments, colonies or coenobium). These species will not have a biovol
      value represented by 0 in the dictionary. Species that form filaments, colonies or coenobium is not covered in the
      scope of this script.
    
    '''
# todo- Somehow get the species different in the species list and then add all the dicts together. Also fix the function
# todo- for the bioconversion..

SpecDiff = dfSpN['spec_name'].str.replace('_', ' ')
Species = Species + SpecDiff
print(Species)


    # if ord not in SpOr_merge:
    #     SpOr_merge[ord] = {spec: species_avg_BV_specieslevel.get(spec, 0)}
    # else:
    #     SpOr_merge[ord] = species_avg_BV_specieslevel.get(spec, 0)

# for spec, ord in zip(Species, Order):
#     print(f"Processing species: {spec}, order: {ord}")
#     if ord not in SpOr_merge:
#         SpOr_merge[ord] = {spec: species_avg_BV_specieslevel.get(spec, 0)}
#         print(f"Adding new order: {ord}")
#     else:
#         SpOr_merge[ord][spec] = species_avg_BV_specieslevel.get(spec, 0)
#         print(f"Adding species to existing order: {ord}")

pprint.pprint(SpOr_merge)



# for spec, SpecDiff, ord in zip(Species, SpecDiff, Order):
#     if ord not in SpOr_merge:
#         SpOr_merge[ord] = {SpecDiff: species_avg_BV_specieslevel.get(spec, 0)} # create a new key value (genus) with the
#         # corresponding species and biovolume nested as value.
#     else:
#         SpOr_merge[ord][SpecDiff] = species_avg_BV_specieslevel.get(spec, 0) # nest the species avg biovolume dictionary
#         # for the corresponding order made in previous step
#
# pprint.pprint(SpOr_merge)

SpOr_merge_BV_orderlevel = {}
''' This dictionary creates an average biovolume for each order, based on the nested species biovolume values.

    - the keys order and species are called from the previous made dictionary.
    - A variable is made for the biovolume values in of the species biovolume values.
    - The average biovolume value per order is added and divided by the amount of values, rounded by 4 decimals.
    
    '''
for order, species in SpOr_merge.items():
   values_list = [biovolume for biovolume in species.values()] # create a list of biovolume values for all species in the order
   avg_ord_BV = round(sum(values_list)/len(values_list), 4) # calculate the average biovolume for the order
   SpOr_merge_BV_orderlevel[order] = avg_ord_BV # store the average biovolume in the new dictionary

# todo make another dictionary based on the species order file. It needs to set the order to the species name that is
# todo - used in certain formats.


pprint.pprint(SpOr_merge_BV_orderlevel)



group_dict = {}
''' Creates the last dictionary needed for the biovolume conversion. This dictionary is an addition to the above dict
    since not all species are defined the same way (due to size classes etc). This dictionary will provide an addition
    in conversion to order level and in later stages to group the data in functional groups.
    '''
SpecDiff = dfSpN['spec_name'].str.replace('_', ' ')
OrdDiff = dfSpN['order']
Group = dfSpN['group']
SpecName_different_Format = dict(zip(SpecDiff, OrdDiff))
for i in range(len(Group)):
    group = Group[i]
    orderDiff = OrdDiff[i]
    species = SpecDiff[i]
    if group not in group_dict:
        group_dict[group] = {orderDiff: [species]}
    else:
        if orderDiff not in group_dict[group]:
            group_dict[group][orderDiff] = [species]
        else:
            group_dict[group][orderDiff].append(species)
# pprint.pprint(group_dict)







def BiovolumeSpecLevel(dataset, datasetColSpec, dataTobBeConverted):
    datasetColSpec = dataset[datasetColSpec]
    species = []
    order = []
    result_spec = []
    result_ord = []
    for i in range(len(datasetColSpec)):
        searchFor = datasetColSpec[i]
        try:
            print(searchFor in group_dict)
            if searchFor in group_dict:

                orderDiff = group_dict[searchFor]
                species.append(searchFor)
                print(f"Match found: {searchFor}")
                if searchFor in species_avg_BV_specieslevel:
                    BV = species_avg_BV_specieslevel[searchFor] * dataset[dataTobBeConverted][0] / 1e+9
                    result_spec.append(BV)
                else:
                    order.append(orderDiff)
                    print(ord_searchFor)
                    BV_ord_searchFor = SpOr_merge_BV_orderlevel[ord_searchFor] * dataset[dataTobBeConverted][0] / 1e+9
                    result_ord.append(BV_ord_searchFor)
                    print(f"Match not found at species level for species {searchFor}, biovolume caluclated on Order level")
        except KeyError:
            print('Match not found in databases')
        #else:
           # print('No match found in databases')
    if len(result_spec) > 0:
        df = pd.DataFrame({
            'Species': species,
            'Biovolume (ml)': result_spec,
            'Order': order,
            'Biovolume Genus (ml)': result_ord})
        return df

df = BiovolumeSpecLevel(
        dataset=SpecData,  # define the dataset (the workbook sheet of the data to be converted)
        datasetColSpec='spec_name',  # Define the column in the sheet where the species name data can be found (string!)
        dataTobBeConverted='conc_cells_per_L')  # define the column in the sheet where the data to be converted to biovolume
        # can be found (string!)


# def BiovolumeSpecLevel(dataset, datasetColSpec, dataTobBeConverted):
#     ''' function 1: Conversion of gathered data to biovolume on species level database (Always run function 2 for species
#     NOT in database)'''
#     datasetColSpec = dataset[datasetColSpec]
#     species = []
#     order = []
#     result_spec = []
#     result_ord = []
#     for i in range(len(datasetColSpec)):
#         searchFor = datasetColSpec[i]
#         if searchFor in species_avg_BV_specieslevel and searchFor == species_avg_BV_specieslevel[searchFor]:
#             species.append(searchFor)
#             print(f"Match found: {searchFor}")
#             BV = species_avg_BV_specieslevel[searchFor] * dataset[dataTobBeConverted][0] / 1e+9
#             result_spec.append(BV)
#         elif searchFor not in species_avg_BV_specieslevel:
#             order.append(searchFor)
#             print((f"Match not found on species level: {searchFor}"))
#             if searchFor == group_dict:
#                 ord_searchFor = SpOr_merge_BV_orderlevel[searchFor] * dataset[dataTobBeConverted][0] / 1e+9
#                 result_ord.append(ord_searchFor)
#             else:
#                 result_ord.append(None)
#         else:
#             print(f"No match found for {searchFor}")
#     if len(result_spec) > 0:
#         df = pd.DataFrame({
#             'Species': species,
#             'Biovolume (ml)': result_spec,
#             'Order': order,
#             'Biovolume Genus (ml)': result_ord})
#         return df
#     else:
#         print('No data found')

# define here the dataset and databases for function 1:
# df = BiovolumeSpecLevel(
#     dataset= SpecData,                       # define the dataset (the workbook sheet of the data to be converted)
#     datasetColSpec= 'spec_name',             # Define the column in the sheet where the species name data can be found (string!)
#     dataTobBeConverted= 'conc_cells_per_L')  # define the column in the sheet where the data to be converted to biovolume
#                                              # can be found (string!)

if df is not None:
    print(df)
else:
    print('No data found')


# creates an excel file with the data from function 1.
with pd.ExcelWriter('Biovolume_test.xlsx') as writer:
    df_empty = pd.DataFrame()
    df_empty.to_excel(writer, sheet_name='Sheet1', index = False)
    df.to_excel(writer, sheet_name='Biovolume', index=False)


print('--------BREAK--------')


# todo CLEAN THIS THING UP! And make a df for function two. Also add the species name in df 1. Maybe combine dataframe
# todo-  of the two functions. Work on conversion to Carbon. Oh and fix the doubling thing in the functions..

''' function 2: conversion of gathered data to biovolume for species NOT in database. Run this function always to be
sure all data is gathered'''

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
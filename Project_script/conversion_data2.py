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
print(SpecData.columns,
dfBV.columns,
test_data.columns)


''' preperation of database by forming dictionaries for taxonomy groups, creating averages of multiple biovolume
 values per species and usage of dictionaries in functions'''

#  defines the columns in the database for the taxonomy groups
divisions = dfBV['Division']
Classes = dfBV['Class']
Order = dfBV['Order']
Genus = dfBV['Genus']
Species = dfBV['Species'].replace('_', ' ')


#  defines dictionaries of different levels in taxonomy groups
SpGe = dict(zip(Species, Genus)) # species on genus level
SpOr = dict(zip(Species, Order)) # species on order level
SpCl = dict(zip(Species, Classes)) # species on class level

# pprint.pprint(SpGe)
# pprint.pprint(Species)
# pprint.pprint(Genus)


# Defines a dictionary for the species biovolume values stored in the database
species_BV_dict = {} # makes a new dictionary
for index, row in dfBV.iterrows(): # iterates over the database
    searchFor = row['Species'] # search for species in the database
    row_info = row.values # makes a variable for the complete row values of the found species
    if row_info[16] == 'cell': # if the unit of the biovolume measurement is cell
        if searchFor not in species_BV_dict: # add the species to the dictionary
            species_BV_dict[searchFor] = [row_info[25]] # row info column 25 is the species value for biovolume in µm3/L
        else: # if the species name is already written in the dictionary, add the found biovolume values
            species_BV_dict[searchFor].append(row_info[25])

# pprint.pprint(species_BV_dict)

# todo make a function of below were the name of the wanted dictionary can be written
# makes a new dictionary and averages the found biovolume for each species (since the database as multiple biovolume
# values per species).
species_avg_BV_specieslevel = {}
for searchFor, values_list in species_BV_dict.items(): # goes trough the species_BV_dict made in previous step
    avg_BV = round(sum(values_list) / len(values_list), 4) # makes an average of biovolume for each species and round it
                                                           # to 4 decimal places
    species_avg_BV_specieslevel[searchFor] = avg_BV # adds the averaged biovolume µm3/L to each species in the new dictionary

pprint.pprint(species_avg_BV_specieslevel)

# To also find the biovolume values at genus level a third dictionary is made to merge genus with species BV.
SpGe_avg = {} # a new dictionary is made
for spec, gen in zip(Species, Genus): # species and genus are zipped together to find the connection
    if gen not in SpGe_avg: # if the genus is not in the new dictionary
        SpGe_avg[gen] = {}
        SpGe_avg[gen] = {spec: species_avg_BV_specieslevel.get(spec, 0)} # create a new key value (genus) with the corresponding
            # species and biovolume nested. 0 means not available!
    else: # and adding the species value
        SpGe_avg[gen][spec] = species_avg_BV_specieslevel.get(spec, 0) # if the genus level is already in the dictionary, add the
            # corresponding species to the dictionary. 0 means not available!



# pprint.pprint(SpGe_avg.items())

# makes a new dictionary and averages the found biovolume for each genus (since not all species biovolume can be found
# on species level)
SpGe_avg_BV_genuslevel = {}
for genus, species in SpGe_avg.items():
   values_list = [biovolume for biovolume in species.values()] # create a list of biovolume values for all species in the genus
   avg_Gen_BV = round(sum(values_list)/len(values_list), 4) # calculate the average biovolume for the genus
   SpGe_avg_BV_genuslevel[genus] = avg_Gen_BV # store the average biovolume in the new dictionary


pprint.pprint(SpGe_avg_BV_genuslevel)

''' function 1: Conversion of gathered data to biovolume on species level database (Always run function 2 for species
NOT in database)'''
# todo fix error in this function..
def BiovolumeSpecLevel(dataset, datasetColSpec, dataTobBeConverted):
    datasetColSpec = dataset[datasetColSpec]
    species = []
    genus = []
    result_spec = []
    result_gen = []
    try:
        for i in range(len(datasetColSpec)):
            searchFor = datasetColSpec[i]
            if searchFor == species_avg_BV_specieslevel[searchFor]:
                species.append(searchFor)
                print(f"Match found: {searchFor}")
                BV = species_avg_BV_specieslevel[searchFor] * dataset[dataTobBeConverted][0] / 1e+9
                result_spec.append(BV)
                continue
            elif searchFor != species_avg_BV_specieslevel[searchFor]:
                genus.append(searchFor)
                print((f"Match not found on species level: {searchFor}"))
                gen_searchFor = SpGe_avg_BV_genuslevel[searchFor] * dataset[dataTobBeConverted][0] / 1e+9
                result_gen.append(gen_searchFor)
                continue
        if len(result_spec) > 0:
            df = pd.DataFrame({'Species': species, 'Biovolume (ml)': result_spec, 'Genus': Genus, 'Biovolume Genus (ml)':
                           result_gen})
            return df
        else:
            print('No data found')
    except KeyError:
        print('NA')

# define here the dataset and databases for function 1:
df = BiovolumeSpecLevel(
   dataset= SpecData, # define the dataset (the workbook sheet of the data to be converted)
   datasetColSpec= 'spec_name', # Define the column in the sheet where the species name data can be found (string!)
   dataTobBeConverted= 'conc_cells_per_L') # define the column in the sheet where the data to be converted to biovolume
                                          # can be found (string!)

if df is not None:
    print(df)
else:
    print('No data found')


# creates an excel file with the data from function 1.
with pd.ExcelWriter('Biovolume_test.xlsx') as writer:
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
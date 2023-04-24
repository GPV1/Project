import openpyxl
from openpyxl.styles import Font
import pandas as pd
import numpy as np
import pprint
import sqlite3


# import datafiles and sheetname. Necessary: database files PEG and Species order and input data
dfBV = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/PEG_BVOL2019_PJ.xlsx',sheet_name='Biovolume file')
dfSpN = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/Species_order_GPV.xlsx',sheet_name='Sheet1')
SpecData = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/Data_2018_GPV.xlsx', sheet_name='Sheet2')
test_data = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/test_data.xlsx', sheet_name='data')

# check the column headers to later choose the right columns for conversion
# print(SpecData.columns)
# print(dfBV.columns)
# print(test_data.columns)

#  define dictionaries for taxonomy groups
divisions = dfBV['Division']
Classes = dfBV['Class']
Order = dfBV['Order']
Genus = dfBV['Genus']
Species = dfBV['Species'].replace('_', ' ')

SpGe = dict(zip(Species, Genus))
SpOr = dict(zip(Species, Order))
SpCl = dict(zip(Species, Classes))

# create a new workbook to store the converted data
# todo make something smarter for below
# BiovolumeWb = openpyxl.Workbook()
# sheet = BiovolumeWb.active
# sheet.title = 'Biovolume'
# BiovolumeWb.save('Biovolume_test.xlsx')
# BiovolumeWb = openpyxl.load_workbook('Biovolume_test.xlsx')
# BiovolumeWbSheet = BiovolumeWb['Biovolume']


''' function 1: Conversion of gathered data to biovolume on species level database (Always run function 2 for species 
NOT in database)'''

#  Todo fix the exception thing

# def BiovolumeSpecLevel(dataset, datasetColSpec, dataTobBeConverted, database, databaseColumn): # define the correct files and
# #input data (see instructions below print(Biovolume)
#    datasetColSpec = dataset[datasetColSpec]
#    for i in range(len(datasetColSpec)): # iterate over the species names in dataset
#        searchFor = datasetColSpec[i]
#        try: # check if the species name is inside the database
#            searchFor == database[databaseColumn]
#        except:
#            print('Biovolume for {} species is not in database'.format(searchFor))
#            continue
#        for index, row in database.iterrows():  # iterate over the species name in the database
#            if searchFor == database[databaseColumn][index]:
#                row_info = row.values  # gives row information to define the unit
#                if row_info[16] == 'cell':
#                    BV = round(row_info[25] * dataset[dataTobBeConverted][0] / 1e+9, 4)  # *1000/1e+12 to convert µm3/L,
#                    # to ml with 4 decimals
#                    print('Biovolume for {} species is {} ml'.format(searchFor, BV))
#                    df = pd.DataFrame(BV, columns=['A'])
#                    return df
#                        # BiovolumeWb.save('Biovolume_test.xlsx')

def BiovolumeSpecLevel(dataset, datasetColSpec, dataTobBeConverted, database, databaseColumn):
    datasetColSpec = dataset[datasetColSpec]
    result = []
    for i in range(len(datasetColSpec)):
        searchFor = datasetColSpec[i]
        for index, row in database.iterrows():  # iterate over the species name in the database
            if searchFor == database[databaseColumn][index]:
                print(f"Match found: {searchFor}")
                row_info = row.values  # gives row information to define the unit
                if row_info[16] == 'cell':
                    BV = round(row_info[25] * dataset[dataTobBeConverted][0] / 1e+9, 4)
                    result.append(BV)
                    continue
    if len(result) > 0:
        df = pd.DataFrame({'Biovolume': result})
        return df
    else:
        print('No data found')


# define here the dataset and databases for function 1:
df = BiovolumeSpecLevel(
   dataset= SpecData, # define the dataset (the workbook sheet of the data to be converted)
   datasetColSpec= 'spec_name', # Define the column in the sheet where the species name data can be found (string!)
   dataTobBeConverted= 'conc_cells_per_L', # define the column in the sheet where the data to be converted to biovolume
                                          # can be found (string!)
   database= dfBV, # define which database must be used for biovolume conversion
   databaseColumn= 'Species') # Define the column where the database species name can be found (string!)

if df is not None:
    print(df)
else:
    print('No data found')

with pd.ExcelWriter('Biovolume_test.xlsx') as writer:
    df.to_excel(writer, sheet_name='Biovolume', index=False)

print('--------BREAK--------')
#iovolumeWb.save('Biovolume_test.xlsx')

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



# todo what if colony, coenobium or filament? --> biovolume per cell so extra step for calculation
# todo print the data in an excel file, make averages and summations

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
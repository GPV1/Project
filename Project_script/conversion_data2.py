import openpyxl
from openpyxl.styles import Font
import pandas as pd
import numpy as np
import pprint
import sqlite3

from pandas import ExcelWriter

# todo check at the end if all the above imports are needed.


# import datafiles and sheetname. Necessary: database files PEG and Species order and input data
dfBV = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/PEG_BVOL2019_PJ.xlsx',sheet_name='Biovolume file')
dfSpN = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/Species_order_GPV.xlsx',sheet_name='Sheet1')
SpecData = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/Data_2018_GPV.xlsx', sheet_name='Sheet1')
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
Species_other = dfSpN['spec_name'].str.replace('_', ' ')
Species_other_order = dfSpN['order']
Species_other_group = dfSpN['group']
All_Species = pd.concat([pd.Series(Species), pd.Series(Species_other)])


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

# pprint.pprint(species_BV_dict)


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

for spec, ord in zip(Species, Order):
    # print(f"Processing species: {spec}, order: {ord}")
    if ord not in SpOr_merge:
        SpOr_merge[ord] = {spec: species_avg_BV_specieslevel.get(spec, int(0))}
        # print(f"Adding new order: {ord}")
    else:
        SpOr_merge[ord][spec] = species_avg_BV_specieslevel.get(spec, int(0))
        # print(f"Adding species to existing order: {ord}")


other_species_dict = dict(zip(Species_other, Species_other_order))
#pprint.pprint(other_species_dict)

for sp, ord in other_species_dict.items():
    if ord in SpOr_merge:
        SpOr_merge[ord][sp] = other_species_dict.get(ord, 0)
    else:
        SpOr_merge[ord] = {sp: other_species_dict.get(ord, 0)}

# pprint.pprint(SpOr_merge)



SpOr_merge_BV_orderlevel = {}
''' This dictionary creates an average biovolume for each order, based on the nested species biovolume values.

    - the keys order and species are called from the previous made dictionary.
    - A variable is made for the biovolume values in of the species biovolume values.
    - The average biovolume value per order is added and divided by the amount of values, rounded by 4 decimals.
    
    '''
for order, species in SpOr_merge.items():
   values_list = species.values() # create a list of biovolume values for all species in the order
   avg_ord_BV = round(sum(values_list)/len(values_list), 4) # calculate the average biovolume for the order
   SpOr_merge_BV_orderlevel[order] = avg_ord_BV # store the average biovolume in the new dictionary


# pprint.pprint(SpOr_merge_BV_orderlevel)

SpOr_merge_BV_orderlevel_Group = {}
''' Creates the last dictionary needed for the biovolume conversion. This dictionary is an addition to the above dict
    since not all species are defined the same way (due to size classes etc). This dictionary will provide an addition
    in conversion to order level and in later stages to group the data in functional groups.
    '''
GroupOrder = dict(zip(Species_other_order, Species_other_group))

for order, biovolume in SpOr_merge_BV_orderlevel.items():
    for orderdiff, group in GroupOrder.items():
        if orderdiff == order:
            if group in SpOr_merge_BV_orderlevel_Group:
                SpOr_merge_BV_orderlevel_Group[group][order] = biovolume
            else:
                SpOr_merge_BV_orderlevel_Group[group] = {order: biovolume}

# pprint.pprint(SpOr_merge_BV_orderlevel_Group)


SpCl_merge_BV = {}
''' 

'''
for spec, biovolume in SpOr_merge.items():
    for species, Cl in SpCl.items():
        for value in SpOr_merge.values():
            if species in value:
                if Cl in SpCl_merge_BV:
                    SpCl_merge_BV[Cl][spec] = biovolume
                else:
                    SpCl_merge_BV[Cl] = {spec: value}

# pprint.pprint(SpCl_merge_BV)

SpCl_merge_BV_avg = {}

for cls, species in SpCl_merge_BV.items():
    avg_BV_species = []
    for values in species.values():
        avg_BV_species.append(sum(values.values())/len(values))
    avg_class_BV = round(sum(avg_BV_species)/len(avg_BV_species), 4) # calculate the average biovolume for the order
    SpCl_merge_BV_avg[cls] = avg_class_BV # store the average biovolume in the new dictionary

# pprint.pprint(SpCl_merge_BV_avg.items())


# An overview of all dictionaries and their content:
''' Overview of all dictonaries and their functions:

- species_BV_dict:                  Species from the PEG database with their corresponding biovolume values (multiple 
                                    values per species). 
                                    * {Keys: Species, values: biovolume}.
                                     
- species_avg_BV_specieslevel:      Average biovolume for species from the PEG database. 
                                    * {Keys: Species, values: average biovolume per species}.

- SpOr_merge:                       The average biovolume per species is nested in the corresponding order. In this 
                                    dictionary also the species from the database SpeciesOrder are added 
                                    (other_species_dict) since these species are not found in the PEG database (The 
                                    biovolume for these species can only be found on order level).
                                    * {Keys: Order {Keys: Species, values: biovolume}}.
                                    
- SpOr_merge_BV_orderlevel:         For each order the average biovolume is based on the species values inside this 
                                    order.
                                    * { Keys: order, values: average biovolume per order}
                                        
- SpOr_merge_BV_orderlevel_Group:   The functional groups (diatom, dionflagellates, flagellates and other) are added to 
                                    the different taxonomy orders.
                                    * {Keys: group {keys: order, values: biovolume per order}}
                                    
- SpCl_merge_BV:

'''

# todo- tomorrow: fix the group in the functions (spec and class level, order level is fixed).
# todo- Add the carbon calculations to the outcomes (new function) and print it in the same excel
# todo- fix some bugs inside the dictionaries and functions
# todo- comment the funtions and clean up the script.


def BiovolumeSpecLevel (dataset, datasetColSpec, dataTobBeConverted):
    datasetColSpec = dataset[datasetColSpec]
    species = []
    result_spec = []
    Group = []
    for searchFor in datasetColSpec:
        if searchFor in species_avg_BV_specieslevel:
            species.append(searchFor)
            BV = round(species_avg_BV_specieslevel[searchFor] * dataset[dataTobBeConverted][0]/1e+9, 4)
            # print(f"{searchFor} found in PEG database. The biovolume for {searchFor} is {BV} ml")
            result_spec.append(BV)
            # print(searchFor)
            #todo: fix this group thing here..
            group = ''
            for order, spec in SpOr_merge.items():
                for value in spec.keys():
                    if searchFor in value:
                        searchFor_order = order
                        # print(searchFor_order)
                        for ord, grp in SpOr_merge_BV_orderlevel_Group.items():
                            if searchFor_order in grp:
                                group = grp
                                break
                        break
            Group.append(group)
        break
        # else:
        #     print(f"{searchFor} not found in PEG database at any taxonomy level")
    if len(result_spec) > 0:
        df = pd.DataFrame({
            'Species': species,
            'Biovolume (ml)': result_spec,
            'Group': Group
             })
        return df

df = BiovolumeSpecLevel(
        dataset = SpecData,  # define the dataset (the workbook sheet of the data to be converted)
        datasetColSpec = 'spec_name',  # Define the col in the sheet where the species name data can be found (string!)
        dataTobBeConverted = 'conc_cells_per_L')  # define the column in the sheet where the data to be converted to
        # biovolume can be found (string!)



def BiovolumeOrderLevel(dataset, datasetColSpec, dataTobBeConverted):
    datasetColSpec = dataset[datasetColSpec]
    species = []
    order = []
    result_ord = []
    Group = []
    for searchFor in datasetColSpec:
        if searchFor not in species_avg_BV_specieslevel:
            for ord, species_dict in SpOr_merge.items():
                if searchFor in species_dict:
                    searchFor_order = ord
                    BV_order = round(SpOr_merge_BV_orderlevel[searchFor_order] * dataset[dataTobBeConverted][0]/1e+9, 4)
                    print(f"Biovolume of {searchFor} found at order level ({ord}): {BV_order} ml")
                    if BV_order > 0.0:
                        species.append(searchFor)
                        order.append(searchFor_order)
                        result_ord.append(BV_order)
                        for ord in SpOr_merge_BV_orderlevel_Group.values():
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
            'Order': order,
            'Group' : Group,
            'Biovolume (ml)': result_ord,
             })
        return df2

df2 = BiovolumeOrderLevel(
        dataset = SpecData,  # define the dataset (the workbook sheet of the data to be converted)
        datasetColSpec = 'spec_name',  # Define the col in the sheet where the species name data can be found (string!)
        dataTobBeConverted = 'conc_cells_per_L')  # define the column in the sheet where the data to be converted to
        # biovolume can be found (string!)


def BiovolumeClassLevel(dataset, datasetColSpec, dataTobBeConverted):
    datasetColSpec = dataset[datasetColSpec]
    species = []
    Class = []
    result_Class = []
    for searchFor in datasetColSpec:
        for ord, species_dict in SpOr_merge.items():
            if searchFor in species_dict:
                searchFor_order = ord
                # print(searchFor_order)
                BV_order = round(SpOr_merge_BV_orderlevel[searchFor_order] * dataset[dataTobBeConverted][0] / 1e+9, 4)
                if BV_order == 0:
                    # print(f"{searchFor} biovolume cannot be calculated on order {searchFor_order} level")
                    for cls in SpCl_merge_BV_avg:
                        species.append(searchFor)
                        searchFor_Class = cls
                        Class.append(searchFor_Class)
                        BV_Class = round(SpCl_merge_BV_avg[searchFor_Class] * dataset[dataTobBeConverted][0] / 1e+9, 10)
                        # print(f"Biovolume of {searchFor} found at class level ({cls}): {BV_Class} ml")
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
        dataset = SpecData,  # define the dataset (the workbook sheet of the data to be converted)
        datasetColSpec = 'spec_name',  # Define the col in the sheet where the species name data can be found (string!)
        dataTobBeConverted = 'conc_cells_per_L')  # define the column in the sheet where the data to be converted to
        # biovolume can be found (string!)


# if df is not None:
#     print(df)
# elif d2 is not None:
#     print(df2)
# elif df3 is not None:
#     print(df3)
# else:
#     print('No data found')

# print(df, df2, df3)

# creates an excel file with the data from function 1.

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



print(df)






















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

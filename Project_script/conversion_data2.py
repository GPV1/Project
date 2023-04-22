import openpyxl
from openpyxl.styles import Font
import pandas as pd
import numpy as np
import pprint

dfBV = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/PEG_BVOL2019_PJ.xlsx',sheet_name='Biovolume file')
dfSpN = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/Species_order_GPV.xlsx',sheet_name='Sheet1')
SpecData = pd.read_excel('C:/GIS_Course/EGM722/Project/Data_files/Data_2018_GPV.xlsx', sheet_name='Sheet1')

print(SpecData.columns)
print(dfBV.columns)

# Is the name in the column spec_name (SpecData) the same as in the column Species (dfBV) then check if the unit is cell
# if cell then multiply biovolume * cells/Liter

#for index, row in SpecData.iterrows():
    #pprint.pprint(SpecData['spec_name'].index)

# todo average biovolumes, spec_name iterating over data file instead of fill in--> same for slice in conc cell/L
# todo what if colony or coenobium or filament? --> biovolume per cell so extra step for calculation
# todo write a function for this whole thing

spec_name = 'Choanoflagellatea'

for index, row in dfBV.iterrows():
    if spec_name == row['Species']:
        row_info = row.values
        if row_info[16] == 'cell':
            BV = round(row_info[25] * SpecData['conc_cells_per_L'][0]/1e+9, 4) # *1000/1e+12 to convert µm3/L to ml,
            # with 3 decimals
            print('Biovolume for {} is {} ml'.format(spec_name, BV))
#break

print(BV)
            #elif row_info[16] == 'colony':
                #continue
            #else:
                #print('NA')

   # print(f"{index}: {row['Species']}")

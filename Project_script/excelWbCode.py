''' Create a new workbook '''
newwb = openpyxl.Workbook()
sheet = newwb.active
sheet.title = 'test'
print(newwb.sheetnames)

''' Save new workbook '''
newwb.save('test.xlsx')

''' Writing values to cells '''
testwb = openpyxl.load_workbook('test.xlsx')
print(testwb.sheetnames)

sheettest = testwb['test']
sheettest['A1'] = 'Check if this is written in the cell'
print(sheettest['A1'].value)

testwb.save('test.xlsx')

'''Setting the font style'''
fontobj1 = Font(name='Calibri', size=20, bold=True, italic=False)
sheettest['A1'].font = fontobj1

testwb.save('test.xlsx')

'''Using Formulas'''
sheettest['A2'] = 200
sheettest['A3'] = 300
sheettest['A4'] = '=SUM(A2:A3)'

testwb.save('test.xlsx')

'''Setting row height and column width'''
sheettest.row_dimensions[1].height = 30
sheettest.column_dimensions['A'].width = 60

testwb.save('test.xlsx')

'''Merging cells'''
sheettest.merge_cells('A1:C1')

testwb.save('test.xlsx')

# unmerge cells
sheettest.unmerge_cells('A1:C1')

testwb.save('test.xlsx')

'''Freezing Panes'''
sheettest.freeze_panes = 'A2'

testwb.save('test.xlsx')

'''Creating charts'''
# get some data in the test sheet
for i in range(1, 11):
    sheettest['C'+str(i)] = i

# create a reference in the sheet to put the graph in
refObj = openpyxl.chart.Reference(sheettest, min_col=3, min_row=1, max_col=3, max_row=10)

# create a series reference by passing in the reference object
seriesObj = openpyxl.chart.Series(refObj, title='First series')

# create a chart object
chartObj = openpyxl.chart.BarChart()
chartObj.tilte = 'Test chart'
chartObj.append(seriesObj)

# Add the chart to the sheet
sheettest.add_chart(chartObj, 'D1')

testwb.save('test.xlsx')

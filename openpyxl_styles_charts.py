#! python3
# 
# Charts in Spreadsheets
import openpyxl 
wb = openpyxl.Workbook()
sheet = wb.active
for i in range(1, 11): # create data in ColA
    sheet['A' + str(i)] = i
# MISSING END OF NEXT LINE
refObj = openpyxl.chart.Reference(sheet, min_col=1, min_row=1, max_col=1, max_row=1)
seriesObj = openpyxl.chart.Series(refObj, title='First series')
chartObj = openpyxl.chart.BarChart()
chartObj.title = 'My Chart'
chartObj.append(seriesObj)
sheet.add_chart(chartObj, 'C5')
wb.save('sampleChart.xlsx')




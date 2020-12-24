from makeCharts_1 import DataSheet, convert_to_BarChart, convert_to_BarChart3D, convert_to_Scatter, Reference
import openpyxl as xl
from openpyxl.chart import Reference, ScatterChart, Series

wb = xl.load_workbook("transactions.xlsx")

# Try it
data_1 = DataSheet('Sheet12', wb)
data_1.get_suitable_chart("a2", 4, 'Fruits')
convert_to_BarChart3D("Sheet1112", wb)
convert_to_Scatter("Sheet1", wb, col_num_x= 2, col_num_y= 3) # col_num_x is the second column of the excel sheet,
# with x-values and col_num_y (3rd column) has y-values
convert_to_Scatter("Sheet1", wb, col_num_x= 1, col_num_y= 2) # col_num_x is the first column of the excel sheet,
# with x-values and col_num_y (2nd column) has y-values
convert_to_BarChart("Sheet12",wb,2,4,1,4,1,2,4)


wb.save("transactions.xlsx")
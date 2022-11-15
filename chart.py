from openpyxl.workbook import Workbook
from openpyxl import load_workbook
from openpyxl.chart import PieChart, PieChart3D, Reference, BarChart, BarChart3D, LineChart, LineChart3D

# Create a workbook object
wb = load_workbook('names3.xlsx')

# Create an active worksheet
ws = wb.active

#Determine the type of chart, it can be changed with all others from line 3
chart = PieChart()

#Designate Labels and Data
labels = Reference(ws, min_col=1, max_col=1, min_row=2, max_row=5)
data = Reference(ws, min_col=4, max_col=4, min_row=1, max_row=5)

#Put it all together
chart.add_data(data, titles_from_data=True)
chart.set_categories(labels)
chart.title = "Employee Salaries"

# Place the chart in the spreadsheet
ws.add_chart(chart, "E2")

# Save an excel spreadsheet
wb.save('names3.xlsx')
print(' File was saved...')
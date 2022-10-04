import pandas as pd
import os
import xlsxwriter as xl

def createexcel(pname, date, time):

    #Creating Pandas DataFrames from Txt files
    summary = pd.read_csv('text11.txt', sep = ':', header = 0, index_col= False, names = ['Device Class', 'Quantity'])
    overall = pd.read_csv('text111.txt', sep = '|', skiprows = 4, header = 0, index_col= False, names = ['Component Class', 'Total on Board', 'Presence Tested', 'Presence Tested (%)', 'Total Pins', 'Untested Pins', 'Total Elements', 'Untested Elements', 'Coverage %', 'Presence untested'])
    icsummary = pd.read_csv('text1111.txt', sep = '|', skiprows = 4, header = 0, index_col= False, names = ['IC Name', 'Optical Test', 'Total Pins', 'Tested Pins', 'VCC Pins', 'GND Pins', 'No Access Pins', 'Parallel Pins', 'NC Pins', 'Coverage %', 'CScan Tested', 'ChipScan Tested', 'Diode Tested', 'Boundary Tested'])
    connectorsummary = pd.read_csv('text11111.txt', sep = '|', skiprows = 4, header = 0, index_col= False, names = ['Connector Name', 'Optical Test', 'Total Pins', 'Tested Pins', 'VCC Pins', 'GND Pins', 'No Access Pins', 'Parallel Pins', 'NC Pins', 'LINK Tested', '% LINK Tested', '% Coverage'])
    alldevices = pd.read_csv('text111111.txt', sep = '|', skiprows = 4, header = 0, index_col= False, names = ['Name', 'Class', 'Electrical Test', 'Optical Test', 'Comment'])
    icpin = pd.read_csv('text1111111.txt', sep = '|', skiprows = 3, header = 0, index_col= False, names = ['Name', 'Pin Number', 'Diode & ChipScan', 'CScan', 'BoundaryScan'])
    connectorpin = pd.read_csv('text11111111.txt', sep = '|', skiprows = 3, header = 0, index_col= False, names = ['Name', 'Pin Number', 'CScan', 'Link'])
    notbds = pd.read_csv('text111111111.txt', sep = '|', skiprows = 3, header = 0, index_col= False, names = ['Name', 'Class', 'Enabled', 'Comment'])

    #Creating Excel File with multiple Sheets for each separate Table
    with pd.ExcelWriter('TestCoverageReport.xlsx', engine = 'xlsxwriter') as writer:
        summary.to_excel(writer, sheet_name='Board', index = False, startrow = 5)
        workbook = writer.book
        worksheet = writer.sheets['Board']
        worksheet.write(1, 0, 'INTEGRATOR - COVERAGE REPORT')
        worksheet.write(2, 0, 'Project: ' + pname)
        worksheet.write(3, 0, 'Date: ' + date)
        worksheet.write(3, 1, 'Time: ' + time)
        cf = workbook.add_format({'bold' : True, 'italic' :  True })
        worksheet.write(4, 0, 'Board Summary', cf)
        worksheet.set_column(0, 1, 25)
        worksheet.set_row(0, 90)
        worksheet.insert_image('A1','acculogiclogo.png')
        chart = workbook.add_chart({'type' : 'column'})
        (max_row, max_col) = summary.shape
        chart.add_series({'values': ['Board', 6, 1, 6, 1], 'name' : ['Board', 6, 0]})
        chart.add_series({'values': ['Board', 7, 1, 7, 1], 'name' : ['Board', 7, 0]})
        chart.add_series({'values': ['Board', 8, 1, 8, 1], 'name' : ['Board', 8, 0]})
        chart.add_series({'values': ['Board', 9, 1, 9, 1], 'name' : ['Board', 9, 0]})
        chart.add_series({'values': ['Board', 10, 1, 10, 1], 'name' : ['Board', 10, 0]})
        chart.add_series({'values': ['Board', 11, 1, 11, 1], 'name' : ['Board', 11, 0]})
        chart.add_series({'values': ['Board', 12, 1, 12, 1], 'name' : ['Board', 12, 0]})
        chart.add_series({'values': ['Board', 13, 1, 13, 1], 'name' : ['Board', 13, 0]})
        chart.add_series({'values': ['Board', 14, 1, 14, 1], 'name' : ['Board', 14, 0]})
        chart.add_series({'values': ['Board', 15, 1, 15, 1], 'name' : ['Board', 15, 0]})
        chart.add_series({'values': ['Board', 16, 1, 16, 1], 'name' : ['Board', 16, 0]})
        chart.add_series({'values': ['Board', 17, 1, 17, 1], 'name' : ['Board', 17, 0]})
        chart.set_x_axis({'visible' : False})
        worksheet.insert_chart(1, 3, chart)

        overall.to_excel(writer, sheet_name='Overall', index = False, startrow = 5)
        workbook = writer.book
        worksheet = writer.sheets['Overall']
        worksheet.write(1, 0, 'INTEGRATOR - COVERAGE REPORT')
        worksheet.write(2, 0, 'Project: ' + pname)
        worksheet.write(3, 0, 'Date: ' + date)
        worksheet.write(3, 1, 'Time: ' + time)
        worksheet.write(4, 0, 'Overall', cf)
        worksheet.set_column(0, 10, 20)
        worksheet.set_row(0, 90)
        worksheet.insert_image('A1','acculogiclogo.png')
        

        icsummary.to_excel(writer, sheet_name='ICs', index = False, startrow = 5)
        workbook = writer.book
        worksheet = writer.sheets['ICs']
        worksheet.write(1, 0, 'INTEGRATOR - COVERAGE REPORT')
        worksheet.write(2, 0, 'Project: ' + pname)
        worksheet.write(3, 0, 'Date: ' + date)
        worksheet.write(3, 1, 'Time: ' + time)
        worksheet.write(4, 0, 'ICs', cf)
        worksheet.set_column(0, 14, 20)
        worksheet.set_row(0, 90)
        worksheet.insert_image('A1','acculogiclogo.png')

        connectorsummary.to_excel(writer, sheet_name='Connectors', index = False, startrow = 5)
        workbook = writer.book
        worksheet = writer.sheets['Connectors']
        worksheet.write(1, 0, 'INTEGRATOR - COVERAGE REPORT')
        worksheet.write(2, 0, 'Project: ' + pname)
        worksheet.write(3, 0, 'Date: ' + date)
        worksheet.write(3, 1, 'Time: ' + time)
        worksheet.write(4, 0, 'Connectors', cf)
        worksheet.set_column(0, 12, 20)
        worksheet.set_row(0, 90)
        worksheet.insert_image('A1','acculogiclogo.png')

        alldevices.to_excel(writer, sheet_name='All_Devices', index = False, startrow = 5)
        workbook = writer.book
        worksheet = writer.sheets['All_Devices']
        worksheet.write(1, 0, 'INTEGRATOR - COVERAGE REPORT')
        worksheet.write(2, 0, 'Project: ' + pname)
        worksheet.write(3, 0, 'Date: ' + date)
        worksheet.write(3, 1, 'Time: ' + time)
        worksheet.write(4, 0, 'All Devices', cf)
        worksheet.set_column(0, 5, 40)
        worksheet.set_row(0, 90)
        worksheet.insert_image('A1','acculogiclogo.png')

        icpin.to_excel(writer, sheet_name='IC_Pins', index = False, startrow = 5)
        workbook = writer.book
        worksheet = writer.sheets['IC_Pins']
        worksheet.write(1, 0, 'INTEGRATOR - COVERAGE REPORT')
        worksheet.write(2, 0, 'Project: ' + pname)
        worksheet.write(3, 0, 'Date: ' + date)
        worksheet.write(3, 1, 'Time: ' + time)
        worksheet.write(4, 0, 'IC Pins', cf)
        worksheet.set_column(0, 5, 20)
        worksheet.set_row(0, 90)
        worksheet.insert_image('A1','acculogiclogo.png')

        connectorpin.to_excel(writer, sheet_name='Connector_Pins', index = False, startrow = 5)
        workbook = writer.book
        worksheet = writer.sheets['Connector_Pins']
        worksheet.write(1, 0, 'INTEGRATOR - COVERAGE REPORT')
        worksheet.write(2, 0, 'Project: ' + pname)
        worksheet.write(3, 0, 'Date: ' + date)
        worksheet.write(3, 1, 'Time: ' + time)
        worksheet.write(4, 0, 'Connector Pins', cf)
        worksheet.set_column(0, 4, 15)
        worksheet.set_row(0, 90)
        worksheet.insert_image('A1','acculogiclogo.png')

        notbds.to_excel(writer, sheet_name='Not_In_BDS', index = False, startrow = 5)
        workbook = writer.book
        worksheet = writer.sheets['Not_In_BDS']
        worksheet.write(1, 0, 'INTEGRATOR - COVERAGE REPORT')
        worksheet.write(2, 0, 'Project: ' + pname)
        worksheet.write(3, 0, 'Date: ' + date)
        worksheet.write(3, 1, 'Time: ' + time)
        worksheet.write(4, 0, 'Not In BDS', cf)
        worksheet.set_column(0, 4, 40)
        worksheet.set_row(0, 90)
        worksheet.insert_image('A1','acculogiclogo.png')

    #Removing files created by textsplit when they're no longer needed
    os.remove("text1.txt")
    os.remove("text11.txt")
    os.remove("text111.txt")
    os.remove("text1111.txt")
    os.remove("text11111.txt")
    os.remove("text111111.txt")
    os.remove("text1111111.txt")
    os.remove("text11111111.txt")
    os.remove("text111111111.txt")

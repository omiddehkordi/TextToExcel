# TxtToExcel


1) The executable Python file is txttoexcel.py. The rest are supporting files with supporting functions.

2) Run the program with ./txttoexcel.py in your Terminal

3) 4 back-to-back windows will appear asking for the following information: Test Coverage Report filepath, Project Name, Date, Time

4) The Program will output an Excel File into the Project Directory

The program relies on the structure of the Report following a structure that allows for splitting the sections according to the rule in textsplit.py (Separating when the line starts with + which indicates a header) and that allows splitting the tables according to the sep= parameter in the pd.read_csv function in excel.py

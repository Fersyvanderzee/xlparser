import openpyxl
import os

file_name = 'result_ddd.txt'
path = 'C:/Dev/Wiltec_tools/files'

result_file = open(file_name, 'w')

print("Iteration started.")

for file in os.listdir(path):

    print(f"Parsing file {file}.")

    try:
        dataframe = openpyxl.load_workbook(f'files/{file}').active
        
        row = 1

        for _ in range(200):
            if dataframe[f'A{row}'].value == "Accountnummer":
                debtornr = dataframe[f'B{row}'].value
            
            if dataframe[f'A{row}'].value == "Dealernummer":
                dealernr = dataframe[f'B{row}'].value

            if dataframe[f'A{row}'].value == "Opmerkingen":
                if dataframe[f'A{row}'].value != None:
                    comments = dataframe[f'B{row}'].value
                else:
                    comments = ""

            if dataframe[f'A{row}'].value == "Pakket 1":       
                start_row = row

            row += 1

        row = start_row

        for _ in range(200):
            if dataframe[f'C{row}'].value != None:
                if dataframe[f'C{row}'].value != 'Artikelnummer':
                    if dataframe[f'D{row}'].value != 0:
                        result_file.write(f"{debtornr};{dealernr};{dataframe[f'C{row}'].value};{dataframe[f'D{row}'].value};{dataframe[f'E{row}'].value};{comments}\n")
            row += 1

    except PermissionError:
        print(f"PermissionError - Permission denied: {file}")

print("Iteration done.")

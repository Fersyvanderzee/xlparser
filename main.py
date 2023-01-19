import openpyxl
import os

file_name = ''
path = ''

result_file = open(file_name, 'w')

print("Iteration started.")

for file in os.listdir(path):

    print(f"Parsing file {file}.")

    try:
        dataframe = openpyxl.load_workbook(f'files/{file}').active

        debtornr = dataframe['B5'].value
        dealernr = dataframe['B7'].value

        row = 18

        for _ in range(41):
            if dataframe[f'C{row}'].value != None:
                if dataframe[f'C{row}'].value != 'Artikelnummer':
                    if dataframe[f'D{row}'].value != 0:
                        result_file.write(f"{debtornr};{dealernr};{dataframe[f'C{row}'].value};{dataframe[f'D{row}'].value};{dataframe[f'E{row}'].value}\n")
            row += 1

    except PermissionError:
        print(f"PermissionError - Permission denied: {file}")

print("Iteration done.")

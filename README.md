# Parts-List-Compiler
### Python program to create an organized xlsx file based on a master parts list

## 1. Editing Source Code

    >The source code for the program is found in the FastenalCompiler.py. The main packages used are: Pandas, tkinter, xlsxwriter.
    Some familiarity with these packages is important. The basic structure of the program is as follows:

        1. import csv
        2. clean and break apart data
        3. write to excel file and format

## 2.  Compiling Source Code

    >Source code can be compiled using the PyInstaller Program, installed via pip. https://pyinstaller.org/en/stable/#
    to compile, navigate to the folder with the source code and run the following command in a terminal emulator:

    >PyInstaller -F FastenalCompiler.py

    The program will now comile the source. The new executible can be found in the 'dist' folder.

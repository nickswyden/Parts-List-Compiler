#imports
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
import os
import xlsxwriter

#desc dict (only for fasteners)
desc_list = [
    '800', '801', '802', '803', '804', '805', '806', '807', '808', '809',
    '810', '811', '812', '813', '814', '815', '816', '817', '818', '819',
    '820', '821', '822', '823', '825', '826', '827', '828', '829', '830',
    '831', '832', '833', '834', '836', '837', '838', '839', '840', '841',
    '842', '843', '844', '848', '850', '851', '852', '853', '854', '858',
    '861', '871', '889', '890', '891', '892'
]

desc_texts = [
    'MISCELLANEOUS FASTENERS', 'SCREWS', 'BOLTS', 'NUTS', 'WASHERS', 'PINS',
    'U-BOLTS', 'SPRINGS', 'SPROCKETS, GEARS', 'ROLLER CHAIN',
    'CYLINDERS, VALVES, SEAL KITS & MANIFOLDS', 'HYDRAULIC FITTINGS & HOSES',
    'CASTINGS', 'CARTONS & PALLETS', 'TIRES & RIMS', 'HUBS & SPINDLES',
    'RUBBER PARTS FULL GO TO 836C', 'PLASTIC PARTS', 'DECALS', 'NAME PLATES',
    'OPENER TILLAGE COMPONENTS', 'PAINT, GREASE, OIL, SOAP', 'BEARINGS',
    'OPTIONAL PRODUCT EQUIPMENT', 'AUSHERMAN FASTENERS',
    'PTOS, GEAR BOXES & ENGINES', 'SPACERS, BUSHINGS & BALLS',
    'AUSHERMAN RESALE COMPONENTS', 'LOW PRESSURE VALVES & GAUGES',
    'LOW PRESSURE FITTINGS & ADAPTORS', 'PUMPS, FILTERS & ACCUMULATORS',
    'SPRAYER NOZZLES', 'ELECTRONIC COMPONENTS',
    'FAB PARTS REQUIRING ADDITIONAL LABOR', 'RUBBER PARTS', 'PLASTIC PARTS',
    'DECALS', 'EXCEL MOWER PARTS/KELLY BACKHOE PARTS',
    'VEHICLES PURCHASE COMPONENTS', 'HYDRAULIC FITTINGS & HOSES', 'BOLTS',
    'ELECTRONIC COMPONENTS', 'DECALS', 'DECALS',
    'CYLINDERS, VALVES, SEALKITS', 'HYDRAULIC FITTINGS & HOSES',
    'BOLTS (ROHS)', 'NUTS (ROHS)', 'WASHERS (ROHS)', 'DECALS',
    'HYDRAULIC FITTING AND HOSES', 'HYDRAULIC FITTING AND HOSES',
    'SALES DEPARTMENT MISC. CATEGORY', 'MISCELLANEOUS',
    'MISC PARTS INCLUDING ELECTRONIC ACRE METER', 'MISCELLANEOUS'
]

descdict = {key: value for key, value in zip(desc_list, desc_texts)}

def main():    
    #obtain file path
    global path
    global dfs
    path = get_file_name()
    masterdf = pd.read_csv(path)

    #clean master
    masterdf = clean_master_data(masterdf)

    #break into C D H E
    letters = ["C", "D", "E", "H"]
    dfs = {letter + "df": split_data(masterdf, letter) for letter in letters}

    #Merge E and H into one H
    dfs['Hdf'] = pd.concat([dfs['Hdf'],dfs['Edf']]) 

    #arrange titles
    dfs["Ddf"] = arrange_titles(dfs['Ddf'], False)
    dfs["Hdf"] = arrange_titles(dfs['Hdf'], False)

    #add fastener type and populate
    dfs['Cdf']['Fastener Type'] = ''
    dfs['Cdf']['Fastener Type'] = dfs['Cdf']['Number'].str[:3].map(descdict)
    dfs['Cdf'] = arrange_titles(dfs['Cdf'], True)

    write_to_excel()
    
    return path, dfs
    
def get_file_name(): #returns file name from windows dialogue box
    #draw file selection window
    root = tk.Tk()
    root.withdraw()

    #select csv
    file_path = filedialog.askopenfilename()
    return file_path

def clean_master_data(masterdf): # drops unneeded data from masterdf, returns new cleaned masterdf
    #remove File extensions
    masterdf['Number'] = masterdf['Number'].str.replace('.ASM', '')
    masterdf['Number'] = masterdf['Number'].str.replace('.PRT', '')

    #remove unneeded columns
    masterdf = masterdf.drop(["Level","Version","State","File Name"], axis=1)

    #remove everything but C, D, H & E
    masterdf = masterdf[masterdf['Number'].str.contains("A|B|F|S|K|_|I|J") == False]

    #add empty Columns
    masterdf['Color'] = ''
    masterdf['Notes'] = ''

    #sum dupes
    masterdf = masterdf.groupby(['Number','Name','Color','Notes'],as_index=False).agg({'Quantity': 'sum'})
    
    return masterdf

def split_data(df, x): #creates new df based Number column and a char, x, and returns
    newdf = df[df['Number'].str[-1] == x] 
    return newdf

def arrange_titles(df,long_title): #rearranges the column headers since they are out of order
    if long_title == True:
        df = df[['Number','Name','Quantity','Color','Fastener Type','Notes']]
    else:
        df = df[['Number','Name','Quantity','Color','Notes']]
    return df

def write_to_excel(): # does the bulk of formatting and writes the dataframes to an .xlsx
    output_name = os.path.basename(path[:-4])
    file_name = output_name + '_output0.xlsx'
    #i = int(file_name[:-6]) + 1
    #if (os.path.isfile(file_name) == True): FIXME: Add functionality to create new sheet with existing sheet in directory (version 1,2,3,...)
    #    file_name = output_name + '_output'+ i +'.xlsx'
    writer = pd.ExcelWriter(file_name, engine="xlsxwriter")
    workbook  = writer.book

    #create sheets
    dataframes = {
        'C - Fasteners': dfs['Cdf'],
        'D - Fabricated': dfs['Ddf'],
        'H E - Weldment': dfs['Hdf']
    }

    for sheet_name, dataframe in dataframes.items():
        dataframe.to_excel(writer, index=False, startrow=4, sheet_name=sheet_name)

    worksheet_dict = {name: writer.sheets[name] for name in ['C - Fasteners', 'D - Fabricated', 'H E - Weldment']}
    worksheet_C, worksheet_D, worksheet_H = worksheet_dict.values()

    #define merged colors
    colors_format_o = workbook.add_format(
        {
            "border": 1,
            "align": "center",
            "fg_color": "#ff6600",
        }
    )
    colors_format_b = workbook.add_format(
        {
            "border": 1,
            "align": "center",
            "fg_color": "#6699ff",
        }
    )
    colors_format_g = workbook.add_format(
        {
            "border": 1,
            "align": "center",
            "fg_color": "#00cc66",
        }
    )

    #create colors at the top of sheet
    worksheets = [worksheet_C, worksheet_D, worksheet_H]
    titles = ["Common", "New/Rarely Used", "Change Proposal"]
    colors_formats = [colors_format_o, colors_format_b, colors_format_g]

    for i, worksheet in enumerate(worksheets): #FIXME: Colors are messed up (each ws only has one color)
        for row, title in enumerate(titles):
            worksheet.merge_range(f"A{row+1}:F{row+1}", title, colors_formats[row])


    #set cell widths
    worksheets = [worksheet_C, worksheet_D, worksheet_H]
    column_widths = [(0, 10), (1, 35), (2, 8), (3, 8), (4, 32), (5, 35)]

    for worksheet in worksheets:
        for col, width in column_widths:
            worksheet.set_column(col, col, width)


    #Write Data to Sheets
    #format data as table
    worksheet_C.add_table('A5:F' + str(len(dfs['Cdf'].index) + 5),{'style': 'Table Style Medium 15',
                                                            'columns': [{'header': 'Number'},
                                                                        {'header': 'Name'},
                                                                        {'header': 'Quantity'},
                                                                        {'header': 'Color'},
                                                                        {'header': 'Fastener Type'},
                                                                        {'header': 'Notes'}]} )

    worksheet_D.add_table('A5:E' + str(len(dfs['Ddf'].index) + 5),{'style': 'Table Style Medium 15',
                                                            'columns': [{'header': 'Number'},
                                                                        {'header': 'Name'},
                                                                        {'header': 'Quantity'},
                                                                        {'header': 'Color'},
                                                                        {'header': 'Notes'}]} )

    worksheet_H.add_table('A5:E' + str(len(dfs['Hdf'].index) + 5),{'style': 'Table Style Medium 15',
                                                            'columns': [{'header': 'Number'},
                                                                        {'header': 'Name'},
                                                                        {'header': 'Quantity'},
                                                                        {'header': 'Color'},
                                                                        {'header': 'Notes'}]} )

    writer.close()

main()
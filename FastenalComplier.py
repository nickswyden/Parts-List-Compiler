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

#draw file selection window
root = tk.Tk()
root.withdraw()

#select csv
file_path = filedialog.askopenfilename()
masterdf = pd.read_csv(file_path)

#clean data and format to specs
#remove File extensions
masterdf['Number'] = masterdf['Number'].str.replace('.ASM', '')
masterdf['Number'] = masterdf['Number'].str.replace('.PRT', '')

#remove unneeded columns
masterdf = masterdf.drop("Level", axis=1)
masterdf = masterdf.drop("Version", axis=1)
masterdf = masterdf.drop("State", axis=1)
masterdf = masterdf.drop('File Name', axis=1)

#remove everything but C, D, H & E
masterdf = masterdf[masterdf['Number'].str.contains("A|B|F|S|K|_|I|J") == False]

#add empty Columns
masterdf['Color'] = ''
masterdf['Notes'] = ''

#sum dupes
masterdf = masterdf.groupby(['Number','Name','Color','Notes'],as_index=False).agg({'Quantity': 'sum'})

#break into C D H and sum duplicates
Cdf = masterdf[masterdf['Number'].str[-1] == 'C']
Ddf = masterdf[masterdf['Number'].str[-1] == 'D']
Ddf = Ddf[['Number','Name','Quantity','Color','Notes']]
Hdf = masterdf[masterdf['Number'].str[-1] == 'H'] #FIXME: implement sorting for E weldments
Hdf = Hdf[['Number','Name','Quantity','Color','Notes']]
Edf = masterdf[masterdf['Number'].str[-1] == 'E']
Edf = Edf[['Number','Name','Quantity','Color','Notes']]
Hdf = pd.concat([Hdf,Edf]) #Merge E and H into one H

#add fastener type and populate
Cdf['Fastener Type'] = ''
Cdf['Fastener Type'] = Cdf['Number'].str[:3].map(descdict)
Cdf = Cdf[['Number','Name','Quantity','Color','Fastener Type','Notes']]

#write to file
output_name = os.path.basename(file_path[:-4])
writer = pd.ExcelWriter(output_name + '_output.xlsx', engine="xlsxwriter")
Cdf.to_excel(writer, index= False,startrow=4,sheet_name= 'C - Fasteners')
Ddf.to_excel(writer, index= False,startrow=4,sheet_name= 'D - Fabricated')
Hdf.to_excel(writer, index= False,startrow=4,sheet_name= 'H E - Weldment')
workbook  = writer.book
worksheet_C = writer.sheets['C - Fasteners']
worksheet_D = writer.sheets['D - Fabricated']
worksheet_H = writer.sheets['H E - Weldment']

#format file
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
worksheet_C.merge_range("A1:F1", "Common", colors_format_o)
worksheet_C.merge_range("A2:F2", "New/Rarely Used", colors_format_b)
worksheet_C.merge_range("A3:F3", "Change Proposal", colors_format_g)

worksheet_D.merge_range("A1:F1", "Common", colors_format_o)
worksheet_D.merge_range("A2:F2", "New/Rarely Used", colors_format_b)
worksheet_D.merge_range("A3:F3", "Change Proposal", colors_format_g)

worksheet_H.merge_range("A1:F1", "Common", colors_format_o)
worksheet_H.merge_range("A2:F2", "New/Rarely Used", colors_format_b)
worksheet_H.merge_range("A3:F3", "Change Proposal", colors_format_g)

#set cell widths
worksheet_C.set_column(0, 0, 10)
worksheet_C.set_column(1, 1, 35)
worksheet_C.set_column(2, 3, 8)
worksheet_C.set_column(4, 4, 32)
worksheet_C.set_column(5, 5, 35)

worksheet_D.set_column(0, 0, 10)
worksheet_D.set_column(1, 1, 35)
worksheet_D.set_column(2, 3, 8)
worksheet_D.set_column(4, 4, 32)
worksheet_D.set_column(5, 5, 35)

worksheet_H.set_column(0, 0, 10)
worksheet_H.set_column(1, 1, 35)
worksheet_H.set_column(2, 3, 8)
worksheet_H.set_column(4, 4, 32)
worksheet_H.set_column(5, 5, 35)

#Write Data to Sheets


#format data as table
worksheet_C.add_table('A5:F' + str(len(Cdf.index) + 5),{'style': 'Table Style Medium 15',
                                                        'columns': [{'header': 'Number'},
                                                                    {'header': 'Name'},
                                                                    {'header': 'Quantity'},
                                                                    {'header': 'Color'},
                                                                    {'header': 'Fastener Type'},
                                                                    {'header': 'Notes'}]} )

worksheet_D.add_table('A5:E' + str(len(Ddf.index) + 5),{'style': 'Table Style Medium 15',
                                                        'columns': [{'header': 'Number'},
                                                                    {'header': 'Name'},
                                                                    {'header': 'Quantity'},
                                                                    {'header': 'Color'},
                                                                    {'header': 'Notes'}]} )

worksheet_H.add_table('A5:E' + str(len(Hdf.index) + 5),{'style': 'Table Style Medium 15',
                                                        'columns': [{'header': 'Number'},
                                                                    {'header': 'Name'},
                                                                    {'header': 'Quantity'},
                                                                    {'header': 'Color'},
                                                                    {'header': 'Notes'}]} )

writer.close()
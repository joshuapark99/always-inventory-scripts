import openpyxl
from pathlib import Path
import glob
from collections import defaultdict



# takes too long

masterSheet = Path("Master_List_With_ASIN.xlsx")

dest = "Master_List_With_All_Info.xlsx"

workbooks = defaultdict(lambda: "not yet")

other_info = ['item_dimensions_unit_of_measure', 'website_shipping_weight', 'website_shipping_weight_unit_of_measure', 'item_length', 'item_width', 'item_height','fabric_type']

lastUsedWBPath = None
wb = None
style_sheet = None

def getInfo(sku,path,row):
    global lastUsedWBPath
    global wb
    global style_sheet
    object = defaultdict(lambda: "not here")
    if(lastUsedWBPath == None or lastUsedWBPath != path):
        print("NEW PATH")
        if(wb != None):
            wb.close()
        wb = openpyxl.load_workbook(Path(path),read_only=True)
        lastUsedWBPath = path
        style_sheet = ''
        sheets = wb.sheetnames;
        #object = defaultdict(lambda: "not here")
        if('Template' in sheets):
            style_sheet = wb['Template']
        else:
            for sheet in sheets:
                temp = wb[sheet].cell(1,1).value
                if type(temp) == type(' ') and temp[:12] == 'TemplateType':
                    style_sheet = wb[sheet]
                    break
    #wb = openpyxl.load_workbook(Path(path),read_only=True)



    for col2 in range(1,style_sheet.max_column+1):
        if style_sheet.cell(3,col2).value in other_info:
            object[style_sheet.cell(3,col2).value] = style_sheet.cell(row,col2).value
    #wb.close()
    return object

if __name__ == "__main__":
    #global lastUsedWBPath
    m_wb = openpyxl.load_workbook(masterSheet,read_only=True)
    new_wb = openpyxl.Workbook()
    path = ''
    read_sheet = m_wb['Sheet']

    write_sheet = new_wb['Sheet']

    for counter in range(2,read_sheet.max_row+1):
        sku = read_sheet.cell(counter,1).value
        if(path != lastUsedWBPath):
            path = read_sheet.cell(counter,3).value
        row = read_sheet.cell(counter,4).value
        #col = read_sheet.cell(counter,5).value
        object = getInfo(sku, path, row)
        write_sheet.cell(counter,1).value = sku
        write_sheet.cell(counter,2).value = read_sheet.cell(counter,2).value
        write_sheet.cell(counter,3).value = path
        for x,y in enumerate(other_info,start=6):
            if object[y] != "not here":
                write_sheet.cell(counter,x).value = object[y]
        print(f'done with item {counter}/{read_sheet.max_row}')

    new_wb.save(dest)
    new_wb.close()
    m_wb.close()

import xlsxwriter as xl
import pandas as pd
import numpy as np
import random

write_file_loc = "/Users/charlie//Documents/personal_tools/Excel_Tools/excel_files/test_cells.xlsx"

def makeTestFrame():
    print("...making test frame")
    cols = ["ard","bel","cow","effver"]
    row_length = 20
    row_amp = 100

    clist = []
    for i in range(len(cols)):
        row = []
        for j in range(row_length):
            num = random.random() * row_amp
            num = int(num)
            row.append(num)
        row = pd.Series(row,name=cols[i])
        clist.append(row)

    dataframe = pd.concat(clist,axis=1)

    return dataframe

def writeFrame(args):
    args = vars(args)

    dataframe = args["dataframe"]
    file_location = args["file_location"]
    sheet_name = "Main"
    safe = not args["unsafe"]
    offset = args["offset"].split("x")
    for i in range(2):
        offset[i] = int(offset[i])

    print()
    print(type(file_location))
    print()

    if dataframe == "None":
        dataframe = makeTestFrame()

    #safe means checking that the excel file is empty before
    # writing over whatever content it has.
    try:
        if safe:
            safe_check_frame = pd.read_excel(file_location)
            
            if not safe_check_frame.empty:
                print("Excel file not empty! If you still want to overwrite it, run this function with the --unsafe flag!")
                return False

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(file_location, engine='xlsxwriter')

        # Convert the dataframe to an XlsxWriter Excel object.
        dataframe.to_excel(writer, sheet_name=sheet_name,startcol=offset[1],
                       startrow=offset[0])

        # Close the Pandas Excel writer and output the Excel file.
        writer.save()
        
        return True
    
    except:
        
        return False



if __name__ == "__main__":
    print("Start!")

    cols = ["ard","bel","cow","effver"]
    row_length = 20
    row_amp = 100

    clist = []
    for i in range(len(cols)):
        row = []
        for j in range(row_length):
            num = random.random() * row_amp
            num = int(num)
            row.append(num)
        row = pd.Series(row,name=cols[i])
        clist.append(row)

    #dataframe = pd.DataFrame(clist[0])
    #clist[0] = dataframe
    dataframe = pd.concat(clist,axis=1)

    print(dataframe.columns)

    print(writeFrame(dataframe,write_file_loc,safe=False))
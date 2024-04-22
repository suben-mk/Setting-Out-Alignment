# Topic; Cross Fall (%) Array
# Created By; Suben Mukem (SBM) as Survey Engineer.
# Updated; 29/01/2024

# Import Module
import numpy as np
import pandas as pd
import time

t0 = time.time() # Start time

#------------------- Main Cross Fall (%) Array ------------------#
# Path files
Import_data_path = "Import Setting-Out Alignment Data.xlsx"
Export_data_path = "Export X-Fall Array.xlsx"

# Import Cross Fall (%) data excel file to Data frame (CROWN NAME, CHAINAGE (M.), CROSS FALL (%))
df_XFall_DATA = pd.read_excel(Import_data_path, "X-FALL DATA")

# Count Cross Fall (%) data
totalXFall = df_XFall_DATA["CHAINAGE (M.)"].count()

# Record to Data Frame
ColumnNames_XFallAR = ["CROWN NAME", "LOOP NO.", "CH.START (m.)", "CH.END (m.)", "X-FALL.START (%)", "X-FALL.END (%)", "TYPE", "REMARK"]
df_XFall_ARRAY = pd.DataFrame(columns= ColumnNames_XFallAR)

for i in range(totalXFall):
    if i <= totalXFall - 2:
        CrownName = df_XFall_DATA["CROWN NAME"][i]
        CH1 = df_XFall_DATA["CHAINAGE (M.)"][i]
        CH2 = df_XFall_DATA["CHAINAGE (M.)"][i + 1]
        XFall1 = df_XFall_DATA["CROSS FALL (%)"][i]
        XFall2 = df_XFall_DATA["CROSS FALL (%)"][i + 1]
        LoopNo = totalXFall - i

        # X-Fall Type
        if XFall1 == XFall2:
            XFallType = "N"
        elif XFall1 != XFall2:
            XFallType = "V"
        else:
            XFallType = False

        # Add Cross Fall (%) data to data frame df_XFall_ARRAY
        df_XFall_AR = pd.DataFrame([[CrownName, LoopNo, CH1, CH2, XFall1, XFall2, XFallType, ""]], columns=ColumnNames_XFallAR)
        df_XFall_ARRAY = df_XFall_ARRAY._append(df_XFall_AR, ignore_index=True)

    else:
        CrownName = df_XFall_DATA["CROWN NAME"][i]
        CH1 = df_XFall_DATA["CHAINAGE (M.)"][i]
        CH2 = df_XFall_DATA["CHAINAGE (M.)"][i] + 0.001
        XFall1 = df_XFall_DATA["CROSS FALL (%)"][i]
        XFall2 = df_XFall_DATA["CROSS FALL (%)"][i]
        LoopNo = totalXFall - i

        # X-Fall Type
        if XFall1 == XFall2:
            XFallType = "N"
        elif XFall1 != XFall2:
            XFallType = "V"
        else:
            XFallType = False
        
        # Add Cross Fall (%) data to data frame df_XFall_ARRAY
        df_XFall_AR = pd.DataFrame([[CrownName, LoopNo, CH1, CH2, XFall1, XFall2, XFallType, ""]], columns=ColumnNames_XFallAR)
        df_XFall_ARRAY = df_XFall_ARRAY._append(df_XFall_AR, ignore_index=True)

# Export Cross Fall (%) Array Result
df_XFall_ARRAY["REMARK"][[0, 1]] = ["V = Vary", "N = Normal"]

with pd.ExcelWriter(Export_data_path) as writer:
    df_XFall_ARRAY.to_excel(writer, sheet_name="XFALL-ARRAY", index = False)

t = time.time() # End time
print("Cross Fall (%) Array was created completely!, {:.3f}sec.".format(t-t0))
#----------------- End Main Cross Fall (%) Array ----------------#
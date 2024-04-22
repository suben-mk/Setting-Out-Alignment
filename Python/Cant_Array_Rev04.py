# Topic; Cant Array
# Created By; Suben Mukem (SBM) as Survey Engineer.
# Updated; 29/01/2024

# Import Module
import numpy as np
import pandas as pd
import time

t0 = time.time() # Start time

#------------------- Main Cant Array ------------------#
# Path files
Import_data_path = "Import Setting-Out Alignment Data.xlsx"
Export_data_path = "Export Cant Array.xlsx"

# Import CANT DATA excel file to Data frame (HIP NO. / NAME, MAIN POINT, CHAINAGE (M.), CANT (MM.))
df_CANT_DATA = pd.read_excel(Import_data_path, "CANT DATA")

# Count CANT DATA
totalCANT = df_CANT_DATA["MAIN POINT"].count()

# Record to Data Frame
ColumnNames_CNTAR = ["HIP NO.", "MAIN POINT", "LOOP NO.", "CH.START (m.)", "CH.END (m.)", "CANT START (MM.)", "CANT END (MM.)", "TYPE", "REMARK"]
df_CANT_ARRAY = pd.DataFrame(columns= ColumnNames_CNTAR)

for i in range(totalCANT):
    if i <= totalCANT - 2:
        HipNo = df_CANT_DATA["HIP NO. / NAME"][i]
        Pnt = df_CANT_DATA["MAIN POINT"][i]
        CH1 = df_CANT_DATA["CHAINAGE (M.)"][i]
        CH2 = df_CANT_DATA["CHAINAGE (M.)"][i + 1]
        Cant1 = df_CANT_DATA["CANT (MM.)"][i]
        Cant2 = df_CANT_DATA["CANT (MM.)"][i + 1]
        LoopNo = totalCANT - i

        # Cant Type
        if Cant1 == Cant2:
            CantType = "N"
        elif Cant1 != Cant2:
            CantType = "V"
        else:
            CantType = False

        # Add Cant data to data frame df_CANT_ARRAY
        df_CANT_AR = pd.DataFrame([[HipNo, Pnt, LoopNo, CH1, CH2, Cant1, Cant2, CantType, ""]], columns=ColumnNames_CNTAR)
        df_CANT_ARRAY = df_CANT_ARRAY._append(df_CANT_AR, ignore_index=True)

    else:
        HipNo = df_CANT_DATA["HIP NO. / NAME"][i]
        Pnt = df_CANT_DATA["MAIN POINT"][i]
        CH1 = df_CANT_DATA["CHAINAGE (M.)"][i]
        CH2 = df_CANT_DATA["CHAINAGE (M.)"][i] + 0.001
        Cant1 = df_CANT_DATA["CANT (MM.)"][i]
        Cant2 = df_CANT_DATA["CANT (MM.)"][i]
        LoopNo = totalCANT - i

        # Cant Type
        if Cant1 == Cant2:
            CantType = "N"
        elif Cant1 != Cant2:
            CantType = "V"
        else:
            CantType = False
        
        # Add Cant data to data frame df_CANT_ARRAY
        df_CANT_AR = pd.DataFrame([[HipNo, Pnt, LoopNo, CH1, CH2, Cant1, Cant2, CantType, ""]], columns=ColumnNames_CNTAR)
        df_CANT_ARRAY = df_CANT_ARRAY._append(df_CANT_AR, ignore_index=True)

# Export Cant Array Result
df_CANT_ARRAY["REMARK"][[0, 1]] = ["V = Vary", "N = Normal"]

with pd.ExcelWriter(Export_data_path) as writer:
    df_CANT_ARRAY.to_excel(writer, sheet_name="CANT-ARRAY", index = False)

t = time.time() # End time
print("Cant Array was created completely!, {:.3f}sec.".format(t-t0))
#----------------- End Main Cant Array ----------------#
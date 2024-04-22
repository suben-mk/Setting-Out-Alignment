# Topic; Tunnel Offset Array
# Created By; Suben Mukem (SBM) as Survey Engineer.
# Updated; 29/01/2024

# Import Module
import numpy as np
import pandas as pd
import time

t0 = time.time() # Start time

#------------------- Main Tunnel Offset Array ------------------#
# Path files
Import_data_path = "Import Setting-Out Alignment Data.xlsx"
Export_data_path = "Export Tunnel-OS Array.xlsx"

# Import Tunnel Offset data excel file to Data frame (HIP NO. / NAME, MAIN POINT, CHAINAGE (M.), HOR.TUNNEL OFFSET (M.), VER.TUNNEL OFFSET (M.))
df_TUOS_DATA = pd.read_excel(Import_data_path, "TUNNEL OFFSET DATA")

# Count tunnel offset data
totalTUOS = df_TUOS_DATA["MAIN POINT"].count()

# Record to Data Frame
ColumnNames_TUOSAR = ["HIP NO.", "MAIN POINT", "LOOP NO.", "CH.START (m.)", "CH.END (m.)", "HOR.OS START (M.)", "HOR.OS END (M.)", "VER.OS START (M.)", "VER.OS END (M.)", "HOR. TYPE", "VER. TYPE", "REMARK"]
df_TUOS_ARRAY = pd.DataFrame(columns= ColumnNames_TUOSAR)

for i in range(totalTUOS):
    if i <= totalTUOS - 2:
        HipNo = df_TUOS_DATA["HIP NO. / NAME"][i]
        Pnt = df_TUOS_DATA["MAIN POINT"][i]
        CH1 = df_TUOS_DATA["CHAINAGE (M.)"][i]
        CH2 = df_TUOS_DATA["CHAINAGE (M.)"][i + 1]
        HorOS1 = df_TUOS_DATA["HOR.TUNNEL OFFSET (M.)"][i]
        HorOS2 = df_TUOS_DATA["HOR.TUNNEL OFFSET (M.)"][i + 1]
        VerOS1 = df_TUOS_DATA["VER.TUNNEL OFFSET (M.)"][i]
        VerOS2 = df_TUOS_DATA["VER.TUNNEL OFFSET (M.)"][i + 1]
        LoopNo = totalTUOS - i

        # Tunnel Offset Type
        if HorOS1 == HorOS2:
            HorType = "N"
        elif HorOS1 != HorOS2:
            HorType = "V"
        else:
            HorType = False

        if VerOS1 == VerOS2:
            VerType = "N"
        elif VerOS1 != VerOS2:
            VerType = "V"
        else:
            VerType = False

        # Add tunnel offset data to data frame df_TUNOS_ARRAY
        df_TUOS_AR = pd.DataFrame([[HipNo, Pnt, LoopNo, CH1, CH2, HorOS1, HorOS2, VerOS1, VerOS2, HorType, VerType, ""]], columns=ColumnNames_TUOSAR)
        df_TUOS_ARRAY = df_TUOS_ARRAY._append(df_TUOS_AR, ignore_index=True)

    else:
        HipNo = df_TUOS_DATA["HIP NO. / NAME"][i]
        Pnt = df_TUOS_DATA["MAIN POINT"][i]
        CH1 = df_TUOS_DATA["CHAINAGE (M.)"][i]
        CH2 = df_TUOS_DATA["CHAINAGE (M.)"][i] + 0.001
        HorOS1 = df_TUOS_DATA["HOR.TUNNEL OFFSET (M.)"][i]
        HorOS2 = df_TUOS_DATA["HOR.TUNNEL OFFSET (M.)"][i]
        VerOS1 = df_TUOS_DATA["VER.TUNNEL OFFSET (M.)"][i]
        VerOS2 = df_TUOS_DATA["VER.TUNNEL OFFSET (M.)"][i]
        LoopNo = totalTUOS - i

        # Tunnel Offset Type
        if HorOS1 == HorOS2:
            HorType = "N"
        elif HorOS1 != HorOS2:
            HorType = "V"
        else:
            HorType = False

        if VerOS1 == VerOS2:
            VerType = "N"
        elif VerOS1 != VerOS2:
            VerType = "V"
        else:
            VerType = False
        
        # Add tunnel offset data to data frame df_TUNOS_ARRAY
        df_TUOS_AR = pd.DataFrame([[HipNo, Pnt, LoopNo, CH1, CH2, HorOS1, HorOS2, VerOS1, VerOS2, HorType, VerType, ""]], columns=ColumnNames_TUOSAR)
        df_TUOS_ARRAY = df_TUOS_ARRAY._append(df_TUOS_AR, ignore_index=True)

# Export Tunnel Offset Array Result
df_TUOS_ARRAY["REMARK"][[0, 1]] = ["V = Vary", "N = Normal"]

with pd.ExcelWriter(Export_data_path) as writer:
    df_TUOS_ARRAY.to_excel(writer, sheet_name="TUOS-ARRAY", index = False)

t = time.time() # End time
print("Tunnel Offset Array was created completely!, {:.3f}sec.".format(t-t0))
#----------------- End Main Tunnel Offset Array ----------------#
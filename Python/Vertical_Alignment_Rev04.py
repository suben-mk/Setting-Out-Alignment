# Topic; Setting Out Vertical Alignment
# Created By; Suben Mukem (SBM) as Survey Engineer.
# Updated; 29/01/2024

# Import Module
import numpy as np
import pandas as pd
import time

t0 = time.time() # Start time

#------------------- Main Vertical Alignment Computation ------------------#
# Path files
Import_data_path = "Import Setting-Out Alignment Data.xlsx"
Export_data_path = "Export Ver-Alignment.xlsx"

# Import VIP DATA excel file to Data frame (VIP NO. / NAME, CHAINAGE, ELEVATION, LVC, LVC 1, LVC 2)
df_VIP_DATA = pd.read_excel(Import_data_path, "VIP DATA")

# Count HIP DATA
totalVIP = df_VIP_DATA["VIP NO. / NAME"].count() - 1

# When compute alignment finish then record to Data Frame
ColumnNames_VSO = ["VIP NO.", "POINT", "CHAINAGE (m.)", "ELEVATION (m.)", "GRADIENT 1 (%)", "GRADIENT 2 (%)", "LVC (m.)", "LVC 1 (m.)", "LVC 2 (m.)", "REMARKS"]
df_VER_SETTING_OUT = pd.DataFrame(columns= ColumnNames_VSO)

ColumnNames_VAR = ["VIP NO.", "MAIN POINT", "LOOP NO.", "CH.START (m.)", "CH.END (m.)", "ELEVATION (m.)", "GRADIENT 1 (%)", "GRADIENT 2 (%)", "LVC (m.)", "LVC 1 (m.)", "LVC 2 (m.)", "CURVE TYPE", "REMARK"]
df_VER_ARRAY = pd.DataFrame(columns= ColumnNames_VAR)

# Beginning of point (BOP)
VipNo = df_VIP_DATA["VIP NO. / NAME"][0] # VIP name of BOP
CBP = df_VIP_DATA["CHAINAGE (M.)"][0] # Chainage of BOP
ELBP = df_VIP_DATA["ELEVATION (M.)"][0] # Elevation of of BOP
CPINext = df_VIP_DATA["CHAINAGE (M.)"][1] # Next VIP curve
ELPINext = df_VIP_DATA["ELEVATION (M.)"][1] # Next VIP curve
# Gradient (%)
GBP2 = ((ELPINext - ELBP) / (CPINext - CBP)) * 100
GBP1 = GBP2
# Add BOP data to data frame df_VER_SETTING_OUT
df_VER_SO = pd.DataFrame([[VipNo, "BOP", CBP, ELBP, GBP1, GBP2, "", "", "", ""]], columns=ColumnNames_VSO)
df_VER_SETTING_OUT = df_VER_SETTING_OUT._append(df_VER_SO, ignore_index=True)
# Add BOP data to data frame df_VER_ARRAY
df_VER_AR = pd.DataFrame([[VipNo, "BOP", "", CBP, "", ELBP, GBP1, GBP2, "", "", "", "T", ""]], columns=ColumnNames_VAR)
df_VER_ARRAY = df_VER_ARRAY._append(df_VER_AR, ignore_index=True)

# Palabolic curve
for i in range(1, totalVIP):
    VipNo = df_VIP_DATA["VIP NO. / NAME"][i] # VIP name
    CPI = df_VIP_DATA["CHAINAGE (M.)"][i] # Chainage of PVI
    ELPI = df_VIP_DATA["ELEVATION (M.)"][i] # Elevation of PVI
    LVC = df_VIP_DATA["LVC (M.)"][i] # Length of vertical curve (Symmetry)
    LVC1 = df_VIP_DATA["LVC 1 (M.)"][i] # Length 1 of vertical curve (Unsymmetry)
    LVC2 = df_VIP_DATA["LVC 2 (M.)"][i] # Length 2 of vertical curve (Unsymmetry)
    CPIBack = df_VIP_DATA["CHAINAGE (M.)"][i - 1] # Chainage of PVI back
    ELPIBack = df_VIP_DATA["ELEVATION (M.)"][i - 1] # Elevation of PVI back
    CPINext = df_VIP_DATA["CHAINAGE (M.)"][i + 1] # Chainage of PVI next
    ELPINext = df_VIP_DATA["ELEVATION (M.)"][i + 1] # Elevation of PVI next

    # Symmetric Curve
    if LVC != 0 and LVC1 == 0 and LVC2 == 0:
        G1 = ((ELPI - ELPIBack) / (CPI - CPIBack)) * 100 # Gradient 1 (%)
        G2 = ((ELPINext - ELPI) / (CPINext - CPI)) * 100 # Gradient 2 (%)
        CPVC = CPI - LVC / 2 # Chainage of PVC
        ELPVC = ELPI - (G1 / 100) * (LVC / 2) # Elevation of PVC
        CPVT = CPVC + LVC # Chainage of PVT
        ELPVT = ELPI + (G2 / 100) * (LVC / 2) # Elevation of PVT
        CurveType = "S"
        
    # Symmetric Curve
    elif LVC == 0 and LVC1 != 0 and LVC2 != 0:
        G1 = ((ELPI - ELPIBack) / (CPI - CPIBack)) * 100 # Gradient 1 (%)
        G2 = ((ELPINext - ELPI) / (CPINext - CPI)) * 100 # Gradient 2 (%)
        CPVC = CPI - LVC1 # Chainage of PVC
        ELPVC = ELPI - (G1 / 100) * LVC1 # Elevation of PVC
        CPVT = CPVC + LVC1 + LVC2 # Chainage of PVT
        ELPVT = ELPI + (G2 / 100) * LVC2 # Elevation of PVT
        CurveType = "U"
    else:
        False

    # Add PVC, PVI and PVT data frame df_VER_SETTING_OUT
    df_VER_SO = pd.DataFrame([["", "PVC", CPVC, ELPVC, G1, G2, "", "", "", ""],
                              [VipNo, "PVI", CPI, ELPI, "", "", LVC, LVC1, LVC2, ""],
                              ["", "PVT", CPVT, ELPVT, G2, G2, "", "", "", ""]],
                             columns=ColumnNames_VSO)
    df_VER_SETTING_OUT = df_VER_SETTING_OUT._append(df_VER_SO, ignore_index=True)
    # Add PVC and PVT data frame df_VER_ARRAY
    df_VER_AR = pd.DataFrame([[VipNo, "PVC", "", CPVC, "", ELPVC, G1, G2, LVC, LVC1, LVC2, CurveType, ""],
                              ["", "PVT", "", CPVT, "", ELPVT, G2, G2, "", "", "", "T", ""]],
                             columns=ColumnNames_VAR)
    df_VER_ARRAY = df_VER_ARRAY._append(df_VER_AR, ignore_index=True)

# Ending of point (EOP)
VipNo = df_VIP_DATA["VIP NO. / NAME"][totalVIP] # VIP name
CEP = df_VIP_DATA["CHAINAGE (M.)"][totalVIP] # Chainage of end point
ELEP = df_VIP_DATA["ELEVATION (M.)"][totalVIP] # Elevation of end point
CPIBack = df_VIP_DATA["CHAINAGE (M.)"][totalVIP - 1] # Chainage of PVI back
ELPIBack = df_VIP_DATA["ELEVATION (M.)"][totalVIP - 1] # Elevation of PVI back
# Gradient 1&2 (%)
GEP1 = ((ELEP - ELPIBack) / (CEP - CPIBack)) * 100
GEP2 = GEP1
# Add EOP data to data frame df_VER_SETTING_OUT
df_VER_SO = pd.DataFrame([[VipNo, "EOP", CEP, ELEP, GEP1, GEP2, "", "", "", ""]], columns=ColumnNames_VSO)
df_VER_SETTING_OUT = df_VER_SETTING_OUT._append(df_VER_SO, ignore_index=True)
# Add EOP data to data frame df_VER_ARRAY
df_VER_AR = pd.DataFrame([[VipNo, "EOP", "", CEP, "", ELEP, GEP1, GEP2, "", "", "", "T", ""]], columns=ColumnNames_VAR)
df_VER_ARRAY = df_VER_ARRAY._append(df_VER_AR, ignore_index=True)

# Export Vertical Alignment Result
df_VER_ARRAY["REMARK"][[0, 1, 2]] = ["T = Tangent", "S = Symmetric Curve", "U = Unsymmetric Curve"]
totalVER_ARRAY = df_VER_ARRAY["MAIN POINT"].count()
for i in range(totalVER_ARRAY):
    if i <= totalVER_ARRAY - 2:
        df_VER_ARRAY["LOOP NO."][i] = totalVER_ARRAY - i
        df_VER_ARRAY["CH.END (m.)"][i] = df_VER_ARRAY["CH.START (m.)"][i + 1]
    else:
        df_VER_ARRAY["LOOP NO."][i] = totalVER_ARRAY - i
        df_VER_ARRAY["CH.END (m.)"][i] = df_VER_ARRAY["CH.START (m.)"][i] + 0.001

with pd.ExcelWriter(Export_data_path) as writer:
    df_VER_SETTING_OUT.to_excel(writer, sheet_name="VER-SETTING OUT", index = False)
    df_VER_ARRAY.to_excel(writer, sheet_name="VER-ARRAY", index = False)

t = time.time() # End time
print("Vertical Alignment was computed completely!, {:.3f}sec.".format(t-t0))
#----------------- End Main Vertical Alignment Computation ----------------#
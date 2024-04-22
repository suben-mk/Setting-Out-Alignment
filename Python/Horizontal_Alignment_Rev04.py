# Topic; Setting Out Horizontal Alignment
# Created By; Suben Mukem (SBM) as Survey Engineer.
# Updated; 29/01/2024

# Import Module
import math # for General Survey Function
import numpy as np
import pandas as pd
import time

t0 = time.time() # Start time

#---------------------------- All Function ----------------------------#
## General Survey ##
# Convert Degrees to Radians
def DegtoRad(d):
    ang = d * math.pi / 180.0
    return ang

# Convert Radians to Degrees
def RadtoDeg(d):
    ang = d * 180 / math.pi
    return ang

# Compute Distance and Azimuth from 2 Points
def DirecAziDist(EStart, NStart, EEnd, NEnd):
    dE = EEnd - EStart
    dN = NEnd - NStart
    Dist = math.sqrt(dE**2 + dN**2)

    if dN != 0:
        ang = math.atan2(dE, dN)
    else:
        Azi = False

    if ang >= 0:
        Azi = RadtoDeg(ang)
    else:
        Azi = RadtoDeg(ang) + 360
    return Dist, Azi

# Compute Grid Coordinate (E, N) by Local Coordinate (Y, X), Coordinate of Center and Azimuth.
def CoorYXtoNE(ECL, NCL, AZCL, Y, X):
    Ei = ECL + Y * math.sin(DegtoRad(AZCL)) + X * math.sin(DegtoRad(90 + AZCL))
    Ni = NCL + Y * math.cos(DegtoRad(AZCL)) + X * math.cos(DegtoRad(90 + AZCL))
    return Ei, Ni

# Convert Degrees to dd-mm-ss (String)
def DegtoDMSStr1(deg):
    d = abs(deg)
    mm, ss = divmod(d * 3600, 60)
    dd, mm = divmod(mm, 60)
    return '{:.0f}-{:.0f}-{:.2f}'.format(dd, mm, ss)

## Spiral Curve (Clothoid) ##
def Spiral(Ls, Rc, d):
    # 1.Spiral Angle and Circular Angle (rad.)
    Qs = Ls / (2 * Rc)

    # 2.Offset Xs, Ys
    C1 = 1 / 3 ; C2 = -1 / 10 ; C3 = -1 / 42 ; C4 = 1 / 216 ; C5 = 1 / 1320 ; C6 = -1 / 9360 ; C7 = -1 / 75600 ; C8 = 1 / 685440
    Xs = Ls * (1 + (C2 * Qs ** 2) + (C4 * Qs ** 4) + (C6 * Qs ** 6) + (C8 * Qs ** 8))
    Ys = Ls * ((C1 * Qs) + (C3 * Qs ** 3) + (C5 * Qs ** 5) + (C7 * Qs ** 7))

    # 3.Offset from PCO Tangent to New Curve (m.)
    P = Ys - Rc * (1 - np.cos(Qs))

    # 4.Distance from PCO Tangent to New Curve (m.)
    K = Xs - Rc * np.sin(Qs)

    # 5.Tangent Line from TS to PI or ST to PI (m.)
    Ts = (Rc + P) * np.tan(DegtoRad(d / 2)) + K
    return Qs, K, P, Xs, Ys, Ts

def SpiralIN(Ls1, Ls2, Rc, d):
    if Ls1 == Ls2:
        Qs1, K1, P1, Xs1, Ys1, Ts1 = Spiral(Ls1, Rc, d)
    else:
        Qs1, K1, P1, Xs1, Ys1, Ts1 = Spiral(Ls1, Rc, d)
        Qs2, K2, P2, Xs2, Ys2, Ts2 = Spiral(Ls2, Rc, d)
        Ts1 = K1 + Rc * np.tan(DegtoRad(d / 2)) + P2 / np.sin(DegtoRad(d)) - P1 / np.tan(DegtoRad(d))
    return Xs1, Ys1, Ts1

def SpiralOUT(Ls1, Ls2, Rc, d):
    if Ls1 == Ls2:
        Qs2, K2, P2, Xs2, Ys2, Ts2 = Spiral(Ls2, Rc, d)
    else:
        Qs1, K1, P1, Xs1, Ys1, Ts1 = Spiral(Ls1, Rc, d)
        Qs2, K2, P2, Xs2, Ys2, Ts2 = Spiral(Ls2, Rc, d)
        Ts2 = K2 + Rc * np.tan(DegtoRad(d / 2)) + P1 / np.sin(DegtoRad(d)) - P2 / np.tan(DegtoRad(d))
    return Xs2, Ys2, Ts2
#------------------------ End All Function ------------------------#

#------------------- Main Horizontal Alignment Computation ------------------#
# Path files
Import_data_path = "Import Setting-Out Alignment Data.xlsx"
Export_data_path = "Export Hor-Alignment.xlsx"

# Input beginning point as list[Chainage, Easting, Northing]
BEGIN_POINT = [7202.834, 662670.304, 1521355.848]

# Import HIP DATA excel file to Data frame (HIP NO. / NAME, EASTING, NORTHING, RADIUS, Ls1, Ls2)
df_HIP_DATA = pd.read_excel(Import_data_path, "HIP DATA")

# Count HIP DATA
totalHIP = df_HIP_DATA["HIP NO. / NAME"].count() - 1

# When compute alignment finish then record to Data Frame
ColumnNames_HSO = ["HIP NO.", "POINT", "CHAINAGE (m.)", "EASTING (m.)", "NORTHING (m.)", "AZIMUTH (dd-mm-ss)", "AZIMUTH (deg.)", "AZIMUTH TANGENT (dd-mm-ss)", "AZIMUTH TANGENT (deg.)", "DEFLECTION ANGLE (dd-mm-ss)", "DEFLECTION ANGLE (deg.)", "LT/RT", "RADIUS (m.)", "Ls IN (m.)", "Lc (m.)", "Ls OUT (m.)", "REMARKS"]
df_HOR_SETTING_OUT = pd.DataFrame(columns= ColumnNames_HSO)

ColumnNames_HAR = ["HIP NO.", "MAIN POINT", "LOOP NO.", "CH.START (m.)", "CH.END (m.)", "E.START (m.)", "N.START (m.)", "AZIMUTH (deg.)", "RADIUS (m.)", "CURVE TYPE", "REMARK"]
df_HOR_ARRAY = pd.DataFrame(columns= ColumnNames_HAR)

# Beginning of point (BOP)
HipNo = df_HIP_DATA["HIP NO. / NAME"][0] # HIP No. of BOP
CBP = BEGIN_POINT[0] # Chainage of BOP
EBP = df_HIP_DATA["EASTING (M.)"][0] # Easting of BOP
NBP = df_HIP_DATA["NORTHING (M.)"][0] # Northing of BOP
EPINext = df_HIP_DATA["EASTING (M.)"][1] # Easting of Next HIP
NPINext = df_HIP_DATA["NORTHING (M.)"][1] # Northing of Next HIP
DistT2, AzT2 = DirecAziDist(EBP, NBP, EPINext, NPINext) # Distance and Azimuth of tangent 2
AzT2DMS = DegtoDMSStr1(AzT2) # Azimuth of BOP (dd-mm-ss)
# Add BOP data to data frame df_HOR_SETTING_OUT
df_HOR_SO = pd.DataFrame([[HipNo, "BOP", CBP, EBP, NBP, AzT2DMS, AzT2, "", "", "", "", "", "", "", "", "", ""]], columns=ColumnNames_HSO)
df_HOR_SETTING_OUT = df_HOR_SETTING_OUT._append(df_HOR_SO, ignore_index=True)
BEGIN_POINT = [CBP, EBP, NBP] # Change beginning point of next curve
# Add BOP data to data frame df_HOR_ARRAY
df_HOR_AR = pd.DataFrame([[HipNo, "BOP", "", CBP, "", EBP, NBP, AzT2, 0, "T", ""]], columns=ColumnNames_HAR)
df_HOR_ARRAY = df_HOR_ARRAY._append(df_HOR_AR, ignore_index=True)

# Case of Curve type e.g. PI no curve, Circular curve and Spiral curve 
for i in range(1, totalHIP):
    HipNo = df_HIP_DATA["HIP NO. / NAME"][i]
    EPI = df_HIP_DATA["EASTING (M.)"][i]
    NPI = df_HIP_DATA["NORTHING (M.)"][i]
    EPIBack = df_HIP_DATA["EASTING (M.)"][i - 1]
    NPIBack = df_HIP_DATA["NORTHING (M.)"][i - 1]
    EPINext = df_HIP_DATA["EASTING (M.)"][i + 1] # Easting of Next HIP
    NPINext = df_HIP_DATA["NORTHING (M.)"][i + 1] # Northing of Next HIP
    Radius = df_HIP_DATA["RADIUS (M.)"][i]
    Ls1 = df_HIP_DATA["Ls1 (M.)"][i]
    Ls2 = df_HIP_DATA["Ls2 (M.)"][i]
    DistT1, AzT1 = DirecAziDist(EPIBack, NPIBack, EPI, NPI)
    AzT1DMS = DegtoDMSStr1(AzT1)
    DistT2, AzT2 = DirecAziDist(EPI, NPI, EPINext, NPINext)
    AzT2DMS = DegtoDMSStr1(AzT2)

    # Left turn (LT) or Rigth turn (RT) and Deflection angle
    Delta = AzT2 - AzT1
    if np.abs(Delta) > 180:
        DefAngle = Delta - np.sign(Delta) * 360
        DefAngleDMS = DegtoDMSStr1(np.abs(DefAngle))
    else:
        DefAngle = Delta
        DefAngleDMS = DegtoDMSStr1(np.abs(DefAngle))
    
    if DefAngle < 0:
        TurnLR = "LT"
    else:
        TurnLR = "RT"

    # Case PI no Curve
    if Radius == 0:
        # PI point
        CBP, EBP, NBP = BEGIN_POINT
        DistBPtoPI, AzBPtoPI = DirecAziDist(EBP, NBP, EPI, NPI)
        CPI = CBP + DistBPtoPI
        # Add PI data to data frame df_HOR_SETTING_OUT
        df_HOR_SO = pd.DataFrame([[HipNo, "PI", CPI, EPI, NPI, AzT2DMS, AzT2, AzT1DMS, AzT1, "{} {}".format(DefAngleDMS, TurnLR), np.abs(DefAngle), TurnLR, Radius, Ls1, 0, Ls2, ""]], columns=ColumnNames_HSO)
        df_HOR_SETTING_OUT = df_HOR_SETTING_OUT._append(df_HOR_SO, ignore_index=True)
        BEGIN_POINT = [CPI, EPI, NPI] # Change beginning point of next curve
        # Add PI data to data frame df_HOR_ARRAY
        df_HOR_AR = pd.DataFrame([[HipNo, "PI", "", CPI, "", EPI, NPI, AzT2, 0, "T", ""]], columns=ColumnNames_HAR)
        df_HOR_ARRAY = df_HOR_ARRAY._append(df_HOR_AR, ignore_index=True)

    # Case Circular Curve
    elif Radius > 0 and Ls1 == 0 and Ls2 == 0:
        # Circular curve parameter
        Lc = Radius * DegtoRad(np.abs(DefAngle))
        Tc = Radius * np.tan(DegtoRad(np.abs(DefAngle) / 2))
        # PC Point
        EPC = EPI - Tc * np.sin(DegtoRad(AzT1))
        NPC = NPI - Tc * np.cos(DegtoRad(AzT1))
        AzPC = AzT1
        AzPCDMS = DegtoDMSStr1(AzPC)
        # PT Point
        EPT = EPI + Tc * np.sin(DegtoRad(AzT2))
        NPT = NPI + Tc * np.cos(DegtoRad(AzT2))
        AzPT = AzT2
        AzPTDMS = DegtoDMSStr1(AzPT)
        # Chainage of PC, PI and PT
        CBP, EBP, NBP = BEGIN_POINT
        DistBPtoPI, AzBPtoPI = DirecAziDist(EBP, NBP, EPI, NPI)
        CPI = CBP + DistBPtoPI
        CPC = CPI - Tc
        CPT = CPC + Lc
        # Add PC, PI and PT data to data frame df_HOR_SETTING_OUT
        df_HOR_SO = pd.DataFrame([["", "PC", CPC, EPC, NPC, AzPCDMS, AzPC, "", "", "", "", "", "", "", "", "", ""],
                                  [HipNo, "PI", CPI, EPI, NPI, "", "", AzT1DMS, AzT1, "{} {}".format(DefAngleDMS, TurnLR), np.abs(DefAngle), TurnLR, Radius, Ls1, Lc, Ls2, ""],
                                  ["", "PT", CPT, EPT, NPT, AzPTDMS, AzPT, "", "", "", "", "", "", "", "", "", ""]],
                                  columns=ColumnNames_HSO)
        df_HOR_SETTING_OUT = df_HOR_SETTING_OUT._append(df_HOR_SO, ignore_index=True)
        BEGIN_POINT = [CPT, EPT, NPT] # Change beginning point of next curve
        # Add PC and PT data to data frame df_HOR_ARRAY
        df_HOR_AR = pd.DataFrame([[HipNo, "PC", "", CPC, "", EPC, NPC, AzPC, Radius * np.sign(DefAngle), "C", ""],
                                  ["", "PT", "", CPT, "", EPT, NPT, AzPT, 0, "T", ""]],
                                 columns=ColumnNames_HAR)
        df_HOR_ARRAY = df_HOR_ARRAY._append(df_HOR_AR, ignore_index=True)
    
    # Case Spiral Curve
    elif Radius > 0 and Ls1 > 0 and Ls2 > 0:
        # Spiral In cumputation
        # Parameter
        Xs1, Ys1, Ts1 = SpiralIN(Ls1, Ls2, Radius, np.abs(DefAngle))
        Qs1, K, P, Xs, Ys, Ts  = Spiral(Ls1, Radius, np.abs(DefAngle))
        # TS Point
        ETS = EPI - Ts1 * np.sin(DegtoRad(AzT1))
        NTS = NPI - Ts1 * np.cos(DegtoRad(AzT1))
        AzTS = AzT1
        AzTSDMS = DegtoDMSStr1(AzTS)
        # SC point
        ESC, NSC = CoorYXtoNE(ETS, NTS, AzTS, Xs1, Ys1 * np.sign(DefAngle))
        AzSC = AzTS + RadtoDeg(Qs1) * np.sign(DefAngle)
        AzSCDMS = DegtoDMSStr1(AzSC)

        # Spiral Out cumputation
        # Parameter
        Xs2, Ys2, Ts2 = SpiralOUT(Ls1, Ls2, Radius, np.abs(DefAngle))
        Qs2, K, P, Xs, Ys, Ts  = Spiral(Ls2, Radius, np.abs(DefAngle))
        # ST point
        EST = EPI + Ts2 * np.sin(DegtoRad(AzT2))
        NST = NPI + Ts2 * np.cos(DegtoRad(AzT2))
        AzST = AzT2
        AzSTDMS = DegtoDMSStr1(AzST)
        # CS point
        ECS, NCS = CoorYXtoNE(EST, NST, AzST, -Xs2, Ys2 * np.sign(DefAngle))
        AzCS = AzST - RadtoDeg(Qs2) * np.sign(DefAngle)
        AzCSDMS = DegtoDMSStr1(AzCS)

        # Circular Parameter
        Qc = DegtoRad(np.abs(DefAngle)) - (Qs1 + Qs2)
        Lc = Radius * Qc
        # Chainage of TS, SC, PI, CS and ST
        CBP, EBP, NBP = BEGIN_POINT
        DistBPtoPI, AzBPtoPI = DirecAziDist(EBP, NBP, EPI, NPI)
        CPI = CBP + DistBPtoPI
        CTS = CPI - Ts1
        CSC = CTS + Ls1
        CCS = CSC + Lc
        CST = CCS + Ls2
        # Add TS, SC, PI, CS and ST data to data frame df_HOR_SETTING_OUT
        df_HOR_SO = pd.DataFrame([["", "TS", CTS, ETS, NTS, AzTSDMS, AzTS, "", "", "", "", "", "", "", "", "", ""],
                                  ["", "SC", CSC, ESC, NSC, AzSCDMS, AzSC, "", "", "", "", "", "", "", "", "", ""],
                                  [HipNo, "PI", CPI, EPI, NPI, "", "", AzT1DMS, AzT1, "{} {}".format(DefAngleDMS, TurnLR), np.abs(DefAngle), TurnLR, Radius, Ls1, Lc, Ls2, ""],
                                  ["", "CS", CCS, ECS, NCS, AzCSDMS, AzCS, "", "", "", "", "", "", "", "", "", ""],
                                  ["", "ST", CST, EST, NST, AzSTDMS, AzST, "", "", "", "", "", "", "", "", "", ""]],
                                  columns=ColumnNames_HSO)
        df_HOR_SETTING_OUT = df_HOR_SETTING_OUT._append(df_HOR_SO, ignore_index=True)
        BEGIN_POINT = [CST, EST, NST] # Change beginning point of next curve
        # Add TS, SC, CS and ST data to data frame df_HOR_ARRAY
        df_HOR_AR = pd.DataFrame([[HipNo, "TS", "", CTS, "", ETS, NTS, AzTS, Radius * np.sign(DefAngle), "SPIN", ""],
                                  ["", "SC", "", CSC, "", ESC, NSC, AzSC, Radius * np.sign(DefAngle), "C", ""],
                                  ["", "CS", "", CCS, "", ECS, NCS, AzCS, Radius * np.sign(DefAngle), "SPOT", ""],
                                  ["", "ST", "", CST, "", EST, NST, AzST, 0, "T", ""]],
                                 columns=ColumnNames_HAR)
        df_HOR_ARRAY = df_HOR_ARRAY._append(df_HOR_AR, ignore_index=True)

    else:
        False
    
# Ending of point (EOP)
HipNo = df_HIP_DATA["HIP NO. / NAME"][totalHIP] # HIP No. of EOP
EEP = df_HIP_DATA["EASTING (M.)"][totalHIP]
NEP = df_HIP_DATA["NORTHING (M.)"][totalHIP]
EPIBack = df_HIP_DATA["EASTING (M.)"][totalHIP - 1] # Easting of Next HIP
NPIBack = df_HIP_DATA["NORTHING (M.)"][totalHIP - 1] # Northing of Next HIP
DistT1, AzT1 = DirecAziDist(EPIBack, NPIBack, EEP, NEP) # Distance and Azimuth of tangent 1
AzT1DMS = DegtoDMSStr1(AzT1) # Azimuth of EOP (dd-mm-ss)

# Chainage of TS, SC, PI, CS and ST
CBP, EBP, NBP = BEGIN_POINT
DistBPtoPI, AzBPtoPI = DirecAziDist(EBP, NBP, EEP, NEP)
CEP = CBP + DistBPtoPI
# Add EOP data to data frame df_HOR_SETTING_OUT
df_HOR_SO = pd.DataFrame([[HipNo, "EOP", CEP, EEP, NEP, AzT1DMS, AzT1, AzT1DMS, AzT1, "", "", "", "", "", "", "", ""]], columns=ColumnNames_HSO)
df_HOR_SETTING_OUT = df_HOR_SETTING_OUT._append(df_HOR_SO, ignore_index=True)
BEGIN_POINT = [CEP, EEP, NEP] # Change beginning point of next curve
# Add EOP data to data frame df_HOR_ARRAY
df_HOR_AR = pd.DataFrame([[HipNo, "EOP", "", CEP, "", EEP, NEP, AzT1, 0, "T", ""]], columns=ColumnNames_HAR)
df_HOR_ARRAY = df_HOR_ARRAY._append(df_HOR_AR, ignore_index=True)

# Export Horizontal Alignment Result
df_HOR_ARRAY["REMARK"][[0, 1, 2, 3]] = ["T = Straight Line", "SPIN = Spiral Curve In", "C = Circular Curve", "SPOT = Spiral Curve Out"] # Add Remark to df_HOR_ARRAY
totalHOR_ARRAY = df_HOR_ARRAY["MAIN POINT"].count()
for i in range(totalHOR_ARRAY):
    if i <= totalHOR_ARRAY - 2:
        df_HOR_ARRAY["LOOP NO."][i] = totalHOR_ARRAY - i
        df_HOR_ARRAY["CH.END (m.)"][i] = df_HOR_ARRAY["CH.START (m.)"][i + 1]
    else:
        df_HOR_ARRAY["LOOP NO."][i] = totalHOR_ARRAY - i
        df_HOR_ARRAY["CH.END (m.)"][i] = df_HOR_ARRAY["CH.START (m.)"][i] + 0.001

with pd.ExcelWriter(Export_data_path) as writer:
    df_HOR_SETTING_OUT.to_excel(writer, sheet_name="HOR-SETTING OUT", index = False)
    df_HOR_ARRAY.to_excel(writer, sheet_name="HOR-ARRAY", index = False)

t = time.time() # End time
print("Horizontal Alignment was computed completely!, {:.3f}sec.".format(t-t0))
#----------------- End Main Horizontal Alignment Computation ----------------#
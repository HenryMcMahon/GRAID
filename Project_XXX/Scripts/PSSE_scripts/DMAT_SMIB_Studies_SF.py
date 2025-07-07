import sys, os 
# import xlwings as xw
import pandas as pd
import traceback,time
import numpy as np
import openpyxl
import shutil
from math import sqrt, pi

# PSSE SETUP #################################################################
sys_path_PSSE = r"C:\Program Files (x86)\PTI\PSSE34\PSSPY37"
sys.path.append(sys_path_PSSE)
os_path_PSSE = r"C:\Program Files (x86)\PTI\PSSE34\PSSBIN"
os.environ['PATH'] += ';' + os_path_PSSE
import psse34
import psspy, dyntools
from psspy import _i,_f,_s
psspy.throwPsseException=True
tmpstdout = sys.stdout
import redirect
def hide_output(hide=True):
    
    #WRITE TO 'NOWHERE' FILE (USED FOR HIDEPROGRESS) ----
    class DummyFile(object):
        def write(self, x): pass
    if hide == True:
        redirect.psse2py()
        sys.stdout = DummyFile()
    else:
        sys.stdout=tmpstdout


# OUTPUT AND LOG FILE LOCATION ################################################
setupdir = 'E:\\Projects\\0124_MiddlemountBESS\\PSSE\\CaseCreates\\'
DMATdir = 'DMAT_SMIB_Tests\\SolarFarm\\' 
LogDir = setupdir + DMATdir + "results\\Logs\\"
OutDir = setupdir + DMATdir + "results\\OutFiles\\"
ResultsDir = setupdir + DMATdir + "results\\csvFiles\\"
BaseCasesDir = setupdir + "Base_Cases_SF\\"

if not os.path.exists(LogDir):
    os.makedirs(LogDir) #Create results Log dir if doesnt exists
if not os.path.exists(OutDir):
    os.makedirs(OutDir) #Create results Output dir if doesnt exists
if not os.path.exists(ResultsDir):
    os.makedirs(ResultsDir) #Create results Csv dir if doesnt exists

# READ FILES ##################################################################
excel_file = setupdir + DMATdir + 'SMIB_studies_full_SF.xlsx'
dyrFile = "MMHY_SF.dyr"

# ************************************************************************** #
# Input Plant Parameters (for scaling purposes)
# ***************************************************************************#
ExportMWrating = 26
SFMWrating = 30.0

# BESSMWrating = 34.2
MWrating = 26 #MW

Sbase=100 #GRID
Vbase=66 #GRID
Zbase=Vbase**2/Sbase
Pmax = MWrating

#define busses within PSSE
INF_BUS = 8888
DUMMYBUS = 600 
POC_BUS = 100
HV_BUS = 200
MV_BUS = 300
INV_SF = 400
MV_SF_BUS = 401
INV_BESS = 500
MV_BESS_BUS = 501

### DMAT PSSE RUNTIME Parameters #############################################
iterations = 600
sfactor = 0.2
tolerance = 0.0001
dT = 0.001 #ms
ffilter = 0.008

stable_time = 10 #s Time for starting faults/tests
run_faults_for = 25 #s Time

out_txt_steps = 1000 #ms
plot_steps = 10 #ms
SF_INV_VAR = 1 + 0 #(Add six for PLB file)
#BESS_INV_VAR = 1 + 0 #(Add six for PLB file)
PPC_VAR = 627 + 0 #(Add six for PLB file) #627 with BESS.dyr

# DMAT PSSE EQUIPMENT NAMES ###################################################

PPC_MODEL = r"""SMAHYCF23"""
SF_INV_model = r"""SMASC193"""
# BESS_INV_model = r"""SMAGF308"""

# ***************************************************************************#
#def add_channels(SF_INV_VAR,BESS_INV_VAR, PPC_VAR): 
def add_channels(SF_INV_VAR, PPC_VAR): 
    
    
    psspy.machine_array_channel([-1,2,INV_SF],r"""1""",r"""SF PELEC""") #1
    psspy.machine_array_channel([-1,3,INV_SF],r"""1""",r"""SF QELEC""")
    psspy.machine_array_channel([-1,9,INV_SF],r"""1""",r"""SF INV IdCmd""")
    psspy.machine_array_channel([-1,5,INV_SF],r"""1""",r"""SF INV Qcmd""")
    psspy.machine_array_channel([-1,8,INV_SF],r"""1""",r"""SF INV Pcmd""")    
    psspy.machine_array_channel([-1,12,INV_SF],r"""1""",r"""SF INV IqCmd""")
    psspy.machine_array_channel([-1,2,INV_BESS],r"""1""",r"""BESS PELEC""")
    psspy.machine_array_channel([-1,3,INV_BESS],r"""1""",r"""BESS QELEC""")
    psspy.machine_array_channel([-1,9,INV_BESS],r"""1""",r"""BESS INV Id 1""")
    psspy.machine_array_channel([-1,5,INV_BESS],r"""1""",r"""BESS Qcmd PPC""")
    psspy.machine_array_channel([-1,8,INV_BESS],r"""1""",r"""BESS Pcmd PPC""")    
    psspy.machine_array_channel([-1,12,INV_BESS],r"""1""",r"""BESS INV Iq 1""") #12
    
    psspy.voltage_and_angle_channel([-1,-1,-1,POC_BUS],[r"""POC V""",r"""POC A"""]) #13
    psspy.voltage_and_angle_channel([-1,-1,-1,INV_SF],[r"""SF INV V""",r"""SF INV A"""]) 
    psspy.voltage_and_angle_channel([-1,-1,-1,INV_BESS],[r"""BESS INV V""",r"""BESS INV A"""]) 
    psspy.voltage_and_angle_channel([-1,-1,-1,MV_SF_BUS],[r"""SF MV V""",r"""SF MV A"""]) 
    psspy.voltage_and_angle_channel([-1,-1,-1,MV_BESS_BUS],[r"""BESS MV V""",r"""BESS MV A"""]) 
    psspy.voltage_and_angle_channel([-1,-1,-1,INF_BUS],[r"""GRID V""",r"""GRID A"""])
    psspy.voltage_and_angle_channel([-1,-1,-1,DUMMYBUS],[r"""DUMMY V""",r"""DUMMY A"""]) #22
    
    psspy.bus_frequency_channel([-1,POC_BUS],r"""POC FREQ""") #23
    psspy.bus_frequency_channel([-1,INF_BUS],r"""GRID FREQ""") #24

    psspy.branch_p_and_q_channel([-1,-1,-1,POC_BUS,DUMMYBUS],r"""1""",[r"""POC P""",r"""POC Q"""]) #25
    psspy.branch_p_and_q_channel([-1,-1,-1,INV_BESS,MV_BESS_BUS],r"""1""",[r"""BESS INV P""",r"""BESS INV Q"""])
    psspy.branch_p_and_q_channel([-1,-1,-1,INV_SF,MV_SF_BUS],r"""1""",[r"""SF INV P""",r"""SF INV Q"""])
   
    
    # PPC VAR Channels
    
    psspy.var_channel([-1,PPC_VAR+13],r"""PPC_BESS_W_Spnt""") #31
    psspy.var_channel([-1,PPC_VAR+14],r"""PPC_Q_Spnt""")
    psspy.var_channel([-1,PPC_VAR+15],r"""PPC_PF_Spnt""")
    psspy.var_channel([-1,PPC_VAR+16],r"""PPC_V_Spnt""")
    psspy.var_channel([-1,PPC_VAR+17],r"""PPC_Hz_Spnt""")
    psspy.var_channel([-1,PPC_VAR+18],r"""PPC_P_Spnt""")
    psspy.var_channel([-1,PPC_VAR+19],r"""PPC_ExtPwrAtLimLo""")
    psspy.var_channel([-1,PPC_VAR+20],r"""PPC_ExtPwrAtLimHi""")
    psspy.var_channel([-1,PPC_VAR+28],r"""PPC_PvPwrAtAvail""")
    psspy.var_channel([-1,PPC_VAR+29],r"""PPC_BatPwrAtAvail""")
    psspy.var_channel([-1,PPC_VAR+43],r"""PPC_Ppv_com""")
    psspy.var_channel([-1,PPC_VAR+44],r"""PPC_Qpv_com""")
    psspy.var_channel([-1,PPC_VAR+45],r"""PPC_Pbess_com""")
    psspy.var_channel([-1,PPC_VAR+46],r"""PPC_Qbess_com""")
    psspy.var_channel([-1,PPC_VAR+55],r"""PPC_BatPwrAtAvail_2""")
    psspy.var_channel([-1,PPC_VAR+60],r"""PPC_BatWSptMax""")
    psspy.var_channel([-1,PPC_VAR+61],r"""PPC_BatWSptMin""")
    psspy.var_channel([-1,PPC_VAR+62],r"""PPC_Ppv_cmd_inv""")
    psspy.var_channel([-1,PPC_VAR+63],r"""PPC_Qpv_cmd_inv""")
    psspy.var_channel([-1,PPC_VAR+64],r"""PPC_Pbess_cmd_inv""")
    psspy.var_channel([-1,PPC_VAR+65],r"""PPC_Qbess_cmd_inv""")
    psspy.var_channel([-1,PPC_VAR+79],r"""PPC_PoiFrqComp""")
    psspy.var_channel([-1,PPC_VAR+80],r"""PPC_FrqPoi_processed""")
    psspy.var_channel([-1,PPC_VAR+81],r"""PPC_Pfrq_Poi_Filtered_processed""")
    psspy.var_channel([-1,PPC_VAR+86],r"""PPC_PwrAtSpnt_FCAS""")
    psspy.var_channel([-1,PPC_VAR+98],r"""PPC_FrqRespPwrAtMaxNom""")
    psspy.var_channel([-1,PPC_VAR+99],r"""PPC_FrqRespPwrAtMinNom""")
    psspy.var_channel([-1,PPC_VAR+100],r"""PPC_FrqRespPwrAtSpnt""")
    psspy.var_channel([-1,PPC_VAR+200],r"""PPC_ExtPwrRtSpntOut""")
    psspy.var_channel([-1,PPC_VAR+201],r"""PPC_PwrRtSpnt""")
    psspy.var_channel([-1,PPC_VAR+266],r"""PPC_PvPwrAtSpnt""")
    psspy.var_channel([-1,PPC_VAR+268],r"""PPC_BatPwrAtSpnt""")
    psspy.var_channel([-1,PPC_VAR+75],r"""FRT Active signal""")
    psspy.var_channel([-1,PPC_VAR+129],r"""FRQ Droop Active""")
    psspy.var_channel([-1,PPC_VAR+451],r"""PPC_PvPwrRtSpnt""")
    psspy.var_channel([-1,PPC_VAR+452],r"""PPC_BatPwrRtSpnt""")
    psspy.var_channel([-1,PPC_VAR+270],r"""Automode""")
    

    # SF VAR Chanels
    psspy.var_channel([-1,SF_INV_VAR+3],r"""SF_init_bus_voltage""")
    psspy.var_channel([-1,SF_INV_VAR+9],r"""SF_Iq_cmd_before_dyn_limit_block""")
    psspy.var_channel([-1,SF_INV_VAR+73],r"""SF_VolSpt_modelInit""")
    psspy.var_channel([-1,SF_INV_VAR+74],r"""SF_VolSpt_ssSol""")
    psspy.var_channel([-1,SF_INV_VAR+163],r"""SF FRT Detect""")
    psspy.var_channel([-1,SF_INV_VAR+251],r"""SF PLL SubStt""")
    psspy.var_channel([-1,SF_INV_VAR+186],r"""SF LVRT Detected Flag""")
    psspy.var_channel([-1,SF_INV_VAR+187],r"""SF HVRT Detected Flag""")
    psspy.var_channel([-1,SF_INV_VAR+16],r"""SF Measured bus frequency""")


def fault_impedance(Vbase,MVArating,SCR,X_R,fault_factor):
    # Calculate Fault Impedance  
    Znet= float(Vbase*Vbase/(SCR*MVArating))
    Rnet= float(Znet/((1+X_R**2))**(1.0/2.0))
    Xnet= float(Rnet*X_R)
    global Rf, Xf
    Xf = float(fault_factor*Xnet)
    if fault_factor == 0:
        Rf = 0.001
    else:
        Rf = float(fault_factor*Rnet)
        
    return Rf,Xf

def save_parameters(Case_no): 
    # Obtain POC Q flow and Grid V setpoint solution from PSSE
    ierr, rval_V=psspy.busdat(INF_BUS, 'PU') 
    ierr, rval_Txtap = psspy.xfrdat(POC_BUS,DUMMYBUS,'1','RATIO') 
    
    try:
        for book in (xw.books):
            if book.fullname == excel_file:
                book.close()
                time.sleep(1)
    except:
        pass
    #         print(book)
    #         wb = book
    #         ws =  wb.sheets[0]
    #         vsource_col = ws.range('1:1').value.index('Vsource_i') + 1
    #         tap_col = ws.range('1:1').value.index('Trf_Tap') + 1
    #         Rf_Ohms_col = ws.range('1:1').value.index('Rf_Ohms') + 1
    #         LH_col = ws.range('1:1').value.index('L_H') + 1
    #         ws.range(vsource_col,Case_no)
    #         #print(column_index)
            #break
    try:
        # # Print Fault Impedance, POC Q flow and Grid V setpoint to CSV File                 
        df = pd.read_excel(excel_file, index_col='Case')
        df.loc[float(Case_no),'Vsource_i'] = rval_V
        df.loc[float(Case_no),'Trf_Tap'] = rval_Txtap
        df.loc[float(Case_no),'Rf_Ohms'] = Rf
        df.loc[float(Case_no),'L_H'] = Xf/(2*np.pi*50)
        # df.loc[float(Case_no),'PPCVset'] = rval_VSpt
        df.to_excel(excel_file)
    except:
        print("Couldnt update excel file for Case %s" %Case_no)

def is_string_in_file(file_path, search_string):
    with open(file_path, 'r') as file:
        for line in file:
            if search_string in line and not line.startswith('/'):
                return True
    return False

def check_plb_in_dyr(dyr,Test_type):
    
    #9999 'USRMDL' 1 'PLBVFU1', 1, 1, 3, 4, 3, 6, 0,1,'OF3s',1,50,0,0/
    #9999 'USRMDL' 1 'PLBVFU1', 1, 1, 3, 4, 3, 6, 0,1,'OF4hz',1,50,0,0/
    #9999 'USRMDL' 1 'PLBVFU1', 1, 1, 3, 4, 3, 6, 0,1,'UF3s',1,50,0,0/
    #9999 'USRMDL' 1 'PLBVFU1', 1, 1, 3, 4, 3, 6, 0,1,'UF4hz',1,50,0,0/
    #9999 'USRMDL' 1 'PLBVFU1', 1, 1, 3, 4, 3, 6, 1,0,'VT_ramp',1,50,0,0/
    #9999 'USRMDL' 1 'PLBVFU1', 1, 1, 3, 4, 3, 6, 1,0,'Vdiprecover',1,50,0,0/
    #9999 'USRMDL' 1 'PLBVFU1', 1, 1, 3, 4, 3, 6, 1,0,'HVLVRT',1,50,0,0 /Playback
    
    if Test_type == "DMAT_OverFreq_3s":
        if is_string_in_file(dyr,"OF3s") != 1: print("PLB NOT IN DYR FILE! %s" %dyr)
        
    if Test_type == "DMAT_UnderFreq_3s":
        if is_string_in_file(dyr,"UF3s") != 1: print("PLB NOT IN DYR FILE! %s" %dyr)
    
def create_result_csv(ResultsDir,infname):
    print(OutDir + infname + ".out")
    # Saving the output in csv file format
    chnfobj=dyntools.CHNF(OutDir + infname + ".out")
    chnfobj.csvout(channels=[], csvfile = setupdir + DMATdir +'temp.csv')
    # chnfobj.csvout(channels=[], csvfile='temp.csv')
    # Deleting the first row of csv file and saving the data in new csv file
    data = "".join(open("temp.csv").readlines()[1:])
    open(ResultsDir + "PSSE_" + infname + ".csv","wb").write(data)

    # Deleting the temporary csv file
    os.remove('temp.csv')
               
def init_sim(Test_Type,infname,dyrfile,savefile,SCR,X_R,Case_no,FRT_dip,T_flt):
    # global bus_flt,bus_IDTRF,bus_inf
    SF_INV_VAR = 1 + 0 #(Add six for PLB file)
    #BESS_INV_VAR = 301 + 0 #(Add six for PLB file)
    PPC_VAR = 627 + 0 #(Add six for PLB file)
    
    print("Running %s --> %s" %(Test_Type,infname))
    # PSSE Start and Load Flow Solve
    hide_output()
    psspy.psseinit()
    psspy.lines_per_page_one_device(1,60)
    It = infname
    psspy.lines_per_page_one_device(1,60)	
    psspy.progress_output(2, LogDir + It + "_Progress.txt",[0,0])
    hide_output(0)
    print('savefile: ',savefile)
    psspy.case(savefile + ".sav")
    

    psspy.fnsl([0,0,0,0,0,0,0,0])
    psspy.fnsl([0,0,0,0,0,0,0,0])
    psspy.fnsl([0,0,0,0,0,0,0,0])

    
    if Test_Type == "GridPhaseAngle_step":
        print('not running different config for tx')# Vgrid_phase_angle_config()###
    elif Test_Type in ["S52513_Vgridsteps","DMAT_Vgridsteps","S52513_Vgridsteps_Qmode"]:
        plb_test = "VT_5pc"
        Vgrid_Step_config(plb_test, del_V1 = 0.05)
        dyrfile = modify_dyr_plb(dyrfile, plb_test) 
        SF_INV_VAR += 6 #(Add six for PLB file)
        #BESS_INV_VAR += 6 #(Add six for PLB file)
        PPC_VAR += 6 #(Add six for PLB file)     
    elif Test_Type in ["S5253_OFtrip","S5253_UFtrip","S52511_OF","S52511_UF","FDroopCheck","S52511_small","DMAT_OverFreq_3s","DMAT_UnderFreq_3s", "DMAT_OverFreq_4Hz","DMAT_UnderFreq_4Hz"]:
        if Test_Type == "S5253_OFtrip":
            plb_test = "S5253_OFtrip"
        elif Test_Type == "S5253_UFtrip":
            plb_test = "S5253_UFtrip"
        elif Test_Type == "S52511_OF":
            plb_test = "S52511_OF"
        elif Test_Type == "S52511_UF":
            plb_test = "S52511_UF"
        elif Test_Type == "FDroopCheck":
            plb_test = "FDroopCheck"
        elif Test_Type == "S52511_small":
            plb_test = "S52511_small"
        elif Test_Type == "DMAT_OverFreq_3s":
            plb_test = "OF3s"
        elif Test_Type == "DMAT_UnderFreq_3s":
            plb_test = "UF3s"
        elif Test_Type == "DMAT_OverFreq_4Hz":
            plb_test = "OF4Hz"
        elif Test_Type == "DMAT_UnderFreq_4Hz":
            plb_test = "UF4Hz"
        Freq_test_config(plb_test)
        dyrfile = modify_dyr_plb(dyrfile, plb_test)  
        SF_INV_VAR += 6 #(Add six for PLB file)
        #BESS_INV_VAR += 6 #(Add six for PLB file)
        PPC_VAR += 6 #(Add six for PLB file)
    elif Test_Type in ["GridVchange_ramp"]:
        plb_test = "VT_ramp"
        Vgrid_ramp_config(plb_test, del_V1=0.10)
        dyrfile = modify_dyr_plb(dyrfile, plb_test)  
        SF_INV_VAR += 6 #(Add six for PLB file)
        #BESS_INV_VAR += 6 #(Add six for PLB file)
        PPC_VAR += 6 #(Add six for PLB file)
    elif Test_Type in ["S5254_LVRT", "S5254_HVRT", "S5254_CUO"]:
        if Test_Type == "S5254_LVRT":
            plb_test = "S5254_LVRT"
        elif Test_Type == "S5254_HVRT":
            plb_test = "S5254_HVRT"
        elif Test_Type == "S5254_CUO":
            plb_test = "S5254_CUO"
        Vgrid_Step_2_config(plb_test, 0.10,T_flt)
        dyrfile = modify_dyr_plb(dyrfile, plb_test)  
        SF_INV_VAR += 6 #(Add six for PLB file)
        #BESS_INV_VAR += 6 #(Add six for PLB file)
        PPC_VAR += 6 #(Add six for PLB file)
    elif Test_Type in ["S5255_HVRT"]:
        plb_test = "S5255_HVRT"
        Vgrid_Step_2_config(plb_test, FRT_dip,T_flt)
        dyrfile = modify_dyr_plb(dyrfile, plb_test)  
        SF_INV_VAR += 6 #(Add six for PLB file)
        #BESS_INV_VAR += 6 #(Add six for PLB file)
        PPC_VAR += 6 #(Add six for PLB file)
    elif Test_Type in ["Extended_Vdip_Recovery_0.8pu","Extended_Vdip_Recovery_0.5pu","Extended_Vdip_Recovery_0.1pu"]:
        plb_test = "Vdiprecover"
        Vgrid_dip_config(plb_test, Test_Type)
        dyrfile = modify_dyr_plb(dyrfile, plb_test)  
        SF_INV_VAR += 6 #(Add six for PLB file)
        #BESS_INV_VAR += 6 #(Add six for PLB file)
        PPC_VAR += 6 #(Add six for PLB file)
    elif Test_Type in ["Additional_ROCOF_0.5", "Additional_ROCOF_0.1", "Additional_ROCOF_1.0", "Additional_ROCOF_4.0"]:
        plb_test = Test_Type
        if Test_Type == "Additional_ROCOF_0.1":
            plb_test = "ROCOF_01"
        elif Test_Type == "Additional_ROCOF_0.5":
            plb_test = "ROCOF_05"
        elif Test_Type == "Additional_ROCOF_1.0":
            plb_test = "ROCOF_1"
        elif Test_Type == "Additional_ROCOF_4.0":
            plb_test = "ROCOF_4"
        # ROCOF_Freq_test_config(plb_test)
        dyrfile = modify_dyr_plb(dyrfile, plb_test)  
        SF_INV_VAR += 6 #(Add six for PLB file)
        #BESS_INV_VAR += 6 #(Add six for PLB file)
        PPC_VAR += 6 #(Add six for PLB file)
    #########################DMATS#########################
    elif Test_Type in ["LVRT","HVRT"]:
        plb_test = "HVLVRT"
        HVRT_LVRT_config(FRT_dip)
        dyrfile = modify_dyr_plb(dyrfile, plb_test)  
        SF_INV_VAR += 6 #(Add six for PLB file)
        #BESS_INV_VAR += 6 #(Add six for PLB file)
        PPC_VAR += 6 #(Add six for PLB file)    
    elif Test_Type in ["Extended_Vdip_Recovery_0.8pu","Extended_Vdip_Recovery_0.5pu","Extended_Vdip_Recovery_0.1pu"]:
        plb_test = "Vdiprecover"
        Vgrid_dip_config(plb_test, Test_Type)
        dyrfile = modify_dyr_plb(dyrfile, plb_test)  
        SF_INV_VAR += 6 #(Add six for PLB file)
        #BESS_INV_VAR += 6 #(Add six for PLB file)
        PPC_VAR += 6 #(Add six for PLB file)
    elif Test_Type in ["GridVchange_step"]:
        plb_test = "VT_step_10pc"
        Vgrid_Step_2_config(plb_test, 0.10,T_flt)
        dyrfile = modify_dyr_plb(dyrfile, plb_test)  
        SF_INV_VAR += 6 #(Add six for PLB file)
        #BESS_INV_VAR += 6 #(Add six for PLB file)
        PPC_VAR += 6 #(Add six for PLB file)
    elif Test_Type in ["DMAT_OverFreq_3s","DMAT_UnderFreq_3s", "DMAT_OverFreq_4Hz","DMAT_UnderFreq_4Hz"]:
        if Test_Type == "DMAT_OverFreq_3s":
            plb_test = "OF3s"
        elif Test_Type == "DMAT_UnderFreq_3s":
            plb_test = "UF3s"
        elif Test_Type == "DMAT_OverFreq_4Hz":
            plb_test = "OF4Hz"
        elif Test_Type == "DMAT_UnderFreq_4Hz":
            plb_test = "UF4Hz"
        Freq_test_config(plb_test)
        dyrfile = modify_dyr_plb(dyrfile, plb_test)  
        SF_INV_VAR += 6 #(Add six for PLB file)
        #BESS_INV_VAR += 6 #(Add six for PLB file)
        PPC_VAR += 6 #(Add six for PLB file)
    elif Test_Type in ["DMAT_Vgridsteps"]:
        plb_test = "VT_5pc"
        Vgrid_Step_config(plb_test, del_V1 = 0.05)
        dyrfile = modify_dyr_plb(dyrfile, plb_test) 
        SF_INV_VAR += 6 #(Add six for PLB file)
        #BESS_INV_VAR += 6 #(Add six for PLB file)
        PPC_VAR += 6 #(Add six for PLB file)
    elif Test_Type in ["S5255_IqTest"]:
        plb_test = "IqVgrids"
        Iq_Vgrid(plb_test, FRT_dip)
        dyrfile = modify_dyr_plb(dyrfile, plb_test) 
        SF_INV_VAR += 6 #(Add six for PLB file)
        #BESS_INV_VAR += 6 #(Add six for PLB file)
        PPC_VAR += 6 #(Add six for PLB file)
    # elif Test_Type in ["S52511_small"]:
    #     plb_test = "S52511_small"
    #     Freq_test_config(plb_test)
    #     dyrfile = modify_dyr_plb(dyrfile, plb_test) 
    #     # SF_INV_VAR += 6 #(Add six for PLB file)
    #     BESS_INV_VAR += 6 #(Add six for PLB file)
    #     PPC_VAR += 6 #(Add six for PLB file)

    psspy.cong(0)
    psspy.conl(0,1,1,[0,0],[ 1.0, 1.0, 1.0, 0.1])
    psspy.conl(0,1,2,[0,0],[ 1.0, 1.0, 1.0, 0.1])
    psspy.conl(0,1,3,[0,0],[ 1.0, 1.0, 1.0, 0.1])
    psspy.ordr(0)
    psspy.ordr(0)
    psspy.fact()
    psspy.tysl(1)
    psspy.dyre_new([1,1,1,1],dyrfile,"","","")
    
    dll_files = [file for file in os.listdir(setupdir) if file.endswith('.dll')]
    # print(dll_files)
    for dll in dll_files:
        # ierr = psspy.addmodellibrary(dll)
        # print('error code',ierr)
        # print(setupdir+dll)
        psspy.addmodellibrary(setupdir+dll)
        
    # Vgrid_Step_config()
    # Freq_test_config()
    # Vgrid_Step_2_config()
    # Vgrid_ramp_config()
    # HVRT_LVRT_config(FRT_dip)
    # Vgrid_dip_config(Test_type=Test_Type)
    
    psspy.dynamics_solution_param_2([iterations,_i,_i,_i,_i,_i,_i,_i],[sfactor,tolerance,dT,ffilter,_f,_f,_f,_f]) #IN ORDER: iterations, acceleration factor, tolerance, timestep, frequency filter.	
    psspy.set_netfrq(0)
    add_channels(SF_INV_VAR, PPC_VAR)

    # PSSE OUT File Setup -------------------------------------------------------update if base run works
    psspy.strt(0,OutDir + infname + ".out")
    psspy.strt(0,OutDir + infname + ".out")
    psspy.strt(0,OutDir + infname + ".out")

def modify_dyr_plb(dyrfile, plb_test):
    # Change the dyr file to include the PLB model
    base_dyr = dyrfile.split("\\")[-1]
    #plb_dyr = setupdir + base_dyr.split('.')[0] + "_" + plb_test + ".dyr"
    plb_dyr = setupdir + DMATdir + base_dyr.split('.')[0] + "_" + plb_test + ".dyr"
    if os.path.exists(plb_dyr):
        os.remove(plb_dyr)
    if not os.path.exists(plb_dyr):
        shutil.copy(dyrfile, plb_dyr)
        
        with open(plb_dyr, 'r') as file:
            lines = file.readlines()

        # Modify the lines
        modified_lines = []
        for line in lines:
            if line.strip().startswith(str(INF_BUS) + " 'GENCLS'"):
                modified_lines.append("//" + line)
                # Add the plb line for the infinite bus generator
                # plb_test = 'TCT_NV'
                plb_model_config = str(INF_BUS) + " 'USRMDL' 1 'PLBVFU1' 1 1 3 4 3 6 1 1 " + "'%s'"%plb_test + " 1.0 50.0 0.0 0.0 // \n"
                modified_lines.append(plb_model_config)
            else:
                modified_lines.append(line)

        # Write the modified lines back to the file
        with open(plb_dyr, 'w') as file:
            file.writelines(modified_lines)
  
    return plb_dyr

def Vgrid_Step_config(plb_test, del_V1 = 0.05):
    ### Read Playback (PLB) file and adjust setpoints for each study###
    ierr, rval_V1 = psspy.busdat(INF_BUS, 'PU')  #read voltage from grid
    my_file = open(plb_test + ".plb", "w")     #Open text file (PLB) used by DYR file
    my_file.write("1.00,"+str(rval_V1)+",50\n")  #Set your PLB script text here
    my_file.write("10.00,"+str(rval_V1)+",50\n")   #Order = Sim Time (s), Voltage (pu), Frequency (Hz)
    my_file.write("10.001,"+str(rval_V1+del_V1)+",50\n")
    my_file.write("25.00,"+str(rval_V1+del_V1)+",50\n")
    my_file.write("25.001,"+str(rval_V1)+",50\n")
    my_file.write("40.00,"+str(rval_V1)+",50\n")
    my_file.write("40.001,"+str(rval_V1-del_V1)+",50\n")
    my_file.write("55.00,"+str(rval_V1-del_V1)+",50\n")
    my_file.write("55.001,"+str(rval_V1)+",50")        
    my_file = open(plb_test + ".plb")          #Opening PLB file again saves the file.

def Iq_Vgrid(plb_test, FRT_dip):
    ### Read Playback (PLB) file and adjust setpoints for each study###
    ierr, rval_V1 = psspy.busdat(INF_BUS, 'PU')  #read voltage from grid

    my_file = open(plb_test + ".plb", "w")     #Open text file (PLB) used by DYR file
    my_file.write("1.00,"+str(rval_V1)+",50\n")  #Set your PLB script text here
    my_file.write("15.00,"+str(rval_V1)+",50\n")   #Order = Sim Time (s), Voltage (pu), Frequency (Hz)
    my_file.write("15001,"+str(FRT_dip)+",50\n")
    my_file.write("15.430,"+str(FRT_dip)+",50\n")
    my_file.write("15.431,"+str(rval_V1)+",50\n")
    my_file.write("25.00,"+str(rval_V1)+",50\n")        
    my_file = open(plb_test + ".plb")          #Opening PLB file again saves the file.

def Freq_test_config(plb_test):
# UF Trip ------------
# 0.001,1.00,50
# 50.00,1.00,50
# 50.75,1.00,47
# 175.00,1.00,47
# 175.25,1.00,48
# 655.00,1.00,48
# 655.50,1.00,50
# 750.00,1.00,50
# ------------
# OF Trip ------------
# 0.001,1.00,50
# 50.00,1.00,50
# 50.50,1.00,52
# 655.00,1.00,52
# 655.50,1.00,50
# 750.00,1.00,50
# ------------
    if plb_test == "S5253_OFtrip":
        ierr, rval_V1=psspy.busdat(INF_BUS, 'PU') #read grid voltage
        #print('rval_V1')
        #print(rval_V1)
        my_file = open(plb_test + ".plb", "w")     
        my_file.write("0.001,"+str(rval_V1)+",50\n")  
        my_file.write("50.00,"+str(rval_V1)+",50\n")   
        my_file.write("50.50,"+str(rval_V1)+",52\n")
        my_file.write("650.50,"+str(rval_V1)+",52\n")
        my_file.write("651.00,"+str(rval_V1)+",50\n")
        my_file.write("750.00,"+str(rval_V1)+",50\n")
        my_file = open(plb_test + ".plb")
        
    elif plb_test == "S5253_UFtrip":
        ierr, rval_V1=psspy.busdat(INF_BUS, 'PU') #read grid voltage
        my_file = open(plb_test + ".plb", "w")     #Open text file (PLB) used by DYR file
        my_file.write("0.001,"+str(rval_V1)+",50\n")  
        my_file.write("50.00,"+str(rval_V1)+",50\n")   
        my_file.write("50.75,"+str(rval_V1)+",47\n")
        my_file.write("170.75,"+str(rval_V1)+",47\n")
        my_file.write("171.00,"+str(rval_V1)+",48\n")
        my_file.write("650.50,"+str(rval_V1)+",48\n")
        my_file.write("651.00,"+str(rval_V1)+",50\n")
        my_file.write("750.00,"+str(rval_V1)+",50\n")
        my_file = open(plb_test + ".plb")

    elif plb_test == "S52511_UF":
        ierr, rval_V1=psspy.busdat(INF_BUS, 'PU') #read grid voltage
        #print('rval_V1')
        #print(rval_V1)
        my_file = open(plb_test + ".plb", "w")     
        my_file.write("5.00,"+str(rval_V1)+",50\n")  
        my_file.write("5.25,"+str(rval_V1)+",49\n")   
        my_file.write("15.00,"+str(rval_V1)+",49\n")
        my_file.write("15.25,"+str(rval_V1)+",50\n")
        my_file.write("25.00,"+str(rval_V1)+",50\n")
        my_file.write("25.50,"+str(rval_V1)+",48\n")
        my_file.write("35.00,"+str(rval_V1)+",48\n")
        my_file.write("35.50,"+str(rval_V1)+",50\n")
        my_file.write("45.00,"+str(rval_V1)+",50\n")
        my_file = open(plb_test + ".plb")
      
    elif plb_test == "S52511_OF":
        ierr, rval_V1=psspy.busdat(INF_BUS, 'PU') #read grid voltage
        my_file = open(plb_test + ".plb", "w")     #Open text file (PLB) used by DYR file

        
        my_file.write("5.00,"+str(rval_V1)+",50\n")  
        my_file.write("5.25,"+str(rval_V1)+",51.01\n")   
        my_file.write("15.00,"+str(rval_V1)+",51.01\n")
        my_file.write("15.25,"+str(rval_V1)+",50\n")
        my_file.write("25.00,"+str(rval_V1)+",50\n")
        my_file.write("25.50,"+str(rval_V1)+",52.02\n")
        my_file.write("35.00,"+str(rval_V1)+",52.02\n")
        my_file.write("35.50,"+str(rval_V1)+",50\n")
        my_file.write("40.00,"+str(rval_V1)+",50\n")
        my_file = open(plb_test + ".plb")

    elif plb_test == "S52511_small":
        ierr, rval_V1=psspy.busdat(INF_BUS, 'PU') #read grid voltage
        my_file = open(plb_test + ".plb", "w")     #Open text file (PLB) used by DYR file
        my_file.write("5.00,"+str(rval_V1)+",50\n")  
        my_file.write("5.25,"+str(rval_V1)+",50.25\n")   
        my_file.write("15.00,"+str(rval_V1)+",50.25\n")
        my_file.write("15.25,"+str(rval_V1)+",50\n")
        my_file.write("25.00,"+str(rval_V1)+",50\n")
        my_file.write("25.50,"+str(rval_V1)+",50.5\n")
        my_file.write("35.00,"+str(rval_V1)+",50.5\n")
        my_file.write("35.50,"+str(rval_V1)+",50\n")
        my_file.write("45.00,"+str(rval_V1)+",50\n")
        my_file.write("45.50,"+str(rval_V1)+",49.75\n")
        my_file.write("55.00,"+str(rval_V1)+",49.75\n")
        my_file.write("55.50,"+str(rval_V1)+",50\n")
        my_file.write("65.00,"+str(rval_V1)+",50\n")
        my_file.write("65.50,"+str(rval_V1)+",49.5\n")
        my_file.write("75.00,"+str(rval_V1)+",49.5\n")
        my_file.write("75.50,"+str(rval_V1)+",50\n")
        my_file.write("85.00,"+str(rval_V1)+",50\n")
        my_file = open(plb_test + ".plb")

    elif plb_test == "FDroopCheck":
        ierr, rval_V1=psspy.busdat(INF_BUS, 'PU') #read grid voltage
        my_file = open(plb_test + ".plb", "w")     #Open text file (PLB) used by DYR file
        my_file.write("5.00,"+str(rval_V1)+",50\n")  
        my_file.write("5.25,"+str(rval_V1)+",50.015\n")   
        my_file.write("15.00,"+str(rval_V1)+",50.015\n")
        my_file.write("15.25,"+str(rval_V1)+",50\n")
        my_file.write("25.00,"+str(rval_V1)+",50\n")
        my_file.write("25.50,"+str(rval_V1)+",51.766\n")
        my_file.write("35.00,"+str(rval_V1)+",51.766\n")
        my_file.write("35.50,"+str(rval_V1)+",50\n")
        my_file.write("45.00,"+str(rval_V1)+",50\n")
        my_file.write("45.50,"+str(rval_V1)+",53.517\n")
        my_file.write("55.00,"+str(rval_V1)+",53.517\n")
        my_file.write("55.50,"+str(rval_V1)+",50\n")
        my_file.write("65.00,"+str(rval_V1)+",50\n")
        my_file.write("65.50,"+str(rval_V1)+",49.985\n")
        my_file.write("75.00,"+str(rval_V1)+",49.985\n")
        my_file.write("75.50,"+str(rval_V1)+",50\n")
        my_file.write("85.00,"+str(rval_V1)+",50\n")
        my_file.write("85.50,"+str(rval_V1)+",48.234\n")
        my_file.write("95.00,"+str(rval_V1)+",48.234\n")
        my_file.write("95.50,"+str(rval_V1)+",50\n")
        my_file.write("105.00,"+str(rval_V1)+",50\n")
        my_file.write("105.50,"+str(rval_V1)+",46.483\n")
        my_file.write("115.00,"+str(rval_V1)+",46.483\n")
        my_file.write("115.50,"+str(rval_V1)+",50\n")
        my_file.write("125.00,"+str(rval_V1)+",50\n")
        my_file = open(plb_test + ".plb")

    elif plb_test == "OF3s":
        ierr, rval_V1=psspy.busdat(INF_BUS, 'PU') #read grid voltage
        #print('rval_V1')
        #print(rval_V1)
        my_file = open(plb_test + ".plb", "w")     #Open text file (PLB) used by DYR file
        my_file.write("15.00,"+str(rval_V1)+",50\n")  #Set your PLB script text here
        my_file.write("18.00,"+str(rval_V1)+",52\n")   #Order = Sim Time (s), Voltage (pu), Frequency (Hz)
        my_file.write("25.00,"+str(rval_V1)+",52\n")
        my_file.write("28.00,"+str(rval_V1)+",50\n")
        my_file = open(plb_test + ".plb")
        
    elif plb_test == "OF4Hz":
        ierr, rval_V1=psspy.busdat(INF_BUS, 'PU') #read grid voltage
        my_file = open(plb_test + ".plb", "w")     #Open text file (PLB) used by DYR file
        my_file.write("15.00,"+str(rval_V1)+",50\n")  #Set your PLB script text here
        my_file.write("15.375,"+str(rval_V1)+",51.5\n")   #Order = Sim Time (s), Voltage (pu), Frequency (Hz)
        my_file.write("25.00,"+str(rval_V1)+",51.5\n")
        my_file.write("25.375,"+str(rval_V1)+",50\n")
        my_file.write("35.00,"+str(rval_V1)+",50\n")  
        my_file.write("35.4375,"+str(rval_V1)+",51.75\n")  
        my_file.write("45.00,"+str(rval_V1)+",51.75\n")
        my_file.write("45.4375,"+str(rval_V1)+",50\n")
        my_file.write("55,"+str(rval_V1)+",50\n")
        my_file.write("55.5,"+str(rval_V1)+",52\n")  
        my_file.write("65.00,"+str(rval_V1)+",52\n")
        my_file.write("65.5,"+str(rval_V1)+",50\n")
        my_file.write("75,"+str(rval_V1)+",50\n")
                
        my_file = open(plb_test + ".plb")
    
    elif plb_test == "UF3s":
        ierr, rval_V1=psspy.busdat(INF_BUS, 'PU') #read grid voltage
        #print('rval_V1')
        #print(rval_V1)
        my_file = open(plb_test + ".plb", "w")     #Open text file (PLB) used by DYR file
        my_file.write("15.00,"+str(rval_V1)+",50\n")  #Set your PLB script text here
        my_file.write("18.00,"+str(rval_V1)+",47\n")   #Order = Sim Time (s), Voltage (pu), Frequency (Hz)
        my_file.write("25.00,"+str(rval_V1)+",47\n")
        my_file.write("28.00,"+str(rval_V1)+",50\n")
        my_file = open(plb_test + ".plb")
    
    elif plb_test == "UF4Hz":    
        ierr, rval_V1=psspy.busdat(INF_BUS, 'PU') #read grid voltage
        #print('rval_V1')
        #print(rval_V1)
        my_file = open(plb_test + ".plb", "w")     #Open text file (PLB) used by DYR file
        my_file.write("15.00,"+str(rval_V1)+",50\n")  #Set your PLB script text here
        my_file.write("15.750,"+str(rval_V1)+",47\n")   #Order = Sim Time (s), Voltage (pu), Frequency (Hz)
        my_file.write("25.00,"+str(rval_V1)+",47\n")
        my_file.write("25.750,"+str(rval_V1)+",50\n")
        my_file = open(plb_test + ".plb")

def Vgrid_Step_2_config(plb_test, del_V1, T_flt):
    if plb_test == "S5254_CUO":
        ierr, rval_V1=psspy.busdat(INF_BUS, 'PU')
        my_file = open(plb_test + ".plb", "w")             #Open text file (PLB) used by DYR file
        my_file.write("0.01,"+str(rval_V1)+",50\n")       #Set your PLB script text here. 
        my_file.write("10.00,"+str(rval_V1)+",50\n")  #Order = Sim Time (s), Voltage (pu), Frequency (Hz)
        my_file.write("10.001,"+str(1.1)+",50\n")
        my_file.write("30.00,"+str(1.1)+",50\n")
        my_file.write("30.001,"+str(rval_V1)+",50\n")
        my_file.write("50.00,"+str(rval_V1)+",50\n")
        my_file.write("50.001,"+str(0.9)+",50\n")
        my_file.write("70.00,"+str(0.9)+",50\n") 
        my_file.write("70.001,"+str(rval_V1)+",50\n")
        my_file.write("90,"+str(rval_V1)+",50\n")        
        my_file = open(plb_test + ".plb")                   #Opening PLB file again saves the file.
    elif plb_test == "S5254_HVRT":
        ierr, rval_V1=psspy.busdat(INF_BUS, 'PU')
        my_file = open(plb_test + ".plb", "w")             #Open text file (PLB) used by DYR file
        my_file.write("1.00,"+str(rval_V1)+",50\n")       #Set your PLB script text here. 
        my_file.write("10.00,"+str(rval_V1)+",50\n")  #Order = Sim Time (s), Voltage (pu), Frequency (Hz)
        my_file.write("10.001,"+str(1.35)+",50\n")
        my_file.write("10.02,"+str(1.35)+",50\n")
        my_file.write("10.021,"+str(1.3)+",50\n")
        my_file.write("10.2,"+str(1.3)+",50\n")
        my_file.write("10.201,"+str(1.25)+",50\n")
        my_file.write("12.00,"+str(1.25)+",50\n")
        my_file.write("12.001,"+str(1.2)+",50\n")
        my_file.write("30.00,"+str(1.2)+",50\n")
        my_file.write("30.001,"+str(1.15)+",50\n")      #1.149
        my_file.write("130.00,"+str(1.15)+",50\n")      #1.149
        my_file.write("130.001,"+str(rval_V1)+",50\n")
        my_file.write("145.00,"+str(rval_V1)+",50\n")          
        my_file = open(plb_test + ".plb")                   #Opening PLB file again saves the file.
    elif plb_test == "S5254_LVRT":
        ierr, rval_V1=psspy.busdat(INF_BUS, 'PU')
        my_file = open(plb_test + ".plb", "w")             #Open text file (PLB) used by DYR file
        my_file.write("1.00,"+str(rval_V1)+",50\n")       #Set your PLB script text here. 
        my_file.write("10.00,"+str(rval_V1)+",50\n")  #Order = Sim Time (s), Voltage (pu), Frequency (Hz)
        my_file.write("10.001,"+str(0.7)+",50\n")
        my_file.write("12.00,"+str(0.7)+",50\n")
        my_file.write("12.001,"+str(0.8)+",50\n")
        my_file.write("20.00,"+str(0.8)+",50\n")
        my_file.write("20.001,"+str(rval_V1)+",50\n")
        my_file.write("35.00,"+str(rval_V1)+",50\n")          
        my_file = open(plb_test + ".plb")                   #Opening PLB file again saves the file.
    elif plb_test == "S5255_HVRT":
        ierr, rval_V1=psspy.busdat(INF_BUS, 'PU')
        my_file = open(plb_test + ".plb", "w")             #Open text file (PLB) used by DYR file
        my_file.write("1.00,"+str(rval_V1)+",50\n")       #Set your PLB script text here. 
        my_file.write("10.00,"+str(rval_V1)+",50\n")  #Order = Sim Time (s), Voltage (pu), Frequency (Hz)
        my_file.write("10.001,"+str(1-del_V1)+",50\n")#
        my_file.write(str(10+T_flt)+","+str(1-del_V1)+",50\n")#
        my_file.write(str(10+T_flt+0.001)+","+str(rval_V1)+",50\n")
        my_file.write("15.00,"+str(rval_V1)+",50\n")    
        my_file = open(plb_test + ".plb")                   #Opening PLB file again saves the file.
    else:
        ierr, rval_V1=psspy.busdat(INF_BUS, 'PU')
        my_file = open(plb_test + ".plb", "w")             #Open text file (PLB) used by DYR file
        my_file.write("15.00,"+str(rval_V1)+",50\n")       #Set your PLB script text here. 
        my_file.write("15.0500,"+str(rval_V1+del_V1)+",50\n")  #Order = Sim Time (s), Voltage (pu), Frequency (Hz)
        my_file.write("25.00,"+str(rval_V1 + del_V1)+",50\n")
        my_file.write("25.0500,"+str(rval_V1 - del_V1)+",50\n")
        my_file.write("35.00,"+str(rval_V1 - del_V1)+",50\n")
        my_file.write("35.0500,"+str(rval_V1 + del_V1)+",50\n")
        my_file.write("45.00,"+str(rval_V1 + del_V1)+",50\n")
        my_file.write("45.0500,"+str(rval_V1)+",50")        
        my_file = open(plb_test + ".plb")                   #Opening PLB file again saves the file.

def Vgrid_ramp_config(plb_test, del_V1=0.10):    
    ierr, rval_V1=psspy.busdat(INF_BUS, 'PU')
    # del_V1=0.10
    my_file = open(plb_test + ".plb", "w")               #Open text file (PLB) used by DYR file
    my_file.write("15.000,"+str(rval_V1)+",50\n")        #Set your PLB script text here. 
    my_file.write("21.000,"+str(rval_V1 - del_V1)+",50\n")  #Order = Sim Time (s), Voltage (pu), Frequency (Hz)
    my_file.write("30.00,"+str(rval_V1 - del_V1)+",50\n")
    my_file.write("36.000,"+str(rval_V1 + del_V1)+",50\n")
    my_file.write("45.000,"+str(rval_V1 + del_V1)+",50\n")
    my_file.write("51.000,"+str(rval_V1 - del_V1)+",50\n")
    my_file.write("60.000,"+str(rval_V1 - del_V1)+",50\n")
    my_file.write("66.000,"+str(rval_V1)+",50\n")
    my_file.write("75.000,"+str(rval_V1)+",50")        
    my_file = open(plb_test + ".plb")

def Vgrid_dip_config(plb_test, Test_Type):

    vdip = float(Test_Type.split("_")[-1][:-2])

    ierr, rval_V1=psspy.busdat(INF_BUS, 'PU')
    #print('rval_V1')
    #print(rval_V1)
    my_file = open(plb_test + ".plb", "w")     #Open text file (PLB) used by DYR file
    my_file.write("15.00,"+str(rval_V1)+",50\n")  #Set your PLB script text here
    my_file.write("15.001,"+str(vdip)+",50\n")   #Order = Sim Time (s), Voltage (pu), Frequency (Hz) #grid voltage is targeted to be 0.1,0.5 and 0.8 pu so set accordingly
    my_file.write("15.430,"+str(vdip)+",50\n")   #grid voltage is targeted to be 0.1,0.5 and 0.8 pu so set accordingly
    my_file.write("15.431,"+str(0.8)+",50\n")   #20ms ramp to match PSCAD so no need to change this value
    my_file.write("16.430,"+str(rval_V1)+",50\n")     
    my_file = open(plb_test + ".plb")          #Opening PLB file again saves the file.


def HVRT_LVRT_config(FRT_dip):
        #9999 'USRMDL' 1 'PLBVFU1', 1, 1, 3, 4, 3, 6, 1,0,'HVLVRT',1,50,0,0 /Playback

    ierr, rval_V1=psspy.busdat(INF_BUS, 'PU')
    my_file = open("HVLVRT.plb", "w")     #Open text file (PLB) used by DYR file
    my_file.write("10.00,"+str(rval_V1)+",50\n")  #Set your PLB script text here
    my_file.write("10.01,"+str(rval_V1+FRT_dip)+",50\n")   #Order = Sim Time (s), Voltage (pu), Frequency (Hz)
    my_file.write("12.50,"+str(rval_V1+FRT_dip)+",50\n")
    my_file.write("12.51,"+str(rval_V1)+",50\n")
    my_file.write("25.00,"+str(rval_V1)+",50\n")
    my_file = open("HVLVRT.plb")

#********************************************************************** # 
#***************************** DMAT TESTS ***************************** #  
#********************************************************************** #   
def Balanced_Fault(T_flt, Rf, Xf, P_POC):
    # print(Rf,Xf)
    psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+18, P_POC*1000)     # Change DirectSales Setpoint   #
    psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+13, MWrating*1000)  # Change BessW Setpoint         #26000
    psspy.run(0,10,out_txt_steps,plot_steps,0)
    psspy.dist_3phase_bus_fault(DUMMYBUS,0,3, Vbase,[Rf, Xf])
    psspy.run(0, 10+T_flt,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 25,out_txt_steps,plot_steps,0)

def MFRT_Protection_1(Rf, Xf):
    psspy.run(1, 10.0,out_txt_steps,plot_steps,0) 
    # psspy.dist_bus_fault(100,3, Vbase,[Rf,Xf])
    # Bolted faults 0.12s duration 
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[0.01*Rf,0.01*Xf,0.0,0.0,0.0,0.0])
    psspy.run(1, 45.0,out_txt_steps,plot_steps,0) 
    
def MFRT_Protection_2(Rf, Xf):
    psspy.run(1, 10.0,out_txt_steps,plot_steps,0) 
    # psspy.dist_bus_fault(100,3, Vbase,[Rf,Xf])
    # Bolted faults 0.12s duration 
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[1*Rf,1*Xf,0.0,0.0,0.0,0.0])
    psspy.run(1, 45.0,out_txt_steps,plot_steps,0)                    
    
def MFRT_S1(Rf, Xf):
    ### Run MFRT Sequence P1 DMAT Study ###
    psspy.run(1, 5.0,out_txt_steps,plot_steps,0) 
                    
    # Bolted faults 0.12s duration 
    psspy.dist_bus_fault_3(1, Vbase,[0,1,DUMMYBUS,0,1],[0,-0.2E+10,0.0,0.0,0.0,0.0])
    psspy.run(1, 5.12,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 5.13,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(1, Vbase,[0,1,DUMMYBUS,0,1],[0,-0.2E+10,0.0,0.0,0.0,0.0])
    psspy.run(0, 5.25,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 5.26,out_txt_steps,plot_steps,0) 
    # Zf= 0.2Zs 0.12s duration
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[0.2*Rf,0.2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 5.38,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 5.58,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[0.2*Rf,0.2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 5.7,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 5.9,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[0.2*Rf,0.2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0,6.02,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 6.52,out_txt_steps,plot_steps,0) 
    # Zf=Zs 0.12s duration
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 6.64,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 7.14,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    # psspy.run(0,3.26,out_txt_steps,plot_steps,0) 
    psspy.run(0,7.26,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 8.01,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 8.13,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 9.13,out_txt_steps,plot_steps,0) 
    # Zf =Zs 0.220s duration 
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 9.35,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 10.85,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 11.07,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 13.07,out_txt_steps,plot_steps,0) 
    
    #Zf= 2Zs 0.220s duration 
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[2*Rf,2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0,13.29,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 15.29,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[2*Rf,2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 15.51,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 18.51,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[2*Rf,2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 18.73,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 23.73,out_txt_steps,plot_steps,0) 
    
    #Zf=3.5Zs 0.220 and 0.430s fault duration 
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[3.5*Rf,3.5*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 23.95,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 30.95,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[3.5*Rf,3.5*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 31.38,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 45,out_txt_steps,plot_steps,0)                       
    
def MFRT_S2(Rf, Xf):
    psspy.run(1, 5.0,out_txt_steps,plot_steps,0) 
    
    # Bolted faults 0.12s duration 
    psspy.dist_bus_fault_3(1, Vbase,[0,1,DUMMYBUS,0,1],[0,-0.2E+10,0.0,0.0,0.0,0.0])
    psspy.run(1, 5.43,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 5.44,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(1, Vbase,[0,1,DUMMYBUS,0,1],[0,-0.2E+10,0.0,0.0,0.0,0.0])
    psspy.run(0, 5.66,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0,5.67,out_txt_steps,plot_steps,0) 
    # Zf= 0.2Zs 0.12s duration
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[0.2*Rf,0.2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 5.89,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 6.09,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[0.2*Rf,0.2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 6.31,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 6.51,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[0.2*Rf,0.2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 6.73,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 7.23,out_txt_steps,plot_steps,0) 
    # Zf=Zs 0.12s duration
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 7.45,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 7.95,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 8.17,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0,8.92,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 9.04,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0,10.04,out_txt_steps,plot_steps,0) 
    # Zf =Zs 0.220s duration 
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 10.16,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 11.66,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 11.78,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 13.78,out_txt_steps,plot_steps,0) 
    
    #Zf= 2Zs 0.220s duration 
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[2*Rf,2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 13.9,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 15.9,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[2*Rf,2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 16.02,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0,19.02,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[2*Rf,2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 19.14,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 24.14,out_txt_steps,plot_steps,0) 
    
    #Zf=3.5Zs 0.220 and 0.430s fault duration 
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[3.5*Rf,3.5*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 24.26,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 31.26,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[3.5*Rf,3.5*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0,31.38,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 45,out_txt_steps,plot_steps,0)                                        
    
def MFRT_S3(Rf, Xf):
    psspy.run(1, 5.0,out_txt_steps,plot_steps,0) 

    # Bolted faults 0.12s duration 
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[0.2*Rf,0.2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(1, 5.22,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 15.22,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[0.2*Rf,0.2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 15.44,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0,22.44,out_txt_steps,plot_steps,0) 
    
    # Zf= 0.2Zs 0.12s duration
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[0.2*Rf,0.2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 22.66,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 27.66,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(1, Vbase,[0,1,DUMMYBUS,0,1],[0,-0.2E+10,0.0,0.0,0.0,0.0])
    psspy.run(0, 27.88,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 30.88,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(1, Vbase,[0,1,DUMMYBUS,0,1],[0,-0.2E+10,0.0,0.0,0.0,0.0])
    psspy.run(0, 31.1,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 33.1,out_txt_steps,plot_steps,0) 
    # Zf=Zs 0.12s duration
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[2*Rf,2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0,33.32,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 35.32,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[2*Rf,2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 35.44,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0,36.94,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[2*Rf,2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0,37.06,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0,38.06,out_txt_steps,plot_steps,0) 
    # Zf =Zs 0.220s duration 
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 38.18,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0,38.93,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 39.05,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 39.55,out_txt_steps,plot_steps,0) 
    
    #Zf= 2Zs 0.220s duration 
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 39.67,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 40.17,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[3.5*Rf,3.5*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 40.29,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0,40.49,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[3.5*Rf,3.5*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 40.61,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 40.81,out_txt_steps,plot_steps,0) 
    
    #Zf=3.5Zs 0.220 and 0.430s fault duration 
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 40.93,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 40.94,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0,41.37,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 45,out_txt_steps,plot_steps,0)                          

def MFRT_S4(Rf, Xf):
    psspy.run(1, 5.0,out_txt_steps,plot_steps,0) 
    
    # Bolted faults 0.12s duration 
    psspy.dist_bus_fault_3(1, Vbase,[0,1,DUMMYBUS,0,1],[0,-0.2E+10,0.0,0.0,0.0,0.0])
    psspy.run(1, 5.12,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 5.32,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(1, Vbase,[0,1,DUMMYBUS,0,1],[0,-0.2E+10,0.0,0.0,0.0,0.0])
    psspy.run(0, 5.44,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 5.64,out_txt_steps,plot_steps,0) 
    # Zf= 0.2Zs 0.12s duration
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[0.2*Rf,0.2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 5.76,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 6.26,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[0.2*Rf,0.2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 6.38,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 6.88,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[0.2*Rf,0.2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0,7,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 7.75,out_txt_steps,plot_steps,0) 
    # Zf=Zs 0.12s duration
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 7.87,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 7.88,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0,8,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 8.01,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 8.13,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 13.13,out_txt_steps,plot_steps,0) 
    # Zf =Zs 0.220s duration 
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 13.35,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 20.35,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 20.57,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 30.57,out_txt_steps,plot_steps,0) 
    
    #Zf= 2Zs 0.220s duration 
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[2*Rf,2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 30.79,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 31.79,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[2*Rf,2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 32.01,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 33.51,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[2*Rf,2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 33.73,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 35.73,out_txt_steps,plot_steps,0) 
    
    #Zf=3.5Zs 0.220 and 0.430s fault duration 
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[3.5*Rf,3.5*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 35.95,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 37.95,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[3.5*Rf,3.5*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 38.38,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 45,out_txt_steps,plot_steps,0)     
                 
def MFRT_S5(Rf, Xf):
    psspy.run(1, 5.0,out_txt_steps,plot_steps,0) 
    # psspy.dist_bus_fault(100,3, Vbase,[Rf,Xf])
    # Bolted faults 0.12s duration 
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[0.2*Rf,0.2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(1, 5.22,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 5.42,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[0.2*Rf,0.2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 5.64,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0,5.84,out_txt_steps,plot_steps,0) 
    # Zf= 0.2Zs 0.12s duration
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[0.2*Rf,0.2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 6.06,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 6.56,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 6.78,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 7.28,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 7.71,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 8.46,out_txt_steps,plot_steps,0) 
    # Zf=Zs 0.12s duration
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0,8.58,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 8.59,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(1, Vbase,[0,1,DUMMYBUS,0,1],[0,-0.2E+10,0.0,0.0,0.0,0.0])
    psspy.run(0, 8.71,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0,8.72,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(1, Vbase,[0,1,DUMMYBUS,0,1],[0,-0.2E+10,0.0,0.0,0.0,0.0])
    psspy.run(0, 8.84,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0,13.84,out_txt_steps,plot_steps,0) 
    # Zf =Zs 0.220s duration 
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[3.5*Rf,3.5*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 13.96,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 20.96,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[3.5*Rf,3.5*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 21.08,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 31.08,out_txt_steps,plot_steps,0) 
    
    #Zf= 2Zs 0.220s duration 
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[2*Rf,2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 31.2,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0,32.2,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[2*Rf,2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 32.32,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 33.82,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[2*Rf,2*Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 33.94,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 35.94,out_txt_steps,plot_steps,0) 
    
    #Zf=3.5Zs 0.220 and 0.430s fault duration 
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 36.16,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 38.16,out_txt_steps,plot_steps,0) 
    
    psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0,38.38,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 45,out_txt_steps,plot_steps,0)                                                           
 
def PFsetp(P_POC): 
    PFsetp_VAR_index = 15   
    Qflag_CON_index = 171
    # The SMAHYCF14 model has initially configured to Site voltage control, hence, the SMAHYCF14 model CON(j+204) must be set for PFsetp
    psspy.change_cctbusomod_con(POC_BUS, PPC_MODEL,1+Qflag_CON_index, 1074)     # change PwrRt.PwrRtCtrlMode => 1074
    psspy.change_cctbusomod_con(POC_BUS, PPC_MODEL,2+Qflag_CON_index, 1)        # change PwrRtCtrl.CtrlKi to => 1
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+18, P_POC*1000) # Change P Setpoint 
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+13, 0) # Change BESS W Setpoint. kick the BESS
    setp = 0.9701425 # cosphi=P/sqrt(P^2+Q^2) P=1 Q=0.395*P
    # setp = 0.98
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+PFsetp_VAR_index, 1) # PF Unity (0Mvar) set point
    psspy.run(0, 20.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+PFsetp_VAR_index, -setp) # Q step down 30% (derived by PF=COS(ATAN(Qpoi/Ppoi)))
    psspy.run(0, 50.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+PFsetp_VAR_index, 1) # Q step-up 30%
    psspy.run(0, 80.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+PFsetp_VAR_index, setp) # Q step down 30%
    psspy.run(0, 110.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+PFsetp_VAR_index, 1) # PF Unity (0Mvar) set point
    psspy.run(0, 140.0,out_txt_steps,plot_steps,0)                      

def PFsetp_DMAT(P_POC): 
    PFsetp_VAR_index = 15
    Qflag_CON_index = 171
    # The SMAHYCF14 model has initially configured to Site voltage control, hence, the SMAHYCF14 model CON(j+204) must be set for PFsetp
    psspy.change_cctbusomod_con(POC_BUS, PPC_MODEL,1+Qflag_CON_index, 1074)
    psspy.change_cctbusomod_con(POC_BUS, PPC_MODEL,2+Qflag_CON_index, 1)
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+18, P_POC*1000) # Change P Setpoint 
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+13, -17593) # Change P Setpoint 
    setp = 0.957826285 # cosphi=P/sqrt(P^2+Q^2) P=1 Q=0.3*P #0.957826285
    # setp = 0.98
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+PFsetp_VAR_index, 1) # PF Unity (0Mvar) set point
    psspy.run(0, 20.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+PFsetp_VAR_index, -setp) # Q step down 30% (derived by PF=COS(ATAN(Qpoi/Ppoi)))
    psspy.run(0, 40.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+PFsetp_VAR_index, setp) # Q step-up 30%
    psspy.run(0, 60.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+PFsetp_VAR_index, -setp) # Q step down 30%
    psspy.run(0, 80.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+PFsetp_VAR_index, 1) # PF Unity (0Mvar) set point
    psspy.run(0, 130.0,out_txt_steps,plot_steps,0)

def PFsetp_DMAT005(P_POC): 
    PFsetp_VAR_index = 15
    Qflag_CON_index = 171
    # The SMAHYCF14 model has initially configured to Site voltage control, hence, the SMAHYCF14 model CON(j+204) must be set for PFsetp
    psspy.change_cctbusomod_con(POC_BUS, PPC_MODEL,1+Qflag_CON_index, 1074)
    psspy.change_cctbusomod_con(POC_BUS, PPC_MODEL,2+Qflag_CON_index, 1)
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+18, P_POC*1000) # Change P Setpoint 
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+13, -17593) # Change P Setpoint 
    setp = 0.1644 # cosphi=P/sqrt(P^2+Q^2) P=0.05*MWrating Q = 0.3*MWrating
    # setp = 0.98
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+PFsetp_VAR_index, 1) # PF Unity (0Mvar) set point
    psspy.run(0, 20.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+PFsetp_VAR_index, -setp) # Q step down 30% (derived by PF=COS(ATAN(Qpoi/Ppoi)))
    psspy.run(0, 40.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+PFsetp_VAR_index, setp) # Q step-up 30%
    psspy.run(0, 60.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+PFsetp_VAR_index, -setp) # Q step down 30%
    psspy.run(0, 80.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+PFsetp_VAR_index, 1) # PF Unity (0Mvar) set point
    psspy.run(0, 130.0,out_txt_steps,plot_steps,0)

def TOV(Fixed_Shunt,T_flt):
    # print(Fixed_Shunt)
    psspy.run(0, 15.0,out_txt_steps,plot_steps,0)  
    psspy.shunt_data(DUMMYBUS,r"""1""",1,[0.0,Fixed_Shunt]) # connect a fixed 
    psspy.run(0,15+T_flt,out_txt_steps,plot_steps,0)  
    psspy.shunt_chng(DUMMYBUS,r"""1""",0,[_f,_f]) # switch off the connected fixed shunt 
    psspy.run(0, 30.0,out_txt_steps,plot_steps,0)  

def Qsetp_step(): 
    Qsetp_VAR_index = 14
    Qflag_CON_index = 171
    # The TSC model has initially configured to Site voltage control, hence, the TSC model ICON(M+24) must be set for Qsetp_step
    psspy.change_cctbusomod_con(POC_BUS, PPC_MODEL,1+Qflag_CON_index, 303)
    psspy.change_cctbusomod_con(POC_BUS, PPC_MODEL,2+Qflag_CON_index, 1)
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+13, 0) # Change BESS W Setpoint. kick the BESS
    Scale_P_KW = MWrating*1000 #65000
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+Qsetp_VAR_index, 0) # Q original set point
    psspy.run(0, 20.0,out_txt_steps,plot_steps,0) 
    # print(-0.1975*Scale_P_KW)
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+Qsetp_VAR_index, -0.1975*Scale_P_KW) # Q step down (39.5%/2) 
    psspy.run(0, 50.0,out_txt_steps,plot_steps,0) 
    # print(0.1975*Scale_P_KW)
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+Qsetp_VAR_index, 0*Scale_P_KW) # Q step-up (39.5%/2)
    psspy.run(0, 80.0,out_txt_steps,plot_steps,0) 
    # print(-0.1975*Scale_P_KW)
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+Qsetp_VAR_index, 0.1975*Scale_P_KW) # Q step down (39.5%/2)
    psspy.run(0, 110.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+Qsetp_VAR_index, 0) # Q original set point
    psspy.run(0, 140.0,out_txt_steps,plot_steps,0)                                         
                                         

def Qsetp_DMAT(): 
    Qsetp_VAR_index = 14
    Qflag_CON_index = 171
    # The TSC model has initially configured to Site voltage control, hence, the TSC model ICON(M+24) must be set for Qsetp_step
    psspy.change_cctbusomod_con(POC_BUS, PPC_MODEL,1+Qflag_CON_index, 303)
    psspy.change_cctbusomod_con(POC_BUS, PPC_MODEL,2+Qflag_CON_index, 1)
    Scale_P_KW = MWrating*1000
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+Qsetp_VAR_index, 0) # Q original set point
    psspy.run(0, 20.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+Qsetp_VAR_index, -0.30*Scale_P_KW) # Q step down 30%
    psspy.run(0, 40.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+Qsetp_VAR_index, 0.30*Scale_P_KW) # Q step-up 30%
    psspy.run(0, 60.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+Qsetp_VAR_index, -0.30*Scale_P_KW) # Q step down 30%
    psspy.run(0, 80.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+Qsetp_VAR_index, 0) # Q original set point
    psspy.run(0, 100.0,out_txt_steps,plot_steps,0)

def Vgrid_step(P_POC): #Verify function Vgrid_Step_config(del_V1 = 0.05)
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+18, P_POC*1000) # Change DirectSales Setpoint
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+13, 0) # Change BessW Setpoint
    psspy.run(0, 90.0,out_txt_steps,plot_steps,0) 

def Vgrid_step_Qmode(P_POC): #Verify function Vgrid_Step_config(del_V1 = 0.05)
    Qsetp_VAR_index = 14
    Qflag_CON_index = 171
    
    psspy.change_cctbusomod_con(POC_BUS, PPC_MODEL,1+Qflag_CON_index, 303)
    psspy.change_cctbusomod_con(POC_BUS, PPC_MODEL,2+Qflag_CON_index, 1)
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+18, P_POC*1000) # Change DirectSales Setpoint
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+13, 0) # Change BessW Setpoint
    psspy.run(0, 90.0,out_txt_steps,plot_steps,0)

def IqTest(P_POC): #Verify function Vgrid_Step_config(del_V1 = 0.05)
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+18, P_POC*1000) # Change DirectSales Setpoint
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+13, 0) # Change BessW Setpoint
    psspy.run(0, 40.0,out_txt_steps,plot_steps,0) 

def Vref_step(): 
    Vref_VAR_index = 16
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+13, 0) #kick the BESS
    print(PPC_VAR)
    ierr, rval_V1 = psspy.dsrval('VAR', PPC_VAR + Vref_VAR_index) # Get initial PPC V setpoint value
    print(rval_V1)
    ierr, rval_P_POC=psspy.brnmsc(POC_BUS,DUMMYBUS, '1', 'P')
    print(rval_P_POC)
    #change BESSWsetpoint
    # psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+13, 0)
    del_v1 = 0.05
    #del_v1=0.05 #5% step to be applied
    
    psspy.run(0, 1,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL, 1+Vref_VAR_index, rval_V1) # voltage step-up 5%
    print(rval_V1)
    print(rval_P_POC)

    psspy.run(0, 10,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL, 1+Vref_VAR_index, rval_V1+del_v1) # voltage step-up 5%
    print(psspy.chnval(34))
    print(rval_P_POC)

    psspy.run(0, 30.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL, 1+Vref_VAR_index, rval_V1) # voltage original set point
    print(psspy.chnval(34))
    print(rval_P_POC)

    psspy.run(0, 50.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL, 1+Vref_VAR_index, rval_V1-del_v1) # voltage step down 5%
    print(psspy.chnval(34))
    print(rval_P_POC)

    psspy.run(0, 70.0,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL, 1+Vref_VAR_index, rval_V1) # voltage original set point
    psspy.run(0, 90.0,out_txt_steps,plot_steps,0)  

def Vref_step_into_Qlimits(Test_Type):
    Vref_VAR_index = 16
    # The TSC model has initially configured to Site voltage control, hence, there is no need to set the ICON(M+24) for Vref_test
    ierr, rval_V1 = psspy.dsrval('VAR', PPC_VAR + Vref_VAR_index) # Get initial PPC V setpoint value
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+13, 0) #kick the BESS
    #del_v1=0.05 #5% step to be applied
    psspy.run(0, 1,out_txt_steps,plot_steps,0) 
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL, 1+Vref_VAR_index, rval_V1) # voltage step-up 5%
    psspy.run(0, 10,out_txt_steps,plot_steps,0)
    
    if Test_Type == "S52513_Vsetpoints_CAPlimit":
        psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL, 1+Vref_VAR_index, rval_V1-0.025) # voltage step down 2.5%
        psspy.run(0, 30.0,out_txt_steps,plot_steps,0) 
        psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL, 1+Vref_VAR_index, rval_V1+0.025) # voltage step up 5% to hit Qmax
        psspy.run(0, 50.0,out_txt_steps,plot_steps,0)
        psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL, 1+Vref_VAR_index, rval_V1-0.025) # voltage step down 2.5%
        psspy.run(0, 70.0,out_txt_steps,plot_steps,0)        
 
    elif Test_Type == "S52513_Vsetpoints_INDlimit":
        psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL, 1+Vref_VAR_index, rval_V1+0.025) # voltage step up 2.5%
        psspy.run(0, 30.0,out_txt_steps,plot_steps,0) 
        psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL, 1+Vref_VAR_index, rval_V1-0.025) # voltage step down 5% to hit Qmin
        psspy.run(0, 50.0,out_txt_steps,plot_steps,0)
        psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL, 1+Vref_VAR_index, rval_V1+0.025) # voltage step up 2.5%
        psspy.run(0, 70.0,out_txt_steps,plot_steps,0) 

def Force_tap(Trf_Tap): # not used at the moment
    if Trf_Tap != 0:
        print("Forcing Taps to %s" %Trf_Tap)
        psspy.three_wnd_winding_data_5(200,300,600,r"""1""",1,[_i,_i,_i,_i,_i,_i],[ Trf_Tap,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
        psspy.three_wnd_winding_data_5(200,800,1000,r"""1""",1,[_i,_i,_i,_i,_i,_i],[ Trf_Tap,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
        psspy.three_wnd_winding_data_5(200,1200,1400,r"""1""",1,[_i,_i,_i,_i,_i,_i],[ Trf_Tap,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
    
def Pref_step(P_POC,Test_Type):
    
    POC_max_kW = P_POC*1000 
    print('0',POC_max_kW)
    Pref_VAR_index = 18

    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+13, 0) #kick the BESS
    if Test_Type == "S52514_Pdispatch":
        psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index, POC_max_kW*(1.0)) # P Setpoint  
  
        psspy.run(0, 15,100,100,0)
        psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index, POC_max_kW*(0.5)) # P Setpoint  

        psspy.run(0, 345.0,100,100,0)
        psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index, POC_max_kW*(1.0)) # P Setpoint  

        psspy.run(0, 675.0,100,100,0)
    elif Test_Type == "S52514_Pstep":
        psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index, POC_max_kW*(1.0)) # P Setpoin
        psspy.run(0, 10,10,10,0)
        psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index, POC_max_kW*(0.5)) # P Setpoint 
 
        psspy.run(0, 15.0,10,10,0)
        psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index, POC_max_kW*(0.05)) # P Setpoint
  
        psspy.run(0, 25,10,10,0)
        psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index, POC_max_kW*(1.0)) # P Setpoint
 
        psspy.run(0, 35.0,10,10,0)
    else:
   
        psspy.run(0,10,out_txt_steps,plot_steps,0) 
        psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index, POC_max_kW*(0.5)) # Change P Setpoint

        psspy.run(0,20,out_txt_steps,plot_steps,0)  
        psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index, POC_max_kW*(0.05)) # Change P Setpoint

        psspy.run(0,30,out_txt_steps,plot_steps,0)  
        psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index, POC_max_kW*(1.0)) # Change P Setpoint

        psspy.run(0,40,out_txt_steps,plot_steps,0)             

def Freq_test(P_POC, Test_Type, extras = 0):
    psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+18, P_POC*1000) #kick the PPC

    if extras != 0:
        psspy.change_plmod_con(INV_SF,r"""1""",SF_INV_model,1, extras) # Set Initial Irradiance to 80% (i.e. not Max)
        psspy.run(0, 5.0,out_txt_steps,plot_steps,0) 

    if Test_Type in ["S5253_OFtrip","S5253_UFtrip"]: #3s step
        psspy.run(0, 752.5,out_txt_steps,plot_steps,0) 
    elif Test_Type in ["S52511_OF","S52511_UF"]: #3s step
        psspy.run(0, 55.0,out_txt_steps,plot_steps,0)  
    elif Test_Type in ["DMAT_OverFreq_3s","DMAT_UnderFreq_3s", "DMAT_UnderFreq_4Hz"]: #3s step
        psspy.run(0, 35.0,out_txt_steps,plot_steps,0)
    elif Test_Type in ["FDroopCheck"]: #3s step
        psspy.run(0, 125.0,out_txt_steps,plot_steps,0)
    elif Test_Type in ["S52511_small"]: #3s step
        psspy.run(0, 85.0,out_txt_steps,plot_steps,0)
    else: # "DMAT_OverFreq_4Hz"
        psspy.run(0, 75.0,out_txt_steps,plot_steps,0) 
                
def Vgrid_step_2(Test_Type):
    psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+18,26000) # Change Lim at Sales 
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+13, -13000) # Change P Setpoint 
    if Test_Type == "S5254_LVRT": 
        psspy.run(0, 45.0,out_txt_steps,plot_steps,0) 
    elif Test_Type == "S5254_HVRT": 
        psspy.run(0, 160.0,out_txt_steps,plot_steps,0)
    elif Test_Type == "S5254_CUO": 
        psspy.run(0, 7,out_txt_steps,plot_steps,0)
        psspy.change_cct2wtdmod_con(HV_BUS,MV_BUS,r"""1""",r"""OLTC1T""",1, 1000.0)
        psspy.change_cct2wtdmod_con(HV_BUS,MV_BUS,r"""1""",r"""OLTC1T""",3, 1000.0)
        psspy.run(0, 90.0,out_txt_steps,plot_steps,0)
    elif Test_Type == "S5255_HVRT": 
        psspy.run(0, 15.0,out_txt_steps,plot_steps,0)
    else:
        psspy.run(0, 55.0,out_txt_steps,plot_steps,0)

def Vgrid_ramp():
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+13, 0) #kick the BESS
    psspy.run(0, 85.0,out_txt_steps,plot_steps,0) 

def Vgrid_Extended_dip():
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+13, 0) #kick the BESS
    psspy.run(0, 40.0,out_txt_steps,plot_steps,0) 

def Vgrid_phase_angle(): 
    #bus configuration: POC (200): Dummy(100): Inf(8888)
    #want to put a new bus (9999) between Dummy(100) and Inf(8888)
    #want to put transformer between Dummy(100) and IDTRF(9999)
    bus_IDTRF = 9999
    # psspy.ltap(DUMMYBUS,INF_BUS,r"""1""", 0.0001,bus_IDTRF,r"""IDTRF""", _f) # inserts a new new dummy bus along on the line / 0.0001 is new bus total line length "From bus" in pu 
    # #new config:POC (200): Dummy(100): IDTRF(9999): Inf(8888)
    # psspy.purgbrn(DUMMYBUS,bus_IDTRF,r"""1""") # deletes the branch
    # #place new transformer between Dummy (100) and IDTRF (8888)
    # psspy.two_winding_data_3(DUMMYBUS,bus_IDTRF,r"""1""",[1,DUMMYBUS,1,0,0,0,33,0,DUMMYBUS,0,1,0,1,1,1],[0.0, 0.0001, 100.0, 1.0,0.0,0.0, 1.0,0.0,0.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0,0.0,0.0, 1.1, 0.9, 1.1, 0.9,0.0,0.0,0.0],r"""DUMMY_TX""")

    
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+13, 0) #kick the BESS
    psspy.run(0, 15.0,out_txt_steps,plot_steps,5) 
    psspy.two_winding_data_6(DUMMYBUS,bus_IDTRF,r"""1""",realari6 = 40)
    psspy.run(0, 20.0,out_txt_steps,plot_steps,5) 
    psspy.two_winding_data_6(DUMMYBUS,bus_IDTRF,r"""1""",realari6 = 0)
    psspy.run(0, 25.0,out_txt_steps,plot_steps,5) 
    psspy.two_winding_data_6(DUMMYBUS,bus_IDTRF,r"""1""",realari6 = -40)
    psspy.run(0, 30.0,out_txt_steps,plot_steps,5) 
    psspy.two_winding_data_6(DUMMYBUS,bus_IDTRF,r"""1""",realari6 = 0)
    psspy.run(0, 40.0,out_txt_steps,plot_steps,5) 
    
    #Grid Voltage Phase Angle +/- 60 degree Change
    psspy.two_winding_data_6(DUMMYBUS,bus_IDTRF,r"""1""",realari6 = 60)
    psspy.run(0, 45.0,out_txt_steps,plot_steps,5) 
    psspy.two_winding_data_6(DUMMYBUS,bus_IDTRF,r"""1""",realari6 = 0)
    psspy.run(0, 50.0,out_txt_steps,plot_steps,5) 
    psspy.two_winding_data_6(DUMMYBUS,bus_IDTRF,r"""1""",realari6 = -60)
    psspy.run(0, 55.0,out_txt_steps,plot_steps,5) 
    psspy.two_winding_data_6(DUMMYBUS,bus_IDTRF,r"""1""",realari6 = 0)
    psspy.run(0, 60.0,out_txt_steps,plot_steps,5)                                       
              
def SCR1_Pref_steps(P_POC):
    if P_POC>0:
        POC_max_kW = MWrating*1000 # PSSE TSC model takes pu value for Pdir_REF
    else:
        POC_max_kW = -MWrating*1000
    Pref_VAR_index = 18
    #Pref_VAR_index_BESS = 13
    # print(POC_max_kW)
    #psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index_BESS, POC_max_kW*(0.05))
    psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index, POC_max_kW*(0.05)) # Change P Setpoint                
    psspy.run(0,15,out_txt_steps,plot_steps,0) 
    #psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index_BESS, POC_max_kW*(0.2)) # Change P Setpoint
    psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index, POC_max_kW*(0.2))
    psspy.run(0,25,out_txt_steps,plot_steps,0)  
    #psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index_BESS, POC_max_kW*(0.4)) # Change P Setpoint
    psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index, POC_max_kW*(0.4)) # Change P Setpoint
    psspy.run(0,35,out_txt_steps,plot_steps,0)  
    #psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index_BESS, POC_max_kW*(0.6)) # Change P Setpoint
    psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index, POC_max_kW*(0.6)) # Change P Setpoint
    psspy.run(0,45,out_txt_steps,plot_steps,0) 
    #psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index_BESS, POC_max_kW*(0.8)) # Change P Setpoint
    psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index, POC_max_kW*(0.8)) # Change P Setpoint
    psspy.run(0,55,out_txt_steps,plot_steps,0)  
    #psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index_BESS, POC_max_kW*(1.0)) # Change P Setpoint
    psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+Pref_VAR_index, POC_max_kW*(1.0)) # Change P Setpoint
    psspy.run(0,65,out_txt_steps,plot_steps,0)     

def SCR1_POC_FRT_test(T_flt, Rf, Xf, X_R):  #Get Rpu and Xpu from UHHY_Source_Impedance_X&R_Calculator
    # Declare R and X for SCR = 1 conditions:
    R_MinX_R = 1.2163
    X_MinX_R = 3.6488
    R_MaxX_R = 0.2740
    X_MaxX_R = 3.8364
    # if fault_factor == 0:
    psspy.run(0, 15.0,out_txt_steps,plot_steps,0) 
    psspy.dist_3phase_bus_fault(DUMMYBUS,0,3, Vbase,[Rf, Xf])
    # psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0])
    psspy.run(0, 15+T_flt,out_txt_steps,plot_steps,0)
    psspy.dist_clear_fault(1)
    if X_R == 3:
        psspy.branch_chng_3(DUMMYBUS,INF_BUS,r"""1""",[_i,_i,_i,_i,_i,_i],[R_MinX_R,X_MinX_R,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s) #Change Grid Strength from SCR=3 to SCR=1, [Rpu, Xpu, ...]
    elif X_R == 14:
        psspy.branch_chng_3(DUMMYBUS,INF_BUS,r"""1""",[_i,_i,_i,_i,_i,_i],[R_MaxX_R,X_MaxX_R,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s) #Change Grid Strength from SCR=3 to SCR=1, 
    # psspy.run(1, 15.0,out_txt_steps,plot_steps,0)                                                   
    psspy.run(0, 30.0,out_txt_steps,plot_steps,0)              

def Site_specific_FRT_test(T_flt, Rf, Xf):
    psspy.run(0,15,out_txt_steps,plot_steps,0)       
    # if fault_factor == 0:
        # psspy.dist_3phase_bus_fault(DUMMYBUS,0,1,0.0,[0.0,-0.2E+10])
    # else:              
    psspy.dist_3phase_bus_fault(DUMMYBUS,0,3, Vbase,[Rf, Xf])    
    # psspy.dist_bus_fault_3(3, Vbase,[0,1,DUMMYBUS,0,1],[Rf,Xf,0.0,0.0,0.0,0.0]) # R and X in ohms
    # psspy.run(1, 10+0.35,out_txt_steps,plot_steps,0)

    psspy.run(0, 15+T_flt,out_txt_steps,plot_steps,0) 
    psspy.dist_clear_fault(1)
    psspy.run(0, 30,out_txt_steps,plot_steps,0) #Check this time after run

def Irradiance_Step(): #CHECK BUS NUMBER
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+13, 0)
    psspy.run(0,10,out_txt_steps,plot_steps,0) 
    psspy.change_plmod_con(INV_SF,r"""1""",SF_INV_model,1, 0.2) # Set Initial Irradiance to 80% (i.e. not Max)       
    psspy.run(0,20,out_txt_steps,plot_steps,0)  
    psspy.change_plmod_con(INV_SF,r"""1""",SF_INV_model,1, 0.8) # Change Inverter Irradiance 
    psspy.run(0,30,out_txt_steps,plot_steps,0)  
    psspy.change_plmod_con(INV_SF,r"""1""",SF_INV_model,1, 1) # Change Inverter Irradiance    
    psspy.run(0,40,out_txt_steps,plot_steps,0)          
    
def Irradiance_Step_226(): #if we need
    original_irradiance = 1
    #psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+13, BESSMWrating*1000)
    # psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+43, 1)
    psspy.change_plmod_con(INV_SF,r"""1""",SF_INV_model,1, original_irradiance) # Set Initial Irradiance to 80% (i.e. not Max)       
    psspy.run(0,15,out_txt_steps,plot_steps,0) 
    psspy.change_plmod_con(INV_SF,r"""1""",SF_INV_model,1, original_irradiance - 0.2) # Set Initial Irradiance to 80% (i.e. not Max)       
    psspy.run(0,20,out_txt_steps,plot_steps,0)  
    psspy.change_plmod_con(INV_SF,r"""1""",SF_INV_model,1, original_irradiance) # Change Inverter Irradiance 
    psspy.run(0,25,out_txt_steps,plot_steps,0)  
    
def Irradiance_Step_227(): #if we need
    #psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+13, BESSMWrating*1000)
    # psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+43, 1)
    psspy.change_plmod_con(INV_SF,r"""1""",SF_INV_model,1, 1)
    psspy.run(0,15,out_txt_steps,plot_steps,0) 
    psspy.change_plmod_con(INV_SF,r"""1""",SF_INV_model,1, 1.2) # Set Initial Irradiance to 80% (i.e. not Max)       
    psspy.run(0,60,out_txt_steps,plot_steps,0)  

def Irradiance_Step_228(): #if we need
    original_irradiance = 0.5
    #psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+13, BESSMWrating*1000)
    # psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+43, 1)
    psspy.change_plmod_con(INV_SF,r"""1""",SF_INV_model,1, original_irradiance) # Set Initial Irradiance to 80% (i.e. not Max)       
    psspy.run(0,15,out_txt_steps,plot_steps,0) 
    psspy.change_plmod_con(INV_SF,r"""1""",SF_INV_model,1, original_irradiance + 0.2) # Set Initial Irradiance to 80% (i.e. not Max)       
    psspy.run(0,20,out_txt_steps,plot_steps,0)  
    psspy.change_plmod_con(INV_SF,r"""1""",SF_INV_model,1, original_irradiance) # Change Inverter Irradiance 
    psspy.run(0,25,out_txt_steps,plot_steps,0)  
    
def Irradiance_Step_229(): #if we need
    #psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+13, BESSMWrating*1000)
    psspy.change_cctbusomod_var(POC_BUS, PPC_MODEL,1+43, 1)
    psspy.change_plmod_con(INV_SF,r"""1""",SF_INV_model,1, 0.5)
    psspy.run(0,50,out_txt_steps,plot_steps,0) 
    psspy.change_plmod_con(INV_SF,r"""1""",SF_INV_model,1, 0.3) # Set Initial Irradiance to 80% (i.e. not Max)       
    psspy.run(0,60,out_txt_steps,plot_steps,0)    

def HVRT_LVRT():
    psspy.run(0, 25,out_txt_steps,plot_steps,0)  
            
def main():
    # print('setupdir:',setupdir)
    # print('Logdir:',LogDir)
    # print('Outdir:',OutDir) 
    # print('Resultsdir:',ResultsDir) 

    # Get data from XLSX file
    test_workbook = openpyxl.load_workbook(excel_file, data_only=True)
    DMAT_baseCases = test_workbook["SMIB_studies_full"]
    
    cases_run = {}
    headings = []
    for ridx, row in enumerate(DMAT_baseCases.iter_rows()):
        if ridx == 0:
            headings = [elm.value for elm in row]
        else:
            if row[1].value == "Yes":
                cases_run[str(row[0].value)+"_"+row[2].value] = {}
                for eidx, elm in enumerate(headings):
                    cases_run[str(row[0].value)+"_"+row[2].value][headings[eidx]] = row[eidx] 

    for case in cases_run:
        # try:
        row = cases_run[case]
        # print(row)   
            
        Case_no = row['Case'].value
        Run_test = row['Run_test'].value
        Test_Type = row['Test_Type'].value
        P = float(row['ActiveP'].value)
        Q = float(row['ReactiveP'].value)
        P_POC = float(row['ActiveP'].value) * MWrating
        Q_POC = float(row['ReactiveP'].value) * MWrating
        SCR = float(row['SCR'].value) 
        SCP = float(row['SCR'].value) * MWrating
        X_R = float(row['X_R'].value)

        extras = row['Extras'].value
        if extras is not None and str(extras).isdigit():
            extras = float(extras)
        
        Vs_init = float(row['Vsource_i'].value)
        Vref = float(row['PPCVref'].value)
        Fixed_Shunt = float(row['Fixed_Shunt'].value)
        T_flt = float(row['Fault_duration'].value)

        Rf = float(row['Rf_Ohms'].value)
        Xf = float(row['L_H'].value) * (2*np.pi*50)
        
        GS_R = float(row['GS_R'].value)
        GS_X = float(row['GS_L'].value) * (2*np.pi*50)
        Trf_Tap = float(row['Trf_Tap'].value)
        FRT_dip = float(row['FRT_dip'].value)
        infname = str(row['Test_File_Names'].value)
        savefile = BaseCasesDir + str(row['SAV_File'].value)
        fault_factor = float(row['Fault Factor'].value)
        
        dyrfile = setupdir + dyrFile
        # print('Case_no, SCP, X_R, P_POC, Q_POC, Vref, infname, csv_file_DMAT')
        # print(Case_no, SCP, X_R, P_POC, Q_POC, Vref, infname)
        # print(Test_Type,infname,dyrfile,savefile,SCR,X_R,Case_no,FRT_dip)
        try:                
            init_sim(Test_Type,infname,dyrfile,savefile,SCR,X_R,Case_no,FRT_dip,T_flt)
            #psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1+13, 26000) #kick the BESS
            if Test_Type in ["3Ph_Balanced_Fault", "S5255_BalancedFault"]:
                #print("Trying Balanced Fault")
                Balanced_Fault(T_flt, Rf, Xf,P_POC)
                
            if Test_Type == "S52513_PFsetpoints":
                PFsetp(P_POC)
                
            if Test_Type == "S52513_Qsetpoints":    # 0.5*0.395=0.1975 steps
                Qsetp_step()

            #if Test_Type == "S52513_Qsetpoints": # use if you want full 0.395 steps
                #Qsetp_step2()

            if Test_Type == "S52513_Vgridsteps":
                Vgrid_step(P_POC)

            if Test_Type == "S52513_Vgridsteps_Qmode":
                Vgrid_step_Qmode(P_POC)
                
            if Test_Type == "S52513_Vsetpoints":
                Vref_step()    

            if Test_Type in ["S52513_Vsetpoints_CAPlimit", "S52513_Vsetpoints_INDlimit"]:
                Vref_step_into_Qlimits(Test_Type) 
                
            if Test_Type in ["DMAT_Psetpoints", "S52514_Pdispatch", "S52514_Pstep"]:
                Pref_step(P_POC, Test_Type)        
                
            if Test_Type in ["S5253_OFtrip","S5253_UFtrip","S52511_OF","S52511_UF","FDroopCheck","S52511_small"]:  
                Freq_test(P_POC, Test_Type) 
                
            if Test_Type in ["S5254_LVRT", "S5254_HVRT", "S5254_CUO","S5255_HVRT"]:
                Vgrid_step_2(Test_Type)
                
            if Test_Type == "GridPhaseAngle_step":
                Vgrid_phase_angle()
            
            if Test_Type == "POC_SCR_1_Pref_step":
                SCR1_Pref_steps(P_POC)
                
            if Test_Type == "POC_FRT_Test":
                SCR1_POC_FRT_test(T_flt, Rf, Xf, X_R)
                
            if Test_Type == "Site_Specific_FRT_Test":
                Site_specific_FRT_test(T_flt, Rf, Xf)  
                  
            #if Test_Type in ["Extended_Vdip_Recovery_0.8pu","Extended_Vdip_Recovery_0.5pu","Extended_Vdip_Recovery_0.1pu"]:
            #    Vgrid_Extended_dip()  
            ###################################DMATS##########################################

            if Test_Type == "MFRT_Protection_1":
                MFRT_Protection_1(Rf, Xf)
            if Test_Type == "MFRT_Protection_2":
                MFRT_Protection_2(Rf, Xf)   
            if Test_Type == "MFRT_S1":
                MFRT_S1(Rf, Xf)    
            if Test_Type == "MFRT_S2":
                MFRT_S2(Rf, Xf)     
            if Test_Type == "MFRT_S3":
                MFRT_S3(Rf, Xf)                 
            if Test_Type == "MFRT_S4":
                MFRT_S4(Rf, Xf)             
            if Test_Type == "MFRT_S5":
                MFRT_S5(Rf, Xf)

            if Test_Type in ["DMAT_TempOV_1.15pu","DMAT_TempOV_1.2pu"]:
                TOV(Fixed_Shunt,T_flt=T_flt)                

            if Test_Type == "DMAT_PFsetpoints":
                PFsetp_DMAT(P_POC)

            if Test_Type == "DMAT_PFsetpoints005":
                PFsetp_DMAT005(P_POC)

            if Test_Type == "DMAT_Qsetpoints":
                Qsetp_DMAT()
                
            if Test_Type == "DMAT_Vgridsteps":
                Vgrid_step(P_POC)
                
            if Test_Type == "DMAT_Vsetpoints":
                Vref_step()   
                
            if Test_Type in ["DMAT_OverFreq_3s","DMAT_UnderFreq_3s","DMAT_OverFreq_4Hz","DMAT_UnderFreq_4Hz"]:
                Freq_test(P_POC, Test_Type, extras)

            if Test_Type == "GridVchange_ramp":
                Vgrid_ramp()
               
            if Test_Type == "GridVchange_step":
                Vgrid_step_2(Test_Type)

            if Test_Type in ["Extended_Vdip_Recovery_0.8pu","Extended_Vdip_Recovery_0.5pu","Extended_Vdip_Recovery_0.1pu"]:
                Vgrid_Extended_dip()

            if Test_Type == "Flat_Run_test_300s":
                psspy.run(0, 300.0,out_txt_steps,plot_steps,0)
            
            if Test_Type == "Flat_Run_test_5s":
                psspy.run(0, 5.0,out_txt_steps,plot_steps,0)
                
            if Test_Type in ["Irradiance_step_226"]:
                Irradiance_Step_226()
                    
            if Test_Type in ["Irradiance_step_227"]:
                #print('TEST NOT RUNNING')
                #psspy.run(0, 300.0,out_txt_steps,plot_steps,0)
                Irradiance_Step_227()
                
            if Test_Type in ["Irradiance_step_228"]:
                #print('TEST NOT RUNNING')
                #psspy.run(0, 300.0,out_txt_steps,plot_steps,0)
                Irradiance_Step_228()
                
            if Test_Type in ["Irradiance_step_229"]:
                #print('TEST NOT RUNNING')
                #psspy.run(0, 300.0,out_txt_steps,plot_steps,0)
                Irradiance_Step_229()
                
                
            if (Test_Type == "LVRT") or (Test_Type == "HVRT"):
                HVRT_LVRT(P_POC)
            
            if (Test_Type == "S5255_IqTest"):
                IqTest(P_POC) 

            create_result_csv(ResultsDir=ResultsDir,infname=infname)
            
            psspy.pssehalt_2()
        except Exception as e:
            print("!!!!!!!!!!! ERROR : %s ---> Skipping Case %s \n\n" %(e,row))
            print(traceback.format_exc())

main()
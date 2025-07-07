#PSS(R)E release 34.7.00
# Case Create Script for Hybrid Plants
# Case Create used for: Upper Hunter Hybrid BESS + PV. APR 2024

import sys, os
import numpy as np
import pandas as pd
import csv
import openpyxl
import shutil
import math

# PSSE SETUP #################################################################
sys_path_PSSE = r"C:\Program Files (x86)\PTI\PSSE34\PSSPY27"
sys.path.append(sys_path_PSSE)
os_path_PSSE = r"C:\Program Files (x86)\PTI\PSSE34\PSSBIN"
os.environ['PATH'] += ';' + os_path_PSSE
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
setupdir = os.getcwd() + '\\'
LogDir = setupdir + "Log_Base_Cases_SF" + "\\"
SavDir = setupdir + "Base_Cases_SF" + "\\"
if not os.path.exists(SavDir):
    os.makedirs(SavDir)
if not os.path.exists(LogDir):
    os.makedirs(LogDir)

# READ DYR AND SLD FILE #######################################################

dyr_qconst = "MMHY_Hybrid_Updated.dyr"
sld = "MMHY_SLD.sld"
savefile = "MMHY_Base_Case_SF.sav"
Case_Create_filename = "Case_Create_SF.xlsx"
Case_Create_sheetname = "Case_Create"


# INPUT PLANT PARAMETERS (FOR SCALING PURPOSES) ###############################
ExportMWrating = 26 #POC MW ratings
#SFMWrating = 30         #Was 29
#BESSMWrating = 34.2 
# may not be used
#BESSVARrating = 13.509 #no other place is used
SFMWrating = 30 #change from 84
SFVARrating = 50.4
MWrating = 26
Sbase=100.0
Vbase=66.0
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

PPC_MODEL = "SMAHYCF23"


SF_INV_VAR = 1 + 0 #(Add six for PLB file)
#PPC_VAR = 301 + 0 #(Add six for PLB file) 
PPC_VAR = 301 + 0 #(Add six for PLB file) 

maxRuns = 1 # Changed from 9
### DMAT PSSE RUNTIME Parameters #############################################Stable
iterations = 600
sfactor = 0.2
tolerance = 0.0001
dT = 0.001
ffilter = 0.008


### SAVE OUTPUT CSV ##########################################################
def create_result_csv(ResultsDir,infname):
    print(SavDir + infname + ".out")
    # Saving the output in csv file format
    chnfobj=dyntools.CHNF(SavDir + infname + ".out")
    chnfobj.csvout(channels=[], csvfile='temp.csv')

    # Deleting the first row of csv file and saving the data in new csv file
    data = "".join(open("temp.csv").readlines()[1:])
    open(ResultsDir + "PSSE_" + infname + ".csv","wb").write(data)

    # Deleting the temporary csv file
    os.remove('temp.csv')

### CASE SETUP ###############################################################
# def Case_Setup(Case_no, SCP, SCR, X_R, P_POC, Q_POC, Vref, infname, csv_file_DMAT):
def Case_Setup(cases_run, Case_no, SCP, SCR, X_R, P_POC, Q_POC, Vref, infname):
    # PARAMETERS
    print('SCP',SCP)
    Znet= float(Vbase*Vbase/SCP)
    Rnet= float(Znet/((1+X_R**2))**(1.0/2.0))
    Xnet= float(Rnet*X_R)
    # Rnet1 = Rnet*SCR
    # Xnet1 = Xnet*SCR
    # Lnet1 = Xnet1/(2*np.pi*50)
    Znetpu= float(Znet/Zbase)
    Xnetpu= float(Xnet/Zbase)
    Rnetpu= float(Rnet/Zbase)

    # PSSE START
    psspy.psseinit()
    psspy.lines_per_page_one_device(1,60)
    It = infname
    psspy.lines_per_page_one_device(1,60)	
    psspy.progress_output(2, LogDir + It + "_Progress.txt",[0,0])
    psspy.case(savefile)
    
    #Change Grid Strength Impedance to match desired DMAT Scenario                
    print(' New simulation parameters ')   
    print(' Xnetpu, Rnetpu, Zbase, Vbase, SCP, X_R, P_POC, Q_POC')
    print(Xnetpu, Rnetpu, Zbase, Vbase, SCP, X_R, P_POC, Q_POC)

    # Solution - Load Flow Parameters
    psspy.solution_parameters_4([_i,100,_i,_i,10],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])                                
    ival = 9
    LF_ITER = 0
    
    # Loop runs full newton until it solves - used to initialise a good state
    while (ival>0):
        psspy.fnsl([1,0,0,1,1,0,99,0])
        # psspy.fnsl([1,0,1,1,1,4,99,1])                                    
        ival = psspy.solved()
        LF_ITER= LF_ITER+1
        if LF_ITER>20:
            print("LF did not converged 1.")
            quit()
            break                                            

    # Change SC Impedance slowly if the X is too large it will cause psse to get stuck
    if (Xnetpu>0.02396):
        LF_ITER = 0
        ij=0
        Xnetpu_Step=Xnetpu/20
        Rnetpu_Step=Rnetpu/20
        while (ij<=20): 
            psspy.branch_chng_3(DUMMYBUS,INF_BUS,r"""1""",[_i,_i,_i,_i,_i,_i],[Rnetpu_Step*ij,Xnetpu_Step*ij,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)
            print(' SC impedance changed')
            print('Xnetpu_Step*ij   ij')
            print(Xnetpu_Step*ij,ij)                                        
            ij=ij+1
            
            ival = 9
            LF_ITER = 0                                        
            while (ival>0):
                psspy.rsol([1,0,0,0,0,0,0,0,0,1],[ 500.0, 5.0])                                              
                ival = psspy.solved()
                LF_ITER= LF_ITER+1
                print('Load flow loop after step SC impedance change',LF_ITER)
                if LF_ITER>60:
                    print("LF did not converged 2 - After updating step SC impedance")
                    print(LF_ITER)
                    quit()
                    break                                                                                                  
    # Change SC Impedance in 1 step
    else:
        psspy.branch_chng_3(DUMMYBUS,INF_BUS,r"""1""",[_i,_i,_i,_i,_i,_i],[Rnetpu,Xnetpu,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)
        ival = 9
        LF_ITER = 0
        while (ival>0):
            psspy.rsol([1,0,0,0,0,0,0,0,0,1],[ 500.0, 5.0])                                          
            ival = psspy.solved()
            LF_ITER= LF_ITER+1
            print('Load flow loop after SC impedance change',LF_ITER)
            if LF_ITER>60:
                print("LF did not converge 3 - After updating SC Impedance")
                prnt(LF_ITER)
                quit()
                break
    print(' SC impedance changed successfully ')    

# Above code that changes SC impedance can be separated into its own function#
#############################################################################
    # Change Active and Reactive Power of Inverter Generator/s to regulate POC P and Q flows
    # The logic of PQ tuning has been modified by Vincent to tune P and Q at the same time.
    ij=0
    ival = 9
    maxIter = 300
    TuneDone = 0
    # P_REQ=0
    # Q_REQ=0
    Perror = 0.005
    Qerror = 0.02
    
    while (ij<maxIter and TuneDone==0):                               
        # CHECK ACTIVE POWER OUTPUT AT THE POC    
        ierr, rval_P=psspy.brnmsc(POC_BUS,DUMMYBUS, '1', 'P')
        ierr, rval_Txtap = psspy.xfrdat(HV_BUS,MV_BUS,'1','RATIO')
        Delt_P = P_POC - rval_P
        print('CHECK ACTIVE POWER OUTPUT AGAINST DESIRED BEFORE TUNING')
        print('Delt_P  rval_P  P_POC[n1]  TxTap')
        print(Delt_P, rval_P, P_POC, rval_Txtap)  
        
        if abs(Delt_P)>=Perror:
            # If error exceeds acceptable error, then correct Active power at the POC by updating inverter generator output
            #BESS ACTIVE POWER TUNING##########################################
            ierr, rval_P = psspy.brnmsc(POC_BUS,DUMMYBUS, '1', 'P')
            Delt_P = (P_POC - rval_P)
            print("The current active power at the POC",rval_P)
            print("Difference between desired power at the POC and current:",Delt_P)
            
            ierr, rval_P_BESS = psspy.brnmsc(INV_SF, MV_SF_BUS, '1', 'P')
            New_P = rval_P_BESS + Delt_P
            print('New_P',New_P)
            # Limit P dispatch of Inverter to machine Pmax/Pmin limits.
            if New_P>=SFMWrating: 
                New_P=SFMWrating
                print("SF reached upper limit of active P output. New_P:",New_P)
            if New_P<0:     #-SFMWrating
                New_P=0     #-SFMWrating
                print("SF reached lower limit of active P output. New_P:",New_P)

            psspy.machine_chng_2(INV_SF,r"""1""",[_i,_i,_i,_i,_i,_i],[ New_P, _f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])                             
            
            #psspy.fnsl([1,0,1,1,1,0,99,0])
            psspy.rsol([1,0,0,0,0,0,0,0,0,1],[0.0,0.0]) 
            #psspy.rsol([1,1,0,0,0,0,0,0,0,1],[0.0,0.0])            
            ival = psspy.solved()  
            print("ival:",ival,"if PSSE solved, ival = 0")
            #BESS ACTIVE POWER TUNING^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
        elif abs(Delt_P)<Perror:
            print('Active power tuning achieved at - Achieved %s'%rval_P)
            #SLACK BUS VOLTAGE TUNING##########################################
            ierr, rval_Q=psspy.brnmsc(POC_BUS,DUMMYBUS, '1', 'Q')
            ierr, rval_V=psspy.busdat(INF_BUS, 'PU')
            Delt_Q = rval_Q - Q_POC
            print ('rval_Q : Q_POC : Delt_Q : rval_V : ij : rval_P ')
            print (rval_Q, Q_POC, Delt_Q, rval_V, ij, rval_P,)
            
            #EXIT-LOOP CONDITION
            if abs(Delt_Q)<Qerror:                                                            
                print('rval_Q   rval_V New_U Q - Achieved ')
                print(rval_Q, rval_V)
                break 
                
            Delt_U=(Delt_Q/SCP)/2                                                                                
            New_U=rval_V+Delt_U
            # Limit slack bus voltage within its limits.
            if New_U>=1.55:
                New_U=1.54
                print("slack bus reached upper limit of voltage. New_U:",New_U)
            if New_U<0.6:
                New_U=0.61
                print("slack bus reached lower limit of voltage. New_U:",New_U)

            psspy.plant_data_4(INF_BUS,0,[_i,_i],[New_U,_f])

            #psspy.fnsl([1,0,1,1,1,0,99,0]) #changed this
            psspy.rsol([1,1,0,0,0,0,0,0,0,1],[0.0,0.0])            
            ival = psspy.solved() 
            print("ival:",ival,"if PSSE solved, ival = 0")            
            #SLACK BUS VOLTAGE TUNING^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
        # else: 
            # psspy.fnsl([1,0,1,1,1,0,99,0]) #changed this
                # psspy.rsol([1,0,0,0,0,0,0,0,0,1],[0.0,0.0])                                                                           
                # ival = psspy.solved() 
        # else:
            # print("Fatal error: no criterias were met to tune the plant")
            # quit()
            # break   
        if LF_ITER==maxIter:
            print("LF did not converge and reached maximum iterations 4.")
            quit()
            break
        
        print("Iteration %s"%ij)
        ij += 1
    
    #TUNING COMPLETE - SAVE RESULTS###########################################
    #psspy.fnsl([1,0,0,0,0,0,99,0]) #changed this
    psspy.rsol([1,0,0,0,0,0,0,0,0,1],[0.0, 2.0])
    #psspy.rsol([1,0,0,0,0,0,0,0,0,0],[0.0, 2.0])
    
    #saving the .sav file
    psspy.save(os.path.join(SavDir,str(It)+".sav"))
   
    ## Print Final POC Q flow and Grid V setpoint solution ##   
    ierr, rval_P=psspy.brnmsc(POC_BUS,DUMMYBUS, '1', 'P')
    ierr, rval_Q=psspy.brnmsc(POC_BUS,DUMMYBUS, '1', 'Q')
    ierr, rval_V=psspy.busdat(INF_BUS, 'PU') 
    ierr, rval_Txtap = psspy.xfrdat(200,300,'1','RATIO')
    print('Final solution rval_P:',rval_P,'rval_Q:',rval_Q,'rval_V:',rval_V,'Txtap:',rval_Txtap)
   
    # Update the dictionary of cases_run, which directly update the xlsx file
    cases_run[Case_no]['Vsource_i'].value = "%.4f"%rval_V
    cases_run[Case_no]['Trf_Tap'].value = "%.4f"%rval_Txtap

    # Update CSV File with new Grid Voltage Setpoint for PSCAD use ##
    # df = pd.read_csv(csv_file_DMAT, index_col='Case')
    # df.loc[float(Case_no),'Vsource_i'] = rval_V #round(rval_V,4)
    # df.loc[float(Case_no),'Trf_Tap'] = rval_Txtap #round(rval_Txtap,4)
    # df.to_csv(csv_file_DMAT)

##############################################################################################
# Run PSSE SAV Case Creation Script
##############################################################################################
# def add_channels(SF_INV_VAR,BESS_INV_VAR, PPC_VAR):
def add_channels(SF_INV_VAR, PPC_VAR):
    
    psspy.machine_array_channel([-1,2,INV_SF],r"""1""",r"""SF PELEC""")
    psspy.machine_array_channel([-1,3,INV_SF],r"""1""",r"""SF QELEC""")
    psspy.machine_array_channel([-1,9,INV_SF],r"""1""",r"""SF INV IdCmd""")
    psspy.machine_array_channel([-1,5,INV_SF],r"""1""",r"""SF INV Qcmd""")
    psspy.machine_array_channel([-1,8,INV_SF],r"""1""",r"""SF INV Pcmd""")    
    psspy.machine_array_channel([-1,12,INV_SF],r"""1""",r"""SF INV IqCmd""")
    # psspy.machine_array_channel([-1,2,INV_BESS],r"""1""",r"""BESS PELEC""")
    # psspy.machine_array_channel([-1,3,INV_BESS],r"""1""",r"""BESS QELEC""")
    # psspy.machine_array_channel([-1,9,INV_BESS],r"""1""",r"""BESS INV Id""")
    # psspy.machine_array_channel([-1,5,INV_BESS],r"""1""",r"""BESS Qcmd PPC""")
    # psspy.machine_array_channel([-1,8,INV_BESS],r"""1""",r"""BESS Pcmd PPC""")    
    # psspy.machine_array_channel([-1,12,INV_BESS],r"""1""",r"""BESS INV Iq""")   
    psspy.voltage_and_angle_channel([-1,-1,-1,POC_BUS],[r"""POC V""",r"""POC A"""])
    psspy.voltage_and_angle_channel([-1,-1,-1,INV_SF],[r"""SF INV V""",r"""SF INV A"""]) 
    # psspy.voltage_and_angle_channel([-1,-1,-1,INV_BESS],[r"""BESS INV V""",r"""BESS INV A"""]) 
    psspy.bus_frequency_channel([-1,POC_BUS],r"""POC FREQ""")
    psspy.voltage_and_angle_channel([-1,-1,-1,INF_BUS],[r"""GRID V""",r"""GRID A"""])
    psspy.voltage_and_angle_channel([-1,-1,-1,DUMMYBUS],[r"""DUMMY V""",r"""DUMMY A"""])

    psspy.branch_p_and_q_channel([-1,-1,-1,POC_BUS,DUMMYBUS],r"""1""",[r"""POC P""",r"""POC Q"""])
    psspy.branch_p_and_q_channel([-1,-1,-1,INV_SF,MV_SF_BUS],r"""1""",[r"""SF INV P""",r"""SF INV Q"""])
    
    psspy.var_channel([-1,PPC_VAR+16],r"""PPC_Vref""")
    psspy.var_channel([-1,PPC_VAR+60],r"""PPC_BatWSptMax""")
    psspy.var_channel([-1,PPC_VAR+61],r"""PPC_BatWSptMin""")
    psspy.var_channel([-1,PPC_VAR+62],r"""PPC_Ppv_cmd_inv""")
    psspy.var_channel([-1,PPC_VAR+63],r"""PPC_Qpv_cmd_inv""")
    # psspy.var_channel([-1,PPC_VAR+64],r"""PPC_Pbess_cmd_inv""")
    # psspy.var_channel([-1,PPC_VAR+65],r"""PPC_Qbess_cmd_inv""")


def Shunt_TOV_Cal(cases_run, infname, Case_no, SCP, dyrfile, target_U, T_flt = 0.9):
    psspy.psseinit()
    psspy.lines_per_page_one_device(1,60)
    It = infname
    psspy.lines_per_page_one_device(1,60)	
    psspy.progress_output(2, LogDir + infname + "_Progress.txt",[0,0])
    hide_output(0)
    psspy.case(SavDir +infname + ".sav")
    
    # Tap Line between POC and Grid Voltage Source to create DUMMY bus for fault test
    # SHUNTBUS = 777 
    # psspy.ltap(INF_BUS,DUMMYBUS,r"""1""", 0.5,SHUNTBUS,r"""SHUNT""", Vbase) #ltap taps the line
    
    # psspy.fnsl([1,0,0,1,1,0,99,0])
    # psspy.fnsl([1,0,0,1,1,0,99,0])
    # psspy.fnsl([1,0,0,1,1,0,99,0])
    psspy.fnsl([0,0,0,0,0,0,0,0])
    psspy.fnsl([0,0,0,0,0,0,0,0])
    psspy.fnsl([0,0,0,0,0,0,0,0])
    psspy.rsol([1,0,0,0,0,0,0,0,0,1],[0.0, 2.0])
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
    for dll in dll_files:
        psspy.addmodellibrary(dll)
    
    psspy.dynamics_solution_param_2([iterations,_i,_i,_i,_i,_i,_i,_i],[sfactor,tolerance,dT,ffilter,_f,_f,_f,_f]) #IN ORDER: iterations, acceleration factor, tolerance, timestep, frequency filter.    	
    SF_INV_VAR = 1 + 0 #(Add six for PLB file) #no SF
    #BESS_INV_VAR = 301 + 0 #(Add six for PLB file) # no SF
    PPC_VAR = 301 + 0 #(Add six for PLB file)
    # add_channels(SF_INV_VAR, BESS_INV_VAR, PPC_VAR)
    
    # PSSE OUT File Setup
    psspy.strt(0,SavDir + infname + ".out") # obsolete function from v30.2.0
    ierr, Bus_V=psspy.busdat(POC_BUS, 'PU') # 4001                                                  
    dif_U=float(target_U) - float(Bus_V)                                   
    Yf = SCP * dif_U
    psspy.run(1, 10.0,5000,10,0)
    
    # read PPC Vref from the PSCAD model
    Vref_VAR_index = 16
    ierr, PPC_Vref = psspy.dsrval('VAR', PPC_VAR + Vref_VAR_index)
    print(PPC_Vref)
    
    dif_Up=abs(100.0*dif_U/target_U)

    count = 0   
    
    iters = 0
    dt_test=0.0        
    while (dif_Up>0.5 and iters < 100):
        psspy.shunt_data(DUMMYBUS,r"""1""",1,[0.0,Yf]) # connect a fixed shunt impedance to invoke Overvoltage at the PoC
        
        dt_test=dt_test+0.01    
        print('dt_test = ', dt_test)
        psspy.run(1,10+dt_test,5000,10,0)   
        
        ierr, Bus_V=psspy.busdat(POC_BUS, 'PU') # 4001
        
        ierr, rval_Q_POC=psspy.brnmsc(POC_BUS,DUMMYBUS, '1', 'Q') 
        
        ierr, rval_Q_SF=psspy.brnmsc(INV_SF, MV_SF_BUS, '1', 'Q')  
        # ierr, rval_Q_BESS=psspy.brnmsc(INV_BESS, MV_BESS_BUS, '1', 'Q')
        dif_Up=abs(100.0*dif_U/target_U)                                        
        dif_U=float(target_U) - float(Bus_V) 
        dYf = SCP * dif_U
        Yf=Yf+dYf
        print('dif_Up:',dif_Up,' dif_U:',dif_U,' dYf:',dYf)
        # print(' II target_U,  Bus_V, Yf ,rval_Q_POC, rval_Q_SF, rval_Q_BESS')
        # print(target_U, Bus_V, Yf, rval_Q_POC, rval_Q_SF, rval_Q_BESS)
        print(' II target_U,  Bus_V, Yf ,rval_Q_POC, rval_Q_SF')
        print(target_U, Bus_V, Yf, rval_Q_POC, rval_Q_SF)
        print('iters',iters)
        
        iters += 1
        
    psspy.run(1,10+dt_test+T_flt,5000,10,0)
    # print(' --Shunt(%s pu) >>> target_U,  Bus_V, Yf ,rval_Q_POC, rval_Q_SF, rval_Q_BESS'%target_U)
    # print(target_U, Bus_V, Yf, rval_Q_POC, rval_Q_SF, rval_Q_BESS)
    print(' --Shunt(%s pu) >>> target_U,  Bus_V, Yf ,rval_Q_POC, rval_Q_SF'%target_U)
    print(target_U, Bus_V, Yf, rval_Q_POC, rval_Q_SF)
    return Yf, PPC_Vref

def runFlat(cases_run, infname, Case_no, SCP, dyrfile, P_POC, Q_POC,counter):
    psspy.psseinit()
    psspy.lines_per_page_one_device(1,60)
    It = infname
    psspy.lines_per_page_one_device(1,60)	
    psspy.progress_output(2, LogDir + infname + "_Progress.txt",[0,0])
    hide_output(0)
    psspy.case(SavDir +infname + ".sav")
    
    # psspy.fnsl([1,0,0,1,1,0,99,0])
    # psspy.fnsl([1,0,0,1,1,0,99,0])
    # psspy.fnsl([1,0,0,1,1,0,99,0])
    psspy.fnsl([0,0,0,0,0,0,0,0])
    psspy.fnsl([0,0,0,0,0,0,0,0])
    psspy.fnsl([0,0,0,0,0,0,0,0])
    
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
    for dll in dll_files:
        psspy.addmodellibrary(dll)
    
    psspy.dynamics_solution_param_2([iterations,_i,_i,_i,_i,_i,_i,_i],[sfactor,tolerance,dT,ffilter,_f,_f,_f,_f]) #IN ORDER: iterations, acceleration factor, tolerance, timestep, frequency filter.	
    SF_INV_VAR = 1 + 0 #(Add six for PLB file)
    # BESS_INV_VAR = 301 + 0 #(Add six for PLB file)
    PPC_VAR = 301+ 0 #(Add six for PLB file)
    if counter == 0:
        # add_channels(SF_INV_VAR, BESS_INV_VAR, PPC_VAR)
        add_channels(SF_INV_VAR, PPC_VAR)

    # PSSE OUT File Setup
    psspy.strt(0,SavDir + infname + ".out")                             
    ierr, Bus_V=psspy.busdat(POC_BUS, 'PU') # 4001     
    psspy.run(0, 5,1000,10,0)  

    # Update setpoint to match solution - same BESS and Plant in BESS only case
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1 + 18, P_POC*1000) # Plant Active Power
    psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1 + 13, P_POC*1000) # BESS W Setpoint
    #psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1 + 14,  Q_POC*1000) # Plant Reactive Power
    print('Q_POC',Q_POC)
    psspy.run(0, 25.0,1000,10,0)
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1 + 13, P_POC*1000) # BESS W Setpoint
    # psspy.run(0, 50.0,1000,10,0)
    # psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1 + 13, P_POC*950) # BESS W Setpoint
    # psspy.run(0, 75.0,1000,10,0)
    
    # read PPC Vref from the PSCAD model
    Vref_VAR_index = 16
    ierr, PPC_Vref = psspy.dsrval('VAR', PPC_VAR + Vref_VAR_index)
    ierr, PPC_Vref2 = psspy.chnval(30)
    # ierr, BESS_P = psspy.chnval(7)
    # ierr, BESS_Q = psspy.chnval(8)
    ierr, SF_P = psspy.chnval(1)
    ierr, SF_Q = psspy.chnval(2)
    ierr, GRID_V = psspy.chnval(20)
    ierr, POC_Pflat = psspy.chnval(24)
    ierr, POC_Qflat = psspy.chnval(25)
    print('PPC_Vref:',PPC_Vref)
    print('PPC_Vref2:',PPC_Vref2)
    return PPC_Vref, PPC_Vref2,POC_Pflat,POC_Qflat, SF_P, SF_Q, GRID_V #,BESS_P, BESS_Q 

def main():
    # Directory of CSV file with PSCAD Cases to be run 
    xlsx_file_DMAT = setupdir + Case_Create_filename
    
    # Get data from XLSX file
    test_workbook = openpyxl.load_workbook(xlsx_file_DMAT, data_only=True)
    DMAT_baseCases = test_workbook[Case_Create_sheetname]
    
    cases_run = {}
    headings = []
    for ridx, row in enumerate(DMAT_baseCases.iter_rows()):
        if ridx == 0:
            headings = [elm.value for elm in row]
            print(headings)
        else: 
            if row[1].value == "Yes":
                cases_run[row[0].value] = {}
                for eidx, elm in enumerate(headings):
                    cases_run[row[0].value][headings[eidx]] = row[eidx] 
    
    if cases_run == {}:
        print('Error: you must select a case to be created in the excel')
    
    counter=0
    for case in cases_run:
        row = cases_run[case]
        print(row)   
            
        Case_no = row['Case'].value
        Run_test = row['Run_test'].value
        Test_Type = row['Test_Type'].value
        P_POC = float(row['ActiveP'].value) * ExportMWrating
        Q_POC = float(row['ReactiveP'].value) * ExportMWrating
        SCR = float(row['SCR'].value) 
        SCP = float(row['SCR'].value) * MWrating
        X_R = float(row['X_R'].value)
        Vref = float(row['PPC_Vref'].value)
        Trf_Tap = float(row['Trf_Tap'].value)
        infname = str(row['Test_File_Names'].value)
        
        print('Case_no, SCP, X_R, P_POC, Q_POC, Vref, infname, csv_file_DMAT')
        print(Case_no, SCP, X_R, P_POC, Q_POC, Vref, infname)
        Case_Setup(cases_run, Case_no, SCP, SCR, X_R, P_POC, Q_POC, Vref, infname)
        dyrfile = setupdir + dyr_qconst
        
        # PPC_Vref, PPC_Vref2,POC_Pflat,POC_Qflat, BESS_P, BESS_Q, SF_P, SF_Q, GRID_V = runFlat(cases_run, infname, Case_no, SCP, dyrfile, P_POC, Q_POC,counter)
        PPC_Vref, PPC_Vref2, POC_Pflat, POC_Qflat, SF_P, SF_Q, GRID_V = runFlat(cases_run, infname, Case_no, SCP, dyrfile, P_POC, Q_POC,counter)
        counter=1
        cases_run[Case_no]['PPC_Vref'].value = "%.4f"%PPC_Vref
        # cases_run[Case_no]['SF_P'].value = "%.4f"%SF_P
        # cases_run[Case_no]['SF_Q'].value = "%.4f"%SF_Q
        # cases_run[Case_no]['BESS_P'].value = "%.4f"%BESS_P
        # cases_run[Case_no]['BESS_Q'].value = "%.4f"%BESS_Q
        # cases_run[Case_no]['POC_Pflat'].value = "%.4f"%POC_Pflat
        # cases_run[Case_no]['POC_Qflat'].value = "%.4f"%POC_Qflat
        #run shunts
        if cases_run[Case_no]['Shunt(1.15pu)'].value == "Yes":
            Yf, PPC_Vref = Shunt_TOV_Cal(cases_run, infname, Case_no, SCP, dyrfile, target_U=1.15, T_flt=0.9)
            cases_run[Case_no]['Shunt(1.15pu)'].value = "%.4f"%Yf
            cases_run[Case_no]['PPC_Vref'].value = "%.4f"%PPC_Vref
        if cases_run[Case_no]['Shunt(1.2pu)'].value == "Yes":
            Yf, PPC_Vref = Shunt_TOV_Cal(cases_run, infname, Case_no, SCP, dyrfile, target_U=1.2, T_flt=0.1)
            cases_run[Case_no]['Shunt(1.2pu)'].value = "%.4f"%Yf
            cases_run[Case_no]['PPC_Vref'].value = "%.4f"%PPC_Vref
            
        test_workbook.save(xlsx_file_DMAT)
        create_result_csv(ResultsDir=SavDir,infname=infname)

if (__name__ == "__main__"):
    main()
    psspy.stop()

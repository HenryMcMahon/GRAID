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
LogDir = setupdir + "Log_Base_Cases_Hybrid" + "\\"
SavDir = setupdir + "Base_Cases_Hybrid" + "\\"
if not os.path.exists(SavDir):
    os.makedirs(SavDir)
if not os.path.exists(LogDir):
    os.makedirs(LogDir)

# READ DYR AND SLD FILE #######################################################

dyr_qconst = "MMHY_Hybrid.dyr"
sld = "MMHY_SLD.sld"
savefile1pu = "MMHY_Base_Case_Hybrid_1pu.sav"
savefile = "MMHY_Base_Case_Hybrid.sav"
Case_Create_filename = "Case_Create_Hybrid.xlsx"
Case_Create_sheetname = "Case_Create"


# INPUT PLANT PARAMETERS (FOR SCALING PURPOSES) ###############################
BESSMWrating = 34.2 
SFMWrating = 30 
MWrating = 26 #POC MW ratings

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
BESS_INV_VAR = 301 + 0 #(Add six for PLB file)
PPC_VAR = 627 + 0 #(Add six for PLB file) #625

maxRuns = 10 # Changed from 9
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
def Case_Setup(cases_run, Case_no, SCP, SCR, X_R, P_POC, Q_POC, Vref, infname, runNum,BESSUpdateP, Target_POC_V,deratingtemp):
    print('we are in case setup, bessupdatep is:',BESSUpdateP)
    # PARAMETERS
    Znet= float(Vbase*Vbase/SCP)
    Rnet= float(Znet/((1+X_R**2))**(1.0/2.0))
    Xnet= float(Rnet*X_R)
    Rnet1 = Rnet*SCR
    Xnet1 = Xnet*SCR
    Lnet1 = Xnet1/(2*np.pi*50)
    Znetpu= float(Znet/Zbase)
    Xnetpu= float(Xnet/Zbase)
    Rnetpu= float(Rnet/Zbase)

    # PSSE START
    psspy.psseinit()
    psspy.lines_per_page_one_device(1,60)
    It = infname
    psspy.lines_per_page_one_device(1,60)	
    psspy.progress_output(2, LogDir + It + "_Progress.txt",[0,0])
    if runNum==0:
        psspy.case(savefile)
        print(savefile,' loaded into PSSE run')
     
    elif runNum==1:
        psspy.case(SavDir + infname + '.sav')
        print(SavDir,infname,'.sav',' loaded into PSSE run')
    
    #Change Grid Strength Impedance to match desired DMAT Scenario                
    print(' Start of Case Setup simulation parameters ', SCR,X_R,P_POC,Q_POC,BESSUpdateP)   
    print(' Xnetpu, Rnetpu, Zbase, Vbase, SCP, X_R, P_POC, Q_POC')
    print(Xnetpu, Rnetpu, Zbase, Vbase, SCP, X_R, P_POC, Q_POC)

    # Solution - Load Flow Parameters
    psspy.solution_parameters_4([_i,100,_i,_i,10],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])                                
    ival = 9
    LF_ITER = 0

    # Update Gen Target_POC_V
    psspy.plant_chng_4(INV_SF,0,[_i,_i],[ Target_POC_V,_f])
    psspy.plant_chng_4(INV_BESS,0,[_i,_i],[ Target_POC_V,_f])

    #Apply a derating scale. if we do this again, make sure to do the calc. if derating !=0, then calc ratings.
    if deratingtemp == 25:
        print('25 degrees is nothing silly')
    elif deratingtemp == 50:
        print('this is something huh')
        # psspy.machine_chng_2(INV_SF,r"""1""",[_i,_i,_i,_i,_i,_i],[ _f,_f, 79.520, -79.520, 79.520,_f, 79.520,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
        # psspy.machine_chng_2(INV_BESS,r"""1""",[_i,_i,_i,_i,_i,_i],[_f,_f, 35.4805,-35.4805, 35.4805,-35.4805, 35.4805,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
    psspy.fdns([1,0,0,1,1,0,99,0])
    psspy.fdns([1,0,0,1,1,0,99,0])

    # Loop runs full newton until it solves - used to initialise a good state
    while (ival>0):
                # psspy.fnsl([1,0,1,1,1,4,99,1])     
        psspy.rsol([1,1,0,0,0,0,0,0,0,1],[ 500.0, 5.0]) 
        ival = psspy.solved()
        LF_ITER= LF_ITER+1
        if LF_ITER>20:
            print("LF did not converged 1.")
            quit()
            break                                            
    
    if runNum==0:
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
                    psspy.rsol([1,0,0,0,0,0,0,0,0,1],[ 200.0, 0.4])                                              
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
                psspy.rsol([1,0,0,0,0,0,0,0,0,1],[ 200.0, 0.4])                                           
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
    # P_REQ=0
    # Q_REQ=0
    Perror = 0.01
    Qerror = 0.02
    #update the BESS output to be the headroom
    psspy.machine_chng_2(INV_BESS,r"""1""",[_i,_i,_i,_i,_i,_i],[ BESSUpdateP, _f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
    psspy.rsol([1,1,0,0,0,0,0,0,0,1],[0.0,0.0])
    # ierr, test_BESS_P = psspy.brnmsc(INV_BESS,MV_BESS_BUS,'1','P')
    # print('BBBBBBBBBBBBBBBBBBBBB',test_BESS_P)
    while (ij<maxIter):                               
        # CHECK ACTIVE POWER OUTPUT AT THE POC    
        ierr, rval_P=psspy.brnmsc(POC_BUS,DUMMYBUS, '1', 'P')
        ierr, rval_Txtap = psspy.xfrdat(HV_BUS,MV_BUS,'1','RATIO')
        Delt_P = P_POC - rval_P
        print('CHECK ACTIVE POWER OUTPUT AGAINST DESIRED BEFORE TUNING')
        print('Delt_P  rval_P  P_POC[n1]  TxTap')
        print(Delt_P, rval_P, P_POC, rval_Txtap)  
       
        if abs(Delt_P)>=Perror:
            # If error exceeds acceptable error, then correct Active power at the POC by updating inverter generator output
            #SF ACTIVE POWER TUNING##########################################
            ierr, rval_P = psspy.brnmsc(POC_BUS,DUMMYBUS, '1', 'P')
            Delt_P = (P_POC - rval_P)
            print("The current active power at the POC",rval_P)
            print("Difference between desired power at the POC and current:",Delt_P)
            
            ierr, rval_P_SF = psspy.brnmsc(INV_SF, MV_SF_BUS, '1', 'P')
            ierr, rval_Q_SF = psspy.brnmsc(INV_SF, MV_SF_BUS, '1', 'Q')
            #only want to increase the P slightly
            
            New_P = rval_P_SF + Delt_P*0.999
            print('New_P:',New_P)
            # Limit P dispatch of Inverter to machine Pmax/Pmin limits.
            if math.sqrt(pow(New_P,2)+pow(rval_Q_SF,2))>=SFMWrating: 
                New_P=math.sqrt(pow(SFMWrating*0.995,2)-pow(rval_Q_SF,2))
                print("SF reached upper limit of active P output. New_P:",New_P)
                #we have to therefore drop the BESS power a little bit if its charging
                ierr, rval_P_BESS = psspy.brnmsc(INV_BESS,MV_BESS_BUS,'1','P')
                if rval_P_BESS<0:
                    New_P_BESS = rval_P_BESS + abs(0.02+rval_P_SF-New_P)
                    psspy.machine_chng_2(INV_BESS,r"""1""",[_i,_i,_i,_i,_i,_i],[ New_P_BESS, _f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    print('BESS charging has slowed to:',New_P_BESS,'from:',rval_P_BESS)
            if New_P<0:
                New_P=0
                print("ERROR, SF P output is trying to go negative. New_P:",New_P)
            psspy.machine_chng_2(INV_SF,r"""1""",[_i,_i,_i,_i,_i,_i],[29.999, _f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
            psspy.rsol([1,1,0,0,0,0,0,0,0,1],[0.0,0.0]) 
                       
            ival = psspy.solved()  
            print("ival:",ival,"if PSSE solved, ival = 0")

            #SF ACTIVE POWER TUNING^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
        elif abs(Delt_P)<Perror:
            print('Active power tuning achieved at - Achieved %s'%rval_P)
            #SLACK BUS VOLTAGE TUNING##########################################
            ierr, rval_Q=psspy.brnmsc(POC_BUS,DUMMYBUS, '1', 'Q')
            ierr, rval_V=psspy.busdat(POC_BUS, 'PU')
            ierr, grid_rval_V=psspy.busdat(INF_BUS, 'PU')
            ierr, rval_P_SF = psspy.brnmsc(INV_SF, MV_SF_BUS, '1', 'P')
            ierr, rval_P_BESS = psspy.brnmsc(INV_BESS,MV_BESS_BUS,'1','P')
            Delt_Q = rval_Q - Q_POC
            print ('rval_Q',rval_Q)
            print ('Q_POC',Q_POC)
            print ('Delt_Q',Delt_Q)
            print ('rval_V',rval_V)
            print ('grid_rval_V',grid_rval_V)
            print ('ij',ij)
            print ('rval_P',rval_P)
            print ('P_SF',rval_P_SF)
            print ('P_BESS',rval_P_BESS)
            
            #EXIT-LOOP CONDITION
            if abs(Delt_Q)<Qerror:                                                            
                print('rval_Q  , grid_rval_V ,New_U ,Q - Achieved ')
                print(rval_Q, grid_rval_V, rval_V, Delt_Q)
                break 
                
            Delt_U=(Delt_Q/SCP)/2                                                                                
            New_U=grid_rval_V+Delt_U
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
 
        if LF_ITER==maxIter:
            print("LF did not converge and reached maximum iterations 4.")
            quit()
            break
        
        print("Iteration %s"%ij)
        ij += 1
    
    #TUNING COMPLETE - SAVE RESULTS###########################################
    print('AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA')
    psspy.rsol([1,0,0,0,0,0,0,0,0,1],[ 0.0, 0.0]) 
    psspy.save(os.path.join(SavDir,str(It)+".sav")) 
    ierr, Case_POC_P=psspy.brnmsc(POC_BUS,DUMMYBUS, '1', 'P')
    ierr, Case_POC_Q=psspy.brnmsc(POC_BUS,DUMMYBUS, '1', 'Q')
    ierr, Case_POC_V=psspy.busdat(POC_BUS, 'PU') 
    ierr, Case_GRID_V=psspy.busdat(INF_BUS, 'PU') 
    ierr, Case_Txtap = psspy.xfrdat(200,300,'1','RATIO')
    ierr, Case_SF_P=psspy.brnmsc(INV_SF,MV_SF_BUS,'1','P')
    ierr, Case_SF_Q=psspy.brnmsc(INV_SF,MV_SF_BUS,'1','Q')
    ierr, Case_BESS_P=psspy.brnmsc(INV_BESS,MV_BESS_BUS,'1','P')
    ierr, Case_BESS_Q=psspy.brnmsc(INV_BESS,MV_BESS_BUS,'1','Q')
    
    cases_run[Case_no]['Vsource_i'].value = "%.4f"%Case_GRID_V
    cases_run[Case_no]['Trf_Tap'].value = "%.4f"%Case_Txtap
    cases_run[Case_no]['PPC_Vref'].value = "%.4f"%Case_POC_V
    cases_run[Case_no]['SF_P'].value = "%.4f"%Case_SF_P #"%.4f"%(SF_P*100)
    cases_run[Case_no]['SF_Q'].value = "%.4f"%Case_SF_Q #"%.4f"%(SF_Q*100)
    cases_run[Case_no]['BESS_P'].value = "%.4f"%Case_BESS_P #"%.4f"%(BESS_P*100)
    cases_run[Case_no]['BESS_Q'].value = "%.4f"%Case_BESS_Q #"%.4f"%(BESS_Q*100)
    cases_run[Case_no]['POC_P'].value = "%.4f"%Case_POC_P #"%.4f"%POC_Pflat
    cases_run[Case_no]['POC_Q'].value = "%.4f"%Case_POC_Q #"%.4f"%POC_Qflat
    
    return Case_SF_P, Case_SF_Q, Case_BESS_P, Case_BESS_Q, Case_POC_P, Case_POC_Q, Case_POC_V
    

##############################################################################################
# Run PSSE SAV Case Creation Script
##############################################################################################
def add_channels(SF_INV_VAR,BESS_INV_VAR, PPC_VAR):
    
    psspy.machine_array_channel([-1,2,INV_SF],r"""1""",r"""SF PELEC""") #1
    psspy.machine_array_channel([-1,3,INV_SF],r"""1""",r"""SF QELEC""")
    psspy.machine_array_channel([-1,9,INV_SF],r"""1""",r"""SF INV IdCmd""")
    psspy.machine_array_channel([-1,5,INV_SF],r"""1""",r"""SF INV Qcmd""")
    psspy.machine_array_channel([-1,8,INV_SF],r"""1""",r"""SF INV Pcmd""")    
    psspy.machine_array_channel([-1,12,INV_SF],r"""1""",r"""SF INV IqCmd""")
    psspy.machine_array_channel([-1,2,INV_BESS],r"""1""",r"""BESS PELEC""")
    psspy.machine_array_channel([-1,3,INV_BESS],r"""1""",r"""BESS QELEC""")
    psspy.machine_array_channel([-1,9,INV_BESS],r"""1""",r"""BESS INV Id""")
    psspy.machine_array_channel([-1,5,INV_BESS],r"""1""",r"""BESS Qcmd PPC""")
    psspy.machine_array_channel([-1,8,INV_BESS],r"""1""",r"""BESS Pcmd PPC""")    
    psspy.machine_array_channel([-1,12,INV_BESS],r"""1""",r"""BESS INV Iq""")   
    psspy.voltage_and_angle_channel([-1,-1,-1,POC_BUS],[r"""POC V""",r"""POC A"""])
    psspy.voltage_and_angle_channel([-1,-1,-1,INV_SF],[r"""SF INV V""",r"""SF INV A"""]) 
    psspy.voltage_and_angle_channel([-1,-1,-1,INV_BESS],[r"""BESS INV V""",r"""BESS INV A"""]) 
    psspy.bus_frequency_channel([-1,POC_BUS],r"""POC FREQ""")
    psspy.voltage_and_angle_channel([-1,-1,-1,INF_BUS],[r"""GRID V""",r"""GRID A"""])
    psspy.voltage_and_angle_channel([-1,-1,-1,DUMMYBUS],[r"""DUMMY V""",r"""DUMMY A"""]) #23

    psspy.branch_p_and_q_channel([-1,-1,-1,POC_BUS,DUMMYBUS],r"""1""",[r"""POC P""",r"""POC Q"""]) #25
    psspy.branch_p_and_q_channel([-1,-1,-1,INV_BESS,MV_BESS_BUS],r"""1""",[r"""BESS INV P""",r"""BESS INV Q"""]) #27
    psspy.branch_p_and_q_channel([-1,-1,-1,INV_SF,MV_SF_BUS],r"""1""",[r"""SF INV P""",r"""SF INV Q"""]) #29
    
    psspy.var_channel([-1,PPC_VAR+16],r"""PPC_Vref""") #30
    psspy.var_channel([-1,PPC_VAR+60],r"""PPC_BatWSptMax""")
    psspy.var_channel([-1,PPC_VAR+61],r"""PPC_BatWSptMin""")
    psspy.var_channel([-1,PPC_VAR+62],r"""PPC_Ppv_cmd_inv""")
    psspy.var_channel([-1,PPC_VAR+63],r"""PPC_Qpv_cmd_inv""")
    psspy.var_channel([-1,PPC_VAR+64],r"""PPC_Pbess_cmd_inv""")
    psspy.var_channel([-1,PPC_VAR+65],r"""PPC_Qbess_cmd_inv""")

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
    psspy.rsol([1,0,0,0,0,0,0,0,0,1],[0.0, 2.0]) #check??
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
    BESS_INV_VAR = 301 + 0 #(Add six for PLB file) # no SF 
    PPC_VAR = 627 + 0 #(Add six for PLB file)    #625
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
    
    dif_Up=abs(100.0*dif_U/target_U)

    count = 0   
    
    iters = 0
    dt_test=0.0        
    while (dif_Up>0.5 and iters < 100):
        psspy.shunt_data(DUMMYBUS,r"""1""",1,[0.0,Yf]) # connect a fixed shunt impedance to invoke Overvoltage at the PoC
        
        dt_test=dt_test+0.01    
        psspy.run(1,10+dt_test,5000,10,0)   
        ierr, Bus_V=psspy.busdat(POC_BUS, 'PU') # 4001
        ierr, rval_Q_POC=psspy.brnmsc(POC_BUS,DUMMYBUS, '1', 'Q') 
        ierr, rval_Q_SF=psspy.brnmsc(INV_SF, MV_SF_BUS, '1', 'Q')  
        ierr, rval_Q_BESS=psspy.brnmsc(INV_BESS, MV_BESS_BUS, '1', 'Q')
        dif_Up=abs(100.0*dif_U/target_U)                                        
        dif_U=float(target_U) - float(Bus_V) 
        dYf = SCP * dif_U
        Yf=Yf+dYf
        print('dif_Up:',dif_Up,' dif_U:',dif_U,' dYf:',dYf)
        print(' II target_U,  Bus_V, Yf ,rval_Q_POC, rval_Q_SF, rval_Q_BESS')
        print(target_U, Bus_V, Yf, rval_Q_POC, rval_Q_SF, rval_Q_BESS)
        print('iters',iters)
        
        iters += 1
        
    psspy.run(1,10+dt_test+T_flt,5000,10,0)
    print(' --Shunt(%s pu) >>> target_U,  Bus_V, Yf ,rval_Q_POC, rval_Q_SF, rval_Q_BESS'%target_U)
    print(target_U, Bus_V, Yf, rval_Q_POC, rval_Q_SF, rval_Q_BESS)
    return Yf, PPC_Vref

 #print('PPC_Vref:',PPC_Vref)

def runFlat(cases_run, infname, Case_no, SCP, dyrfile, P_POC, Q_POC, runNum,counter):
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
    # SF_INV_VAR = 1 + 0 #(Add six for PLB file)
    # BESS_INV_VAR = 301 + 0 #(Add six for PLB file)
    # PPC_VAR = 625 + 0 #(Add six for PLB file)
    if counter==0:
        add_channels(SF_INV_VAR, BESS_INV_VAR, PPC_VAR)
        
    # PSSE OUT File Setup
    psspy.strt(0,SavDir + infname + ".out")                             
    ierr, Bus_V=psspy.busdat(POC_BUS, 'PU')    
    psspy.run(0, 5,1000,10,0)  

    # Update setpoint to match solution - same BESS and Plant in BESS only case
    psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1 + 18, P_POC*1000) # Plant Active Power
    psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1 + 13, 0) # BESS W Setpoint
    #psspy.change_cctbusomod_var(POC_BUS,PPC_MODEL,1 + 14,  Q_POC*1000) # Plant Reactive Power
    
    psspy.run(0, 10.0,1000,10,0)
    
    # read PPC Vref from the PSCAD model
    Vref_VAR_index = 16
    ierr, PPC_Vref = psspy.dsrval('VAR', PPC_VAR + Vref_VAR_index)
    ierr, PPC_Vref2 = psspy.chnval(30)
    # ierr, dummyvalP = psspy.chnval(26)
    # ierr, dummyvalQ = psspy.chnval(27)
    ierr, BESS_P = psspy.chnval(7)
    ierr, BESS_Q = psspy.chnval(8)
    ierr, SF_P = psspy.chnval(1)
    ierr, SF_Q = psspy.chnval(2)
    ierr, GRID_V = psspy.chnval(20)
    ierr, POC_V = psspy.chnval(13)
    ierr, POC_Pflat = psspy.chnval(24)
    ierr, POC_Qflat = psspy.chnval(25)
    # ierr, POC_tflat = psspy.chnval(26)
    print('PPC_Vref:',PPC_Vref)
    print('PPC_Vref2:',PPC_Vref2)
    print('POC_Pflat:',POC_Pflat)
    print('POC_Qflat:',POC_Qflat)
    # print('POC_tflat:',POC_tflat)
    
    return PPC_Vref2, GRID_V, POC_Pflat, POC_Qflat

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
    counter=0
    for case in cases_run:
        row = cases_run[case]
        print('row',row)   
            
        Case_no = row['Case'].value
        Run_test = row['Run_test'].value
        Test_Type = row['Test_Type'].value
        savefilepick = row['savefile'].value
        P_POC = float(row['ActiveP'].value) * MWrating
        Q_POC = float(row['ReactiveP'].value) * MWrating
        SCR = float(row['SCR'].value) 
        SCP = float(row['SCR'].value) * MWrating
        X_R = float(row['X_R'].value)
        Vref = float(row['PPC_Vref'].value)
        Trf_Tap = float(row['Trf_Tap'].value)
        infname = str(row['Test_File_Names'].value)
        B_P_xl = float(row['BESS_P'].value)
        Target_POC_V = float(row['Target POC V'].value)
        deratingtemp = float(row['derating'].value)

        print('Case_no, SCP, X_R, P_POC, Q_POC, Vref, infname, csv_file_DMAT')
        print(Case_no, SCP, X_R, P_POC, Q_POC, Vref, infname)
        
        # if savefilepick == "1pu": #this is a chainsaw. useful but globals are so dangerous. tread carefully
        #     global savefile
        #     savefile = savefile1pu

        #print('runNum:',runNum)
        print('maxRuns:',maxRuns)
        BESSUpdateP = B_P_xl; #should equal 1 if not calculated beforehand  
        for runNum in range(maxRuns):
            print('start',BESSUpdateP)
            SF_P, SF_Q, BESS_P, BESS_Q, Case_POC_P, Case_POC_Q, Case_POC_V = Case_Setup(cases_run, Case_no, SCP, SCR, X_R, P_POC, Q_POC, Vref, infname,runNum,BESSUpdateP, Target_POC_V,deratingtemp)
            dyrfile = setupdir + dyr_qconst
            # PPC_Vref, PPC_Vref2, POC_Pflat,POC_Qflat, BESS_P, BESS_Q, SF_P, SF_Q, GRID_V,POC_V = runFlat(cases_run, infname, Case_no, SCP, dyrfile, P_POC, Q_POC, runNum,counter)
            # counter=1
            print('Calculating BESS charge rate')
            print('SF_P:',SF_P)
            print('SF_Q:',SF_Q)
            print('BESS_P:',BESS_P)
            print('BESS_Q:',BESS_Q)
            #Update the BESS Power to charge the amount of headroom
            SF_ApparentPower=math.sqrt(pow((SF_P),2)+pow((SF_Q),2))   #SF Apparent  S = MVA
            BESS_ApparentPower=math.sqrt(pow((BESS_P),2)+pow((BESS_Q),2))   #BESS Apparent  S = MVA

            print('SF_ApparentPower:',SF_ApparentPower)
            print('BESS_ApparentPower:',BESS_ApparentPower)

            SFAPHeadroom = SFMWrating-SF_ApparentPower
            BESSAPHeadroom=BESSMWrating-BESS_ApparentPower
            
            print('SFAPHeadroom:',SFAPHeadroom)
            print('BESSAPHeadroom:',BESSAPHeadroom)

            #similar triangles
            SFPHeadroom = (SF_P)*(SFAPHeadroom/SF_ApparentPower)
            BESSPHeadroom = math.sqrt(pow((BESSMWrating),2)-pow((BESS_Q),2))
            print('SFPHeadroom',SFPHeadroom)
            print('BESSPHeadroom',BESSPHeadroom)
            print('runNum:',runNum)
            print('Case_POC_P:',Case_POC_P)
            print('Case_POC_Q:',Case_POC_Q)
            print('Case_POC_V:',Case_POC_V)
            print('maxRuns:',maxRuns)
            print('choosing if statement')
            if SFPHeadroom<0:
                print('SFHeadroom is going negative - ',SFPHeadroom)
                break
                # return 0
            if SFPHeadroom>=SFMWrating:
                print('SF has no more headroom - ',SFPHeadroom)
                break
                # return 0
            if runNum>5 and SFPHeadroom<0.2:
                print('SFHeadroom is small - ',SFPHeadroom)
                # break
                # return 0
            # elif abs(BESS_P)>BESSMWrating:
                # print('BESS cannot charge faster - ',BESS_P)
                # break
                # return 0
            if runNum<maxRuns-1:
                #charge the BESS slowly to ensure limits are not breached
                print('BESS_P',BESS_P)
                print('BESS_Q',BESS_Q)
                print('BESS_ApparentPower',BESS_ApparentPower)
                print('BESSAPHeadroom',BESSAPHeadroom)
                print('BESSPHeadroom',BESSPHeadroom)
                print('SF_P',SF_P)
                print('SF_Q',SF_Q)
                print('SF_ApparentPower',SF_ApparentPower)
                print('SFAPHeadroom',SFAPHeadroom)
                print('SFPHeadroom',SFPHeadroom)
                BESSUpdateP = 0
                buffer=0.98
                BESSUpdateP = -min((abs(BESS_P)+SFPHeadroom*buffer),BESSPHeadroom*buffer)
                # print('BESSUpdateP',BESSUpdateP)
                # while abs(BESSUpdateP) < abs(BESSPHeadroom*0.96) and abs(BESSUpdateP) < abs(abs(BESS_P)+SFPHeadroom*0.96):
                #     BESSUpdateP = -abs(BESSUpdateP-0.01)
                    # print('BESS Update Power - ',BESSUpdateP)
                # print('BESS cannot charge faster - ',BESSUpdateP)
                print('1.BESS active power has been updated to - ',BESSUpdateP)
                
        PPC_Vref2, GRID_V, POC_Pflat, POC_Qflat = runFlat(cases_run, infname, Case_no, SCP, dyrfile, P_POC, Q_POC, runNum,counter)            
        counter=1
        cases_run[Case_no]['PPC_Vref_flat'].value = "%.4f"%PPC_Vref2
        cases_run[Case_no]['Vsource_i_flat'].value = "%.4f"%GRID_V
        cases_run[Case_no]['POC_Pflat'].value = "%.4f"%POC_Pflat
        cases_run[Case_no]['POC_Qflat'].value = "%.4f"%POC_Qflat
     #Transient OverVoltage   
        if cases_run[Case_no]['Shunt(1.15pu)'].value == "Yes":
            Yf, PPC_Vref = Shunt_TOV_Cal(cases_run, infname, Case_no, SCP, dyrfile, target_U=1.15, T_flt=0.9)
            cases_run[Case_no]['Shunt(1.15pu)'].value = "%.4f"%Yf
            cases_run[Case_no]['PPC_Vref'].value = "%.4f"%PPC_Vref
        if cases_run[Case_no]['Shunt(1.2pu)'].value == "Yes":
            Yf, PPC_Vref = Shunt_TOV_Cal(cases_run, infname, Case_no, SCP, dyrfile, target_U=1.2, T_flt=0.1)
            cases_run[Case_no]['Shunt(1.2pu)'].value = "%.4f"%Yf
            cases_run[Case_no]['PPC_Vref'].value = "%.4f"%PPC_Vref
            
        test_workbook.save(xlsx_file_DMAT)
        # create_result_csv(ResultsDir=SavDir,infname=infname)
        counter+=1

if (__name__ == "__main__"):
    main()
    psspy.stop()

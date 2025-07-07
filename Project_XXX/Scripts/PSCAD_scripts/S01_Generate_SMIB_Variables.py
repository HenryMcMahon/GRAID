# ----------------------------------------------------------------------------------
# PSCAD S01_Generate_DMAT_Variables.py
# Version: 1.0.0 (2023/07/24)
#
# Description: This script is used to generate the variables used in the lookup
# table for the DMAT model. The variables are generated based on the data in the
# DMAT Study selection file.
#
# N.B.: You can check the version of the script above matches the version in the
# 01_DMAT_studies.xlsx configuration workbook.
#
# Prepared by : Marty Johnson
# Email : marty.johnson@epecgroup.com.au
# ----------------------------------------------------------------------------------


# ----------------------------------------------------------------------------------
# Import Essential libraries
# ----------------------------------------------------------------------------------
import os
import logging
import csv
import pandas as pd


# Working directory
working_dir = os.getcwd() + "\\"

# ----------------------------------------------------------------------------------
# Input Data
# ----------------------------------------------------------------------------------

# Input File Name (xlsx)
input_file_name = "01_SMIB_Studies-BESS-C.xlsx"
# Variable File Names (csv)
variable_file_name = "02_SMIB_Variables"
# Output File Names (csv)
output_file_name = "03_SMIB_Test_Names.csv"


# ----------------------------------------------------------------------------------
# Main Program
# ----------------------------------------------------------------------------------


def main():
    generateVariables()


def getStudyType(studyName):    #Add new numbers for SMIB studies
    studyTypes = {
        #DMAT Tests
        "Flat_Run_test_300s": 1,
        "Flat_Run_test_5s": 2,
        "3Ph_Balanced_Fault": 3,
        "2Ph_Unbalanced_Fault": 4,
        "1Ph_Unbalanced_Fault": 5,
        "L_L_Unbalanced_Fault": 6,
        "MFRT_S1": 7,
        "MFRT_S2": 8,
        "MFRT_S3": 9,
        "MFRT_S4": 10,
        "MFRT_S5": 11,
        "MFRT_Protection": 12,
        "DMAT_TempOV_1_15pu": 13,
        "DMAT_TempOV_1_2pu": 14,
        "DMAT_Vsetpoints": 15,
        "DMAT_Vgridsteps": 16,
        "DMAT_Qsetpoints": 17,
        "DMAT_PFsetpoints": 18,
        "DMAT_Psetpoints": 19,
        "DMAT_OverFreq_4Hz": 20,
        "DMAT_OverFreq_3s": 21,
        "DMAT_UnderFreq_4Hz": 22,
        "DMAT_UnderFreq_3s": 23,
        "DMAT_UnderFreq_4Hz_AP50": 24,
        "DMAT_UnderFreq_3s_AP50": 25,
        "DMAT_UnderFreq_4Hz_AP5": 26,
        "DMAT_UnderFreq_3s_AP5": 27,
        "GridVchange_ramp": 28,
        "GridVchange_step": 29,
        "Extended_Vdip_Recovery_0.8pu": 30,
        "Extended_Vdip_Recovery_0.5pu": 31,
        "Extended_Vdip_Recovery_0.1pu": 32,
        "GridPhaseAngle_step": 33,
        "POC_SCR_1_Pref_step": 34,
        "POC_FRT_Test": 35,
        "Site_Specific_FRT_Test": 36,
        "Irradiance_Neg_step": 37,
        "Irradiance_Pos_step": 38,
        "LVRT": 39,
        "HVRT": 40,
        "Additional_ROCOF_0.1": 41,
        "Additional_ROCOF_0.5": 42,
        "Additional_ROCOF_1.0": 43,
        "Additional_ROCOF_4.0": 44,
        "Freq_Oscillation_Rejection": 45,
        "Volt_Oscillation_Rejection": 46,
        #Preliminary Study Tests
        "CUO_PIA": 51,
        "CUO_PIA_10P": 52,
        "CUO_PIA_1_1pu": 53,
        "CUO_PIA_0_9pu": 54,
        "CUO_LVRT_0_7pu": 55,
        "CUO_LVRT_0_8pu": 56,
        "CUO_HVRT_1_15pu": 57,
        "CUO_HVRT_1_2pu": 58,
        "CUO_HVRT_1_25pu": 59,
        "CUO_HVRT_1_3pu": 60,
        "CUO_HVRT_1_35pu": 61,
        "PPC_Function_Flt_Vdist": 62,
        "PPC_Function_Flt_Freq": 63,
        #SMIB Studies
        "S52513_Vsetpoints": 64,
        "S52513_Vsetpoints_CAPlimit": 65,
        "S52513_Vsetpoints_INDlimit": 66,
        "S52513_Vgridsteps": 67,
        "S52513_Qsetpoints": 68,
        "S52513_PFsetpoints": 69,
        "S5253_OFtrip": 70,
        "S5253_UFtrip": 71,
        "S52511_OF": 72,
        "S52511_UF": 73,
        "S5254_LVRT": 74,
        "S5254_HVRT": 75,
        "S5254_CUO": 76,
        "S52514_Pdispatch": 77,
        "S52514_Pstep": 78,
        "S5255_BalancedFault": 79,
        "S5255_Unbalanced_2Ph_Fault": 80,
        "S5255_Unbalanced_1Ph_Fault": 81,
        "S5255_HVRT": 82,
        #"HVRT_Iq": 82,
        #"S52514_Pdispatch": 83,
    }
    return studyTypes.get(studyName, 0)


# Run Test
def generateVariables():
    # Remove old files
    if os.path.exists(working_dir + output_file_name):
        os.remove(working_dir + output_file_name)
    if os.path.exists(working_dir + variable_file_name + "_1.csv"):
        os.remove(working_dir + variable_file_name + "_1.csv")
    if os.path.exists(working_dir + variable_file_name + "_2.csv"):
        os.remove(working_dir + variable_file_name + "_2.csv")
    logging.info("Reading script input data...")
    # Read the input file
    input_file = pd.read_excel(
        working_dir + input_file_name, sheet_name="SMIB_studies", header=0, index_col=0
    )
    # Create the output file
    with open(working_dir + output_file_name, "w", newline="") as output_file:
        with open(working_dir + variable_file_name + "_1.csv", "w", newline="") as csvfile1:
            with open(
                working_dir + variable_file_name + "_2.csv", "w", newline=""
            ) as csvfile2:
                # Create the csv writer
                writer1 = csv.writer(csvfile1, delimiter=",")
                writer2 = csv.writer(csvfile2, delimiter=",")
                writer3 = csv.writer(output_file, delimiter=",")
                # Write the header rows
                writer1.writerow(
                    [
                        "! SimNum",
                        "! StudyType",
                        "! ActiveP",
                        "! ReactiveP",
                        "! SCR",
                        "! X_R",
                        "! Vsource_i",
                        "! Trf_Tap",
                        "! PPCVset",
                    ]
                )
                writer2.writerow(
                    [
                        "! SimNum",
                        "! Fixed_Shunt",
                        "! Fault_duration",
                        "! Rf_Ohms",
                        "! L_Ohms",
                        "! FRT_dip",
                        "! GS_R",
                        "! GS_L",
                        "! Run_Duration",
                        "! Oscillation_Freq",
                    ]
                )
                writer3.writerow(
                    [
                        "SimNum",
                        "Test_Name",
                    ]
                )

                # Define iterator
                iterator = 1

                # Iterate through the rows of the input file
                for index, row in input_file.iterrows():
                    # Check if study is enabled
                    if row["Run_test"] == "No":
                        continue
                    # Get the study type
                    studyType = getStudyType(row["Test_Type"])
                    # Check if the study type was found
                    if studyType == 0:
                        logging.warning(
                            "Study Type not found for "
                            + row["Test_Type"]
                            + " on case "
                            + str(index)
                        )
                        
                    # Write the data to CSV for the test name sheet.
                    writer3.writerow(
                        [
                            iterator,
                            row["Test_File_Names"],
                        ]
                    )

                    # Write the data to the output file
                    writer1.writerow(
                        [
                            iterator,
                            studyType,
                            row["ActiveP"],
                            row["ReactiveP"],
                            row["SCR"],
                            row["X_R"],
                            row["Vsource_i"],
                            row["Trf_Tap"],
                            row["PPCVset"],
                        ]
                    )
                    writer2.writerow(
                        [
                            iterator,
                            row["Fixed_Shunt"],
                            row["Fault_duration"],
                            row["Rf_Ohms"],
                            row["L_Ohms"],
                            row["FRT_dip"],
                            row["GS_R"],
                            row["GS_L"],
                            row["Run_Duration"],
                            row["Oscillation_Freq"],
                        ]
                    )
                    # Increment the iterator
                    iterator += 1
                                
                if iterator < 3:
                    # Write the data to the output file
                    writer1.writerow(
                        [
                            iterator,
                            1,
                            1,
                            1,
                            1,
                            1,
                            1,
                            1,
                            1,
                        ]
                    )
                    writer2.writerow(
                        [
                            iterator,
                            1,
                            1,
                            1,
                            1,
                            1,
                            1,
                            1,
                            1,
                            1,
                        ]
                    )
    logging.info("Script input data read successfully.")


# Calling the main Program Function
if __name__ == "__main__":
    main()

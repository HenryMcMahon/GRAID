#----------------------------------------------------------------------------------
# Prepared by : Marty Johnson
# Email : marty.johnson@epecgroup.com.au
#----------------------------------------------------------------------------------


#----------------------------------------------------------------------------------
# Import Essential libraries
#----------------------------------------------------------------------------------
import sys, os
sys.path.append(r"C:\Python27\Lib\site-packages\mhrc")
import automation.controller
import logging
import csv
from automation.utilities.file import OutFile
import time
import mhi.pscad
from mhi.pscad.utilities.file import OutFile

#----------------------------------------------------------------------------------
# Import Local Scripts
#----------------------------------------------------------------------------------
import S01_Generate_SMIB_Variables


#----------------------------------------------------------------------------------
# Configuration the PSCAD
#----------------------------------------------------------------------------------
fortran_version = "Intel(R) Visual Fortran Compiler XE 15.0.7.287"
fortran_ext = ".if18_x86"
SIMULATOR = True
RESULTOR = True
PLOTTING = True
SMALLFILE = True
controller = automation.controller.Controller()
#----------------------------------------------------------------------------------

# Working directory
working_dir = os.getcwd() + "\\"

#----------------------------------------------------------------------------------
# Input Data
#----------------------------------------------------------------------------------

# Output File Names (csv)
output_file_names = "03_SMIB_Test_Names.csv"

# Workspace, Project & Simulation Set Name
workspace_name = "MMHY_SMIB"
project_name = "MMHY_SMIB"
sim_set_name = "DMAT_Study_Set"

available_cores = 4
Simulation_time = 200

#----------------------------------------------------------------------------------
# Main Program
#----------------------------------------------------------------------------------


# Main Program
def main():
    Run_Tests(Simulation_time)



# Run Test
def Run_Tests(Tend):
    # Test List with Test Setup and File Names
    test_list_fnames = working_dir + output_file_names

    # Total Tests
    total_tests = len(open(test_list_fnames).readlines())-1

    convert_out_to_csv(test_list_fnames)


# Function to Convert the PSCAD out files to CSV file Format
def convert_out_to_csv(test_list_fnames):
    Out_No = 1
    with open(test_list_fnames, "r") as csvfl:
        csv_reader = csv.reader(csvfl)
        Header = next(csv_reader)
        
        for row in csv_reader:
            resfilename = row[1]
                
            # Selecting the .inf file to convert to csv
            result_file = working_dir + "..\\model\\" + project_name + fortran_ext + "\\" + project_name + "_" + str(Out_No).zfill(2)
                
            # Collecting all the out files for the corresponding .inf file
            out_file = OutFile(result_file)
                
            # Converting the out files to a csv files
            out_file.toCSV()
            
            # Getting the temporary File
            result_file = working_dir + "..\\model\\" + project_name + fortran_ext + "\\" + project_name + "_" + str(Out_No).zfill(2) + ".csv"
                
            # A new name for the out.CSV file, which contains the case name
            fname = working_dir + "..\\results\\" + resfilename + ".csv"
                
            try:
                # Removing the previous CSV file if already exists
                os.remove(fname)
            except:
                pass
                # Renaming the out.CSV file to caseName.csv
            os.rename(result_file, fname)
                
            print("Results for " + resfilename + " is generated")
                
            Out_No += 1



# Calling the main Program Function
if (__name__ == "__main__"):
    main()

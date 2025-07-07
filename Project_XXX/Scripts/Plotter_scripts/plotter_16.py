#PLEASE DONT TOUCH IT, IF YOU BREAK IT IM NOT GOING TO FIX IT
#PLEASE DONT TOUCH IT, IF YOU BREAK IT IM NOT GOING TO FIX IT
#PLEASE DONT TOUCH IT, IF YOU BREAK IT IM NOT GOING TO FIX IT
#PLEASE DONT TOUCH IT, IF YOU BREAK IT IM NOT GOING TO FIX IT
#PLEASE DONT TOUCH IT, IF YOU BREAK IT IM NOT GOING TO FIX IT

######################################################################################################################################
########################################     Developed by Luciano Roco       #########################################################
########################################             version 16              #########################################################
######################################################################################################################################

import sys
import subprocess
import importlib
profile = 0
if profile:
    import cProfile, pstats, io
    
    profiler = cProfile.Profile()
    profiler.enable()

def install_packages(packages):
    """
    Install the given list of packages using pip.
    """
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", *packages])
    except subprocess.CalledProcessError as e:
        print(f"An error occurred while installing packages: {e}")
        sys.exit(1)

# List of required packages with their import names and install names if different
required_packages = {
    'pandas': 'pandas',
    'xlwings': 'xlwings',
    'matplotlib': 'matplotlib',
    'numpy': 'numpy',
    'PyPDF2': 'PyPDF2',
    'scipy': 'scipy',
    'openpyxl': 'openpyxl',
    'psutil': 'psutil',
}

class color:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    END = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    BLACK = '\033[30m'
    RED = '\033[31m'
    GREEN = '\033[32m'
    YELLOW = '\033[33m'
    BLUE = '\033[34m'
    MAGENTA = '\033[35m'
    CYAN = '\033[36m'
    WHITE = '\033[37m'
    UNDERLINE = '\033[4m'
    RESET = '\033[0m'
    
missing_packages = []

# Check for missing packages
for import_name, package_name in required_packages.items():
    try:
        importlib.import_module(import_name)
    except ImportError:
        missing_packages.append(package_name)

# Install missing packages if any
if missing_packages:
    print(f"Installing missing packages: {', '.join(missing_packages)}")
    install_packages(missing_packages)
else:
    #print("All required packages are already installed.")
    pass

# Now, perform all imports
import time
import itertools
import os  # Needed for os.system if you still intend to use it
import matplotlib.pyplot as plt
import matplotlib.image as image
import matplotlib as mpl
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
from matplotlib.ticker import MaxNLocator
from collections import defaultdict
from cycler import cycler
import ctypes
import re
import shutil
import traceback
from datetime import datetime
import numpy as np
from numpy.lib.stride_tricks import sliding_window_view
import pandas as pd
import xlwings as xw
import matplotlib.pyplot as plt
import PyPDF2
from PyPDF2 import PdfWriter, PdfReader
import bisect

import warnings

import psutil



warnings.filterwarnings("ignore")
mpl.rcParams['axes.prop_cycle'] = cycler(color='rbgmcy')
pd.set_option('mode.chained_assignment', None)
mpl.use('pdf')

#mpl.use('Qt5Agg')
#mpl.set(gcf,'Renderer','zbuffer')
#mpl.style.use('fast')




if len(sys.argv) > 1:
    #print("Reading template: ", sys.argv[1])
    config_excel = sys.argv[1]
    if ".xlsx" not in config_excel: config_excel += ".xlsx"
else: 
    #config_excel = 'plotter_CBESS.xlsx'
    print(r"Select a excel Template in the batchfile (after the % symbol), it should be in the same folder as the batch file (dont forget the .xlsx ext)")
    exit()
    
force_ncols = None #If None it will do it Auto
output_folder = "Plots_PDF_" + config_excel.split(".xlsx")[0]
output_folder_html = output_folder + os.sep + "HTML"

rec_perc = 0.95
errband_perc = 0.1
#resample_to = 0.0001 #Improving time accuracy in recovery/settling/rising analysis

    
min_yticks = 5
min_xticks = 4
legend_size = 5.5
axis_size = 6.5
line_size = 0.3

path_parent = os.path.dirname(os.getcwd())
working_dir = os.getcwd() + "\\"


    
    
def to_df(sheet_name):
    
    # Access the specified sheet
    sheet = wb.sheets[sheet_name]
    
    # Read the data from the sheet
    data_range = sheet.used_range
    data = data_range.options(pd.DataFrame, index=False, header=True).value
    
    # Close the workbook
    #wb.close()
    
    # Filter out rows with empty first column
    data = data[data.iloc[:, 0].notnull()]
    data.reset_index(drop=True, inplace=True)
    return data
def extract_values_point(string):
    pattern = r'point\((\d+\.?\d*)\)'
    matches = re.findall(pattern, string)
    return [float(match) for match in matches]

def extract_values_point_all(string):
    pattern = r'point_all\((\d+\.?\d*)\)'
    matches = re.findall(pattern, string)
    return [float(match) for match in matches]

def extract_values_point_all(string):
    pattern = r'point_all\(([\d\.]+(?:,\s*[\d\.]+)*)\)'
    matches = re.findall(pattern, string)
    return [float(val) for match in matches for val in match.split(',')]

def extract_values_ort_gral(string):
    # Updated pattern to handle spaces around values
    #Vgrid, Vpoc, Qpoc, freq
    pattern = r'ORT\s*\(\s*([^,]+?)\s*,\s*([^,]+?)\s*,\s*([^,]+?)\s*,\s*(\d*\.?\d*)\s*\)'
    matches = re.findall(pattern, string)
    return [(match[0].strip(), match[1].strip(),match[2].strip(), float(match[3])) for match in matches]

def extract_values_ort_gral(string):
    # Updated pattern to handle spaces around values
    # Vgrid, Vpoc, Qpoc, freq, tmin, tmax
    pattern = r'ORT\s*\(\s*([^,]+?)\s*,\s*([^,]+?)\s*,\s*([^,]+?)\s*,\s*(\d*\.?\d*)\s*(?:,\s*(\d*\.?\d*)\s*,\s*(\d*\.?\d*))?\s*\)'
    matches = re.findall(pattern, string)
    
    result = []
    for match in matches:
        Vgrid = match[0].strip()
        Vpoc = match[1].strip()
        Qpoc = match[2].strip()
        freq = float(match[3])
        if match[4]:
            tmin = float(match[4]) 
        else:
            #print(f"{color.WARNING} ORT Assessment: Using tmin = 10s {color.END}")
            tmin = "def"
            
        if match[5]:
            tmax = float(match[5]) 
        else:
            tmax = "def"
            #print(f"{color.WARNING} ORT Assessment: Using tmax = 29s {color.END}")
        result.append((Vgrid, Vpoc, Qpoc, freq, tmin, tmax))
    
    return result


def extract_values_insert_signal(string):
    # Updated pattern to capture three numeric values followed by quoted strings
    pattern = r'insert_signal\s*\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*["\']([^"\']*)["\']\s*,\s*["\']([^"\']*)["\']\s*(?:,\s*["\']([^"\']*)["\'])?\s*(?:,\s*["\']([^"\']*)["\'])?\s*(?:,\s*["\']([^"\']*)["\'])?\s*\)'

    matches = re.findall(pattern, string)
    # print(matches)  # Debug: Show the matched groups
    result = []
    
    for match in matches:
        # Extract values
        row = float(match[0])
        col = float(match[1])
        file_id = float(match[2])
        
        # Handle strings inside quotes
        name = match[3].strip()  
        legend = match[4].strip() if match[4] else None
        axis_text = match[5].strip() if match[5] else None
        style = match[6].strip() if match[5] else None
        color = match[7].strip() if match[6] else None
        
        result.append((file_id, row, col, name, legend, axis_text, style, color))

    return result


def extract_values_point_gral(string):
    #pattern = r'point\(([^,]+),(\d+\.?\d*)\)'
    #matches = re.findall(pattern, string)
    #return [(match[0], float(match[1])) for match in matches]
    pattern = r'point\(([^,]+),(\d+\.?\d*),?(\d*)\)'
    matches = re.findall(pattern, string)
    return [(match[0], float(match[1]), int(match[2]) if match[2] else 3) for match in matches]

def extract_values_shift(string):
    pattern = r'xshift\(([^,]+),(-?\d+\.?\d*)\)'
    matches = re.findall(pattern, string)
    return [(match[0], float(match[1])) for match in matches]



def extract_values_shift_y(string):
    pattern = r'yshift\(\s*([^,]+)\s*,\s*([^,]+?)\s*,\s*(-?\d*\.?\d*)\s*\)'
    pattern = r'yshift\(\s*([^,]+)\s*,\s*([^,]+)\s*,\s*(-?\d*\.?\d*)\s*\)'
    matches = re.finditer(pattern, string)
    return [(match.group(1), match.group(2), float(match.group(3))) for match in matches]



def extract_ncols(string):
    pattern = r'ncols\((\d+)\)'
    matches = re.findall(pattern, string)
    return [int(match) for match in matches]

def extract_timecol(string):
    pattern = r'timecol\((\d+)\)'
    matches = re.findall(pattern, string)
    return [int(match) for match in matches]

def extract_legendsize(string):
    pattern = r'legend_size\((\d*\.?\d+)\)'
    matches = re.findall(pattern, string)
    return [int(match) for match in matches]
def extract_axissize(string):
    pattern = r'axis_size\((\d*\.?\d+)\)'
    matches = re.findall(pattern, string)
    return [int(match) for match in matches]

#    max(Ir_rms)
def extract_max_point_gral(string):
    pattern = r'max\(([^,]+),?(\d+\.?\d*),?(\d+\.?\d*)\),?(\d*)\)?'
    matches = re.findall(pattern, string)
    return [(match[0], float(match[1]) if match[1] else 0, float(match[2]) if match[2] else 100, int(match[3]) if match[3] else 3) for match in matches]

def extract_hide_legend(string):
    pattern = r'hide_legend\(([^)]+)\)'
    matches = re.findall(pattern, string)
    return matches

def extract_delay_rise(string):
    pattern = r'delay_rise\((\d*\.?\d+)\)'
    matches = re.findall(pattern, string)
    return [int(match) for match in matches]

def extract_avg_rise(string):
    pattern = r'avg_rise\((\d*\.?\d+)\)'
    matches = re.findall(pattern, string)
    return [int(match) for match in matches]

def extract_min_point_gral(string):
    #pattern = r'point\(([^,]+),(\d+\.?\d*)\)'
    #matches = re.findall(pattern, string)
    #return [(match[0], float(match[1])) for match in matches]
    pattern = r'min\(([^,]+),(\d+\.?\d*),?(\d+\.?\d*)\),?(\d*)\)'
    matches = re.findall(pattern, string)
    return [(match[0], float(match[1]) if match[1] else 0, float(match[2]) if match[2] else 100, int(match[2]) if match[3] else 3) for match in matches]

def extract_values_line(string):
    pattern = r'line\((\d+\.?\d*)\)'
    matches = re.findall(pattern, string)
    return [float(match) for match in matches]


try:
    wb = xw.Book(config_excel)
except Exception as e:
    print(f"{color.RED}xlwings failed with error: {e}{color.END}")
    print(f"{color.WARNING}Trying to open with Excel directly...{color.END}")
    os.startfile(config_excel)  # Open with the default Excel application
    time.sleep(5)  # Give Excel time to open the file

    try:
        wb = xw.Book(config_excel)
    except Exception as e2:
        print(f"{color.RED}Still failed to open with xlwings after launching Excel: {e2}{color.END}")
        raise
    

company_logo = image.imread(r"N:\LucianoRoco\Scripts\Plotter\EPEC_IMG.png")


df_cols = ["CSV","Page","Initial Time of analysis","Final Time of analysis","Resampling","Extra Info","Signal","Parameter","Delta Time Measured from","Actual Time","Delta Time"]
data_report = []




def generate_linspace(tim, sample_time):
    # Get the start and end times from tim
    #print("AWDAWD")
    #print(len(tim))
    start = tim[0]
    end = tim[-1]
    
    # Calculate the number of points
    num_points = int((end - start) / sample_time) + 1
    
    # Generate the linspace
    linspace = np.linspace(start, end, num_points)
    
    return linspace



if not os.path.exists(output_folder):
    os.makedirs(output_folder)
    


def save_string_to_file(string: str, path: str):
    # Create the file path
    file_path = os.path.join(path)
    try:
        # Check if the file exists
        if os.path.exists(file_path):
            # Append the string as a new line
            with open(file_path, 'a') as f:
                f.write('\n' + string)
        else:
            # Write the string to the file
            with open(file_path, 'w') as f:
                f.write(string)
    except Exception as e:
        print(f"ERROR: Couldnt update the txt file with time results -  {e}")



def increase_sampling(signal, time, new_t_sample):
    signal = list(signal)
    time = list(time)
    new_signal = []
    new_time = []
    i = 0
    
    if abs(time[1] - time[0]) <= new_t_sample:
        return time, signal
    
        
    while i < len(signal) - 1:
        if time[i] == time[i + 1] and signal[i] != signal[i+1]:
            # Handle step in signal
            step_value_before = signal[i]
            step_value_after = signal[i + 1]
            
            # Add the value before the step
            new_signal.append(step_value_before)
            new_time.append(time[i])
            print("def",time[i],signal[i],time[i+1],signal[i+1],time[i+2],signal[i+2])
            
            # Divide the step into 20 samples
            step_duration = 0.0000001  # Increment for time
            Nsamples = 2
            for j in range(1, Nsamples+1):  # Generate 20 intermediate points
                interp_time = time[i] + j * step_duration
                interp_signal = step_value_before + (step_value_after - step_value_before) * (j / Nsamples)
                new_time.append(interp_time)
                new_signal.append(interp_signal)
                
                #print(interp_time,interp_signal)
            
            i += 1  # Move past the step
            continue
        
        
        if time[i] + new_t_sample > time[i + 1]:
            i += 1
            continue
        
        # Calculate slope between the two samples
        slope = (signal[i + 1] - signal[i]) / (time[i + 1] - time[i])
        num_new_samples = int((time[i + 1] - time[i]) / new_t_sample)
        
        # Append original sample
        new_signal.append(signal[i])
        new_time.append(time[i])
        
        # Append new samples
        for j in range(num_new_samples):
            new_signal.append(signal[i] + slope * (j + 1) * new_t_sample)
            new_time.append(time[i] + (j + 1) * new_t_sample)
        if new_time[-1] < 10.05 and new_time[-1]> 9.98:
            print("",new_time[-1],new_signal[-1])
        
        i += 1
    
    # Append the last sample
    new_signal.append(signal[-1])
    new_time.append(time[-1])
    
    
    return new_time, new_signal


def increase_sampling(signal, time, new_t_sample):
    signal = list(signal)
    time = list(time)
    new_signal = []
    new_time = []
    i = 0

    while i < len(time) - 1:
        # Step detection
        if (abs(time[i+1] - time[i]) < 1e-5) and (abs(signal[i+1] - signal[i]) > 1e-5) :
            # 1. Add the value *before* the step
            new_time.append(time[i])
            new_signal.append(signal[i])

            # 2. Insert 20 samples for the step transition
            for j in range(1, 21):
                new_time_step = time[i] + 1e-7 * j  # Very small time increment
                new_signal_step = signal[i] + (signal[i+1] - signal[i]) * (j / 20)
                new_time.append(new_time_step)
                new_signal.append(new_signal_step)

            i += 1  # Move to the next sample in the *original* signal (after the step)

            # Append value *AFTER* the step, (no interpolation yet).  This is KEY!
            new_time.append(time[i])  # Use the *next* time value (the one after the step)
            new_signal.append(signal[i]) # And its corresponding signal

            continue # Jump the normal interpolation




        # ------ Standard Interpolation ------
        if time[i] + new_t_sample <= time[i+1]:

            # Number of new samples between the current and next original points
            num_new_samples = int(np.floor((time[i+1] - time[i]) / new_t_sample))

            # Add the original sample (before interpolating)
            new_time.append(time[i])
            new_signal.append(signal[i])

            if num_new_samples > 0:
                 # Calculate the slope
                slope = (signal[i+1] - signal[i]) / (time[i+1] - time[i])
                 # Generate interpolated samples.

                for j in range(1, num_new_samples + 1):
                    interp_time = time[i] + j * new_t_sample
                    interp_signal = signal[i] + slope * (interp_time - time[i])
                    new_time.append(interp_time)
                    new_signal.append(interp_signal)

            i += 1 # Increment to the next original sample
        # if next sample is nearer, just add original points
        else:
            new_time.append(time[i])
            new_signal.append(signal[i])
            i+=1

    # Append the final sample
    new_time.append(time[-1])
    new_signal.append(signal[-1])
    
    for k in range(len(new_time)):
        if 10 < new_time[k] < 10.011:  # Adjusted range to capture relevant points
            #print(f"{new_time[k]:.8f}, {new_signal[k]:.8f}")
            pass
            
    return new_time, new_signal
    


def calculate_ncols(fig, ax, labels, font_size=6):
    # Get subplot width in inches
    subplot_width = ax.get_window_extent().transformed(fig.dpi_scale_trans.inverted()).width
    if len(labels) > 6:
        kk = 100
    else:
        # Calculate label length in inches for each label
        kk = 75
        for i in labels:
            if "Settling" in i or "Rising" in i:
                kk = 110
                break
            
    label_lengths = [len(label) * font_size / kk for label in labels]

    # Calculate total length of all labels
    total_length = sum(label_lengths)

    # Calculate number of columns based on total length and subplot width
    ncols = int(np.floor(subplot_width / (total_length / len(labels))))

    # Limit number of columns to 4
    if ncols >= 6:
        ncols = 6
    if ncols <= 1:
        ncols = 1
    return ncols

def sanitize_legend_entries(handles, labels, hide_legends):
    # Filter out entries that contain "none" or match hide_legends
    filtered = [
        (h, l) for h, l in zip(handles, labels) 
        if "none" not in l.lower() and all(hide not in l for hide in hide_legends)
    ]
    if not filtered:
        return [], []
    # Unzip filtered pairs back into separate lists
    handles_filtered, labels_filtered = zip(*filtered)
    return list(handles_filtered), list(labels_filtered)


def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]
def extract_values_from_filename(filename):
    p_start = 1
    g_start = filename.find("_g") + 2
    t_start = filename.find("_t") + 2
    pdf_end = filename.find(".pdf")

    if p_start > 0 and g_start > p_start and t_start > g_start and pdf_end > t_start:
        p_value = filename[p_start:g_start - 2]
        g_value = filename[g_start:t_start - 2]
        t_value = filename[t_start:pdf_end]
        return p_value, g_value, t_value

    return None, None, None  # Return None for each value if the filename doesn't match the expected pattern

def merge_txt_to_xlsx_by_group(input_folder: str, output_folder: str, output_filename: str):
    
    # Create the output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Create a dictionary to store the dataframes by group
    df_groups = {}
    countt = 0
    # Get all the txt files in the input folder
    txt_files = [f for f in os.listdir(input_folder) if f.endswith('.txt')]
    txt_files.sort(key=natural_sort_key)
    # Read each txt file and append its contents to the appropriate group in the dictionary
    for txt_file in txt_files:
        file_path = os.path.join(input_folder, txt_file)
        reading_tries = 0
        while 1:
            try:
                #print("trying again..")
                df = pd.read_csv(file_path, delimiter=',', header=None,keep_default_na=False)
                break
            except Exception as e:
                reading_tries += 1
                time.sleep(3)
            if reading_tries > 5:
                #print("ERROR merging txt files to a xlsx. %s")
                break
        if reading_tries > 5: 
            print("Problem converting to csv: %s" %txt_file)
            continue
        
        #df = df.applymap(lambda x: pd.to_numeric(x, errors='ignore'))
        df = df.applymap(lambda x: round(pd.to_numeric(x, errors='ignore'), 4) if isinstance(x, (int, float)) else x)
        # Check if the filename contains the '_g' and '_t' substrings
        if '_g' in txt_file and '_t' in txt_file:
            # Extract the group number from the filename
            group_num = txt_file.split('_g')[1].split('_t')[0]
            
            # Add the dataframe to the appropriate group in the dictionary
            if group_num not in df_groups:
                df_groups[group_num] = []
            df_groups[group_num].append(df)
            countt +=1
    if countt != 0:
        print("\nCreating xlsx files with time analysis (recovery,settling,rise)")
        # Merge and save each group of dataframes to a separate xlsx file
        for group_num, df_list in df_groups.items():
            merged_df = pd.concat(df_list, axis=0, ignore_index=True)
            
            # Create the output file path for this group
            output_file_path = os.path.join(output_folder, f'{output_filename}_{group_num}.xlsx')
            
            # Save the merged dataframe to the xlsx file
            merged_df.to_excel(output_file_path, index=False)
    
    
        
def merge_pdfs_by_group(input_folder, output_folder, output_filename):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    intcount = 0
    pdf_groups = {}

    pdf_files = [f for f in os.listdir(input_folder) if f.endswith('.pdf')]
    pdf_files.sort(key=natural_sort_key)  # Sort the files using natural sort

    for pdf_file in pdf_files:
        p, g, t = extract_values_from_filename(pdf_file)
        if g:
            group_num = g
            if group_num not in pdf_groups:
                pdf_groups[group_num] = []
            pdf_groups[group_num].append(os.path.join(input_folder, pdf_file))

    for group_num, pdf_list in pdf_groups.items():
        output_pdf = PyPDF2.PdfWriter()
        print(f"Merging pdf group {group_num}")
        intcount += 1
        for pdf_path in pdf_list:
            pdf_reader = PyPDF2.PdfReader(pdf_path)
            for page in pdf_reader.pages:
                output_pdf.add_page(page)

        # Save the merged PDF to the output folder
        merged_output_path = os.path.join(output_folder, f'{output_filename}_g{group_num}.pdf')
        with open(merged_output_path, 'wb') as output_file:
            output_pdf.write(output_file)

        # Create a folder for the group with the same name
        group_output_folder = os.path.join(output_folder, f'Group_{group_num}')
        os.makedirs(group_output_folder, exist_ok=True)

        # Move the individual PDF files to the group's folder
        for pdf_path in pdf_list:
            if pdf_path != merged_output_path:  # Exclude the merged PDF
                shutil.move(pdf_path, os.path.join(group_output_folder, os.path.basename(pdf_path)))
    return intcount
                
def move_pdfs_to_group_folders(input_folder, output_folder):
    print("Moving pdf files to the group folders")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    pdf_files = [f for f in os.listdir(input_folder) if f.endswith('.pdf')]
    pdf_files.sort(key=natural_sort_key)  # Sort the files using natural sort

    for pdf_file in pdf_files:
        p, g, t = extract_values_from_filename(pdf_file)
        if g:
            group_num = g
            # Create a folder for the group with the same name
            group_output_folder = os.path.join(output_folder, f'Group_{group_num}')
            os.makedirs(group_output_folder, exist_ok=True)
            # Move the individual PDF file to the group's folder
            shutil.move(os.path.join(input_folder, pdf_file), os.path.join(group_output_folder, os.path.basename(pdf_file)))

def merge_pdfs_in_group_folders(output_folder, output_filename):
    intcount = 0
    lastpdf = ""
    for group_output_folder in os.listdir(output_folder):
        if not group_output_folder.startswith('Group_'):
            continue
        group_num = group_output_folder.split('Group_')[1]
        # Merge the PDF files in the group's folder
        output_pdf = PdfWriter()
        print(f"{color.GREEN}Merging pdf group {group_num}{color.END}")
        
        
        pdf_files = [f for f in os.listdir(os.path.join(output_folder, group_output_folder)) if f.endswith('.pdf')]
        pdf_files.sort(key=natural_sort_key)
        
        for pdf_path in pdf_files:
            if not pdf_path.endswith('.pdf'):
                continue
            pdf_reader = PdfReader(os.path.join(output_folder, group_output_folder, pdf_path))
            for page in pdf_reader.pages:
                output_pdf.add_page(page)

        # Save the merged PDF to the output folder
        merged_output_path = os.path.join(output_folder, f'{group_num}.pdf')
        with open(merged_output_path, 'wb') as output_file:
            output_pdf.write(output_file)
        intcount += 1
        lastpdf = merged_output_path
    return intcount,lastpdf

    
def last_row(sheet):
    return sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row




#company_logo = image.imread(working_dir + "04_epec_logo.jpg")
def get_last_modification_date(file_path):
    if os.path.exists(file_path):
        file_stats = os.stat(file_path)
        modification_time = file_stats.st_mtime
        formatted_date = time.strftime("%d/%m/%Y - %H:%M", time.localtime(modification_time))
        return formatted_date
    else:
        return "File not found."
# List to store text instances




# Flag to track the number of clicks
click_count = 0

# List to store text instances
text_instances = []
last_x = 1
# Initialize variables for storing the coordinates of the first click
x_first_click = None
y_first_click = None

def delete_txt_files(folder_path: str, strings: list):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path) and filename.endswith('.txt'):
            with open(file_path, 'r') as f:
                file_content = f.read()
            if any(string in file_content for string in strings):
                try:
                    os.remove(file_path)
                except Exception as e:
                    print(f"Error: Couldn't remove the file {file_path} --  {e}")
                    
def delete_txt_files(folder_path: str, strings: list):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path) and filename.endswith('.txt'):
            if any(string in filename for string in strings):
                try:
                    os.remove(file_path)
                    #print(f"Deleted file: {file_path}")
                except Exception as e:
                    #print(f"Error: Couldn't remove the file {file_path} --  {e}")
                    pass
            #else:
                #print(f"No matching strings found in filename: {filename}")
        #else:
            #print(f"Not a .txt file: {filename}")
            
def delete_folders(folder_path: str, strings: list):
    print("Removing temporary folders... :)")
    for dirname in os.listdir(folder_path):
        dir_path = os.path.join(folder_path, dirname)
        if os.path.isdir(dir_path):
            if any(string in dirname for string in strings):
                try:
                    shutil.rmtree(dir_path)
                    #print(f"Deleted directory: {dir_path}")
                except Exception as e:
                    print(f"Error: Couldn't remove the directory {dir_path} --  {e}")
                    pass
                
def get_index(lista,value,backwards = 0):
    if backwards == 0:
       # print(len(lista))
        for t in range(len(lista)):
            #print(lista[t])
            if float(lista[t]) >= float(value):
                
                return t-1,lista[t-1]
    else:
        for t in reversed(range(len(lista))):
            if lista[t] <= value:
                return t,lista[t]
        
    #if backwards == 0:
    #    return len(lista)-1,lista[-1]
    #else:
    #    return 0,lista[0]
    
def get_index(lista, value, backwards=False):
    """
    Returns a tuple: (index, lista[index]) where:
      - If backwards=False (default): find the position just before the first element >= value
      - If backwards=True: find the position of the last element <= value

    NOTE: Requires that `lista` is sorted in ascending order.
    """
    if not backwards:
        # Find the insertion point to keep lista sorted if value were inserted.
        idx = bisect.bisect_left(lista, value)

        # Handle edge cases:
        if idx == 0:
            # Original code returned t-1 => -1 if the first item >= value is at index 0
            # That effectively means the last element in the list. 
            return -1, lista[-1]
        elif idx >= len(lista):
            # Means all elements in lista are < value => return the last element
            return len(lista) - 1, lista[-1]
        else:
            return idx - 1, lista[idx - 1]

    else:
        # Find the last element <= value
        # bisect_right gives the insertion point to the right of any existing entries of value
        idx = bisect.bisect_right(lista, value) - 1

        # If idx == -1, value is smaller than all elements
        if idx < 0:
            return None, None  # or handle however you prefer
        return idx, lista[idx]
    


def moving_average(a, n=3):
    return np.average(sliding_window_view(a, window_shape = n), axis=1)
def window_rms(a, window_size):
  a2 = np.power(a,2)
  window = np.ones(window_size)/float(window_size)
  return np.sqrt(np.convolve(a2, window, 'same'))
def window_median(a, window_size):
  if window_size % 2 == 0:  # if window_size is even
    window_size += 1  # make it odd
  return medfilt(a, window_size)
def smooth(a, alpha=0.02):
    #THIS USES Exp Mov Avg
    alpha = 1-alpha
    result = np.zeros_like(a)
    result[0] = a[0]
    for i in range(1, len(a)):
        result[i] = alpha * a[i] + (1 - alpha) * result[i-1]
    return result

def evaluate_formula(formula, df, tim, ttt_min, ttt_max):
    # Define a list of valid operators
    valid_operators = ['+', '-', '*', '/', '(', ')', 'sqrt', 'max', 'min', 'vini','vinird', 'vend','abs','mavg','rms','ema','median','smooth','round','cos','sin','tan','acos','asin','atan',]

    # Split the formula into tokens
    tokens = re.findall(r'\b[\w\.]+\b|\W', formula)

    # Initialize an empty list to store the processed tokens
    processed_tokens = []
    #((df["INV1_1001_VOEL_IQ"].values)-0.448217))/np.abs(0.9-(df["WDBESS_INV_1001_V"].values))
    
    # Iterate over the tokens
    for token in tokens:
        # Check if the token is a valid operator
        if token in valid_operators:
            # If the token is a valid operator, replace it with its numpy equivalent (if necessary)
            if token == 'sqrt':
                token = 'np.sqrt'
            elif token == 'max':
                token = 'np.amax'
            elif token == 'min':
                token = 'np.amin'
            elif token == 'abs':
                token = 'np.abs'
            elif token == 'sin':
                token = 'np.sin'
            elif token == 'cos':
                token = 'np.cos'
            elif token == 'tan':
                token = 'np.tan'
            elif token == 'asin':
                token = 'np.asin'    
            elif token == 'acos':
                token = 'np.acos'
            elif token == 'atan':
                token = 'np.atan'
                
                
                    
            elif token == 'vini':
                next_token_index = tokens.index(token)+2
                next_token = tokens[next_token_index]
                if next_token in df.columns:
                    token = str(df[next_token].values[ttt_min])
                    tokens[next_token_index] = token
                    tokens.remove('vini')
                    if '(' in tokens: tokens.remove('(')
                    if ')' in tokens: tokens.remove(')')
            elif token == 'vinird':
                next_token_index = tokens.index(token)+2
                next_token = tokens[next_token_index]
                if next_token in df.columns:
                    token = str(np.round(df[next_token].values[ttt_min],3))
                    tokens[next_token_index] = token
                    tokens.remove('vinird')
                    if '(' in tokens: tokens.remove('(')
                    if ')' in tokens: tokens.remove(')')

            elif token == 'vend':
                next_token_index = tokens.index(token)+2
                next_token = tokens[next_token_index]
                if next_token in df.columns:
                    token = str(df[next_token].values[ttt_max])
                    tokens[next_token_index] = token
                    tokens.remove('vend')
                    if '(' in tokens: tokens.remove('(')
                    if ')' in tokens: tokens.remove(')')


            elif token == 'mavg':
                next_token = tokens[tokens.index(token)+2]
                N = int(tokens[tokens.index(token)+4])
                if next_token in df.columns:
                    token = f'uniform_filter1d(df["{next_token}"].values, size={N}, mode="nearest")'
                    tokens.remove(next_token)
                    tokens.remove('(')
                    tokens.remove(',')
                    tokens.remove(str(N))
                    tokens.remove(')')
            
            elif token == 'rms':
                next_token = tokens[tokens.index(token)+2]
                N = int(tokens[tokens.index(token)+4])
                if next_token in df.columns:
                    token = f'window_rms(df["{next_token}"].values,{N})'
                    #print(token)
                    tokens.remove(next_token)
                    tokens.remove('(')
                    tokens.remove(',')
                    tokens.remove(str(N))
                    tokens.remove(')')
            elif token == 'ema':
                next_token = tokens[tokens.index(token)+2]
                N = int(tokens[tokens.index(token)+4])
                if next_token in df.columns:
                    token = f'df["{next_token}"].ewm(span={N}).mean()'
                    #print(token)
                    tokens.remove(next_token)
                    tokens.remove('(')
                    tokens.remove(',')
                    tokens.remove(str(N))
                    tokens.remove(')')        
            elif token == 'median':
                next_token = tokens[tokens.index(token)+2]
                N = int(tokens[tokens.index(token)+4])
                if next_token in df.columns:
                    token = f'window_median(df["{next_token}"].values,{N})'
                    #print(token)
                    tokens.remove(next_token)
                    tokens.remove('(')
                    tokens.remove(',')
                    tokens.remove(str(N))
                    tokens.remove(')')    
            elif token == 'round':
                next_token = tokens[tokens.index(token)+2]
                N = int(tokens[tokens.index(token)+4])
                if next_token in df.columns:
                    token = f'np.round(df["{next_token}"].values,{N})'
                    #print(token)
                    tokens.remove(next_token)
                    tokens.remove('(')
                    tokens.remove(',')
                    tokens.remove(str(N))
                    tokens.remove(')')   
            elif token == 'smooth':
                next_token = tokens[tokens.index(token)+2]
                N = float(tokens[tokens.index(token)+4])
                if next_token in df.columns:
                    token = f'smooth(df["{next_token}"].values,{N})'
                    #print(token)
                    tokens.remove(next_token)
                    tokens.remove('(')
                    tokens.remove(',')
                    tokens.remove(str(N))
                    tokens.remove(')')  
            # Append the processed operator to the list of processed tokens
            processed_tokens.append(token)
        elif token.strip().isdigit() or '.' in token:
            # If the token is a number, append it to the list of processed tokens
            processed_tokens.append(token)
        elif token.isspace():
            # If the token is a space, ignore it
            continue
        else:
            # If the token is not a valid operator or a number, assume it's a variable name
            if token in df.columns:
                # If the variable exists in the dataframe, retrieve its data and append it to the list of processed tokens
                processed_tokens.append(f'df["{token}"].values')
            elif str(token).upper() in df.columns:
                processed_tokens.append(f'df["{str(token).upper()}"].values')
            elif "column" in token:
                column_number = int(token.replace("column", "").strip())
                processed_tokens.append(f'df.iloc[:, {column_number}].values')
                #processed_tokens.append(f'df["{token}"].values')
            else:
                # If the variable doesn't exist in the dataframe, print a message and use a default value of 0
                print(f"{color.RED}Variable '{token}' not found in dataframe.{color.WARNING} Using default value of 0.{color.END}")
                processed_tokens.append('np.full_like(tim, 0.0000001111123111)')
                #return np.zeros_like(tim)
                return np.full_like(tim, 0.0000001111123111)

    # Join the processed tokens into a new formula string
    new_formula = ''.join(processed_tokens)

    try:
        # Evaluate the new formula using Python's eval() function
        result = eval(new_formula)
        #print(new_formula)
        #print(result)
    except Exception as e:
        # If an error occurs during evaluation, return "hline"
        result = "hline"
        print("Error evaluating formula: %s" %new_formula)
        print(e,"AAA")
        print()
        #print(result)
    return result

##################
def evaluate_condition(condition_str, df, tim):
    """
    Evaluates a condition string and returns a boolean array where the condition is True.
    
    Supports logical operators: AND, OR, NOT
    Supports comparison operators: >, >=, <, <=, ==, !=
    Supports mathematical operations: +, -, *, /, sqrt, etc.
    
    Example condition:
    OR(AND(signal1*sqrt(3)>1.4, signal2<=2), signal3 + signal2 >= 5)
    """
    # Define a mapping for logical operators and functions with case-insensitive patterns
    replacements = {
        r'\bAND\b': '&',
        r'\bOR\b': '|',
        r'\bNOT\b': '~',
        r'\bSQRT\b': 'np.sqrt',
        # Add more functions here if needed
    }
    
    # Replace logical operators and functions using regex with case-insensitive flags
    for pattern, replacement in replacements.items():
        condition_str = re.sub(pattern, replacement, condition_str, flags=re.IGNORECASE)
    
    # Extract all unique tokens (variables/functions) in the condition
    tokens = set(re.findall(r'\b\w+\b', condition_str))
    
    # Replace signal names with DataFrame columns
    for token in tokens:
        varnotfound = 0
        # Skip if token is a NumPy function or a Python keyword/operator
        if token.lower() in ['and', 'or', 'not', 'sqrt', 'np']:
            continue
        # Check if the token is a number (integer or float)
        if re.fullmatch(r'\d+(\.\d+)?', token):
            # It's a number, skip replacement
            continue
        # Check if the token exists in DataFrame columns (case-sensitive)
        if token in df.columns:
            # Replace whole word matches only
            condition_str = re.sub(r'\b{}\b'.format(re.escape(token)), f'df["{token}"].values', condition_str)
        else:
            # Handle undefined tokens (raise error or continue)
            varnotfound = 1
            print(f"{color.RED}Error: Variable '{color.WARNING}{token}{color.RED}' not found in DataFrame columns.{color.END}")
    
    # Debug: Print the transformed condition string
    #print(f"Transformed condition string: {condition_str}")
    
    try:
        # Evaluate the condition expression safely
        condition_result = eval(condition_str)
        if isinstance(condition_result, np.ndarray):
            # Ensure the result is boolean
            return condition_result.astype(bool)
        else:
            #print("Condition did not evaluate to a boolean array. Using all False.")
            return np.zeros_like(tim, dtype=bool)
    except Exception as e:
        if not varnotfound == 1:
            print(f"{color.RED}Error evaluating condition: {color.WARNING}{condition_str}{color.END}")
            
        print(e)
        return np.zeros_like(tim, dtype=bool)
    
def find_condition_index(condition_result, tim, ttt_min_idx, ttt_max_idx):
    """
    Finds the first index where the condition is True within the specified index range.
    
    Parameters:
        condition_result (np.ndarray): Boolean array where True indicates the condition is met.
        tim (np.ndarray): Array of time values.
        ttt_min_idx (int): Minimum index threshold.
        ttt_max_idx (int): Maximum index threshold.
    
    Returns:
        tuple: (index, time) of the first occurrence where the condition is met within the index range.
    """
    if not isinstance(condition_result, np.ndarray) or condition_result.dtype != bool:
        #print("Invalid condition result. Returning default index.")
        return 0, tim[0]
    
    # Find all indices where the condition is True
    indices = np.where(condition_result)[0]
    #print(f"Found {len(indices)} indices where condition is True.")
    
    # Ensure `tim` is a NumPy array
    if not isinstance(tim, np.ndarray):
        tim = np.array(tim)
        #print("Converted `tim` to NumPy array.")
    
    # Ensure `indices` is a 1D array of integers
    if not isinstance(indices, np.ndarray) or indices.dtype.kind not in {'i', 'u'}:
        indices = np.array(indices, dtype=int)
        #print("Converted `indices` to 1D integer NumPy array.")
    
    #print(f"Shape of tim: {tim.shape}")
    #print(f"Shape of indices: {indices.shape}")
    
    # Validate ttt_min_idx and ttt_max_idx
    if not (0 <= ttt_min_idx < len(tim)):
       # print(f"ttt_min_idx ({ttt_min_idx}) is out of bounds. Using 0.")
        ttt_min_idx = 0
    if not (0 <= ttt_max_idx < len(tim)):
        #print(f"ttt_max_idx ({ttt_max_idx}) is out of bounds. Using last index.")
        ttt_max_idx = len(tim) - 1
    
    #print(f"Filtering indices between {ttt_min_idx} and {ttt_max_idx} (inclusive).")
    
    # Filter indices within the specified index range
    valid_indices = indices[(indices >= ttt_min_idx) & (indices <= ttt_max_idx)]
    #print(f"Found {len(valid_indices)} valid indices within the index range {ttt_min_idx} to {ttt_max_idx}.")
    
    if valid_indices.size > 0:
        # Since indices are sorted, the first one is the earliest occurrence
        first_index = valid_indices[0]
        first_time = tim[first_index]
        #print(f"First condition met at index {first_index}, time {first_time} seconds.")
        return first_index, first_time
    else:
        #print("Condition never met within the specified index range. Using default index.")
        return 0, tim[0]  # Default to the start if condition is never met
    
# Function to check if a string is a condition
def is_condition(commencement_type):
    return commencement_type.lower() not in ["time", "yes" , "no"]

def extract_required_columns(formulas):
    # Define valid tokens that are not columns
    valid_ops = {
        '+', '-', '*', '/', '(', ')', 'sqrt', 'max', 'min', 'vini', 'vinird', 'vend',
        'abs', 'mavg', 'rms', 'ema', 'median', 'smooth', 'round', 'cos', 'sin', 'tan',
        'acos', 'asin', 'atan'
    }
    required = set()
    for formula in formulas:
        # Extract word-like tokens (ignoring punctuation)
        formula = formula.replace(" ","_").replace("&","")
        tokens = re.findall(r'\b[\w\.]+\b', formula)
        for token in tokens:
            # Skip if token is a known operator or a number
            if token in valid_ops:
                continue
            try:
                float(token)
                continue  # token is numeric
            except ValueError:
                pass
            required.add(token)
    return required

def clean_column(col):
    # Apply the same chain of replacements as in your plotting tool
    col = col.replace('"', '')
    col = col.replace("'", '')
    col = col.strip()
    col = col.replace(" ", "_")
    col = col.replace("(", "")
    col = col.replace(")", "")
    col = col.replace(".", "_")
    col = col.replace(":", "_")
    col = col.replace("-", "_")
    col = col.replace("&", "")
    col = col.replace("+", "_")
    col = col.replace("*", "_")
    col = col.replace("/", "_")
    col = col.replace("%", "perc")
    return col

def get_col_mapping(file_path):
    # Read only the header row
    header_df = pd.read_csv(file_path, nrows=0)
    original_cols = header_df.columns.tolist()
    cleaned_cols = [clean_column(c) for c in original_cols]
    
    # Create a mapping: cleaned -> original.
    mapping = {cleaned: orig for cleaned, orig in zip(cleaned_cols, original_cols)}
    # Also include uppercase mapping if needed
    for cleaned, orig in zip(cleaned_cols, original_cols):
        mapping[cleaned.upper()] = orig
    return mapping, cleaned_cols

# Main Script
def main_script(args):
            global text_instances,initime,figurass,axis_size#,resample_to
            
            skip_page = 0
            shorter_x_range = 100000
            do_ORT = ""
            s52513 = 0
            S5255_rec = 0
            text_instances = []
            set_clearing_lines = 0
            counter_points_gral_all = 0
            yoffset = 6
            sig_report = pd.DataFrame()
            
            npdf, sheet_plots, sheet_templates, sheet_files,initime,results_list = args
            if not "_gNone_" in sheet_plots['pdf_name'][npdf]:
                print(f"Plotting row {npdf+2}: {sheet_plots['pdf_name'][npdf]}.pdf")
                
            else:
                return
            npage = sheet_plots["Page"][npdf]
            ngroup = sheet_plots["pdf Group"][npdf]
            ntemplate = sheet_plots["Template Id"][npdf]
            
            template_df = sheet_templates[sheet_templates["Template N"] == ntemplate]
            if template_df.empty and not sheet_files["pdf Group"].empty:
                print(f"ERROR: The template assigned in sheet 'Plots' (row {npdf+2}) does NOT exist in sheet 'Templates'")
                #if 
                return
            
            insert_signals = extract_values_insert_signal(str(sheet_plots["Extras"][npdf]))
            if len(insert_signals) != 0:
                # Create a DataFrame from insert_signals with the specified columns
                insert_signal_columns = ["File Id","ROW","COLUMN","Signal Name","Legend","Y_Label","Style","Color"]
                temp_df = pd.DataFrame(insert_signals, columns=insert_signal_columns)
                
                #print("------------------------------------------")
                #print(temp_df)
                # Get the existing columns from template_df
                existing_columns = template_df.columns
                
                # Reindex temp_df to include all columns in template_df (fill missing columns with NaN)
                temp_df = temp_df.reindex(columns=existing_columns, fill_value=np.nan)

                # Concatenate temp_df with template_df
                template_df = pd.concat([template_df, temp_df], ignore_index=True)
            #print(template_df.iloc[-2].to_string())
            #print(template_df.iloc[-1].to_string())
            #print(template_df)
            template_df['Extras'] = template_df['Extras'].fillna(1)
            template_df['X_Shift'] = template_df['X_Shift'].fillna(0)
            template_df['Y_Shift'] = template_df['Y_Shift'].fillna(0)
            template_df['Y_Label'] = template_df['Y_Label'].fillna(" ")
            template_df['Style'] = template_df['Style'].fillna("-")
            template_df["Legend"] = template_df["Legend"].fillna(template_df["Signal Name"])
            template_df['Margin'] = template_df['Margin'].fillna(0.1)
            template_df["Recovery Time"] = template_df["Recovery Time"].fillna("No")
            template_df["Settling Time"] = template_df["Settling Time"].fillna("No")
            #template_df["Settling Time 2"] = template_df["Settling Time 2"].fillna("No")
            template_df["Rising Time"] = template_df["Rising Time"].fillna("No")
            
            if "Commencement time" not in template_df.columns:
                template_df["Commencement time"] = "No"


            template_df["Commencement time"] = template_df["Commencement time"].fillna("No")
            
            #print(template_df)
            if str(sheet_plots["pdf Group"][npdf]) == "0.0" or str(sheet_plots["pdf Group"][npdf]) == "0":
                return
            #exit()
            try:
                nrows = int(template_df["ROW"].max())
                ncols = int(template_df["COLUMN"].max())
            except Exception as e:
                print(f"ERROR: The template {ntemplate} assigned in sheet 'Plots' (row {npdf+2}) does NOT exist in sheet 'Templates'. Error: {e}")
                #continue
            
            #print(template_df)
            
            if "_Zoomed" in str(sheet_plots["pdf Group"][npdf]):
                file_df = sheet_files[sheet_files["pdf Group"] == sheet_plots["pdf Group"][npdf].replace("_Zoomed","")]
            else:
                file_df = sheet_files[sheet_files["pdf Group"] == sheet_plots["pdf Group"][npdf]]
            #file_df = sheet_files[sheet_files["pdf Group"].str.lower() == sheet_plots["pdf Group"][npdf].str.replace("_Zoomed", "").lower()]
           
                
            file_df = file_df[file_df["Page N"] == npage]
            file_df = file_df.sort_values("Id")
            #template_df = template_df.sort_values("File Id")
            #print(file_df)
            
            file_df.reset_index(drop=True, inplace=True)
            template_df.reset_index(drop=True, inplace=True)
            
            if file_df.empty:
                print(f"ERROR: The template {ntemplate} assigned in sheet 'Plots' (row {npdf+2}) does NOT exist in sheet 'Templates' or doesnt have a page assigned in the sheet 'Files'")
            #print(file_df)
            #if ncols >= nrows:
            #    fig, axs = plt.subplots(nrows, ncols, figsize=(11.75, 7.25),squeeze=0)
            #else:
            #    fig, axs = plt.subplots(nrows, ncols, figsize=(7.25,11.75),squeeze=0)
            if ncols >= 5 or "extend" in str(sheet_plots['Extras'][npdf]).lower():
                figsize = (19, 9)
                
            else:
                figsize = (11, 8.50)
                
            try:
                if "," in str(sheet_plots["X_max"][npdf]) or "," in str(sheet_plots["X_min"][npdf]):
                    fig, axs = plt.subplots(nrows, ncols, figsize=figsize,squeeze=0, sharex='col')
                else:
                    fig, axs = plt.subplots(nrows, ncols, figsize=figsize,squeeze=0, sharex='all')
                
                #fig2 = make_subplots(rows=nrows, cols=ncols, shared_xaxes=True)
            except Exception as e:
                #print("ERROR")
                print(traceback.format_exc())
                print(e)
                return
            
            db_loc = []
            overall_min_max = {}
            margins = {}
            temp_52513 = ""
            temp_52513_titles = "Signal,Page N"
            temp_5255 = ""
            temp_5255_titles = "Signal,Page N"
            

            for nfile in range(len(file_df["Page N"])):
                #if skip_page == 1:
                #    return
                file_path = file_df["loc"][nfile]
                col_mapping, cleaned_cols = get_col_mapping(file_path)
                
                formulas = template_df.loc[template_df["File Id"] == file_df["Id"][nfile], "Signal Name"].tolist()
                required_cleaned = extract_required_columns(formulas)
                # 2. Add the time column if available (using the same cleaning logic)
                if "Time_column" in file_df.columns and pd.notna(file_df["Time_column"][nfile]):
                    try:
                        time_index = int(file_df["Time_column"][nfile])
                        if time_index < len(cleaned_cols):
                            required_cleaned.add(cleaned_cols[time_index])
                    except Exception as e:
                        print(f"Error processing Time_column: {e}")
                else:
                    # Default to the first column from the CSV if "Time_column" doesn't exist or is NaN
                    if cleaned_cols:
                        required_cleaned.add(cleaned_cols[0])
                
                        
                usecols = []
                for col in required_cleaned:
                    if col in col_mapping:
                        usecols.append(col_mapping[col])
                    elif col.upper() in col_mapping:
                        usecols.append(col_mapping[col.upper()])
                
                
                
                usecols = list(set(usecols))  # remove duplicates
                #print(required_cleaned)
                #print(usecols)
                
                #exit()
                #print(usecols)
                try:
                    signals_data = pd.read_csv(file_df["loc"][nfile],usecols=usecols)
                    signals_data.columns = signals_data.columns.str.replace('"', '')
                    signals_data.columns = signals_data.columns.str.replace("'","")
                    signals_data.columns = signals_data.columns.str.strip()
                    signals_data.columns = signals_data.columns.str.replace(" ","_")
                    signals_data.columns = signals_data.columns.str.replace("(","")
                    signals_data.columns = signals_data.columns.str.replace(")","")
                    signals_data.columns = signals_data.columns.str.replace(".","_")
                    signals_data.columns = signals_data.columns.str.replace(":","_")
                    signals_data.columns = signals_data.columns.str.replace("-","_")
                    signals_data.columns = signals_data.columns.str.replace("&","")
                    signals_data.columns = signals_data.columns.str.replace("+","_")
                    signals_data.columns = signals_data.columns.str.replace("-","_")
                    signals_data.columns = signals_data.columns.str.replace("*","_")
                    signals_data.columns = signals_data.columns.str.replace("/","_")
                    signals_data.columns = signals_data.columns.str.replace("%","perc")
                    
                    #signals_data.columns = signals_data.columns.str.strip()
                    #for i in signals_data.columns:
                    #    if "ref" in i.lower():
                    #        print(i)
                    db_loc.append(f"[{str(len(db_loc))}]-"+file_df["loc"][nfile] + " -- " + get_last_modification_date(file_df["loc"][nfile]))
                except Exception as e:
                    print(f"Sheet: 'Files' - Row N{npdf+2} -- {color.RED}{e}{color.END}")
                    db_loc.append(f"[{str(len(db_loc))}]-"+file_df["loc"][nfile] + " -- ERROR NOT FOUND")  
                    for nsignal in range(len(template_df["Template N"])):
                        if template_df["File Id"][nsignal] == file_df["Id"][nfile]:
                            Report_results = {
                            "Group": sheet_plots["pdf Group"][npdf],
                            "Page": sheet_plots["Page"][npdf],
                            "File_id": sheet_plots["Template Id"][npdf],
                            "File_loc": file_df["loc"][nfile],
                            "Fname": os.path.splitext(os.path.basename(file_df["loc"][nfile]))[0],
                            "Signal": template_df["Legend"][nsignal],
                            "Title1": sheet_plots["Title 1"][npdf],
                            "Title2": sheet_plots["Title 2"][npdf],
                            "Title3": sheet_plots["Title 3"][npdf],
                            "Title4": sheet_plots["Title 4"][npdf],
                            "Extra Info": "File NOT FOUND",
                            }
                            #print(template_df)
                            if "Yes" in template_df["Recovery Time"][nsignal]:
                                Report_results["Recovery time"] = f"NOT_FOUND"
                            if "1st" in template_df["Recovery Time"][nsignal]:
                                Report_results["1st Recovery time"] = f"NOT_FOUND"
                                
                    if "skip_missing" in str(sheet_plots["Extras"][npdf]):
                        skip_page = 1
                        print("Skipping page")
                        return
                    continue
                if skip_page == 1:
                    return
                
                
                
                ylims = []
                for nsignal in range(len(template_df["Template N"])):
                    if template_df["File Id"][nsignal] == file_df["Id"][nfile]:
                        try:
                            row = int(template_df["ROW"][nsignal])-1
                            col = int(template_df["COLUMN"][nsignal])-1
                        except: continue

                        
                        
                        Report_results = {
                            "Group": sheet_plots["pdf Group"][npdf],
                            "Page": sheet_plots["Page"][npdf],
                            "File_id": template_df["File Id"][nsignal],
                            "File_loc": file_df["loc"][nfile],
                            "Fname": os.path.splitext(os.path.basename(file_df["loc"][nfile]))[0],
                            "Signal": template_df["Legend"][nsignal],
                            "Title1": sheet_plots["Title 1"][npdf],
                            "Title2": sheet_plots["Title 2"][npdf],
                            "Title3": sheet_plots["Title 3"][npdf],
                            "Title4": sheet_plots["Title 4"][npdf],
                            "Extra Info": "-",
                            }
                        #print(template_df["Scale"][nsignal])
                        template_df["Signal Name"][nsignal] = template_df["Signal Name"][nsignal].strip().replace(" ","_").replace("&","")

                            
                        if "hline" in template_df["Legend"][nsignal] or "vline" in template_df["Legend"][nsignal]:
                            labell = "none"
                        else:
                            labell = template_df["Legend"][nsignal]
                        if "hline" in template_df["Signal Name"][nsignal]:
                            axs[row,col].axhline(y=float(template_df["Y_Shift"][nsignal]), color=template_df["Color"][nsignal], linestyle=template_df["Style"][nsignal], label=labell,linewidth=0.5)
                            continue
                        elif "vline" in template_df["Signal Name"][nsignal]:
                            axs[row,col].axvline(x=float(template_df["X_Shift"][nsignal]), color=template_df["Color"][nsignal], linestyle=template_df["Style"][nsignal], label=labell,linewidth=0.5)
                            continue
                        
                        else:
                            try:
                                if "force_init" in str(sheet_plots["Extras"][npdf]):
                                    tfix = sheet_plots["Settling/Rising Times From"][npdf]
                                    if not tfix == "nan":
                                        tfix,_ = get_index(signals_data.iloc[:, 0],tfix)
                                        signals_data.iloc[:tfix, 1:] = signals_data.iloc[tfix][1:]
                            except Exception as e:
                                print("ERROR: Couldnt force settled init")
                                print(e)
                                
                            try:
                                if "light" in str(sheet_plots["Extras"][npdf]) or "light" in template_df["Extras"][nsignal]:
                                    rasterized = True
                                else:
                                    rasterized = False
                                   
                            except Exception as e:
                                rasterized = False
                                
                            try:
                                points_gral = extract_values_point_gral(str(sheet_plots["Extras"][npdf]))
                            except Exception as e:
                                print("Error: couldnt extract values from Sheet 'Plots' to draw points in signal %s" %str(sheet_plots["Extras"][npdf]))
                            
                            
                                
                            try:
                                points_gral_all = extract_values_point_all(str(sheet_plots["Extras"][npdf]))
                                
                            except Exception as e:
                                print("Error: couldnt extract values from Sheet 'Plots' to draw points in signal %s" %str(sheet_plots["Extras"][npdf]))
                            try:
                                max_gral = extract_max_point_gral(str(sheet_plots["Extras"][npdf]))
                            except Exception as e:
                                print("Error: couldnt extract time range from Sheet 'Plots' to draw max points in signal %s" %str(sheet_plots["Extras"][npdf]))
                            
                            try:
                                data_ort = extract_values_ort_gral(str(sheet_plots["Extras"][npdf]))[0]
                                #print(data_ort)
                            except Exception as e:
                                data_ort = [0]
                                #print("Error: couldnt extract values from Sheet 'Plots' to assess Oscillation Rej Test %s" %str(sheet_plots["Extras"][npdf]))
                                pass
                            
                            try:
                                force_ncols = extract_ncols(str(sheet_plots["Extras"][npdf]))[0]
                                
                                if force_ncols == []:
                                    force_ncols = None
                                #print(force_ncols)
                            except Exception as e:
                                force_ncols = None
                                #print("Error: couldnt force the number  of columns for that page %s" %str(sheet_plots["Extras"][npdf]))
                            try:
                                legend_size = extract_legendsize(str(sheet_plots["Extras"][npdf]))[0]
                                
                                if legend_size == []:
                                    legend_size = 5.5
                                #print(legend_size)
                            except Exception as e:
                                legend_size = 5.5
                                #print("Error: couldnt force the number  of columns for that page %s" %str(sheet_plots["Extras"][npdf]))
                            try:
                                axis_size = extract_axissize(str(sheet_plots["Extras"][npdf]))[0]
                                
                                if axis_size == []:
                                    axis_size = 6.5
                                #print(axis_size)
                            except Exception as e:
                                axis_size = 6.5
                            try:
                                delay_rise = extract_delay_rise(str(sheet_plots["Extras"][npdf]))[-1]
                                
                                if delay_rise == []:
                                    delay_rise = 0.00000000001
                                #print(delay_rise)
                            except Exception as e:
                                delay_rise = 0.00000000001

                            try:
                                avg_rise = extract_avg_rise(str(sheet_plots["Extras"][npdf]))[-1]
                                
                                if avg_rise == []:
                                    avg_rise = 50
                                #print(avg_rise)
                            except Exception as e:
                                avg_rise = 50

                            try:
                                hide_legends = extract_hide_legend(str(sheet_plots["Extras"][npdf]))
                                #print(hide_legends)
                            except Exception as e:
                                 print("ERROR: Could not hide legends: %s" %e)
                            try:
                                if pd.notna(file_df["Time_column"][nfile]):
                                    timecol = int(file_df["Time_column"][nfile])
                                    print("Using Time_column %s" % timecol)
                                else:
                                    #print("Time_column is empty, using default value 0")
                                    timecol = 0
                            except:
                                timecol = 0
                                
                            tim = signals_data[signals_data.columns[timecol]]+template_df["X_Shift"][nsignal]
                            try:
                                if "xzoom_shorter" in str(sheet_plots["Extras"][npdf]):
                                    ttt_max,_ = get_index(tim,1000,1)
                                    if _ < shorter_x_range:
                                        shorter_x_range = _
                            except: pass
                            try:
                                shiftx = extract_values_shift(str(sheet_plots["Extras"][npdf]))
                                #print(shiftx)
                                if len(shiftx) != 0:
                                    for i in shiftx:
                                        #shiftx = shiftx[-1]
                                        if float(i[0]) == float(template_df["File Id"][nsignal]):
                                            #print(i)
                                            tim += i[1]
                                        
                            except Exception as e:
                                
                                print("Error: couldnt shiftx value: shiftx(file_id,xshift): %s - %s" %(str(sheet_plots["Extras"][npdf]),e))
                                
                            
                            #print(template_df["X_Shift"][nsignal])
                            try:
                                ttt_min,_ = get_index(tim,sheet_plots["X_min"][npdf])
                            except Exception as e:
                                print(e)
                                try:
                                    ttt_min,_ = get_index(tim,0.01)
                                except:
                                    print(f"{color.FAIL} Please check CSV Files, seems like some are empty \n({db_loc[-1]}){color.END}")
                                    break
                                    #exit()
                            try:
                                ttt_max,_ = get_index(tim,sheet_plots["X_max"][npdf],1)
                            except:
                                ttt_max,_ = get_index(tim,1000,1)
                                
                            try:
                                #sig = signals_data[template_df["Signal Name"][nsignal]]*float(template_df["Scale"][nsignal])+float(template_df["Y_Shift"][nsignal])
                                sig = evaluate_formula(template_df["Signal Name"][nsignal],signals_data,tim,ttt_min,ttt_max)+float(template_df["Y_Shift"][nsignal])
                                if isinstance(sig, float):
                                        
                                        sig = np.full_like(tim, sig)
                                mask = np.vectorize(lambda x: isinstance(x, str))(sig)
                                sig[mask] = 0.0
                                
                                
                            except Exception as e:
                                print(str(e), "error code: EEE")
                                
                                if sheet_plots["print fnames"][npdf] == "Yes":
                                    llabl = template_df["Legend"][nsignal] + f" [{len(db_loc)-1}]"
                                else:
                                    llabl = template_df["Legend"][nsignal]
                                axs[row,col].axhline(y=template_df["Signal Name"][nsignal]+" [ERROR!]",linestyle=":",linewidth=line_size,label=llabl)
                                continue
                            
                            try:
                                shifty_all = extract_values_shift_y(str(sheet_plots["Extras"][npdf]))
                                #print(shifty)
                                if len(shifty_all) != 0:
                                    for shifty in shifty_all:
                                        #shifty = shifty[-1]
                                        if float(shifty[0]) == float(template_df["File Id"][nsignal]) and str(shifty[1]).strip() in str(template_df["Legend"][nsignal]):
                                            print(f"Shifting signal {shifty[1]} of file id {shifty[0]} by {shifty[2]}")
                                            sig += shifty[2]
                                        
                            except Exception as e:

                                print("Error: couldnt shifty value [yshift(file_id,yshift)] %s - %s" %(str(sheet_plots["Extras"][npdf]),e))
                            #L = axs[row,col].legend(bbox_to_anchor=(0,1.02,1,0.2), loc="lower left", mode="None", borderaxespad=0, frameon=False, fontsize=8)
                            axs[row,col].set_xlabel(sheet_plots["X_Label"][npdf],rotation = 0, fontname="Times New Roman", fontweight="ultralight", fontsize=6)
                            if not "axis" in str(template_df['Extras'][nsignal]):
                                axs[row,col].set_ylabel(template_df["Y_Label"][nsignal], rotation = 90, fontname="Times New Roman", fontweight="ultralight", fontsize=6.5)
                                
                                
                            try:
                                # Update overall min and max values for the subplot
                                if not "axis" in str(template_df['Extras'][nsignal]):
                                    if (row, col) not in overall_min_max:
                                        overall_min_max[(row, col)] = [sig[ttt_min:ttt_max].min(), sig[ttt_min:ttt_max].max()]
                                        margins[(row,col)] = template_df["Margin"][nsignal]
                                    else:
                                        overall_min_max[(row, col)][0] = min(overall_min_max[(row, col)][0], sig[ttt_min:ttt_max].min())
                                        overall_min_max[(row, col)][1] = max(overall_min_max[(row, col)][1], sig[ttt_min:ttt_max].max())
                                        margins[(row,col)] = template_df["Margin"][nsignal]
                                #axs[row,col].set_ylim([template_df["Y_min"][nsignal],template_df["Y_max"][nsignal]])
                            except Exception as e:
                                print(e)
                                pass
                            
                            if sheet_plots["print fnames"][npdf] == "Yes":
                                llabl = template_df["Legend"][nsignal] + f" [{len(db_loc)-1}]"
                            else:
                                llabl = template_df["Legend"][nsignal]
                                
                            #print(template_df)
                            #print(f"Plotting signal: {template_df['Signal Name'][nsignal]}")
                            #print(template_df['Extras'][nsignal])
                            if np.all(sig == 0.0000001111123111):
                                llabl += " [ERROR]"
                                
                                axs[row,col].plot(tim,sig,label = llabl,linestyle=template_df["Style"][nsignal],color=template_df["Color"][nsignal],linewidth=0.01)
                                
                                    
                            
                                
                            elif "axis" in str(template_df['Extras'][nsignal]):
                                template_df["Margin"][nsignal] = 0.1
                                #print("AAAA")
                                axs2 = axs[row,col].twinx()
                                axs2min = min(sig[ttt_min:ttt_max])
                                axs2max = max(sig[ttt_min:ttt_max])
                                try:
                                    axs2._get_lines.prop_cycler = axs[row,col]._get_lines.prop_cycler
                                except:
                                    pass
                                axs2.plot(tim,sig,linestyle=template_df["Style"][nsignal],color=template_df["Color"][nsignal],linewidth=line_size,rasterized=rasterized)
                                #axs2.set_ylim([min(sig[ttt_min:ttt_max]),max(sig[ttt_min:ttt_max])+(max(sig[ttt_min:ttt_max])-min(sig[ttt_min:ttt_max]))*0.1])
                                #try:
                                axs2.set_ylim([axs2min-(axs2max-axs2min)*(0.1+template_df["Margin"][nsignal]/100),axs2max+(axs2max-axs2min)*(0.1+template_df["Margin"][nsignal]/100)])
                                #except:pass
                                axs2.set_ylabel(template_df["Y_Label"][nsignal], rotation = 90, fontname="Times New Roman", fontweight="ultralight", fontsize=6.5)
                                plt.setp(axs2.yaxis.get_ticklabels(), rotation = 0, fontname="Times New Roman", fontweight="ultralight", fontsize=axis_size)
                                try:
                                    
                                    axs2.tick_params(colors=axs2.lines[-1].get_color(), which='both')
                                    axs[row, col].plot([], [], label=llabl, linestyle=template_df["Style"][nsignal], color=template_df["Color"][nsignal], linewidth=line_size, rasterized=rasterized)
                                except:

                                    try:
                                        handles2, labels2 = axs2.get_legend_handles_labels()
                                        axs2.legend(handles2, labels2,loc='lower right', bbox_to_anchor=(0,1.02,1,0.2), 
                                                    fancybox=True, shadow=True, mode="None", borderaxespad=0, frameon=False, fontsize=legend_size,ncols = ncol_legends)
                                    except Exception as e:
                                        pass
                                n_yticks = len(axs[row, col].get_yticks())
                                axs2.locator_params(axis='y', nbins=n_yticks-1)
                            else:
                                linestyle_tuple = [
                                ('loosely dotted',        (0, (1, 10))),
                                ('dotted',                (0, (1, 1))),
                                
                                ('densely dotted',        (0, (1, 1))),
                                ('long dash with offset', (5, (10, 3))),
                                ('loosely dashed',        (0, (5, 10))),
                                ('dashed',                (0, (5, 5))),
                                
                                ('densely dashed',        (0, (5, 1))),

                                ('loosely dashdotted',    (0, (3, 10, 1, 10))),
                                ('dashdotted',            (0, (3, 5, 1, 5))),
                                ('densely dashdotted',    (0, (3, 1, 1, 1))),

                                ('dashdotdotted',         (0, (3, 5, 1, 5, 1, 5))),
                                ('loosely dashdotdotted', (0, (3, 10, 1, 10, 1, 10))),
                                ('densely dashdotdotted', (0, (3, 1, 1, 1, 1, 1)))]
                                # Create a dictionary from linestyle_tuple
                                linestyle_dict = dict(linestyle_tuple)
                                
                                try:
                                # Use the dictionary to get the linestyle from the name
                                    #print(llabl)
                                    linestyle = linestyle_dict[template_df["Style"][nsignal]]                                
                                    axs[row, col].plot(tim, sig, label=llabl, linestyle=linestyle, color=template_df["Color"][nsignal], linewidth=line_size, rasterized=rasterized)
                                    
                                except:
                                    try:
                                        axs[row, col].plot(tim, sig, label=llabl, linestyle=template_df["Style"][nsignal], color=template_df["Color"][nsignal], linewidth=line_size, rasterized=rasterized)
                                    except Exception as e:
                                        if "could not convert string to float" in str(e):
                                            #sig = np.zeros_like(tim)
                                            print(f"{color.RED} \nERROR in {color.WARNING}Page: {sheet_plots['Page'][npdf]} - PDF: {sheet_plots['pdf Group'][npdf]} - Signal: '{template_df['Legend'][nsignal]}'{color.RED}>>>>>>>>>>>>>>>>> Something is wrong with the file or signal, maybe sim crashed\nErr: {e}\n{color.END}")
                                            return
                                            break
                                        else:
                                            print(f"{color.RED} \nERROR in {color.WARNING}Page: {sheet_plots['Page'][npdf]} - PDF: {sheet_plots['pdf Group'][npdf]} - Signal: '{template_df['Legend'][nsignal]}'{color.RED}>>>>>>>>>>>>>>>>> Something is wrong with the file or signal\nErr: {e}\n{color.END}")
                                            

                            band_value = None
                            if "errorbands(" in str(template_df['Extras'][nsignal]):
                                # "errorbands(MVA)" branch: use 2% of the provided MVA value.
                                try:
                                    mva_str = str(template_df['Extras'][nsignal]).split("errorbands(")[1].split(")")[0]
                                    mva_value = float(mva_str)
                                    # Divide by 2 to apply half the band above and half below the signal.
                                    band_value = mva_value * 0.02 / 2
                                    labelerrorbands = "Errbands 2% of Plant MVA"
                                except Exception as e:
                                    print("Error parsing MVA value for errorbands:", e)
                            elif "errorbands" in str(template_df['Extras'][nsignal]):
                                # "errorbands" alone: calculate 10% of the change in the signal.
                                vini = sig[ttt_min]
                                maxval = sig[ttt_min:ttt_max].max()
                                minval = sig[ttt_min:ttt_max].min()
                                maxdelta = max(abs(maxval - vini), abs(minval - vini))
                                band_value = maxdelta * 0.1
                                labelerrorbands = "Errbands 10% of change"

                            if band_value is not None:
                                # Generate a higher resolution time base for a smoother fill
                                reduced_tim = generate_linspace(list(tim), 0.01)
                                indices = np.searchsorted(tim, reduced_tim)
                                reduced_sig = sig[indices]

                                # Set upper and lower error bands
                                b1 = band_value
                                b2 = -band_value

                                # Plot the error bands on the appropriate subplot.
                                errb = axs[row, col].fill_between(
                                    reduced_tim, reduced_sig + b1, reduced_sig + b2,
                                    color="silver", alpha=0.4, label=labelerrorbands
                                )

                                
                            if "maxhide" in str(template_df['Extras'][nsignal]):
                                maxval = sig[ttt_min:ttt_max].max()
                                maxval_index = np.argmax(sig[ttt_min:ttt_max]) + ttt_min
                                x = tim[maxval_index]
                                x = tim[ttt_min:ttt_max][maxval_index]
                                y = maxval
                                #axs[row,col].axhline(y=y,color = axs[row,col].lines[-1].get_color(),linestyle="--", label=f"Max({x:.3f}s) = {y:.3f}",linewidth=line_size)
                                save_string_to_file(f"Max Value signal,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{x},y:,{y}",working_dir + os.sep+ output_folder +os.sep+ f"p1_gTot_{initime}_tTot.csv")
                                
                            elif "maxp" in str(template_df['Extras'][nsignal]) or "maxpoint" in str(template_df['Extras'][nsignal]):
                                maxval = sig[ttt_min:ttt_max].max()
                                maxval_index = np.argmax(sig[ttt_min:ttt_max]) + ttt_min
                                x = tim[maxval_index]
                                y = maxval
                                axs[row,col].scatter(x, y, marker='x',s=15, color=axs[row,col].lines[-1].get_color(),linewidth=line_size)
                                
                                y_range = axs[row, col].get_ylim()[1] - axs[row, col].get_ylim()[0]
                                offset = y_range * 0.05  # Adjust the multiplier as needed
                                axs[row, col].annotate(f"{y:.3f}", (x, y), textcoords="offset points", xytext=(0, offset), ha='center', fontsize=legend_size, color=axs[row, col].lines[-1].get_color())
                                save_string_to_file(f"Max Value signal,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{x},y:,{y}",working_dir + os.sep+ output_folder +os.sep+ f"p1_gTot_{initime}_tTot.csv")
                                
                            elif ("max" in str(template_df['Extras'][nsignal]) or "maxline" in str(template_df['Extras'][nsignal])) and not "maxp" in str(template_df['Extras'][nsignal]) and not "maxhide" in str(template_df['Extras'][nsignal]):
                                maxval = sig[ttt_min:ttt_max].max()
                                maxval_index = np.argmax(sig[ttt_min:ttt_max]) + ttt_min
                                x = tim[maxval_index]
                                x = tim[ttt_min:ttt_max][maxval_index]
                                y = maxval
                                axs[row,col].axhline(y=y,color = axs[row,col].lines[-1].get_color(),linestyle="--", label=f"Max({x:.3f}s) = {y:.3f}",linewidth=line_size)
                                save_string_to_file(f"Max Value signal,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{x},y:,{y}",working_dir + os.sep+ output_folder +os.sep+ f"p1_gTot_{initime}_tTot.csv")
                                
                            
                                
                            
                                
                                
                            if "min" in str(template_df['Extras'][nsignal]) or "minline" in str(template_df['Extras'][nsignal]):
                                maxval = sig[ttt_min:ttt_max].min()
                                maxval_index = np.argmin(sig[ttt_min:ttt_max]) + ttt_min
                                x = tim[maxval_index]
                                x = tim[ttt_min:ttt_max][maxval_index]
                                y = maxval
                                axs[row,col].axhline(y=y,color = axs[row,col].lines[-1].get_color(),linestyle="--", label=f"Min({x:.3f}s) = {y:.3f}",linewidth=line_size)
                                save_string_to_file(f"Min Value signal,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{x},y:,{y}",working_dir + os.sep+ output_folder +os.sep+ f"p1_gTot_{initime}_tTot.csv")
                                            
                            if "minp" in str(template_df['Extras'][nsignal]) or "minpoint" in str(template_df['Extras'][nsignal]):
                                maxval = sig[ttt_min:ttt_max].max()
                                maxval_index = np.argmin(sig[ttt_min:ttt_max]) + ttt_min
                                x = tim[maxval_index]
                                y = maxval
                                axs[row,col].scatter(x, y, marker='x',s=15, color=axs[row,col].lines[-1].get_color(),linewidth=line_size)
                                y_range = axs[row, col].get_ylim()[1] - axs[row, col].get_ylim()[0]
                                offset = y_range * 0.05  # Adjust the multiplier as needed
                                axs[row, col].annotate(f"{y:.3f}", (x, y), textcoords="offset points", xytext=(0, -offset), ha='center', fontsize=legend_size, color=axs[row, col].lines[-1].get_color())
                                save_string_to_file(f"Min Value signal,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{x},y:,{y}",working_dir + os.sep+ output_folder +os.sep+ f"p1_gTot_{initime}_tTot.csv")
                            
                                
                            try:
                                if len(points_gral) != 0:
                                    for ppoint in points_gral:
                                        if ppoint[0] == str(template_df["Legend"][nsignal]).strip():
                                            ttt_point,_ = get_index(tim,ppoint[1])
                                            x = tim[ttt_point]
                                            y = sig[ttt_point]
                                            axs[row,col].scatter(x, y, marker='x',s=15, color=axs[row,col].lines[-1].get_color(),linewidth=line_size)
                                            #axs[row,col].annotate(f"{y:.3f}", (x, y), textcoords="offset points", xytext=(0,-3), ha='center', fontsize=7,color="k")
                                            y_range = axs[row, col].get_ylim()[1] - axs[row, col].get_ylim()[0]
                                            x_range = axs[row, col].get_xlim()[1] - axs[row, col].get_xlim()[0]
                                            offset = min(y_range * 0.05, 10)  # Adjust the multiplier as needed and set a maximum limit
                                            offsetx = min(x_range * 0.5, 10)
                                            axs[row, col].annotate(f"{y:.{ppoint[2]}f}", (x, y), textcoords="offset points", xytext=(+offsetx, -offset), ha='center', fontsize=6, color=axs[row, col].lines[-1].get_color())
                                            #save_string_to_file(f"{sheet_plots['pdf Group'][npdf]}")
                                            save_string_to_file(f"Value signal,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{x},y:,{y}",working_dir + os.sep+ output_folder +os.sep+ f"p1_gTot_{initime}_tTot.csv")
                                            #save_string_to_file(f"Rising time,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{tmin},{tim[indt5]},{delta}",working_dir + os.sep+ output_folder +os.sep+ f"p1_gTot_{initime}_tTot.csv")
                                
                            except Exception as e:
                                #print("Problem drawing points from extra: %s" %e)
                                pass
                                
                            try:
                                #print(points_gral_all)
                                if len(points_gral_all) != 0:
                                    #counter_points_gral_all += 1
                                    #yoffset *= -1
                                    for ppoint in points_gral_all:
                                            #counter_points_gral_all += 1
                                            if sheet_plots["X_max"][npdf] is None: 
                                                sheet_plots["X_max"][npdf] = get_index(tim,1000,1)[1]
                                            if  ppoint > sheet_plots["X_max"][npdf] or ppoint < sheet_plots["X_min"][npdf]: 
                                                continue
                                            ttt_point,_ = get_index(tim,ppoint)
                                            x = tim[ttt_point]
                                            y = sig[ttt_point]
                                            axs[row,col].scatter(x, y, marker='x',s=15, color=axs[row,col].lines[-1].get_color(),linewidth=line_size)
                                            yoffset = 5
                                            if template_df["File Id"][nsignal] in {1, 3, 5, 7}:
                                                yoffset = yoffset
                                            else:
                                                yoffset = -yoffset*1.3
                                        
                                            #print(ppoint)
                                            if abs(y) == 0:
                                                formatted_y = f"{y:.0f}"
                                            elif abs(y) < 1.35:
                                                formatted_y = f"{y:.3f}"
                                            elif 1.35 <= abs(y) < 10:
                                                formatted_y = f"{y:.2f}"
                                            else:
                                                formatted_y = f"{y:.1f}"
                                                
                                            axs[row, col].annotate(formatted_y, (x, y), textcoords="offset points", xytext=(0, yoffset), ha='center', fontsize=6, color=axs[row, col].lines[-1].get_color())
                                            save_string_to_file(f"Value signal,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{x},y:,{y}",working_dir + os.sep+ output_folder +os.sep+ f"p1_gTot_{initime}_tTot.csv")
                                            
                            except Exception as e:
                                print("Problem drawing points_all from extra: %s" %traceback.format_exc())
                                pass
                            #print(axs2)
                            try: #ORT ASSESSMENT
                                #if data_ort != [0]:
                                #    print(data_ort)
                                #print(len(data_ort))
                                
                                if len(data_ort) == 6:
                                    #print(data_ort)
                                    Qposition = (0,0)
                                    Vposition = (0,0)
                                    
                                    if  data_ort[0] == str(template_df["Legend"][nsignal]).strip():
                                        Vsig_grid_ort = sig.copy()
                                        do_ORT += "Vgridsig"
                                        
                                    if  data_ort[1] == str(template_df["Legend"][nsignal]).strip() and not "Vsig" in do_ORT:
                                        Vsig_ort = sig.copy()
                                        do_ORT += "Vsig"
                                        Vposition = ((template_df["ROW"][nsignal]),(template_df["COLUMN"][nsignal]))
                                        
                                        
                                    if  data_ort[2] == str(template_df["Legend"][nsignal]).strip() and not "Qsig" in do_ORT:
                                        Qsig_ort = sig.copy()
                                        Qposition = ((template_df["ROW"][nsignal]),(template_df["COLUMN"][nsignal]))
                                        do_ORT += "Qsig"
                                    
                                    Fsig_ort = data_ort[3]
                                    cycle_time = 1/Fsig_ort
                                    
                                    cycle_asses_tol = 1.5 #(This should grab a full cycle of Q and V to work correctly)
                                    
                                    
                                    
                                    if "Vgridsig" in do_ORT and "Vsig" in do_ORT and "Qsig" in do_ORT and not "ORT DONE" in do_ORT:
                                        if data_ort[4] == "def":
                                            #data_ort[4] = 10
                                            print(f"{color.WARNING} ORT Assessment: Using tmin = 10s {color.END}")
                                            x_min_ort = 10
                                        else:
                                            x_min_ort = data_ort[4]
                                            
                                        if data_ort[5] == "def":
                                            #data_ort[5] = 29
                                            x_max_ort = 29
                                            #print(f"{color.WARNING} ORT Assessment: Using tmax = 29s {color.END}")
                                        else:
                                            x_max_ort = data_ort[5]
                                        #Vsig_ort_filt = filter_frequency(Vsig_ort,tim,Fsig_ort)
                                        #print(Vsig_ort_filt)
                                        #axs[0, 0].plot(tim,Vsig_ort_filt,color="k",linewidth=line_size, linestyle='solid',label=f"Filtered {Fsig_ort}")

                                        #do_ORT = 4
                                        #x_min_ort = float(sheet_plots["X_min"][npdf])
                                        #x_max_ort = float(sheet_plots["X_max"][npdf])
                                        #print(x_min_ort,x_max_ort)
                                        try:
                                            ind_tmin,tmin = get_index(tim,x_min_ort)
                                        except:
                                            ind_tmin= ttt_min
                                            tmin = tim[ind_tmin]
                                        try:
                                            ind_tmax,tmax = get_index(tim,x_max_ort,1) 
                                        except: 
                                            ind_tmax= ttt_max
                                            tmax = tim[ind_tmax]
                                            #ind_tmax,tmax = get_index(tim,ttt_max)
                                        if cycle_time*cycle_asses_tol >= tmax-tmin:
                                            print(f"{color.RED}ORT ERROR: The selected X range is to small to assess ORT on these signals. {color.WARNING}F={Fsig_ort}Hz --> t={cycle_asses_tol}*1/F={cycle_asses_tol/Fsig_ort}s in Page {npage} {color.END}")
                                            raise Exception("Please Increase the xrange, usually 8 to 29 works fine")
                                        
                                        ind_tmin_asses,tmin_asses = get_index(tim,tmax-cycle_time*cycle_asses_tol)
                                        
                                        max_index_Vsig = np.argmax(Vsig_ort[ind_tmin_asses:ind_tmax])+ind_tmin_asses
                                        min_index_Vsig = np.argmin(Vsig_ort[ind_tmin_asses:ind_tmax])+ind_tmin_asses
                                        max_index_Vsig_grid = np.argmax(Vsig_grid_ort[ind_tmin_asses:ind_tmax])+ind_tmin_asses
                                        min_index_Vsig_grid = np.argmin(Vsig_grid_ort[ind_tmin_asses:ind_tmax])+ind_tmin_asses
                                        
                                        max_index_Qsig = np.argmax(Qsig_ort[ind_tmin_asses:ind_tmax])+ind_tmin_asses
                                        #min_index_Qsig = np.argmin(Qsig_ort[ind_tmin_asses:ind_tmax])+ind_tmin_asses
                                        
                                        
                                        #axs[row,col].scatter(tim[max_index_Vsig],Vsig_ort[max_index_Vsig], marker='x',s=15, color="k",linewidth=line_size)
                                        #axs[row,col].scatter(tim[min_index_Vsig],Vsig_ort[min_index_Vsig], marker='x',s=15, color="k",linewidth=line_size)
                                        #print(tim[max_index_Vsig],Vsig_ort[max_index_Vsig])                                                                                                    
                                        #ind_tmin_asses,tmin_asses = get_index(tim,tmax-cycle_time*1.5)
                                        
                                        
                                        #print(tmin_asses,tim[min_index_Vsig],tim[max_index_Vsig])
                                        
                                        Vgain = (Vsig_ort[max_index_Vsig] - Vsig_ort[min_index_Vsig])/(Vsig_grid_ort[max_index_Vsig_grid] - Vsig_grid_ort[min_index_Vsig_grid])
                                        
                                        #VPht = tim[max_index_Vsig] #Set as reference
                                        PhDiff = ((tim[max_index_Qsig]-tim[max_index_Vsig])/cycle_time)*360 #degrees 
                                        if PhDiff < 0:
                                            PhDiff += 360
                                        if PhDiff < 0:
                                            PhDiff += 360
                                        if PhDiff > 360:
                                            PhDiff -= 360
                                        #print(Vgain)
                                        #if QPh >
                                        #awdawd
                                        axs[int(Vposition[0]), int(Vposition[1])].plot([0],[0],color="lime",linewidth=5, linestyle='solid',label=f"V Gain = {Vgain:.2f}pu")
                                        axs[int(Qposition[0])-1, int(Qposition[1])-1].plot([0],[0],color="gray",linewidth=5, linestyle='solid',label=f"Phase Diff = {int(PhDiff)}deg")
                                        do_ORT += "ORT DONE"
                                        #save_string_to_file(f"ORT_Vgain,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{tmin_asses},Vgain:,{Vgain}",working_dir + os.sep+ output_folder +os.sep+ f"p1_gTot_{initime}_tTot.csv")
                                        #save_string_to_file(f"ORT_PhaseDiff,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{tmin_asses},PhaseDiff:,{PhDiff}",working_dir + os.sep+ output_folder +os.sep+ f"p1_gTot_{initime}_tTot.csv")
                                        #save_string_to_file(f"ORT_Assessment,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},Frec,{Fsig_ort},Vgain:,{Vgain},Phase Diff:,{PhDiff}",working_dir + os.sep+ output_folder +os.sep+ f"ORT_{sheet_plots['pdf_name'][npdf]}.txt")
                                        Report_results["ORT Freq"] = f"{Fsig_ort:.1f}"
                                        Report_results["ORT Gain"] = f"{Vgain:.3f}"
                                        Report_results["ORT Phase Err"] = f"{PhDiff:.0f}"
                                        Report_results["Signal"] = f"ORT_Results"
                                        results_list.append(Report_results)
                                        #print(Report_results)
                                        sheet_plots["X_min zoom"][npdf] = tmax-1/Fsig_ort*1.52
                                        sheet_plots["X_max zoom"][npdf] = tmax
                            except Exception as e:
                                print(f"Problem assessing ORT in page {npage}  - ERROR:  {traceback.format_exc()}")
                                pass
                            
                            try:
                                
                                if len(max_gral) != 0:
                                    #print(max_gral)
                                    for ppoint in max_gral:
                                        print(ppoint)
                                        if ppoint[0] == str(template_df["Legend"][nsignal]).strip():
                                            ttt_ini_min,_ = get_index(tim,ppoint[1])
                                            ttt_ini_max,_ = get_index(tim,ppoint[2])
                                            max_index = np.argmax(sig[ttt_ini_min:ttt_ini_max])+ttt_ini_min
                                            x = tim[max_index]
                                            y = sig[max_index]
                                            print(x,y)
                                            axs[row,col].scatter(x, y, marker='x',s=15, color=axs[row,col].lines[-1].get_color(),linewidth=line_size)
                                            #axs[row,col].annotate(f"{y:.3f}", (x, y), textcoords="offset points", xytext=(0,-3), ha='center', fontsize=7,color="k")
                                            y_range = axs[row, col].get_ylim()[1] - axs[row, col].get_ylim()[0]
                                            x_range = axs[row, col].get_xlim()[1] - axs[row, col].get_xlim()[0]
                                            offset = min(y_range * 0.05, 10)  # Adjust the multiplier as needed and set a maximum limit
                                            offsetx = min(x_range * 0.5, 10)
                                            axs[row, col].annotate(f"max: {y:.{ppoint[3]}f}", (x, y), textcoords="offset points", xytext=(+offsetx, -offset), ha='center', fontsize=6, color=axs[row, col].lines[-1].get_color())
                                            #save_string_to_file(f"{sheet_plots['pdf Group'][npdf]}")
                                            save_string_to_file(f"Max signal,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},t:{ppoint[1]}:{ppoint[2]},y:,{y}",working_dir + os.sep+ output_folder +os.sep+ f"p1_gTot_{initime}_tTot.csv")
                                    
                            except Exception as e:
                                print("Problem drawing points from extra: %s" %e)
                                
                            try:
                                points = extract_values_point(str(template_df['Extras'][nsignal]))
                                
                                if len(points) != 0:
                                    for ppoint in points:
                                        ttt_point,_ = get_index(tim,ppoint)
                                        x = tim[ttt_point]
                                        y = sig[ttt_point]
                                        axs[row,col].scatter(x, y, marker='x',s=15, color=axs[row,col].lines[-1].get_color(),linewidth=line_size)
                                        
                                        # Adjust offset based on the total number of subplot rows
                                        offset = 0.2 / axs.shape[0]
                                        
                                        # Adjust offsetx based on the total number of subplot columns
                                        offsetx = 0 / axs.shape[1]
                                        
                                        # Transform offset from axes coordinates to data coordinates
                                        inv = axs[row, col].transData.inverted()
                                        dx, dy = inv.transform(axs[row, col].transAxes.transform((offsetx, offset))) - inv.transform(axs[row, col].transAxes.transform((0, 0)))
                                        
                                        # Determine whether to place the annotation above or below the point
                                        subymax = axs[row, col].get_ylim()[1]
                                        subymin = axs[row, col].get_ylim()[0]
                                        
                                        y_annot = y+dy
                                        # while y_annot >= subymax:
                                        #     y_annot -= dy/10
                                        # while y_annot <= subymin:
                                        #     y_annot += dy/10
                                        
                                            
                                        if y > (subymax + subymin) / 2:
                                            y_annot += -dy * 4  # Place the annotation below the point
                                        
                                        axs[row, col].annotate(f"{y:.3f}", (x, y), textcoords="data", xytext=(x + dx, y_annot), ha='center', fontsize=5, color=axs[row, col].lines[-1].get_color())
                                        
                            except Exception as e:
                                print("Problem drawing points from extra: %s" %e)

                            
                            try:
                                points = extract_values_line(str(template_df['Extras'][nsignal]))
                                
                                if len(points) != 0:
                                    for ppoint in points:
                                        #print(ppoint)
                                        ttt_point,_ = get_index(tim,ppoint)
                                        x = tim[ttt_point]
                                        y = sig[ttt_point]
                                        axs[row,col].axvline(x=x,color = axs[row,col].lines[-1].get_color(),linestyle="--", label=f"y({x:.3f}s) = {y:.3f}",linewidth=line_size)
                                
    
                                        
                            except Exception as e:
                                print("Problem drawing lines from extra: %s" %e)
                                    
                            #plt.draw()
                            #fig2.add_trace(go.Scatter(x=tim, y=sig,name=llabl), row=row+1, col=col+1)
                            #print(ax.xaxis.get_ticklabels())
                            if "gain" in str(template_df['Extras'][nsignal]):
                                gaindiv = str(str(template_df['Extras'][nsignal]).split("gain(")[1]).split(")")[0]
                                gaindiv = float(gaindiv)
                                #print(max(sig),min(sig))
                                sig_ngain = sig[ttt_min:ttt_max] - np.mean(sig[ttt_min:ttt_max])
                                gain = (sig_ngain.max() - sig_ngain.min())/gaindiv
                                
                                sig_ngain = None
                                axs[row,col].axvline(x=-1000,color = "gray",linestyle="solid", label=f"Gain = {gain:.3f}",linewidth=5)
                                
                            resample_to = file_df["Resample_to"][nfile]

                            ########################################
                            ############## SETTLING TIMES ##########
                            report_52513 = []
                           
                            if "Yes" in template_df["Settling Time"][nsignal] or "Yes" in template_df["Recovery Time"][nsignal] or "1st" in template_df["Recovery Time"][nsignal] or "5258" in template_df["Recovery Time"][nsignal]  or "Yes" in template_df["Rising Time"][nsignal] or template_df["Commencement time"][nsignal] != "No":
                                if resample_to != 0.123456:
                                    tim,sig= increase_sampling(sig,tim,resample_to)    
                                else:
                                    tim,sig = list(tim),list(sig)
                                
                                
                                #if sheet_plots["SetRise_Times_to"][npdf] == "" or sheet_plots["SetRise_Times_to"][npdf] == " ":
                                if pd.isna(sheet_plots["SetRise_Times_to"][npdf]) or sheet_plots["SetRise_Times_to"][npdf] in ["", " "]:
                                    sheet_plots["SetRise_Times_to"][npdf] = tim[-1]
                                        
                                tini = sheet_plots["Settling/Rising Times From"][npdf]
                                settimecon = template_df["Settling Time"][nsignal]
                                
                                if not tini == "No" and "Yes" in settimecon:
                                    x_min_values = str(sheet_plots["Settling/Rising Times From"][npdf]).split(',')
                                    x_max_values = str(sheet_plots["SetRise_Times_to"][npdf]).split(',')
                                    col_x_min = float(x_min_values[col]) if col < len(x_min_values) else float(x_min_values[0])
                                    col_x_max = float(x_max_values[col]) if col < len(x_max_values) else float(x_max_values[0])
                                    #print(col,col_x_max)
                                    try:
                                        ind_tmin,tmin = get_index(tim,col_x_min)
                                        #print(tmin,0)
                                    except:
                                        ind_tmin,tmin = get_index(tim,0)
                                    try:
                                        ind_tmax,tmax = get_index(tim,col_x_max) 
                                        #print(tmax,col_x_max, 1)
                                    except:
                                        try:
                                            ind_tmax,tmax = get_index(tim,ttt_max)
                                        except:
                                            ind_tmax,tmax = get_index(tim,10000,1)
                                        #print(tmax,2)
                                    try:
                                        ind_tini,tini = get_index(tim,tini)
                                    except:
                                        ind_tini,tini = get_index(tim,ttt_min)
                                        
                                    vend = sig[ind_tmax]
                                    vini = sig[ind_tini]
                                    #print(vini,vend)
                                    
                                    #REACTIVE POWER:
                                    if "Q" in settimecon:
                                        labelerrorbands = "none"#f"Errorbands {int(errband_perc*100)}%"
                                        b1 = vend+((vend-vini)*errband_perc)
                                        b2 = vend-((vend-vini)*errband_perc)
                                        if b1 > b2:
                                            b1 = vend-((vend-vini)*errband_perc)
                                            b2 = vend+((vend-vini)*errband_perc)
                                        
                                        
                                        
                                    #ACTIVE POWER
                                    if "P" in settimecon:
                                        labelerrorbands = "none"#f"Errorbands {int(errband_perc*100)}%"
                                        maxdelta = 0
                                        indmax = 0
                                        for t in reversed(range(len(tim))):
                                            if tim[t] <= tini:
                                                break
                                            if tim[t] > tmax:
                                                continue
                                            if abs(sig[t]-vini) >= maxdelta:
                                                maxdelta = abs(sig[t]-vini)
                                                indmax = t
                                        b1 = vend+((maxdelta)*errband_perc)
                                        b2 = vend-((maxdelta)*errband_perc)
                                        if b1> b2:
                                            b1 = vend-((maxdelta)*errband_perc)
                                            b2 = vend+((maxdelta)*errband_perc)
                                    
                                    
                                    #VOLTAGE
                                    
                                    if "V" in settimecon:
                                        labelerrorbands = "none"#f"Errorbands {int(errband_perc*100)}%"
                                        maxdelta = 0
                                        ind_max_delta = 0
                                        for t in reversed(range(len(tim))):
                                            if tim[t] <= tini:
                                                break
                                            if tim[t] > tmax:
                                                continue
                                            if abs(sig[t]-vini) >= maxdelta:
                                                maxdelta = abs(sig[t]-vini)
                                                ind_max_delta = t
                                               # print(tim[ind_max_delta])
                                        mean = (vini+sig[ind_max_delta])/2
                                        #print(vini,sig[ind_max_delta],tini,sig[ind_tini])
                                        
                                        if vend >= mean:
                                            b1 = vend-((vend-vini)*errband_perc)
                                            b2 = vend+((vend-vini)*errband_perc)
                                        else:
                                            b1 = vend+((maxdelta)*errband_perc)
                                            b2 = vend-((maxdelta)*errband_perc)
                                            
                                        if b1> b2:
                                            b1 = vend-((maxdelta)*errband_perc)
                                            b2 = vend+((maxdelta)*errband_perc)    
                                        
                                    for t in reversed(range(len(tim))):
                                            if t> ind_tmax:
                                                continue
                                            if sig[t] <=b1 or sig[t]>=b2:
                                                ll = sig[t]
                                                ll2 = tim[t]
                                                break
                                    delta = tim[t]-tini
                                    
                                    if 1:  
                                        if "-" in settimecon:
                                                leg = f"Settle time (10%): {delta:.3f}s"
                                                leg2 = f"Settle time (10%): <10ms"
                                        else:
                                                leg = f"Settle time: {delta:.3f}s"
                                                leg2 = f"Settle time: <10ms"
                                                  
                                        if delta > 0.01:
                                            #axs[row,col].axvline(x=tim[t], color="k", linestyle="--", label=f"Settling time: {tim[t]:.3f}s (▲: {delta:.3f}s)",linewidth=0.4)
                                            axs[row,col].axvline(x=tim[t], color="k", linestyle="--", label=leg,linewidth=0.4)
                                        else:
                                            axs[row,col].axvline(x=-100, color="k", linestyle="--", label=leg2)
                                            delta = 0.01
                                            
                                        errb = axs[row,col].fill_between(generate_linspace(tim,0.1), b1, b2, color="limegreen", alpha=0.1,label=labelerrorbands)
                                        axs[row,col].axhline(y=b1, color="limegreen", linestyle="--",linewidth=0.5,label="none")
                                        axs[row,col].axhline(y=b2, color="limegreen", linestyle="--",linewidth=0.5,label="none")
                                        # data_report.append([file_df["loc"][nfile],file_df["Page N"][nfile],tim[ind_tini],tim[ind_tmax],1,"Errorband = [%.3f-%.3f]" %(b1,b2),template_df["Signal Name"][nsignal],"Settle Time",tim[ind_tini],tim[t],tim[t]-tim[ind_tini]])
                                        if "-" in settimecon:
                                            #Report_results["Set Time(s)(N/A - 2%% MVA Used)"] = f"{delta:.3f}"
                                            pass
                                        else:
                                            Report_results["Settling time (s)"] = f"{delta:.3f}"
                                        
                                    if 1 and ("P" in settimecon or "Q" in settimecon) and "-" in settimecon:
                                        labelerrorbands = f"Errorbands 2% MVA"
                                        #print(b1,b2)
                                        maxdelta = 0
                                        indmax = 0
                                        errband_perc2 = 0.02
                                        for t in reversed(range(len(tim))):
                                            if tim[t] <= tini:
                                                break
                                            if abs(sig[t]-vini) >= maxdelta:
                                                maxdelta = abs(sig[t]-vini)
                                                indmax = t
                                        errorband = float(settimecon.split("-")[1])*errband_perc2
                                        b1 = vend+errorband
                                        b2 = vend-errorband
                                        if b1> b2:
                                            b1 = vend-errorband
                                            b2 = vend+errorband
                                        for t in reversed(range(len(tim))):
                                            if t> ind_tmax:
                                                continue
                                            if sig[t] <=b1 or sig[t]>=b2:
                                                ll = sig[t]
                                                ll2 = tim[t]
                                                break
                                        delta = tim[t]-tini
                                        if delta > 0.002:
                                            axs[row,col].axvline(x=tim[t], color="k", linestyle="--", label=f"Settle time (2% MVA): {delta:.3f}s",linewidth=0.4)
                                        else:
                                            delta = 0.01
                                            axs[row,col].axvline(x=-100, color="k", linestyle="--", label=f"Settle time (2% MVA): <10ms")
                                        errb = axs[row,col].fill_between(generate_linspace(tim,0.1), b1, b2, color="lightblue", alpha=0.3,label=labelerrorbands)
                                        axs[row,col].axhline(y=b1, color="lightblue", linestyle="--",linewidth=0.4,label="none")
                                        axs[row,col].axhline(y=b2, color="lightblue", linestyle="--",linewidth=0.4,label="none")
                                        Report_results["Settling time 2% MVA (s)"] = f"{delta:.3f}"
                                    #save_string_to_file(f"Settling time {settimecon.replace('Yes','')},{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{tini},{tim[t]},{delta}",working_dir + os.sep+ output_folder +os.sep+ f"{sheet_plots['pdf_name'][npdf]}.txt")
                                    #save_string_to_file(f"Settling time {settimecon.replace('Yes','')},{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{tini},{tim[t]},{delta}",working_dir + os.sep+ output_folder +os.sep+ f"p1_gTot_{initime}_tTot.csv")
                                    #report_52513.append()
                                        
                                    #temp_52513_titles += f",Settling time {settimecon.replace('Yes','')},"
                                    #temp_52513 += f",{delta}"
                                
                                if 1:
                                    
                                    ########################################
                                    ############## Rising TIMES xxxx##########
                                    
                                    tini = sheet_plots["Settling/Rising Times From"][npdf]

                                    if not tini == "No" and template_df["Rising Time"][nsignal] == "Yes":
                                        #print(tini)
                                        x_min_values = str(sheet_plots["Settling/Rising Times From"][npdf]).split(',')
                                        x_max_values = str(sheet_plots["SetRise_Times_to"][npdf]).split(',')
                                        col_x_min = float(x_min_values[col]) if col < len(x_min_values) else float(x_min_values[0])
                                        col_x_max = float(x_max_values[col]) if col < len(x_max_values) else float(x_max_values[0])
                                        
                                        try:
                                            ind_tmin,tmin = get_index(tim,col_x_min)
                                        except:
                                            ind_tmin,tmin = get_index(tim,0)
                                        try:
                                            ind_tmax,tmax = get_index(tim,col_x_max) 
                                            #print(tmax,1)
                                        except Exception as e:
                                            #print(e)
                                            try:
                                                ind_tmax,tmax = get_index(tim,float(sheet_plots["SetRise_Times_to"][npdf]))
                                                
                                            except Exception as e:
                                                ind_tmax,tmax = get_index(tim,10000,1)
                                                
                                        
                                        
                                        try:
                                            #print(tini,1)
                                            #print(len(tim))
                                            ind_tini,tini = get_index(tim,tini)
                                            #print(tini,2)
                                        except:
                                            ind_tini,tini = get_index(tim,ttt_min)
                                            #print(tini,3)
                                        #print(tini,tmax,2)
                                        
                                        vend = sig[ind_tmax]
                                        vini = sig[ind_tini]
                                        #print(vini,vend,33)
                                        if not pd.isna(sheet_plots["Extras"][npdf]):
                                            if "avg_rise" in sheet_plots["Extras"][npdf]:
                                                sig2 = moving_average(sig,avg_rise)
                                                #axs[row,col].plot(tim[0:ind_tmax],sig2[0:ind_tmax],label=f"Filtered",rasterized=rasterized)
                                        
                                            else:
                                                sig2 = sig
                                        else:
                                            sig2 = sig
                                            
                                        if vend > vini:
                                            
                                            perc1 = (max(sig2[ind_tini:ind_tmax])-vini)*errband_perc+vini
                                            perc2 = (max(sig2[ind_tini:ind_tmax])-vini)*(1-errband_perc)+vini  
                                            
                                            #print(perc1,perc2,222)
                                            for indt4 in (range(len(tim)-1)):
                                                if tim[indt4] <= tim[ind_tmax] and tim[indt4] >= tim[ind_tini] and sig2[indt4+1] >= perc1 and sig2[indt4+1]-sig2[indt4] > 0:
                                                    break
                                                
                                            for indt5 in range(len(tim)-1):
                                                if tim[indt5] - tim[ind_tini] < delay_rise:
                                                    continue
                                                
                                                if tim[indt5] > tmax :
                                                    #print("ERROR RISING TIME NOT FOUND")
                                                    break
                                                
                                                if  tim[indt5] < tini:
                                                        continue

                                                if sig2[indt5+1] >= perc2 and tim[indt5+1]-tim[indt5] == 0 and sig2[indt5+1]-sig2[indt5] != 0:
                                                    #Step Detected!
                                                    break
                                                
                                                if sig2[indt5] >= perc2 and indt5>=indt4: #sig2[indt5+1]-sig2[indt5] > 0:
                                                    break 
                                                

                                        
                                        else:
                                            

                                            perc1 = (min(sig2[ind_tini:ind_tmax])-vini)*errband_perc+vini
                                            perc2 = (min(sig2[ind_tini:ind_tmax])-vini)*(1-errband_perc)+vini  
                                            
                                            for indt4 in (range(len(tim)-1)):
                                                if tim[indt4] <= tim[ind_tmax] and tim[indt4] >= tim[ind_tini] and sig2[indt4+1] <= perc1 and sig2[indt4+1]-sig2[indt4] < 0:
                                                    break
                                            for indt5 in range(len(tim)-1):
                                                if tim[indt5] - tim[ind_tini] < delay_rise:
                                                        continue
                                                if tim[indt5] > tmax :
                                                    #print("ERROR RISING TIME NOT FOUND")
                                                    break
                                                if  tim[indt5] < tini:
                                                        continue
                                                if sig2[indt5+1] <= perc2 and tim[indt5+1]-tim[indt5] == 0 and sig2[indt5+1]-sig2[indt5] != 0:
                                                    #Step Detected!
                                                    break
                                                if sig2[indt5] <= perc2 and indt5>=indt4: #sig2[indt5+1]-sig2[indt5] < 0:
                                                    break 
                                                
                                                # if tim[indt5] <= tim[ind_tmax] and tim[indt5] >= tim[ind_tini] and sig2[indt5] <= perc2 and sig2[indt5+1]-sig2[indt5] < 0:
                                                    
                                                #     break 
                                                
                                        #rise = axs[row,col].fill_between(tim[indt4:indt5], sig2[indt4], sig2[indt5], color="gold", alpha=0.2,label=f"Rising Time 10-90%: {(tim[indt5]-tim[indt4]):.2f}s")
                                        delta = round(tim[indt5]-tim[indt4],3)
                                        
                                        if delta <= 0.01: 
                                            rise = axs[row,col].plot(tim[indt4:indt5],sig2[indt4:indt5],color="gold",alpha=0.5,linewidth=3,label=f"Rise Time 10-90%: <10ms",rasterized=rasterized)
                                            #delta = "<10ms"
                                            Report_results["Rising time (s)"] = 0.01
                                        else:
                                            rise = axs[row,col].plot(tim[indt4:indt5],sig2[indt4:indt5],color="gold",alpha=0.5,linewidth=3,label=f"Rise Time 10-90%: {delta:.3f}s",rasterized=rasterized)
                                            Report_results["Rising time (s)"] = f"{delta:.3f}"
                                        axs[row,col].axvline(x=tim[indt4], color="orange", linestyle="--",linewidth=0.7,label="none")
                                        axs[row,col].axvline(x=tim[indt5], color="orange", linestyle="--",linewidth=0.7,label="none")

                                        #save_string_to_file(f"Rising time,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{tmin},{tim[indt5]},{delta}",working_dir + os.sep+ output_folder +os.sep+ f"{sheet_plots['pdf_name'][npdf]}.txt")
                                        #save_string_to_file(f"Rising time,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{tmin},{tim[indt5]},{delta}",working_dir + os.sep+ output_folder +os.sep+ f"p1_gTot_{initime}_tTot.csv")
                                        


                                        # data_report.append([file_df["loc"][nfile],file_df["Page N"][nfile],"Values [%s-%s] | Times [%s-%s]" %(sig[indt4],sig[indt5],tim[indt4],tim[indt5]),template_df["Signal Name"][nsignal],"Rise 10-90 Time",tim[indt4],tim[indt5],tim[indt5]-tim[indt4]])
                                
                                if 1:
                                    
                                    ########################################
                                    ############## Commencement TIMES ##########
                               
                                    
                                    commencement_type = str(template_df["Commencement time"][nsignal]).strip().lower()

                                    if not sheet_plots["Settling/Rising Times From"][npdf] == "No" and (commencement_type == "time" or commencement_type == "yes"):
                                        # Existing time-based commencement handling
                                        tini = sheet_plots["Settling/Rising Times From"][npdf]
                                        
                                        x_min_values = str(sheet_plots["Settling/Rising Times From"][npdf]).split(',')
                                        x_max_values = str(sheet_plots["SetRise_Times_to"][npdf]).split(',')
                                        col_x_min = float(x_min_values[col]) if col < len(x_min_values) else float(x_min_values[0])
                                        col_x_max = float(x_max_values[col]) if col < len(x_max_values) else float(x_max_values[0])
                                        
                                        try:
                                            ind_tmin, tmin = get_index(tim, col_x_min)
                                        except:
                                            ind_tmin, tmin = get_index(tim, 0)
                                        try:
                                            ind_tmax, tmax = get_index(tim, col_x_max, 1) 
                                        except:
                                            try:
                                                ind_tmax, tmax = get_index(tim, float(sheet_plots["SetRise_Times_to"][npdf]))
                                            except:
                                                ind_tmax, tmax = get_index(tim, tmin + 0.1)
                                        try:
                                            ind_tini, tini = get_index(tim, tini)
                                        except:
                                            ind_tini, tini = get_index(tim, ttt_min)
                                        
                                        vend = sig[ind_tmax]
                                        vini = sig[ind_tini]

                                        if vend > vini:
                                            perc1 = (max(sig[ind_tini:ind_tmax]) - vini) * errband_perc + vini
                                            perc2 = (max(sig[ind_tini:ind_tmax]) - vini) * (1 - errband_perc) + vini  
                                            for indt4 in range(len(tim)-1):
                                                if tim[indt4] <= tim[ind_tmax] and tim[indt4] >= tim[ind_tini] and sig[indt4+1] >= perc1 and sig[indt4+1] - sig[indt4] > 0:
                                                    break
                                            for indt5 in range(len(tim)-1):
                                                if tim[indt5] <= tim[ind_tmax] and tim[indt5] >= tim[ind_tini] and sig[indt5] >= perc2 and sig[indt5+1] - sig[indt5] > 0:
                                                    if tim[indt5] - tim[ind_tini] < delay_rise:
                                                        continue
                                                    break 
                                        else:
                                            perc1 = (min(sig[ind_tini:ind_tmax]) - vini) * errband_perc + vini
                                            perc2 = (min(sig[ind_tini:ind_tmax]) - vini) * (1 - errband_perc) + vini  
                                            for indt4 in range(len(tim)-1):
                                                if tim[indt4] <= tim[ind_tmax] and tim[indt4] >= tim[ind_tini] and sig[indt4+1] <= perc1 and sig[indt4+1] - sig[indt4] < 0:
                                                    break
                                            

                                        delta = round((tim[indt4] - tim[ind_tini]) * 1000)
                                        rise = axs[row, col].plot(tim[ind_tini:indt4], sig[ind_tini:indt4], color="lime", alpha=0.5, linewidth=1, label=f"Commencement Time: {delta:.0f}ms", rasterized=rasterized)
                                        
                                        axs[row, col].axvline(x=tim[ind_tini], color="lime", linestyle="--", linewidth=0.4, label="none")
                                        axs[row, col].axvline(x=tim[indt4], color="lime", linestyle="--", linewidth=0.4, label="none")

                                        Report_results["Commencement time (ms)"] = f"{delta}"
                                        temp_52513_titles += f",Commencement time,"
                                        temp_52513 += f",{delta}"

                                    elif is_condition(commencement_type):

                                        # Handling condition-based commencement
                                        condition_expression = template_df["Commencement time"][nsignal]
                                        condition_result = evaluate_condition(condition_expression, signals_data, tim)
                                        ind_tini, tini = find_condition_index(condition_result, tim, ttt_min, ttt_max)
                                        
                                        # Proceed with the existing logic using the determined tini
                                        x_max_values = str(sheet_plots["SetRise_Times_to"][npdf]).split(',')
                                        col_x_max = float(x_max_values[col]) if col < len(x_max_values) else float(x_max_values[0])
                                        
                                        try:
                                            ind_tmax, tmax = get_index(tim, col_x_max, 1) 
                                        except:
                                            try:
                                                ind_tmax, tmax = get_index(tim, float(sheet_plots["SetRise_Times_to"][npdf]))
                                            except:
                                                ind_tmax, tmax = get_index(tim, tini + 0.07)
                                    
                                        vend = sig[ind_tmax]
                                        vini = sig[ttt_min]

                                        if vend > vini:
                                            perc1 = (max(sig[ttt_min:ind_tmax]) - vini) * errband_perc + vini
                                            perc2 = (max(sig[ttt_min:ind_tmax]) - vini) * (1 - errband_perc) + vini  
                                            for indt4 in range(len(tim)-1):
                                                if tim[indt4] <= tim[ind_tmax] and tim[indt4] >= tim[ind_tini] and sig[indt4+1] >= perc1 and sig[indt4+1] - sig[indt4] > 0:
                                                    break
                                            for indt5 in range(len(tim)-1):
                                                if tim[indt5] <= tim[ind_tmax] and tim[indt5] >= tim[ind_tini] and sig[indt5] >= perc2 and sig[indt5+1] - sig[indt5] > 0:
                                                    if tim[indt5] - tim[ind_tini] < delay_rise:
                                                        continue
                                                    break 
                                        else:
                                            perc1 = (min(sig[ttt_min:ind_tmax]) - vini) * errband_perc + vini
                                            perc2 = (min(sig[ttt_min:ind_tmax]) - vini) * (1 - errband_perc) + vini  
                                            for indt4 in range(len(tim)-1):
                                                if tim[indt4] <= tim[ind_tmax] and tim[indt4] >= tim[ind_tini] and sig[indt4+1] <= perc1 and sig[indt4+1] - sig[indt4] < 0:
                                                    break
                                            
                                        
                                        delta = round((tim[indt4] - tim[ind_tini]) * 1000)
                                        rise = axs[row, col].plot(tim[ind_tini:indt4], sig[ind_tini:indt4], color="lime", alpha=0.5, linewidth=1, label=f"Commencement Time: {delta:.0f}ms", rasterized=rasterized)
                                        
                                        axs[row, col].axvline(x=tim[ind_tini], color="lime", linestyle="--", linewidth=0.4, label="none")
                                        axs[row, col].axvline(x=tim[indt4], color="lime", linestyle="--", linewidth=0.4, label="none")

                                        Report_results["Commencement time (ms)"] = f"{delta}"
                                        

                                
                                

                                
                                #calculate_recovery_time()
                                if 1:
                                    ########################################
                                    ############## RECOVERY TIMES ##########
                                    rec_perc = 0.95
                                    if pd.isna(sheet_plots["Recovery_Times_to"][npdf]) or sheet_plots["Recovery_Times_to"][npdf] in ["", " "]:
                                        sheet_plots["Recovery_Times_to"][npdf] = tim[-1]
                                        #print("WAWDADWAWD",sheet_plots["Recovery_Times_to"][npdf])
                                        
                                    rec_tini = sheet_plots["Recovery Times From"][npdf]
                                    if not rec_tini == "No" and "Yes" in template_df["Recovery Time"][nsignal]:
                                        ind_tmin,tmin = get_index(tim,sheet_plots["X_min"][npdf])
                                        ind_tmax,tmax = get_index(tim,sheet_plots["Recovery_Times_to"][npdf],1) 
                                        ind_tini,rec_tini = get_index(tim,rec_tini)
                                        vend = sig[ind_tmax]
                                        vini = sig[ind_tmin]
                                    
                                        delta = -100
                                        try:
                                            #print(vini)
                                            #print(rec_perc)
                                            axs[row,col].axhline(y=vini*rec_perc, color="orange", linestyle="--", label=f"none",linewidth=0.5)
                                            
                                        except:
                                            print(f"{color.RED}ERROR: page {sheet_plots['Page'][npdf]} - PDF: {sheet_plots['pdf Group'][npdf]}\n{traceback.format_exc()}{color.END}")
                                        for t in reversed(range(len(tim))):
                                            if t <= ind_tini:
                                                axs[row,col].fill_between(tim[0:1], -100, -100, color="orange", alpha=0.3,label=f"Rec. Time: <10ms")
                                                # save_string_to_file(f"Recovery time,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{rec_tini},N/A,<10ms",working_dir + os.sep+ output_folder +os.sep+ f"{sheet_plots['pdf_name'][npdf]}.txt")
                                                # save_string_to_file(f"Recovery time,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{rec_tini},N/A,<10ms",working_dir + os.sep+ output_folder +os.sep+ f"p1_gTot_{initime}_tTot.csv")
                                               
                                                delta = 0.010
                                                
                                                break #NOT FOUND
                                            if t>= ind_tmax:continue
                                            if sig[ind_tmax] > 0:
                                                if sig[t] <= vini*rec_perc:
                                                    delta = tim[t] - rec_tini
                                                    #print(tim[t])
                                                    axs[row,col].fill_between(tim[ind_tini:t+1], sig[ind_tini:t+1], np.maximum(vini * rec_perc, sig[ind_tini:t + 1]), color="orange", alpha=0.3,label=f"Rec. Time: {delta:.3f}s")
                                                    axs[row,col].axvline(x=tim[t], color="orange", linestyle=":", label=f"none",linewidth=0.5)
                                                    break #FOUND TIME
                                            else:
                                                if sig[t] >= vini*rec_perc:
                                                    delta = tim[t] - rec_tini
                                                    axs[row,col].fill_between(tim[ind_tini:t+1], sig[ind_tini:t+1], np.minimum(vini * rec_perc, sig[ind_tini:t + 1]), color="orange", alpha=0.3,label=f"Rec. Time: {delta:.3f}s")
                                                    axs[row,col].axvline(x=tim[t], color="orange", linestyle=":", label=f"none",linewidth=0.5)
    
                                                    break #FOUND TIME
                                                
                                        Report_results["Recovery time (s)"] = f"{delta:.3f}"
                                                    
                                        if delta == -100:
                                            delta = 0.01
                                       
                                        
                                        # data_report.append([file_df["loc"][nfile],file_df["Page N"][nfile],"Limit (%.2fpu): %.4f [initial value = %.4f]" %(rec_perc,vini*rec_perc,sig[t]),template_df["Signal Name"][nsignal],"Rec. Time%.2fpu" %rec_perc,tim[ind_tini],tim[t],delta])
                                    ########################################
                                    ############## 1st RECOVERY TIME ##########
                                    
                                    if pd.isna(sheet_plots["Recovery_Times_to"][npdf]) or sheet_plots["Recovery_Times_to"][npdf] in ["", " "]:
                                        sheet_plots["Recovery_Times_to"][npdf] = tim[-1]
                                        #print("WAWDADWAWD",sheet_plots["Recovery_Times_to"][npdf])
                                    #print(template_df["Recovery Time"][nsignal])
                                    rec_tini = sheet_plots["Recovery Times From"][npdf]
                                    if not rec_tini == "No"  and "1st" in template_df["Recovery Time"][nsignal]:
                                        #print("Doing 1st")
                                        ind_tmin,tmin = get_index(tim,sheet_plots["X_min"][npdf])
                                        ind_tmax,tmax = get_index(tim,sheet_plots["Recovery_Times_to"][npdf],1) 
                                        ind_tini,rec_tini = get_index(tim,rec_tini)
                                        vend = sig[ind_tmax]
                                        vini = sig[ind_tmin]
                                        tol_1strec = 0.011
                                        delta = -100
                                        axs[row,col].axhline(y=vini*rec_perc, color="orange", linestyle="--", label=f"none",linewidth=line_size)

                                        for t in (range(len(tim))):
                                            if t<= ind_tini:continue
                                            if t >= ind_tmax:
                                                axs[row,col].fill_between(tim[0:1], -100, -100, color="orange", alpha=0.3,label=f"1st Rec. Time: <10ms")
                                                delta = 0.01
                                                break #NOT FOUND

                                            if sig[ind_tmax] > 0:
                                                if sig[t] >= vini*rec_perc:
                                                    delta = tim[t] - rec_tini
                                                    if delta < tol_1strec:
                                                        leg = f"1st Rec. Time: <10ms"
                                                        Report_results["1st Recovery time (s)"] = 0.01
                                                    else:
                                                        leg = f"1st Rec. Time: {delta:.3f}s"
                                                        Report_results["1st Recovery time (s)"] = f"{delta:.3f}"
                                                        
                                                    axs[row,col].axvline(x=tim[t], color="orange", linestyle=":", label=f"none",linewidth=0.5)
                                                    axs[row,col].fill_between(tim[ind_tini:t+1], sig[ind_tini:t+1], np.maximum(vini * rec_perc, sig[ind_tini:t + 1]), color="orange", alpha=0.3,label=leg)
                                                    
                                                    break #FOUND TIME
                                            else:
                                                if sig[t] <= vini*rec_perc:
                                                    delta = tim[t] - rec_tini
                                                    
                                                    if delta < tol_1strec:
                                                        leg = f"1st Rec. Time: <10ms"
                                                        Report_results["1st Recovery time (s)"] = 0.01
                                                    else:
                                                        leg = f"1st Rec. Time: {delta:.3f}s"
                                                        Report_results["1st Recovery time (s)"] = f"{delta:.3f}"
                                                        
                                                    axs[row,col].fill_between(tim[ind_tini:t+1], sig[ind_tini:t+1], np.minimum(vini * rec_perc, sig[ind_tini:t + 1]), color="orange", alpha=0.3,label=leg)
                                                    axs[row,col].axvline(x=tim[t], color="orange", linestyle=":", label=f"none",linewidth=0.5)
        
                                                    break #FOUND TIME
                                        
                                        
                                        if 0:
                                            if delta > 0.002:
                                                axs[row,col].axvline(x=tim[t], color="k", linestyle=":", label=f"Recovery {rec_perc*100}%: {delta:.3f}s")
                                                
                                            else:
                                                axs[row,col].axvline(x=-100, color="k", linestyle=":", label=f"Recovery {rec_perc*100}%: N/A")
                                        
                                        #temp_5255_titles += f",1st Rec time,"
                                        #temp_5255 += f",{delta}"
                                        # data_report.append([file_df["loc"][nfile],file_df["Page N"][nfile],"Limit (%.2fpu): %.4f [initial value = %.4f]" %(rec_perc,vini*rec_perc,sig[t]),template_df["Signal Name"][nsignal],"Recovery Time %.2fpu" %rec_perc,tim[ind_tini],tim[t],delta])
                                    
                                    ########################################
                                    ############## 5258 TIME ##########
                                    if 0:
                                        if pd.isna(sheet_plots["Recovery_Times_to"][npdf]) or sheet_plots["Recovery_Times_to"][npdf] in ["", " "]:
                                            sheet_plots["Recovery_Times_to"][npdf] = tim[-1]
                                            #print("WAWDADWAWD",sheet_plots["Recovery_Times_to"][npdf])
                                        #print(template_df["Recovery Time"][nsignal])
                                        rec_tini = sheet_plots["Recovery Times From"][npdf]
                                        if not rec_tini == "No"  and "5258" in template_df["Recovery Time"][nsignal]:
                                            #print("Doing 1st")
                                            rec_perc = 0.5
                                            ind_tmin,tmin = get_index(tim,sheet_plots["X_min"][npdf])
                                            ind_tmax,tmax = get_index(tim,sheet_plots["Recovery_Times_to"][npdf],1) 
                                            ind_tini,rec_tini = get_index(tim,rec_tini)
                                            vend = sig[ind_tmax]
                                            vini = sig[ind_tini]
                                            
                                            delta = -100
                                            axs[row,col].axhline(y=vini*rec_perc, color="lime", linestyle="--", label=f"none",linewidth=0.5)

                                            for t in (range(len(tim))):
                                                if t<= ind_tini:continue
                                                if t >= ind_tmax:
                                                    axs[row,col].fill_between(tim[0:1], -100, -100, color="lime", alpha=0.3,label=f"S5.2.5.8 Time {int(rec_perc*100)}%: N/A")
                                                    #save_string_to_file(f"S5.2.5.8 Time,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{rec_tini},N/A,N/A",working_dir + os.sep+ output_folder +os.sep+ f"{sheet_plots['pdf_name'][npdf]}.txt")
                                                    #save_string_to_file(f"S5.2.5.8 Time,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{rec_tini},N/A,N/A",working_dir + os.sep+ output_folder +os.sep+ f"p1_gTot_{initime}_tTot.csv")
                                            
                                                    break #NOT FOUND
                                                
                                                if sig[ind_tini] > 0:
                                                    #print(tim[t],sig[t])
                                                    if sig[t] <= vini*rec_perc:
                                                        delta = tim[t] - rec_tini
                                                    # print(tim[t],sig[t],"------------------")
                                                        axs[row,col].fill_between(tim[ind_tini:t+1], sig[ind_tini:t+1], np.maximum(vini * rec_perc, sig[ind_tini:t + 1]), color="lime", alpha=0.3,label=f"S5.2.5.8 Time {rec_perc*100}%: {tim[t]:.3f}s (▲: {delta:.3f}s)")
                                                        axs[row,col].axvline(x=tim[t], color="lime", linestyle=":", label=f"none",linewidth=0.5)
                                                        save_string_to_file(f"S5.2.5.8 Time,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{rec_tini},{tim[t]},{delta:.4f}",working_dir + os.sep+ output_folder +os.sep+ f"{sheet_plots['pdf_name'][npdf]}.txt")
                                                        #save_string_to_file(f"S5.2.5.8 Time,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{delta:.4f}",working_dir + os.sep+ output_folder +os.sep+ f"{sheet_plots['pdf_name'][npdf]}.txt")
                                                        save_string_to_file(f"S5.2.5.8 Time,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{rec_tini},{tim[t]},{delta}",working_dir + os.sep+ output_folder +os.sep+ f"p1_gTot_{initime}_tTot.csv")
                                                        break #FOUND TIME
                                                else:
                                                    if sig[t] >= vini*rec_perc:
                                                        delta = tim[t] - rec_tini
                                                        axs[row,col].fill_between(tim[ind_tini:t+1], sig[ind_tini:t+1], np.minimum(vini * rec_perc, sig[ind_tini:t + 1]), color="lime", alpha=0.3,label=f"S5.2.5.8 Time {rec_perc*100}%: {tim[t]:.3f}s (▲: {delta:.3f}s)")
                                                        axs[row,col].axvline(x=tim[t], color="lime", linestyle=":", label=f"none",linewidth=0.5)
                                                        save_string_to_file(f"S5.2.5.8 Time,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{rec_tini},{tim[t]},{delta}",working_dir + os.sep+ output_folder +os.sep+ f"{sheet_plots['pdf_name'][npdf]}.txt")
                                                        save_string_to_file(f"S5.2.5.8 Time,{file_df['loc'][nfile]},{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]},{sheet_plots['pdf Group'][npdf]},{rec_tini},{tim[t]},{delta}",working_dir + os.sep+ output_folder +os.sep+ f"p1_gTot_{initime}_tTot.csv")
                                            
                                                        break #FOUND TIME
                                            
                                        
                                        
                                    
                                        # data_report.append([file_df["loc"][nfile],file_df["Page N"][nfile],"Limit (%.2fpu): %.4f [initial value = %.4f]" %(0.5,vini*0.5,sig[t]),template_df["Signal Name"][nsignal],"S5.2.5.8 Time %.2fpu" %0.5,tim[ind_tini],tim[t],delta])
                                    results_list.append(Report_results)
                                    ##############################################################################################################        
                                    if "Yes" in template_df["Settling Time"][nsignal] and not "No" in str(sheet_plots["Settling/Rising Times From"][npdf]):
                                        axs[row,col].axvline(x=tini, color="k", linestyle=":", label=f"noneStep Applied: {tini:.1f}s",linewidth=0.4)
                                        s52513 = 1
                                    elif "Yes" in template_df["Rising Time"][nsignal] and not "No" in str(sheet_plots["Settling/Rising Times From"][npdf]): 
                                        axs[row,col].axvline(x=tini, color="k", linestyle=":", label=f"noneStep Applied: {tini:.1f}s",linewidth=0.4)
                                        s52513 = 1
                                    if ("Yes" in template_df["Recovery Time"][nsignal] or "1st" in template_df["Recovery Time"][nsignal]) and not rec_tini == "No":
                                        axs[row,col].axvline(x=rec_tini, color="k", linestyle=":", label=f"noneClearing time: {rec_tini:.3f}s", linewidth=0.4)
                                        
                                        set_clearing_lines = 1
                                        s52513 = 1
                                        S5255_rec = 1
                                    #temp_52513_data = f"{template_df['Legend'][nsignal]},{sheet_plots['Page'][npdf]}"
                                    #save_string_to_file(f"{temp_52513_titles}",working_dir + os.sep+ output_folder +os.sep+ f"S52513_Report_{initime}.csv")
            #if "Yes" in template_df["Settling Time"][nsignal] or "Yes" in template_df["Rising Time"][nsignal]:              
            
            #if s52513 == 1:
            #    save_string_to_file(f"{sheet_plots['pdf Group'][npdf]},{sheet_plots['Page'][npdf]}{temp_52513}",working_dir + os.sep+ output_folder +os.sep+ f"S52513_Report.csv")
            #if S5255_rec == 1:
            #    save_string_to_file(f"{sheet_plots['pdf Group'][npdf]},{sheet_plots['Page'][npdf]}{temp_5255}",working_dir + os.sep+ output_folder +os.sep+ f"S5255_Report.csv")
                
             
            #except Exception as e:
            #    print(f"ERROR {e}")
            # Split the X_min and X_max fields by comma
            if do_ORT != "":
                if "Vgridsig" not in do_ORT:
                    print(f"{color.RED}ORT Error: {color.WARNING}Grid V {color.RED} signal not found. Please check legend names or Extras function{color.END}")
                if "Vsig" not in do_ORT:
                    print(f"{color.RED}ORT Error: {color.WARNING}POC V {color.RED} signal not found. Please check legend names or Extras function{color.END}")
                if "Qsig" not in do_ORT:
                    print(f"{color.RED}ORT Error: {color.WARNING}POC Q {color.RED} signal not found. Please check legend names or Extras function{color.END}")
                
            if 1:
                if set_clearing_lines == 1:
                    #print(rec_tini)
                    for ax in axs.flatten():
                        try:
                            line_exists = any("Clearing time" in line.get_label() for line in ax.lines)
                            if not line_exists and len(ax.lines) != 0:
                                ax.axvline(x=rec_tini, color="k", linestyle=":", label=f"noneClearing time", linewidth=0.4)
                            
                        except Exception as e:
                            #print(e)
                            pass
            x_min_values = str(sheet_plots["X_min"][npdf]).split(',')
            x_max_values = str(sheet_plots["X_max"][npdf]).split(',')

            # Get the number of columns in your grid of subplots
            num_columns = len(axs[0])
            
            for iind, ax in enumerate(axs.flatten()):
                if len(ax.lines) != 0:
                    handles, labels = ax.get_legend_handles_labels()

                    # Sanitize and filter handles/labels together to preserve order
                    handles_filtered, labels_filtered = sanitize_legend_entries(handles, labels, hide_legends)

                    # If after filtering we have no legends, skip the legend creation
                    if not labels_filtered:
                        # Still apply tick formatting, grid, etc.
                        ax.ticklabel_format(useOffset=False)
                        ax.grid(True, alpha=0.3, color="black", linestyle="dotted")
                        ax.tick_params(which='both', width=0.2)
                        ax.tick_params(labelbottom=1)
                        plt.setp(ax.xaxis.get_ticklabels(), rotation=0, fontname="Times New Roman", 
                                fontweight="ultralight", fontsize=axis_size)
                        plt.setp(ax.yaxis.get_ticklabels(), rotation=0, fontname="Times New Roman", 
                                fontweight="ultralight", fontsize=axis_size)
                        ax.xaxis.set_major_locator(MaxNLocator(integer=False, min_n_ticks=min_xticks, nbins="auto"))
                        ax.yaxis.set_major_locator(MaxNLocator(integer=False, min_n_ticks=min_yticks, nbins="auto"))
                        ax.margins(x=0)
                        continue

                    # Remove existing legend before creating a new one
                    if ax.legend_ is not None:
                        ax.legend_.remove()
                    box = ax.get_window_extent().transformed(fig.dpi_scale_trans.inverted())
                    subplot_width = box.width
                    # Calculate the number of columns for the legend if not forced
                    ncol_legends = calculate_ncols(fig,ax,labels_filtered,font_size=legend_size) if force_ncols is None else force_ncols#fit_legend_ncols(ax, handles_filtered, labels_filtered, subplot_width, font_size=legend_size) if force_ncols is None else force_ncols

                    # Create the new legend
                    legend = ax.legend(
                        handles_filtered, labels_filtered,
                        loc='lower left', bbox_to_anchor=(0, 1.02, 1, 0.2),
                        fancybox=True, shadow=True, mode=None, borderaxespad=0,
                        frameon=False, fontsize=legend_size, ncol=ncol_legends
                    )

                    # Update legend text formatting
                    for text in legend.get_texts():
                        updated_text = text.get_text().replace(r'\n', '\n')
                        text.set_text(updated_text)
                        if "[ERROR]" in text.get_text():
                            text.set_color("red")

                # Apply axes formatting regardless of whether legend exists
                ax.ticklabel_format(useOffset=False)
                ax.grid(True, alpha=0.3, color="black", linestyle="dotted")
                ax.tick_params(which='both', width=0.2)
                ax.tick_params(labelbottom=1)

                plt.setp(ax.xaxis.get_ticklabels(), rotation=0, fontname="Times New Roman",
                        fontweight="ultralight", fontsize=axis_size)
                plt.setp(ax.yaxis.get_ticklabels(), rotation=0, fontname="Times New Roman",
                        fontweight="ultralight", fontsize=axis_size)
                ax.xaxis.set_major_locator(MaxNLocator(integer=False, min_n_ticks=min_xticks, nbins="auto"))
                ax.yaxis.set_major_locator(MaxNLocator(integer=False, min_n_ticks=min_yticks, nbins="auto"))
                ax.margins(x=0)
                
                
                



                # Get the column index
                column_index = iind % num_columns
                
                # Use the column index to access the corresponding X_min and X_max
                x_min = float(x_min_values[column_index]) if column_index < len(x_min_values) else float(x_min_values[0])
                try:
                    x_max = float(x_max_values[column_index]) if column_index < len(x_max_values) else float(x_max_values[0])
                except:
                    try:
                        x_max = get_index(tim,10000,1)[1]
                    except:
                        continue
                #print(column_index,x_max)
                # Set the x-axis limits
                if not np.isnan(x_min) and not np.isinf(x_min):
                    ax.set_xlim(left=x_min)
                if not np.isnan(x_max) and not np.isinf(x_max):
                    ax.set_xlim(right=x_max)
                else:
                    try:
                        if "xzoom_shorter" in sheet_plots["Extras"][npdf]:
                            ax.set_xlim(right=shorter_x_range)
                    except: pass
            # Loop through subplots to apply y-axis limits with margins
            for (row, col), (overall_min, overall_max) in overall_min_max.items():
                ax = axs[row, col]
                margin = margins[(row,col)]  # Adjust as needed
                #ax.relim()
                #ax.autoscale_view()
                # Check if margin contains '>'
                try:
                    if '>>>' in str(margin): #This will have at least that number when calculating the yrange
                        minrange = float(margin.split('>>>')[1])
                        yrange = (overall_max - overall_min)*1.2

                        if yrange < minrange:
                            yrange = minrange

                        mid_point = (overall_max + overall_min) / 2
                        ylimn = [mid_point - yrange/2, mid_point + yrange/2]
                        #print(mid_point, ylimn)
                        ax.set_ylim(ylimn)
                        
                    elif '>>' in str(margin): #This will have at least that number of decimals+1 when calculating the yrange (beta)
                        mindecimals = float(margin.split('>>')[1])
                        yrange = overall_max - overall_min
                        if yrange < 10**(-mindecimals):
                            yrange = 10**(-mindecimals)
                        ylimn = [overall_min - yrange, overall_max + yrange]
                        ax.set_ylim(ylimn)
                        
                    elif '>' in str(margin):
                        lower_limit, upper_limit = map(float, margin.split('>'))
                        ax.set_ylim([lower_limit, upper_limit])
                    else:
                        
                            yrange = (overall_max - overall_min) * float(margin)
                            ylimn = [overall_min - yrange, overall_max + yrange]
                            ax.set_ylim(ylimn)
                except Exception as e:
                    #print(e)
                    pass
            
            #if 1:
                
            [fig.delaxes(ax) for ax in axs.flatten() if not ax.has_data()]
            #fig.text(0, 1, "Page # " + str(npage), transform=fig.transFigure, horizontalalignment="right", fontname="Times New Roman", fontweight="normal", fontsize=10)

            #fig.text(max(ax.get_position().x1 for ax in axs.flatten()), min(ax.get_position().y0 for ax in axs.flatten()) - 0.05, "Page # " + str(int(npage)), transform=fig.transFigure, horizontalalignment="right", fontname="Times New Roman", fontweight="normal", fontsize=9)
            fig.text(0.95, 0.025, "Page # " + str(int(npage)), transform=fig.transFigure, horizontalalignment="right", fontname="Times New Roman", fontweight="normal", fontsize=9)
            
            #fig.figimage(company_logo, xo=0, yo=fig.bbox.ymax - company_logo.shape[0])  # Adjust the coordinates as needed

            #fig.text(0.5, min(ax.get_position().y0 for ax in axs.flatten()) - 0.05, "Text Under Lowest Row", transform=fig.transFigure, horizontalalignment="right", fontname="Times New Roman", fontweight="normal", fontsize=10)
            yloc = 0
            for loctxt in db_loc:
                yloc += 0.013
                if sheet_plots["print fnames"][npdf] == "Yes":
                    fig.text(0.8, yloc, loctxt + "\n", transform=fig.transFigure, horizontalalignment="right", fontname="Times New Roman", fontweight="normal", fontsize=8)
            
            title = ""
            if  sheet_plots["Title 1"][npdf] != "__":
                title += str(sheet_plots["Title 1"][npdf]) + "\n"
            if  sheet_plots["Title 2"][npdf] != "__":
                title += str(sheet_plots["Title 2"][npdf]) + "\n"
            if  sheet_plots["Title 3"][npdf] != "__":
                title += str(sheet_plots["Title 3"][npdf]) + "\n"
            if  sheet_plots["Title 4"][npdf] != "__":
                title += str(sheet_plots["Title 4"][npdf]) + "\n"

            fig.suptitle(title, fontname="Times New Roman", fontweight="bold", fontsize=12, x=0.5, y=0.95)
            #fig.subplots_adjust(hspace = 0.6, wspace=.40)
            addLogo = OffsetImage(company_logo, zoom=0.5)
            # Use AnnotationBbox to position the logo at a relative position within the figure
            logo_box = AnnotationBbox(addLogo, (0.05, 0.95),  # (x, y) position in figure coordinates
                                    xycoords='figure fraction',  # use figure-relative coordinates
                                    frameon=False)  # No box around the logo
            fig.add_artist(logo_box)
            fig.tight_layout(rect=[0.01, yloc+0.01, 0.98, 0.97],pad=0.2, h_pad=0.05, w_pad=0.08)
            # Connect the click event to the function

            fig.savefig(working_dir + os.sep+ output_folder +os.sep+ f"{sheet_plots['pdf_name'][npdf]}.pdf", dpi=300) #, bbox_inches="tight"
            #
            
                    
            zoomed = 0        
            
            for iind,ax in enumerate(axs.flatten()):
                try:
                    x_min_values = str(sheet_plots["X_min zoom"][npdf]).split(',')
                    x_max_values = str(sheet_plots["X_max zoom"][npdf]).split(',')
                    column_index = iind % num_columns
                    col_x_min = float(x_min_values[column_index]) if column_index < len(x_min_values) else float(x_min_values[0])
                    col_x_max = float(x_max_values[column_index]) if column_index < len(x_max_values) else float(x_max_values[0])
                except:continue
                if not (sheet_plots["X_min zoom"][npdf]) == "No":
                    ax.set_xlim(left=col_x_min)
                    zoomed = 1
                if not (sheet_plots["X_max zoom"][npdf]) == "No":
                    ax.set_xlim(right=col_x_max)
                    zoomed = 1

            

                
            if zoomed ==1:
                fig.savefig(working_dir + os.sep+ output_folder +os.sep+ f"{sheet_plots['pdf_name'][npdf].split('_t')[0]+'_Zoomed_t'+str(npdf)}.pdf", bbox_inches="tight", dpi=400)
            
            #Close Figure to free RAM - Liam
            plt.close(fig)
            
            
delete_txt_files(output_folder,["_g","_t"])
#merge_txt_to_xlsx_by_group(output_folder,output_folder,"Time_Results")
#exit()

def truncate_filename(filename, max_length=30):
    if len(filename) > max_length:
        return filename[:max_length - 3] + "..."
    return filename

def get_file_size_in_gb(file_path):
    """Returns the size of the file in GB."""
    file_size = os.path.getsize(file_path)  # Get file size in bytes
    return file_size / (1024 ** 3)  # Convert bytes to GB

def calculate_total_size_and_average(sheet_files, plot_dict):
    page_size_map = defaultdict(float)  # Dictionary to store total size per page (in GB)
    total_pages = 0
    dont_exist_flag = 0
    # Iterate through the files and sum up their sizes per page
    for ind2, file_path in enumerate(sheet_files["loc"]):
        if isinstance(file_path, float):
            #print(f"{color.RED}UPS, seems like you have a missing file location in the 'Files' Sheet{color.END}")
            continue
        key = (sheet_files["pdf Group"][ind2], sheet_files["Page N"][ind2])
        if key in plot_dict:
            if os.path.exists(file_path):
                # Get the size of the file in GB
                file_size_gb = get_file_size_in_gb(file_path)
                # Add file size to the corresponding page
                page_size_map[key] += file_size_gb
            else:
                print(f"{color.RED}File missing{color.END} for Group {color.WARNING}{sheet_files['pdf Group'][ind2]} - Page {sheet_files['Page N'][ind2]}{color.END}\nFile: {file_path}\n")
                dont_exist_flag = 1
    total_size_gb = sum(page_size_map.values())
    total_pages = len(page_size_map)  # Count how many unique pages have files

    # Calculate the average file size per page
    if total_pages > 0:
        avg_size_gb_per_page = total_size_gb / total_pages
    else:
        avg_size_gb_per_page = 0
    print(f"{color.WARNING}---------------------------------------------------------------------------------------------------------------------------------------------------------{color.END}")
    print(f"Total file size: {total_size_gb:.4f} GB")
    #print(f"Total number of unique pages: {total_pages}")
    print(f"Average file size per page: {avg_size_gb_per_page:.4f} GB")

    return avg_size_gb_per_page,dont_exist_flag

def get_commit_charge():
    class MEMORYSTATUSEX(ctypes.Structure):
        _fields_ = [("dwLength", ctypes.c_ulong), ("dwMemoryLoad", ctypes.c_ulong),
                    ("ullTotalPhys", ctypes.c_ulonglong), ("ullAvailPhys", ctypes.c_ulonglong),
                    ("ullTotalPageFile", ctypes.c_ulonglong), ("ullAvailPageFile", ctypes.c_ulonglong),
                    ("ullTotalVirtual", ctypes.c_ulonglong), ("ullAvailVirtual", ctypes.c_ulonglong),
                    ("ullAvailExtendedVirtual", ctypes.c_ulonglong)]
    
    memory_status = MEMORYSTATUSEX()
    memory_status.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
    ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory_status))
    
    used_commit_charge = memory_status.ullTotalPageFile - memory_status.ullAvailPageFile
    total_commit_charge = memory_status.ullTotalPageFile

    # Calculate percentage of used commit charge
    commit_charge_percentage = (used_commit_charge / total_commit_charge) * 100

    # Convert used and total commit charge to GB
    used_commit_charge_gb = used_commit_charge / (1024 ** 3)  # Convert from bytes to GB
    total_commit_charge_gb = total_commit_charge / (1024 ** 3)  # Convert from bytes to GB

    return commit_charge_percentage, used_commit_charge_gb, total_commit_charge_gb

def extract_split_pdf(input_string):
    # Use regex to match 'split_pdf(' followed by values in parentheses
    pattern = r'split_pdf\(([^,]+),\s*([^,\)]+)(?:,\s*(\d+))?\)'  # Matches the specific function call
    matches = re.findall(pattern, input_string)
    
    # Extract and clean values from all matches
    result = []
    for match in matches:
        value1 = match[0].strip()  # First value
        value2 = match[1].strip()  # Second value
        optional_value = str(match[2]) if match[2] else "0"  # Optional value, defaulting to 1
        result.append([value1, value2, optional_value])
    
    return result  # Return a list of lists


def sanitize_sheet_name(name):
    invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
    for char in invalid_chars:
        name = name.replace(char, '_')
    return name[:31]  # Excel sheet name limit

def split_group_into_subgroups(group_name):
    """
    Split the group name by '_' and return the list of subgroups.
    """
    if pd.isna(group_name):
        return []
    return group_name.split('_')



def Create_Reports(output_folder, results_list):
    """
    Genera diversos informes a partir de la `results_list` de diccionarios.
    Solo se guardan hojas transpuestas con prefijo 'T' y formato 'IdN_TM'.
    Luego imprime (y/o guarda) una tabla resumen con los valores máximos de cualquier
    columna nueva que no era parte del conjunto "estándar".

    Args:
        output_folder (str): Ruta a la carpeta donde se guardan los informes.
        results_list (list): Lista de diccionarios; cada diccionario tiene
                             un conjunto de claves que describen los datos (p.ej., "Group",
                             "Signal", posiblemente claves nuevas adicionales).
    """

    # Definir columnas estándar al inicio
    standard_columns = [
        "Group", "Page", "File_id", "File_loc", "Fname", "Signal",
        "Title1", "Title2", "Title3", "Title4", "Extra Info"
    ]
    identifier_columns = [col for col in standard_columns if col != "Signal"]
    
    # Convertir la lista de resultados a un DataFrame
    results_df = pd.DataFrame(list(results_list))
    if results_df.empty:
        print(f"{color.WARNING}No data available to create reports... ignoring{color.END}")
        return

    # Validar columnas esenciales
    essential_columns = ["Group", "File_id"]
    for col in essential_columns:
        if col not in results_df.columns:
            print(f"{color.WARNING}'{col}' not found.{color.END}")
            return

    # Si 'Page' no existe, añadir una columna vacía (opcional según necesidad)
    if "Page" not in results_df.columns:
        results_df["Page"] = ""

    # Convertir columnas después de 'Extra Info' a float donde sea posible
    if "Extra Info" in results_df.columns:
        start_col_index = results_df.columns.get_loc("Extra Info") + 1
        for col in results_df.columns[start_col_index:]:
            results_df[col] = pd.to_numeric(results_df[col], errors="coerce")

    # Recolectar todos los subgrupos de 'Group'
    all_subgroups = set()
    for group in results_df["Group"].dropna():
        subgroups = split_group_into_subgroups(group)  # función definida por el usuario
        all_subgroups.update(subgroups)
    all_subgroups = {sg for sg in all_subgroups if sg}  # Eliminar cadenas vacías

    # Identificar columnas nuevas (no estándar)
    all_new_columns = sorted(set(results_df.columns) - set(standard_columns))

    # Lista para almacenar el resumen de valores máximos
    summary_records = []

    
    print("\nCreating reports...")

    # Ordenar por 'File_id' y 'Page' para consistencia
    results_df = results_df.sort_values(by=["File_id", "Page"])

    # ----------------------------------------------------------------
    # 1) Crear REPORT_ALL.xlsx (solo hojas transpuestas)
    # ----------------------------------------------------------------
    all_report_path = os.path.join(output_folder, "REPORT_ALL.xlsx")
    with pd.ExcelWriter(all_report_path, engine="openpyxl") as writer:
        used_sheet_names = set()

        # Hoja transpuesta general
        pivot_df = results_df.copy()
        measurement_columns = [col for col in pivot_df.columns if col not in standard_columns]
        pivot_df = pivot_df.melt(
            id_vars=standard_columns,
            value_vars=measurement_columns,
            var_name="Measurement",
            value_name="Value"
        )
        pivot_df["Measurement_Signal"] = pivot_df["Measurement"] + " - " + pivot_df["Signal"]
        transposed_unique_df = pivot_df.pivot_table(
            index=identifier_columns,
            columns="Measurement_Signal",
            values="Value",
            aggfunc="first"
        ).reset_index()
        transposed_unique_df = transposed_unique_df.sort_values(by=["File_id", "Page"])

        # Aplanar MultiIndex si existe
        if isinstance(transposed_unique_df.columns, pd.MultiIndex):
            transposed_unique_df.columns = [
                "_".join(col).strip() if col[1] else col[0] 
                for col in transposed_unique_df.columns.values
            ]
        
        # Nombre de la hoja transpuesta general
        transposed_sheet_name = "All_Results"
        transposed_sheet_name = sanitize_sheet_name(transposed_sheet_name, used_sheet_names)
        transposed_unique_df.to_excel(writer, index=False, sheet_name=transposed_sheet_name)
        used_sheet_names.add(transposed_sheet_name)

        # Hojas transpuestas por File_id
        unique_file_ids = transposed_unique_df["File_id"].unique()
        for file_id in unique_file_ids:
            formatted_file_id = format_file_id(file_id)
            sheet_name = f"Id{formatted_file_id}_All_Results"
            sheet_name = sanitize_sheet_name(sheet_name, used_sheet_names)
            file_df = pivot_df[pivot_df["File_id"] == file_id]
            transposed_file_df = file_df.pivot_table(
                index=standard_columns,
                columns="Measurement_Signal",
                values="Value",
                aggfunc="first"
            ).reset_index()
            transposed_file_df = transposed_file_df.sort_values(by=["File_id", "Page"])

            if isinstance(transposed_file_df.columns, pd.MultiIndex):
                transposed_file_df.columns = [
                    "_".join(col).strip() if col[1] else col[0] 
                    for col in transposed_file_df.columns.values
                ]

            transposed_file_df.to_excel(writer, index=False, sheet_name=sheet_name)
            used_sheet_names.add(sheet_name)

    print(f"{color.GREEN}File created: REPORT_ALL.xlsx{color.END}")

    # Registro de resumen para REPORT_ALL
    record_all = {"Report_file": "REPORT_ALL.xlsx"}
    for col in all_new_columns:
        record_all[col] = results_df[col].max() if col in results_df.columns else "N/A"
    summary_records.append(record_all)

    # ----------------------------------------------------------------
    # 2) Crear informes por grupo (solo hojas transpuestas)
    # ----------------------------------------------------------------
    unique_groups = results_df["Group"].dropna().unique()
    for group in unique_groups:
        group_df = results_df[results_df["Group"] == group]
        sanitized_group = sanitize_sheet_name(str(group))
        group_report_filename = f"REPORT_{sanitized_group}.xlsx"
        group_path = os.path.join(output_folder, group_report_filename)

        with pd.ExcelWriter(group_path, engine="openpyxl") as writer:
            used_sheet_names = set()

            # Hoja transpuesta general del grupo
            pivot_group_df = group_df.copy()
            measurement_columns_group = [col for col in pivot_group_df.columns if col not in standard_columns]
            pivot_group_df = pivot_group_df.melt(
                id_vars=standard_columns,
                value_vars=measurement_columns_group,
                var_name="Measurement",
                value_name="Value"
            )
            pivot_group_df["Measurement_Signal"] = pivot_group_df["Measurement"] + " - " + pivot_group_df["Signal"]
            transposed_unique_group_df = pivot_group_df.pivot_table(
                index=identifier_columns,
                columns="Measurement_Signal",
                values="Value",
                aggfunc="first"
            ).reset_index()
            transposed_unique_group_df = transposed_unique_group_df.sort_values(by=["File_id", "Page"])

            if isinstance(transposed_unique_group_df.columns, pd.MultiIndex):
                transposed_unique_group_df.columns = [
                    "_".join(col).strip() if col[1] else col[0] 
                    for col in transposed_unique_group_df.columns.values
                ]

            transposed_group_sheet_name = f"All_{sanitized_group}"
            transposed_group_sheet_name = sanitize_sheet_name(transposed_group_sheet_name, used_sheet_names)
            transposed_unique_group_df.to_excel(writer, index=False, sheet_name=transposed_group_sheet_name)
            used_sheet_names.add(transposed_group_sheet_name)

            # Hojas transpuestas por File_id dentro del grupo
            unique_file_ids_group = transposed_unique_group_df["File_id"].unique()
            for file_id in unique_file_ids_group:
                formatted_file_id = format_file_id(file_id)
                sheet_name = f"Id{formatted_file_id}_{sanitized_group}"
                sheet_name = sanitize_sheet_name(sheet_name, used_sheet_names)
                file_df_group = pivot_group_df[pivot_group_df["File_id"] == file_id]
                transposed_file_group_df = file_df_group.pivot_table(
                    index=standard_columns,
                    columns="Measurement_Signal",
                    values="Value",
                    aggfunc="first"
                ).reset_index()
                transposed_file_group_df = transposed_file_group_df.sort_values(by=["File_id", "Page"])

                if isinstance(transposed_file_group_df.columns, pd.MultiIndex):
                    transposed_file_group_df.columns = [
                        "_".join(col).strip() if col[1] else col[0] 
                        for col in transposed_file_group_df.columns.values
                    ]

                transposed_file_group_df.to_excel(writer, index=False, sheet_name=sheet_name)
                used_sheet_names.add(sheet_name)

        print(f"{color.GREEN}File Created: {group_report_filename}{color.END}")

        # Registro de resumen para el grupo
        record_group = {"Report_file": group_report_filename}
        for col in all_new_columns:
            record_group[col] = group_df[col].max() if col in group_df.columns else "N/A"
        summary_records.append(record_group)

    # ----------------------------------------------------------------
    # 3) Crear informes por subgrupo (solo hojas transpuestas)
    # ----------------------------------------------------------------
    for subgroup in all_subgroups:
        mask = results_df["Group"].apply(
            lambda x: subgroup in split_group_into_subgroups(x) if pd.notna(x) else False
        )
        subgroup_df = results_df[mask]
        if subgroup_df.empty:
            continue

        sanitized_subgroup = sanitize_sheet_name(str(subgroup))
        subgroup_report_filename = f"REPORT_{sanitized_subgroup}.xlsx"
        subgroup_path = os.path.join(output_folder, subgroup_report_filename)

        with pd.ExcelWriter(subgroup_path, engine="openpyxl") as writer:
            used_sheet_names = set()

            # Hoja transpuesta general del subgrupo
            pivot_subgroup_df = subgroup_df.copy()
            measurement_columns_subgroup = [col for col in pivot_subgroup_df.columns if col not in standard_columns]
            pivot_subgroup_df = pivot_subgroup_df.melt(
                id_vars=standard_columns,
                value_vars=measurement_columns_subgroup,
                var_name="Measurement",
                value_name="Value"
            )
            pivot_subgroup_df["Measurement_Signal"] = pivot_subgroup_df["Measurement"] + " - " + pivot_subgroup_df["Signal"]
            transposed_unique_subgroup_df = pivot_subgroup_df.pivot_table(
                index=standard_columns,
                columns="Measurement_Signal",
                values="Value",
                aggfunc="first"
            ).reset_index()
            transposed_unique_subgroup_df = transposed_unique_subgroup_df.sort_values(by=["File_id", "Page"])

            if isinstance(transposed_unique_subgroup_df.columns, pd.MultiIndex):
                transposed_unique_subgroup_df.columns = [
                    "_".join(col).strip() if col[1] else col[0] 
                    for col in transposed_unique_subgroup_df.columns.values
                ]

            transposed_subgroup_sheet_name = f"All_{sanitized_subgroup}"
            transposed_subgroup_sheet_name = sanitize_sheet_name(transposed_subgroup_sheet_name, used_sheet_names)
            transposed_unique_subgroup_df.to_excel(writer, index=False, sheet_name=transposed_subgroup_sheet_name)
            used_sheet_names.add(transposed_subgroup_sheet_name)

            # Hojas transpuestas por File_id dentro del subgrupo
            unique_file_ids_subgroup = transposed_unique_subgroup_df["File_id"].unique()
            for file_id in unique_file_ids_subgroup:
                formatted_file_id = format_file_id(file_id)
                sheet_name = f"Id{formatted_file_id}_{sanitized_subgroup}"
                sheet_name = sanitize_sheet_name(sheet_name, used_sheet_names)
                file_df_subgroup = pivot_subgroup_df[pivot_subgroup_df["File_id"] == file_id]
                transposed_file_subgroup_df = file_df_subgroup.pivot_table(
                    index=standard_columns,
                    columns="Measurement_Signal",
                    values="Value",
                    aggfunc="first"
                ).reset_index()
                transposed_file_subgroup_df = transposed_file_subgroup_df.sort_values(by=["File_id", "Page"])

                if isinstance(transposed_file_subgroup_df.columns, pd.MultiIndex):
                    transposed_file_subgroup_df.columns = [
                        "_".join(col).strip() if col[1] else col[0] 
                        for col in transposed_file_subgroup_df.columns.values
                    ]

                transposed_file_subgroup_df.to_excel(writer, index=False, sheet_name=sheet_name)
                used_sheet_names.add(sheet_name)

        print(f"{color.GREEN}File Created: {subgroup_report_filename}{color.END}")

        # Registro de resumen para el subgrupo
        record_subgroup = {"Report_file": subgroup_report_filename}
        for col in all_new_columns:
            record_subgroup[col] = subgroup_df[col].max() if col in subgroup_df.columns else "N/A"
        summary_records.append(record_subgroup)

    # ----------------------------------------------------------------
    # 4) Generar tabla de resumen de valores máximos
    # ----------------------------------------------------------------
    summary_df = pd.DataFrame(summary_records)
    print("\nSummary of max values by report:\n")
    print(summary_df.to_string(index=False))
    summary_df.to_excel(os.path.join(output_folder, "REPORT_MAX_SUMMARY.xlsx"), index=False)

    
        
def split_group_into_subgroups(group):
    """
    Función definida por el usuario para dividir una cadena de grupo en subgrupos.
    Implementa esta función según el formato específico de tu cadena de grupo.
    """
    # Implementación de ejemplo (modificar según sea necesario)
    return group.split('_')

def sanitize_sheet_name(name, existing_names=None):
    """
    Sanitiza el nombre de la hoja para cumplir con las reglas de nombres de hojas de Excel.
    Elimina o reemplaza caracteres inválidos, recorta el nombre a 31 caracteres,
    y asegura la unicidad si `existing_names` es proporcionado.

    Args:
        name (str): Nombre original de la hoja.
        existing_names (set, optional): Conjunto de nombres ya usados en el workbook.

    Returns:
        str: Nombre sanitizado y único para la hoja.
    """
    invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
    for char in invalid_chars:
        name = name.replace(char, '_')
    
    name = name.strip().rstrip('.')  # Eliminar espacios y puntos al final
    
    name = name[:31]  # Recortar a 31 caracteres
    
    # Asegurar unicidad
    if existing_names is not None:
        original_name = name
        counter = 1
        while name in existing_names:
            suffix = f"_{counter}"
            if len(original_name) + len(suffix) > 31:
                name = original_name[:31 - len(suffix)] + suffix
            else:
                name = original_name + suffix
            counter += 1
    return name

def format_file_id(file_id):
    """
    Formatea el File_id para evitar puntos innecesarios en los nombres de las hojas.
    Si el File_id es un número flotante sin parte decimal, lo convierte a entero.
    Si tiene decimales significativos, reemplaza el punto por un guión bajo.

    Args:
        file_id (int, float, str): El File_id original.

    Returns:
        str: File_id formateado como cadena.
    """
    if isinstance(file_id, float):
        if file_id.is_integer():
            return str(int(file_id))
        else:
            return str(file_id).replace('.', '_')
    return str(file_id)

if __name__ == "__main__":
    import multiprocessing 
    
    initime = ""#datetime.now().strftime('_%H_%M')
    os.system(f"title Plotting Tool v16 - By Luciano Roco - {os.getcwd().split(os.sep)[-2]}/{os.getcwd().split(os.sep)[-1]} - {config_excel}" )
    os.system("mode con:cols=160 lines=30")
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.GetStdHandle(-11)
    buffer_size = ctypes.wintypes._COORD(160, 1000) 
    kernel32.SetConsoleScreenBufferSize(handle, buffer_size)
    
    os.system("")
    try: 
        os.remove("figcounter_DELETE.txt")
    except:pass
    # try:
    #     os.remove(working_dir + output_folder +os.sep+ f"p1_gTot_{initime}_tTot.csv")
    #     os.remove(working_dir + output_folder +os.sep+ f"S52513_Report.csv")
    #     os.remove(working_dir + output_folder +os.sep+ f"S5255_Report.csv")
    # except Exception as e:
    #     if "The system cannot find" not in str(e):
    #         print(e)

    print(f"""{color.CYAN}
=========================================================================================================================================================
                                                            Plotting Tool {color.RED}v16{color.GREEN} 
                                                            By Luciano Roco {color.CYAN}
========================================================================================================================================================= {color.END}\n""".center(10))
    
    print(f"{color.WARNING}New features:")
    print(f"{color.END}         - General:{color.WARNING}  Report Generation{color.END} is working now. It wont skip any page and you can find all the assessment times there {color.WARNING}{color.END}")   
    print(f"{color.END}         - Settle Time:{color.WARNING}  2% MVA errorbands is being calculated internally now {color.END} - Use as'Yes P - PlantMVA'{color.END}")   
    print(f"{color.END}         - Commencement Time:{color.WARNING}  Now supported in v15 - Add the following column:  {color.END} {color.END}") 
    print(f"{color.END}         -                   {color.WARNING}  Templates >> {color.WARNING}Commencement time  {color.END} >> Only supports {color.WARNING}time{color.END} as value to run it{color.END}")
    print(f"{color.END}                             {color.END} Works with the Settling/Rising Times From  >> to times {color.END}")
    print(f"{color.END}         - Function (Plots>Extras):{color.WARNING}  ORT{color.GREEN}(GRID_V_Legend, POC_V_Legend , POC_Q_Legend, Freq) {color.END} - Only one assessment per page.. Sorry :)")
    print(f'{color.END}         - Function (Plots>Extras):{color.WARNING}  split_pdf{color.GREEN}(File_Id, Extra_Text, optional: 1 or 0){color.END} will create one more pdf file using only that file ID (for DMAT){color.END} ')
    print(f"{color.END}                                   {color.END}  opt number 1: add text to the end of pdf_Group name, 0: at the beggining (Default is 0) {color.END}")
    
    print(f"{color.WARNING}---------------------------------------------------------------------------------------------------------------------------------------------------------{color.END}")
    
    if os.path.isfile("D:\LucianoRoco\Scripts\PSSE3_DONTDELETE.txt"):
        print(f"{color.BLUE}PSSE 3 Server detected {color.WARNING} --> Reducing Performance{color.END}")
        average_cpu_per_page = 4 #percent
        cpu_charge_lim = 90
        commmit_charge_lim = 0.75
        ncoresmin = 8
        print(f"{color.WARNING}---------------------------------------------------------------------------------------------------------------------------------------------------------{color.END}")
    else:
        average_cpu_per_page = 2 #percent
        cpu_charge_lim = 90
        commmit_charge_lim = 0.85
        ncoresmin = 10
    
    while 1:
        try:
            print(f"{color.END}Reading templates..{color.END}")
            sheet_files = to_df("Files")
            sheet_templates = to_df("Templates")
            sheet_plots = to_df("Plots")
            break
        except:
            print(f"{color.RED}Error. I cannot read the excel. Please avoid clicking or changing cell values until it starts plotting{color.END}")
            print(f"{color.WARNING}Trying again in 2s...{color.END}")
            time.sleep(2)
            
    sheet_plots["X_Label"] = sheet_plots["X_Label"].fillna("Time (s)")
    sheet_plots["Title 1"] = sheet_plots["Title 1"].fillna("__")
    sheet_plots["Title 2"] = sheet_plots["Title 2"].fillna("__")
    sheet_plots["Title 3"] = sheet_plots["Title 3"].fillna("__")
    sheet_plots["Title 4"] = sheet_plots["Title 4"].fillna("__")
    sheet_plots["Recovery Times From"] = sheet_plots["Recovery Times From"].fillna(0)
    sheet_plots["Recovery_Times_to"] = sheet_plots["Recovery_Times_to"].fillna(sheet_plots["X_max"])
    sheet_plots["SetRise_Times_to"] = sheet_plots["SetRise_Times_to"].fillna(sheet_plots["X_max"])
    sheet_plots["X_min zoom"] = sheet_plots["X_min zoom"].fillna("No")
    sheet_plots["X_max zoom"] = sheet_plots["X_max zoom"].fillna("No")
    sheet_plots["Create HTML"] = sheet_plots["Create HTML"].fillna("No")
    #sheet_plots = sheet_plots.dropna(subset=['Page'])
    sheet_plots['Page'] = sheet_plots['Page'].fillna(0)
    try:
        sheet_plots['pdf_name'] = sheet_plots.apply(lambda row: f"p{int(row['Page'])}_g{row['pdf Group']}_t{row.name}", axis=1)
    except:
        print(f"{color.RED}Ups, seems like there is nothing to plot. Please add Yes to the Run_pdf column{color.END}")
        exit()
    dir_files = sheet_files["Directory"]
    name_files = sheet_files["Filename (csv)"].str.strip().str.replace(".csv","")
    sheet_files["loc"] = dir_files + os.sep + name_files + ".csv"
    
    sheet_files["Resample_to"] = sheet_files["Resample_to"].fillna(0.123456)
    nonexi = 0
    


    print("Checking files..")
    out_files = []
     # Create a dictionary for faster lookup. Ty for the tip Marty :)
    plot_dict = {(row['pdf Group'], row['Page']): row for _, row in sheet_plots.iterrows() if row['Run_pdf'] == 'Yes'}
    average_size_per_page,files_not_found = calculate_total_size_and_average(sheet_files, plot_dict)
    commit_charge_perc,commit_charge_GB,commit_charge_GB_tot = get_commit_charge()
    
    cpu_charge = psutil.cpu_percent()
    
    if average_size_per_page > 6:
        ncoresmin = 1
        print(f"{color.WARNING}Files are too big, reducing performance to save memory{color.END}")
    elif average_size_per_page > 3:
        ncoresmin = 5  
        print(f"{color.WARNING}Files are too big, reducing performance to save memory{color.END}")
        
        
    if average_size_per_page == 0:
        average_size_per_page = 1
        print(f"{color.WARNING}Ups. Seems like there are no files to plot{color.END}")
    try:
        ncoresmax_memory = int((commit_charge_GB_tot*commmit_charge_lim-commit_charge_GB)/average_size_per_page)
    except:
        ncoresmax_memory = ncoresmin
    
    ncoresmax_cpu = int((cpu_charge_lim-cpu_charge)/average_cpu_per_page)
    
    
        
    # if ncoresmax_memory >= 35:
    #     ncoresmax_memory = 35
        
    if files_not_found == 1:
        print("")
        input(f"{color.WARNING}>>>>>>>>{color.GREEN}Press {color.WARNING}ENTER{color.GREEN} to continue{color.END}")
    #print(f"\nAverage file size per page: {average_size_per_page:.4f} GB")
    #print(f"Commit charge: {commit_charge_perc:.2f}% ({commit_charge_GB:.2f} GB) >>> Available until {commmit_charge_lim*100}%: {commit_charge_GB_tot*commmit_charge_lim-commit_charge_GB:.2f} GB  >>> {color.WARNING}Max Plots limited by memory: {ncoresmax_memory}{color.END}")
    #print(f"CPU usage: {cpu_charge:.2f}% >>> Available until {cpu_charge_lim}%: {cpu_charge_lim-cpu_charge:.2f} %  >>> {color.WARNING}Max Plots limited by CPU: {ncoresmax_cpu}{color.END}")
    #print(f"{color.WARNING}Min Plots: {ncoresmin}{color.END}")
    print()
    ncores = max(min(ncoresmax_cpu,ncoresmax_memory),ncoresmin)

    print(f"{color.CYAN}Plots in parallel =  {ncores}{color.END}")
    print(f"{color.WARNING}---------------------------------------------------------------------------------------------------------------------------------------------------------\n{color.END}")
    
        
    
    
    
    
    
    #if psutil.cpu_percent() <= 80:
    #    ncores = ncoresmax #DONT SET TOO HIGH>> Could crash the server.. 
    #else:
    #    print(f"{color.WARNING}Ups.. Server is busy: Using less CPU power {color.RED}=({color.END}")
    #    ncores = ncoresmin
    
    manager = multiprocessing.Manager()
    results_list = manager.list()
    
    npdfs = []
    split_extras = 0
    initial_len = range(len(sheet_plots["Run_pdf"]))
    for npdf in initial_len:
        if sheet_plots["Run_pdf"][npdf] == "Yes":
            npdfs.append((npdf, sheet_plots, sheet_templates, sheet_files,initime,results_list))
            try:
                split_extras = extract_split_pdf(sheet_plots["Extras"][npdf])

                
                if len(split_extras) != 0:
                    for i in split_extras:
                        Id = i[0]
                        add_text = i[1]
                        add_end = i[2]
                        
                        sheet_plots_split = sheet_plots.copy()
                        sheet_files_split = sheet_files.copy()
                        
                        oldpdf = sheet_plots["pdf Group"].iloc[npdf]
                        if str(add_end) == "1":
                            sheet_plots_split.iloc[npdf, sheet_plots_split.columns.get_loc('pdf Group')] +=  "_" + str(add_text)
                            sheet_files_split.loc[sheet_files_split['pdf Group'].str.contains(oldpdf, na=False) & sheet_files_split['Id'].astype(str).str.contains(Id, na=False), 'pdf Group'] += "_" + str(add_text)
                        else:
                            sheet_plots_split.iloc[npdf, sheet_plots_split.columns.get_loc('pdf Group')] = str(add_text) + "_" + sheet_plots_split.iloc[npdf, sheet_plots_split.columns.get_loc('pdf Group')]
                            sheet_files_split.loc[sheet_files_split['pdf Group'].str.contains(oldpdf, na=False) & sheet_files_split['Id'].astype(str).str.contains(Id, na=False), 'pdf Group'] = str(add_text) + "_" + sheet_files_split.loc[sheet_files_split['pdf Group'].str.contains(oldpdf, na=False) & sheet_files_split['Id'].astype(str).str.contains(Id, na=False), 'pdf Group']
                        
                        print(f'{color.WARNING}Splitting{color.END} {sheet_plots["pdf Group"][npdf]}, File ID: {Id} {color.WARNING}>>{color.END} {sheet_plots_split.iloc[npdf, sheet_plots_split.columns.get_loc("pdf Group")]}{color.END}')
                        
                        sheet_plots_split['pdf_name'] = sheet_plots_split.apply(lambda row: f"p{int(row['Page'])}_g{row['pdf Group']}_t{row.name}", axis=1)
                        npdfs.append((npdf, sheet_plots_split, sheet_templates, sheet_files_split,initime,results_list))
            except:pass
    #exit()     
    with multiprocessing.Pool(processes=min(ncores,len(npdfs))) as pool:
        try:
            results = pool.map(main_script, npdfs)
        except Exception:
            print(traceback.format_exc())
    
    #pool.join()
    #input("Press ENTER to merge all the groups of pdf :)")
    #merge_pdfs_by_group(output_folder,output_folder,"MERGED")
    
    while 1:
        try:
            print("\n")
            move_pdfs_to_group_folders(output_folder,output_folder)
            n_pdf_files,lastpdffile = merge_pdfs_in_group_folders(output_folder,output_folder)
            break
        except Exception as e:
            
            print(f"{color.RED}ERROR MERGING PDF: {e}{color.END}")
            input("Press ENTER to try to merge all the groups of pdf again :)")
    while 1:
        try:
            Create_Reports(output_folder,results_list)
            break
        except Exception as e:
            print(f"\nERROR CREATING REPORTS: {e}")
            input(f"{color.WARNING}>>>>>>>>{color.GREEN}Press {color.WARNING}ENTER{color.GREEN} to try again :){color.END}")
    
    delete_txt_files(output_folder,["_g","_t"])
    delete_folders(output_folder,["Group"])
    print(f"\n{color.GREEN}Done :){color.END}")
    #os.system(f'explorer "{output_folder}"')
    if n_pdf_files != 1:
        os.startfile(output_folder)
    else:
        os.startfile(lastpdffile)

if profile:
    profiler.disable()
    # Output profiling results to stdout
    s = io.StringIO()
    sortby = 'tottime'
    ps = pstats.Stats(profiler, stream=s).sort_stats(sortby)
    ps.dump_stats("profile_output.prof")
    #subprocess.Popen(["snakeviz", "profile_output.prof"])
    from snakeviz import cli
    cli.main(["profile_output.prof"])
    ps.print_stats(20)  # Print top 20 lines; adjust as needed

    print(s.getvalue())
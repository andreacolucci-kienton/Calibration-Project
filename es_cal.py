import sys
from tabulate import tabulate
import copy
from tkinter.filedialog import askopenfilename
import os
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
os.system('cls')
dataset1_file = askopenfilename()
dataset1_file_name = Path(dataset1_file).stem
dataset1 = open(dataset1_file, "r")
lines_dataset1 = dataset1.readlines()
list_dataset1 = []
i_flg1=0
#---------- FUNCTIONS ---------#
def FESTWERT_parse(line):
    line_split = line.split() #[FESTWERT, CAL]
    name = line_split[0]
    cal = line_split[1]
    return name, cal
def FESTWERTEBLOCK_parse(line):
    line_split = line.split() #[FESTWERTEBLOCK, CAL, N_VALORI]
    name = line_split[0]
    cal = line_split[1]
    n_val = line_split[2]
    return name, cal, n_val
def GRUPPENKENNLINIE_parse(line):
    line_split = line.split() #[GRUPPENKENNLINIE, CAL, N_VALORI]
    name = line_split[0]
    cal = line_split[1]
    n_val = line_split[2]
    return name, cal, n_val
def STUETZSTELLENVERTEILUNG_parse(line):
    line_split = line.split() #[STUETZSTELLENVERTEILUNG, CAL, N_VALORI]
    name = line_split[0]
    cal = line_split[1] 
    return name, cal
""" def KENNLINIE_parse(line):
    line_split = line.split() #[KENNLINIE, CAL, N_VALORI]
    name = line_split[0]
    cal = line_split[1]
    n_val = line_split[2]
    return name, cal, n_val
def KENNFELD_parse(line):
    line_split = line.split() #[KENNFELD, CAL, N_VALORI]
    name = line_split[0]
    cal = line_split[1]
    x = line_split[2]
    y = line_split[3]
    return name, cal, x, y """

def FESTWERT_value(line):
    line_split = line.split() #[WERT, VAL]
    val = line_split[1]
    if("TRUE" in val or "FALSE" in val):
        if("TRUE" in val):
            val = "TRUE"
        else:
            val = "FALSE"
    else:
        try:
            val = "{:.2f}".format(float(val))
        except:
            pass
        finally:
            val = "N.A."
    return val
def FESTWERTEBLOCK_value(line, n_val):
    line_split = line.split() #[FESTWERTEBLOCK, N_VAL]
    i = 0
    val="N.A."
    while i < len(line_split)-1:
        try:
            val + ("{:.2f}".format(float(line_split[i+1]))) + " "
        except:
            pass
        i = i+1
    return val
""" def GRUPPENKENNLINIE_value():
    val="N.A."
    return val """
def STUETZSTELLENVERTEILUNG_value(line):
    line_split = line.split() #[STUETZSTELLENVERTEILUNG, CAL, N_VALORI]
    cal = line_split[1] 
    return cal, n_val
""" def KENNLINIE_value(line):
    line_split = line.split() #[KENNLINIE, CAL, N_VALORI]
    cal = line_split[1]
    n_val = line_split[2]
    return cal, n_val
def KENNFELD_value(line):
    line_split = line.split() #[KENNFELD, CAL, N_VALORI]
    cal = line_split[1]
    x = line_split[2]
    y = line_split[3]
    return cal, x, y """

#---------- DATASET 1 ----------#
cal_dataset1=""
module_dataset1=""
val_dataset1=""
cal_dataset2=""
module_dataset2=""
val_dataset2=""
n = len(lines_dataset1)
i = 1
print()
print("Parsing the first dataset...")
for str_dataset1 in lines_dataset1:
    sys.stdout.write(str(int((i*100)/n))+"%")
    sys.stdout.write("\b\b\b")
    sys.stdout.flush()
    str_dataset1=str_dataset1.lstrip()
    if str_dataset1.startswith("FESTWERT") or str_dataset1.startswith("FESTWERTEBLOCK") or str_dataset1.startswith("GRUPPENKENNLINIE") or str_dataset1.startswith("STUETZSTELLENVERTEILUNG"):
    #if str_dataset1.startswith("FESTWERT") or str_dataset1.startswith("FESTWERTEBLOCK") or str_dataset1.startswith("GRUPPENKENNLINIE") or str_dataset1.startswith("STUETZSTELLENVERTEILUNG") or str_dataset1.startswith("KENNLINIE") or str_dataset1.startswith("KENNFELD"):
        if str_dataset1.startswith("FESTWERT"):
            [name, cal_dataset1] = FESTWERT_parse(str_dataset1)      
        if str_dataset1.startswith("FESTWERTEBLOCK"):
            [name, cal_dataset1, n_val] = FESTWERTEBLOCK_parse(str_dataset1)
        if str_dataset1.startswith("GRUPPENKENNLINIE"):
            [name, cal_dataset1, n_val] = GRUPPENKENNLINIE_parse(str_dataset1)
        if str_dataset1.startswith("STUETZSTELLENVERTEILUNG"):
            [name, cal_dataset1] = STUETZSTELLENVERTEILUNG_parse(str_dataset1)
        """ if str_dataset1.startswith("KENNLINIE"):
            [name, cal_dataset1, n_val] = KENNLINIE_parse(str_dataset1)
        if str_dataset1.startswith("KENNFELD"):
            [name, cal_dataset1, x, y] = KENNFELD_parse(str_dataset1)  """   
    else:
        if i_flg1 == 0:
            i_flg1 = 1
    if str_dataset1.startswith("FUNKTION"):
        module_dataset1=str_dataset1.removeprefix("FUNKTION").lstrip()
    if str_dataset1.startswith("WERT") or str_dataset1.startswith("TEXT"):
        if name == "FESTWERT":
            val_dataset1 = FESTWERT_value(str_dataset1)      
        if name == "FESTWERTEBLOCK":
            val_dataset1 = FESTWERTEBLOCK_value(str_dataset1, n_val) 
        if name == "GRUPPENKENNLINIE":
            val_dataset1 = "N.A."
        if name == "STUETZSTELLENVERTEILUNG":
            val_dataset1 = "N.A."
        """ if name == "KENNLINIE":
            val_dataset1 = "N.A."
        if name == "KENNFELD":
            val_dataset1 = "N.A." """
    if cal_dataset1 != "" and module_dataset1 != "" and val_dataset1 != "":
        dict_temp = {"Calibration":cal_dataset1,"Module":module_dataset1,"Value":val_dataset1}
        cal_dataset1 = ""
        module_dataset1 = ""
        val_dataset1 = ""
        if dict_temp not in list_dataset1:
            list_dataset1.append(dict_temp)
    i = i+1
try:
    del dict_temp
except:
    pass
#---------- DATASET 2 ----------#

dataset2_file = askopenfilename()
dataset2_file_name = Path(dataset2_file).stem
dataset2 = open(dataset2_file, "r")
lines_dataset2 = dataset2.readlines()
list_dataset2 = []
i_flg2=0
n = len(lines_dataset2)
i = 1
print()
print("Parsing the second dataset...")
for str_dataset2 in lines_dataset2:
    sys.stdout.write(str(int((i*100)/n))+"%")
    sys.stdout.write("\b\b\b")
    sys.stdout.flush()
    str_dataset2=str_dataset2.lstrip()
    if str_dataset2.startswith("FESTWERT") or str_dataset2.startswith("FESTWERTEBLOCK") or str_dataset2.startswith("GRUPPENKENNLINIE") or str_dataset2.startswith("STUETZSTELLENVERTEILUNG"):
    #if str_dataset2.startswith("FESTWERT") or str_dataset2.startswith("FESTWERTEBLOCK") or str_dataset2.startswith("GRUPPENKENNLINIE") or str_dataset2.startswith("STUETZSTELLENVERTEILUNG") or str_dataset2.startswith("KENNLINIE") or str_dataset2.startswith("KENNFELD"):
        if str_dataset2.startswith("FESTWERT"):
            [name, cal_dataset2] = FESTWERT_parse(str_dataset2)       
        if str_dataset2.startswith("FESTWERTEBLOCK"):
            [name, cal_dataset2, n_val] = FESTWERTEBLOCK_parse(str_dataset2)
        if str_dataset2.startswith("GRUPPENKENNLINIE"):
            [name, cal_dataset2, n_val] = GRUPPENKENNLINIE_parse(str_dataset2)
        if str_dataset2.startswith("STUETZSTELLENVERTEILUNG"):
            [name, cal_dataset2] = STUETZSTELLENVERTEILUNG_parse(str_dataset2)
        """ if str_dataset2.startswith("KENNLINIE"):
            [name, cal_dataset2, n_val] = KENNLINIE_parse(str_dataset2)
        if str_dataset2.startswith("KENNFELD"):
            [name, cal_dataset2, x, y] = KENNFELD_parse(str_dataset2) """    
    else:
        if i_flg2 == 0:
            i_flg2 = 1
    if str_dataset2.startswith("FUNKTION"):
        module_dataset2=str_dataset2.removeprefix("FUNKTION").lstrip()
    if str_dataset2.startswith("WERT") or str_dataset2.startswith("TEXT"):
        if name == "FESTWERT":
            val_dataset2 = FESTWERT_value(str_dataset2)       
        if name == "FESTWERTEBLOCK":
            val_dataset2 = FESTWERTEBLOCK_value(str_dataset2, n_val) 
        if name == "GRUPPENKENNLINIE":
            val_dataset2 = "N.A."
        if name == "STUETZSTELLENVERTEILUNG":
            val_dataset2 = "N.A."
        """ if name == "KENNLINIE":
            val_dataset2 = "N.A."
        if name == "KENNFELD":
            val_dataset2 = "N.A." """
    if cal_dataset2 != "" and module_dataset2 != "" and val_dataset2 != "":
        dict_temp = {"Calibration":cal_dataset2,"Module":module_dataset2,"Value":val_dataset2}
        cal_dataset1 = ""
        module_dataset1 = ""
        val_dataset1 = ""
        if dict_temp not in list_dataset2:
            list_dataset2.append(dict_temp)
    i = i+1
try:
    del dict_temp
except:
    pass
list_dataset = []
list_dataset = copy.deepcopy(list_dataset1)
list_dataset.extend(list_dataset2)
dataset = []      
n = len(list_dataset)
i = 1
print()
print("Creating the list without duplicates...")                    
for element in list_dataset:
    sys.stdout.write(str(int((i*100)/n))+"%")
    sys.stdout.write("\b\b\b")
    sys.stdout.flush()
    #os.system('cls')
    if element not in dataset:
        dataset.append(element)
    i = i+1
final = []
n = len(dataset)
i = 1
print()
print("Creating the final list...")
for element_dataset in dataset:
    sys.stdout.write(str(int((i*100)/n))+"%")
    sys.stdout.write("\b\b\b")
    sys.stdout.flush()
    #os.system('cls')
    if element_dataset in list_dataset1:
        data1 = 1
    else:
        data1 = 0
    if element_dataset in list_dataset2:
        data2 = 1
    else:
        data2 = 0
    if [data1,data2] != [1,1]:
        if(data1 == 1):
            final.append([element_dataset['Calibration'],element_dataset['Module'],element_dataset['Value'],"-"])
        else:
            final.append([element_dataset['Calibration'],element_dataset['Module'],"-",element_dataset['Value']])
    sys.stdout.flush()
    i = i+1
print()
print("Creating the Excel file...")
headers=["Calibration","Module",dataset1_file_name,dataset2_file_name]
name = dataset1_file_name+"vs"+dataset2_file_name+".xlsx"
path = os.path.join("C:\\Users\\Andrea\\Desktop\\",name)
wb = Workbook()
ws = wb.active
ws.append(headers)
for row in final:
    ws.append(row)
wb.save(path)   
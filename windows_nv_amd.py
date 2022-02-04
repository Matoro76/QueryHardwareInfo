import os
import xml.etree.ElementTree as xml
from multiprocessing import freeze_support
import wmi     #pip install wmi (Can't install on Linux platform)
import psutil  #pip install psutil
import cpuinfo #pip install py-cpuinfo
import datetime
import re
import sys
import logging

def error():
    logging.basicConfig(level = logging.INFO,
    filename = "error.txt",
    filemode = "w",
    format = '[%(asctime)s %(levelname)-8s] %(message)s',
    datefmt = '%Y%m%d %H:%M:%S')

#calculate the size of memory
def get_size(bytes, suffix="B"):
    factor = 1024
    for unit in ["", "K", "M", "G", "T", "P"]:
        if bytes < factor:
            bytes= round(bytes)
            return f"{bytes}{unit}{suffix}"
        bytes /= factor

def remove_character(original, s):
    removed = original.replace(s, "")
    return removed

#first function to collecting data
def find_OS():
    print("Collecting Data ......")
    w = wmi.WMI()
    for obj in w.Win32_OperatingSystem():
        if obj.Status == 'OK':
            os = obj.Caption
            break
    return os

def find_CpuNum():
    w = wmi.WMI()
    num  = 0
    #get CpuNum info
    i = 0
    while(i<len(w.Win32_Processor())):
        if w.Win32_Processor()[i].CpuStatus == 1: 
            num  = num  + 1 
        i = i + 1
    return num 

def find_GraphicsCard_and_CardNum():
    GraphicsCard = ""
    CardNum = 0
    w = wmi.WMI()
    for obj in w.Win32_VideoController():
        if obj.AdapterCompatibility =="Advanced Micro Devices, Inc." or obj.AdapterCompatibility =="NVIDIA" or obj.AdapterCompatibility =="Intel Corporation":
            GraphicsCard = obj.Name
            CardNum = CardNum + 1 
    return GraphicsCard, CardNum

def find_GraphicsDriver():
    w = wmi.WMI()
    Driver = ""
    target = ""
    for obj in w.Win32_VideoController():
        if obj.AdapterCompatibility =="Advanced Micro Devices, Inc." or obj.AdapterCompatibility == "Intel Corporation":
            Driver = obj.DriverVersion
        #Transfer DriverVersion indicated digit 
        elif obj.AdapterCompatibility =="NVIDIA":
            if os.path.isfile(r"C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe"):
                os.system(r'C:\"Program Files"\"NVIDIA Corporation"\NVSMI\nvidia-smi.exe -q > ./temp.log"')
            elif os.path.isfile(r"C:\Program Files (x86)\NVIDIA Corporation\NVSMI\nvidia-smi.exe"):
                os.system(r'C:\"Program Files (x86)"\"NVIDIA Corporation"\NVSMI\nvidia-smi.exe -q > ./temp.log"')
            else:
                os.system("nvidia-smi -q > temp.log")
            state = 0 
            for component in os.listdir('./'):
                if component.startswith("temp"):
                    with open(component) as log:
                        log = log.readlines()
                    state = 1
            if state == 1:
                for line in log:
                    if "Driver Version" in line:
                        target = remove_character(line, " ")
                        target = remove_character(target, "\r")
                        target = remove_character(target, "\n")
                        target = remove_character(target, "\t")

                a = re.search(r"DriverVersion:(.+)", target).group(1)
                Driver =str(a)
            else:
                logging.error("Failed to create temp.log file !")
    return Driver    

def find_VRAM():
    w = wmi.WMI()
    vram = 0
    found = False
    target = ""
    for obj in w.Win32_VideoController():
        if obj.AdapterCompatibility =="NVIDIA" :
        #Transfer DriverVersion indicated digit 
            state = 0 
            for component in os.listdir('./'):
                if component.startswith("temp"):
                    with open(component) as log:
                        log = log.readlines()
                    state = 1
            if state == 1:
                for line in log:
                    if found == True:
                        target = remove_character(line, " ")
                        target = remove_character(target, "\r")
                        target = remove_character(target, "\n")
                        target = remove_character(target, "\t")
                        break
                    if "FB Memory Usage" in line:
                        found = True

                a = re.search(r"Total:(.+)", target).group(1)
                a = a.replace("MiB", "")
                vram =round(int(a) / 1024)
                os.remove("temp.log") 
            else:
                logging.error("Failed to create temp.log file !")
            break

        elif obj.AdapterCompatibility == "Advanced Micro Devices, Inc." or obj.AdapterCompatibility == "Intel Corporation":
            print("Creating XML File......")
            timea = datetime.datetime.now()
            os.system("dxdiag /dontskip /whql:off /64bit /x dxdiag.xml")
            #loading OS info
            timeb = datetime.datetime.now()
            if not os.path.isfile("dxdiag.xml"):
                logging.error("Failed to create XML file !")
                sys.exit(1)
            print("Time of creating XML file: ", timeb - timea)
            #Has been created XML

            root = xml.parse('dxdiag.xml').getroot()
            d = root.find('DisplayDevices')
            for display in d.findall('DisplayDevice'):
                #if node <DedicatedMemory> doesn't exist, remain the vram spec blank
                DedicatedMemory = display.find("DedicatedMemory")
                if DedicatedMemory is None:
                    logging.error("There is no <DedicatedMemory> node in dxdiag.xml !")
                    continue
                else:
                    v = DedicatedMemory.text
                    v_int = int(v[:-3])
                    if v_int > vram :
                        vram = v_int
                        vram = round (vram / 1024)
            if os.path.isfile("./dxdiag.xml"):
                os.remove("dxdiag.xml")
            break

        else:
            logging.error("GPU AdapterCompatibility are not AMD, NVIDIA, or Intel.") 
    return vram

def find_cpu():
    cpu = ""
    cpu = cpuinfo.get_cpu_info()['brand_raw']
    if  not cpu.endswith("GHz"):
        hz = cpuinfo.get_cpu_info()['hz_advertised_friendly']
        hz = hz.replace("GHz", "")
        hz = hz.replace(" ", "")
        hz = float(hz)
        hz = "{:.2f}".format(hz)
        hz = str(hz)
        cpu = cpu +" @ " + hz +" GHz"
    return cpu

def find_ram():
    ram = 0 
    svmem=psutil.virtual_memory()
    ram = get_size(svmem.total)
    return ram

def find_core_and_logic():
    #get core & logic info
    Processor = ""
    Core = int(psutil.cpu_count(logical = False)/CpuNum)
    Logic = int(psutil.cpu_count(logical = True)/CpuNum)
    Processor = str(Core)+"/"+str(Logic)
    return Processor

def write_into_txt():
    #open file
    name ="Hardware_Spec_"
    current_date = datetime.datetime.today()
    current_date_string = current_date.strftime("%Y%m%d_%H%M%S")
    extension = ".txt"
    file_name =name + current_date_string + extension
    f = open(file_name, 'w')

    #Graphics Cards
    f.write("Graphics Cards=")
    f.write(str(GraphicsCard))
    #Card Num
    f.write("\nCard Num=")
    f.write(str(CardNum))

    #VRAM
    f.write("\nVRAM=")
    if vram == 0:
        f.write("")
    else:
        f.write(str(vram))
        f.write("GB")

    #Graphics Driver
    f.write("\nGraphics Driver=")
    f.write(str(Driver))

    #OS
    f.write("\nOS=")
    f.write(str(OperatingSystem))

    #CPU
    f.write("\nCPU=")
    f.write(str(cpu))

    #CPU Num
    f.write("\nCPU Num=")
    f.write(str(CpuNum))

    #Core/Logic Processor
    f.write("\nCore/Logic Processor=")
    f.write(Processor)
    f.write("\nRAM=")
    f.write(str(ram))
    #STOP writing and close the file
    f.close()
    print("Done !")


if __name__ == '__main__':
    #make cmd won't throw run_time_error(), because the time of creating xml might be a little bit long
    freeze_support()
    error()
    OperatingSystem = find_OS()
    CpuNum = find_CpuNum()
    GraphicsCard, CardNum = find_GraphicsCard_and_CardNum()
    Driver = find_GraphicsDriver()
    vram = find_VRAM()
    cpu = find_cpu()
    ram = find_ram()
    Processor = find_core_and_logic()
    write_into_txt()

    
    
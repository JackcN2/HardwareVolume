from __future__ import print_function
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume
import time
import serial
import os

file_path = "config.txt"
if not os.path.isfile(file_path):
    with open(file_path, 'w') as f:
        f.write("Master\n")
        f.write("Discord.exe\n")
        f.write("example.exe\n")
        f.write("otherexample.exe\n")
        f.write("chrome.exe\n")
        f.write("COM3\n")
        f.write("9600\n")
    
        
with open(file_path, 'r') as file:
    lines = file.readlines()
    if len(lines) >= 6:
        # Set the 6th line to the string, stripping any trailing newline characters
                comport = lines[5].strip()
    if len(lines) >= 7:
        # Set the 6th line to the string, stripping any trailing newline characters
                baud = lines[6].strip()
                baud = int(baud)
                

ser = serial.Serial(comport, baud)

#inputdata = inputdata[:-3]



#testing values only

#


def main():
    vol1, vol2, vol3, vol4, vol5 = None, None, None, None, None 
    knob1, knob2, knob3, knob4, knob5 = read_config(file_path)


    
    serial_data = ser.readline().decode('utf-8').strip()
    int_list = []
    if ser.in_waiting > 0:
        for item in serial_data.split('|'):
            if item.isdigit():  # Check if the item is a valid integer
                int_list.append(int(item))

            if len(int_list) >= 5:
                vol1, vol2, vol3, vol4, vol5 = int_list[:5]
                vol1 = vol1 / 1023
                vol2 = vol2 / 1023
                vol3 = vol3 / 1023
                vol4 = vol4 / 1023
                vol5 = vol5 / 1023
                
    sessions = AudioUtilities.GetAllSessions()
    #prevent erroring
    if vol1:
        for session in sessions:
            
        
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)

            if knob1 == "Master":
                volvar = vol1
                
                master(volvar)
                
            elif session.Process and session.Process.name() == knob1:
               
                volume.SetMasterVolume(vol1, None)
                
            if knob2 == "Master":
                volvar = vol2
                
                master(volvar)
                
            elif session.Process and session.Process.name() == knob2:
               
                volume.SetMasterVolume(vol2, None)
                
            if knob3 == "Master":
                volvar = vol3
                
                master(volvar)
                
            elif session.Process and session.Process.name() == knob3:
               
                volume.SetMasterVolume(vol3, None)
            if knob4 == "Master":
                volvar = vol4
                
                master(volvar)
                
            elif session.Process and session.Process.name() == knob4:
               
                volume.SetMasterVolume(vol5, None)
            if knob5 == "Master":
                
                volvar = vol5
                
                master(volvar)
                
            elif session.Process and session.Process.name() == knob5:
                
                volume.SetMasterVolume(vol5, None)


            
def master(volvar):
    # Get the default audio device (speakers)
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMasterVolumeLevelScalar(volvar, None)
    

def read_config(file_path):
    # Initialize variables for the first five knobs
    knob1, knob2, knob3, knob4, knob5 = None, None, None, None, None
    
    try:
        with open(file_path, 'r') as file:
            # Read knobs and assign them to variables
            knobs = file.readlines()
            if len(knobs) >= 1:
                knob1 = knobs[0].strip()
            if len(knobs) >= 2:
                knob2 = knobs[1].strip()
            if len(knobs) >= 3:
                knob3 = knobs[2].strip()
            if len(knobs) >= 4:
                knob4 = knobs[3].strip()
            if len(knobs) >= 5:
                knob5 = knobs[4].strip()
            if len(knobs) >= 6:
        # Set the 6th line to the string, stripping any trailing newline characters
                comport = knobs[5].strip()
                

            
    except FileNotFoundError:
        print(f"Error, config file moved")
    return knob1, knob2, knob3, knob4, knob5

    

while True:
    time.sleep(0.1)
    main()


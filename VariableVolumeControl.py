#imports
from __future__ import print_function
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume
import time
import serial
import os



#Setup the file refrenced throuout
file_path = "config.txt"
#Create the config file on the first time the program is run
if not os.path.isfile(file_path):
    with open(file_path, 'w') as f:
        f.write("Master\n")
        f.write("Discord.exe\n")
        f.write("example.exe\n")
        f.write("otherexample.exe\n")
        f.write("chrome.exe\n")
        f.write("COM3\n")
        f.write("9600\n")
        f.write("Lines 1-5 are the program that each knob refrences, line 6 is the serial port and line 7 is the baud rate")
    
#read the port  and baud lines of the file        
with open(file_path, 'r') as file:
    lines = file.readlines()
    if len(lines) >= 6:
        # Set the 6th line to the string, stripping any trailing newline characters
                comport = lines[5].strip()
    if len(lines) >= 7:
        # Set the 6th line to the string, stripping any trailing newline characters
                baud = lines[6].strip()
                baud = int(baud)

#open the serial monitor with the data we just got               
try:
    ser = serial.Serial(comport, baud)
    
#Make the user restart the program if the monitor cant be opened
except:
    print("no device detected")
    input("connect a  device and restart the program")
    exit()



#the function to run continuously
def main():
    #Initialize the first volume to ensure the program does not run out of order and break
    vol1 = None
    #read the config file for what programs corrispond to what knobs
    knob1, knob2, knob3, knob4, knob5 = read_config(file_path)


    #read the lines from serial
    serial_data = ser.readline().decode('utf-8').strip()
    int_list = []
    if ser.in_waiting > 0:
        for item in serial_data.split('|'):
            if item.isdigit():  # Check if the item is a valid integer
                int_list.append(int(item))
            #Convert the volume values between the ones outputted by arduino and  the ones accepted by pycaw
            if len(int_list) >= 5:
                vol1, vol2, vol3, vol4, vol5 = int_list[:5]
                vol1 = vol1 / 1023
                vol2 = vol2 / 1023
                vol3 = vol3 / 1023
                vol4 = vol4 / 1023
                vol5 = vol5 / 1023
    #enter the audio settings         
    sessions = AudioUtilities.GetAllSessions()
    
    #prevent erroring
    if vol1:
        for session in sessions:
            
            
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        #for each knob, set the volume to the serial data we recive
            #if the knob is set to master, call the master function
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


#called if a knob is set to master            
def master(volvar):
    #set the master volume to what the knob is set to
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMasterVolumeLevelScalar(volvar, None)
    

def read_config(file_path):
    # Initialize variables for the first five knobs
    knob1, knob2, knob3, knob4, knob5 = None, None, None, None, None
    
    try:
        #read each line and set it to a knob
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

            
    except FileNotFoundError:
        print(f"Error, config file moved")
    return knob1, knob2, knob3, knob4, knob5

    
#run the program every 10th of a second 
while True:
    time.sleep(0.1)
    main()


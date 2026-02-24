import win32com.client
import pythoncom
import math
import os
import time

#Solidworks API
swDocPART = 1   #1 - part / 2 - asembly / 3 - drawing
swSaveAsCurrentVersion = 0
swSaveAsOptions_Silent = 1
swSaveAsOptions_Copy = 2

#Text colors
class Color:
    RED = '\033[91m'
    GREEN = '\033[92m'
    MAGENTA = '\033[95m'
    ORANGE = '\033[93m'
    RESET = '\033[0m'

def run_aero_sweep():
    #configuration
    angleName = input (f"{Color.MAGENTA}Enter base angle name: {Color.RESET}")
    baseFileName = input (f"{Color.MAGENTA}Enter base file name: {Color.RESET}")
    clenFileName = baseFileName[:-7]
    inputAngles = input(f"{Color.MAGENTA}Enter new angles separated by space: {Color.RESET}").split()
    newAngles = [float(angle) for angle in inputAngles]

    workingDirectory = os.getcwd()
    baseDirectory = os.path.join(workingDirectory, baseFileName)

    if not os.path.exists(baseDirectory):
        print(f"{Color.RED}Error: file not found {baseDirectory}{Color.RESET}")
        return
    
    print (f"{Color.ORANGE}Launching SolidWorks in the background...{Color.RESET}")
    swApp = win32com.client.Dispatch("SldWorks.Application")
    swApp.Visible = False    #Change to true to show SolidWorks while working

    try:
        print(f"{Color.ORANGE}Opening base model: {baseFileName}...{Color.RESET}")

        argErr = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        argWarn = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

        Model = swApp.OpenDoc6(baseDirectory, swDocPART, 1, "", argErr, argWarn)   #(directory, file type, silent mode, configuration, error code, error code)

        if Model is None:
            print (f"{Color.RED}Error: fialed to open file{Color.RESET}")
            return
        
        swDimension = Model.Parameter(angleName)
        if swDimension is None:
            print(f"{Color.RED}Critical error: angle {angleName} not found, check the name in SolidWorks{Color.RESET}")
            return
        
        print(f"{Color.ORANGE}Generationg variants...{Color.RESET}")

        for angleDeg in newAngles:
            angleRad = angleDeg * (math.pi / 180.0)     #Conversion to radians
            swDimension.SystemValue = angleRad
            Model.EditRebuild3

            newFileName = f"{clenFileName}_{angleDeg}deg.sldprt"
            saveDirectory = os.path.join(workingDirectory, newFileName)
            saveOptions = swSaveAsOptions_Silent + swSaveAsOptions_Copy

            success = Model.SaveAs3(saveDirectory, swSaveAsCurrentVersion, saveOptions)

            if success == 0:
                print(f"{Color.GREEN}Generation successful: {newFileName}{Color.RESET}")
            
            else:
                print(f"{Color.RED}File save error{Color.RESET}")

            time.sleep(0.5) #Pause for file handling

        print(f"{Color.GREEN}Variants generation successful, new files should be present in folder{Color.RESET}")
    
    except Exception as e:
        print(f"\n{Color.RED}Unexpecter error: {e}{Color.RESET}")

    finally:
        if 'Model' in locals() and Model is not None:
            swApp.CloseDoc(Model.GetTitle)
            print(f"{Color.ORANGE}Base model file closed{Color.RESET}")

if __name__ == "__main__":
    run_aero_sweep()

input(f"{Color.MAGENTA}Press enter to close...{Color.RESET}")
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

def run_aero_sweep():
    #configuration
    angleName = input ("Enter base angle name: ")
    baseFileName = input ("Enter base file name: ")
    clenFileName = baseFileName[:-7]
    inputAngles = input("Enter new angles separated by space: ").split()
    newAngles = [float(angle) for angle in inputAngles]

    workingDirectory = os.getcwd()
    baseDirectory = os.path.join(workingDirectory, baseFileName)

    if not os.path.exists(baseDirectory):
        print(f"Error: file not found {baseDirectory}")
        return
    
    print ("Launching SolidWorks in the background...")
    swApp = win32com.client.Dispatch("SldWorks.Application")
    swApp.Visible = False    #Change to true to show SolidWorks while working

    try:
        print(f"Opening base model: {baseFileName}...")

        argErr = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        argWarn = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

        Model = swApp.OpenDoc6(baseDirectory, swDocPART, 1, "", argErr, argWarn)   #(directory, file type, silent mode, configuration, error code, error code)

        if Model is None:
            print ("Error: fialed to open file")
            return
        
        swDimension = Model.Parameter(angleName)
        if swDimension is None:
            print(f"Critical error: angle {angleName} not found, check the name in SolidWorks")
            return
        
        print("Generationg variants...")

        for angleDeg in newAngles:
            angleRad = angleDeg * (math.pi / 180.0)     #Conversion to radians
            swDimension.SystemValue = angleRad
            Model.EditRebuild3

            newFileName = f"{clenFileName}_{angleDeg}deg.sldprt"
            saveDirectory = os.path.join(workingDirectory, newFileName)
            saveOptions = swSaveAsOptions_Silent + swSaveAsOptions_Copy

            success = Model.SaveAs3(saveDirectory, swSaveAsCurrentVersion, saveOptions)

            if success == 0:
                print(f"Generation successful: {newFileName}")
            
            else:
                print(f"File save error")

            time.sleep(0.5) #Pause for file handling

        print("Variants generation successful, new files should be present in folder")
    
    except Exception as e:
        print(f"\nUnexpecter error: {e}")

    finally:
        if 'Model' in locals() and Model is not None:
            swApp.CloseDoc(Model.GetTitle)
            print("Base model file closed")

if __name__ == "__main__":
    run_aero_sweep()

input("Press enter to close...")
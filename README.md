# 🏎️ AoA Tool

A Python-based script for automating angle of attack (AoA) changes in SolidWorks models. This tool significantly speeds up the process of preparing geometry for CFD simulations, eliminating the need to manually modify and save each wing variant.

## ✨ Main Features
* **Interactive CLI:** Users are prompted to input the base angle name, base file name, and a list of target angles directly in the terminal.
* **CAD Automation:** Direct connection to the SolidWorks API (runs in silent mode in the background).
* **Dynamic Geometry Change:** Automatically converts user-defined degrees to radians and applies them to the specified sketch dimension.
* **Safe Save:** Modified variants are automatically rebuilt and saved as new independent '.SLDPRT' files (Save As Copy), protecting your base file.

## 🛠️ System Requirements
To run the script, you must have:
1. **SolidWorks** installed (the script uses the native engine).
2. **Python** (version 3.x) installed.
3. An external COM communication library. To install it, open your terminal and run: `pip install pywin32`

## 🚀 Quick Start
1. Download the 'AoA Tool.py' script file.
2. Place the script in the exact same folder as your base wing '.SLDPRT' file.
3. Open a terminal in that folder and run the command: `python "AoA Tool.py"` or simply double click the file.
4. Follow the on-screen prompts to enter your base file name, angle dimension name, and desired angles.
   * *💡 Note:* If you are testing the script using the provided **`GT3Wing.SLDPRT`** example file from this repository, the default angle name you should enter is: **`AoA_Rear_Wing@AirfoilSketch`**.
5. The completed '.SLDPRT' files will appear in the same directory.

## 📝 Known Issues
* The script currently assumes the base file extension is exactly 7 characters long (e.g., '.sldprt' or '.SLDPRT') due to hardcoded string slicing.
* Ensure SolidWorks is not actively locking the base file in another window before running the script.

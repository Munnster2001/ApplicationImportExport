Application Import/Export with winget
This PowerShell script provides a graphical user interface (GUI) to simplify the process of exporting and importing your installed Windows applications using the winget command-line tool. It's a handy way to back up your application list and restore it on another machine.

Features
Export Applications: Generates a CSV file and a JSON manifest file listing all applications detected by winget list. The CSV is easy to read, while the JSON is the format required for the import function.

Import Applications: Uses the JSON manifest file to automatically install all applications that were previously exported.

User-friendly GUI: A simple interface allows you to browse for a file path, and provides a progress bar and verbose output to track the process.

Robust Parsing: The script uses a regular expression to reliably parse the text output of winget list, ensuring that application names with spaces and other details are captured correctly.

How to Use
Save the Script: Save the code as a .ps1 file (e.g., winget-gui.ps1).

Run with PowerShell: Right-click the script and select Run with PowerShell.

Export:

Click the Browse button to choose a location and file name for your export files.

Click Export Applications. The script will create both a .json manifest and a .csv file.

Import:

Place the .json file you created on the machine you want to set up.

Click the Browse button and select the .json file.

Click Import Applications. winget will then begin installing all the listed applications.

Script Details
The script is a self-contained PowerShell GUI built with WPF (Windows Presentation Foundation) and runs without needing any additional modules.

buttonExport.Add_Click:

Executes winget list to get a list of all applications.

Uses a regular expression (^(?<Name>.*?)\s{2,}(?<Id>\S+)\s{2,}(?<Version>\S*)\s{2,}(?<Source>\S+)$) to parse the output and capture the Name, Id, Version, and Source of each application.

Generates a formatted CSV file for easy viewing.

Generates a winget-compatible JSON manifest file from the same parsed data for use with the import function.

buttonImport.Add_Click:

Reads the path to the JSON manifest file from the text box.

Executes the command winget import -i [file_path] to install the applications listed in the manifest.

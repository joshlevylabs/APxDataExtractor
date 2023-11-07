# APx Automation Script

This script automates the process of gathering data from APx software and exports the checked data into an Excel file.

## Prerequisites

- Python 3.7+
- Audio Precision APx500 Software
- .NET framework

## Getting Started

### Setting up a virtual environment:

1. Navigate to the project's root directory.
2. Create a virtual environment:
   python -m venv venv
3. Activate the virtual environment:
   - On Windows:
      .\venv\Scripts\activate
   - On macOS and Linux:
      source venv/bin/activate

### Installing necessary packages:

Once the virtual environment is activated, you can install the necessary packages using:

pip install openpyxl
pip install pythonnet

## Running the Script
After setting up the virtual environment and installing the packages, you can run the script as follows:

python script_name.py --filename <YourExcelFileName.xlsx>
Replace script_name.py with the name of your script and <YourExcelFileName.xlsx> with your desired Excel file name.

### Arguments:
--filename or -name: Name of the Excel file (required).
Notes
Ensure the paths in the script for Audio Precision APx500 Software DLLs are correct and match the version installed on your machine.

## Troubleshooting
If you encounter any DLL related issues, ensure the .NET framework is installed and the references to the DLL files in the script are accurate.
Check the logs for any errors or warnings for more detailed troubleshooting.
## License
Add any licensing information here.

## Author
Joshua Levy
Josh Levy Labs
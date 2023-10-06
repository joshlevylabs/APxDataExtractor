AP Sequence Runner - README.txt
======================================================================

ABOUT THE PROGRAM:
------------------
AP Sequence Runner is a comprehensive software tool designed to automate, enhance, and simplify the execution of AP sequences and the retrieval of results. By offering an intuitive GUI, the software caters to a spectrum of users - from those just getting started with AP sequences to experienced professionals who require advanced functionalities.

INSTALLING WITH VIRTUAL ENVIRONMENT:
------------------------------------
1. If you haven't already, ensure you have Python installed on your machine.
2. Navigate to the root directory of the cloned repository.
3. Create a virtual environment: python -m venv apxenv
4. 4. Activate the virtual environment:
- On Windows:
  ```
  .\apxenv\Scripts\activate
  ```
- On macOS and Linux:
  ```
  source apxenv/bin/activate
  ```
5. Once the virtual environment is active, install the dependencies using: pip install -r requirements.txt
6. After installation, you can run the script within this virtual environment.


KEY FEATURES:
-------------
1. SELECTION AND EXECUTION:
   - Choose specific signal paths, measurements, and results to be executed within the AP sequence.
   - Efficiently execute the entire AP sequence or specific components based on the user's preference.

2. RESULT HANDLING:
   - View a comprehensive summary of results after sequence execution.
   - Easily pinpoint failed, passed, or errored tests.
   - Export results to Excel format for further analysis and reporting.

3. INTEGRATION:
   - Seamless integration with the APx API to offer consistent and reliable results.
   - Re-run specific sequences or measurements directly from the software.

4. EXTENSIBILITY:
   - Open-source and modular design to ensure ease of customization and extension for future requirements.

REQUIREMENTS:
-------------
1. Operating System: Windows 10 or higher.
2. Python 3.8 or higher.
3. Dependencies: tkinter, logging, os, and APx API.

SETUP & INSTALLATION:
---------------------
1. Clone the repository from GitHub.
2. Navigate to the cloned directory in your terminal or command prompt.
3. Install any required Python packages.
4. Run the script using Python to launch the GUI.

USAGE:
------
1. SELECTING SEQUENCES:
   - Use the provided ListBoxes to select desired signal paths, measurements, and results.
   - Utilize the 'Pin Selected...' buttons to add selections to the pinned list.

2. EXECUTING SEQUENCES:
   - Click the 'Run Sequence' button to initiate the AP sequence based on your selections.

3. RESULT HANDLING:
   - View results in the GUI.
   - Use the 'Export...' buttons to save results in Excel format.
   
4. RERUNNING SEQUENCES:
   - Use the 'Rerun Pinned Tests' button to execute sequences or measurements from the pinned list.

CONTRIBUTING:
-------------
Feedback, bug reports, and pull requests are welcome on GitHub.

LICENSE:
--------


CREDITS:
--------
Developed by [Joshua Levy at Josh Levy Labs].

CONTACT:
--------
For any queries or support, please contact:
- Email: josh@joshlevylabs.com
- GitHub: https://github.com/joshlevylabs


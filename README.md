# ADAPT PT/RC Reactions Tool - Documentation
This script prints reactions outputted from ADAPT in a DL-SDL-LL format, grouped by support instead of ADAPT's default, by load case.
It allows the user to copy reactions from ADAPT's output quickly, insted of having to manually extract each reaction from the report.

This script is intended for the structural engineer designing concrete beams in ADAPT, to aid in laying out reactions from ADAPT's output. It is assumed that you are working with a Windows PC.

## Disclaimer
**None of the contributors will be liable for any element of design or construction related to the use of this software. It is the engineer's responsibility to determine adequacy of design. This tool is merely intended to aid in simplifying the design process.**

## Installation Instructions

### Recommended Method

Download ***ADAPT Reactions Tool V1.0.exe*** and run it from your PC.

### Running via Python Interpreter

1. Install Python 3.8+ if you don't already. Refer to python.org for more info.
2. Download ***adapt_reactions_tool_xls.py***
3. Use pip to install *xlrd* using the following command in the Python 3 interpreter:

        pip3 install xlrd
    *xlrd* is a Python library used to parse legacy .xls files.

You're good to go!

## Use

### Output a Report from your ADAPT Run
Once your ADAPT run is complete, output a design report for your run with 'Skip Live Load' disabled. Your report must contain the *'Moments, Shears, and Reactions'* tables under *'Tabular Reports - Compact'*. The script will currently not differentiate from a report with skipped live loads enabled vs. disabled. It is important for the engineer to be careful about this.

Ensure that the 'Create Optional XLS Report' option is checked. ***adapt_reactions_tool_xls.py*** uses this file to output its results. Save this .xls file and copy its full path:

1. Hold Shift and right-click the .xls file within Windows Explorer
2. Click **'Copy as Path'**

### Open the Script
Simply run the *adapt_reactions_tool_xls.py* file using Python. You can either double click the file if Python is your default program to run *.py* files, or right click on the file and navigate to *Open With... -> Python*. Alternatively, run the script from your IDE of choice. 

### Output the Reactions
Almost there!

***adapt_reactions_tool_xls.py*** will prompt you for a path to an Excel sheet with reactions. Paste the path we copied before opening the script and hit Enter. The script will print the reactions from ADAPT grouped by support, and ordered by load case: 
- DL
- SDL
- LL

Note: ***adapt_reactions_tool_xls.py*** does not yet support ADAPT's X load case.

## Sample Output

```
Format:
DL
SDL
LL

DL Reaction multiplied by 1
SDL and LL Reactions multiplied by 1
Support 1:
DL: 13.27 k
SD: 34.22 k
LL: 11.80 k 

Support 2:
DL: 13.27 k
SD: 34.22 k
LL: 11.80 k 
```

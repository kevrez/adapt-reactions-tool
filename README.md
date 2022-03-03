# ADAPT PT/RC Reactions Tool - Documentation
This script prints reactions outputted from ADAPT in a DL-SDL-LL format, grouped by support instead of ADAPT's default, by load case.
It allows the user to copy reactions from ADAPT's output quickly, instead of having to manually extract each reaction from the report.

This script is intended for the structural engineer designing concrete beams in ADAPT, to aid in laying out reactions from ADAPT's output. It is assumed that you are working with a Windows PC.

## Disclaimer
**None of the contributors will be liable for any element of design or construction related to the use of this software. It is the engineer's responsibility to determine adequacy of design. This tool is merely intended to aid in simplifying the design process.**

## Installation Instructions

### Recommended Method

Download ***ADAPT Reactions Tool V1.0.exe*** and run it from your PC.

### Running via Python Interpreter

1. Install Python 3.8+ if you don't already. Refer to python.org for more info.
2. Download ***adapt_reactions_tool.py*** and ***adapt_reactions_parser.py*** and save them to the same folder.
3. Use pip to install *xlrd* using the following command in the Python 3 interpreter:

        pip3 install xlrd
    *xlrd* is a Python library used to parse legacy .xls files.

You're good to go!

## Use

### Output a Report from your ADAPT Run
Once your ADAPT run is complete, output a design report for your run with 'Skip Live Load' disabled. Your report must contain the *'Moments, Shears, and Reactions'* tables under *'Tabular Reports - Compact'*.

Ensure that the 'Create Optional XLS Report' option is checked. ***ADAPT Reactions Tool*** uses this file to output its results. Save this .xls file and copy its full path:

1. Hold Shift and right-click the .xls file within Windows Explorer
2. Click **'Copy as Path'**

### Open the Program
Open the executable or run the ***adapt_reactions_tool.py*** file using your Python interpreter. 

### Output the Reactions
Almost there!

***ADAPT Reactions Tool*** will prompt you for a path to an Excel sheet with reactions. Paste the path you copied before opening the script and hit Enter. The script will print the reactions from ADAPT grouped by support, and ordered by load case: 
- DL
- SDL
- LL

Note: ***ADAPT Reactions Tool*** does not yet support ADAPT's X load case.

## Sample Output

```
Support 1:
DL: 11.15 k
SD: 2.86 k
LL: 12.17 k

Support 2:
DL: 11.15 k
SD: 2.86 k
LL: 12.17 k
```

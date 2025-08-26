# BasinVisPy
A python-based toolbox for geospatial calculations in MS Excel and LibreOffice Calc.

This package aims to bring the functionality of BasinVis directly into spreadsheet software.
By removing the requirement for Matlab (TM) and relying entirely on Python our toolbox is now accessible to everyone.

# Installation Guide (Windows)

* Please install or upgrade python to at least version 3.11 or more. Installation instructions for python can be found in the [Python Wiki](https://wiki.python.org/moin/BeginnersGuide/Download).
* Clone this git repository and enter the directory in the file explorer.
* Hold shift and right click into the directory window (do not click on a file while doing this) and select "Open PowerShell window here..."
* Run the command "pip install -r requirements.txt". This should install the dependencies
  * [numpy](https://numpy.org/) - A powerful scientific computing package for python.
  * [sympy](https://www.sympy.org) - A symbolic math solver to evaluate backstripping equations.
  * [xlwings](https://www.xlwings.org/) - An Excel addin to access python functionality within spreadsheets.
* Install the Excel addin by running "xlwings addin install"
* Start Excel and set a reference to xlwings in the VBA editor, instructions for this can be found on the [xlwings add-in page](https://docs.xlwings.org/en/stable/addin.html)
* With this, Excel should be prepared for BasinVis calculations

* To start a project, copy the "BasinVis-example.xlsm" file and the "BasinVis.py" file into a new directory.
* You may rename the spreadsheet file, as long as you keep the xlsm file ending. "BasinVis.py" needs to keep the same filename and stay in the same directory as the spreadsheet for the scripts to work.
* Inside the Spreadsheet there are currently two functionalities available.
  * "decomp( phi_0, c, top_p, bottom_p, top_decomp)" is a new formula, that performs decompaction based on the following parameters:
    (Initial porosity, Coefficient c, Present top depth, Present bottom depth, Decompacted top depth)
  * "subdata (cellrange)" is a helper function that allows you to select in which cells the input well data for backstripping are located.
* The core functionality is provided by xlwings "run main" option. Rather than a simple cell calculation, this will build a interactive spreadsheet calculating tectonic and total subsidence based on your selected "subdata" cells.
  * Depending on the amount of units, this might take a while. The console window provides completion progress bar.

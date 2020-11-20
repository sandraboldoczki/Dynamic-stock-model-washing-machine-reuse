# Hybrid-Dynamic-stock-model-washing-machine-reuse

Guide to the hybrid dynamic stock model for washing machine reuse

Titel: Does increased circularity lead to environmental sustainability? The case of washing machine reuse in Germany (submitted to the Journal of Industrial Ecology)	<br /> 	<br /> 
Sandra Boldoczki*, Andrea Thorenz, Axel Tuma <br /> 
Resource Lab, University of Augsburg, Augsburg, Germany <br /> <br /> 
To whom correspondence should be addressed: *) Sandra Boldoczki, Universitaetsstr. 16, 86159, Augsburg, Germany, sandra.boldoczki@wiwi.uni-augsburg.de


This document provides a guide to the Python code of the hybrid dynamic stock model for washing machine reuse.
The GitHub repository contains three subfolders:

• A folder ‘Data’, where the Excel workbook ‘Input_DSMWM.xlsx’ with all parameters is located

• A folder ‘Results’, which is empty and where the script stores the model results

• A folder ‘Script’, which contains the model script ‘DSM_WM_reuse.py’.

The model script is a standalone script, which apart from standard Python modules does not need further software. To run the hybrid dynamic stock model for a specific parameter constellation, one needs to adapt the respective parameters in the datafile.

The script will then create a subfolder in the results folder with the name structure DateStamp_TimeStamp. In this folder, a copy of the script, and the model results as .xls files are stored.

Note: The hybrid dynamic stock model script can be run and modified for research and teaching purposes. No additional support is provided by the authors.

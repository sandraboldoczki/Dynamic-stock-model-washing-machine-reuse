# Dynamic-stock-model-washing-machine-reuse

Python code of the dynamic stock model for washing machine reuse (submitted)

Journal of Industrial Ecology
Does increased circularity lead to environmental sustainability? The case of washing machine reuse in Germany

Sandra Boldoczki*, Andrea Thorenz, Axel Tuma

Resource Lab, University of Augsburg, Augsburg, Germany

Guide to Supporting Information S19 (of 19)

To whom correspondence should be addressed: *) Sandra Boldoczki, Universitaetsstr. 16, 86159, Augsburg, Germany, sandra.boldoczki@wiwi.uni-augsburg.de

GUIDE TO THE DYNAMIC STOCK MODEL FOR WASHING MACHINE REUSE

This paper comes with a third supplementary file, a zip archive ‘DSM_WM_reuse_SI19.zip’. This archive contains three subfolders:

• A folder ‘Data’, where the Excel workbook ‘Input_DSMWM.xlsx’ with the above described parameters is located

• A folder ‘Results’, which is empty and where the scripts store the model results

• A folder ‘Script’, which contains the model script ‘DSM_WM_reuse.py’.

To run the dynamic stock model, one needs to extract the zip folder and copy its content to a convenient location. Then, in line 80 of the main script, the path of the model folder needs to be specified. The model script is a standalone script, which apart from standard Python modules does not need further software.

To run the dynamic stock model for a specific parameter constellation, one needs to

• Define this constellation in columns H-V of the sheet ‘Parameter_Overview’ of the MaTrace Global datafile. Column H contains the scenario name, column I the scenario description, column J the modus of the model run (at the moment, only ‘TraceSingleProduct’ is supported), column K the start year (e.g., 2015), column J the time horizon (e.g., 2100), column M the test product (at the moment, only ‘Car’ is supported), column N the country index of where the test product is consumed initially (1-25), and columns O-V contain the different improvement options or alternative values for the model parameters that are described above and that can be selected. Row 3 contains comments that describe the possible valid entries for these columns.

• Select this constellation by indicating the index number (column G) of the parameter constellation in cell C4.

• Run the MaTrace Global main script.

The script will then create a subfolder in the results folder with the name structure Results DSM_ScenarioName_DateStamp. In this folder, a copy of the script, and the model results as .xls files are stored.

Note: The dynamic stock model scripts are published under the MIT license and can be run and modified for research and teaching purposes. No additional support is provided by the authors. The main script is also posted on this GitHub repository to track future modification

# Hybrid-dynamic-stock-model-washing-machine-reuse

Python code of the hybrid dynamic stock model for washing machine reuse (submitted)

Journal of Industrial Ecology
Does increased circularity lead to environmental sustainability? The case of washing machine reuse in Germany

Sandra Boldoczki*, Andrea Thorenz, Axel Tuma

Resource Lab, University of Augsburg, Augsburg, Germany

Guide to Supporting Information S19 (of 19)

To whom correspondence should be addressed: *) Sandra Boldoczki, Universitaetsstr. 16, 86159, Augsburg, Germany, sandra.boldoczki@wiwi.uni-augsburg.de

GUIDE TO THE HYBRID DYNAMIC STOCK MODEL FOR WASHING MACHINE REUSE

This paper comes with a third supplementary file, a zip archive ‘DSM_WM_reuse_SI19.zip’. This archive contains three subfolders:

• A folder ‘Data’, where the Excel workbook ‘Input_DSMWM.xlsx’ with the above described parameters is located

• A folder ‘Results’, which is empty and where the scripts store the model results

• A folder ‘Script’, which contains the model script ‘DSM_WM_reuse.py’.

To run the hybrid dynamic stock model, one needs to extract the zip folder and copy its content to a convenient location. Then, in line 34 of the script, the path of the model folder needs to be specified. The model script is a standalone script, which apart from standard Python modules does not need further software.

To run the hybrid dynamic stock model for a specific parameter constellation, one needs to adapt the respective parameters in the datafile.

The script will then create a subfolder in the results folder with the name structure Results DSM_ScenarioName_DateStamp_Timestamp. In this folder, a copy of the script, and the model results as .xls files are stored.

Note: The hybrid dynamic stock model script can be run and modified for research and teaching purposes. No additional support is provided by the authors. The script is also posted on this GitHub repository to track future modification.

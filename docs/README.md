R Addin provides an easy way to define and run R script interactions started from Excel.

# Using RAddin

Running an R script is simply done by selecting the appropriate Rdefinition and pressing "run Rdefinition/shell" on the R Addin Ribbon Tab. 
Selecting the Rdefinition highlights the specified definition area (see below).

# Defining RAddin script interactions (Rdefinitions)

Rscript interactions (Rdefinitions) are defined using a 3 column named area (1st col: definition type, 2nd: definition value, 3rd: definition path):

The Rdefinition area name must start with R_Addin and can have an optional postfix as an additional description of the definition.

An area name can be Workbook global or worksheet local, the first Rdefinition area is being taken as the default definition (used for running when not selecting a Rdefinition).
the worksheet name (for worksheet local names) or the workbook name (for workbook global names) is prepended to the postfix definition description.

So for the 7 areas currently defined in the test workbook testRAddin.xlsx, there should be 7 entries in the Rdefinition dropdown: 

- Test_RdotNet, 
- Test_OtherSheet, 
- testRAddin.xlsx, 
- Test_OtherSheetAnotherDef, 
- testRAddin.xlsxAnotherDef,
- Test_scriptRngScriptCell and
- Test_scriptRngScriptRange 

In the 1st column of the Rdefinition range are the definition types, possible types are 
- rexec: full path to an executable, being able to shell-run the script in line "script" (usually Rscript.exe). When running cmd shell files a special entry "cmd" is used for rexec. Only needed when overriding the ExePath in the AppSettings in the Raddin-AddIn-packed.xll.config file.
- rpath: path to the folder with the R dlls, in case you want to use the in-memory option with RDotNet. Only needed when overriding the RPath in the AppSettings in the Raddin-AddIn-packed.xll.config file. 
- dir: the path where below files (scripts, args, results and diagrams) are stored. 
- script or debug (put debug here to capture & return error output): full path of the executable script. 
- arg (R input objects, txt files): path where the (input) arguments are stored. 
- res (R output objects, txt files): path where the (output) results are expected.
- diag (R output diagrams, png format): path where the (output) diagrams are expected.
- scriptrng (R scripts directly within Excel): either a range, where this script is stored or directly as a value in the 2nd column

In the 2nd column of the Rdefinition range are the definition values as described above.

Absolute Paths in dir or the definition path column are defined by starting with \\ or X:\ (X being a drive letter)
For arg, res, scriptrng and diag these are range names referring to the respective ranges to be taken as arg, res, scriptrng or diag target in the excel workbook.

The range names that are referred in arg, res, scriptrng and diag types can also be either workbook global (having no ! in front of the name) or worksheet local (having the worksheet name + ! in front of the name)

The definitions are loaded into the Rdefinition dropdown either on opening/activating a Workbook with above named areas or by pressing the small dialogBoxLauncher "refresh Rdefinitions" on the R Addin Ribbon Tab.

# Installation of RAddin

put Raddin-AddIn-packed.xll and Raddin-AddIn-packed.xll.config (Raddin-AddIn64-packed.xll and Raddin-AddIn64-packed.xll.config for 64 bit excel) into %appdata%\Microsoft\AddIns 
and run AddIns("Raddin-AddIn-packed.xll").Installed = True in Excel (or add the Addin manually).

# Building RAddin

Building: all packages necessary for building are contained, simply open Raddin.sln and build the solution.
# RAddin

R Addin provides an easy way to define and run R script interactions started from Excel.

Running an R script is simply done by selecting the appropriate Rdefinition and pressing "run this Rdefinition" on the R Addin Ribbon Tab. 
Selecting the Rdefinition highlights the specified definition area (see below).

How to define Rscript interactions (Rdefinitions) using a 3 column named area (1st col: definition type, 2nd: definition value, 3rd: definition path):

Rdefinition area name must start with R_Addin and can have an optional postfix as an additional description of the definition.

An area name can be Workbook global or worksheet local, the first Rdefinition area is being taken as the default definition (used for running when not selecting a Rdefinition).
the worksheet name (for worksheet local names) or the workbook name (for workbook global names) is prepended to the postfix definition description.

So for the 4 areas defined in this test workbook there should be 4 entries in the Rdefinition dropdown: Input_toy, test.xlsm, Input_toyAnotherDef, test.xlsmAnotherDef

1st column: definition types, possible types are rexec, dir, script (put debug here to capture & return error output), arg (R input objects, txt files), res (R output objects, txt files),
scriptrng (R scripts within Excel, either a range, where this script is stored or directly in the 2nd column) and diag (R output diagrams, png format)

rexec is the full path to an executable, being able to run the script in line "script" (usually Rscript.exe). It is only needed when overriding the ExePath in the AppSettings in the Raddin-AddIn-packed.xll.config file.

When running cmd shell files a special entry "cmd" is used for rexec.

2nd column: definition values, for rexec, dir and script (debug) this is simply the path/name of the executable/Rdefinition directory/script.

Absolute Paths in dir or the definition path column are defined by starting with \\ or X:\ (X being a drive letter)
For arg, res, scriptrng and diag these are range names referring to the respective ranges to be taken as arg, res, scriptrng or diag target in the excel workbook.

The range names that are referred in arg, res, scriptrng and diag types can also be either workbook global (having no ! in front of the name) or worksheet local (having the worksheet name + ! in front of the name)

The definitions are loaded into the Rdefinition dropdown either on opening/activating a Workbook with above named areas or by pressing "refresh Rdefinitions" on the R Addin Ribbon Tab.

Installation: put Raddin-AddIn-packed.xll and Raddin-AddIn-packed.xll.config (Raddin-AddIn64-packed.xll and Raddin-AddIn64-packed.xll.config for 64 bit excel) into %appdata%\Microsoft\AddIns 
and run AddIns("Raddin-AddIn-packed.xll").Installed = True in Excel (or add the Addin manually).

Building: all packages necessary for building are contained, simply open Raddin.sln and build the solution.
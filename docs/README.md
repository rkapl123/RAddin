R Addin provides an easy way to define and run R script interactions started from Excel.

# Using RAddin

Running an R script is simply done by selecting the desired invocation method ("run via shell" or "run via RdotNet") on the R Addin Ribbon Tab and clicking "run <Rdefinition>" 
beneath the Sheet-button in the Ribbon group "Run R-Scripts defined in WB/sheets names". Activating the "debug script" toggle button leaves the cmd window open when invocation was done "via shell".
Selecting the Rdefinition in the Rdefinition dropdown highlights the specified definition range.

When running r scripts, following is executed:

## run via shell

the input arguments (arg, see below) are written to files, the scripts defined inside Excel are written and called using the executable located in ExePath/rexec, the defined results/diagrams that were written to file are read and placed in Excel.

## run via RdotNet

the input arguments are passed to a new RdotNet instance, the scripts defined inside Excel or stored on disk are read and called using R Dlls (located in RPath/rpath), the defined results that were created in the R instance are passed to and placed in Excel (no diagrams yet!).


![Image of screenshot1](https://raw.githubusercontent.com/rkapl123/RAddin/master/docs/screenshot1.png)

# Defining RAddin script interactions (Rdefinitions)

Rscript interactions (Rdefinitions) are defined using a 3 column named range (1st col: definition type, 2nd: definition value, 3rd: (optional) definition path):

The Rdefinition range name must start with "R_Addin" and can have a postfix as an additional definition name. 
If there is no postfix after "R_Addin", the script is called "MainScript" in the Workbook/Worksheet (depending whether the range name.

A range name can be Workbook global or worksheet local.
In the Rdefinition dropdowns the worksheet name (for worksheet local names) or the workbook name (for workbook global names) is prepended to the additional postfixed definition name.

So for the 8 definitions (range names) currently defined in the test workbook testRAddin.xlsx, there should be 8 entries in the Rdefinition dropdown: 

- Test_RdotNet, (MainScript)
- Test_OtherSheet, (MainScript)
- testRAddin.xlsx, (MainScript)
- Test_OtherSheetAnotherDef, 
- testRAddin.xlsxAnotherDef,
- testRAddin.xlsxErrorInDef,
- Test_scriptRngScriptCell and
- Test_scriptRngScriptRange 

In the 1st column of the Rdefinition range are the definition types, possible types are 
- rexec: full path to an executable, being able to shell-run the script in line "script" (usually Rscript.exe). When running cmd shell files a special entry "cmd" is used for rexec. Only needed when overriding the ExePath in the AppSettings in the Raddin-AddIn-packed.xll.config file.
- rpath: path to the folder with the R dlls, in case you want to use the in-memory option with RDotNet. Only needed when overriding the RPath in the AppSettings in the Raddin-AddIn-packed.xll.config file. 
- dir: the path where below files (scripts, args, results and diagrams) are stored. 
- script: full path of an executable script. 
- arg/arg[rc] (R input objects, txt files): R variable name and path/filename, where the (input) arguments are stored. For Rdotnet creation of the dataframe, if the definition type ends with "r", "c" or both, the input argument is assumed to contain (c)olumn names, (r)ow names or both.
- res/rres (R output objects, txt files): R variable name and path/filename, where the (output) results are expected. If the definition type is rres, results are removed before saving
- diag (R output diagrams, png format): path/filename, where (output) diagrams are expected.
- scriptrng/scriptcell (R scripts directly within Excel): either ranges, where a script is stored (scriptrng) or directly as a cell value (text content or formula result) in the 2nd column (scriptcell)

Scripts (defined with the script, scriptrng or scriptcell definition types) are executed in sequence of their appearance. Although rexec, rpath and dir definitions can appear more than once, only the last definition is used.

In the 2nd column are the definition values as described above.
- For arg, res, scriptrng and diag these are range names referring to the respective ranges to be taken as arg, res, scriptrng or diag target in the excel workbook.
- The range names that are referred in arg, res, scriptrng and diag types can also be either workbook global (having no ! in front of the name) or worksheet local (having the worksheet name + ! in front of the name)

In the 3rd column are the definition paths of the files referred to in arg, res and diag
- Absolute Paths in dir or the definition path column are defined by starting with \\ or X:\ (X being a mapped drive letter)

The definitions are loaded into the Rdefinition dropdown either on opening/activating a Workbook with above named areas or by pressing the small dialogBoxLauncher "refresh Rdefinitions" on the R Addin Ribbon Tab.

When saving the Workbook the input arguments (definition with arg) defined in the currently selected Rdefinition dropdown are stored as well. If nothing is selected, the first Rdefinition of the dropdown is chosen.

Issues:

- [x] Implement a faster way to read textfiles into excel (done via querytables/textfiles)
- [ ] Implement a faster way to save textfiles from excel
- [ ] Improve RdotNet integration (lists are missing, orientation of vectors, missing diagrams, etc.)

# Installation of RAddin

put Raddin-AddIn-packed.xll and Raddin-AddIn-packed.xll.config (Raddin-AddIn64-packed.xll and Raddin-AddIn64-packed.xll.config for 64 bit excel) into %appdata%\Microsoft\AddIns 
and run AddIns("Raddin-AddIn-packed.xll").Installed = True in Excel (or add the Addin manually).

Adapt the settings in Raddin-AddIn<64>-packed.xll.config:

```XML
  <appSettings file="O:\SOFTWARE\TRIT\R\RAddinSettings.config"> : This is a redirection to a central config file containing the same information below
    <add key="ExePath" value="C:\Program Files\R\R-3.4.0\bin\x64\Rscript.exe" /> : The Executable Path used by the shell invocation method
    <add key="rHome" value="C:\Program Files\R\R-3.4.0" /> : rHome for the RdotNet invocation method, to get the R-DLL-Path the rPath<bitness>bit setting below is used 
    <add key="rPath64bit" value="bin\x64" /> : the folder for the 64 bit R-DLLs 
    <add key="rPath32bit" value="bin\i386" /> : the folder for the 32 bit R-DLLs
    <add key="presetSheetButtonsCount" value="24"/> : the preset maximum Button Count for Sheets (if you expect more sheets with Rdefinitions set it accordingly) 
    <add key="runShell" value="True"/> : the default setting for invocation method shell
    <add key="runRdotNet" value="False"/> : the default setting for invocation method RdotNet 
```

For the RdotNet invocation method always keep in mind that a 32 bit Excel instance can only work with 32 bit R-DLLs and a 64 bit Excel instance can only work with 64 bit R-DLLs !!!

# Building RAddin

Building: all packages necessary for building are contained, simply open Raddin.sln and build the solution.
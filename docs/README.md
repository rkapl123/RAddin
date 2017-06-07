R Addin provides an easy way to define and run R script interactions started from Excel.

# Using RAddin

Running an R script is simply done by selecting the desired invocation method (run via shell or RdotNet) on the R Addin Ribbon Tab and clicking "run <Rdefinition>" 
beneath the Workbook/Sheet name. 
Selecting the Rdefinition in the dropdown highlights the specified definition area.

![Image of screenshot1](https://raw.githubusercontent.com/rkapl123/RAddin/master/docs/screenshot1.png)

# Defining RAddin script interactions (Rdefinitions)

Rscript interactions (Rdefinitions) are defined using a 3 column named area (1st col: definition type, 2nd: definition value, 3rd: definition path):

The Rdefinition area name must start with R_Addin and can have an optional postfix as an additional description of the definition. 
If there is no postfix, the script is called the "MainScript" of this Workbook/Worksheet.

An area name can be Workbook global or worksheet local.
In the Rdefinition dropdown the worksheet name (for worksheet local names) or the workbook name (for workbook global names) is prepended to the postfix definition description.

So for the 8 areas currently defined in the test workbook testRAddin.xlsx, there should be 8 entries in the Rdefinition dropdown: 

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
- script or debug (put debug here to capture & return error output): full path of the executable script. 
- arg (R input objects, txt files): path where the (input) arguments are stored. 
- res (R output objects, txt files): path where the (output) results are expected.
- diag (R output diagrams, png format): path where the (output) diagrams are expected.
- scriptrng or debugrng (R scripts directly within Excel, debugrng is analogous to debug above): either a range, where this script is stored or directly as a value in the 2nd column

In the 2nd column of the Rdefinition range are the definition values as described above.

Absolute Paths in dir or the definition path column are defined by starting with \\ or X:\ (X being a drive letter)
For arg, res, scriptrng and diag these are range names referring to the respective ranges to be taken as arg, res, scriptrng or diag target in the excel workbook.

The range names that are referred in arg, res, scriptrng/debugrng and diag types can also be either workbook global (having no ! in front of the name) or worksheet local (having the worksheet name + ! in front of the name)

The definitions are loaded into the Rdefinition dropdown either on opening/activating a Workbook with above named areas or by pressing the small dialogBoxLauncher "refresh Rdefinitions" on the R Addin Ribbon Tab.

Still open Issues:

- [ ] Implement a faster way to read textfiles into excel (currently this is terribly slow for large files)
- [ ] Improve RdotNet integration (lists are missing, orientation of vectors, etc.)

# Installation of RAddin

put Raddin-AddIn-packed.xll and Raddin-AddIn-packed.xll.config (Raddin-AddIn64-packed.xll and Raddin-AddIn64-packed.xll.config for 64 bit excel) into %appdata%\Microsoft\AddIns 
and run AddIns("Raddin-AddIn-packed.xll").Installed = True in Excel (or add the Addin manually).

Adapt the settings in Raddin-AddIn<64>-packed.xll.config:

```XML
  <appSettings file="O:\SOFTWARE\TRIT\MRO\RAddinSettings.config"> : This is a redirection to a central config file containing the same information below
    <add key="ExePath" value="C:\Program Files\Microsoft\MRO\R-3.3.2\bin\x64\Rscript.exe" /> : The Executable Path used by the shell invocation method
    <add key="rPath" value="C:\Program Files\Microsoft\MRO\R-3.3.2\bin" /> : The R-DLL-Path stub (bitness is added using below settings) for the RdotNet invocation method
    <add key="rPath64bit" value="x64" /> : the folder for the 64 bit R-DLLs 
    <add key="rPath32bit" value="i386" /> : the folder for the 32 bit R-DLLs
	<add key="presetSheetButtonsCount" value="24"/> : the preset maximum Button Count for Sheets (if you expect more sheets with Rdefinitions set it accordingly) 
    <add key="runShell" value="True"/> : the default setting for invocation method shell
    <add key="runRdotNet" value="False"/> : the default setting for invocation method RdotNet 
```

For the RdotNet invocation method alwys keep in mind that a 32 bit Excel instance can only work with 32 bit R-DLLs and a 64 bit Excel instance can only work with 64 bit R-DLLs !!!

# Building RAddin

Building: all packages necessary for building are contained, simply open Raddin.sln and build the solution.
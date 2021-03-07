R Addin provides an easy way to define and run R script interactions started from Excel.

# Using RAddin

Running an R script is simply done by selecting the desired invocation method ("run via shell" or "run via RdotNet") on the R Addin Ribbon Tab and clicking "run <Rdefinition>"
beneath the Sheet-button in the Ribbon group "Run R-Scripts defined in WB/sheets names". Activating the "debug script" toggle button leaves the cmd window open when invocation was done "via shell" and writes additional trace messages to the log (see below). Also, if "run via RdotNet", the output of the R script is also written to the log.
Selecting the Rdefinition in the Rdefinition dropdown highlights the specified definition range.

When running R scripts, following is executed:

## run via shell

The input arguments (arg, see below) are written to files, the scripts defined inside Excel are written and called using the executable located in ExePath/rexec (see settings), the defined results/diagrams that were written to file are read and placed in Excel.

## run via RdotNet

The input arguments are passed to a new RdotNet instance, the scripts defined inside Excel or stored on disk are read and called using R Dlls (located in rHome/rPath<bitness>, see settings), the defined results that were created in the R instance are passed to and placed in Excel (diagrams are passed via file by now!). 
The RdotNet feature is still experimental, not everything will work as expected in a production environment, RdotNet also doesn't work with R versions above 4.0.0.


![Image of screenshot1](https://raw.githubusercontent.com/rkapl123/RAddin/master/docs/screenshot1.png)

# Defining RAddin script interactions (Rdefinitions)

Rscript interactions (Rdefinitions) are defined using a 3 column named range (1st col: definition type, 2nd: definition value, 3rd: (optional) definition path):

The Rdefinition range name must start with "R_Addin" and can have a postfix as an additional definition name.
If there is no postfix after "R_Addin", the script is called "MainScript" in the Workbook/Worksheet.

A range name can be at Workbook level or worksheet level.
In the Rdefinition dropdowns the worksheet name (for worksheet level names) or the workbook name (for workbook level names) is prepended to the additional postfixed definition name.

So for the 10 definitions (range names) currently defined in the test workbook testRAddin.xlsx, there should be 10 entries in the Rdefinition dropdown:

- testRAddin.xlsx, (Workbooklevel name, runs as MainScript)
- testRAddin.xlsxAnotherDef (Workbooklevel name),
- testRAddin.xlsxErrorInDef (Workbooklevel name),
- testRAddin.xlsxNewResDiagDir (Workbooklevel name),
- testRAddin.xlsxOtherPlot (Workbooklevel name),
- Test_OtherSheet, (name in Test_OtherSheet)
- Test_OtherSheetAnotherDef (name in Test_OtherSheet),
- Test_RdotNet, (name in Test_RdotNet)
- Test_scriptRngScriptCell (Test_scriptRng) and
- Test_scriptRngScriptRange (Test_scriptRng)

In the 1st column of the Rdefinition range are the definition types, possible types are
- rexec: an executable, being able to shell-run the script in line "script" (usually Rscript.exe, but can be any executable). This is only needed when overriding the ExePath in the AppSettings in the Raddin-AddIn-packed.xll.config file.
- rpath: path to the folder with the R dlls, in case you want to use the in-memory option with RDotNet. Only needed when overriding the RPath in the AppSettings in the Raddin-AddIn-packed.xll.config file.
- dir: the path where below files (scripts, args, results and diagrams) are stored.
- script: full path of an executable script.
- arg/arg[rc] (R input objects, txt files): R variable name and path/filename, where the (input) arguments are stored. For Rdotnet creation of the dataframe, if the definition type ends with "r", "c" or both, the input argument is assumed to contain (c)olumn names, (r)ow names or both.
- res/rres (R output objects, txt files): R variable name and path/filename, where the (output) results are expected. If the definition type is rres, results are removed from excel before saving and rerunning the R script
- diag (R output diagrams, png format): path/filename, where (output) diagrams are expected.
- scriptrng/scriptcell (R scripts directly within Excel): either ranges, where a script is stored (scriptrng) or directly as a cell value (text content or formula result) in the 2nd column (scriptcell)

Scripts (defined with the script, scriptrng or scriptcell definition types) are executed in sequence of their appearance. Although rexec, rpath and dir definitions can appear more than once, only the last definition is used.

In the 2nd column are the definition values as described above.
- For arg, res, scriptrng and diag these are range names referring to the respective ranges to be taken as arg, res, scriptrng or diag target in the excel workbook.
- The range names that are referred in arg, res, scriptrng and diag types can also be either workbook global (having no ! in front of the name) or worksheet local (having the worksheet name + ! in front of the name)
- for rexec this can either be the full path for the executable, or - in case the executable is already in the path - a simple filename (like cmd.exe or perl.exe)

In the 3rd column are the definition paths of the files referred to in arg, res and diag
- Absolute Paths in dir or the definition path column are defined by starting with \\ or X:\ (X being a mapped drive letter)
- not existing paths for arg, res, scriptrng/scriptcell and diag are created automatically, so dynamical paths can be given here.
- for rexec, additional commandline switches can be passed here to the executable (like "/c" for cmd.exe, this is required to start the subsequent script)

The definitions are loaded into the Rdefinition dropdown either on opening/activating a Workbook with above named areas or by pressing the small dialogBoxLauncher "refresh Rdefinitions" on the R Addin Ribbon Tab and clicking "refresh Rdefinitions":  
![Image of screenshot2](https://raw.githubusercontent.com/rkapl123/RAddin/master/docs/screenshot2.png)

The mentioned hyperlink to the local help file can be configured in the app config file (Raddin.xll.config) with key "LocalHelp".
When saving the Workbook the input arguments (definition with arg) defined in the currently selected Rdefinition dropdown are stored as well. If nothing is selected, the first Rdefinition of the dropdown is chosen.

The error messages are logged to a diagnostic log provided by ExcelDna, which can be accessed by clicking on "show Log". The log level can be set in the `system.diagnostics` section of the app-config file (Raddin.xll.config):
Either you set the switchValue attribute of the source element to prevent any trace messages being generated at all, or you set the initializeData attribute of the added LogDisplay listener to prevent the generated messages to be shown (below the chosen level)  

Issues:

- [x] Implement a faster way to read textfiles into excel (done via querytables/textfiles)
- [ ] Implement a faster way to save textfiles from excel
- [ ] Improve RdotNet integration (lists are missing, orientation of vectors, missing diagrams, etc.)
- [ ] Make RdotNet integration work with R versions above 4.0.0

# Installation of RAddin

run Distribution/deployAddin.cmd (this puts Raddin32.xll/Raddin64.xll as Raddin.xll and Raddin.xll.config into %appdata%\Microsoft\AddIns and starts installRAddinInExcel.vbs (setting AddIns("Raddin.xll").Installed = True in Excel)).

Adapt the settings in Raddin.xll.config:

```XML
  <system.diagnostics>
    <sources>
      <source name="ExcelDna.Integration" switchValue="All">
        <listeners>
          <remove name="System.Diagnostics.DefaultTraceListener" />
          <add name="LogDisplay" type="ExcelDna.Logging.LogDisplayTraceListener,ExcelDna.Integration">
            <!-- EventTypeFilter takes a SourceLevel as the initializeData:
                    Off, Critical, Error, Warning (default), Information, Verbose, All -->
            <filter type="System.Diagnostics.EventTypeFilter" initializeData="Warning" />
          </add>
        </listeners>
      </source>
    </sources>
  </system.diagnostics>
  <appSettings file="your.Central.Configfile.Path"> : This is a redirection to a central config file containing the same information below
    <add key="LocalHelp" value="C:\YourPathToLocalHelp\LocalHelp.htm" /> : If you download this page to your local site, put it there.
    <add key="ExePath" value="C:\Program Files\R\R-4.0.4\bin\x64\Rscript.exe" /> : The Executable Path used by the shell invocation method
    <add key="rHome" value="C:\Program Files\R\R-4.0.4" /> : rHome for the RdotNet invocation method, to get the R-DLL-Path the rPath<bitness>bit setting below is used
    <add key="rPath64bit" value="bin\x64" /> : the folder for the 64 bit R-DLLs
    <add key="rPath32bit" value="bin\i386" /> : the folder for the 32 bit R-DLLs
    <add key="presetSheetButtonsCount" value="24"/> : the preset maximum Button Count for Sheets (if you expect more sheets with Rdefinitions set it accordingly)
    <add key="runShell" value="True"/> : the default setting for invocation method shell
    <add key="runRdotNet" value="False"/> : the default setting for invocation method RdotNet
  </appSettings>
```

For the RdotNet invocation method always keep in mind that a 32 bit Excel instance can only work with 32 bit R-DLLs and a 64 bit Excel instance can only work with 64 bit R-DLLs !!!

# Building

All packages necessary for building are contained, simply open Raddin.sln and build the solution. The script deployForTest.cmd can be used to deploy the built xll and config to %appdata%\Microsoft\AddIns

# RAddin
Simple Excel-DNA based Add-in for handling R-scripts from Excel, storing input objects (scalars/vectors/matrices) and 
retrieving result objects (scalars/vectors/matrices) as text files (currently restricted to tab separated).
Graphics are retrieved from produced png files into Excel to be displayed as diagrams.

Installation: put Raddin-AddIn-packed.xll into %appdata%\Microsoft\AddIns and run AddIns("Raddin-AddIn-packed.xll").Installed = True in Excel (or add the Addin manually).

Documentation in accompanying testDocumentationRAddin.xlsx

Building: Put Excel-DNA (0.33.9) into folder ..\ExcelDNA and build the solution.
Imports Microsoft.Office.Interop.Excel
Imports ExcelDna.Integration
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Configuration
Imports ExcelDna.Integration.CustomUI

Public Module MyFunctions
    Public currWb As Workbook
    Public Rdefinitions As Range
    Public Rcalldefnames As String() = {}
    Public Rcalldefs As Range() = {}
    Public theRibbon As ExcelDna.Integration.CustomUI.IRibbonUI
    Public rexec As String
    Dim dirglobal As String

    ' definitions of current R invocation (scripts, args, results, diags...)
    Dim RdefDic As Dictionary(Of String, String()) = New Dictionary(Of String, String())

    ' prepare Parameters (script, args, results, diags) for usage in invokeScripts, storeArgs, getResults and getDiags 
    Function prepareParams(c As Integer, name As String, ByRef RDataRange As Range, ByRef returnName As String, ByRef returnPath As String, ext As String) As String
        Dim value As String = RdefDic(name)(c)
        ' only for args, results and diags (scripts dont have a target range)
        If name = "args" Or name = "results" Or name = "diags" Then
            Try
                RDataRange = currWb.Names.Item(value).RefersToRange
            Catch ex As Exception
                Return "Error occured when looking up " + name + " range '" + value + "' (defined correctly ?), " + ex.Message
            End Try
        End If
        ' if argvalue refers to a WS Name, cut off WS name prefix for R file name...
        Dim posWSseparator = InStr(value, "!")
        If posWSseparator > 0 Then
            value = value.Substring(posWSseparator)
        End If
        If Len(RdefDic(name + "paths")(c)) > 0 Then
            returnPath = RdefDic(name + "paths")(c)
        End If
        returnName = value + ext
        Return vbNullString
    End Function

    ' invokes current scripts/args/results definition
    Public Function invokeScripts() As String
        Dim script As String = vbNullString
        Dim scriptpath As String

        scriptpath = dirglobal
        For c As Integer = 0 To RdefDic("scripts").Length - 1
            Dim ErrMsg As String = prepareParams(c, "scripts", Nothing, script, scriptpath, "")
            If Len(ErrMsg) > 0 Then Return ErrMsg

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\scriptpath
            Dim curWbPrefix As String = IIf(Left(scriptpath, 2) = "\\" Or Mid(scriptpath, 2, 2) = ":\", "", currWb.Path + "\")
            Dim fullScriptPath = curWbPrefix + scriptpath
            If Not File.Exists(fullScriptPath + "\" + script) Then
                Return "Script '" + fullScriptPath + "\" + script + "' not found!" + vbCrLf
            End If
            If Not File.Exists(rexec) And rexec <> "cmd" Then
                Return "Executable '" + rexec + "' not found!" + vbCrLf
            End If
            Try
                Dim cmd As Process
                cmd = New Process()
                cmd.StartInfo.FileName = IIf(rexec = "cmd", script, rexec)
                cmd.StartInfo.Arguments = IIf(rexec = "cmd", "", script)
                cmd.StartInfo.RedirectStandardInput = False
                cmd.StartInfo.RedirectStandardOutput = IIf(rexec = "cmd", False, RdefDic("debug")(c))
                cmd.StartInfo.RedirectStandardError = IIf(rexec = "cmd", False, RdefDic("debug")(c))
                cmd.StartInfo.CreateNoWindow = False
                cmd.StartInfo.UseShellExecute = (rexec = "cmd")
                cmd.StartInfo.WorkingDirectory = fullScriptPath
                cmd.Start()
                cmd.WaitForExit()
                If RdefDic("debug")(c) And rexec <> "cmd" Then
                    MsgBox("returned error/output from process: " + cmd.StandardError.ReadToEnd())
                End If
            Catch ex As Exception
                Return "Error occured when invoking script '" + script + "' in path '" + currWb.Path + IIf(Len(scriptpath) > 0, "\" + scriptpath, vbNullString) + "', using '" + rexec + "'" + ex.Message + vbCrLf
            End Try
        Next
        Return vbNullString
    End Function

    ' creates Inputfiles for defined arg ranges, tab separated, decimalpoint always ".", dates are stored as "yyyy-MM-dd" 
    ' otherwise:  "what you see is what you get"
    Public Function storeArgs() As String
        Dim argFilename As String = vbNullString, argdir As String
        Dim RDataRange As Range = Nothing
        Dim outputFile As StreamWriter = Nothing

        argdir = dirglobal
        For c As Integer = 0 To RdefDic("args").Length - 1
            Try
                Dim errMsg As String
                errMsg = prepareParams(c, "args", RDataRange, argFilename, argdir, ".txt")
                If Len(errMsg) > 0 Then Return errMsg

                ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\argdir
                Dim curWbPrefix As String = IIf(Left(argdir, 2) = "\\" Or Mid(argdir, 2, 2) = ":\", "", currWb.Path + "\")
                ' remove any existing input files...
                If File.Exists(curWbPrefix + argdir + "\" + argFilename) Then
                    File.Delete(curWbPrefix + argdir + "\" + argFilename)
                End If

                outputFile = New StreamWriter(curWbPrefix + argdir + "\" + argFilename)
                ' make sure we're writing a dot decimal separator
                Dim customCulture As System.Globalization.CultureInfo
                customCulture = System.Threading.Thread.CurrentThread.CurrentCulture.Clone()
                customCulture.NumberFormat.NumberDecimalSeparator = "."
                System.Threading.Thread.CurrentThread.CurrentCulture = customCulture

                ' write the RDataRange to file
                Dim i As Integer = 1
                Do
                    Dim j As Integer = 1
                    Dim writtenLine As String = ""
                    If RDataRange(i, 1).Value2.ToString <> "" Then
                        Do
                            Dim printedValue As String
                            If RDataRange(i, j).NumberFormat.ToString().Contains("yy") Then
                                printedValue = DateTime.FromOADate(RDataRange(i, j).Value2).ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)
                            ElseIf IsNumeric(RDataRange(i, j).Value2) Then
                                printedValue = String.Format("{0:###################0.################}", RDataRange(i, j).Value2)
                            Else
                                printedValue = RDataRange(i, j).Value2
                            End If
                            writtenLine = writtenLine + printedValue + vbTab
                            j = j + 1
                        Loop Until j > RDataRange.Columns.Count
                        outputFile.WriteLine(Left(writtenLine, Len(writtenLine) - 1))
                    End If
                    i = i + 1
                Loop Until i > RDataRange.Rows.Count
            Catch ex As Exception
                If outputFile IsNot Nothing Then outputFile.Close()
                Return "Error occured when creating inputfile '" + argFilename + "', " + ex.Message
            End Try
            If outputFile IsNot Nothing Then outputFile.Close()
        Next
        Return vbNullString
    End Function

    ' get Outputfiles for defined results ranges, tab separated
    ' otherwise:  "what you see is what you get"
    Public Function getResults() As String
        Dim resFilename As String = vbNullString, readdir As String
        Dim RDataRange As Range = Nothing
        Dim errMsg As String = vbNullString

        readdir = dirglobal
        For c As Integer = 0 To RdefDic("results").Length - 1
            errMsg = prepareParams(c, "results", RDataRange, resFilename, readdir, ".txt")
            If Len(errMsg) > 0 Then Return errMsg

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\readdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            Dim infile As StreamReader = Nothing
            Try
                infile = New StreamReader(curWbPrefix + readdir + "\" + resFilename)
            Catch ex As Exception
                Return "Error occured in getResults when opening '" + currWb.Path + "\" + readdir + "\" + resFilename + "', " + ex.Message
            End Try

            ' parse the actual file line by line
            Dim i As Integer = 1, currentRecord As String(), currentLine As String
            RDataRange.ClearContents()
            Do While Not infile.EndOfStream
                Try
                    currentLine = infile.ReadLine
                    currentRecord = currentLine.Split(vbTab)
                Catch ex As FileIO.MalformedLineException
                    If infile IsNot Nothing Then infile.Close()
                    Return "Error occured in getResults when parsing file '" + resFilename + "', " + ex.Message
                End Try
                ' Put parsed data into target range column by column
                For j = 1 To currentRecord.Count()
                    Try
                        RDataRange.Cells(i, j).Value2 = currentRecord(j - 1)
                    Catch ex As Exception
                        If infile IsNot Nothing Then infile.Close()
                        Return "Error occured in getResults when writing data into '" + RDataRange.Parent.name + "!" + RDataRange.Address + "', " + ex.Message
                    End Try
                Next
                i = i + 1
            Loop
            If infile IsNot Nothing Then infile.Close()
        Next
        Return vbNullString
    End Function

    ' get Output diagrams (png) for defined diags ranges
    Public Function getDiags() As String
        Dim diagFilename As String = vbNullString, readdir As String
        Dim RDataRange As Range = Nothing
        Dim errMsg As String = vbNullString

        readdir = dirglobal
        For c As Integer = 0 To RdefDic("diags").Length - 1
            errMsg = prepareParams(c, "diags", RDataRange, diagFilename, readdir, ".png")
            If Len(errMsg) > 0 Then Return errMsg

            ' clean previously set shapes...
            For Each oldShape As Shape In RDataRange.Worksheet.Shapes
                If oldShape.Name = diagFilename Then
                    oldShape.Delete()
                    Exit For
                End If
            Next
            ' add new shape from picture
            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\readdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            Try
                With RDataRange.Worksheet.Shapes.AddPicture(Filename:=curWbPrefix + readdir + "\" + diagFilename, _
                    LinkToFile:=False, SaveWithDocument:=True, Left:=RDataRange.Left, Top:=RDataRange.Top, Width:=-1, Height:=-1)
                    .Name = diagFilename
                End With
            Catch ex As Exception
                Return "Error occured when placing the diagram into target range '" + RdefDic("diags")(c) + "', " + ex.Message
            End Try
        Next
        Return vbNullString
    End Function

    ' remove result and diagram files from arg(ument) ranges
    Public Function removeResultsAndDiags() As String
        Dim resFilename As String = vbNullString, diagFilename As String = vbNullString, readdir As String
        Dim RDataRange As Range = Nothing
        Dim errMsg As String = vbNullString

        readdir = dirglobal
        For c As Integer = 0 To RdefDic("results").Length - 1
            errMsg = prepareParams(c, "results", RDataRange, resFilename, readdir, ".txt")
            If Len(errMsg) > 0 Then Return errMsg

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\argdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            ' remove any existing result files...
            If File.Exists(curWbPrefix + readdir + "\" + resFilename) Then
                Try
                    File.Delete(curWbPrefix + readdir + "\" + resFilename)
                Catch ex As Exception
                    Return "Error occured when trying to remove '" + curWbPrefix + readdir + "\" + resFilename + "', " + ex.Message
                End Try
            End If
        Next
        For c As Integer = 0 To RdefDic("diags").Length - 1
            errMsg = prepareParams(c, "diags", RDataRange, diagFilename, readdir, ".png")
            If Len(errMsg) > 0 Then Return errMsg

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\argdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            ' remove any existing diagram files...
            If File.Exists(curWbPrefix + readdir + "\" + diagFilename) Then
                Try
                    File.Delete(curWbPrefix + readdir + "\" + diagFilename)
                Catch ex As Exception
                    Return "Error occured when trying to remove '" + curWbPrefix + readdir + "\" + diagFilename + "', " + ex.Message
                End Try
            End If
        Next
        Return vbNullString
    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' startRprocess: started from GUI (button) and accessible from VBA (via Application.Run)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function startRprocess() As String
        Dim errStr As String
        ' get the definition range
        errStr = getRDefinitions()
        If errStr <> vbNullString Then Return "Failed getting Rdefinitions: " + errStr
        ' remove result and diagram files from arg(ument) ranges
        errStr = removeResultsAndDiags()
        If errStr <> vbNullString Then Return "removing  result and diagram files returned error: " + errStr
        ' store input files from arg(ument) ranges
        errStr = storeArgs()
        If errStr <> vbNullString Then Return "storing input files returned error: " + errStr
        ' invoke r script(s)
        errStr = invokeScripts()
        If errStr <> vbNullString Then Return "invoking scripts returned error: " + errStr
        ' get and write output files into res(ult) ranges
        errStr = getResults()
        If errStr <> vbNullString Then Return "fetching/placing result files/content returned error: " + errStr
        ' get and put result diagrams/pictures into dia(gram) ranges
        errStr = getDiags()
        If errStr <> vbNullString Then Return "fetching/placing result diagrams returned error: " + errStr
        ' all ís OK = return nullstring
        Return vbNullString
    End Function

    Public Function startRdefRefresh() As String
        Dim errStr As String
        ' always reset Rdefinitions when changing Workbooks, otherwise this is not being refilled in getRNames
        Rdefinitions = Nothing
        ' get the defined R_Addin Names
        errStr = getRNames()
        If errStr = "no definitions" Then
            Return vbNullString
        ElseIf errStr <> vbNullString Then
            Return "Error while getRNames in startRdefRefresh: " + errStr
        End If
        MyFunctions.theRibbon.Invalidate()
        Return vbNullString
    End Function

    ' gets defined named ranges for R script invocation in the current workbook 
    Function getRNames() As String
        ' get default rexec path from user (or overriden in appSettings tag as redirect to global) settings. This can be overruled by individual rexec settings in Rdefinitions
        Try
            rexec = ConfigurationManager.AppSettings("ExePath").ToString()
        Catch ex As Exception
            Return "no ExePath when retrieving AppSettings in getRNames: " + ex.Message
        End Try

        ReDim Preserve Rcalldefnames(-1)
        ReDim Preserve Rcalldefs(-1)
        For Each namedrange As Name In currWb.Names
            Dim cleanname As String = Replace(namedrange.Name, namedrange.Parent.Name & "!", "")
            If Left(cleanname, 7) = "R_Addin" Then
                If namedrange.RefersToRange.Columns.Count <> 3 Then Return "Rdefinitions range " + namedrange.Parent.name + "!" + namedrange.Name + " doesn't have 3 columns !"
                ' final name of entry is without R_Addin and !
                Dim finalname As String = Replace(Replace(namedrange.Name, "R_Addin", ""), "!", "")
                ' first definition as standard definition (works without selecting a Rdefinition)
                If Rdefinitions Is Nothing Then Rdefinitions = namedrange.RefersToRange
                If Not InStr(namedrange.Name, "!") > 0 Then
                    finalname = currWb.Name + finalname
                End If
                ReDim Preserve Rcalldefnames(Rcalldefnames.Length)
                ReDim Preserve Rcalldefs(Rcalldefs.Length)
                Rcalldefnames(Rcalldefnames.Length - 1) = finalname
                Rcalldefs(Rcalldefs.Length - 1) = namedrange.RefersToRange
            End If
        Next
        If UBound(Rcalldefnames) = -1 Then Return "no RNames"
        Return vbNullString
    End Function

    ' gets definitions from  current selected R script invocation range (Rdefintions)
    Function getRDefinitions() As String
        Try
            RdefDic("args") = {}
            RdefDic("argspaths") = {}
            RdefDic("results") = {}
            RdefDic("resultspaths") = {}
            RdefDic("diags") = {}
            RdefDic("diagspaths") = {}
            RdefDic("scripts") = {}
            RdefDic("scriptspaths") = {}
            RdefDic("debug") = {}
            For Each defRow As Range In Rdefinitions.Rows
                Dim deftype As String, defval As String, deffilepath As String
                deftype = LCase(defRow.Cells(1, 1).Value2)
                defval = defRow.Cells(1, 2).Value2
                deffilepath = defRow.Cells(1, 3).Value2
                If deftype = "rexec" Then
                    rexec = defval
                ElseIf deftype = "arg" Then
                    ReDim Preserve RdefDic("args")(RdefDic("args").Length)
                    RdefDic("args")(RdefDic("args").Length - 1) = defval
                    ReDim Preserve RdefDic("argspaths")(RdefDic("argspaths").Length)
                    RdefDic("argspaths")(RdefDic("argspaths").Length - 1) = deffilepath
                ElseIf deftype = "res" Then
                    ReDim Preserve RdefDic("results")(RdefDic("results").Length)
                    RdefDic("results")(RdefDic("results").Length - 1) = defval
                    ReDim Preserve RdefDic("resultspaths")(RdefDic("resultspaths").Length)
                    RdefDic("resultspaths")(RdefDic("resultspaths").Length - 1) = deffilepath
                ElseIf deftype = "diag" Then
                    ReDim Preserve RdefDic("diags")(RdefDic("diags").Length)
                    RdefDic("diags")(RdefDic("diags").Length - 1) = defval
                    ReDim Preserve RdefDic("diagspaths")(RdefDic("diagspaths").Length)
                    RdefDic("diagspaths")(RdefDic("diagspaths").Length - 1) = deffilepath
                ElseIf deftype = "script" Or deftype = "debug" Then
                    ReDim Preserve RdefDic("debug")(RdefDic("debug").Length)
                    RdefDic("debug")(RdefDic("debug").Length - 1) = (deftype = "debug")
                    ReDim Preserve RdefDic("scripts")(RdefDic("scripts").Length)
                    RdefDic("scripts")(RdefDic("scripts").Length - 1) = defval
                    ReDim Preserve RdefDic("scriptspaths")(RdefDic("scriptspaths").Length)
                    RdefDic("scriptspaths")(RdefDic("scriptspaths").Length - 1) = deffilepath
                ElseIf deftype = "dir" Then
                    dirglobal = defval
                End If
            Next
            If rexec = "" Then Return "Error in getRDefinitions: no rexec defined"
            If RdefDic("scripts").Count = 0 Then Return "Error in getRDefinitions: no script(s) defined"
        Catch ex As Exception
            Return "Error in getRDefinitions: " + ex.Message
        End Try
        Return vbNullString
    End Function
End Module

' Events from Excel (Workbook_Save ...)
Public Class AddIn
    Implements IExcelAddIn

    WithEvents Application As Application
    WithEvents Button As Microsoft.Office.Core.CommandBarButton

    ' connect to Excel when opening Addin
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        Application = ExcelDnaUtil.Application
    End Sub

    'has to be implemented
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
    End Sub

    Private Sub Workbook_Save(Wb As Workbook, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
        Dim errStr As String
        errStr = doDefinitions(Wb)
        If errStr = "no RNames" Then Exit Sub
        If errStr <> vbNullString Then
            MsgBox("Error when getting definitions in Workbook_Save: " + errStr)
            Exit Sub
        End If
        errStr = storeArgs()
        If errStr <> "" Then MsgBox("Error when saving inputfiles in Workbook_Save: " + errStr)
    End Sub

    Private Sub Workbook_Open(Wb As Workbook) Handles Application.WorkbookOpen
        ' is being treated in Workbook_Activate...
    End Sub

    Private Sub Workbook_Activate(Wb As Workbook) Handles Application.WorkbookActivate
        Dim errStr As String
        errStr = doDefinitions(Wb)
        If errStr = "no RNames" Then Exit Sub
        If errStr <> vbNullString Then
            MsgBox("Error when getting definitions in Workbook_Activate: " + errStr)
            Exit Sub
        End If
        MyFunctions.theRibbon.Invalidate()
    End Sub

    Private Function doDefinitions(Wb As Workbook) As String
        Dim errStr As String
        currWb = Wb
        ' always reset Rdefinitions when changing Workbooks (may not be the current one, if saved programmatically!), otherwise this is not being refilled in getRNames
        Rdefinitions = Nothing
        ' get the defined R_Addin Names
        errStr = getRNames()
        If errStr = "no RNames" Then Return "no RNames"
        If errStr <> vbNullString Then
            Return "Error while getRNames in doDefinitions: " + errStr
        End If
        ' get the definitions from the current defined range (first name in R_Addin Names)
        errStr = getRDefinitions()
        If errStr <> vbNullString Then Return "Error while getRdefinitions in doDefinitions: " + errStr
        Return vbNullString
    End Function
End Class

' Events from Ribbon
<ComVisible(True)>
Public Class MyRibbon
    Inherits ExcelRibbon

    Public Sub startRprocess(control As ExcelDna.Integration.CustomUI.IRibbonControl)
        Dim errStr As String
        If UBound(Rcalldefnames) = -1 Then
            MsgBox("no Rdefinitions found for R_Addin in current Workbook (3 column named range (type/value/path), minimum types: rexec and script)!")
            Exit Sub
        End If
        If MyFunctions.Rdefinitions Is Nothing Then
            MsgBox("Rdefinitions Is Nothing (this shouldn't actually happen) !")
            Exit Sub
        End If
        errStr = MyFunctions.startRprocess()
        If errStr <> "" Then MsgBox(errStr)
    End Sub


    Public Sub refreshRdefs(control As ExcelDna.Integration.CustomUI.IRibbonControl)
        Dim errStr As String
        errStr = MyFunctions.startRdefRefresh()
        If errStr <> vbNullString Then
            MsgBox(errStr)
        Else
            MsgBox("refreshed Rdefinitions from current Workbook !")
        End If
    End Sub

    Public Function GetItemCount(control As ExcelDna.Integration.CustomUI.IRibbonControl) As Integer
        Return (MyFunctions.Rcalldefnames.Length)
    End Function

    Public Function GetItemLabel(control As ExcelDna.Integration.CustomUI.IRibbonControl, index As Integer) As String
        Return MyFunctions.Rcalldefnames(index)
    End Function

    Public Function GetItemID(control As ExcelDna.Integration.CustomUI.IRibbonControl, index As Integer) As String
        Return MyFunctions.Rcalldefnames(index)
    End Function

    Public Sub selectItem(control As ExcelDna.Integration.CustomUI.IRibbonControl, id As String, index As Integer)
        MyFunctions.Rdefinitions = Rcalldefs(index)
        MyFunctions.Rdefinitions.Parent.Select()
        MyFunctions.Rdefinitions.Select()
    End Sub
    Public Sub ribbonLoaded(myribbon As ExcelDna.Integration.CustomUI.IRibbonUI)
        MyFunctions.theRibbon = myribbon
    End Sub

End Class
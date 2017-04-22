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
                Return "Error occured when looking up " + name + " range '" + value + "', " + ex.Message
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
        Dim errMsg As String = vbNullString

        scriptpath = dirglobal
        For c As Integer = 0 To RdefDic("scripts").Length - 1
            errMsg = prepareParams(c, "scripts", Nothing, script, scriptpath, "")
            If Len(errMsg) > 0 Then Exit For

            Dim fullScriptPath = currWb.Path + IIf(Len(scriptpath) > 0, "\" + scriptpath, vbNullString)
            If Not File.Exists(fullScriptPath + "\" + script) Then
                Return "Script '" + fullScriptPath + "\" + script + "' not found!"
            End If
            If Not File.Exists(rexec) Then
                Return "Executable '" + rexec + "' not found!"
            End If
            Try
                Dim cmd As Process
                cmd = New Process()
                cmd.StartInfo.FileName = rexec
                cmd.StartInfo.Arguments = script
                cmd.StartInfo.RedirectStandardInput = False
                cmd.StartInfo.RedirectStandardOutput = False
                cmd.StartInfo.CreateNoWindow = False
                cmd.StartInfo.UseShellExecute = False
                cmd.StartInfo.WorkingDirectory = fullScriptPath
                'cmd.Start()
                'cmd.WaitForExit()
            Catch ex As Exception
                Return "Error occured when invoking script '" + script + "' in path '" + currWb.Path + IIf(Len(scriptpath) > 0, "\" + scriptpath, vbNullString) + "', using '" + rexec + "'" + ex.Message
            End Try
        Next
        Return errMsg
    End Function

    ' creates Inputfiles for defined arg ranges, tab separated, decimalpoint always ".", dates are stored as "yyyy-MM-dd"
    ' otherwise:  "what you see is what you get"
    Public Function storeArgs() As String
        Dim argFilename As String = vbNullString, argdir As String
        Dim RDataRange As Range = Nothing
        Dim outputFile As StreamWriter = Nothing
        Dim errMsg As String = vbNullString

        argdir = dirglobal
        For c As Integer = 0 To RdefDic("args").Length - 1
            Try
                errMsg = prepareParams(c, "args", RDataRange, argFilename, argdir, ".txt")
                If Len(errMsg) > 0 Then Exit For

                outputFile = New StreamWriter(currWb.Path + "\" + argdir + "\" + argFilename)
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
                                printedValue = String.Format("{0:##################.########}", RDataRange(i, j).Value2)
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
                errMsg = "Error occured when creating inputfile '" + argFilename + "', " + ex.Message
            Finally
                If outputFile IsNot Nothing Then outputFile.Close()
            End Try
        Next
        Return errMsg
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
            If Len(errMsg) > 0 Then Exit For

            Dim afile As StreamReader = Nothing
            Try
                afile = New StreamReader(currWb.Path + "\" + readdir + "\" + resFilename)
            Catch ex As Exception
                Return "Error occured when opening '" + currWb.Path + "\" + readdir + "\" + resFilename + "', " + ex.Message
            End Try

            ' parse the actual file line by line
            Dim i As Integer = 1, currentRecord As String(), currentLine As String
            Do While Not afile.EndOfStream
                Try
                    currentLine = afile.ReadLine
                    currentRecord = currentLine.Split(vbTab)
                Catch ex As FileIO.MalformedLineException
                    afile.Close()
                    Return "Error occured when parsing file '" + resFilename + "', " + ex.Message
                End Try
                RDataRange.Clear()
                ' Put parsed data into target range column by column
                For j = 1 To currentRecord.Count()
                    RDataRange.Cells(i, j).Value2 = currentRecord(j - 1)
                Next
                i = i + 1
            Loop
        Next
        Return errMsg
    End Function

    ' get Output diagrams (png) for defined diags ranges
    Public Function getDiags() As String
        Dim diagFilename As String = vbNullString, readdir As String
        Dim RDataRange As Range = Nothing
        Dim errMsg As String = vbNullString

        readdir = dirglobal
        For c As Integer = 0 To RdefDic("diags").Length - 1
            errMsg = prepareParams(c, "diags", RDataRange, diagFilename, readdir, ".png")
            If Len(errMsg) > 0 Then Exit For
            ' clean previously set shapes...
            For Each oldShape As Shape In RDataRange.Worksheet.Shapes
                If oldShape.TopLeftCell.Address = RDataRange.Address Then
                    oldShape.Delete()
                    Exit For
                End If
            Next
            ' add new shape from picture
            Try
                RDataRange.Worksheet.Shapes.AddPicture( _
                    Filename:=currWb.Path + "\" + readdir + "\" + diagFilename, _
                    LinkToFile:=False, SaveWithDocument:=True, Left:=RDataRange.Left, Top:=RDataRange.Top, Width:=-1, Height:=-1)
            Catch ex As Exception
                Return "Error occured when placing the diagram into target range '" + RdefDic("diags")(c) + "', " + ex.Message
            End Try
        Next
            Return errMsg
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' startRprocess: started from GUI (button) and accessible from VBA (via Application.Run)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function startRprocess() As String
        Dim errStr As String
        ' get the definition range
        errStr = getRDefinitions()
        If errStr <> vbNullString Then Return "Failed getting Rdefinitions: " + errStr
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

    ' gets defined named ranges for R script invocation in the current workbook
    Function getRNames() As String
        ReDim Preserve Rcalldefnames(-1)
        ReDim Preserve Rcalldefs(-1)
        For Each namedrange As Name In currWb.Names
            Dim cleanname As String = Replace(namedrange.Name, namedrange.Parent.Name & "!", "")
            If Left(cleanname, 7) = "R_Addin" Then
                ' final name of entry is without R_Addin and !
                Dim finalname As String = Replace(Replace(namedrange.Name, "R_Addin", ""), "!", "")
                ' first workbook level definition as standard definition
                If Not InStr(namedrange.Name, "!") > 0 Then
                    finalname = currWb.Name + finalname
                    If Rdefinitions Is Nothing Then Rdefinitions = namedrange.RefersToRange
                End If
                ReDim Preserve Rcalldefnames(Rcalldefnames.Length)
                ReDim Preserve Rcalldefs(Rcalldefs.Length)
                Rcalldefnames(Rcalldefnames.Length - 1) = finalname
                Rcalldefs(Rcalldefs.Length - 1) = namedrange.RefersToRange
            End If
        Next
        If UBound(Rcalldefnames) = -1 Then Return "no definitions"
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
                ElseIf deftype = "script" Then
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
        Try
            MyFunctions.rexec = ConfigurationManager.AppSettings("RscriptPath").ToString()
        Catch ex As Exception
            MsgBox("Error when retrieving settings: " + ex.Message)
        End Try
    End Sub

    'has to be implemented
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
    End Sub

    Private Sub Workbook_Save(Wb As Workbook, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
        If UBound(Rcalldefnames) = -1 Or MyFunctions.Rdefinitions Is Nothing Then Exit Sub
        currWb = Wb
        ' get the definition range
        Dim errStr As String
        ' get the definition range
        errStr = getRDefinitions()
        If errStr <> vbNullString Then MsgBox("Error while getting Rdefinitions: " + errStr)

        errStr = storeArgs()
        If errStr <> "" Then MsgBox("Error when saving inputfiles: " + errStr)
    End Sub

    Private Sub Workbook_Open(Wb As Workbook) Handles Application.WorkbookOpen
        currWb = Wb
        ' get the definition range
        Dim errStr As String
        errStr = getRNames()
        If errStr = "no definitions" Then Exit Sub
        ' get the definition range
        errStr = getRDefinitions()
        If errStr <> vbNullString Then MsgBox("Error while getting Rdefinitions: " + errStr)
        MyFunctions.theRibbon.Invalidate()
    End Sub

    Private Sub Workbook_Activate(Wb As Workbook) Handles Application.WorkbookActivate
        currWb = Wb
        ' get the definition range
        Dim errStr As String
        errStr = getRNames()
        If errStr = "no definitions" Then Exit Sub
        ' get the definition range
        errStr = getRDefinitions()
        If errStr <> vbNullString Then MsgBox("Error while getting Rdefinitions: " + errStr)
        MyFunctions.theRibbon.Invalidate()
    End Sub
End Class

' Events from Ribbon
<ComVisible(True)>
Public Class MyRibbon
    Inherits ExcelRibbon

    Public Sub startRprocess(control As ExcelDna.Integration.CustomUI.IRibbonControl)
        Dim errStr As String
        If UBound(Rcalldefnames) = -1 Then
            MsgBox("no Rdefinitions found for R_Addin (3 column named range (type/value/path), minimum types: rexec and script)!")
            Exit Sub
        End If
        If MyFunctions.Rdefinitions Is Nothing Then
            MsgBox("no Rdefinition selected for starting R script!")
            Exit Sub
        End If
        errStr = MyFunctions.startRprocess()
        If errStr <> "" Then MsgBox(errStr)
    End Sub

    Public Function GetItemCount(control As ExcelDna.Integration.CustomUI.IRibbonControl) As Integer
        Return(MyFunctions.Rcalldefnames.Length)
    End Function

    Public Function GetItemLabel(control As ExcelDna.Integration.CustomUI.IRibbonControl, index As Integer) As String
        Return MyFunctions.Rcalldefnames(index)
    End Function

    Public Function GetItemID(control As ExcelDna.Integration.CustomUI.IRibbonControl, index As Integer) As String
        Return MyFunctions.Rcalldefnames(index)
    End Function

    Public Sub selectItem(control As ExcelDna.Integration.CustomUI.IRibbonControl, id As String, index As Integer)
        MyFunctions.Rdefinitions = Rcalldefs(index)
    End Sub
    Public Sub ribbonLoaded(myribbon As ExcelDna.Integration.CustomUI.IRibbonUI)
        MyFunctions.theRibbon = myribbon
    End Sub

End Class
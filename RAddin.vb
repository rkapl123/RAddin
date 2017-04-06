Imports Microsoft.Office.Interop.Excel
Imports ExcelDna.Integration
Imports System.IO
Imports System.Runtime.InteropServices
Imports ExcelDna.Integration.CustomUI

Public Module MyFunctions
    Public currWb As Workbook
    Public Rdefinitions As Range
    Public Rcalldefnames As String() = {}
    Public Rcalldefs As Range() = {}
    Public theRibbon As ExcelDna.Integration.CustomUI.IRibbonUI

    Dim scripts As String() = {}
    Dim scriptspaths As String() = {}
    Dim rexec As String
    Dim dirglobal As String
    Dim args As String() = {}
    Dim argspaths As String() = {}
    Dim results As String() = {}
    Dim resultspaths As String() = {}
    Dim diags As String() = {}
    Dim diagspaths As String() = {}

    ' creates Inputfiles for defined arg ranges, tab separated, decimalpoint always ".", dates are stored as "yyyy-MM-dd"
    ' otherwise:  "what you see is what you get"
    Public Function storeInput() As String
        Dim argFilename As String = vbNullString, writedir As String, RDataRange As Range
        Dim outputFile As StreamWriter = Nothing
        Dim errMsg As String = vbNullString

        writedir = dirglobal
        For c As Integer = 0 To args.Length - 1
            Try
                Dim argvalue As String
                argvalue = args(c)
                If Len(argspaths(c)) > 0 Then writedir = argspaths(c)
                argFilename = argvalue + ".txt"
                Try
                    RDataRange = currWb.Names.Item(argvalue).RefersToRange
                Catch ex As Exception
                    Return "Error occured when looking up arg range '" + argvalue + "', " + ex.Message
                End Try
                outputFile = New StreamWriter(currWb.Path + "\" + writedir + "\" + argFilename)
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
                If Not IsNothing(outputFile) Then outputFile.Close()
            End Try
        Next
        Return errMsg
    End Function

    Public Function invokeScripts() As String
        Dim script As String
        Dim scriptpath As String

        scriptpath = dirglobal
        For c As Integer = 0 To scripts.Length - 1
            script = scripts(c)
            If Len(scriptspaths(c)) > 0 Then scriptpath = scriptspaths(c)
            Try
                Dim cmd As Process
                cmd = New Process()
                cmd.StartInfo.FileName = rexec
                cmd.StartInfo.Arguments = script
                cmd.StartInfo.RedirectStandardInput = False
                cmd.StartInfo.RedirectStandardOutput = False
                cmd.StartInfo.CreateNoWindow = False
                cmd.StartInfo.UseShellExecute = False
                cmd.StartInfo.WorkingDirectory = currWb.Path + IIf(Len(scriptpath) > 0, "\" + scriptpath, vbNullString)
                cmd.Start()
                cmd.WaitForExit()
            Catch ex As Exception
                Return "Error occured when invoking script '" + script + "' in path '" + currWb.Path + IIf(Len(scriptpath) > 0, "\" + scriptpath, vbNullString) + "', using '" + rexec + "'" + ex.Message
            End Try
        Next
        Return vbNullString
    End Function

    Public Function getOutput() As String
        Dim resFilename As String, readdir As String, RDataRange As Range
        readdir = dirglobal
        For c As Integer = 0 To results.Length - 1
            Dim resvalue As String
            resvalue = results(c)
            If resultspaths(c) <> vbNullString Then readdir = resultspaths(c)
            resFilename = resvalue + ".txt"
            Try
                RDataRange = currWb.Names.Item(resvalue).RefersToRange
            Catch ex As Exception
                Return "Error occured when looking up result range '" + resvalue + "', " + ex.Message
            End Try

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
                ' Put parsed data into target range column by column
                For j = 1 To currentRecord.Count()
                    RDataRange.Cells(i, j).Value2 = currentRecord(j - 1)
                Next
                i = i + 1
            Loop
        Next
        Return vbNullString
    End Function

    Public Function getOutDiagrams() As String
        Dim diagFilename As String, readdir As String, RDataRange As Range

        readdir = dirglobal

        For c As Integer = 0 To diags.Length - 1
            Dim diagvalue As String
            diagvalue = diags(c)
            If diagspaths(c) <> vbNullString Then readdir = diagspaths(c)
            diagFilename = diagvalue + ".png"
            Try
                RDataRange = currWb.Names.Item(diagvalue).RefersToRange
            Catch ex As Exception
                Return "Error occured when looking up diagram target range '" + diagvalue + "', " + ex.Message
            End Try

            Try
                With RDataRange.Parent.Pictures.Insert(Filename:=currWb.Path + "\" + readdir + "\" + diagFilename, LinkToFile:=False, SaveWithDocument:=True)
                    .Left = RDataRange.Left
                    .Top = RDataRange.Top
                    .Placement = 1
                    .PrintObject = True
                End With
            Catch ex As Exception
                Return "Error occured when placing the diagram into target range '" + diagvalue + "', " + ex.Message
            End Try
        Next
        Return vbNullString
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' accessible from VBA (via Application.Run)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function startRprocess() As String
        Dim errStr As String
        ' get the definition range
        errStr = getRDefinitions()
        If errStr <> vbNullString Then
            Return "Failed getting Rdefinitions: " + errStr
        End If

        ' store input files from arg(ument) ranges
        errStr = storeInput()
        If errStr <> vbNullString Then
            Return "storing input files returned error: " + errStr
        End If

        ' invoke r script(s)
        errStr = invokeScripts()
        If errStr <> vbNullString Then
            Return "invoking scripts returned error: " + errStr
        End If

        ' get and write output files into res(ult) ranges
        errStr = getOutput()
        If errStr <> vbNullString Then
            Return "fetching/placing result files/content returned error: " + errStr
        End If

        ' get and put result diagrams/pictures into dia(gram) ranges
        errStr = getOutDiagrams()
        If errStr <> vbNullString Then
            Return "fetching/placing result diagrams returned error: " + errStr
        End If

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
                Dim finalname As String = Replace(cleanname, "R_Addin", "")
                ' first workbook level definition as standard definition
                If IsNothing(Rdefinitions) And Not InStr(namedrange.Name, "!") > 0 Then Rdefinitions = namedrange.RefersToRange
                If Len(finalname) = 0 Then
                    finalname = namedrange.Parent.Name
                End If
                ReDim Preserve Rcalldefnames(Rcalldefnames.Length)
                ReDim Preserve Rcalldefs(Rcalldefs.Length)
                Rcalldefnames(Rcalldefnames.Length - 1) = finalname
                Rcalldefs(Rcalldefs.Length - 1) = namedrange.RefersToRange
            End If
        Next
        If UBound(Rcalldefnames) = 0 Then Return "no definitions"
        Return vbNullString
    End Function

    ' gets definitions from  R script invocation ranges in the current workbook
    Function getRDefinitions() As String
        Try
            ReDim Preserve args(-1)
            ReDim Preserve argspaths(-1)
            ReDim Preserve results(-1)
            ReDim Preserve resultspaths(-1)
            ReDim Preserve diags(-1)
            ReDim Preserve diagspaths(-1)
            ReDim Preserve scripts(-1)
            ReDim Preserve scriptspaths(-1)
            For Each defRow As Range In Rdefinitions.Rows
                Dim deftype As String, defval As String, deffilepath As String
                deftype = LCase(defRow.Cells(1, 1).Value2)
                defval = defRow.Cells(1, 2).Value2
                deffilepath = defRow.Cells(1, 3).Value2
                If deftype = "rexec" Then
                    rexec = defval
                ElseIf deftype = "arg" Then
                    ReDim Preserve args(args.Length)
                    args(args.Length - 1) = defval
                    ReDim Preserve argspaths(argspaths.Length)
                    argspaths(argspaths.Length - 1) = deffilepath
                ElseIf deftype = "res" Then
                    ReDim Preserve results(results.Length)
                    results(results.Length - 1) = defval
                    ReDim Preserve resultspaths(resultspaths.Length)
                    resultspaths(resultspaths.Length - 1) = deffilepath
                ElseIf deftype = "diag" Then
                    ReDim Preserve diags(diags.Length)
                    diags(diags.Length - 1) = defval
                    ReDim Preserve diagspaths(diagspaths.Length)
                    diagspaths(diagspaths.Length - 1) = deffilepath
                ElseIf deftype = "script" Then
                    ReDim Preserve scripts(scripts.Length)
                    scripts(scripts.Length - 1) = defval
                    ReDim Preserve scriptspaths(scriptspaths.Length)
                    scriptspaths(scriptspaths.Length - 1) = deffilepath
                ElseIf deftype = "dir" Then
                    dirglobal = defval
                End If
            Next
            If rexec = "" Then Return "Error in getRDefinitions: no rexec defined"
            If scripts.Count = 0 Then Return "Error in getRDefinitions: no script(s) defined"
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
        If UBound(Rcalldefnames) = 0 Or IsNothing(MyFunctions.Rdefinitions) Then Exit Sub
        currWb = Wb
        ' get the definition range
        Dim errStr As String
        ' get the definition range
        errStr = getRDefinitions()
        If errStr <> vbNullString Then MsgBox("Error while getting Rdefinitions: " + errStr)

        errStr = storeInput()
        If errStr <> "" Then MsgBox("Error when saving inputfiles: " + errStr)
    End Sub

    Private Sub Workbook_Open(Wb As Workbook) Handles Application.WorkbookOpen
        currWb = Wb
        ' get the definition range
        Dim errStr As String
        errStr = getRNames()
        If errStr = "no Rdefinitions" Then Exit Sub
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
        If errStr = "no Rdefinitions" Then Exit Sub
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
        If UBound(Rcalldefnames) = 0 Then
            MsgBox("no Rdefinitions found for R_Addin (3 column named range (type/value/path), minimum types: rexec and script)!")
            Exit Sub
        End If
        If IsNothing(MyFunctions.Rdefinitions) Then
            MsgBox("no Rdefinition selected for starting R script!")
            Exit Sub
        End If
        errStr = MyFunctions.startRprocess()
        If errStr <> "" Then MsgBox(errStr)
    End Sub

    Public Function GetItemCount(control As ExcelDna.Integration.CustomUI.IRibbonControl) As Integer
        Return MyFunctions.Rcalldefnames.Length
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
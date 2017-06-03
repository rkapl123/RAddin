Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Configuration
Imports RDotNet
Imports RDotNet.NativeLibrary

Public Module RAddin

    Public rdotnetengine As REngine = Nothing
    Public currWb As Workbook
    Public Rdefinitions As Range
    Public Rcalldefnames As String() = {}
    Public Rcalldefs As Range() = {}
    Public rdefsheetColl As Dictionary(Of String, Dictionary(Of String, Range))
    Public rdefsheetMap As Dictionary(Of String, String)
    Public theRibbon As ExcelDna.Integration.CustomUI.IRibbonUI
    Public rexec As String
    Public rpath As String
    Public avoidFurtherMsgBoxes As Boolean
    Dim dirglobal As String

    ' definitions of current R invocation (scripts, args, results, diags...)
    Dim RdefDic As Dictionary(Of String, String()) = New Dictionary(Of String, String())

    ' prepare Parameters (script, args, results, diags) for usage in invokeScripts, storeArgs, getResults and getDiags 
    Private Function prepareParams(c As Integer, name As String, ByRef RDataRange As Range, ByRef returnName As String, ByRef returnPath As String, ext As String) As String
        Dim value As String = RdefDic(name)(c)
        ' only for args, results and diags (scripts dont have a target range)
        If name = "args" Or name = "results" Or name = "diags" Or name = "scriptrng" Then
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

    ' creates Inputfiles for defined arg ranges, tab separated, decimalpoint always ".", dates are stored as "yyyy-MM-dd" 
    ' otherwise:  "what you see is what you get"
    Public Function storeArgs() As Boolean
        Dim argFilename As String = vbNullString, argdir As String
        Dim RDataRange As Range = Nothing
        Dim outputFile As StreamWriter = Nothing

        argdir = dirglobal
        For c As Integer = 0 To RdefDic("args").Length - 1
            Try
                Dim errMsg As String
                errMsg = prepareParams(c, "args", RDataRange, argFilename, argdir, ".txt")
                If Len(errMsg) > 0 Then
                    If Not myMsgBox(errMsg) Then Return False
                End If

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
                    If RDataRange(i, 1).Value2 IsNot Nothing Then
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
                If Not myMsgBox("Error occured when creating inputfile '" + argFilename + "', " + ex.Message) Then Return False
            End Try
            If outputFile IsNot Nothing Then outputFile.Close()
        Next
        Return True
    End Function

    ' creates script files for defined scriptRng ranges 
    Private Function storeScriptRng() As Boolean
        Dim scriptRngFilename As String = vbNullString, scriptRngdir = vbNullString, scriptText = vbNullString
        Dim RDataRange As Range = Nothing
        Dim outputFile As StreamWriter = Nothing

        scriptRngdir = dirglobal
        For c As Integer = 0 To RdefDic("scriptrng").Length - 1
            Try
                Dim errMsg As String
                errMsg = prepareParams(c, "scriptrng", RDataRange, scriptRngFilename, scriptRngdir, ".R")
                If Len(errMsg) > 0 Then
                    scriptText = RdefDic("scriptrng")(c)
                    scriptRngFilename = "RDataRangeRow" + c.ToString() + ".R"
                End If

                ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\scriptRngdir
                Dim curWbPrefix As String = IIf(Left(scriptRngdir, 2) = "\\" Or Mid(scriptRngdir, 2, 2) = ":\", "", currWb.Path + "\")
                ' remove any existing input files...
                If File.Exists(curWbPrefix + scriptRngdir + "\" + scriptRngFilename) Then
                    File.Delete(curWbPrefix + scriptRngdir + "\" + scriptRngFilename)
                End If

                outputFile = New StreamWriter(curWbPrefix + scriptRngdir + "\" + scriptRngFilename)

                ' reuse the script invocation methods by setting the respective parameters
                ReDim Preserve RdefDic("debug")(RdefDic("debug").Length)
                RdefDic("debug")(RdefDic("debug").Length - 1) = RdefDic("debugrng")(c)
                ReDim Preserve RdefDic("scripts")(RdefDic("scripts").Length)
                RdefDic("scripts")(RdefDic("scripts").Length - 1) = scriptRngFilename
                ReDim Preserve RdefDic("scriptspaths")(RdefDic("scriptspaths").Length)
                RdefDic("scriptspaths")(RdefDic("scriptspaths").Length - 1) = scriptRngdir

                ' write the RDataRange or scriptText (if script directly in value) to file
                If Not IsNothing(scriptText) Then
                    outputFile.WriteLine(scriptText)
                Else
                    Dim i As Integer = 1
                    Do
                        Dim j As Integer = 1
                        Dim writtenLine As String = ""
                        If RDataRange(i, 1).Value2 IsNot Nothing Then
                            Do
                                writtenLine = writtenLine + RDataRange(i, j).Value2
                                j = j + 1
                            Loop Until j > RDataRange.Columns.Count
                            outputFile.WriteLine(writtenLine)
                        End If
                        i = i + 1
                    Loop Until i > RDataRange.Rows.Count
                End If
            Catch ex As Exception
                If outputFile IsNot Nothing Then outputFile.Close()
                If Not myMsgBox("Error occured when creating script file '" + scriptRngFilename + "', " + ex.Message) Then Return False
            End Try
            If outputFile IsNot Nothing Then outputFile.Close()
        Next
        Return True
    End Function

    ' invokes current scripts/args/results definition
    Private Function invokeScripts() As Boolean
        Dim script As String = vbNullString
        Dim scriptpath As String

        scriptpath = dirglobal
        For c As Integer = 0 To RdefDic("scripts").Length - 1
            Dim ErrMsg As String = prepareParams(c, "scripts", Nothing, script, scriptpath, "")
            If Len(ErrMsg) > 0 Then
                If Not myMsgBox(ErrMsg) Then Return False
            End If

                ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\scriptpath
                Dim curWbPrefix As String = IIf(Left(scriptpath, 2) = "\\" Or Mid(scriptpath, 2, 2) = ":\", "", currWb.Path + "\")
            Dim fullScriptPath = curWbPrefix + scriptpath
            If Not File.Exists(fullScriptPath + "\" + script) Then
                If Not myMsgBox("Script '" + fullScriptPath + "\" + script + "' not found!" + vbCrLf) Then Return False
            End If
            If Not File.Exists(rexec) And rexec <> "cmd" Then
                If Not myMsgBox("Executable '" + rexec + "' not found!" + vbCrLf) Then Return False
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
                If Not myMsgBox("Error occured when invoking script '" + fullScriptPath + "\" + script + "', using '" + rexec + "'" + ex.Message + vbCrLf) Then Return False
            End Try
        Next
        Return True
    End Function

    ' get Outputfiles for defined results ranges, tab separated
    ' otherwise:  "what you see is what you get"
    Private Function getResults() As Boolean
        Dim resFilename As String = vbNullString, readdir As String
        Dim RDataRange As Range = Nothing
        Dim errMsg As String = vbNullString

        readdir = dirglobal
        For c As Integer = 0 To RdefDic("results").Length - 1
            errMsg = prepareParams(c, "results", RDataRange, resFilename, readdir, ".txt")
            If Len(errMsg) > 0 Then
                If Not myMsgBox(errMsg) Then Return False
            End If

                ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\readdir
                Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            Dim infile As StreamReader = Nothing
            Try
                infile = New StreamReader(curWbPrefix + readdir + "\" + resFilename)
            Catch ex As Exception
                If Not myMsgBox("Error occured in getResults when opening '" + currWb.Path + "\" + readdir + "\" + resFilename + "', " + ex.Message) Then Return False
            End Try

            ' parse the actual file line by line
            Dim i As Integer = 1, currentRecord As String() = Nothing, currentLine As String
            RDataRange.ClearContents()
            Do While Not infile.EndOfStream
                Try
                    currentLine = infile.ReadLine
                    currentRecord = currentLine.Split(vbTab)
                Catch ex As Exception
                    If infile IsNot Nothing Then infile.Close()
                    If Not myMsgBox("Error occured in getResults when parsing file '" + resFilename + "', " + ex.Message) Then Return False
                End Try
                ' Put parsed data into target range column by column
                For j = 1 To currentRecord.Count()
                    Try
                        RDataRange.Cells(i, j).Value2 = currentRecord(j - 1)
                    Catch ex As Exception
                        If infile IsNot Nothing Then infile.Close()
                        If Not myMsgBox("Error occured in getResults when writing data into '" + RDataRange.Parent.name + "!" + RDataRange.Address + "', " + ex.Message) Then Return False
                    End Try
                Next
                i = i + 1
            Loop
            If infile IsNot Nothing Then infile.Close()
        Next
        Return True
    End Function

    ' get Output diagrams (png) for defined diags ranges
    Private Function getDiags() As Boolean
        Dim diagFilename As String = vbNullString, readdir As String
        Dim RDataRange As Range = Nothing
        Dim errMsg As String = vbNullString

        readdir = dirglobal
        For c As Integer = 0 To RdefDic("diags").Length - 1
            errMsg = prepareParams(c, "diags", RDataRange, diagFilename, readdir, ".png")
            If Len(errMsg) > 0 Then
                If Not myMsgBox(errMsg) Then Return False
            End If

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
                With RDataRange.Worksheet.Shapes.AddPicture(Filename:=curWbPrefix + readdir + "\" + diagFilename,
                    LinkToFile:=False, SaveWithDocument:=True, Left:=RDataRange.Left, Top:=RDataRange.Top, Width:=-1, Height:=-1)
                    .Name = diagFilename
                End With
            Catch ex As Exception
                If Not myMsgBox("Error occured when placing the diagram into target range '" + RdefDic("diags")(c) + "', " + ex.Message) Then Return False
            End Try
        Next
        Return True
    End Function

    ' remove result, diagram and temporary R script files
    Private Function removeFiles() As Boolean
        Dim filename As String = vbNullString, readdir As String
        Dim RDataRange As Range = Nothing
        Dim errMsg As String = vbNullString

        readdir = dirglobal
        ' remove result files
        For c As Integer = 0 To RdefDic("results").Length - 1
            errMsg = prepareParams(c, "results", RDataRange, filename, readdir, ".txt")
            If Len(errMsg) > 0 Then
                If Not myMsgBox(errMsg) Then Return False
            End If

                ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\argdir
                Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            ' remove any existing result files...
            If File.Exists(curWbPrefix + readdir + "\" + filename) Then
                Try
                    File.Delete(curWbPrefix + readdir + "\" + filename)
                Catch ex As Exception
                    If Not myMsgBox("Error occured when trying to remove '" + curWbPrefix + readdir + "\" + filename + "', " + ex.Message) Then Return False
                End Try
            End If
        Next
        ' remove diagram files
        For c As Integer = 0 To RdefDic("diags").Length - 1
            errMsg = prepareParams(c, "diags", RDataRange, filename, readdir, ".png")
            If Len(errMsg) > 0 Then
                If Not myMsgBox(errMsg) Then Return False
            End If

                ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\argdir
                Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            ' remove any existing diagram files...
            If File.Exists(curWbPrefix + readdir + "\" + filename) Then
                Try
                    File.Delete(curWbPrefix + readdir + "\" + filename)
                Catch ex As Exception
                    If Not myMsgBox("Error occured when trying to remove '" + curWbPrefix + readdir + "\" + filename + "', " + ex.Message) Then Return False
                End Try
            End If
        Next
        ' remove temporary R script files
        For c As Integer = 0 To RdefDic("scriptrng").Length - 1
            errMsg = prepareParams(c, "scriptrng", RDataRange, filename, readdir, ".R")
            If Len(errMsg) > 0 Then
                filename = "RDataRangeRow" + c.ToString() + ".R"
            End If

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\argdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            ' remove any existing diagram files...
            If File.Exists(curWbPrefix + readdir + "\" + filename) Then
                Try
                    File.Delete(curWbPrefix + readdir + "\" + filename)
                Catch ex As Exception
                    If Not myMsgBox("Error occured when trying to remove '" + curWbPrefix + readdir + "\" + filename + "', " + ex.Message) Then Return False
                End Try
            End If
        Next
        Return True
    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' startRprocess: started from GUI (button) and accessible from VBA (via Application.Run)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function startRprocess(runShell As Boolean, runRdotNet As Boolean) As String
        Dim errStr As String
        avoidFurtherMsgBoxes = False
        ' get the definition range
        errStr = getRDefinitions()
        If errStr <> vbNullString Then Return "Failed getting Rdefinitions: " + errStr
        If runShell Then ' shell invocation
            Try
                If Not removeFiles() Then Return vbNullString
                If Not storeArgs() Then Return vbNullString
                If Not storeScriptRng() Then Return vbNullString
                If Not invokeScripts() Then Return vbNullString
                If Not getResults() Then Return vbNullString
                If Not getDiags() Then Return vbNullString
            Catch ex As Exception
                Return "Exception in shell Rdefinitions run: " + ex.Message + ex.StackTrace
            End Try
        End If
        If runRdotNet Then ' RDotNet invocation
            Try
                errStr = initializeRDotNet()
                If errStr <> vbNullString Then Return "initializing RdotNet returned error: " + errStr
                If Not prepareParamsInvokeScriptsAndGetResults() Then Return vbNullString
            Catch ex As Exception
                Return "Exception in RdotNet Rdefinitions run: " + ex.Message + ex.StackTrace
            End Try
        End If
        ' all is OK = return nullstring
        Return vbNullString
    End Function

    ' Msgbox that avoids further Msgboxes (click Yes) or cancels run altogether (click Cancel)
    Private Function myMsgBox(message As String) As Boolean
        If avoidFurtherMsgBoxes Then Return True
        Dim retval As MsgBoxResult = MsgBox(message + vbCrLf + "Avoid further Messages (Yes/No) or abort Rdefinition (Cancel)", MsgBoxStyle.YesNoCancel)
        If retval = MsgBoxResult.Yes Then avoidFurtherMsgBoxes = True
        Return (retval = MsgBoxResult.Yes Or retval = MsgBoxResult.No)
    End Function

    ' initialize RdotNet engine
    Private Function initializeRDotNet() As String
        Dim logInfo As String = vbNullString
        Try
            Dim fullrpath As String = rpath + IIf(rpath.EndsWith("\"), "", "\") + IIf(System.Environment.Is64BitProcess, ConfigurationManager.AppSettings("rPath64bit"), ConfigurationManager.AppSettings("rPath32bit"))
            Dim rHome As String = vbNullString
            logInfo = NativeUtility.FindRPaths(fullrpath, rHome) + ", Is64BitProcess: " + System.Environment.Is64BitProcess.ToString()
            REngine.SetEnvironmentVariables(rPath:=fullrpath)
            rdotnetengine = REngine.GetInstance()
            rdotnetengine.Initialize()
        Catch ex As Exception
            Return "Error initializing RDotNet: " + ex.Message + ",logInfo from FindRPaths: " + logInfo
        End Try
        Return vbNullString
    End Function

    ' prepare arguments for rdotnet, run scripts and return results to excel
    Private Function prepareParamsInvokeScriptsAndGetResults() As String

        ' First import arguments...
        For c As Integer = 0 To RdefDic("args").Length - 1
            Dim argname As String = RdefDic("args")(c)
            ' if argvalue refers to a WS Name, cut off WS name prefix for R argname...
            Dim posWSseparator = InStr(argname, "!")
            If posWSseparator > 0 Then argname = argname.Substring(posWSseparator)
            Dim RDataRange As Range = currWb.Names.Item(RdefDic("args")(c)).RefersToRange
            Dim dfDataColumns() As IEnumerable
            ReDim dfDataColumns(RDataRange.Columns.Count - 1)

            Try
                ' write the RDataRange to dfDataColumns
                Dim j As Integer = 1
                Do
                    Dim columnValues(RDataRange.Rows.Count) As String
                    Dim i As Integer = 1
                    Do
                        If RDataRange(i, 1).Value2 IsNot Nothing Then
                            columnValues(i) = RDataRange(i, j).Value2
                        End If
                        i = i + 1
                    Loop Until i > RDataRange.Rows.Count
                    dfDataColumns(j - 1) = columnValues
                    j = j + 1
                Loop Until j > RDataRange.Columns.Count
                ' write the dfDataColumns to rdotnet dataframe
                Dim targetArg As RDotNet.DataFrame = rdotnetengine.CreateDataFrame(dfDataColumns)
                ' set the symbol to the correct name
                rdotnetengine.SetSymbol(argname, targetArg)
            Catch ex As Exception
                myMsgBox("Error occured when creating RdotNet arg '" + argname + "', " + ex.Message)
            End Try
        Next

        ' then evaluate excel stored scripts
        For c As Integer = 0 To RdefDic("scriptrng").Length - 1
            Dim scriptText As String = Nothing
            Dim RDataRange As Range = Nothing
            Try
                RDataRange = currWb.Names.Item(RdefDic("scriptrng")(c)).RefersToRange
            Catch
                scriptText = RdefDic("scriptrng")(c)
            End Try

            If Not IsNothing(scriptText) Then
                Try
                    rdotnetengine.Evaluate(scriptText)
                Catch ex As Exception
                    myMsgBox("Error occured when evaluating script '" + scriptText + "', " + ex.Message)
                End Try
            Else
                Dim i As Integer = 1
                Do
                    Dim j As Integer = 1
                    Dim evalLine As String = ""
                    If RDataRange(i, 1).Value2 IsNot Nothing Then
                        Do
                            Try
                                evalLine = RDataRange(i, j).Value2
                                rdotnetengine.Evaluate(evalLine)
                            Catch ex As Exception
                                myMsgBox("Error occured when evaluating script line '" + evalLine + "', " + ex.Message)
                            End Try
                            j = j + 1
                        Loop Until j > RDataRange.Columns.Count
                    End If
                    i = i + 1
                Loop Until i > RDataRange.Rows.Count
            End If
        Next

        ' then evaluate filesystem stored scripts
        Dim scriptpath As String = dirglobal
        For c As Integer = 0 To RdefDic("scripts").Length - 1
            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\scriptRngdir
            Dim curWbPrefix As String = IIf(Left(scriptpath, 2) = "\\" Or Mid(scriptpath, 2, 2) = ":\", "", currWb.Path + "\")

            Dim scriptname As String = RdefDic("scripts")(c)
            Try
                rdotnetengine.Evaluate("source('" + curWbPrefix + scriptpath + "\" + scriptname + "')")
            Catch ex As Exception
                myMsgBox("Error occured when evaluating script '" + curWbPrefix + scriptpath + "\" + scriptname + "', " + ex.Message)
            End Try
        Next

        ' then evaluate (return) resultnames
        For c As Integer = 0 To RdefDic("results").Length - 1
            Dim resname As String = RdefDic("results")(c)
            ' if argvalue refers to a WS Name, cut off WS name prefix for R argname...
            Dim posWSseparator = InStr(resname, "!")
            If posWSseparator > 0 Then resname = resname.Substring(posWSseparator)

            ' first get data from RdotNet engine
            Dim resultDataSymbolicExpr As SymbolicExpression
            Dim resultData As Object = Nothing
            Dim columnCount As Integer : Dim rowCount As Integer
            Try
                resultDataSymbolicExpr = rdotnetengine.Evaluate(resname)
                If resultDataSymbolicExpr.IsDataFrame() Then
                    resultData = resultDataSymbolicExpr.AsDataFrame()
                    columnCount = resultDataSymbolicExpr.AsDataFrame().ColumnCount
                    rowCount = resultDataSymbolicExpr.AsDataFrame().RowCount
                ElseIf resultDataSymbolicExpr.IsMatrix() Then
                    resultData = resultDataSymbolicExpr.AsRawMatrix()
                    columnCount = resultDataSymbolicExpr.AsRawMatrix().ColumnCount
                    rowCount = resultDataSymbolicExpr.AsRawMatrix().RowCount
                ElseIf resultDataSymbolicExpr.IsVector() Then
                    resultData = resultDataSymbolicExpr.AsRaw()
                    columnCount = 1
                    rowCount = resultDataSymbolicExpr.AsRaw().Count()
                ElseIf resultDataSymbolicExpr.IsList() Then
                    resultData = resultDataSymbolicExpr.AsList()
                    columnCount = 1
                    rowCount = resultDataSymbolicExpr.AsList().Count()
                End If

            Catch ex As Exception
                myMsgBox("Error occured when evaluating result '" + resname + "', " + ex.Message)
            End Try

            ' then put data into excel range
            Dim RDataRange As Range = currWb.Names.Item(RdefDic("results")(c)).RefersToRange
            Dim i As Integer = 0
            Do
                If columnCount = 1 Then
                    RDataRange(i + 1, 1).Value2 = resultData(i)
                Else
                    Dim j As Integer = 0
                    Do
                        Try
                            RDataRange(i + 1, j + 1).Value2 = resultData(i, j)
                        Catch ex As Exception
                            myMsgBox("Error occured when putting result '" + resname + "', " + ex.Message)
                        End Try
                        j = j + 1
                    Loop Until j > columnCount - 1
                End If
                i = i + 1
            Loop Until i > rowCount - 1
        Next

        Return vbNullString
    End Function

    Public Function startRnamesRefresh() As String
        Dim errStr As String
        ' always reset Rdefinitions when changing Workbooks, otherwise this is not being refilled in getRNames
        Rdefinitions = Nothing
        ' get the defined R_Addin Names
        errStr = getRNames()
        If errStr = "no definitions" Then
            Return vbNullString
        ElseIf errStr <> vbNullString Then
            Return "Error while getRNames in startRnamesRefresh: " + errStr
        End If
        RAddin.theRibbon.Invalidate()
        Return vbNullString
    End Function

    ' gets defined named ranges for R script invocation in the current workbook 
    Public Function getRNames() As String
        ReDim Preserve Rcalldefnames(-1)
        ReDim Preserve Rcalldefs(-1)
        rdefsheetColl = New Dictionary(Of String, Dictionary(Of String, Range))
        rdefsheetMap = New Dictionary(Of String, String)
        Dim i As Integer = 0
        For Each namedrange As Name In currWb.Names
            Dim cleanname As String = Replace(namedrange.Name, namedrange.Parent.Name & "!", "")
            If Left(cleanname, 7) = "R_Addin" Then
                If namedrange.RefersToRange.Columns.Count <> 3 Then Return "Rdefinitions range " + namedrange.Parent.name + "!" + namedrange.Name + " doesn't have 3 columns !"
                ' final name of entry is without R_Addin and !
                Dim finalname As String = Replace(Replace(namedrange.Name, "R_Addin", ""), "!", "")
                Dim nodeName As String = Replace(Replace(namedrange.Name, "R_Addin", ""), namedrange.Parent.Name & "!", "")
                If nodeName = "" Then nodeName = "MainScript"
                ' first definition as standard definition (works without selecting a Rdefinition)
                If Rdefinitions Is Nothing Then Rdefinitions = namedrange.RefersToRange
                If Not InStr(namedrange.Name, "!") > 0 Then
                    finalname = currWb.Name + finalname
                End If
                ReDim Preserve Rcalldefnames(Rcalldefnames.Length)
                ReDim Preserve Rcalldefs(Rcalldefs.Length)
                Rcalldefnames(Rcalldefnames.Length - 1) = finalname
                Rcalldefs(Rcalldefs.Length - 1) = namedrange.RefersToRange
                If Not rdefsheetColl.ContainsKey(namedrange.Parent.Name) Then
                    ' add to new sheet "menu"
                    Dim scriptColl As Dictionary(Of String, Range) = New Dictionary(Of String, Range)
                    scriptColl.Add(nodeName, namedrange.RefersToRange)
                    rdefsheetColl.Add(namedrange.Parent.Name, scriptColl)
                    rdefsheetMap.Add("ID" + i.ToString(), namedrange.Parent.Name)
                    i = i + 1
                Else
                    ' add rdefinition to existing sheet "menu"
                    Dim scriptColl As Dictionary(Of String, Range)
                    scriptColl = rdefsheetColl(namedrange.Parent.Name)
                    scriptColl.Add(nodeName, namedrange.RefersToRange)
                End If
            End If
        Next
        If UBound(Rcalldefnames) = -1 Then Return "no RNames"
        Return vbNullString
    End Function

    ' gets definitions from  current selected R script invocation range (Rdefinitions)
    Public Function getRDefinitions() As String
        Try
            RdefDic("args") = {}
            RdefDic("argspaths") = {}
            RdefDic("results") = {}
            RdefDic("resultspaths") = {}
            RdefDic("diags") = {}
            RdefDic("diagspaths") = {}
            RdefDic("scripts") = {}
            RdefDic("scriptspaths") = {}
            RdefDic("scriptrng") = {}
            RdefDic("scriptrngpaths") = {}
            RdefDic("debugrng") = {}
            RdefDic("debug") = {}
            rpath = Nothing : rexec = Nothing : dirglobal = vbNullString
            For Each defRow As Range In Rdefinitions.Rows
                Dim deftype As String, defval As String, deffilepath As String
                deftype = LCase(defRow.Cells(1, 1).Value2)
                defval = defRow.Cells(1, 2).Value2
                deffilepath = defRow.Cells(1, 3).Value2
                If deftype = "rexec" Then ' setting for shell innvocation
                    rexec = defval
                ElseIf deftype = "rpath" Then ' setting for RdotNet
                    rpath = defval
                ElseIf deftype = "arg" Then
                    ReDim Preserve RdefDic("args")(RdefDic("args").Length)
                    RdefDic("args")(RdefDic("args").Length - 1) = defval
                    ReDim Preserve RdefDic("argspaths")(RdefDic("argspaths").Length)
                    RdefDic("argspaths")(RdefDic("argspaths").Length - 1) = deffilepath
                ElseIf deftype = "scriptrng" Or deftype = "debugrng" Then
                    ReDim Preserve RdefDic("debugrng")(RdefDic("debugrng").Length)
                    RdefDic("debugrng")(RdefDic("debugrng").Length - 1) = (deftype = "debugrng")
                    ReDim Preserve RdefDic("scriptrng")(RdefDic("scriptrng").Length)
                    RdefDic("scriptrng")(RdefDic("scriptrng").Length - 1) = defval
                    ReDim Preserve RdefDic("scriptrngpaths")(RdefDic("scriptrngpaths").Length)
                    RdefDic("scriptrngpaths")(RdefDic("scriptrngpaths").Length - 1) = deffilepath
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
            ' get default rexec path from user (or overriden in appSettings tag as redirect to global) settings. This can be overruled by individual rexec settings in Rdefinitions
            Try
                If rexec Is Nothing Then rexec = ConfigurationManager.AppSettings("ExePath")
            Catch ex As Exception
            End Try
            ' get default Rpath from user (or overriden in appSettings tag as redirect to global) settings. This can be overruled by individual rexec settings in Rdefinitions
            Try
                If rpath Is Nothing Then rpath = ConfigurationManager.AppSettings("rPath")
            Catch ex As Exception
            End Try
            If rexec Is Nothing And rpath Is Nothing Then Return "Error in getRDefinitions: neither rexec nor rpath defined"
            If RdefDic("scripts").Count = 0 And RdefDic("scriptrng").Count = 0 Then Return "Error in getRDefinitions: no script(s) or scriptRng(s) defined"
        Catch ex As Exception
            Return "Error in getRDefinitions: " + ex.Message
        End Try
        Return vbNullString
    End Function
End Module

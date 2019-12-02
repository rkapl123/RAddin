Imports RDotNet
Imports RDotNet.NativeLibrary
Imports Microsoft.Office.Interop.Excel
Imports RDotNet.Devices
Imports RDotNet.Internals

''' <summary>all functions for the Rdotnet invocation method (here the scripts and input args are passed to an inmemory engine and results are retrieved from there with one exception: graphics are still taken from the filesystem)</summary>
Module RdotnetInvocation
    ''' <summary></summary>
    Public rDotNetEngine As REngine = Nothing
    ''' <summary></summary>
    Public rPath As String
    ''' <summary></summary>
    Public rHome As String
    ''' <summary>the "console" used to capture output from R.NET engine, actually writing to Trace</summary>
    Public debugConsole As DebugDevice

    '''<summary>initialize RdotNet engine</summary> 
    ''' <returns>True if success, False otherwise</returns>
    Public Function initializeRDotNet() As Boolean
        ' only instantiate new engine if there is none already (reuse engine !)
        If IsNothing(rDotNetEngine) Then
            Dim logOutput As String = ""
            Try
                REngine.SetEnvironmentVariables(rPath:=rPath, rHome:=rHome)
                logOutput = (New NativeUtility(Nothing)).FindRPaths(rPath, rHome)
                Dim sp As RDotNet.StartupParameter = New StartupParameter With {
                    .Quiet = False,
                    .Interactive = False
                }
                debugConsole = New DebugDevice()
                rDotNetEngine = REngine.GetInstance(device:=debugConsole, parameter:=sp, initialize:=True)
                If Not rDotNetEngine.IsRunning Then rDotNetEngine.Initialize(device:=debugConsole, parameter:=sp, setupMainLoop:=True)
                ' avoid autoprint as this kills performance
                rDotNetEngine.AutoPrint = False
                ' set starting directory to current workbooks path
                rDotNetEngine.Evaluate("setwd(""" + Replace(currWb.Path, "\", "/") + """)")
                LogInfo("initialized RDotNet, FindRPaths Log: " + logOutput)
            Catch ex As Exception
                If Not RAddin.myMsgBox("Error initializing RDotNet: " + ex.Message + ",FindRPaths Log: " + logOutput) Then Return False
            End Try
        End If
        ' remove diagramfiles, in case folder structure is not existing, create it...
        If Not removeDiagFiles() Then Return False
        Return True
    End Function

    ''' <summary>import arguments into RdotNetEngine</summary>
    ''' <returns>True if success, False otherwise</returns>
    Public Function storeArgs() As Boolean
        For c As Integer = 0 To RdefDic("args").Length - 1
            Dim argname As String = RdefDic("args")(c)
            Dim rowcolumn As String = RdefDic("argsrc")(c) ' contains r if ranges rows start with row names, c if ranges columns start with column names
            ' if argvalue refers to a WS Name, cut off WS name prefix for R argname...
            Dim posWSseparator = InStr(argname, "!")
            If posWSseparator > 0 Then argname = argname.Substring(posWSseparator)
            Dim RDataRange As Range = currWb.Names.Item(RdefDic("args")(c)).RefersToRange
            Dim rowOffset As Integer = IIf(InStr(rowcolumn, "c") > 0, 1, 0)
            Dim colOffset As Integer = IIf(InStr(rowcolumn, "r") > 0, 1, 0)
            Dim dfDataColumns(RDataRange.Columns.Count - 1 - colOffset) As IEnumerable
            Dim columnNames(RDataRange.Columns.Count - 1 - colOffset) As String
            Dim rowNames(RDataRange.Rows.Count - rowOffset) As String
            Try
                ' write the RDataRange to dfDataColumns
                Dim j As Integer = 1
                Do
                    Dim columnValues(RDataRange.Rows.Count - 1 - IIf(InStr(rowcolumn, "c") > 0, 1, 0)) As String
                    Dim i As Integer = 1
                    Do
                        ' do we have to write row names?
                        If j = 1 And InStr(rowcolumn, "r") > 0 Then
                            rowNames(i - 1) = IIf(RDataRange(i, j).Value2 IsNot Nothing, RDataRange(i, j).Value2, Nothing)
                        Else
                            ' do we have to write column names?
                            If i = 1 And InStr(rowcolumn, "c") > 0 Then
                                columnNames(j - 1) = IIf(RDataRange(i, j).Value2 IsNot Nothing, RDataRange(i, j).Value2, Nothing)
                            Else
                                columnValues(i - 1 - rowOffset) = IIf(RDataRange(i, j).Value2 IsNot Nothing, IIf(IsNumeric(RDataRange(i, j).Value2), Replace(RDataRange(i, j).Value2, ",", "."), RDataRange(i, j).Value2), Nothing)
                            End If
                        End If
                        i += 1
                    Loop Until i > RDataRange.Rows.Count
                    dfDataColumns(j - 1) = columnValues
                    j += 1
                Loop Until j > RDataRange.Columns.Count
                ' write the dfDataColumns to rdotnet dataframe
                Dim targetArg As RDotNet.DataFrame = rDotNetEngine.CreateDataFrame(dfDataColumns, columnNames:=IIf(InStr(rowcolumn, "c"), columnNames, Nothing), rowNames:=IIf(InStr(rowcolumn, "r"), rowNames, Nothing))
                ' set the symbol to the correct name
                rDotNetEngine.SetSymbol(argname, targetArg)
                LogInfo("set symbol to name, argname:" + argname)
            Catch ex As Exception
                If Not RAddin.myMsgBox("Error occured when creating RdotNet arg '" + argname + "', " + ex.Message) Then Return False
            End Try
        Next
        Return True
    End Function

    ''' <summary>invokes current range stored scripts/args/results</summary>
    ''' <returns>True if success, False otherwise</returns>
    Public Function invokeExcelScripts() As Boolean
        ' then evaluate excel stored scripts
        For c As Integer = 0 To RdefDic("scriptrng").Length - 1
            Dim scriptText As String = Nothing
            Dim RDataRange As Range = Nothing
            If Left(RdefDic("scriptrng")(c), 1) = "=" Then
                scriptText = RdefDic("scriptrng")(c).Substring(1)
            Else
                Try
                    RDataRange = currWb.Names.Item(RdefDic("scriptrng")(c)).RefersToRange
                Catch ex As Exception
                    If Not RAddin.myMsgBox("Error occured when getting range for '" + RdefDic("scriptrng")(c) + "', " + ex.Message) Then Return False
                End Try
            End If

            If Not IsNothing(scriptText) Then
                Try
                    rDotNetEngine.Evaluate(scriptText)
                    LogInfo("Evaluated scriptText: " + scriptText)
                Catch ex As Exception
                    If Not RAddin.myMsgBox("Error occured when evaluating script '" + scriptText + "', " + ex.Message) Then Return False
                End Try
            Else
                Dim i As Integer = 1
                Do
                    Dim j As Integer = 1
                    If RDataRange(i, 1).Value2 IsNot Nothing Then
                        Do
                            scriptText += RDataRange(i, j).Value2 + vbCrLf
                            j += 1
                        Loop Until j > RDataRange.Columns.Count
                        Try
                            rDotNetEngine.Evaluate(scriptText)
                        Catch ex As Exception
                            If Not RAddin.myMsgBox("Error occured when evaluating script '" + scriptText + "', " + ex.Message) Then Return False
                        End Try
                    End If
                    i += 1
                Loop Until i > RDataRange.Rows.Count
            End If
        Next
        Return True
    End Function

    ''' <summary>invokes current filesystem stored scripts/args/results</summary>
    ''' <returns>True if success, False otherwise</returns>
    Public Function invokeFileSysScripts() As Boolean
        ' then evaluate filesystem stored scripts
        Dim scriptpath As String = dirglobal
        For c As Integer = 0 To RdefDic("scripts").Length - 1
            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\scriptRngdir
            Dim curWbPrefix As String = IIf(Left(scriptpath, 2) = "\\" Or Mid(scriptpath, 2, 2) = ":\", "", currWb.Path + "\")

            Dim scriptname As String = RdefDic("scripts")(c)
            Try
                rDotNetEngine.Evaluate("source('" + curWbPrefix + scriptpath + "\" + scriptname + "')")
                LogInfo("Evaluated script in: " + +curWbPrefix + scriptpath + "\" + scriptname)
            Catch ex As Exception
                If Not RAddin.myMsgBox("Error occured when evaluating script '" + curWbPrefix + scriptpath + "\" + scriptname + "', " + ex.Message) Then Return False
            End Try
        Next
        Return True
    End Function


    ''' <summary>get Outputfiles for defined results ranges, tab separated
    ''' otherwise:  "what you see is what you get"
    ''' </summary>
    ''' <returns>True if success, False otherwise</returns>
    Public Function getResults() As Boolean
        ' then evaluate (return) resultnames
        For c As Integer = 0 To RdefDic("results").Length - 1
            Dim resname As String = RdefDic("results")(c)
            ' if argvalue refers to a WS Name, cut off WS name prefix for R argname...
            Dim posWSseparator = InStr(resname, "!")
            If posWSseparator > 0 Then resname = resname.Substring(posWSseparator)

            ' first get data from RdotNet engine
            Dim resultDataSymbolicExpr As SymbolicExpression
            Dim resultData As Object = Nothing
            Dim columnNames As Object = Nothing
            Dim rowNames As Object = Nothing
            Dim columnCount As Integer : Dim rowCount As Integer
            Try
                resultDataSymbolicExpr = rDotNetEngine.Evaluate(resname)
                If resultDataSymbolicExpr.IsDataFrame() Then
                    resultData = resultDataSymbolicExpr.AsDataFrame()
                    columnNames = resultDataSymbolicExpr.AsDataFrame().ColumnNames
                    rowNames = resultDataSymbolicExpr.AsDataFrame().RowNames
                    If IsNumeric(rowNames(0)) Then rowNames = Nothing
                    columnCount = resultDataSymbolicExpr.AsDataFrame().ColumnCount
                    rowCount = resultDataSymbolicExpr.AsDataFrame().RowCount
                ElseIf resultDataSymbolicExpr.IsMatrix() Then
                    resultData = resultDataSymbolicExpr.AsCharacterMatrix()
                    columnCount = resultDataSymbolicExpr.AsCharacterMatrix().ColumnCount
                    rowCount = resultDataSymbolicExpr.AsCharacterMatrix().RowCount
                ElseIf resultDataSymbolicExpr.IsVector() Then
                    resultData = resultDataSymbolicExpr.AsCharacter()
                    columnCount = 1
                    rowCount = resultDataSymbolicExpr.AsCharacter().Count()
                ElseIf resultDataSymbolicExpr.IsList() Then
                    resultData = resultDataSymbolicExpr.AsList()
                    columnCount = 1
                    rowCount = resultDataSymbolicExpr.AsList().Count()
                End If

            Catch ex As Exception
                If Not RAddin.myMsgBox("Error occured when evaluating result '" + resname + "', " + ex.Message) Then Return False
            End Try
            ' if we have row names then add one column
            Dim columnOffset As Integer = IIf(Not IsNothing(rowNames), 1, 0)
            ' if we have column names then add one row
            Dim rowOffset As Integer = IIf(Not IsNothing(columnNames), 1, 0)

            ' put data into excel range
            Dim RDataRange As Range = currWb.Names.Item(RdefDic("results")(c)).RefersToRange
            Dim j As Integer = 0
            Do
                Try
                    If Not IsNothing(columnNames) Then RDataRange(1, j + 1 + columnOffset).Value2 = columnNames(j)
                Catch ex As Exception
                    If Not RAddin.myMsgBox("Error occured when putting headers from '" + resname + "', " + ex.Message) Then Return False
                End Try
                j += 1
            Loop Until j > columnCount - 1

            Dim i As Integer = 0
            Do
                If columnCount = 1 Then
                    RDataRange(i + 1, 1).Value2 = resultData(i)
                Else
                    j = 0
                    Do
                        Try
                            If Not IsNothing(rowNames) And j = 0 Then
                                RDataRange(i + 1 + rowOffset, j + 1).Value2 = rowNames(i)
                            Else
                                RDataRange(i + 1 + rowOffset, j + 1).Value2 = resultData(i, j - columnOffset)
                            End If
                        Catch ex As Exception
                            If Not RAddin.myMsgBox("Error occured when putting result '" + resname + "', " + ex.Message) Then Return False
                        End Try
                        j += 1
                    Loop Until j > columnCount - 1 + columnOffset
                End If
                i += 1
            Loop Until i > rowCount - 1
            LogInfo("Put results in resname:" + resname)
        Next
        Return True
    End Function

    ''' <summary>get Output diagrams (png) for defined diags ranges</summary>
    ''' <returns>True if success, False otherwise</returns>
    Public Function getDiags() As Boolean
        'currently implemented as workaround via file creation...
        Return RscriptInvocation.getDiags()
    End Function

End Module

Public Class DebugDevice
    Implements RDotNet.Devices.ICharacterDevice
    Private Sub WriteConsole(output As String, length As Integer, outputType As ConsoleOutputType) Implements ICharacterDevice.WriteConsole
        If RAddin.debugScript Then Trace.WriteLine("R.NET: " + output)
    End Sub

    Public Sub ShowMessage(message As String) Implements ICharacterDevice.ShowMessage
        If RAddin.debugScript Then Trace.WriteLine("R.NET: " + message)
    End Sub

    Public Sub Busy(which As BusyType) Implements ICharacterDevice.Busy
    End Sub

    Public Sub Callback() Implements ICharacterDevice.Callback
    End Sub

    Public Sub Suicide(message As String) Implements ICharacterDevice.Suicide
        Trace.WriteLine("R.NET(Suicide): " + message)
    End Sub

    Public Sub ResetConsole() Implements ICharacterDevice.ResetConsole
    End Sub

    Public Sub FlushConsole() Implements ICharacterDevice.FlushConsole
    End Sub

    Public Sub ClearErrorConsole() Implements ICharacterDevice.ClearErrorConsole
    End Sub

    Public Sub CleanUp(saveAction As StartupSaveAction, status As Integer, runLast As Boolean) Implements ICharacterDevice.CleanUp
    End Sub

    Public Sub EditFile(file As String) Implements ICharacterDevice.EditFile
    End Sub

    Public Function ReadConsole(prompt As String, capacity As Integer, history As Boolean) As String Implements ICharacterDevice.ReadConsole
        Return Nothing
    End Function

    Public Function Ask(question As String) As YesNoCancel Implements ICharacterDevice.Ask
        Return YesNoCancel.No
    End Function

    Public Function ShowFiles(files() As String, headers() As String, title As String, delete As Boolean, pager As String) As Boolean Implements ICharacterDevice.ShowFiles
        Return False
    End Function

    Public Function ChooseFile(create As Boolean) As String Implements ICharacterDevice.ChooseFile
        Return Nothing
    End Function

    Public Function LoadHistory([call] As Language, operation As SymbolicExpression, args As Pairlist, environment As REnvironment) As SymbolicExpression Implements ICharacterDevice.LoadHistory
        Return environment.Engine.NilValue
    End Function

    Public Function SaveHistory([call] As Language, operation As SymbolicExpression, args As Pairlist, environment As REnvironment) As SymbolicExpression Implements ICharacterDevice.SaveHistory
        Return environment.Engine.NilValue
    End Function

    Public Function AddHistory([call] As Language, operation As SymbolicExpression, args As Pairlist, environment As REnvironment) As SymbolicExpression Implements ICharacterDevice.AddHistory
        Return environment.Engine.NilValue
    End Function

End Class
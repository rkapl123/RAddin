Imports RDotNet
Imports RDotNet.NativeLibrary
Imports Microsoft.Office.Interop.Excel

Module RdotnetInvocation
    Public rdotnetengine As REngine = Nothing
    Public rPath As String
    Public rHome As String

    ' initialize RdotNet engine
    Public Function initializeRDotNet() As Boolean
        Dim logInfo As String = vbNullString
        ' only instantiate new engine if there is none already (reuse engine !)
        If IsNothing(rdotnetengine) Then
            Try
                logInfo = NativeUtility.FindRPaths(rPath, rHome) + ", Is64BitProcess: " + System.Environment.Is64BitProcess.ToString()
                REngine.SetEnvironmentVariables(rPath:=rPath, rHome:=rHome)
                rdotnetengine = REngine.GetInstance()
                rdotnetengine.Initialize()
            Catch ex As Exception
                If Not RAddin.myMsgBox("Error initializing RDotNet: " + ex.Message + ",logInfo from FindRPaths: " + logInfo) Then Return False
            End Try
        End If
        Return True
    End Function

    ' import arguments into RdotNetEngine
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
                        i = i + 1
                    Loop Until i > RDataRange.Rows.Count
                    dfDataColumns(j - 1) = columnValues
                    j = j + 1
                Loop Until j > RDataRange.Columns.Count
                ' write the dfDataColumns to rdotnet dataframe
                Dim targetArg As RDotNet.DataFrame = rdotnetengine.CreateDataFrame(dfDataColumns, columnNames:=IIf(InStr(rowcolumn, "c"), columnNames, Nothing), rowNames:=IIf(InStr(rowcolumn, "r"), rowNames, Nothing))
                ' set the symbol to the correct name
                rdotnetengine.SetSymbol(argname, targetArg)
            Catch ex As Exception
                If Not RAddin.myMsgBox("Error occured when creating RdotNet arg '" + argname + "', " + ex.Message) Then Return False
            End Try
        Next
        Return True
    End Function

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
                    rdotnetengine.Evaluate(scriptText)
                Catch ex As Exception
                    If Not RAddin.myMsgBox("Error occured when evaluating script '" + scriptText + "', " + ex.Message) Then Return False
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
                                If Not RAddin.myMsgBox("Error occured when evaluating script line '" + evalLine + "', " + ex.Message) Then Return False
                            End Try
                            j = j + 1
                        Loop Until j > RDataRange.Columns.Count
                    End If
                    i = i + 1
                Loop Until i > RDataRange.Rows.Count
            End If
        Next
        Return True
    End Function

    Public Function invokeFileSysScripts() As Boolean
        ' then evaluate filesystem stored scripts
        Dim scriptpath As String = dirglobal
        For c As Integer = 0 To RdefDic("scripts").Length - 1
            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\scriptRngdir
            Dim curWbPrefix As String = IIf(Left(scriptpath, 2) = "\\" Or Mid(scriptpath, 2, 2) = ":\", "", currWb.Path + "\")

            Dim scriptname As String = RdefDic("scripts")(c)
            Try
                rdotnetengine.Evaluate("source('" + curWbPrefix + scriptpath + "\" + scriptname + "')")
            Catch ex As Exception
                If Not RAddin.myMsgBox("Error occured when evaluating script '" + curWbPrefix + scriptpath + "\" + scriptname + "', " + ex.Message) Then Return False
            End Try
        Next
        Return True
    End Function

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
                resultDataSymbolicExpr = rdotnetengine.Evaluate(resname)
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
                j = j + 1
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
                        j = j + 1
                    Loop Until j > columnCount - 1 + columnOffset
                End If
                i = i + 1
            Loop Until i > rowCount - 1
        Next
        Return True
    End Function

    Public Function getDiags() As Boolean
        'currently implemented as workaround via file creation...
        Return RscriptInvocation.getDiags()
    End Function

End Module

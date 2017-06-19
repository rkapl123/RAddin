Imports RDotNet
Imports RDotNet.NativeLibrary
Imports System.Configuration
Imports Microsoft.Office.Interop.Excel

Module RdotnetInvocation
    Public rdotnetengine As REngine = Nothing

    ' initialize RdotNet engine
    Public Function initializeRDotNet() As Boolean
        Dim logInfo As String = vbNullString
        Try
            Dim fullrpath As String = rpath + IIf(rpath.EndsWith("\"), "", "\") + IIf(System.Environment.Is64BitProcess, ConfigurationManager.AppSettings("rPath64bit"), ConfigurationManager.AppSettings("rPath32bit"))
            Dim rHome As String = vbNullString
            logInfo = NativeUtility.FindRPaths(fullrpath, rHome) + ", Is64BitProcess: " + System.Environment.Is64BitProcess.ToString()
            REngine.SetEnvironmentVariables(rPath:=fullrpath)
            rdotnetengine = REngine.GetInstance()
            rdotnetengine.Initialize()
        Catch ex As Exception
            If Not RAddin.myMsgBox("Error initializing RDotNet: " + ex.Message + ",logInfo from FindRPaths: " + logInfo) Then Return False
        End Try
        Return True
    End Function

    ' prepare arguments for rdotnet, run scripts and return results to excel
    Public Function prepareParamsInvokeScriptsAndGetResults() As String

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
End Module

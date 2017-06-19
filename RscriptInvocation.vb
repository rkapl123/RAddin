Imports System.IO
Imports Microsoft.Office.Interop.Excel

Module RscriptInvocation
    Public rexec As String

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
                    If Not RAddin.myMsgBox(errMsg) Then Return False
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
                If Not RAddin.myMsgBox("Error occured when creating inputfile '" + argFilename + "', " + ex.Message) Then Return False
            End Try
            If outputFile IsNot Nothing Then outputFile.Close()
        Next
        Return True
    End Function

    ' creates script files for defined scriptRng ranges 
    Public Function storeScriptRng() As Boolean
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

                ' write the RDataRange or scriptText (if script directly in cell/formula right next to scriptrng) to file
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
                If Not RAddin.myMsgBox("Error occured when creating script file '" + scriptRngFilename + "', " + ex.Message) Then Return False
            End Try
            If outputFile IsNot Nothing Then outputFile.Close()
        Next
        Return True
    End Function

    ' invokes current scripts/args/results definition
    Public Function invokeScripts() As Boolean
        Dim script As String = vbNullString
        Dim scriptpath As String

        scriptpath = dirglobal
        For c As Integer = 0 To RdefDic("scripts").Length - 1
            Dim ErrMsg As String = prepareParams(c, "scripts", Nothing, script, scriptpath, "")
            If Len(ErrMsg) > 0 Then
                If Not RAddin.myMsgBox(ErrMsg) Then Return False
            End If

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\scriptpath
            Dim curWbPrefix As String = IIf(Left(scriptpath, 2) = "\\" Or Mid(scriptpath, 2, 2) = ":\", "", currWb.Path + "\")
            Dim fullScriptPath = curWbPrefix + scriptpath
            If Not File.Exists(fullScriptPath + "\" + script) Then
                If Not RAddin.myMsgBox("Script '" + fullScriptPath + "\" + script + "' not found!" + vbCrLf) Then Return False
            End If
            If Not File.Exists(rexec) And rexec <> "cmd" Then
                If Not RAddin.myMsgBox("Executable '" + rexec + "' not found!" + vbCrLf) Then Return False
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
                If Not RAddin.myMsgBox("Error occured when invoking script '" + fullScriptPath + "\" + script + "', using '" + rexec + "'" + ex.Message + vbCrLf) Then Return False
            End Try
        Next
        Return True
    End Function

    ' get Outputfiles for defined results ranges, tab separated
    ' otherwise:  "what you see is what you get"
    Public Function getResults() As Boolean
        Dim resFilename As String = vbNullString, readdir As String
        Dim RDataRange As Range = Nothing
        Dim errMsg As String = vbNullString

        readdir = dirglobal
        For c As Integer = 0 To RdefDic("results").Length - 1
            errMsg = prepareParams(c, "results", RDataRange, resFilename, readdir, ".txt")
            If Len(errMsg) > 0 Then
                If Not RAddin.myMsgBox(errMsg) Then Return False
            End If

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\readdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            If Not File.Exists(curWbPrefix + readdir + "\" + resFilename) Then
                If Not RAddin.myMsgBox("Results file '" + curWbPrefix + readdir + "\" + resFilename + "' not found!") Then Return False
            End If
            Dim newQueryTable As QueryTable
            newQueryTable = RDataRange.Worksheet.QueryTables.Add(Connection:="TEXT;" & curWbPrefix + readdir + "\" + resFilename, Destination:=RDataRange)
            With newQueryTable
                .Name = "Data"
                .FieldNames = True
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .RefreshStyle = XlCellInsertionMode.xlOverwriteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .TextFilePlatform = 850
                .TextFileStartRow = 1
                .TextFileParseType = XlTextParsingType.xlDelimited
                .TextFileTextQualifier = XlTextQualifier.xlTextQualifierDoubleQuote
                .TextFileTabDelimiter = True
                .TextFileTrailingMinusNumbers = True
                .Refresh(BackgroundQuery:=False)
            End With
            newQueryTable.Delete()
        Next
        Return True
    End Function

    ' get Output diagrams (png) for defined diags ranges
    Public Function getDiags() As Boolean
        Dim diagFilename As String = vbNullString, readdir As String
        Dim RDataRange As Range = Nothing
        Dim errMsg As String = vbNullString

        readdir = dirglobal
        For c As Integer = 0 To RdefDic("diags").Length - 1
            errMsg = prepareParams(c, "diags", RDataRange, diagFilename, readdir, ".png")
            If Len(errMsg) > 0 Then
                If Not RAddin.myMsgBox(errMsg) Then Return False
            End If
            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\readdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            If Not File.Exists(curWbPrefix + readdir + "\" + diagFilename) Then
                If Not RAddin.myMsgBox("Diagram file '" + curWbPrefix + readdir + "\" + diagFilename + "' not found!") Then Return False
            End If

            ' clean previously set shapes...
            For Each oldShape As Shape In RDataRange.Worksheet.Shapes
                If oldShape.Name = diagFilename Then
                    oldShape.Delete()
                    Exit For
                End If
            Next
            ' add new shape from picture
            Try
                With RDataRange.Worksheet.Shapes.AddPicture(Filename:=curWbPrefix + readdir + "\" + diagFilename,
                    LinkToFile:=False, SaveWithDocument:=True, Left:=RDataRange.Left, Top:=RDataRange.Top, Width:=-1, Height:=-1)
                    .Name = diagFilename
                End With
            Catch ex As Exception
                If Not RAddin.myMsgBox("Error occured when placing the diagram into target range '" + RdefDic("diags")(c) + "', " + ex.Message) Then Return False
            End Try
        Next
        Return True
    End Function

    ' remove result, diagram and temporary R script files
    Public Function removeFiles() As Boolean
        Dim filename As String = vbNullString, readdir As String
        Dim RDataRange As Range = Nothing
        Dim errMsg As String = vbNullString

        readdir = dirglobal
        ' remove result files
        For c As Integer = 0 To RdefDic("results").Length - 1
            errMsg = prepareParams(c, "results", RDataRange, filename, readdir, ".txt")
            If Len(errMsg) > 0 Then
                If Not RAddin.myMsgBox(errMsg) Then Return False
            End If

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\argdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            ' remove any existing result files...
            If File.Exists(curWbPrefix + readdir + "\" + filename) Then
                Try
                    File.Delete(curWbPrefix + readdir + "\" + filename)
                Catch ex As Exception
                    If Not RAddin.myMsgBox("Error occured when trying to remove '" + curWbPrefix + readdir + "\" + filename + "', " + ex.Message) Then Return False
                End Try
            End If
        Next
        ' remove diagram files
        For c As Integer = 0 To RdefDic("diags").Length - 1
            errMsg = prepareParams(c, "diags", RDataRange, filename, readdir, ".png")
            If Len(errMsg) > 0 Then
                If Not RAddin.myMsgBox(errMsg) Then Return False
            End If

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\argdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            ' remove any existing diagram files...
            If File.Exists(curWbPrefix + readdir + "\" + filename) Then
                Try
                    File.Delete(curWbPrefix + readdir + "\" + filename)
                Catch ex As Exception
                    If Not RAddin.myMsgBox("Error occured when trying to remove '" + curWbPrefix + readdir + "\" + filename + "', " + ex.Message) Then Return False
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
                    If Not RAddin.myMsgBox("Error occured when trying to remove '" + curWbPrefix + readdir + "\" + filename + "', " + ex.Message) Then Return False
                End Try
            End If
        Next
        Return True
    End Function

End Module

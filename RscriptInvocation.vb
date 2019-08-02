Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop.Excel

''' <summary>all functions for the Rscript invocation method (by writing files and retrieving the results from files after invocation)</summary>
Module RscriptInvocation
    ''' <summary>executable name for calling (not only R) scripts (could also be cmd.exe or perl.exe)</summary>
    Public rexec As String
    ''' <summary>optional arguments to executable for calling (not only R) scripts, could be /c for cmd.exe</summary>
    Public rexecArgs As String

    ''' <summary>prepare parameter (script, args, results, diags) for usage in invokeScripts, storeArgs, getResults and getDiags</summary>
    ''' <param name="index">index of parameter to be prepared in RdefDic</param>
    ''' <param name="name">name (type) of parameter: script, scriptrng, args, results, diags</param>
    ''' <param name="RDataRange">returned Range of data area for scriptrng, args, results and diags</param>
    ''' <param name="returnName">returned name of data file for the parameter: same as range name</param>
    ''' <param name="returnPath">returned path of data file for the parameter</param>
    ''' <param name="ext">extension of filename that should be used for file containing data for that type (e.g. txt for args/results or png for diags)</param>
    ''' <returns>True if success, False otherwise</returns>
    Private Function prepareParam(index As Integer, name As String, ByRef RDataRange As Range, ByRef returnName As String, ByRef returnPath As String, ext As String) As String
        Dim value As String = RdefDic(name)(index)
        If value = "" Then
            Return "Empty definition value for parameter " + name + ", index: " + index.ToString()
        End If
        ' only for args, results and diags (scripts dont have a target range)
        If name = "args" Or name = "results" Or name = "diags" Or name = "scriptrng" Then
            Try
                RDataRange = currWb.Names.Item(value).RefersToRange
            Catch ex As Exception
                Return "Error occured when looking up " + name + " range '" + value + "' in Workbook " + currWb.Name + " (defined correctly ?), " + ex.Message
            End Try
        End If
        ' if argvalue refers to a WS Name, cut off WS name prefix for R file name...
        Dim posWSseparator = InStr(value, "!")
        If posWSseparator > 0 Then
            value = value.Substring(posWSseparator)
        End If
        ' get path of data file, if it is defined
        If RdefDic.ContainsKey(name + "paths") Then
            If Len(RdefDic(name + "paths")(index)) > 0 Then
                returnPath = RdefDic(name + "paths")(index)
            End If
        End If
        returnName = value + ext
        Return vbNullString
    End Function

    ''' <summary>creates Inputfiles for defined arg ranges, tab separated, decimalpoint always ".", dates are stored as "yyyy-MM-dd"
    ''' otherwise: "what you see is what you get"
    '''</summary>
    ''' <returns>True if success, False otherwise</returns>
    Public Function storeArgs() As Boolean
        Dim argFilename As String = vbNullString, argdir As String
        Dim RDataRange As Range = Nothing
        Dim outputFile As StreamWriter = Nothing

        argdir = dirglobal
        For c As Integer = 0 To RdefDic("args").Length - 1
            Try
                Dim errMsg As String
                errMsg = prepareParam(c, "args", RDataRange, argFilename, argdir, ".txt")
                If Len(errMsg) > 0 Then
                    If Not RAddin.myMsgBox(errMsg) Then Return False
                End If

                ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\argdir
                Dim curWbPrefix As String = IIf(Left(argdir, 2) = "\\" Or Mid(argdir, 2, 2) = ":\", "", currWb.Path + "\")
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
                            writtenLine += printedValue + vbTab
                            j += +1
                        Loop Until j > RDataRange.Columns.Count
                        outputFile.WriteLine(Left(writtenLine, Len(writtenLine) - 1))
                    End If
                    i += 1
                Loop Until i > RDataRange.Rows.Count
            Catch ex As Exception
                If outputFile IsNot Nothing Then outputFile.Close()
                If Not RAddin.myMsgBox("Error occured when creating inputfile '" + argFilename + "', " + ex.Message + " (maybe defined the wrong cell format for values?)") Then Return False
            End Try
            If outputFile IsNot Nothing Then outputFile.Close()
        Next
        Return True
    End Function

    ''' <summary>creates script files for defined scriptRng ranges</summary>
    ''' <returns>True if success, False otherwise</returns>
    Public Function storeScriptRng() As Boolean
        Dim scriptRngFilename As String = vbNullString, scriptText = vbNullString
        Dim RDataRange As Range = Nothing
        Dim outputFile As StreamWriter = Nothing

        Dim scriptRngdir As String = dirglobal
        For c As Integer = 0 To RdefDic("scriptrng").Length - 1
            Try
                Dim ErrMsg As String
                ' scriptrng beginning with a "=" is a scriptcell (as defined in getRDefinitions) ...
                If Left(RdefDic("scriptrng")(c), 1) = "=" Then
                    scriptText = RdefDic("scriptrng")(c).Substring(1)
                    scriptRngFilename = "RDataRangeRow" + c.ToString() + ".R"
                Else
                    ErrMsg = prepareParam(c, "scriptrng", RDataRange, scriptRngFilename, scriptRngdir, ".R")
                    If Len(ErrMsg) > 0 Then
                        If Not RAddin.myMsgBox(ErrMsg) Then Return False
                    End If
                End If


                ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\scriptRngdir
                Dim curWbPrefix As String = IIf(Left(scriptRngdir, 2) = "\\" Or Mid(scriptRngdir, 2, 2) = ":\", "", currWb.Path + "\")
                outputFile = New StreamWriter(curWbPrefix + scriptRngdir + "\" + scriptRngFilename, False, Encoding.Default)

                ' reuse the script invocation methods by setting the respective parameters
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
                                writtenLine += RDataRange(i, j).Value2
                                j += 1
                            Loop Until j > RDataRange.Columns.Count
                            outputFile.WriteLine(writtenLine)
                        End If
                        i += 1
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

    ''' <summary>invokes current scripts/args/results definition</summary>
    ''' <returns>True if success, False otherwise</returns>
    Public Function invokeScripts() As Boolean
        Dim script As String = vbNullString
        Dim scriptpath As String
        Dim previousDir As String = Directory.GetCurrentDirectory()

        scriptpath = dirglobal
        For c As Integer = 0 To RdefDic("scripts").Length - 1
            Dim ErrMsg As String = prepareParam(c, "scripts", Nothing, script, scriptpath, "")
            If Len(ErrMsg) > 0 Then
                If Not RAddin.myMsgBox(ErrMsg) Then Return False
            End If

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\scriptpath
            Dim curWbPrefix As String = IIf(Left(scriptpath, 2) = "\\" Or Mid(scriptpath, 2, 2) = ":\", "", currWb.Path + "\")
            Dim fullScriptPath = curWbPrefix + scriptpath

            Try ' + """"
                Directory.SetCurrentDirectory(fullScriptPath)
                Shell(IIf(RAddin.debugScript, "cmd.exe /c """, "") + RscriptInvocation.rexec + " " + RscriptInvocation.rexecArgs + " """ + fullScriptPath + "\" + script + """" + IIf(RAddin.debugScript, """ & pause", ""), AppWinStyle.NormalFocus, True)
            Catch ex As Exception
                ' reset current dir
                Directory.SetCurrentDirectory(previousDir)
                If Not RAddin.myMsgBox("Error occured when invoking script '" + fullScriptPath + "\" + script + "', using '" + rexec + "'" + ex.Message + vbCrLf) Then Return False
            End Try
        Next
        ' reset current dir
        Directory.SetCurrentDirectory(previousDir)
        Return True
    End Function

    ''' <summary>get Outputfiles for defined results ranges, tab separated
    ''' otherwise:  "what you see is what you get"
    ''' </summary>
    ''' <returns>True if success, False otherwise</returns>
    Public Function getResults() As Boolean
        Dim resFilename As String = vbNullString, readdir As String
        Dim RDataRange As Range = Nothing
        Dim previousResultRange As Range
        Dim errMsg As String

        readdir = dirglobal
        For c As Integer = 0 To RdefDic("results").Length - 1
            errMsg = prepareParam(c, "results", RDataRange, resFilename, readdir, ".txt")
            If Len(errMsg) > 0 Then
                If Not RAddin.myMsgBox(errMsg) Then Return False
            End If

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\readdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            If Not File.Exists(curWbPrefix + readdir + "\" + resFilename) Then
                If Not RAddin.myMsgBox("Results file '" + curWbPrefix + readdir + "\" + resFilename + "' not found!") Then Return False
            End If
            ' remove previous content, might not exist, so catch any exception
            If RdefDic("rresults")(c) Then
                Try
                    previousResultRange = currWb.Names.Item("___RaddinResult" + RdefDic("results")(c)).RefersToRange
                    previousResultRange.ClearContents()
                Catch ex As Exception
                End Try
            Else ' if we changed from rresults to results, need to remove hiddent ___RaddinResult name, otherwise results would still be removed when saving
                Try
                    currWb.Names.Item("___RaddinResult" + RdefDic("results")(c)).Delete()
                Catch ex As Exception
                End Try
            End If

            Try
                Dim newQueryTable As QueryTable
                newQueryTable = RDataRange.Worksheet.QueryTables.Add(Connection:="TEXT;" & curWbPrefix + readdir + "\" + resFilename, Destination:=RDataRange)
                '                    .TextFilePlatform = 850
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
                    .AdjustColumnWidth = False
                    .RefreshPeriod = 0
                    .TextFileStartRow = 1
                    .TextFileParseType = XlTextParsingType.xlDelimited
                    .TextFileTabDelimiter = True
                    .TextFileSpaceDelimiter = False
                    .TextFileSemicolonDelimiter = False
                    .TextFileCommaDelimiter = False
                    .TextFileDecimalSeparator = "."
                    .TextFileThousandsSeparator = ","
                    .TextFileTrailingMinusNumbers = True
                    .Refresh(BackgroundQuery:=False)
                End With
                If RdefDic("rresults")(c) Then
                    currWb.Names.Add(Name:="___RaddinResult" + RdefDic("results")(c), RefersTo:=newQueryTable.ResultRange, Visible:=False)
                End If
                newQueryTable.Delete()
            Catch ex As Exception
                If Not RAddin.myMsgBox("Error in placing results in to Excel: " + ex.Message) Then Return False
            End Try
        Next
        Return True
    End Function

    ''' <summary>get Output diagrams (png) for defined diags ranges</summary>
    ''' <returns>True if success, False otherwise</returns>
    Public Function getDiags() As Boolean
        Dim diagFilename As String = vbNullString, readdir As String
        Dim RDataRange As Range = Nothing
        Dim errMsg As String

        readdir = dirglobal
        For c As Integer = 0 To RdefDic("diags").Length - 1
            errMsg = prepareParam(c, "diags", RDataRange, diagFilename, readdir, ".png")
            If Len(errMsg) > 0 Then
                If Not RAddin.myMsgBox(errMsg) Then Return False
            End If
            ' clean previously set shape...
            For Each oldShape As Shape In RDataRange.Worksheet.Shapes
                If oldShape.Name = diagFilename Then
                    oldShape.Delete()
                    Exit For
                End If
            Next
            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\readdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            If Not File.Exists(curWbPrefix + readdir + "\" + diagFilename) Then
                If Not RAddin.myMsgBox("Diagram file '" + curWbPrefix + readdir + "\" + diagFilename + "' not found!") Then Return False
            End If

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

    ''' <summary>remove result, diagram and temporary R script files</summary>
    ''' <returns>True if success, False otherwise</returns>
    Public Function removeFiles() As Boolean
        Dim filename As String = vbNullString
        Dim readdir As String = dirglobal
        Dim RDataRange As Range = Nothing
        Dim errMsg As String

        ' check for script existence before creating any potential missing folders below...
        For c As Integer = 0 To RdefDic("scripts").Length - 1
            Dim script As String = vbNullString
            ' returns script and readdir !
            errMsg = prepareParam(c, "scripts", Nothing, script, readdir, "")
            If Len(errMsg) > 0 Then
                If Not RAddin.myMsgBox(errMsg) Then Return False
            End If

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\scriptpath
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            Dim fullScriptPath = curWbPrefix + readdir
            If Not File.Exists(fullScriptPath + "\" + script) Then
                RAddin.myMsgBox("Script '" + fullScriptPath + "\" + script + "' not found!" + vbCrLf)
                Return False
            End If
            ' check if executable exists or exists somewhere in the path....
            Dim foundExe As Boolean = False
            Dim exe As String = Environment.ExpandEnvironmentVariables(rexec)
            If Not File.Exists(exe) Then
                If Path.GetDirectoryName(exe) = String.Empty Then
                    For Each test In (Environment.GetEnvironmentVariable("PATH")).Split(";")
                        Dim thePath As String = test.Trim()
                        If Len(thePath) > 0 And File.Exists(Path.Combine(thePath, exe)) Then
                            foundExe = True
                            Exit For
                        End If
                    Next
                    If Not foundExe Then
                        RAddin.myMsgBox("Executable '" + rexec + "' not found!" + vbCrLf)
                        Return False
                    End If
                End If
            End If
        Next

        ' remove input argument files
        For c As Integer = 0 To RdefDic("args").Length - 1
            ' returns filename and readdir !
            errMsg = prepareParam(c, "args", RDataRange, filename, readdir, ".txt")
            If Len(errMsg) > 0 Then
                If Not RAddin.myMsgBox(errMsg) Then Return False
            End If

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\argdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            ' special comfort: if containing folder doesn't exist, create it now:
            If Not Directory.Exists(curWbPrefix + readdir) Then
                Try
                    Directory.CreateDirectory(curWbPrefix + readdir)
                Catch ex As Exception
                    If Not RAddin.myMsgBox("Error occured when trying to create input arguments containing folder '" + curWbPrefix + readdir + "', " + ex.Message) Then Return False
                End Try
            End If
            ' remove any existing input files...
            If File.Exists(curWbPrefix + readdir + "\" + filename) Then
                File.Delete(curWbPrefix + readdir + "\" + filename)
            End If
        Next

        ' remove result files
        For c As Integer = 0 To RdefDic("results").Length - 1
            ' returns filename and readdir !
            errMsg = prepareParam(c, "results", RDataRange, filename, readdir, ".txt")
            If Len(errMsg) > 0 Then
                If Not RAddin.myMsgBox(errMsg) Then Return False
            End If

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\argdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            ' special comfort: if containing folder doesn't exist, create it now:
            If Not Directory.Exists(curWbPrefix + readdir) Then
                Try
                    Directory.CreateDirectory(curWbPrefix + readdir)
                Catch ex As Exception
                    If Not RAddin.myMsgBox("Error occured when trying to create result containing folder '" + curWbPrefix + readdir + "', " + ex.Message) Then Return False
                End Try
            End If
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
            ' returns filename and readdir !
            errMsg = prepareParam(c, "diags", RDataRange, filename, readdir, ".png")
            If Len(errMsg) > 0 Then
                If Not RAddin.myMsgBox(errMsg) Then Return False
            End If
            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\argdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            ' special comfort: if containing folder doesn't exist, create it now:
            If Not Directory.Exists(curWbPrefix + readdir) Then
                Try
                    Directory.CreateDirectory(curWbPrefix + readdir)
                Catch ex As Exception
                    If Not RAddin.myMsgBox("Error occured when trying to create diagram containing folder '" + curWbPrefix + readdir + "', " + ex.Message) Then Return False
                End Try
            End If
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
            ' returns filename and readdir !
            errMsg = prepareParam(c, "scriptrng", RDataRange, filename, readdir, ".R")
            If Len(errMsg) > 0 Then
                filename = "RDataRangeRow" + c.ToString() + ".R"
            End If

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\argdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            ' special comfort: if containing folder doesn't exist, create it now:
            If Not Directory.Exists(curWbPrefix + readdir) Then
                Try
                    Directory.CreateDirectory(curWbPrefix + readdir)
                Catch ex As Exception
                    If Not RAddin.myMsgBox("Error occured when trying to create script containing folder '" + curWbPrefix + readdir + "', " + ex.Message) Then Return False
                End Try
            End If
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

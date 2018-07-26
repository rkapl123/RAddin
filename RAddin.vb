﻿Imports Microsoft.Office.Interop.Excel
Imports System.Configuration

Public Module RAddin

    Public currWb As Workbook
    Public Rdefinitions As Range
    Public Rcalldefnames As String() = {}
    Public Rcalldefs As Range() = {}
    Public rdefsheetColl As Dictionary(Of String, Dictionary(Of String, Range))
    Public rdefsheetMap As Dictionary(Of String, String)
    Public theRibbon As ExcelDna.Integration.CustomUI.IRibbonUI
    Public avoidFurtherMsgBoxes As Boolean
    Public dirglobal As String
    Public debugScript As Boolean
    Public dropDownSelected As Boolean

    ' definitions of current R invocation (scripts, args, results, diags...)
    Public RdefDic As Dictionary(Of String, String()) = New Dictionary(Of String, String())

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' startRprocess: started from GUI (button) and accessible from VBA (via Application.Run)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function startRprocess(runShell As Boolean, runRdotNet As Boolean) As String
        Dim errStr As String
        avoidFurtherMsgBoxes = False
        ' get the definition range
        errStr = getRDefinitions()
        If errStr <> vbNullString Then Return "Failed getting Rdefinitions: " + errStr
        If runShell Then ' Shell invocation
            Try
                If Not RscriptInvocation.removeFiles() Then Return vbNullString
                If Not RscriptInvocation.storeArgs() Then Return vbNullString
                If Not RscriptInvocation.storeScriptRng() Then Return vbNullString
                If Not RscriptInvocation.invokeScripts() Then Return vbNullString
                If Not RscriptInvocation.getResults() Then Return vbNullString
                If Not RscriptInvocation.getDiags() Then Return vbNullString
            Catch ex As Exception
                Return "Exception in Shell Rdefinitions run: " + ex.Message + ex.StackTrace
            End Try
        End If
        If runRdotNet Then ' RDotNet invocation
            Try
                If Not RdotnetInvocation.initializeRDotNet() Then Return vbNullString
                If Not RdotnetInvocation.storeArgs() Then Return vbNullString
                If Not RdotnetInvocation.invokeExcelScripts() Then Return vbNullString
                If Not RdotnetInvocation.invokeFileSysScripts() Then Return vbNullString
                If Not RdotnetInvocation.getResults() Then Return vbNullString
                If Not RdotnetInvocation.getDiags() Then Return vbNullString
            Catch ex As Exception
                Return "Exception in RdotNet Rdefinitions run: " + ex.Message + ex.StackTrace
            End Try
        End If
        ' all is OK = return nullstring
        Return vbNullString
    End Function

    ' Msgbox that avoids further Msgboxes (click Yes) or cancels run altogether (click Cancel)
    Public Function myMsgBox(message As String) As Boolean
        If avoidFurtherMsgBoxes Then Return True
        Dim retval As MsgBoxResult = MsgBox(message + vbCrLf + "Avoid further Messages (Yes/No) or abort Rdefinition (Cancel)", MsgBoxStyle.YesNoCancel)
        If retval = MsgBoxResult.Yes Then avoidFurtherMsgBoxes = True
        Return (retval = MsgBoxResult.Yes Or retval = MsgBoxResult.No)
    End Function

    ' refresh Rnames from Workbook on demand (currently when invoking about box)
    Public Function startRnamesRefresh() As String
        Dim errStr As String
        If currWb Is Nothing Then Return "No Workbook active to refresh RNames from..."
        ' always reset Rdefinitions when changing Workbooks, otherwise this is not being refilled in getRNames
        Rdefinitions = Nothing
        ' get the defined R_Addin Names
        errStr = getRNames()
        If errStr = "no definitions" Then
            Return vbNullString
        ElseIf errStr <> vbNullString Then
            Return "Error while getRNames in startRnamesRefresh: " + errStr
        End If
        theRibbon.Invalidate()
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

                Dim scriptColl As Dictionary(Of String, Range)
                If Not rdefsheetColl.ContainsKey(namedrange.Parent.Name) Then
                    ' add to new sheet "menu"
                    scriptColl = New Dictionary(Of String, Range)
                    scriptColl.Add(nodeName, namedrange.RefersToRange)
                    rdefsheetColl.Add(namedrange.Parent.Name, scriptColl)
                    rdefsheetMap.Add("ID" + i.ToString(), namedrange.Parent.Name)
                    i = i + 1
                Else
                    ' add rdefinition to existing sheet "menu"
                    scriptColl = rdefsheetColl(namedrange.Parent.Name)
                    scriptColl.Add(nodeName, namedrange.RefersToRange)
                End If
            End If
        Next
        If UBound(Rcalldefnames) = -1 Then Return "no RNames"
        Return vbNullString
    End Function

    Public Sub resetRDefinitions()
        RdefDic("args") = {}
        RdefDic("argsrc") = {}
        RdefDic("argspaths") = {}
        RdefDic("results") = {}
        RdefDic("rresults") = {}
        RdefDic("resultspaths") = {}
        RdefDic("diags") = {}
        RdefDic("diagspaths") = {}
        RdefDic("scripts") = {}
        RdefDic("scriptspaths") = {}
        RdefDic("scriptrng") = {}
        RdefDic("scriptrngpaths") = {}
        rPath = Nothing : rexec = Nothing : dirglobal = vbNullString
    End Sub

    ' gets definitions from  current selected R script invocation range (Rdefinitions)
    Public Function getRDefinitions() As String
        resetRDefinitions()
        Try
            For Each defRow As Range In Rdefinitions.Rows
                Dim deftype As String, defval As String, deffilepath As String
                deftype = LCase(defRow.Cells(1, 1).Value2)
                defval = defRow.Cells(1, 2).Value2
                deffilepath = defRow.Cells(1, 3).Value2
                If deftype = "rexec" Then ' setting for shell innvocation
                    RscriptInvocation.rexec = defval
                ElseIf deftype = "rpath" Then ' setting for RdotNet
                    RdotnetInvocation.rPath = defval
                ElseIf deftype = "arg" Or deftype = "argr" Or deftype = "argc" Or deftype = "argrc" Or deftype = "argcr" Then
                    ReDim Preserve RdefDic("argsrc")(RdefDic("argsrc").Length)
                    RdefDic("argsrc")(RdefDic("argsrc").Length - 1) = Replace(deftype, "arg", "")
                    ReDim Preserve RdefDic("args")(RdefDic("args").Length)
                    RdefDic("args")(RdefDic("args").Length - 1) = defval
                    ReDim Preserve RdefDic("argspaths")(RdefDic("argspaths").Length)
                    RdefDic("argspaths")(RdefDic("argspaths").Length - 1) = deffilepath
                ElseIf deftype = "scriptrng" Or deftype = "scriptcell" Then
                    ReDim Preserve RdefDic("scriptrng")(RdefDic("scriptrng").Length)
                    RdefDic("scriptrng")(RdefDic("scriptrng").Length - 1) = IIf(Right(deftype, 4) = "cell", "=", "") + defval
                    ReDim Preserve RdefDic("scriptrngpaths")(RdefDic("scriptrngpaths").Length)
                    RdefDic("scriptrngpaths")(RdefDic("scriptrngpaths").Length - 1) = deffilepath
                ElseIf deftype = "res" Or deftype = "rres" Then
                    ReDim Preserve RdefDic("rresults")(RdefDic("rresults").Length)
                    RdefDic("rresults")(RdefDic("rresults").Length - 1) = (deftype = "rres")
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
            ' get default rexec path from user (or overriden in appSettings tag as redirect to global) settings. This can be overruled by individual rexec settings in Rdefinitions
            Try
                If RscriptInvocation.rexec Is Nothing Then RscriptInvocation.rexec = ConfigurationManager.AppSettings("ExePath")
            Catch ex As Exception
                Return "Error in getRDefinitions: " + ex.Message
            End Try
            ' get default rHome from user (or overriden in appSettings tag as redirect to global) settings. This can be overruled by individual rexec settings in Rdefinitions
            Try
                RdotnetInvocation.rHome = ConfigurationManager.AppSettings("rHome")
                If RdotnetInvocation.rPath Is Nothing Then RdotnetInvocation.rPath = RdotnetInvocation.rHome + IIf(RdotnetInvocation.rHome.EndsWith("\"), "", "\") + IIf(System.Environment.Is64BitProcess, ConfigurationManager.AppSettings("rPath64bit"), ConfigurationManager.AppSettings("rPath32bit"))
            Catch ex As Exception
                Return "Error in getRDefinitions: " + ex.Message
            End Try
            If rexec Is Nothing And rPath Is Nothing Then Return "Error in getRDefinitions: neither rexec nor rpath (for Rdotnet) defined"
            If RdefDic("scripts").Count = 0 And RdefDic("scriptrng").Count = 0 Then Return "Error in getRDefinitions: no script(s) or scriptRng(s) defined in " + Rdefinitions.Name.Name
        Catch ex As Exception
            Return "Error in getRDefinitions: " + ex.Message
        End Try
        Return vbNullString
    End Function

    ' remove results in all result Ranges (before saving)
    Public Function removeResultsDiags() As Boolean
        For Each namedrange As Name In currWb.Names
            If Left(namedrange.Name, 15) = "___RaddinResult" Then
                namedrange.RefersToRange.Clear()
                namedrange.Delete()
            End If
        Next
        Return True
    End Function

End Module

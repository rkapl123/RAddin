Imports Microsoft.Office.Interop.Excel
Imports System.Configuration

''' <summary>The main functions for working with RDefinitions (named ranges in Excel) and starting the R processes (writing input, invoking R scripts and retrievng results)</summary>
Public Module RAddin

    ''' <summary></summary>
    Public currWb As Workbook
    ''' <summary></summary>
    Public Rdefinitions As Range
    ''' <summary></summary>
    Public Rcalldefnames As String() = {}
    ''' <summary></summary>
    Public Rcalldefs As Range() = {}
    ''' <summary></summary>
    Public rdefsheetColl As Dictionary(Of String, Dictionary(Of String, Range))
    ''' <summary></summary>
    Public rdefsheetMap As Dictionary(Of String, String)
    ''' <summary>reference object for the Addins ribbon</summary>
    Public theRibbon As ExcelDna.Integration.CustomUI.IRibbonUI
    ''' <summary>ribbon menu handler</summary>
    Public theMenuHandler As MenuHandler
    ''' <summary></summary>
    Public avoidFurtherMsgBoxes As Boolean
    ''' <summary></summary>
    Public dirglobal As String
    ''' <summary></summary>
    Public debugScript As Boolean
    ''' <summary></summary>
    Public dropDownSelected As Boolean

    ''' <summary>definitions of current R invocation (scripts, args, results, diags...)</summary>
    Public RdefDic As Dictionary(Of String, String()) = New Dictionary(Of String, String())

    ''' <summary>startRprocess: started from GUI (button) and accessible from VBA (via Application.Run)</summary>
    ''' <param name="runShell"></param>
    ''' <param name="runRdotNet"></param>
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

    ''' <summary>Msgbox that avoids further Msgboxes (click Yes) or cancels run altogether (click Cancel)</summary>
    ''' <param name="message"></param>
    ''' <returns>True if further Msgboxes should be avoided, False otherwise</returns>
    Public Function myMsgBox(message As String, Optional noAvoidChoice As Boolean = False) As Boolean
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Dim caller As String = theMethod.ReflectedType.FullName & "." & theMethod.Name

        Trace.TraceWarning("{0}: {1}", caller, message)
        If noAvoidChoice Then
            MsgBox(message, MsgBoxStyle.OkOnly, "R-Addin Message")
            Return False
        Else
            If avoidFurtherMsgBoxes Then Return True
            Dim retval As MsgBoxResult = MsgBox(message + vbCrLf + "Avoid further Messages (Yes/No) or abort Rdefinition (Cancel)", MsgBoxStyle.YesNoCancel, "R-Addin Message")
            If retval = MsgBoxResult.Yes Then avoidFurtherMsgBoxes = True
            Return (retval = MsgBoxResult.Yes Or retval = MsgBoxResult.No)
        End If
    End Function

    ''' <summary>refresh Rnames from Workbook on demand (currently when invoking about box)</summary>
    ''' <returns></returns>
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

    ''' <summary>gets defined named ranges for R script invocation in the current workbook</summary>
    ''' <returns></returns>
    Public Function getRNames() As String
        ReDim Preserve Rcalldefnames(-1)
        ReDim Preserve Rcalldefs(-1)
        rdefsheetColl = New Dictionary(Of String, Dictionary(Of String, Range))
        rdefsheetMap = New Dictionary(Of String, String)
        Dim i As Integer = 0
        For Each namedrange As Name In currWb.Names
            Dim cleanname As String = Replace(namedrange.Name, namedrange.Parent.Name & "!", "")
            If Left(cleanname, 7) = "R_Addin" Then
                If InStr(namedrange.RefersTo, "#REF!") > 0 Then Return "Rdefinitions range " + namedrange.Parent.name + "!" + namedrange.Name + " contains #REF!"
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
                    i += 1
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

    ''' <summary>reset all RDefinition representations</summary>
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

    ''' <summary>gets definitions from current selected R script invocation range (Rdefinitions)</summary>
    ''' <returns></returns>
    Public Function getRDefinitions() As String
        resetRDefinitions()
        Try
            RscriptInvocation.rexecArgs = "" ' reset (r)exec arguments as they might have been set elsewhere...
            For Each defRow As Range In Rdefinitions.Rows
                Dim deftype As String, defval As String, deffilepath As String
                deftype = LCase(defRow.Cells(1, 1).Value2)
                defval = defRow.Cells(1, 2).Value2
                deffilepath = defRow.Cells(1, 3).Value2
                If deftype = "rexec" Then ' setting for shell innvocation
                    RscriptInvocation.rexec = defval
                    RscriptInvocation.rexecArgs = deffilepath
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

    ''' <summary>remove results in all result Ranges (before saving)</summary>
    Public Sub removeResultsDiags()
        For Each namedrange As Name In currWb.Names
            If Left(namedrange.Name, 15) = "___RaddinResult" Then
                namedrange.RefersToRange.ClearContents()
                namedrange.Delete()
            End If
        Next
    End Sub

End Module

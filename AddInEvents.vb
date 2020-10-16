Imports ExcelDna.Integration
Imports Microsoft.Office.Interop.Excel

''' <summary>Events from Excel (Workbook_Save ...)</summary>
Public Class AddInEvents
    Implements IExcelAddIn

    ''' <summary>the Application object for event registration</summary>
    WithEvents Application As Application

    ''' <summary>connect to Excel when opening Addin</summary>
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        Application = ExcelDnaUtil.Application
        Dim Wb As Workbook = Application.ActiveWorkbook
        If Not Wb Is Nothing Then
            Dim errStr As String = doDefinitions(Wb)
            RAddin.dropDownSelected = False
            If errStr = "no RNames" Then
                RAddin.resetRDefinitions()
            ElseIf errStr <> vbNullString Then
                myMsgBox("Error when getting definitions in Workbook_Activate: " + errStr, True)
            End If
        End If
        theMenuHandler = New MenuHandler
        Trace.Listeners.Add(New ExcelDna.Logging.LogDisplayTraceListener())
    End Sub

    ''' <summary>clean up when Raddin is deactivated</summary>
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        If RdotnetInvocation.rDotNetEngine IsNot Nothing Then RdotnetInvocation.rDotNetEngine.Dispose()
        theMenuHandler = Nothing
        Application = Nothing
    End Sub

    ''' <summary>save arg ranges to text files as well </summary>
    Private Sub Workbook_Save(Wb As Workbook, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
        Dim errStr As String
        ' avoid resetting Rdefinition when dropdown selected for a specific RDefinition !
        If RAddin.dropDownSelected Then
            errStr = RAddin.getRDefinitions()
            If errStr <> vbNullString Then myMsgBox("Error while getRdefinitions (dropdown selected !) in Workbook_Save: " + errStr, True)
        Else
            errStr = doDefinitions(Wb) ' includes getRDefinitions - for top sorted Rdefinition
            If errStr = "no RNames" Then Exit Sub
            If errStr <> vbNullString Then
                myMsgBox("Error when getting definitions in Workbook_Save: " + errStr, True)
                Exit Sub
            End If
        End If
        RAddin.avoidFurtherMsgBoxes = False
        RscriptInvocation.storeArgs()
        RAddin.removeResultsDiags() ' remove results specified by rres
    End Sub

    ''' <summary>refresh ribbon is being treated in Workbook_Activate...</summary>
    Private Sub Workbook_Open(Wb As Workbook) Handles Application.WorkbookOpen
    End Sub

    ''' <summary>refresh ribbon with current workbook's Rnames</summary>
    Private Sub Workbook_Activate(Wb As Workbook) Handles Application.WorkbookActivate
        Dim errStr As String = doDefinitions(Wb)
        RAddin.dropDownSelected = False
        If errStr = "no RNames" Then
            RAddin.resetRDefinitions()
        ElseIf errStr <> vbNullString Then
            myMsgBox("Error when getting definitions in Workbook_Activate: " + errStr, True)
        End If
        RAddin.theRibbon.Invalidate()
    End Sub

    ''' <summary>get Rnames of current workbook and load Rdefinitions of first name in R_Addin Names</summary>
    Private Function doDefinitions(Wb As Workbook) As String
        Dim errStr As String
        currWb = Wb
        ' always reset Rdefinitions when changing Workbooks (may not be the current one, if saved programmatically!), otherwise this is not being refilled in getRNames
        RdefinitionRange = Nothing
        ' get the defined R_Addin Names
        errStr = RAddin.getRNames()
        If errStr = "no RNames" Then Return errStr
        If errStr <> vbNullString Then
            Return "Error while getRNames in doDefinitions: " + errStr
        End If
        ' get the definitions from the current defined range (first name in R_Addin Names)
        errStr = RAddin.getRDefinitions()
        If errStr <> vbNullString Then Return "Error while getRdefinitions in doDefinitions: " + errStr
        LogInfo("done RDefinitions for workbook " + Wb.Name)
        Return vbNullString
    End Function

    ''' <summary>Close Workbook: remove references to current Workbook and R Definitions</summary>
    Private Sub Application_WorkbookDeactivate(Wb As Workbook) Handles Application.WorkbookDeactivate
        currWb = Nothing
        RAddin.dropDownSelected = False
        ReDim Preserve Rcalldefnames(-1)
        ReDim Preserve Rcalldefs(-1)
        rdefsheetColl = New Dictionary(Of String, Dictionary(Of String, Range))
        rdefsheetMap = New Dictionary(Of String, String)
        RAddin.resetRDefinitions()
        RAddin.theRibbon.Invalidate()
    End Sub
End Class

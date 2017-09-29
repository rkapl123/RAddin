Imports ExcelDna.Integration
Imports Microsoft.Office.Interop.Excel

' Events from Excel (Workbook_Save ...)
Public Class AddInEvents
    Implements IExcelAddIn

    WithEvents Application As Application

    ' connect to Excel when opening Addin
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        Application = ExcelDnaUtil.Application
    End Sub

    ' clean up when Raddin is deactivated
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        If RdotnetInvocation.rdotnetengine IsNot Nothing Then RdotnetInvocation.rdotnetengine.Dispose()
    End Sub

    ' save arg ranges to text files as well 
    Private Sub Workbook_Save(Wb As Workbook, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
        Dim errStr As String
        ' avoid resetting Rdefinition when dropdown selected for a specific RDefinition !
        If RAddin.dropDownSelected Then
            errStr = RAddin.getRDefinitions()
            If errStr <> vbNullString Then MsgBox("Error while getRdefinitions (dropdown selected !) in Workbook_Save: " + errStr)
        Else
            errStr = doDefinitions(Wb) ' includes getRDefinitions - for top sorted Rdefinition
            If errStr = "no RNames" Then Exit Sub
            If errStr <> vbNullString Then
                MsgBox("Error when getting definitions in Workbook_Save: " + errStr)
                Exit Sub
            End If
        End If
        RAddin.avoidFurtherMsgBoxes = False
        RscriptInvocation.storeArgs()
        'RAddin.removeResultsDiags()
    End Sub

    Private Sub Workbook_Open(Wb As Workbook) Handles Application.WorkbookOpen
        ' is being treated in Workbook_Activate...
    End Sub

    ' refresh ribbon with current workbook's Rnames
    Private Sub Workbook_Activate(Wb As Workbook) Handles Application.WorkbookActivate
        Dim errStr As String
        errStr = doDefinitions(Wb)
        RAddin.dropDownSelected = False
        If errStr = "no RNames" Then
            RAddin.resetRDefinitions()
        ElseIf errStr <> vbNullString Then
            MsgBox("Error when getting definitions in Workbook_Activate: " + errStr)
        End If
        RAddin.theRibbon.Invalidate()
    End Sub

    ' get Rnames of current workbook and load Rdefinitions of first name in R_Addin Names
    Private Function doDefinitions(Wb As Workbook) As String
        Dim errStr As String
        currWb = Wb
        ' always reset Rdefinitions when changing Workbooks (may not be the current one, if saved programmatically!), otherwise this is not being refilled in getRNames
        Rdefinitions = Nothing
        ' get the defined R_Addin Names
        errStr = RAddin.getRNames()
        If errStr = "no RNames" Then Return errStr
        If errStr <> vbNullString Then
            Return "Error while getRNames in doDefinitions: " + errStr
        End If
        ' get the definitions from the current defined range (first name in R_Addin Names)
        errStr = RAddin.getRDefinitions()
        If errStr <> vbNullString Then Return "Error while getRdefinitions in doDefinitions: " + errStr
        Return vbNullString
    End Function

End Class

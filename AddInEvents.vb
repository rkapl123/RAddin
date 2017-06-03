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

    'has to be implemented
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        If RAddin.rdotnetengine IsNot Nothing Then RAddin.rdotnetengine.Dispose()
    End Sub

    Private Sub Workbook_Save(Wb As Workbook, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
        Dim errStr As String
        errStr = doDefinitions(Wb)
        If errStr = "no RNames" Then Exit Sub
        If errStr <> vbNullString Then
            MsgBox("Error when getting definitions in Workbook_Save: " + errStr)
            Exit Sub
        End If
        RAddin.avoidFurtherMsgBoxes = False
        RAddin.storeArgs()
    End Sub

    Private Sub Workbook_Open(Wb As Workbook) Handles Application.WorkbookOpen
        ' is being treated in Workbook_Activate...
    End Sub

    Private Sub Workbook_Activate(Wb As Workbook) Handles Application.WorkbookActivate
        Dim errStr As String
        errStr = doDefinitions(Wb)
        If errStr = "no RNames" Then Exit Sub
        If errStr <> vbNullString Then
            MsgBox("Error when getting definitions in Workbook_Activate: " + errStr)
            Exit Sub
        End If
        RAddin.theRibbon.Invalidate()
    End Sub

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

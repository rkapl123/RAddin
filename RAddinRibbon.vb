Imports System.Runtime.InteropServices
Imports ExcelDna.Integration.CustomUI

' Events from Ribbon
<ComVisible(True)>
Public Class RAddinRibbon
    Inherits ExcelRibbon

    Private Sub startRprocess(runShell As Boolean)
        Dim errStr As String
        If UBound(Rcalldefnames) = -1 Then
            MsgBox("no Rdefinitions found for R_Addin in current Workbook (3 column named range (type/value/path), minimum types: rexec and script)!")
            Exit Sub
        End If
        If RAddin.Rdefinitions Is Nothing Then
            MsgBox("Rdefinitions Is Nothing (this shouldn't actually happen) !")
            Exit Sub
        End If
        errStr = RAddin.startRprocess(runShell)
        If errStr <> "" Then MsgBox(errStr)
    End Sub

    Public Sub startRprocessShell(control As ExcelDna.Integration.CustomUI.IRibbonControl)
        startRprocess(True)
    End Sub

    Public Sub startRprocessRdotNet(control As ExcelDna.Integration.CustomUI.IRibbonControl)
        startRprocess(False)
    End Sub

    Public Sub refreshRdefs(control As ExcelDna.Integration.CustomUI.IRibbonControl)
        Dim errStr As String
        errStr = RAddin.startRdefRefresh()
        If errStr <> vbNullString Then
            MsgBox(errStr)
        Else
            MsgBox("refreshed Rdefinitions from current Workbook !")
        End If
    End Sub

    Public Function GetItemCount(control As ExcelDna.Integration.CustomUI.IRibbonControl) As Integer
        Return (RAddin.Rcalldefnames.Length)
    End Function

    Public Function GetItemLabel(control As ExcelDna.Integration.CustomUI.IRibbonControl, index As Integer) As String
        Return RAddin.Rcalldefnames(index)
    End Function

    Public Function GetItemID(control As ExcelDna.Integration.CustomUI.IRibbonControl, index As Integer) As String
        Return RAddin.Rcalldefnames(index)
    End Function

    Public Sub selectItem(control As ExcelDna.Integration.CustomUI.IRibbonControl, id As String, index As Integer)
        Dim errStr As String
        errStr = RAddin.startRdefRefresh()
        If errStr <> vbNullString Then
            MsgBox(errStr)
        End If
        RAddin.Rdefinitions = Rcalldefs(index)
        RAddin.Rdefinitions.Parent.Select()
        RAddin.Rdefinitions.Select()
    End Sub

    Public Sub ribbonLoaded(myribbon As ExcelDna.Integration.CustomUI.IRibbonUI)
        RAddin.theRibbon = myribbon
    End Sub

End Class
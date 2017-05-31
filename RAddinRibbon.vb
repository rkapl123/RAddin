Imports System.Runtime.InteropServices
Imports ExcelDna.Integration.CustomUI

' Events from Ribbon
<ComVisible(True)>
Public Class RAddinRibbon
    Inherits ExcelRibbon

    Public runShell As Boolean

    Public Sub startRprocess(control As IRibbonControl)
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

    Public Function getToggleLabel(control As IRibbonControl) As String
        Return "run via " + IIf(runShell, "Shell", "RdotNet")
    End Function

    Public Sub toggleRunScript(control As IRibbonControl, pressed As Boolean)
        runShell = Not pressed
        RAddin.theRibbon.Invalidate()
    End Sub

    Public Sub refreshRdefs(control As IRibbonControl)
        Dim sModuleInfo As String = vbNullString
        For Each tModule As ProcessModule In Process.GetCurrentProcess().Modules
            Dim sModule As String = tModule.FileName
            If sModule.ToUpper.Contains("RADDIN-ADDIN-PACKED.XLL") Or sModule.ToUpper.Contains("RADDIN-ADDIN64-PACKED.XLL") Then
                sModuleInfo = sModule & ", Buildtime: " & FileDateTime(sModule).ToString()
            End If
        Next
        Dim errStr As String
        errStr = RAddin.startRdefRefresh()
        If errStr <> vbNullString Then
            MsgBox(sModuleInfo & vbCrLf & vbCrLf & "refresh Error: " & errStr)
        Else
            MsgBox(sModuleInfo & vbCrLf & vbCrLf & "refreshed Rdefinitions from current Workbook !")
        End If
    End Sub

    Public Function GetItemCount(control As IRibbonControl) As Integer
        Return (RAddin.Rcalldefnames.Length)
    End Function

    Public Function GetItemLabel(control As IRibbonControl, index As Integer) As String
        Return RAddin.Rcalldefnames(index)
    End Function

    Public Function GetItemID(control As IRibbonControl, index As Integer) As String
        Return RAddin.Rcalldefnames(index)
    End Function

    Public Sub selectItem(control As IRibbonControl, id As String, index As Integer)
        Dim errStr As String
        errStr = RAddin.startRdefRefresh()
        If errStr <> vbNullString Then
            MsgBox(errStr)
        End If
        RAddin.Rdefinitions = Rcalldefs(index)
        RAddin.Rdefinitions.Parent.Select()
        RAddin.Rdefinitions.Select()
    End Sub

    Public Sub ribbonLoaded(myribbon As IRibbonUI)
        RAddin.theRibbon = myribbon
        ' default to run via shell..
        runShell = True
    End Sub

    Public Overrides Function GetCustomUI(RibbonID As String) As String
        Dim customUIXml As String = "<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' loadImage='LoadImage' onLoad='ribbonLoaded' ><ribbon><tabs><tab id='RaddinTab' label='R Addin'>" + _
            "<group id='RaddinGroup' label='Run defined R-Scripts'>" + _
              "<dropDown id='scriptDropDown' label='Rdefinition:' sizeString='123456789012345678901234567890' getItemCount='GetItemCount' getItemID='GetItemID' getItemLabel='GetItemLabel' onAction='selectItem'/>" + _
              "<toggleButton id='Button1' getLabel='getToggleLabel' onAction='toggleRunScript' image='M' size='normal' tag='1' screentip='toggles whether to run R script via Shell/Files or RdotNet' supertip='toggles whether to run R script via Shell/Files or RdotNet' />" + _
              "<button id='Button2' label='run Rdefinion' image='M' size='normal' onAction='startRprocess' tag='2' screentip='run R script from dropdown' supertip='runs R script defined in corresponding range R_Addin' />" + _
              "<dialogBoxLauncher><button id='Button3' label='refresh Rdefinitions and get RAddin Info' onAction='refreshRdefs' tag='3' screentip='refresh Rdefinitions from current Workbook and get Info about RAddin' supertip='refreshes the Rdefinition: dropdown from all ranges in the current Workbook called R_Addin and gets RAddin Info (Buildtime)' /></dialogBoxLauncher>" + _
            "</group></tab></tabs></ribbon></customUI>"
        Return customUIXml
    End Function
End Class
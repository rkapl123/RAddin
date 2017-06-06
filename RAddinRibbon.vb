Imports System.Runtime.InteropServices
Imports ExcelDna.Integration.CustomUI

' Events from Ribbon
<ComVisible(True)>
Public Class RAddinRibbon
    Inherits ExcelRibbon

    Public runShell As Boolean
    Public runRdotNet As Boolean

    Public Sub startRprocess(control As IRibbonControl)
        Dim errStr As String
        ' set Rdefinition to invocaters range... invocating sheet is put into Tag
        RAddin.Rdefinitions = RAddin.rdefsheetColl(control.Tag).Item(control.Id)
        RAddin.Rdefinitions.Parent.Select()
        RAddin.Rdefinitions.Select()
        errStr = RAddin.startRprocess(runShell, runRdotNet)
        If errStr <> "" Then MsgBox(errStr)
    End Sub

    ' reflect the change in the togglebuttons title
    Public Function getImage(control As IRibbonControl) As String
        If (runShell And control.Id = "shell") Or (runRdotNet And control.Id = "rdotnet") Then
            Return "AcceptTask"
        Else
            Return "DeclineTask"
        End If
    End Function

    ' reflect the change in the togglebuttons title
    Public Function getPressed(control As IRibbonControl) As Boolean
        If control.Id = "shell" Then
            Return runShell
        ElseIf control.Id = "rdotnet" Then
            Return runRdotNet
        Else
            Return False
        End If
    End Function

    ' toggle shell or Rdotnet mode buttons
    Public Sub toggleButton(control As IRibbonControl, pressed As Boolean)
        If control.Id = "shell" Then
            runShell = pressed
        ElseIf control.Id = "rdotnet" Then
            runRdotNet = pressed
        End If
        ' invalidate to reflect the change in the togglebuttons image
        RAddin.theRibbon.InvalidateControl(control.Id)
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
        errStr = RAddin.startRnamesRefresh()
        If errStr <> vbNullString Then
            MsgBox(sModuleInfo & vbCrLf & vbCrLf & "refresh Error: " & errStr)
        Else
            If UBound(Rcalldefnames) = -1 Then
                MsgBox(sModuleInfo & vbCrLf & vbCrLf & "no Rdefinitions found for R_Addin in current Workbook (3 column named range (type/value/path), minimum types: rexec and script)!")
            Else
                MsgBox(sModuleInfo & vbCrLf & vbCrLf & "refreshed Rnames from current Workbook !")
            End If
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
        RAddin.Rdefinitions = Rcalldefs(index)
        RAddin.Rdefinitions.Parent.Select()
        RAddin.Rdefinitions.Select()
    End Sub

    Public Sub ribbonLoaded(myribbon As IRibbonUI)
        RAddin.theRibbon = myribbon
        ' set default run via methods ..
        runShell = True
        runRdotNet = False
    End Sub

    ' creates the Ribbon
    Public Overrides Function GetCustomUI(RibbonID As String) As String
        Dim customUIXml As String = "<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='ribbonLoaded' ><ribbon><tabs><tab id='RaddinTab' label='R Addin'>" +
            "<group id='RaddinGroup' label='General settings'>" +
              "<dropDown id='scriptDropDown' label='Rdefinition:' sizeString='123456789012345678901234567890' getItemCount='GetItemCount' getItemID='GetItemID' getItemLabel='GetItemLabel' onAction='selectItem'/>" +
              "<toggleButton id='shell' label='run via shell' onAction='toggleButton' getImage='getImage' getPressed='getPressed' size='normal' tag='1' screentip='toggles whether to run R script via Shell/Files' supertip='toggles whether to run R script via Shell/Files' />" +
              "<toggleButton id='rdotnet' label='run via RdotNet' onAction='toggleButton' getImage='getImage' getPressed='getPressed' size='normal' tag='2' screentip='toggles whether to run R script via RdotNet' supertip='toggles whether to run R script via RdotNet' />" +
              "<dialogBoxLauncher><button id='dialog' label='refresh Rdefinitions and get RAddin Info' onAction='refreshRdefs' tag='3' screentip='refresh Rdefinitions from current Workbook and get Info about RAddin' supertip='refreshes the Rdefinition: dropdown from all ranges in the current Workbook called R_Addin and gets RAddin Info (Buildtime)' /></dialogBoxLauncher></group>" +
              "<group id='RscriptsGroup' label='Run R-Scripts defined in WB/sheet names'>"

        For i As Integer = 0 To 10
            customUIXml = customUIXml + "<dynamicMenu id='ID" + i.ToString() + "' " +
                                            "size='large' getLabel='getSheetLabel' imageMso='SignatureLineInsert' " +
                                            "screentip='Select script to run' " +
                                            "getContent='getDynMenContent' getVisible='getDynMenVisible'/>"
        Next
        customUIXml = customUIXml + "</group></tab></tabs></ribbon></customUI>"
        Return customUIXml
    End Function

    ' set the name of the WB/sheet dropdown to the sheet name (for the WB dropdown this is the WB name) 
    Public Function getSheetLabel(control As IRibbonControl) As String
        getSheetLabel = vbNullString
        If RAddin.rdefsheetMap.ContainsKey(control.Id) Then getSheetLabel = RAddin.rdefsheetMap(control.Id)
    End Function

    ' create the buttons in the WB/sheet dropdown
    Public Function getDynMenContent(control As IRibbonControl) As String
        Dim xmlString As String = "<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'>"
        Dim currentSheet As String = RAddin.rdefsheetMap(control.Id)
        For Each nodeName As String In RAddin.rdefsheetColl(currentSheet).Keys
            xmlString = xmlString + "<button id='" + nodeName + "' label='run " + nodeName + "' imageMso='SignatureLineInsert' onAction='startRprocess' tag ='" + currentSheet + "' screentip='run " + nodeName + " Rdefinition' supertip='runs R script defined in " + nodeName + " R_Addin range on sheet " + currentSheet + "' />"
        Next
        xmlString = xmlString + "</menu>"
        Return xmlString
    End Function

    ' shows the sheet button only if it was collected...
    Public Function getDynMenVisible(control As IRibbonControl) As Boolean
        Return RAddin.rdefsheetMap.ContainsKey(control.Id)
    End Function
End Class
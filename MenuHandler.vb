Imports System.Runtime.InteropServices
Imports ExcelDna.Integration.CustomUI
Imports System.Configuration

''' <summary>Events from Ribbon</summary>
<ComVisible(True)>
Public Class MenuHandler
    Inherits ExcelRibbon
    ''' <summary></summary>
    Public runShell As Boolean
    ''' <summary></summary>
    Public runRdotNet As Boolean

    ''' <summary></summary>
    Public Sub startRprocess(control As IRibbonControl)
        Dim errStr As String
        ' set Rdefinition to invocaters range... invocating sheet is put into Tag
        RAddin.RdefinitionRange = RAddin.rdefsheetColl(control.Tag).Item(control.Id)
        RAddin.RdefinitionRange.Parent.Select()
        RAddin.RdefinitionRange.Select()
        errStr = RAddin.startRprocess(runShell, runRdotNet)
        If errStr <> "" Then myMsgBox(errStr, True)
    End Sub

    ''' <summary>reflect the change in the togglebuttons title</summary>
    ''' <returns></returns>
    Public Function getImage(control As IRibbonControl) As String
        If (runShell And control.Id = "shell") Or (runRdotNet And control.Id = "rdotnet") Or (RAddin.debugScript And control.Id = "debug") Then
            Return "AcceptTask"
        Else
            Return "DeclineTask"
        End If
    End Function

    ''' <summary>reflect the change in the togglebuttons title</summary>
    ''' <returns>True for the respective control if activated</returns>
    Public Function getPressed(control As IRibbonControl) As Boolean
        If control.Id = "shell" Then
            Return runShell
        ElseIf control.Id = "rdotnet" Then
            Return runRdotNet
        ElseIf control.Id = "debug" Then
            Return RAddin.debugScript
        Else
            Return False
        End If
    End Function

    ''' <summary>toggle shell or Rdotnet mode buttons</summary>
    ''' <param name="pressed"></param>
    Public Sub toggleButton(control As IRibbonControl, pressed As Boolean)
        If control.Id = "shell" Then
            runShell = pressed
            runRdotNet = Not pressed
        ElseIf control.Id = "rdotnet" Then
            runRdotNet = pressed
            runShell = Not pressed
        ElseIf control.Id = "debug" Then
            RAddin.debugScript = pressed
            ' invalidate to reflect the change in the togglebuttons image
            RAddin.theRibbon.InvalidateControl(control.Id)
            Exit Sub
        End If
        ' for shell/rdotnet toggle always invalidate both controls
        RAddin.theRibbon.InvalidateControl("shell")
        RAddin.theRibbon.InvalidateControl("rdotnet")
    End Sub

    ''' <summary></summary>
    Public Sub refreshRdefs(control As IRibbonControl)
        Dim myAbout As AboutBox1 = New AboutBox1
        myAbout.ShowDialog()
    End Sub

    ''' <summary></summary>
    ''' <returns></returns>
    Public Function GetItemCount(control As IRibbonControl) As Integer
        Return (RAddin.Rcalldefnames.Length)
    End Function

    ''' <summary></summary>
    ''' <returns></returns>
    Public Function GetItemLabel(control As IRibbonControl, index As Integer) As String
        Return RAddin.Rcalldefnames(index)
    End Function

    ''' <summary></summary>
    ''' <returns></returns>
    Public Function GetItemID(control As IRibbonControl, index As Integer) As String
        Return RAddin.Rcalldefnames(index)
    End Function

    ''' <summary></summary>
    Public Sub selectItem(control As IRibbonControl, id As String, index As Integer)
        ' needed for workbook save (saves selected Rdefinition)
        RAddin.dropDownSelected = True
        RAddin.RdefinitionRange = Rcalldefs(index)
        RAddin.RdefinitionRange.Parent.Select()
        RAddin.RdefinitionRange.Select()
    End Sub

    ''' <summary></summary>
    Public Sub ribbonLoaded(myribbon As IRibbonUI)
        RAddin.theRibbon = myribbon
        ' set default run via methods ..
        Try
            runShell = CBool(ConfigurationManager.AppSettings("runShell"))
            runRdotNet = CBool(ConfigurationManager.AppSettings("runRdotNet"))
        Catch ex As Exception
            myMsgBox("Error reading default run configuration runShell/runDotNet:" + ex.Message, True)
        End Try
    End Sub

    ''' <summary>creates the Ribbon</summary>
    ''' <returns></returns>
    Public Overrides Function GetCustomUI(RibbonID As String) As String
        Dim customUIXml As String = "<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='ribbonLoaded' ><ribbon><tabs><tab id='RaddinTab' label='R Addin'>" +
            "<group id='RaddinGroup' label='General settings'>" +
              "<dropDown id='scriptDropDown' label='Rdefinition:' sizeString='12345678901234567890' getItemCount='GetItemCount' getItemID='GetItemID' getItemLabel='GetItemLabel' onAction='selectItem'/>" +
              "<buttonGroup id='buttonGroup'>" +
              "<toggleButton id='shell' label='run via shell' onAction='toggleButton' getImage='getImage' getPressed='getPressed' tag='1' screentip='toggles whether to run R script via Shell/Files' supertip='toggles whether to run R script via Shell/Files' />" +
              "<toggleButton id='rdotnet' label='run via RdotNet' onAction='toggleButton' getImage='getImage' getPressed='getPressed' tag='2' screentip='toggles whether to run R script via RdotNet' supertip='toggles whether to run R script via RdotNet' />" +
              "</buttonGroup><toggleButton id='debug' label='debug script' onAction='toggleButton' getImage='getImage' getPressed='getPressed' tag='3' screentip='toggles whether to debug R script' supertip='toggles whether to debug R script (leave cmd shell open)' />" +
              "" +
              "<dialogBoxLauncher><button id='dialog' label='About RAddin' onAction='refreshRdefs' tag='3' screentip='Show Aboutbox (refresh Rdefinitions from current Workbook and show Log from there)'/></dialogBoxLauncher></group>" +
              "<group id='RscriptsGroup' label='Run R-Scripts defined in WB/sheet names'>"
        Dim presetSheetButtonsCount As Integer = Int16.Parse(ConfigurationManager.AppSettings("presetSheetButtonsCount"))
        Dim thesize As String = IIf(presetSheetButtonsCount < 15, "normal", "large")
        For i As Integer = 0 To presetSheetButtonsCount
            customUIXml = customUIXml + "<dynamicMenu id='ID" + i.ToString() + "' " +
                                            "size='" + thesize + "' getLabel='getSheetLabel' imageMso='SignatureLineInsert' " +
                                            "screentip='Select script to run' " +
                                            "getContent='getDynMenContent' getVisible='getDynMenVisible'/>"
        Next
        customUIXml += "</group></tab></tabs></ribbon></customUI>"
        Return customUIXml
    End Function

    ''' <summary>set the name of the WB/sheet dropdown to the sheet name (for the WB dropdown this is the WB name)</summary>
    ''' <returns></returns>
    Public Function getSheetLabel(control As IRibbonControl) As String
        getSheetLabel = vbNullString
        If RAddin.rdefsheetMap.ContainsKey(control.Id) Then getSheetLabel = RAddin.rdefsheetMap(control.Id)
    End Function

    ''' <summary>create the buttons in the WB/sheet dropdown</summary>
    ''' <returns></returns>
    Public Function getDynMenContent(control As IRibbonControl) As String
        Dim xmlString As String = "<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'>"
        Dim currentSheet As String = RAddin.rdefsheetMap(control.Id)
        For Each nodeName As String In RAddin.rdefsheetColl(currentSheet).Keys
            xmlString = xmlString + "<button id='" + nodeName + "' label='run " + nodeName + "' imageMso='SignatureLineInsert' onAction='startRprocess' tag ='" + currentSheet + "' screentip='run " + nodeName + " Rdefinition' supertip='runs R script defined in " + nodeName + " R_Addin range on sheet " + currentSheet + "' />"
        Next
        xmlString += "</menu>"
        Return xmlString
    End Function

    ''' <summary>shows the sheet button only if it was collected...</summary>
    ''' <returns></returns>
    Public Function getDynMenVisible(control As IRibbonControl) As Boolean
        Return RAddin.rdefsheetMap.ContainsKey(control.Id)
    End Function

End Class
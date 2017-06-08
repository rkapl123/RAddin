﻿Public NotInheritable Class AboutBox1
    Private Sub AboutBox1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Set the title of the form.
        Dim sModuleInfo As String = vbNullString
        For Each tModule As ProcessModule In Process.GetCurrentProcess().Modules
            Dim sModule As String = tModule.FileName
            If sModule.ToUpper.Contains("RADDIN-ADDIN-PACKED.XLL") Or sModule.ToUpper.Contains("RADDIN-ADDIN64-PACKED.XLL") Then
                sModuleInfo = FileDateTime(sModule).ToString()
            End If
        Next

        Me.Text = String.Format("About {0}", My.Application.Info.Title)
        Me.LabelProductName.Text = My.Application.Info.ProductName
        Me.LabelVersion.Text = String.Format("Version {0} Buildtime {1}", My.Application.Info.Version.ToString, sModuleInfo)
        Me.LabelCopyright.Text = My.Application.Info.Copyright
        Me.LabelCompanyName.Text = My.Application.Info.CompanyName
        Me.TextBoxDescription.Text = My.Application.Info.Description
    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click
        Dim errStr As String
        errStr = RAddin.startRnamesRefresh()
        If errStr <> vbNullString Then
            MsgBox("refresh Error: " & errStr)
        Else
            If UBound(Rcalldefnames) = -1 Then
                MsgBox("no Rdefinitions found for R_Addin in current Workbook (3 column named range (type/value/path), minimum types: rexec and script)!")
            Else
                MsgBox("refreshed Rnames from current Workbook !")
            End If
        End If
        Me.Close()
    End Sub

    Private Sub LabelCompanyName_Click(sender As Object, e As EventArgs) Handles LabelCompanyName.Click
        Process.Start(LabelCompanyName.Text)
    End Sub
End Class
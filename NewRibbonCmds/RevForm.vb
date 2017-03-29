Imports NewRibbonCmds.PrintPDFForm
Public Class RevForm

    Public Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Rev = ComboBox1.SelectedValue
        Me.Hide()
    End Sub
End Class


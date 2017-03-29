Option Explicit On

Imports System.Windows.Forms
Imports System.Linq
Imports MSVB = Microsoft.VisualBasic
Imports NewRibbonCmds.StandardAddInServer ' this is so we can return from our form the main program!
''' <summary>
''' This Form is how we start the AddNewBorder Command!
''' </summary>
''' <remarks></remarks>
Public Class SQLConnectForm
    Private fore As System.Drawing.Color
    Private back As System.Drawing.Color
    Private switch As Boolean
    Private namesCollection As New AutoCompleteStringCollection()

    ''' <summary>
    ''' Get the Autocomplete entries from our db whilst loading the form.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Form2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Using db As New DWGDetailsDataContext
            Try
                Dim MySource As String() = (From d As DETAIL In db.DETAILs Where d.PROTECTIVE_MARKING_1 IsNot Nothing Select d.PROTECTIVE_MARKING_1 Distinct).ToArray()
                tbClassification.AutoCompleteCustomSource.AddRange(MySource)
                tbClassification.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                tbClassification.AutoCompleteSource = AutoCompleteSource.CustomSource

                MySource = (From d As DETAIL In db.DETAILs Where d.DIMS_IN IsNot Nothing Select d.DIMS_IN Distinct).ToArray()
                tbDimsIn.AutoCompleteCustomSource.AddRange(MySource)
                tbDimsIn.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                tbDimsIn.AutoCompleteSource = AutoCompleteSource.CustomSource

                MySource = (From d As DETAIL In db.DETAILs Where d.SUB_CONTRACTOR IsNot Nothing Select d.SUB_CONTRACTOR Distinct).ToArray()
                tbSubCon.AutoCompleteCustomSource.AddRange(MySource)
                tbSubCon.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                tbSubCon.AutoCompleteSource = AutoCompleteSource.CustomSource

                MySource = (From d As DETAIL In db.DETAILs Where d.TOLERANCE_1 IsNot Nothing Select d.TOLERANCE_1 Distinct).ToArray()
                tbTol1.AutoCompleteCustomSource.AddRange(MySource)
                tbTol1.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                tbTol1.AutoCompleteSource = AutoCompleteSource.CustomSource

                MySource = (From d As DETAIL In db.DETAILs Where d.TOLERANCE_2 IsNot Nothing Select d.TOLERANCE_2 Distinct).ToArray()
                tbTol2.AutoCompleteCustomSource.AddRange(MySource)
                tbTol2.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                tbTol2.AutoCompleteSource = AutoCompleteSource.CustomSource

                MySource = (From d As DETAIL In db.DETAILs Where d.TOLERANCE_3 IsNot Nothing Select d.TOLERANCE_3 Distinct).ToArray()
                tbTol3.AutoCompleteCustomSource.AddRange(MySource)
                tbTol3.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                tbTol3.AutoCompleteSource = AutoCompleteSource.CustomSource

                MySource = (From d As DETAIL In db.DETAILs Where d.SURFACE_TEXTURE IsNot Nothing Select d.SURFACE_TEXTURE Distinct).ToArray()
                tbSurface.AutoCompleteCustomSource.AddRange(MySource)
                tbSurface.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                tbSurface.AutoCompleteSource = AutoCompleteSource.CustomSource

                MySource = (From d As DETAIL In db.DETAILs Where d.PROTECTIVE_FINISH IsNot Nothing Select d.PROTECTIVE_FINISH Distinct).ToArray()
                tbProFin.AutoCompleteCustomSource.AddRange(MySource)
                tbProFin.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                tbProFin.AutoCompleteSource = AutoCompleteSource.CustomSource

                MySource = (From d As DETAIL In db.DETAILs Where d.SCALE IsNot Nothing Select d.SCALE Distinct).ToArray()
                tbScale.AutoCompleteCustomSource.AddRange(MySource)
                tbScale.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                tbScale.AutoCompleteSource = AutoCompleteSource.CustomSource
                If Not isDrawing = True Then
                    RadioButton1.Visible = False
                    RadioButton2.Visible = False
                    RadioButton3.Visible = False
                    RadioButton4.Visible = False
                    Panel1.Visible = False
                    label1.Visible = False
                End If
            Catch ex As Exception
                Windows.MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using
        fore = Label2.ForeColor
        back = Label2.BackColor
    End Sub
    ''' <summary>
    ''' Assign whether this is an assembly (Inventor Drawing) or part (Inventor Part or Inventor Drawing)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cmdExecute_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdExecute.Click
        If CheckBox1.Checked = True Then
            IsAssemblyOrNot = True
            ButtonWasClicked = True
            AddDataToDatabase()
            If Not oPartDoc Is Nothing Then
                AddiProperties()
            ElseIf Not oIDWDoc Is Nothing Then
                PreparetoAddBorder()
            End If
            Me.Close()
        ElseIf CheckBox1.Checked = False Then
            IsAssemblyOrNot = False
            ButtonWasClicked = True
            AddDataToDatabase()
            If Not oPartDoc Is Nothing Then
                AddiProperties()
            ElseIf Not oIDWDoc Is Nothing Then
                PreparetoAddBorder()
            End If
            Me.Close()
        End If

    End Sub
#Region "Radiobuttons, CheckBox and label text"
    Private Sub RadioButton1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton1.Click
        Label2.Visible = True
        RadioButton1.Checked = True
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        RadioButton4.Checked = False
        Timer1.Enabled = True
        If CheckBox1.Checked = False Then
            IsAssemblyOrNot = False
            Timer1.Enabled = False
            shtsize = "A3"
        ElseIf CheckBox1.Checked = True Then
            CheckBox1.Checked = False
        End If
    End Sub

    Private Sub RadioButton2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton2.Click
        Label2.Visible = True
        RadioButton1.Checked = False
        RadioButton2.Checked = True
        RadioButton3.Checked = False
        RadioButton4.Checked = False
        CheckBox1.Checked = False
        Timer1.Enabled = True
        If CheckBox1.Checked = False Then
            IsAssemblyOrNot = False
            Timer1.Enabled = False
            shtsize = "A2"
        ElseIf CheckBox1.Checked = True Then
            CheckBox1.Checked = False
        End If
    End Sub

    Private Sub RadioButton3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton3.Click
        Label2.Visible = True
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        RadioButton3.Checked = True
        RadioButton4.Checked = False
        CheckBox1.Checked = False
        Timer1.Enabled = True
        If CheckBox1.Checked = False Then
            IsAssemblyOrNot = False
            Timer1.Enabled = False
            shtsize = "A1"
        ElseIf CheckBox1.Checked = True Then
            CheckBox1.Checked = False
        End If
    End Sub

    Private Sub RadioButton4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton4.Click
        Label2.Visible = True
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        RadioButton4.Checked = True
        CheckBox1.Checked = False
        Timer1.Enabled = True
        If CheckBox1.Checked = False Then
            IsAssemblyOrNot = False
            Timer1.Enabled = False
            shtsize = "A0"
        ElseIf CheckBox1.Checked = True Then
            CheckBox1.Checked = False
        End If
    End Sub
    Private Sub CheckBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.Click
        If RadioButton1.Checked = True Or RadioButton2.Checked = True Or RadioButton3.Checked = True Or RadioButton4.Checked = True And CheckBox1.Checked = True Then
            IsAssemblyOrNot = True
            Timer1.Enabled = False
            Windows.MessageBox.Show("Assembly Enabled")
            Label2.Visible = False
        ElseIf CheckBox1.Checked = False Then
            Timer1.Enabled = True
            Windows.MessageBox.Show("Assembly Disabled")
            Label2.Visible = False
        End If
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick 'makes our label flash as a reminder to the user.
        Label2.ForeColor = IIf(switch, fore, back)
        switch = Not switch
    End Sub
#End Region

    Private Function FormIncomplete() As Boolean
        Dim EmptyItem As Boolean = False
        If tbTitle.Text = "" AndAlso EmptyItem <> True Then
            EmptyItem = True
            ActiveControl = tbTitle
        End If
        If tbClassification.Text = "" AndAlso EmptyItem <> True Then
            EmptyItem = True
            ActiveControl = tbClassification
        End If
        If tbDate.Text = "" AndAlso EmptyItem <> True Then
            EmptyItem = True
            ActiveControl = tbDate
        End If
        If tbDimsIn.Text = "" AndAlso EmptyItem <> True Then
            EmptyItem = True
            ActiveControl = tbDimsIn
        End If
        If tbDwgNum.Text = "" AndAlso EmptyItem <> True Then
            EmptyItem = True
            ActiveControl = tbDwgNum
        End If
        If tbIssue.Text = "" AndAlso EmptyItem <> True Then
            EmptyItem = True
            ActiveControl = tbIssue
        End If
        If tbScale.Text = "" AndAlso EmptyItem <> True Then
            EmptyItem = True
            ActiveControl = tbScale
        End If
        If tbNumSheets.Text = "" AndAlso EmptyItem <> True Then
            EmptyItem = True
            ActiveControl = tbNumSheets
        End If
        If tbSheetNum.Text = "" AndAlso EmptyItem <> True Then
            EmptyItem = True
            ActiveControl = tbSheetNum
        End If
        If tbSubCon.Text = "" AndAlso EmptyItem <> True Then
            EmptyItem = True
            ActiveControl = tbSubCon
        End If
        If tbTol1.Text = "" AndAlso EmptyItem <> True Then
            EmptyItem = True
            ActiveControl = tbTol1
        End If
        If tbTol2.Text = "" AndAlso EmptyItem <> True Then
            EmptyItem = True
            ActiveControl = tbTol2
        End If
        If tbTol3.Text = "" AndAlso EmptyItem <> True Then
            EmptyItem = True
            ActiveControl = tbTol3
        End If
        Return EmptyItem
    End Function

    Private Sub clearform()
        tbChange.Text = ""
        tbProFin.Text = ""
        tbSurface.Text = ""
        tbTitle.Text = ""
        tbClassification.Text = ""
        tbDate.Text = ""
        tbDimsIn.Text = ""
        tbDwgNum.Text = ""
        tbIssue.Text = ""
        tbScale.Text = ""
        tbNumSheets.Text = ""
        tbSheetNum.Text = ""
        tbSubCon.Text = ""
        tbTol1.Text = ""
        tbTol2.Text = ""
        tbTol3.Text = ""
        tbUsed01.Text = ""
        tbUsed02.Text = ""
        tbUsed03.Text = ""
        tbUsed04.Text = ""
        tbUsed05.Text = ""
        tbUsed06.Text = ""
        tbUsed07.Text = ""
        tbUsed08.Text = ""
        tbUsed09.Text = ""
        tbUsed10.Text = ""
    End Sub

    Private Sub AddDataToDatabase()
        Dim spaceloc As Integer
        Dim titleline1 As String = String.Empty
        Dim titleline2 As String = String.Empty
        Dim titleline3 As String = String.Empty
        If Not FormIncomplete() = True Then
            If tbTitle.Text.Length > 15 Then
                Dim strToTrim As String = tbTitle.Text
                spaceloc = InStr(15, strToTrim, " ")
                titleline1 = Microsoft.VisualBasic.Left(strToTrim, spaceloc)
                strToTrim = Microsoft.VisualBasic.Right(strToTrim, strToTrim.Length - spaceloc)
                spaceloc = InStr(15, strToTrim, " ")
                titleline2 = Microsoft.VisualBasic.Left(strToTrim, spaceloc)
                titleline3 = Microsoft.VisualBasic.Right(strToTrim, strToTrim.Length - spaceloc)
            Else
                titleline1 = tbTitle.Text
            End If
            Dim username As String
            If System.Environment.UserName = "jcocking" Then
                username = "jcockings"
            Else
                username = System.Environment.UserName
            End If
            username = UCase(Microsoft.VisualBasic.Left(username, 1) & "." & Microsoft.VisualBasic.Right(username, Len(username) - 1))
            Using db As New DWGDetailsDataContext
                Try
                    dwg = New DETAIL() With { _
                        .DATETIME = DateAndTime.Now(), _
                        .Drawing_Number = tbDwgNum.Text, _
                        .DRAWING_NUMBER_1 = .Drawing_Number, _
                        .DRAWING_NUMBER_2 = .DRAWING_NUMBER_1, _
                        .ISSUE_1 = tbIssue.Text, _
                        .CHANGE_NO_1 = tbChange.Text, _
                        .DATE_1 = tbDate.Text, _
                        .Sheet_Number = tbSheetNum.Text, _
                        .SHEET_NUMBER_1 = .Sheet_Number, _
                        .SHEET_NUMBER_2 = .Sheet_Number, _
                        .SHEET_NUMBER_3 = .Sheet_Number, _
                        .SHEET_NUMBER_4 = .Sheet_Number, _
                        .Number_of_sheets = tbNumSheets.Text, _
                        .NUMBER_OF_SHEETS_1 = tbNumSheets.Text, _
                        .NUMBER_OF_SHEETS_2 = tbNumSheets.Text, _
                        .NUMBER_OF_SHEETS_3 = tbNumSheets.Text, _
                        .NUMBER_OF_SHEETS_4 = tbNumSheets.Text, _
                        .DRAWN = username, _
                        .Drawn1 = username, _
                        .PROTECTIVE_MARKING_1 = tbClassification.Text, _
                        .PROTECTIVE_MARKING_2 = .PROTECTIVE_MARKING_1, _
                        .Protective_Marking_3 = .PROTECTIVE_MARKING_1, _
                        .Protective_Marking_4 = .PROTECTIVE_MARKING_1, _
                        .SUB_CONTRACTOR = tbSubCon.Text, _
                        .DIMS_IN = tbDimsIn.Text, _
                        .Dims_In1 = tbDimsIn.Text, _
                        .SCALE = tbScale.Text, _
                        .Scale1 = tbScale.Text, _
                        .PROTECTIVE_FINISH = tbProFin.Text, _
                        .SURFACE_TEXTURE = tbSurface.Text, _
                        .TOLERANCE_1 = tbTol1.Text, _
                        .TOLERANCE_2 = tbTol2.Text, _
                        .TOLERANCE_3 = tbTol3.Text, _
                        .Tolerance_4 = tbTol1.Text, _
                        .Tolerance_5 = tbTol2.Text, _
                        .Tolerance_6 = tbTol3.Text, _
                        .Sheet_Title = tbTitle.Text, _
                        .TITLE_1 = titleline1, _
                        .TITLE_2 = titleline2, _
                        .TITLE_3 = titleline3}
                    If tbUsed01.Text.Length > 0 Then
                        dwg.USED_ON_01___PREFIX = MSVB.Left(tbUsed01.Text, 4)
                        dwg.USED_ON_01___DRG_NUMBER = MSVB.Right(tbUsed01.Text, tbUsed01.Text.Length - 4)
                    End If
                    If tbUsed02.Text.Length > 0 Then
                        dwg.USED_ON_02___PREFIX = MSVB.Left(tbUsed02.Text, 4)
                        dwg.USED_ON_02___DRG_NUMBER = MSVB.Right(tbUsed02.Text, tbUsed02.Text.Length - 4)
                    End If
                    If tbUsed03.Text.Length > 0 Then
                        dwg.USED_ON_03___PREFIX = MSVB.Left(tbUsed03.Text, 4)
                        dwg.USED_ON_03___DRG_NUMBER = MSVB.Right(tbUsed03.Text, tbUsed03.Text.Length - 4)
                    End If
                    If tbUsed04.Text.Length > 0 Then
                        dwg.USED_ON_04___PREFIX = MSVB.Left(tbUsed04.Text, 4)
                        dwg.USED_ON_04___DRG_NUMBER = MSVB.Right(tbUsed04.Text, tbUsed04.Text.Length - 4)
                    End If
                    If tbUsed05.Text.Length > 0 Then
                        dwg.USED_ON_05___PREFIX = MSVB.Left(tbUsed05.Text, 4)
                        dwg.USED_ON_05___DRG_NUMBER = MSVB.Right(tbUsed05.Text, tbUsed05.Text.Length - 4)
                    End If
                    If tbUsed06.Text.Length > 0 Then
                        dwg.USED_ON_06___PREFIX = MSVB.Left(tbUsed06.Text, 4)
                        dwg.USED_ON_06___DRG_NUMBER = MSVB.Right(tbUsed06.Text, tbUsed06.Text.Length - 4)
                    End If
                    If tbUsed07.Text.Length > 0 Then
                        dwg.USED_ON_07___PREFIX = MSVB.Left(tbUsed07.Text, 4)
                        dwg.USED_ON_07___DRG_NUMBER = MSVB.Right(tbUsed07.Text, tbUsed07.Text.Length - 4)
                    End If
                    If tbUsed08.Text.Length > 0 Then
                        dwg.USED_ON_08___PREFIX = MSVB.Left(tbUsed08.Text, 4)
                        dwg.USED_ON_08___DRG_NUMBER = MSVB.Right(tbUsed08.Text, tbUsed08.Text.Length - 4)
                    End If
                    If tbUsed09.Text.Length > 0 Then
                        dwg.USED_ON_09___PREFIX = MSVB.Left(tbUsed09.Text, 4)
                        dwg.USED_ON_09___DRG_NUMBER = MSVB.Right(tbUsed09.Text, tbUsed09.Text.Length - 4)
                    End If
                    If tbUsed10.Text.Length > 0 Then
                        dwg.USED_ON_10___PREFIX = MSVB.Left(tbUsed10.Text, 4)
                        dwg.USED_ON_10___DRG_NUMBER = MSVB.Right(tbUsed10.Text, tbUsed10.Text.Length - 4)
                    End If
                    db.DETAILs.InsertOnSubmit(dwg)
                Catch ex As Exception
                    Windows.MessageBox.Show("There was a problem creating the new database entry: " & ex.Message)
                Finally
                    If db IsNot Nothing Then
                        db.SubmitChanges()
                        Me.Close()
                    End If
                End Try
            End Using
        Else
            Windows.MessageBox.Show("The form is incomplete!")
        End If
    End Sub

    Private Sub General_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles tbUsed10.KeyDown, tbUsed09.KeyDown, tbUsed08.KeyDown, tbUsed07.KeyDown, tbUsed06.KeyDown, tbUsed05.KeyDown, tbUsed04.KeyDown, tbUsed03.KeyDown, tbUsed02.KeyDown, tbUsed01.KeyDown, tbTol3.KeyDown, tbTol2.KeyDown, tbTol1.KeyDown, tbTitle.KeyDown, tbSurface.KeyDown, tbSubCon.KeyDown, tbSheetNum.KeyDown, tbScale.KeyDown, tbProFin.KeyDown, tbNumSheets.KeyDown, tbIssue.KeyDown, tbDwgNum.KeyDown, tbDimsIn.KeyDown, tbDate.KeyDown, tbClassification.KeyDown, MyBase.KeyDown, tbChange.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                If e.KeyCode = Keys.Enter Then
                    e.Handled = True
                    Me.ProcessTabKey(Not e.Shift)
                Else
                    e.Handled = False
                    MyBase.OnKeyUp(e)
                End If
            ElseIf e.KeyCode = Keys.Tab Then
                If e.KeyCode = Keys.Tab Then
                    e.Handled = True
                    Me.ProcessTabKey(e.Shift)
                Else
                    e.Handled = False
                    MyBase.OnKeyUp(e)
                End If
            End If
        Catch ex As Exception
            Windows.MessageBox.Show("There was an error with the General_Keydown Method: " & ex.Message)
        End Try
    End Sub
End Class
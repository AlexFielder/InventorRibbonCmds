Imports Inventor

' this is so we can return from our form the main program!

Public Class Form_NotFullyConstrained
    Public ThisApplication As Inventor.Application
    Dim mlBadSketchCounter As Long
    Dim mlUnknownSketchCounter As Long
    Dim msContentCenterPath As String

    Private Sub UserForm_Initialize()
        Dim oFileLocations As Inventor.FileLocations
        Dim lNum As Long
        Dim asNames() As String = Nothing
        Dim asPaths() As String = Nothing
        Dim iIndex As Integer

        On Error Resume Next
        oFileLocations = ThisApplication.FileLocations
        oFileLocations.Libraries(lNum, asNames, asPaths)
        If lNum > 0 Then
            For iIndex = 0 To UBound(asNames)
                If LCase(asNames(iIndex)) Like "*content*" Then
                    msContentCenterPath = asPaths(iIndex)
                End If
            Next
        End If

        lblFile.Text = ""
        lblFound.Text = ""

        Me.Left = CSng(GetSetting("SketchesNotFullyConstrainedInStructure", _
        "StartUp", "Left", "0"))
        Me.Top = CSng(GetSetting("SketchesNotFullyConstrainedInStructure", _
        "StartUp", "Top", "0"))
    End Sub

    Private Sub UserForm_QueryClose(ByVal Cancel As Integer, ByVal CloseMode As Integer)
        SaveSetting("SketchesNotFullyConstrainedInStructure", _
        "StartUp", "Left", CStr(Me.Left))
        SaveSetting("SketchesNotFullyConstrainedInStructure", _
        "StartUp", "Top", CStr(Me.Top))
    End Sub

    Private Sub cmdScanActiveDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdScanActiveDoc.Click
        Dim oActiveDoc As Document
        Dim oIdw As DrawingDocument
        Dim oSheet As Sheet
        Dim oDrawingView As DrawingView
        Dim oDocument As Document
        Dim oModels As New Collection
        Dim oModel As Document
        Dim lDocumentType As DocumentTypeEnum

        oActiveDoc = ThisApplication.ActiveDocument
        If oActiveDoc Is Nothing Then
            MsgBox("No assembly, part or drawing (of assembly or part) is " & _
            "open in Inventor.", vbExclamation, "No Model or Part active")
            Exit Sub
        ElseIf oActiveDoc.DocumentType = Inventor.DocumentTypeEnum.kDrawingDocumentObject Then
            oIdw = oActiveDoc
            On Error Resume Next
            For Each oSheet In oIdw.Sheets
                For Each oDrawingView In oSheet.DrawingViews
                    oDocument = oDrawingView. _
                    ReferencedDocumentDescriptor.ReferencedDocument
                    Debug.Print(oDocument.FullFileName)
                    lDocumentType = oDocument.DocumentType
                    If lDocumentType = Inventor.DocumentTypeEnum.kAssemblyDocumentObject Or _
                    lDocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject Then

                        oModels.Add(oDocument, oDocument.FullFileName)
                    End If
                Next
            Next
            On Error GoTo 0
        ElseIf oActiveDoc.DocumentType = Inventor.DocumentTypeEnum.kAssemblyDocumentObject Or _
        oActiveDoc.DocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject Then

            oModels.Add(oActiveDoc)
        Else
            MsgBox("Assembly, part or drawing (of assembly or part) " & _
            "must be open in Inventor", vbCritical, "User Error")
            Exit Sub
        End If


        chkIncludeUnknownConstraintStatus.Visible = False
        lblFound.Text = "0 non-fully constrained sketch found thus far"
        mlBadSketchCounter = 0
        mlUnknownSketchCounter = 0
        List1.Items.Clear()

        For Each oModel In oModels
            If oModel.DocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject Then
                UnderOrOverConstrainedSketches(oModel)
            Else 'Assembly
                For Each oDocument In oModel.AllReferencedDocuments
                    If oDocument.DocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject Then
                        If msContentCenterPath = "" Then
                            UnderOrOverConstrainedSketches(oDocument)
                        ElseIf Not LCase(oDocument.FullFileName) Like _
                        LCase(msContentCenterPath) & "*" Then

                            UnderOrOverConstrainedSketches(oDocument)
                        End If
                    End If
                Next
            End If
        Next

        lblFound.Text = mlBadSketchCounter & _
        " non-fully constrained sketch found"
        lblFile.Text = ""
        chkIncludeUnknownConstraintStatus.Visible = True
        If List1.Items.Count = 0 Then
            List1.Items.Add("No under or over constrained sketches found.")
        End If
        If mlUnknownSketchCounter > 0 Then
            List1.Items.Add(mlUnknownSketchCounter & " sketches have a " & _
            "contraint status of 'unknown' (most often has " & _
            "construction geometry & is fully constrained.).")
        End If
    End Sub

    Private Sub cmdOpenModel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOpenModel.Click

        Dim iNumWorkgroups As Long
        Dim asNames() As String = Nothing
        Dim asPaths() As String = Nothing
        Dim oDoc As Document
        Dim sInitDir As String = Nothing
        Dim oFileDlg As FileDialog = Nothing

        'Get the list of workgroup paths.
        On Error Resume Next
        If ThisApplication.FileLocations.Workspace <> "" Then
            sInitDir = ThisApplication.FileLocations.Workspace
        Else
            ThisApplication.FileLocations.Workgroups(iNumWorkgroups, _
            asNames, asPaths)
            If iNumWorkgroups > 0 Then sInitDir = asPaths(0)
        End If
        Err.Clear()
        ThisApplication.CreateFileDialog(oFileDlg)
        'Setup parameters for the open FileDlg then display it.
        With oFileDlg
            .InitialDirectory = sInitDir
            .DialogTitle = "Select Assembly"
            .Filter = "Inventor Files (*.iam;*.ipt;*.idw)|*.iam;*.ipt|" & _
            "Inventor Assembly (*.iam)|*.iam|Inventor Part (*.ipt)|" & _
            "*.ipt|Inventor Drawing (*.idw)|*.idw"
            .FilterIndex = 0
            .InsertMode = False 'By default is True. This argument controls whether the OnFileOpenDialog event or the OnFileInsertDialog event fires for this instance of FileDialog
            .CancelError = True 'Set the flag so an error will be raised if the user clicks the Cancel button.
            .ShowOpen()
        End With
        If Not Err.Number = 0 Then Exit Sub

        oDoc = ThisApplication.Documents.Open(oFileDlg.FileName)
        If Not Err.Number = 0 Then
            If Err.Number = -2147467259 Then
                MsgBox("Failed to open document. One possiblity is the " & _
                "document could be a newer version than your Inventor's " & _
                "version.", vbCritical, "Open Error")
            Else
                MsgBox(Err.Number & " " & Err.Description, vbCritical, "Error")
            End If
            Exit Sub
        End If
        cmdScanActiveDoc_Click(sender, e)
    End Sub

    Private Sub cmdQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQuit.Click
        Me.Hide()
    End Sub

    Private Sub UnderOrOverConstrainedSketches(ByVal oPartDocument As Inventor.PartDocument)
        Dim oComponentDefinition As Inventor.ComponentDefinition
        Dim sDisplayName As String
        Dim sSketchName As String
        Dim oSketch As Inventor.Sketch
        Dim bAtLeast1 As Boolean
        Dim lConstraintStatus As Inventor.ConstraintStatusEnum
        Dim oSketchEntity As Inventor.SketchEntity

        oComponentDefinition = oPartDocument.ComponentDefinition
        sDisplayName = oPartDocument.DisplayName
        lblFile.Text = sDisplayName

        On Error Resume Next 'Could error out. I haven't investigated why yet.
        For Each oSketch In oComponentDefinition.Sketches
            sSketchName = oSketch.Name
            lConstraintStatus = oSketch.ConstraintStatus

            'Unknown is filtered out because there is no way of knowing
            'what (& how many) situations cause this condition. A fully
            'constrained sketch with contruction lines will come back
            'with Unknown contraint status.
            If lConstraintStatus = Inventor.ConstraintStatusEnum.kUnknownConstraintStatus Then
                mlUnknownSketchCounter = mlUnknownSketchCounter + 1
            End If

            If (lConstraintStatus = Inventor.ConstraintStatusEnum.kUnderConstrainedConstraintStatus Or _
            lConstraintStatus = Inventor.ConstraintStatusEnum.kOverConstrainedConstraintStatus Or _
            IIf(chkIncludeUnknownConstraintStatus.Checked, _
            lConstraintStatus = Inventor.ConstraintStatusEnum.kUnknownConstraintStatus, False)) And _
            oSketch.Parent.Type <> ObjectTypeEnum.kSheetMetalComponentDefinitionObject Then
                'Many sheetmetal sketches are not fully constrained and are
                'hidden because they are created and manipulated by a
                'sheetmetal feature (E.G. bend).

                If Not bAtLeast1 Then
                    List1.Items.Add("")
                    List1.Items.Add(sDisplayName & vbTab & vbTab & _
                    oComponentDefinition.Document.FullFileName)
                    bAtLeast1 = True
                End If

                List1.Items.Add(vbTab & sSketchName & " " & _
                sConstraintStatusEnum(lConstraintStatus) & ".")
                mlBadSketchCounter = mlBadSketchCounter + 1
                lblFound.Text = mlBadSketchCounter & _
                " non-fully constrained sketch found thus far"
                'Used for auto scrolling purposes
                List1.SelectedItem(List1.Items.Count - 1) = True
            End If
        Next
    End Sub

    Function sConstraintStatusEnum(ByVal lConstraintStatusEnum As Inventor.ConstraintStatusEnum) As Inventor.ConstraintStatusEnum
        Select Case lConstraintStatusEnum
            Case 51713
                sConstraintStatusEnum = "is fully constrained"
            Case 51714
                sConstraintStatusEnum = "is under constrained"
            Case 51715
                sConstraintStatusEnum = "is over constrained"
            Case 51716
                sConstraintStatusEnum = "constraint status unknown"
        End Select
    End Function

    'Private Sub List1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    'End Sub

    Private Sub Form_NotFullyConstrained_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Escape Then Me.Close()
    End Sub

    Private Sub List1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles List1.DoubleClick
        Dim sLine As String
        Dim iIndex As Integer

        iIndex = List1.SelectedIndex
        sLine = List1.Items(iIndex)
        If sLine Like vbTab & "*" Then
            Dim sSketch As String
            Dim oPartDocument As PartDocument
            Dim sPathFile As String
            Dim iC As Integer
            Dim oSS As SelectSet
            Dim oControlDef As ControlDefinition
            Dim iStart As Integer

            iStart = 2
            sSketch = Mid(sLine, iStart, InStr(sLine, " ") - iStart)
            iC = iIndex - 1
            While List1.Items(iC) Like vbTab & "*"

                iC = iC - 1
            End While
            sLine = List1.Items(iC)
            sPathFile = Mid(sLine, InStrRev(sLine, vbTab) + 1)
            oPartDocument = ThisApplication.Documents.Open(sPathFile)

            ' Place the desired object into the select set.
            oSS = oPartDocument.SelectSet
            oSS.Clear()
            oSS.Select(oPartDocument.ComponentDefinition.Sketches.Item(sSketch))
            oControlDef = ThisApplication.CommandManager. _
         ControlDefinitions.Item("AppFindInBrowserCtxCmd")
            oControlDef.Execute()

            If chkEditSketch.Checked Then oPartDocument.ComponentDefinition. _
         Sketches.Item(sSketch).Edit()
        End If
    End Sub



End Class
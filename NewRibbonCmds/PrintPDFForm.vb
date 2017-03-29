Imports System
Imports System.IO

Imports Inventor

''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Public Class PrintPDFForm
    Inherits System.Windows.Forms.Form

    Private maxTime As Long = 20

    Private WithEvents _PDFCreator As PDFCreator.clsPDFCreator
    Private pErr As PDFCreator.clsPDFCreatorError
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel

    ' Inventor application object.
    ' Private WithEvents ThisApplication As Inventor.Application
    ' Public ThisApplication As Inventor.Application

    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents CBAllDocs As System.Windows.Forms.CheckBox

    Private ProjectLocation As String
    Private Project As String
    ' Private revisions(100, 4) As String ' 10 is the maximum number of revisions the border can have
    Private revisions() As String
    Public Shared Rev As Integer
    Private ReadyState As Boolean
    Public bRestart As Boolean
    Public Shared ThisApplication As Inventor.Application = NewRibbonCmds.StandardAddInServer.ThisApplication

#Region "From Windows Form Designer generated Code"

    Public Sub New()
        MyBase.New()
        InitializeComponent()
    End Sub

    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    Private components As System.ComponentModel.IContainer

    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
        Me.Button7 = New System.Windows.Forms.Button
        Me.CBAllDocs = New System.Windows.Forms.CheckBox
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Enabled = False
        Me.Button1.Location = New System.Drawing.Point(8, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(152, 40)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Show options dialog"
        '
        'Button2
        '
        Me.Button2.Enabled = False
        Me.Button2.Location = New System.Drawing.Point(8, 56)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(152, 40)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Show logfile dialog"
        '
        'Button3
        '
        Me.Button3.Enabled = False
        Me.Button3.Location = New System.Drawing.Point(8, 102)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(152, 40)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "Print printer testpage"
        '
        'Button4
        '
        Me.Button4.Enabled = False
        Me.Button4.Location = New System.Drawing.Point(8, 150)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(152, 40)
        Me.Button4.TabIndex = 4
        Me.Button4.Text = "Print PDFCreator testpage"
        '
        'Button5
        '
        Me.Button5.Enabled = False
        Me.Button5.Location = New System.Drawing.Point(8, 196)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(152, 40)
        Me.Button5.TabIndex = 5
        Me.Button5.Text = "Convert to PDF"
        '
        'Button6
        '
        Me.Button6.Enabled = False
        Me.Button6.Location = New System.Drawing.Point(8, 244)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(152, 40)
        Me.Button6.TabIndex = 6
        Me.Button6.Text = "Convert to TIFF"
        '
        'Timer1
        '
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 302)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(335, 22)
        Me.StatusStrip1.TabIndex = 7
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(42, 17)
        Me.ToolStripStatusLabel1.Text = "Status:"
        '
        'Button7
        '
        Me.Button7.Enabled = False
        Me.Button7.Location = New System.Drawing.Point(200, 56)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(121, 23)
        Me.Button7.TabIndex = 10
        Me.Button7.Text = "Create PDF(s)"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'CBAllDocs
        '
        Me.CBAllDocs.AutoSize = True
        Me.CBAllDocs.Location = New System.Drawing.Point(200, 102)
        Me.CBAllDocs.Name = "CBAllDocs"
        Me.CBAllDocs.Size = New System.Drawing.Size(129, 17)
        Me.CBAllDocs.TabIndex = 11
        Me.CBAllDocs.Text = "All Open Documents?"
        Me.CBAllDocs.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(335, 324)
        Me.Controls.Add(Me.CBAllDocs)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
        Me.Text = "Sample1 - PDFCreator COM interface"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Parameters As String
        ToolStripStatusLabel1.Text = "Status: Program is started."
        pErr = New PDFCreator.clsPDFCreatorError
        _PDFCreator = New PDFCreator.clsPDFCreator

        Parameters = "/NoProcessingAtStartup"
        Do
            bRestart = False
            _PDFCreator = New PDFCreator.clsPDFCreator
            If _PDFCreator.cStart(Parameters) = False Then
                'PDF Creator is already running.  Kill the existing process
                Shell("taskkill /f /im PDFCreator.exe", AppWinStyle.Hide, True)
                _PDFCreator = Nothing
                bRestart = True
                ToolStripStatusLabel1.Text = "Status: Error[" & pErr.Number & "]: " & pErr.Description
            Else
                Button1.Enabled = True
                Button2.Enabled = True
                Button3.Enabled = True
                Button4.Enabled = True
                Button5.Enabled = True
                Button6.Enabled = True
                Button7.Enabled = True
            End If
        Loop Until bRestart = False
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub PDFCreator_Ready() Handles _PDFCreator.eReady
        ToolStripStatusLabel1.Text = "Status: """ & _PDFCreator.cOutputFilename & """ was created!"
        _PDFCreator.cPrinterStop = True
        ReadyState = True
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub _PDFCreator_eError() Handles _PDFCreator.eError
        pErr = _PDFCreator.cError
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        _PDFCreator.cShowOptionsDialog(True)
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        _PDFCreator.cShowLogfileDialog(True)
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        _PDFCreator.cDefaultPrinter = "PDFCreator"
        ' Wait 1 second
        Timer1.Interval = 1000
        Timer1.Enabled = True
        Do While Timer1.Enabled
            System.Windows.Forms.Application.DoEvents()
        Loop
        _PDFCreator.cPrinterStop = False
        _PDFCreator.cPrintPrinterTestpage()
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        _PDFCreator.cPrintPDFCreatorTestpage()
        _PDFCreator.cPrinterStop = False
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        PrintIt(0) 'pdf
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        PrintIt(5) 'tiff
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '  Dim ThisApplication As Inventor.Application = Instance.ThisApplication
        If Not ThisApplication.Documents.Count = 0 Then
            If CBAllDocs.Checked = True Then ' print all documents
                StartPrinting(1)
            Else
                StartPrinting(0)
            End If
        Else
            MsgBox("You need to have a document open in order to print it!", MsgBoxStyle.Exclamation, "No Document open!")
            Me.Hide()
            Exit Sub
        End If
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="AllDocs"></param>
    ''' <remarks></remarks>
    Public Sub StartPrinting(ByVal AllDocs As Long)
        ' Dim ThisApplication As Inventor.Application = Instance.ThisApplication
        Dim oDoc As Document
        If AllDocs = 0 Then ' single document
            ' get the active document for printing!
            oDoc = ThisApplication.ActiveDocument
            If oDoc.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                ' continue to print
                Print(oDoc)
            Else
                MsgBox("Wrong document Type!", MsgBoxStyle.Critical, "Wrong Document Type!")
                Exit Sub
            End If
        ElseIf AllDocs = 1 Then ' multiple documents
            Dim oDocs As Documents
            oDocs = ThisApplication.Documents
            For Each oDoc In oDocs
                ' set the active document for printing!

                If oDoc.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                    oDoc.Activate()
                    ' continue to print
                    Print(oDoc)
                    'Cleanup()
                End If
            Next
            ToolStripStatusLabel1.Text = "Status: Finished Printing PDFs!"
        End If

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Cleanup()
        _PDFCreator.cClose()
        pErr = Nothing
        _PDFCreator = Nothing
        Shell("taskkill /f /im PDFCreator.exe", AppWinStyle.Hide, True)
        bRestart = True
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="oDocToPrint"></param>
    ''' <remarks></remarks>
    Public Sub Print(ByVal oDocToPrint As Document)
        ' a few variables

        Dim numsheets As Integer

        Dim sht As Sheet
        Dim outName As String
        Dim i As Integer
        Dim Latestrev As String
        Dim Searchstring As String = ""

        ' Set reference to drawing print manager
        Dim oDrgPrintMgr As DrawingPrintManager
        oDrgPrintMgr = oDocToPrint.PrintManager

        ' Set the printer name
        oDrgPrintMgr.Printer = "PDFCreator"

        If UCase(oDocToPrint.FullFileName) Like UCase("*ALI*") Then
            ProjectLocation = "C:\temp\test_pdfs\"
        Else
            ProjectLocation = "\\bas047\Aliquot\pdfs\AWE FINAL DRAWINGS"
        End If
        Project = "ALIQUOT"

        If Not bRestart = False Then StartPrinterAgain()

        ToolStripStatusLabel1.Text = "Status: Start creating pdf..."

        For Each sht In oDocToPrint.Sheets
            _PDFCreator.cPrinterStop = False
            sht.Activate()
            With Timer1
                .Interval = maxTime * 1000
                .Enabled = True
                Do While Not ReadyState And .Enabled
                    System.Windows.Forms.Application.DoEvents()
                Loop
                .Enabled = False
            End With
            'Set the paper size , scale and orientation
            oDrgPrintMgr.ScaleMode = PrintScaleModeEnum.kPrintFullScale ' kPrintBestFitScale
            ' Change the paper size to a custom size. The units are in centimeters.
            Dim shtsize As Long
            shtsize = sht.Size
            oDrgPrintMgr.PaperSize = PaperSizeEnum.kPaperSizeCustom
            If shtsize = 9993 Then ' A0
                oDrgPrintMgr.PaperHeight = 84.1
                oDrgPrintMgr.PaperWidth = 118.9
            ElseIf shtsize = 9994 Then ' A1
                oDrgPrintMgr.PaperHeight = 59.4
                oDrgPrintMgr.PaperWidth = 84.1
            ElseIf shtsize = 9995 Then ' A2
                oDrgPrintMgr.PaperHeight = 42
                oDrgPrintMgr.PaperWidth = 59.4
            ElseIf shtsize = 9996 Then ' A3
                oDrgPrintMgr.PaperHeight = 29.7
                oDrgPrintMgr.PaperWidth = 42
            ElseIf shtsize = 9997 Then ' A4
                oDrgPrintMgr.PaperHeight = 29.7
                oDrgPrintMgr.PaperWidth = 21
            End If
            oDrgPrintMgr.PrintRange = PrintRangeEnum.kPrintCurrentSheet
            If shtsize <> 9997 Then
                oDrgPrintMgr.Orientation = PrintOrientationEnum.kLandscapeOrientation
                oDrgPrintMgr.Rotate90Degrees = True
            Else
                oDrgPrintMgr.Orientation = PrintOrientationEnum.kPortraitOrientation
                oDrgPrintMgr.Rotate90Degrees = False
            End If
            oDrgPrintMgr.AllColorsAsBlack = False

            Latestrev = RetrieveRev()

            If UCase(Project) = "EDD" Then
                Searchstring = "DRAWING NUMBER 1"
            ElseIf UCase(Project) = "ALIQUOT" Then
                If oDocToPrint.FullFileName Like "*ALI*" Then
                    Searchstring = "<Drawing Number>"
                Else
                    ' this deals with the other Borders!
                    Searchstring = "DRAWING NUMBER 1"
                End If
            End If
            outName = RetrievePE(Searchstring, sht) & " REV " & Latestrev & ".pdf"
            If UCase(outName) Like UCase("*AS-*") Then
                maxTime = 20
            Else
                ' maxTime = 20
            End If

            With _PDFCreator
                .cOption("UseAutosave") = 1
                .cOption("UseAutosaveDirectory") = 1
                .cOption("AutosaveDirectory") = ProjectLocation
                .cOption("AutosaveFilename") = outName
                .cOption("AutosaveFormat") = 0                            ' 0 = PDF
                .cClearCache()
            End With

            oDrgPrintMgr.SubmitPrint()
            With Timer1
                .Interval = maxTime * 1000
                .Enabled = True
                Do While Not ReadyState And .Enabled
                    System.Windows.Forms.Application.DoEvents()
                Loop
                .Enabled = False
            End With
            For i = 1 To 8
                revisions(i) = ""
            Next i
            numsheets = numsheets + 1
            If Not ReadyState Then
                MsgBox("Creating printer test page as pdf." & vbCrLf & vbCrLf & _
                 "An error is occured: Time is up!", MsgBoxStyle.Exclamation, Me.Text)
            Else
                ToolStripStatusLabel1.Text = "Status: Printed: " & outName
            End If
            _PDFCreator.cPrinterStop = False
            'Wait until the file shows up before closing PDF Creator
            Do
                System.Windows.Forms.Application.DoEvents()
            Loop Until Dir(ProjectLocation & outName) = outName
        Next
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub StartPrinterAgain()
        Dim Parameters As String
        ToolStripStatusLabel1.Text = "Status: PDFCreator Program is restarted."
        pErr = New PDFCreator.clsPDFCreatorError
        '_PDFCreator = New PDFCreator.clsPDFCreator

        Parameters = "/NoProcessingAtStartup"
        Do
            bRestart = False
            _PDFCreator = New PDFCreator.clsPDFCreator

            If _PDFCreator.cStart(Parameters) = False Then
                'PDF Creator is already running.  Kill the existing process
                Shell("taskkill /f /im PDFCreator.exe", AppWinStyle.Hide, True)
                _PDFCreator = Nothing
                bRestart = True
                ToolStripStatusLabel1.Text = "Status: Error[" & pErr.Number & "]: " & pErr.Description
            End If
        Loop Until bRestart = False
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function RetrieveRev() As String ' will only work whilst the revision is numeric!
        ' 2009-11-06 A.Fielder Changed RetrieveRev to work with the other Borders!
        ' 2009-11-12 A.Fielder Reverted changes to make this once again work with CH2M Borders - will
        '                       need to have a think about making it work with the other borders!
        RetrieveRev = String.Empty
        Dim RevForm As New RevForm

        Dim oDrawDoc As DrawingDocument
        oDrawDoc = ThisApplication.ActiveDocument
        Dim oSheet As Sheet
        oSheet = oDrawDoc.ActiveSheet
        ' Get all the (prompted) revision values from the title block.
        ' and add them to an array so we can sort them.
        ' Debug.Print "Retrieving revision from " & oSheet.Name
        Dim oBorderDef As BorderDefinition

        oBorderDef = oSheet.Border.Definition

        Dim oTextBox As Inventor.TextBox
        Dim bFound As Boolean
        bFound = False
        Dim Revision As String
        Dim cnt As Integer

        Dim i As Integer
        i = 1
        cnt = 0
        For Each oTextBox In oBorderDef.Sketch.TextBoxes
            Revision = GetPromptField(oTextBox.FormattedText)
            If Revision <> "" Then
                If UCase(Revision) Like "*REV*" And Len(Revision) < 13 Or UCase(Revision) Like "*ISSUE*" Then
                    revisions(i) = oSheet.Border.GetResultText(oTextBox) ' Rev
                    cnt = cnt + 1
                End If
                revisions(i) = oSheet.Name
                If cnt = 3 Then
                    cnt = 0
                    i = i + 1
                End If
            End If
        Next

        For i = LBound(revisions) To UBound(revisions)
            If revisions(i) <> "" Then
                Debug.Print(revisions(i))
            End If
        Next i
        AlphaNumericSort.Main(revisions)
        Bubblesort()
        For i = LBound(revisions) To UBound(revisions)
            If revisions(i) <> "" Then
                RetrieveRev = revisions(i)
                If revisions(i + 1) = "" Then ' we reached the highest revision.
                    RetrieveRev = revisions(i)
                    Exit For
                End If
                ' unecessary as we are now re-cycling revisions in the same revision box.
                ' i.e. when we get to 7 we start moving the revisions down.
                'If RetrieveRev = 7 Then
                '    For j = (revisions(i)) To 20
                '        RevForm.ComboBox1.Items.Add(j)
                '    Next j
                '    RevForm.Show()
                '    RetrieveRev = Rev
                '    Exit Function
                'End If
            End If
        Next i

    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Searchstring"></param>
    ''' <param name="oSheet"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function RetrievePE(ByVal Searchstring As String, ByVal oSheet As Sheet) As String
        RetrievePE = String.Empty
        Dim oDrawDoc As DrawingDocument
        oDrawDoc = ThisApplication.ActiveDocument
        ' Dim oSheet As sheet
        ' Set oSheet = oDrawDoc.ActiveSheet

        ' Get the prompted text value from the title block.
        ' This is done by first getting the text box in the title
        ' block definition that defines the prompted text.  Then
        ' you can use this to get the value specified for this
        ' particular title block instance.

        Dim oBorderDef As BorderDefinition

        oBorderDef = oSheet.Border.Definition

        Dim oTextBox As Inventor.TextBox
        Dim bFound As Boolean
        bFound = False
        For Each oTextBox In oBorderDef.Sketch.TextBoxes
            If GetPromptField(oTextBox.FormattedText) = Searchstring Then
                RetrievePE = oSheet.Border.GetResultText(oTextBox)
                Return RetrievePE
                bFound = True
                Exit For
            End If
        Next
        If Not bFound = True Then
            MsgBox("Specified formatted text was not found in the title block.")
        End If

    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="FormattedText"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPromptField(ByVal FormattedText As String) As String

        On Error GoTo ErrorFound

        ' Verify that this is a prompt field.
        If Strings.Left(FormattedText, 7) <> "<Prompt" Then
            GetPromptField = ""
            Exit Function
        End If
        ' Get the text that is to the right of the first ">" symbol
        ' and to the left of the last "<" symbol.
        ' Debug.Print FormattedText
        GetPromptField = Strings.Right(FormattedText, Len(FormattedText) - InStr(FormattedText, ">"))
        GetPromptField = Strings.Left(GetPromptField, InStr(GetPromptField, "<") - 1)
        ' Replace any &lt; or &gt; with < and > symbols.
        If Project = "ALIQUOT" Then
            GetPromptField = Replace(GetPromptField, "&lt;", "<")
            GetPromptField = Replace(GetPromptField, "&gt;", ">")
        End If
        Exit Function
ErrorFound: GetPromptField = ""

    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Bubblesort()
        Dim i As Integer
        Dim iOuter As Long
        Dim iInner As Long
        Dim iLbound As Long
        Dim iUbound As Long
        Dim iTemp As String

        iLbound = LBound(revisions)
        For i = iLbound To UBound(revisions) ' - 1
            If revisions(i) <> "" Then
                iUbound = i
            End If
        Next i

        For iOuter = iLbound To iUbound ' - 1
            'Which comparison
            For iInner = iLbound To iUbound - iOuter - 1
                'Compare this item to the next item
                If revisions(iInner) <> "" Then ' Continue
                    ' Debug.Print "About to sort " & revisions(iInner, 4)
                    If CInt(revisions(iInner)) > CInt(revisions(iInner + 1)) Then
                        'Swap
                        iTemp = revisions(iInner)
                        revisions(iInner) = revisions(iInner + 1)
                        revisions(iInner + 1) = iTemp
                        iTemp = revisions(iInner)
                        revisions(iInner) = revisions(iInner + 1)
                        revisions(iInner + 1) = iTemp
                        iTemp = revisions(iInner)
                        revisions(iInner) = revisions(iInner + 1)
                        revisions(iInner + 1) = iTemp
                        iTemp = revisions(iInner)
                        revisions(iInner) = revisions(iInner + 1)
                        revisions(iInner + 1) = iTemp
                    End If
                End If
            Next iInner
        Next iOuter

        ' MsgBox ("Done Sorting!")
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Filetyp"></param>
    ''' <remarks></remarks>
    Public Sub PrintIt(ByVal Filetyp As Long)
        Dim fname As String, fi As FileInfo, DefaultPrinter As String
        Dim opt As PDFCreator.clsPDFCreatorOptions
        With OpenFileDialog1
            .Multiselect = False
            .CheckFileExists = True
            .CheckPathExists = True
            .Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        End With
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            fi = New FileInfo(OpenFileDialog1.FileName)
            If fi.Name.Length > 0 Then
                If InStr(fi.Name, ".", CompareMethod.Text) > 1 Then
                    fname = Mid(fi.Name, 1, InStr(fi.Name, ".", CompareMethod.Text) - 1)
                Else
                    fname = fi.Name
                End If
            End If
            If Not _PDFCreator.cIsPrintable(fi.FullName) Then
                MsgBox("File '" & fi.FullName & "' is not printable!", MsgBoxStyle.Exclamation, Me.Text)
                Exit Sub
            End If
            opt = _PDFCreator.cOptions
            With opt
                .UseAutosave = 1
                .UseAutosaveDirectory = 1
                .AutosaveDirectory = fi.DirectoryName
                .AutosaveFormat = Filetyp
                ' If Filetyp = 5 Then
                ' .BitmapResolution = 72
                ' End If
                opt.AutosaveFilename = fi.Name
            End With
            With _PDFCreator
                .cOptions = opt
                .cClearCache()
                DefaultPrinter = .cDefaultPrinter
                .cDefaultPrinter = "PDFCreator"
                .cPrintFile(fi.FullName)
                ReadyState = False
                .cPrinterStop = False
            End With

            With Timer1
                .Interval = maxTime * 1000
                .Enabled = True
                Do While Not ReadyState And .Enabled
                    System.Windows.Forms.Application.DoEvents()
                Loop
                .Enabled = False
            End With
            If Not ReadyState Then
                MsgBox("Creating printer test page as pdf." & vbCrLf & vbCrLf & _
                 "An error is occured: Time is up!", MsgBoxStyle.Exclamation, Me.Text)
            End If
            _PDFCreator.cPrinterStop = True
            _PDFCreator.cDefaultPrinter = DefaultPrinter
        End If
        opt = Nothing
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        _PDFCreator.cClose()
        pErr = Nothing
        _PDFCreator = Nothing
        Shell("taskkill /f /im PDFCreator.exe", AppWinStyle.Hide, True)
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub
End Class


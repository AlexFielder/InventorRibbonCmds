Option Explicit On

Imports System.Runtime.InteropServices
Imports System.Collections.Generic
Imports System.Data
Imports System.Drawing
Imports Inventor
Imports System.Data.SqlClient
Imports System.Linq
Imports System.IO
Imports Microsoft.Win32
Imports System.Windows
Imports IPictureDisp = stdole.IPictureDisp


<ProgId("NewRibbonCmds.StandardAddInServer"), _
    GuidAttribute("67082e55-7173-4dc3-b957-a646ff6460cc")> _
Public Class StandardAddInServer
    Implements ApplicationAddInServer
    Friend Shared Instance As StandardAddInServer
    ' Inventor application object.
    Public Shared ThisApplication As Inventor.Application
    'use this to capture errors!
    Public FunctionName As String
    'Style Switcher
    Private WithEvents m_button1Def As ButtonDefinition
    Public Shared WithEvents m_combobox1 As ComboBoxDefinition
    Public Shared oIDWDoc As DrawingDocument
    Public PEDefaults() As String
    Public Layernames() As String
    Public Shared cs As ChangeStyles = New ChangeStyles()
    'Addborder to drawing and Database!
    Public Shared oPartDoc As PartDocument
    Public Shared shtsize As String
    Public Shared IsAssemblyOrNot As Boolean
    Public Shared ButtonWasClicked As Boolean
    Public Shared sqlcmd As SqlCommand
    Public Shared ds As DataSet
    Public Shared dwg As DETAIL
    Public Shared Property isDrawing As Boolean
    Public Shared Property IsFirstTime As Boolean
    Public Shared ILSheet As Sheet = Nothing
    Public Shared DLSheet As Sheet = Nothing
    'PrintPDFs
    Private WithEvents m_PrintPDFsButtonDef As ButtonDefinition
    'Screenshot
    Private WithEvents m_SSButtonDef As ButtonDefinition
    'Check Sketches
    Private WithEvents m_MyCheckButtonDef As ButtonDefinition
    'Add IL & DL to Database - would allow QA on the data contained within them too!
    Private WithEvents m_AddtoDBButtonDef As ButtonDefinition
    'Allow the user to save some settings
    Private WithEvents m_SavesettingsButtonDef As ButtonDefinition

    Public Shared WithEvents m_applicationEvents As ApplicationEvents
    Private WithEvents m_userInterfaceEvents As UserInterfaceEvents
    Private WithEvents m_InputEvents As UserInputEvents
    Private WithEvents m_dialogueEvents As FileDialogEvents

    Private m_ClientID As String
    ' one possible way to create the strings we need:
    Public Shared TitleBlockDetailsList As List(Of PEDetails) = Nothing
    Public Shared DetailsList As List(Of PEDetails) = Nothing

#Region "ApplicationAddInServer Members"
    Public Sub Activate(ByVal addInSiteObject As ApplicationAddInSite, ByVal firstTime As Boolean) _
        Implements ApplicationAddInServer.Activate
        '2010-06-01 AF Added PrintPDFButtonDef 


        ' This method is called by Inventor when it loads the AddIn.
        ' The AddInSiteObject provides access to the Inventor Application object.
        ' The FirstTime flag indicates if the AddIn is loaded for the first time.

        ' Initialize AddIn members.
        Instance = Me
        ThisApplication = addInSiteObject.Application
        ' input events are required if we want to capture the styles in the document when it activates.
        ' technically this could be achieved by simply pulling the styles from the style library.
        m_InputEvents = ThisApplication.CommandManager.UserInputEvents
        m_applicationEvents = ThisApplication.ApplicationEvents
        ' Get the ClassID for this add-in and save it in a member variable to use wherever a ClientID is needed.
        m_ClientID = AddInGuid(GetType(StandardAddInServer))

        ' Set the large icon size based on whether it is the classic or ribbon interface.
        Dim UIManager As UserInterfaceManager = ThisApplication.UserInterfaceManager

        Dim largeIconSize As Integer
        'forgot this bit for the longest time - would have broken everything. I think?
        If ThisApplication.UserInterfaceManager.InterfaceStyle = InterfaceStyleEnum.kRibbonInterface Then
            largeIconSize = 32
        Else
            largeIconSize = 24
        End If
        Dim controlDefs As ControlDefinitions = ThisApplication.CommandManager.ControlDefinitions
        ' define a couple of pictures to use in the toolbar.

        Dim smallPicture As IPictureDisp = ToIPictureDisp(New Icon(My.Resources.Style_Switch, 16, 16))
        Dim _
            largePicture As IPictureDisp = _
                ToIPictureDisp(New Icon(My.Resources.Style_Switch, largeIconSize, largeIconSize))
        ' define 2 more icons for the screenshot tool.
        Dim ssSmallPicture As IPictureDisp = ToIPictureDisp(New Icon(My.Resources.ss, 16, 16))
        Dim ssLargePicture As IPictureDisp = ToIPictureDisp(New Icon(My.Resources.ss, largeIconSize, largeIconSize))
        ' and a further 2 icons for our print pdf button.
        Dim PPdfSmallPicture As IPictureDisp = ToIPictureDisp(New Icon(My.Resources.CH2M_PDf, 16, 16))
        Dim _
            PPdfLargePicture As IPictureDisp = _
                ToIPictureDisp(New Icon(My.Resources.CH2M_PDf, largeIconSize, largeIconSize))
        ' and 2 more icons for our QA Tool.
        Dim MyCheckSmallPicture As IPictureDisp = ToIPictureDisp(New Icon(My.Resources.tick, 16, 16))
        Dim _
            MyCheckLargePicture As IPictureDisp = _
                ToIPictureDisp(New Icon(My.Resources.tick, largeIconSize, largeIconSize))
        ' and 2 more icons for our AddtoDB Tool.
        Dim AddtoDBSmallPicture As IPictureDisp = ToIPictureDisp(New Icon(My.Resources.addtodb, 16, 16))
        Dim _
            AddtoDBLargePicture As IPictureDisp = _
                ToIPictureDisp(New Icon(My.Resources.addtodb, largeIconSize, largeIconSize))
        'and 2 more icons for our save settings button.
        Dim PsavesettingsSmallpic As IPictureDisp = ToIPictureDisp(New Icon(My.Resources.savesettings, 16, 16))
        Dim _
            PsavesettingsLargepic As IPictureDisp = _
                ToIPictureDisp(New Icon(My.Resources.savesettings, largeIconSize, largeIconSize))
        ' and then assign details to the button(s) we want to create and the combobox.
        m_button1Def = _
            controlDefs.AddButtonDefinition("Style Switch", "CH2MRibbonCmdsOne", CommandTypesEnum.kNonShapeEditCmdType, _
                                             m_ClientID, _
                                             "Allows users to automatically switch between available layout styles!", _
                                             "Allows users to automatically switch between available layout styles!", _
                                             smallPicture, largePicture)
        m_SSButtonDef = _
            controlDefs.AddButtonDefinition("Screenshot", "CH2MRibbonCmdsTwo", CommandTypesEnum.kNonShapeEditCmdType, _
                                             m_ClientID, "Allow users to take screenshots of parts/assemblies!", _
                                             "Allow users to take screenshots of parts/assemblies!", ssSmallPicture, _
                                             ssLargePicture)
        'increased the width of m_combobox1 from 50 to 75
        m_combobox1 = _
            controlDefs.AddComboBoxDefinition("Dimension Styles", "CH2MRibbonCmdsCB1", _
                                               CommandTypesEnum.kNonShapeEditCmdType, 75, m_ClientID, , _
                                               "Use this to select available Dimension Styles!")
        m_PrintPDFsButtonDef = _
            controlDefs.AddButtonDefinition("Print PDF Files", "PrintCustomPDFs", CommandTypesEnum.kQueryOnlyCmdType, _
                                             m_ClientID, "Print PDFs to agreed standards!", _
                                             "Print PDFs to Company and client standards", PPdfSmallPicture, _
                                             PPdfLargePicture)
        m_MyCheckButtonDef = _
            controlDefs.AddButtonDefinition("Check the files!", "MyCheck", CommandTypesEnum.kQueryOnlyCmdType, _
                                             m_ClientID, "Check the document to agreed standards!", _
                                             "Checks the file or file(s) currently open to agreed company and client standards!", _
                                             MyCheckSmallPicture, MyCheckLargePicture)
        m_AddtoDBButtonDef = _
            controlDefs.AddButtonDefinition("Add the DL/IL to the Database", "AddDLtoDB", _
                                             CommandTypesEnum.kQueryOnlyCmdType, m_ClientID, _
                                             "Allows the user to add the data contained within the IL &/Or DL to the database", _
                                             "Add the IL OR DL Data to our DB!", AddtoDBSmallPicture, _
                                             AddtoDBLargePicture)
        m_SavesettingsButtonDef = _
            controlDefs.AddButtonDefinition("Save some settings!", "Savesettings", _
                                             CommandTypesEnum.kNonShapeEditCmdType, m_ClientID, _
                                             "Allows the user to save some settings in the %userdata%\ribbon cmds\ folder", _
                                             "Save settings", PsavesettingsSmallpic, PsavesettingsLargepic)

        ' Create a command category and add the buttons to the category.
        Dim cmdCategory As CommandCategory
        cmdCategory = _
            ThisApplication.CommandManager.CommandCategories.Add("CH2M Ribbon Commands", _
                                                                  "CH2M Ribbon Commands Category", m_ClientID)
        cmdCategory.Add(m_SSButtonDef)
        cmdCategory.Add(m_button1Def)
        cmdCategory.Add(m_combobox1)
        cmdCategory.Add(m_PrintPDFsButtonDef)
        cmdCategory.Add(m_AddtoDBButtonDef)

        If firstTime Then
            If UIManager.InterfaceStyle = InterfaceStyleEnum.kRibbonInterface Then
                CreateOrUpdateRibbonInterface()
            End If
        End If

        ' Connect to the user interface events to be able to handle when the user resets the interface.
        m_userInterfaceEvents = UIManager.UserInterfaceEvents
    End Sub

    Private Sub CreateOrUpdateRibbonInterface()
        Dim DrawingRibbon As Ribbon = ThisApplication.UserInterfaceManager.Ribbons.Item("Drawing")
        Dim CH2MHILLDrawingTab As RibbonTab = DrawingRibbon.RibbonTabs.Add("CH2M HILL", "CH2M HILL Drawing Tab", m_ClientID, "id_TabAnnotate")
        Dim CH2MDrawingPanel As RibbonPanel = CH2MHILLDrawingTab.RibbonPanels.Add("CH2M HILL Drawing Tools", "CH2M Drawing Tools", m_ClientID)
        CH2MDrawingPanel.CommandControls.AddButton(m_button1Def, True)
        CH2MDrawingPanel.CommandControls.AddComboBox(m_combobox1)
        CH2MDrawingPanel.CommandControls.AddSeparator("CH2MRibbonCmdsCB1")
        CH2MDrawingPanel.CommandControls.AddButton(m_AddtoDBButtonDef, True)
        CH2MDrawingPanel.CommandControls.AddButton(m_SSButtonDef, True)
        CH2MDrawingPanel.CommandControls.AddButton(m_PrintPDFsButtonDef, True)
        CH2MDrawingPanel.CommandControls.AddButton(m_MyCheckButtonDef, True)

        Dim DrawingAddinsTab As RibbonTab = DrawingRibbon.RibbonTabs.Item("id_AddInsTab")
        Dim DrawingSettingsPanel As RibbonPanel = DrawingAddinsTab.RibbonPanels.Add("CH2M HILL", "CH2M HILL SettingsTab", m_ClientID)
        DrawingSettingsPanel.CommandControls.AddButton(m_SavesettingsButtonDef, True)

        Dim PartRibbon As Ribbon = ThisApplication.UserInterfaceManager.Ribbons.Item("Part")
        Dim CH2MHILLPartTab As RibbonTab = PartRibbon.RibbonTabs.Add("CH2M HILL", "CH2M HILL PartRibbonTab", m_ClientID)
        Dim CH2MPartPanel As RibbonPanel = CH2MHILLPartTab.RibbonPanels.Add("CH2M HILL Part Tools", "CH2M HILL PartPanel", m_ClientID)
        CH2MPartPanel.CommandControls.AddButton(m_SSButtonDef, True)
        CH2MPartPanel.CommandControls.AddButton(m_MyCheckButtonDef, True)
        CH2MPartPanel.CommandControls.AddButton(m_SavesettingsButtonDef, True)
        CH2MPartPanel.CommandControls.AddButton(m_AddtoDBButtonDef, True)

        Dim AssemblyRibbon As Ribbon = ThisApplication.UserInterfaceManager.Ribbons.Item("Assembly")
        Dim CH2MHILLAssemblyTab As RibbonTab = AssemblyRibbon.RibbonTabs.Add("CH2M HILL", "CH2M HILL AssemblyRibbonTab", m_ClientID)
        Dim CH2MAssemblyPanel As RibbonPanel = CH2MHILLAssemblyTab.RibbonPanels.Add("CH2M HILL Assembly Tools", "CH2M HILL AssemblyPanel", m_ClientID)
        CH2MAssemblyPanel.CommandControls.AddButton(m_SSButtonDef, True)
        CH2MAssemblyPanel.CommandControls.AddButton(m_MyCheckButtonDef, True)
        Dim AssemblyAddinsTab As RibbonTab = AssemblyRibbon.RibbonTabs.Item("id_AddInsTab")
        Dim AssemblySettingsPanel As RibbonPanel = AssemblyAddinsTab.RibbonPanels.Add("CH2M HILL", "CH2M HILL SettingsTab", m_ClientID)
        AssemblySettingsPanel.CommandControls.AddButton(m_SavesettingsButtonDef, True)

        Dim ZerodocRibbon As Ribbon = ThisApplication.UserInterfaceManager.Ribbons.Item("ZeroDoc")
        Dim zerodocAddinsTab As RibbonTab = ZerodocRibbon.RibbonTabs.Item("id_AddInsTab")
        Dim CH2MHILLSettingsPanel As RibbonPanel = zerodocAddinsTab.RibbonPanels.Add("CH2M HILL", "CH2M HILL SettingsTab", m_ClientID)
        CH2MHILLSettingsPanel.CommandControls.AddButton(m_SavesettingsButtonDef, True)

    End Sub

    Public Sub Deactivate() Implements ApplicationAddInServer.Deactivate

        ' This method is called by Inventor when the AddIn is unloaded.
        ' The AddIn will be unloaded either manually by the user or
        ' when the Inventor session is terminated.

        ' TODO:  Add ApplicationAddInServer.Deactivate implementation

        ' Release objects.
        Marshal.ReleaseComObject(ThisApplication)
        ThisApplication = Nothing

        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Sub

    Public ReadOnly Property Automation() As Object Implements ApplicationAddInServer.Automation

        ' This property is provided to allow the AddIn to expose an API 
        ' of its own to other programs. Typically, this  would be done by
        ' implementing the AddIn's API interface in a class and returning 
        ' that class object through this property.

        Get
            Return Nothing
        End Get
    End Property

    Public Sub ExecuteCommand(ByVal commandID As Integer) Implements ApplicationAddInServer.ExecuteCommand

        ' Note:this method is now obsolete, you should use the 
        ' ControlDefinition functionality for implementing commands.

    End Sub

#End Region

#Region "COM Registration"

    ' Registers this class as an AddIn for Inventor.
    ' This function is called when the assembly is registered for COM.
    <ComRegisterFunctionAttribute()> _
    Public Shared Sub Register(ByVal t As Type)

        Dim clssRoot As RegistryKey = Registry.ClassesRoot
        Dim clsid As RegistryKey = Nothing
        Dim subKey As RegistryKey = Nothing

        Try
            clsid = clssRoot.CreateSubKey("CLSID\" + AddInGuid(t))
            clsid.SetValue(Nothing, "NewRibbonCmds")
            subKey = clsid.CreateSubKey("Implemented Categories\{39AD2B5C-7A29-11D6-8E0A-0010B541CAA8}")
            subKey.Close()

            subKey = clsid.CreateSubKey("Settings")
            subKey.SetValue("AddInType", "Standard")
            subKey.SetValue("LoadOnStartUp", "1")

            'subKey.SetValue("SupportedSoftwareVersionLessThan", "")
            subKey.SetValue("SupportedSoftwareVersionGreaterThan", "12..")
            'subKey.SetValue("SupportedSoftwareVersionEqualTo", "")
            'subKey.SetValue("SupportedSoftwareVersionNotEqualTo", "")
            'subKey.SetValue("Hidden", "0")
            'subKey.SetValue("UserUnloadable", "1")
            subKey.SetValue("Version", 0)
            subKey.Close()

            subKey = clsid.CreateSubKey("Description")
            subKey.SetValue(Nothing, My.Resources.Version)

        Catch ex As Exception
            Trace.Assert(False)
        Finally
            If Not subKey Is Nothing Then subKey.Close()
            If Not clsid Is Nothing Then clsid.Close()
            If Not clssRoot Is Nothing Then clssRoot.Close()
        End Try

    End Sub

    ' Unregisters this class as an AddIn for Inventor.
    ' This function is called when the assembly is unregistered.
    <ComUnregisterFunctionAttribute()> _
    Public Shared Sub Unregister(ByVal t As Type)

        Dim clssRoot As RegistryKey = Registry.ClassesRoot
        Dim clsid As RegistryKey = Nothing

        Try
            clssRoot = Registry.ClassesRoot
            clsid = clssRoot.OpenSubKey("CLSID\" + AddInGuid(t), True)
            clsid.SetValue(Nothing, "")
            clsid.DeleteSubKeyTree("Implemented Categories\{39AD2B5C-7A29-11D6-8E0A-0010B541CAA8}")
            clsid.DeleteSubKeyTree("Settings")
            clsid.DeleteSubKeyTree("Description")
        Catch
        Finally
            If Not clsid Is Nothing Then clsid.Close()
            If Not clssRoot Is Nothing Then clssRoot.Close()
        End Try

    End Sub

    ' This property uses reflection to get the value for the GuidAttribute attached to the class.
    Public Shared ReadOnly Property AddInGuid(ByVal t As Type) As String
        Get
            Dim guid As String = ""
            Try
                Dim customAttributes() As Object = t.GetCustomAttributes(GetType(GuidAttribute), False)
                Dim guidAttribute As GuidAttribute = CType(customAttributes(0), GuidAttribute)
                guid = "{" + guidAttribute.Value.ToString() + "}"
            Finally
                AddInGuid = guid
            End Try
        End Get
    End Property

#End Region

#Region "Inventor Screenshot Tool"

    Private Sub m_InputEvents_OnActivateCommand(ByVal CommandName As String, ByVal Context As NameValueMap) _
        Handles m_InputEvents.OnActivateCommand

    End Sub

    Private Sub m_InputEvents_OnContextMenu(ByVal SelectionDevice As SelectionDeviceEnum, _
                                             ByVal AdditionalInfo As NameValueMap, ByVal CommandBar As CommandBar) _
        Handles m_InputEvents.OnContextMenu

    End Sub

    Private Sub oUserInput_OnStartCommand(ByVal CommandID As CommandIDEnum) Handles m_InputEvents.OnStartCommand

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Context"></param>
    ''' <remarks></remarks>
    Private Sub m_SSButtonDef_OnExecute(ByVal Context As NameValueMap) Handles m_SSButtonDef.OnExecute
        Dim oScreenShot As New InteractionEventsManager(ThisApplication)
        oScreenShot.StartEvent(True)
    End Sub
#End Region

#Region "PrintPDFs"

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Context"></param>
    ''' <remarks></remarks>
    Private Sub m_PrintPDFsButtonDef_OnExecute(ByVal Context As NameValueMap) Handles m_PrintPDFsButtonDef.OnExecute
        Dim PrintForm As New PrintPDFForm
        PrintForm.Show(New WindowWrapper(ThisApplication.MainFrameHWND))
    End Sub

    'Private Sub m_InputEvents_OnActivateCommand(ByVal CommandName As String, _
    '                                            ByVal Context As Inventor.NameValueMap) Handles m_InputEvents.OnActivateCommand

    '    If ThisApplication.ActiveDocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
    '        If CommandName = "DrawingEditFieldTextCmd" Then
    '            Dim oDrawDoc As DrawingDocument
    '            oDrawDoc = ThisApplication.ActiveDocument
    '            MsgBox("we managed to figure out when the user wants to edit 'field text'!", _
    '                   MsgBoxStyle.Exclamation, _
    '                   "DrawingEditFieldTextCmd grabbed!")

    '        End If
    '    End If
    'End Sub

    'add a button to a context menu.
    'Private Sub m_InputEvents_OnContextMenu(ByVal SelectionDevice As Inventor.SelectionDeviceEnum, _
    '                                        ByVal AdditionalInfo As Inventor.NameValueMap, _
    '                                        ByVal CommandBar As Inventor.CommandBar) Handles m_InputEvents.OnContextMenu
    '    Dim oCmdBarControl As CommandBarControl
    '    oCmdBarControl = CommandBar.Controls.Item("DrawingEditFieldTextCmd")
    '    If Not oCmdBarControl Is Nothing Then
    '        ' add our button before the normal "Edit Field Text" command!
    '        ' MsgBox("OnContextMenu: Commandbar Displayname= " & CommandBar.DisplayName) '& " AdditionalInfo Count= " & AdditionalInfo.Count)
    '        CommandBar.Controls.AddButton(m_EditFieldTextButtonDef, oCmdBarControl.Index)
    '    Else
    '        MsgBox("it didn't work!")
    '    End If
    'End Sub
    'execute said button.
    'Private Sub m_EditFieldTextButtonDef_OnExecute(ByVal Context As Inventor.NameValueMap) Handles m_EditFieldTextButtonDef.OnExecute
    '    Dim dbfrm As New frmDBEditFieldText
    '    dbfrm.Show(New WindowWrapper(ThisApplication.MainFrameHWND))
    'End Sub

#End Region

#Region "Document level events"

    ''' <summary>
    ''' Captures the OnCloseDocument Event
    ''' </summary>
    ''' <param name="DocumentObject">The document object being closed.</param>
    ''' <param name="FullDocumentName">The FullName of the documentobject</param>
    ''' <param name="BeforeOrAfter">Is it before or after the event firing?</param>
    ''' <param name="Context">Ummm</param>
    ''' <param name="HandlingCode">Yeah..</param>
    ''' <remarks></remarks>
    Public Shared Sub m_applicationEvents_OnCloseDocument(ByVal DocumentObject As Inventor._Document, ByVal FullDocumentName As String, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_applicationEvents.OnCloseDocument
        Try
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                If DocumentObject.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                    If oIDWDoc IsNot Nothing Then
                        oIDWDoc = Nothing
                    End If
                ElseIf DocumentObject.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    If oPartDoc IsNot Nothing Then
                        oPartDoc = Nothing
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Captures the OnNewDocument Event
    ''' </summary>
    ''' <param name="DocumentObject">The Document Type (Drawing/Part/Assembly etc.)</param>
    ''' <param name="BeforeOrAfter">Before or After the Event - we need to take control After</param>
    ''' <param name="Context">Denotes a successful command - usually contains information about the new Document</param>
    ''' <param name="HandlingCode"></param>
    ''' <remarks>This Event will only be captured if we're using the Standard.idw template file!</remarks>
    Public Shared Sub m_applicationEvents_OnNewDocument(ByVal DocumentObject As _Document, _
                                                   ByVal BeforeOrAfter As EventTimingEnum, ByVal Context As NameValueMap, _
                                                   ByRef HandlingCode As HandlingCodeEnum) _
        Handles m_applicationEvents.OnNewDocument
        'Dim i As Integer = 0
        Dim sContextName As String = ""
        Dim oContextValue As Object = Nothing
        Try
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                If DocumentObject.DocumentType = DocumentTypeEnum.kDrawingDocumentObject _
                    Or DocumentObject.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    If Context.Count > 0 Then
                        For i As Integer = 1 To Context.Count
                            Dim contextval As Object = Context.Item(i)
                            If TypeOf contextval Is Array Then
                                Dim temparray As Array
                                temparray = CType(Context.Item(i), Array)
                            Else
                                sContextName = Context.Name(i)
                                oContextValue = Context.Item(i)
                                If UCase(oContextValue) Like UCase("*STANDARD.IDW*") Then
                                    isDrawing = True
                                    oIDWDoc = ThisApplication.ActiveDocument
                                    Dim dbform As New SQLConnectForm
                                    dbform.Show((New WindowWrapper(ThisApplication.MainFrameHWND)))
                                    Exit For
                                ElseIf UCase(oContextValue) Like UCase("*Standard (mm).IPT*") Then
                                    oPartDoc = ThisApplication.ActiveDocument
                                    Dim dbform As New SQLConnectForm
                                    dbform.Text = "iProperties for Part Files"
                                    dbform.Show((New WindowWrapper(ThisApplication.MainFrameHWND)))
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("There was an error in executing the form." & vbLf & "Error Message:" & ex.Message, _
                             "DB Form Error!")
        End Try
    End Sub

    ''' <summary>
    ''' Once a drawing document is activated, combobox1 on our ribbonbar is populated.
    ''' </summary>
    ''' <param name="DocumentObject"></param>
    ''' <param name="BeforeOrAfter"></param>
    ''' <param name="Context"></param>
    ''' <param name="HandlingCode"></param>
    ''' <remarks></remarks>
    Public Shared Sub m_applicationEvents_OnActivateDocument(ByVal DocumentObject As _Document, _
                                                        ByVal BeforeOrAfter As EventTimingEnum, _
                                                        ByVal Context As NameValueMap, _
                                                        ByRef HandlingCode As HandlingCodeEnum) _
        Handles m_applicationEvents.OnActivateDocument

        If BeforeOrAfter = EventTimingEnum.kAfter Then
            ' MsgBox("After the document activates!")
            If DocumentObject.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                m_combobox1.Clear()
                ChangeStyles.PopulateCB1()
            End If
        End If
    End Sub
#End Region

#Region "Buttons etc."
    ''' <summary>
    ''' Adds the contents of Drawing Lists Parts Lists to the Database.
    ''' </summary>
    ''' <param name="Context"></param>
    ''' <remarks></remarks>
    Public Sub m_AddtoDBButtonDef_OnExecute(ByVal Context As NameValueMap) Handles m_AddtoDBButtonDef.OnExecute
        If Not IsAssemblyDrawing() = True Then
            MessageBox.Show("This tool can only be run on an Assembly Drawing that has a valid Drawing List!")
        End If
    End Sub

    ''' <summary>
    ''' Executes the ChangeStyle command.
    ''' </summary>
    ''' <param name="Context"></param>
    ''' <remarks></remarks>
    Public Shared Sub m_button1Def_OnExecute(ByVal Context As NameValueMap) Handles m_button1Def.OnExecute
        ChangeStyles.ChangeStyle()
    End Sub

    ''' <summary>
    ''' Executes the Save Settings button
    ''' </summary>
    ''' <param name="Context"></param>
    ''' <remarks>Doesn't actually save anything yet, but it will do.</remarks>
    Private Sub m_SavesettingsButtonDef_OnExecute(ByVal Context As NameValueMap) _
        Handles m_SavesettingsButtonDef.OnExecute
        MessageBox.Show("The new button works then!?")
    End Sub

    ''' <summary>
    ''' Executes the QA Check form.
    ''' </summary>
    ''' <param name="Context"></param>
    ''' <remarks></remarks>
    Private Sub m_MyCheckButtonDef_OnExecute(ByVal Context As NameValueMap) Handles m_MyCheckButtonDef.OnExecute
        Dim QAForm As New Form_NotFullyConstrained
        QAForm.Show(New WindowWrapper(ThisApplication.MainFrameHWND))
    End Sub
#End Region
#Region "AddBorder_And_iProperties"
    ''' <summary>
    ''' Prepare to add the new border to the sheet. (With a shared DataSet)
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub PreparetoAddBorder()
        Try
            If IsAssemblyOrNot = True And ButtonWasClicked = True Then
                IsFirstTime = True
                addborder(dwg)
                IsFirstTime = False
            ElseIf IsAssemblyOrNot = False And ButtonWasClicked = True Then
                IsFirstTime = True
                addborder(dwg)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Something is broken!")
        End Try
    End Sub
    ''' <summary>
    ''' Populate our custom iProperty Values.
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub AddiProperties()
        Try
            CheckiPropertiesExist()
            'perhaps need to add some error checking in case these values are empty!
            Dim invCustomPropSet As PropertySet = oPartDoc.PropertySets.Item("Inventor User Defined Properties")
            For Each invCustomProperty As Inventor.Property In invCustomPropSet
                Select Case invCustomProperty.Name
                    Case "Title"
                        invCustomProperty.Value = dwg.Sheet_Title
                    Case "DwgNum1"
                        invCustomProperty.Value = dwg.Drawing_Number
                    Case "DimsIn"
                        invCustomProperty.Value = dwg.DIMS_IN
                    Case "ProtectiveMarking1"
                        invCustomProperty.Value = dwg.PROTECTIVE_MARKING_1
                    Case "Sheetnum1"
                        invCustomProperty.Value = dwg.Sheet_Number
                    Case "NumSheets1"
                        invCustomProperty.Value = dwg.Number_of_sheets
                    Case "SubContractor"
                        invCustomProperty.Value = dwg.SUB_CONTRACTOR
                    Case "ProtectiveFinish"
                        invCustomProperty.Value = dwg.PROTECTIVE_FINISH
                    Case "SurfaceTexture"
                        invCustomProperty.Value = dwg.SURFACE_TEXTURE
                    Case "UsedOn01"
                        If dwg.USED_ON_01___DRG_NUMBER IsNot Nothing Then
                            invCustomProperty.Value = dwg.USED_ON_01___DRG_NUMBER
                        End If
                    Case "UsedOn02"
                        If dwg.USED_ON_02___DRG_NUMBER IsNot Nothing Then
                            invCustomProperty.Value = dwg.USED_ON_02___DRG_NUMBER
                        End If
                    Case "UsedOn03"
                        If dwg.USED_ON_03___DRG_NUMBER IsNot Nothing Then
                            invCustomProperty.Value = dwg.USED_ON_03___DRG_NUMBER
                        End If
                    Case "UsedOn04"
                        If dwg.USED_ON_04___DRG_NUMBER IsNot Nothing Then
                            invCustomProperty.Value = dwg.USED_ON_04___DRG_NUMBER
                        End If
                    Case "UsedOn05"
                        If dwg.USED_ON_05___DRG_NUMBER IsNot Nothing Then
                            invCustomProperty.Value = dwg.USED_ON_05___DRG_NUMBER
                        End If
                    Case "UsedOn06"
                        If dwg.USED_ON_06___DRG_NUMBER IsNot Nothing Then
                            invCustomProperty.Value = dwg.USED_ON_06___DRG_NUMBER
                        End If
                    Case "UsedOn07"
                        If dwg.USED_ON_07___DRG_NUMBER IsNot Nothing Then
                            invCustomProperty.Value = dwg.USED_ON_07___DRG_NUMBER
                        End If
                    Case "UsedOn08"
                        If dwg.USED_ON_08___DRG_NUMBER IsNot Nothing Then
                            invCustomProperty.Value = dwg.USED_ON_08___DRG_NUMBER
                        End If
                    Case "UsedOn09"
                        If dwg.USED_ON_09___DRG_NUMBER IsNot Nothing Then
                            invCustomProperty.Value = dwg.USED_ON_09___DRG_NUMBER
                        End If
                    Case "UsedOn10"
                        If dwg.USED_ON_10___DRG_NUMBER IsNot Nothing Then
                            invCustomProperty.Value = dwg.USED_ON_10___DRG_NUMBER
                        End If
                End Select
            Next
        Catch ex As Exception
            MessageBox.Show("There was an error auto-populating the iProperties of the part file: " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' Checks for the 'Title' iProperty and adds it if missing.
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub CheckiPropertiesExist()
        Dim invCustomPropSet As PropertySet = oPartDoc.PropertySets.Item("Inventor User Defined Properties")
        Dim invCustomProperty As Inventor.Property = Nothing
        Dim strTitle = ""
        Try
            invCustomProperty = (From cp As Inventor.Property In invCustomPropSet Where cp.Name = "Title").First()
        Catch ex As Exception
            IIf(invCustomProperty IsNot Nothing, True, invCustomPropSet.Add(strTitle, "Title"))
        End Try
    End Sub
    ''' <summary>
    ''' Add the correct borders to the correct 
    ''' </summary>
    ''' <param name="titleblock"></param>
    ''' <remarks></remarks>
    Private Shared Sub addborder(ByVal titleblock As DETAIL)
        Dim osheet As Sheet = Nothing
        Dim oborder As Border = Nothing
        Dim ILborder As Border = Nothing
        Dim DLBorder As Border = Nothing
        Dim BorderDef As BorderDefinition = Nothing
        Dim BorderName As String = String.Empty
        Dim oTitleBlock As TitleBlock = Nothing
        Dim TitleBlockDef As TitleBlockDefinition = Nothing
        Dim TitleBlockName As String = String.Empty
        Dim oSheets As Sheets = oIDWDoc.Sheets
        Dim sPromptStrings As String() = Nothing

        removeExcessSheets(oSheets)
        osheet = oIDWDoc.ActiveSheet
        addBordersAndTitles(osheet)
        populateDrawingSheet(osheet)
        IIf(IsAssemblyOrNot, PopulateDLAndILBorders(ILSheet), False)
        IIf(IsAssemblyOrNot, PopulateDLAndILBorders(DLSheet), False)
    End Sub
    ''' <summary>
    ''' Removes Excess sheets according to IsFirstTime and IsAssemblyOrNot variables
    ''' </summary>
    ''' <param name="sheets"></param>
    ''' <remarks></remarks>
    Public Shared Sub removeExcessSheets(ByVal sheets As Sheets)
        If IsFirstTime And Not IsAssemblyOrNot Then
            For Each sheet As Sheet In sheets
                If Not sheet.Name Like "*" & shtsize & "*" Then
                    sheet.Activate()
                    sheet.Delete()
                End If
            Next
        ElseIf IsFirstTime And IsAssemblyOrNot Then
            For Each sheet As Sheet In sheets
                If Not sheet.Name Like "*" & shtsize & "*" Then
                    If Not sheet.Name Like "*LIST*" Then
                        sheet.Activate()
                        sheet.Delete()
                    Else
                        If sheet.Name Like "*DRAWING LIST*" Then
                            DLSheet = sheet
                        ElseIf sheet.Name Like "*ITEM LIST*" Then
                            ILSheet = sheet
                        End If
                    End If
                End If
            Next
        End If
    End Sub
    ''' <summary>
    ''' Add a relevant border and Title Block to the active sheet.
    ''' </summary>
    ''' <param name="osheet"></param>
    ''' <remarks></remarks>
    Private Shared Sub addBordersAndTitles(ByVal osheet As Sheet)
        Dim oborder As Border = Nothing
        Dim BorderDef As BorderDefinition
        Dim BorderName As String = String.Empty
        Dim oTitleBlock As TitleBlock = Nothing
        Dim TitleBlockDef As TitleBlockDefinition
        Dim TitleBlockName As String = String.Empty
        Dim oSheets As Sheets = oIDWDoc.Sheets
        Dim sPromptStrings As String() = Nothing
        BorderName = "SB-" & shtsize & "_99" & Right(shtsize, 1) & "-5.2(block)"
        BorderDef = oIDWDoc.BorderDefinitions.Item(BorderName)
        sPromptStrings = ReturnCountofPromptedEntries(BorderDef)
        oborder = osheet.AddBorder(BorderDef, sPromptStrings)
        TitleBlockName = shtsize & " CH2M Title Border"
        TitleBlockDef = oIDWDoc.TitleBlockDefinitions.Item(TitleBlockName)
        sPromptStrings = ReturnCountofPromptedEntries(TitleBlockDef)
        oTitleBlock = osheet.AddTitleBlock(TitleBlockDef, , sPromptStrings)
        If IsFirstTime And IsAssemblyOrNot Then
            BorderName = "SB-IL_996-5.2(block)"
            BorderDef = oIDWDoc.BorderDefinitions.Item(BorderName)
            sPromptStrings = ReturnCountofPromptedEntries(BorderDef)
            ILSheet.AddBorder(BorderDef, sPromptStrings)
            BorderName = "SB-DL_995-5.2(block)"
            BorderDef = oIDWDoc.BorderDefinitions.Item(BorderName)
            sPromptStrings = ReturnCountofPromptedEntries(BorderDef)
            DLSheet.AddBorder(BorderDef, sPromptStrings)
        End If
    End Sub
    ''' <summary>
    ''' Populate our drawing sheet with the relevant information.
    ''' </summary>
    ''' <param name="oSheet"></param>
    ''' <remarks></remarks>
    Private Shared Sub populateDrawingSheet(ByVal oSheet As Sheet)
        Try
            Dim oTag As String = String.Empty
            Dim oString As String = String.Empty
            Dim oTitleBlockDef As TitleBlockDefinition = oSheet.TitleBlock.Definition
            For Each oTextBox As TextBox In oTitleBlockDef.Sketch.TextBoxes
                If IsPromptField(oTextBox.FormattedText) Then
                    oTag = UCase(GetPromptField(oTextBox.FormattedText))
                    Select Case oTag
                        Case UCase("CH2M DRG NO")
                            Dim tmpstr = dwg.Drawing_Number
                            If tmpstr IsNot Nothing Then
                                oSheet.TitleBlock.SetPromptResultText(oTextBox, tmpstr)
                            Else
                                oSheet.TitleBlock.SetPromptResultText(oTextBox, "")
                            End If
                        Case UCase("Number of Sheets")
                            Dim tmpstr = dwg.Number_of_sheets
                            If tmpstr IsNot Nothing Then
                                oSheet.TitleBlock.SetPromptResultText(oTextBox, tmpstr)
                            Else
                                oSheet.TitleBlock.SetPromptResultText(oTextBox, "")
                            End If
                        Case UCase("Sheet Number")
                            Dim tmpstr = dwg.Sheet_Number
                            If tmpstr IsNot Nothing Then
                                oSheet.TitleBlock.SetPromptResultText(oTextBox, tmpstr)
                            Else
                                oSheet.TitleBlock.SetPromptResultText(oTextBox, "")
                            End If
                        Case UCase("Revision number")
                            Dim tmpstr = dwg.REV_A
                            If tmpstr IsNot Nothing Then
                                oSheet.TitleBlock.SetPromptResultText(oTextBox, tmpstr)
                            Else
                                oSheet.TitleBlock.SetPromptResultText(oTextBox, "")
                            End If
                        Case UCase("Current Issue Date")
                            Dim tmpstr = dwg.REV_A_DATE
                            If tmpstr IsNot Nothing Then
                                oSheet.TitleBlock.SetPromptResultText(oTextBox, tmpstr)
                            Else
                                oSheet.TitleBlock.SetPromptResultText(oTextBox, "")
                            End If
                    End Select
                End If
            Next
            Dim oBorderDef As BorderDefinition = oSheet.Border.Definition
            For Each oTextBox As TextBox In oBorderDef.Sketch.TextBoxes
                If IsPromptField(oTextBox.FormattedText) Then
                    oTag = UCase(GetPromptField(oTextBox.FormattedText))
                    Select Case oTag
                        Case UCase("Sheet number 1")
                            Dim tmpstr = dwg.Sheet_Number
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Sheet number 2")
                            Dim tmpstr = dwg.Sheet_Number
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Sheet number 3")
                            Dim tmpstr = dwg.Sheet_Number
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Sheet number 4")
                            Dim tmpstr = dwg.Sheet_Number
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Drawing Number 1")
                            Dim tmpstr = dwg.Drawing_Number
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Drawing Number 2")
                            Dim tmpstr = dwg.Drawing_Number
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Change No 1")
                            Dim tmpstr = dwg.CHANGE_NO_1
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Issue 1")
                            Dim tmpstr = dwg.REV_A
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Date 1")
                            Dim tmpstr = dwg.REV_A_DATE
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 1 - prefix")
                            Dim tmpstr = dwg.USED_ON_01___PREFIX
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 1 - drg number")
                            Dim tmpstr = dwg.USED_ON_01___DRG_NUMBER
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 2 - prefix")
                            Dim tmpstr = dwg.USED_ON_02___PREFIX
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 2 - drg number")
                            Dim tmpstr = dwg.USED_ON_02___DRG_NUMBER
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 3 - prefix")
                            Dim tmpstr = dwg.USED_ON_03___PREFIX
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 3 - drg number")
                            Dim tmpstr = dwg.USED_ON_03___DRG_NUMBER
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 4 - prefix")
                            Dim tmpstr = dwg.USED_ON_04___PREFIX
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 4 - drg number")
                            Dim tmpstr = dwg.USED_ON_04___DRG_NUMBER
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 5 - prefix")
                            Dim tmpstr = dwg.USED_ON_05___PREFIX
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 5 - drg number")
                            Dim tmpstr = dwg.USED_ON_05___DRG_NUMBER
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 6 - prefix")
                            Dim tmpstr = dwg.USED_ON_06___PREFIX
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 6 - drg number")
                            Dim tmpstr = dwg.USED_ON_06___DRG_NUMBER
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 7 - prefix")
                            Dim tmpstr = dwg.USED_ON_07___PREFIX
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 7 - drg number")
                            Dim tmpstr = dwg.USED_ON_07___DRG_NUMBER
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 8 - prefix")
                            Dim tmpstr = dwg.USED_ON_08___PREFIX
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 8 - drg number")
                            Dim tmpstr = dwg.USED_ON_08___DRG_NUMBER
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 9 - prefix")
                            Dim tmpstr = dwg.USED_ON_09___PREFIX
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 9 - drg number")
                            Dim tmpstr = dwg.USED_ON_09___DRG_NUMBER
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 10 - prefix")
                            Dim tmpstr = dwg.USED_ON_10___PREFIX
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Used On 10 - drg number")
                            Dim tmpstr = dwg.USED_ON_10___DRG_NUMBER
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If

                        Case UCase("Subcontractor")
                            Dim tmpstr = dwg.SUB_CONTRACTOR
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("drawn")
                            Dim tmpstr = dwg.DRAWN
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Drawn1")
                            Dim tmpstr = dwg.DRAWN
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("protective marking 1")
                            Dim tmpstr = dwg.PROTECTIVE_MARKING_1
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("protective marking 2")
                            Dim tmpstr = dwg.PROTECTIVE_MARKING_1
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("protective marking 3")
                            Dim tmpstr = dwg.PROTECTIVE_MARKING_1
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("protective marking 4")
                            Dim tmpstr = dwg.PROTECTIVE_MARKING_1
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Tolerance1")
                            Dim tmpstr = dwg.TOLERANCE_1
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Tolerance2")
                            Dim tmpstr = dwg.TOLERANCE_2
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Tolerance3")
                            Dim tmpstr = dwg.TOLERANCE_3
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Tolerance4")
                            Dim tmpstr = dwg.TOLERANCE_1
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Tolerance5")
                            Dim tmpstr = dwg.TOLERANCE_2
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                        Case UCase("Tolerance6")
                            Dim tmpstr = dwg.TOLERANCE_3
                            If tmpstr Is Nothing Then
                                oSheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                oSheet.Border.SetPromptResultText(oTextBox, tmpstr)
                            End If
                    End Select
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Something is broken!")
        End Try
        
    End Sub
    ''' <summary>
    ''' Populate the DL and IL Sheet borders.
    ''' </summary>
    ''' <param name="oSheet"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function PopulateDLAndILBorders(ByVal oSheet As Sheet) As Boolean
        Dim oBorderDef As BorderDefinition = oSheet.Border.Definition
        For Each oTextBox As TextBox In oBorderDef.Sketch.TextBoxes
            If IsPromptField(oTextBox.FormattedText) Then
                For Each detail As PEDetails In DetailsList
                    'Need to change this to suit our DETAIL object!
                    If detail.TagName = GetPromptField(oTextBox.FormattedText) Then
                        If String.IsNullOrEmpty(detail.TagStringVal) Then
                            DLSheet.Border.SetPromptResultText(oTextBox, "")
                        Else
                            If detail.TagName = "DRAWING NUMBER 1" Then
                                Dim tmpdwgnum As String = Right(detail.TagStringVal, Len(detail.TagStringVal) - InStr(detail.TagStringVal, "-"))
                                Select Case oSheet.Name
                                    Case "*IL*"
                                        tmpdwgnum = "4IL-" & tmpdwgnum
                                    Case "*DL*"
                                        tmpdwgnum = "4DL-" & tmpdwgnum
                                End Select
                                DLSheet.Border.SetPromptResultText(oTextBox, tmpdwgnum)
                            Else
                                DLSheet.Border.SetPromptResultText(oTextBox, detail.TagStringVal)
                            End If
                        End If
                        Exit For
                    End If
                Next
            End If
        Next
    End Function

    ''' <summary>
    ''' Adds the relevant border to the sheet.
    ''' </summary>
    ''' <param name="ds"></param>
    ''' <remarks></remarks>
    Public Shared Sub addborder(ByVal ds As DataSet) ', ByVal typeint As Integer) '0 = Drawing, 1 = DL, 2= IL
        Dim oTag As String = String.Empty
        Dim oString As String = String.Empty
        Try
            If Not DetailsList Is Nothing Then
                DetailsList.Clear()
                'reset the list just in case we'll add our entries to if it isn't empty!
            Else
                DetailsList = New List(Of PEDetails)
            End If
            If Not TitleBlockDetailsList Is Nothing Then
                TitleBlockDetailsList.Clear()
            Else
                TitleBlockDetailsList = New List(Of PEDetails)
                ' this will always need to be "new'd" here!
            End If
            'need to create the following steps:
            '1 create a list of the data in ds using our PEDetails class - done
            '2 add a border - done
            '3 add values in Detailslist to border & Titleblock. - DONE!
            '4 Add values from dataset to DL & IL Borders
            For Each DTable As DataTable In ds.Tables ' should be one table only here!
                For Each Drow As DataRow In DTable.Rows
                    For Each Dcolumn As DataColumn In DTable.Columns
                        If Dcolumn.ColumnName = "AttributeName" Then ' contains the tag we need to look for later on.
                            oTag = Drow(Dcolumn)
                        ElseIf Dcolumn.ColumnName = "DefaultValue" Then _
                            ' should contain our string, either edited or not.
                            If Not Drow(Dcolumn) = "" Then _
                                ' need to grab the blank entries as well as those with values.
                                oString = Drow(Dcolumn)
                            Else
                                oString = ""
                            End If
                        End If
                    Next
                    DetailsList.Add(New PEDetails(oTag, oString))
                Next
            Next

            Dim osheet As Sheet
            Dim ILSheet As Sheet = Nothing
            Dim DLSheet As Sheet = Nothing
            Dim oborder As Border = Nothing
            Dim ILborder As Border = Nothing
            Dim DLBorder As Border = Nothing
            Dim BorderDef As BorderDefinition
            Dim BorderName As String = String.Empty
            Dim oTitleBlock As TitleBlock = Nothing
            'Dim TitleBlockDef As TitleBlockDefinition
            Dim TitleBlockName As String = String.Empty
            Dim oSheets As Sheets = oIDWDoc.Sheets
            Dim sPromptStrings As String() = Nothing

            If IsFirstTime And Not IsAssemblyOrNot Then
                For Each sheet As Sheet In oSheets
                    If Not sheet.Name Like "*" & shtsize & "*" Then
                        sheet.Activate()
                        sheet.Delete()
                    End If
                Next
            ElseIf IsFirstTime And IsAssemblyOrNot Then
                For Each sheet As Sheet In oSheets
                    If Not sheet.Name Like "*" & shtsize & "*" Then
                        If Not sheet.Name Like "*LIST*" Then
                            sheet.Activate()
                            sheet.Delete()
                        Else
                            If sheet.Name Like "*DRAWING LIST*" Then
                                DLSheet = sheet
                            ElseIf sheet.Name Like "*ITEM LIST*" Then
                                ILSheet = sheet
                            End If
                        End If
                    End If
                Next
            End If
            osheet = oIDWDoc.ActiveSheet

            Dim oBorderDef As BorderDefinition = osheet.Border.Definition
            For Each oTextBox As TextBox In oBorderDef.Sketch.TextBoxes
                If IsPromptField(oTextBox.FormattedText) Then
                    For Each detail As PEDetails In DetailsList
                        If detail.TagName = GetPromptField(oTextBox.FormattedText) Then
                            If String.IsNullOrEmpty(detail.TagStringVal) Then
                                osheet.Border.SetPromptResultText(oTextBox, "")
                            Else
                                osheet.Border.SetPromptResultText(oTextBox, detail.TagStringVal)
                            End If
                            Exit For
                        End If
                    Next
                End If
            Next
            If IsFirstTime And IsAssemblyOrNot Then
                BorderName = "SB-IL_996-5.2(block)"
                BorderDef = oIDWDoc.BorderDefinitions.Item(BorderName)
                sPromptStrings = ReturnCountofPromptedEntries(BorderDef)
                ILborder = ILSheet.AddBorder(BorderDef, sPromptStrings)
                BorderName = "SB-DL_995-5.2(block)"
                BorderDef = oIDWDoc.BorderDefinitions.Item(BorderName)
                sPromptStrings = ReturnCountofPromptedEntries(BorderDef)
                DLBorder = DLSheet.AddBorder(BorderDef, sPromptStrings)
            End If
            If IsFirstTime Then
                Dim oTitleBlockDef As TitleBlockDefinition = osheet.TitleBlock.Definition
                For Each detail As PEDetails In DetailsList
                    If UCase(detail.TagName) = UCase("Drawing Number") Then
                        TitleBlockDetailsList.Add(New PEDetails(detail.TagName, detail.TagStringVal))
                    ElseIf UCase(detail.TagName) = UCase("Number of sheets") Then
                        TitleBlockDetailsList.Add(New PEDetails(detail.TagName, detail.TagStringVal))
                    ElseIf UCase(detail.TagName) = UCase("Sheet number") Then
                        TitleBlockDetailsList.Add(New PEDetails(detail.TagName, detail.TagStringVal))
                    ElseIf UCase(detail.TagName) = UCase("Rev A") Then
                        TitleBlockDetailsList.Add(New PEDetails(detail.TagName, detail.TagStringVal))
                    ElseIf UCase(detail.TagName) = UCase("Rev A Date") Then
                        TitleBlockDetailsList.Add(New PEDetails(detail.TagName, detail.TagStringVal))
                    End If
                Next
                'Dim tmplist As List(Of PEDetails)
                'For Each oTextBox As TextBox In oTitleBlockDef.Sketch.TextBoxes
                '    If IsPromptField(oTextBox.FormattedText) Then
                '        oTag = UCase(GetPromptField(oTextBox.FormattedText))
                '        '1st we need to get the values that provide the information we need (Drawing number, sheet number, number of sheets, latest revision, latest date) for our titleblock from detailslist
                '        '2nd we then need to assign those values to the matching textboxes within the titleblock
                '        If oTag = "CH2M DRG NO" Then
                '            tmplist = TitleBlockDetailsList.FindAll(AddressOf Finddwgnum)
                '            For Each d As PEDetails In tmplist ' should only be one of these?!
                '                If String.IsNullOrEmpty(d.TagStringVal) Then
                '                    osheet.TitleBlock.SetPromptResultText(oTextBox, "")
                '                Else
                '                    osheet.TitleBlock.SetPromptResultText(oTextBox, d.TagStringVal)
                '                End If
                '                Exit For
                '            Next
                '        ElseIf oTag = UCase("Number of Sheets") Then
                '            tmplist = TitleBlockDetailsList.FindAll(AddressOf Findnumsheets)
                '            For Each d As PEDetails In tmplist ' should only be one of these?!
                '                If String.IsNullOrEmpty(d.TagStringVal) Then
                '                    osheet.TitleBlock.SetPromptResultText(oTextBox, "")
                '                Else
                '                    osheet.TitleBlock.SetPromptResultText(oTextBox, d.TagStringVal)
                '                End If
                '                Exit For
                '            Next
                '        ElseIf oTag = UCase("Sheet Number") Then
                '            tmplist = TitleBlockDetailsList.FindAll(AddressOf Findshtnum)
                '            For Each d As PEDetails In tmplist ' should only be one of these?!
                '                If String.IsNullOrEmpty(d.TagStringVal) Then
                '                    osheet.TitleBlock.SetPromptResultText(oTextBox, "")
                '                Else
                '                    osheet.TitleBlock.SetPromptResultText(oTextBox, d.TagStringVal)
                '                End If
                '                Exit For
                '            Next
                '        ElseIf oTag = UCase("Revision number") Then
                '            tmplist = TitleBlockDetailsList.FindAll(AddressOf Findrevnum)
                '            For Each d As PEDetails In tmplist ' should only be one of these?!
                '                If String.IsNullOrEmpty(d.TagStringVal) Then
                '                    osheet.TitleBlock.SetPromptResultText(oTextBox, "")
                '                Else
                '                    osheet.TitleBlock.SetPromptResultText(oTextBox, d.TagStringVal)
                '                End If
                '                Exit For
                '            Next
                '        ElseIf oTag = UCase("Current Issue Date") Then
                '            tmplist = TitleBlockDetailsList.FindAll(AddressOf Finddate)
                '            For Each d As PEDetails In tmplist ' should only be one of these?!
                '                If String.IsNullOrEmpty(d.TagStringVal) Then
                '                    osheet.TitleBlock.SetPromptResultText(oTextBox, "")
                '                Else
                '                    osheet.TitleBlock.SetPromptResultText(oTextBox, d.TagStringVal)
                '                End If
                '                Exit For
                '            Next
                '        End If
                '    End If
                'Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Something is broken!")
            Throw New ApplicationException("Exception Occured")
        Finally
            'a little cleanup!
            TitleBlockDetailsList.Clear()
        End Try
    End Sub

    ''' <summary>
    ''' Sets up a string Array based on the Prompted Entries in a specific BorderDefinition
    ''' </summary>
    ''' <param name="BorderDef">The BorderDefinition we want to parse</param>
    ''' <returns>Returns a String Array equal to the count of Prompted Entries in TitleBlockDef</returns>
    ''' <remarks></remarks>
    Private Shared Function ReturnCountofPromptedEntries(ByVal BorderDef As BorderDefinition) As String()
        Dim i As Integer = 0
        Dim count As String() = New String() {}
        Dim tmpstr As String = ""
        For Each otextbox As TextBox In BorderDef.Sketch.TextBoxes
            If IsPromptField(otextbox.FormattedText) Then
                tmpstr = ""
                'tmpstr = GetPromptField(otextbox.FormattedText)
                'If tmpstr = Nothing Then
                '    tmpstr = ""
                'End If
                ReDim Preserve count(count.Length)
                count(i) = tmpstr
                i += 1
            End If
        Next
        Return count
    End Function

    ''' <summary>
    ''' Sets up a string Array based on the Prompted Entries in a specific TitleBlockDefinition
    ''' </summary>
    ''' <param name="TitleBlockDef">The TitleBlockDefinition we want to parse</param>
    ''' <returns>Returns a String Array equal to the count of Prompted Entries in TitleBlockDef</returns>
    ''' <remarks></remarks>
    Private Shared Function ReturnCountofPromptedEntries(ByVal TitleBlockDef As TitleBlockDefinition) As String()
        Dim i As Integer = 0
        Dim count As String() = New String() {}
        Dim tmpstr As String = ""
        For Each otextbox As TextBox In TitleBlockDef.Sketch.TextBoxes
            If IsPromptField(otextbox.FormattedText) Then
                tmpstr = " "
                'tmpstr = GetPromptField(otextbox.FormattedText)
                'If tmpstr = Nothing Then
                '    tmpstr = ""
                'End If
                ReDim Preserve count(count.Length)
                count(i) = tmpstr
                i += 1
            End If
        Next
        Return count
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="FormattedText"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetPromptField(ByVal FormattedText As String) As String

        On Error GoTo ErrorFound

        ' Verify that this is a prompt field.
        If Left(FormattedText, 7) <> "<Prompt" Then
            GetPromptField = ""
            Exit Function
        End If
        ' Get the text that is to the right of the first ">" symbol
        ' and to the left of the last "<" symbol.
        ' Debug.Print FormattedText
        GetPromptField = Right(FormattedText, Len(FormattedText) - InStr(FormattedText, ">"))
        GetPromptField = Left(GetPromptField, InStr(GetPromptField, "<") - 1)
        ' Replace any &lt; or &gt; with < and > symbols.
        If Left(GetPromptField, 4) = "&lt;" Then
            GetPromptField = Replace(GetPromptField, "&lt;", "<")
            GetPromptField = Replace(GetPromptField, "&gt;", ">")
        End If
        'System.Windows.MessageBox.Show(GetPromptField)
        Exit Function
ErrorFound: GetPromptField = ""

    End Function

    ''' <summary>
    ''' Determines whether our Text is a prompted entry field or not.
    ''' </summary>
    ''' <param name="FormattedText"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsPromptField(ByVal FormattedText As String) As Boolean
        ' Verify that this is a prompt field.
        If Left(FormattedText, 7) <> "<Prompt" Then
            Return False
        Else
            Return True
        End If
    End Function

    ''' <summary>
    ''' Used to retrieve Prompted entry values from the borders.
    ''' </summary>
    ''' <param name="Searchstring"></param>
    ''' <param name="oSheet"></param>
    ''' <returns></returns>
    ''' <remarks>         Get the prompted text value from the title block. _
    ''' This is done by first getting the text box in the title block definition that defines the prompted text.  _
    ''' Then you can use this to get the value specified for this particular title block instance.</remarks>
    Public Function RetrievePE(ByVal Searchstring As String, ByVal oSheet As Sheet) As String
        RetrievePE = String.Empty
        Dim oDrawDoc As DrawingDocument
        oDrawDoc = ThisApplication.ActiveDocument

        Dim oBorderDef As BorderDefinition

        oBorderDef = oSheet.Border.Definition

        Dim oTextBox As TextBox
        Dim bFound As Boolean
        bFound = False
        For Each oTextBox In oBorderDef.Sketch.TextBoxes
            If GetPromptField(oTextBox.FormattedText) = Searchstring Then
                RetrievePE = oSheet.Border.GetResultText(oTextBox)
                bFound = True
                Exit For
            End If
        Next
        If (Not bFound = True) Then
            MsgBox("Specified formatted text was not found in the title block.")
        End If
        Return RetrievePE
    End Function
#End Region

#Region "AddDLtoDB"
    ''' <summary>
    ''' Checks whether the activedocument is a .idw file.
    ''' </summary>
    ''' <returns>Returns true if working with Drawing Document</returns>
    ''' <remarks></remarks>
    Public Shared Function IsAssemblyDrawing() As Boolean
        Dim datatoadd As Boolean = False
        oIDWDoc = ThisApplication.ActiveDocument
        ' need to set this here or hilarity will ensue!
        If oIDWDoc.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
            For Each osheet As Sheet In oIDWDoc.Sheets
                If UCase(osheet.Name) Like "*DL*" Or UCase(osheet.Name) Like "*DRAWING LIST*" Then
                    datatoadd = True
                    NewPartsListToUpload(osheet)
                End If
            Next
        Else
            datatoadd = False
        End If
        Return datatoadd
    End Function

    ''' <summary>
    ''' Queries parts lists on 'DL' sheets and adds the values to the database for later retrieval.
    ''' </summary>
    ''' <param name="idwsheet">The sheet we're processing!</param>
    ''' <remarks>This information could be used to link documents together thus saving us having to manually enter the "Used On" fields.</remarks>
    Private Shared Sub NewPartsListToUpload(ByVal idwsheet As Sheet)
        Using db As New DWGDetailsDataContext
            Try
                Dim oPartList As PartsList
                oPartList = idwsheet.PartsLists.Item(1)
                Dim i As Long
                For i = 1 To oPartList.PartsListRows.Count
                    Dim oRow As PartsListRow
                    oRow = oPartList.PartsListRows.Item(i)
                    Dim j As Long
                    Dim dwgnum As String = String.Empty
                    Dim dwgissue As String = String.Empty
                    Dim dwgpm As String = String.Empty
                    For j = 1 To oPartList.PartsListColumns.Count
                        Dim oCell As PartsListCell
                        oCell = oRow.Item(j)
                        Select Case oPartList.PartsListColumns.Item(j).Title
                            Case "Drawing Number"
                                dwgnum = oCell.Value
                            Case "Issue"
                                dwgissue = oCell.Value
                            Case "PM"
                                dwgpm = oCell.Value
                        End Select
                        'Debug.Print "Row: " & i & ", Column: " & oPartList.PartsListColumns.Item(j).Title & " = "; oCell.Value
                    Next
                    Dim dl As New DLDETAIL() With {
                            .ASSEMBLY = oIDWDoc.FullDocumentName, _
                            .DLNUMBER = idwsheet.Name, _
                            .DRAWINGNUMBER = dwgnum, _
                            .ISSUE = dwgissue, _
                            .PM = dwgpm
                        }
                    db.DLDETAILs.InsertOnSubmit(dl)
                Next
            Catch ex As Exception
                MessageBox.Show("There was a problem creating the new database entry: " & ex.Message)
            Finally
                db.SubmitChanges()
            End Try
        End Using
    End Sub
#End Region

End Class
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
Imports NewRibbonCmds.StandardAddInServer
Public Class ChangeStyles
    ''' <summary>
    ''' Populates the Combobox with the appropriate style names.
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub PopulateCB1()
        'get the current drawing
        Dim oIDWDoc As DrawingDocument
        oIDWDoc = ThisApplication.ActiveDocument

        'get the current activestandardstyle
        'Dim oDimStyle As Style
        'oDimStyle = oIDWDoc.StylesManager.ActiveStandardStyle()
        'harvest the current available styles and add them to the drop-down box we created.
        Dim oIDWStyles As DrawingStylesManager
        oIDWStyles = oIDWDoc.StylesManager
        For i = 1 To oIDWStyles.DimensionStyles.Count
            If _
                oIDWStyles.DimensionStyles.Item(i).Name Like "*CH2M*" Or _
                oIDWStyles.DimensionStyles.Item(i).Name Like "*FRAME*" Then
                m_combobox1.AddItem(oIDWStyles.DimensionStyles.Item(i).Name)
            End If
        Next
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="osheets"></param>
    ''' <remarks></remarks>
    Public Shared Sub GetLayerSettings(ByVal osheets As Sheets)
        'Layers
        DetailsList = New List(Of PEDetails)
        For Each osheet As Sheet In osheets
            Dim oBorderDef As BorderDefinition = osheet.Border.Definition
            Dim oTag As String
            ' attribute tag
            Dim oString As String
            ' attribute string
            Dim oFound As Boolean = False
            Dim layervis As Boolean = False
            For Each oTextBox As TextBox In oBorderDef.Sketch.TextBoxes
                If IsPromptField(oTextBox.FormattedText) Then
                    oTag = GetPromptField(oTextBox.FormattedText)
                    oString = osheet.Border.GetResultText(oTextBox)
                    Select Case oTag
                        Case UCase("<Tolerance 1>")
                            IIf(oString = "", layervis = False, layervis = False)
                            DetailsList.Add(New PEDetails(osheet.Name, oTag, oString, "CH2M - No Tolerances", layervis))
                        Case UCase("<Protective Finish>")
                            IIf(oString = "", layervis = False, layervis = False)
                            DetailsList.Add(New PEDetails(osheet.Name, oTag, oString, "CH2M - No Protective Finish", layervis))
                        Case UCase("<Surface Texture>")
                            IIf(oString = "", layervis = False, layervis = False)
                            DetailsList.Add(New PEDetails(osheet.Name, oTag, oString, "CH2M - No Surface Texture", layervis))
                        Case UCase("Tolerance 1")
                            IIf(oString = "", layervis = False, layervis = False)
                            DetailsList.Add(New PEDetails(osheet.Name, oTag, oString, "FRAME - No Tolerances", layervis))
                        Case UCase("Protective Finish")
                            IIf(oString = "", layervis = False, layervis = False)
                            DetailsList.Add(New PEDetails(osheet.Name, oTag, oString, "FRAME - No Protective Finish", layervis))
                        Case UCase("surface texture")
                            IIf(oString = "", layervis = False, layervis = False)
                            DetailsList.Add(New PEDetails(osheet.Name, oTag, oString, "FRAME - No Surface Texture",layervis))
                    End Select
                End If
            Next
        Next
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub ChangeStyle()
        ' Updates:
        '0.0.0.1: Updated to work with rotated text
        '0.0.0.2: Updated to work with CustomTables.
        '0.0.0.3: Removed AssignStyles section to it's own sub.
        '0.0.0.4: Attempting to simplify the layer change.
        '0.0.0.5: Moved capturing of "layer state" settings to GetLayerSettings sub.
        '0.0.0.6: Redefined List usage according to this page: http://bit.ly/98bdie
        '0.0.0.7: CH2M Hole Table Style added to library.
        '0.0.1.1: Modified to work with the new "layer only" templates.
        Dim oSheets As Sheets = oIDWDoc.Sheets
        Dim UnitsType As UnitsTypeEnum = oIDWDoc.UnitsOfMeasure.LengthUnits
        ' store the current length units value
        oIDWDoc.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits
        ' set the units to millimeters so we get the correct values?
        GetLayerSettings(oSheets)
        Dim LastRefKeyStr As String = ""
        Dim tr As Transaction
        Dim RefkeyManager As ReferenceKeyManager = oIDWDoc.ReferenceKeyManager
        For Each osheet As Sheet In oSheets
            tr = ThisApplication.TransactionManager.StartTransaction(oIDWDoc, "Process Layers")
            Try
                If m_combobox1.Text Like "*CH2M*" Then
                    For Each Layer As Layer In oIDWDoc.StylesManager.Layers
                        If Layer.Name Like "*FRAME*" Then
                            Layer.Visible = False
                        ElseIf Layer.Name = "CH2M Frame Text" Or Layer.Name = "CH2M Frame" Then
                            Layer.Visible = True
                        ElseIf _
                            Layer.Name = "CH2M - No Surface Texture" Or Layer.Name = "CH2M - No Tolerances" Or _
                            Layer.Name = "CH2M - No Protective Finish" Then
                            For Each detail As PEDetails In DetailsList
                                If detail.Sheetname = osheet.Name Then
                                    If detail.LayerName = Layer.Name Then
                                        If detail.LayerVis = False Then
                                            Layer.Visible = False
                                        ElseIf detail.LayerVis = True Then
                                            Layer.Visible = True
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    Next
                ElseIf m_combobox1.Text Like "*FRAME*" Then
                    For Each Layer As Layer In oIDWDoc.StylesManager.Layers
                        If Layer.Name Like "*CH2M*" Then
                            Layer.Visible = False
                        ElseIf Layer.Name = "FRAME TEXT (THICK)" Or Layer.Name = "FRAME" Then
                            Layer.Visible = True
                        ElseIf _
                            Layer.Name = "FRAME - No Tolerances" Or Layer.Name = "FRAME - No Protective Finish" Or _
                            Layer.Name = "FRAME - No Surface Texture" Then
                            For Each detail As PEDetails In DetailsList
                                If detail.Sheetname = osheet.Name Then
                                    If detail.LayerName = Layer.Name Then
                                        If detail.LayerVis = False Then
                                            Layer.Visible = False
                                        ElseIf detail.LayerVis = True Then
                                            Layer.Visible = True
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
            Catch ex As Exception
                tr.Abort()
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "An Error has occured during the layer switch over!")
            End Try
            tr.End()
        Next
        oIDWDoc.UnitsOfMeasure.LengthUnits = UnitsType
        oIDWDoc.Update()
    End Sub

End Class

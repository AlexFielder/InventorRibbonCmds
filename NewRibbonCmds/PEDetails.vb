Public Class PEDetails
    Public Sheetname As String
    Public TagName As String
    Public TagStringVal As String
    Public LayerName As String
    Public LayerVis As Boolean
    ''' <summary>
    ''' Stores our Prompted Entry (PE) details
    ''' </summary>
    ''' <param name="m_Tagname">The Tag Name of the PE</param>
    ''' <param name="m_TagStringVal">The Tag String Value of the PE</param>
    ''' <param name="m_Sheetname">The Sheet we found it on</param>
    ''' <param name="m_Layername">The Layer we found it on</param>
    ''' <param name="m_Layervis">The Visibility of Layername</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal m_Tagname As String, _
                   ByVal m_TagStringVal As String, _
                   Optional ByVal m_Sheetname As String = "", _
                   Optional ByVal m_Layername As String = "", _
                   Optional ByVal m_Layervis As Boolean = False)
        Sheetname = m_Sheetname
        TagStringVal = m_TagStringVal
        TagName = m_Tagname
        LayerName = m_Layername
        LayerVis = m_Layervis
    End Sub


End Class

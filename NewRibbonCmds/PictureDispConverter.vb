' Microsoft Office Outlook 2007 Add-in Sample Code
'
' THIS CODE AND INFORMATION ARE PROVIDED AS IS WITHOUT WARRANTY OF ANY
' KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
' IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' 
Imports System
Imports System.Drawing
Imports System.Runtime.InteropServices

Public Module PictureDispConverter

    'IPictureDisp guid
    Public iPictureDispGuid As Guid = GetType(stdole.IPictureDisp).GUID

    'Converts an Icon into a IPictureDisp
    Public Function ToIPictureDisp(ByVal ico As Icon) As stdole.IPictureDisp

        Dim pictIcon As New PICTDESC.Icon(ico)
        Return PictureDispConverter.OleCreatePictureIndirect(pictIcon, iPictureDispGuid, True)
    End Function

    'Converts an image into a IPictureDisp
    Public Function ToIPictureDisp(ByVal picture As Image) As stdole.IPictureDisp

        Dim bm As Bitmap
        If TypeOf picture Is Bitmap Then
            bm = picture
        Else
            bm = New Bitmap(picture)
        End If
        Dim pictBit As New PICTDESC.Bitmap(bm)
        Return PictureDispConverter.OleCreatePictureIndirect(pictBit, iPictureDispGuid, True)
    End Function


    <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)> _
    Private Function OleCreatePictureIndirect(<MarshalAs(UnmanagedType.AsAny)> ByVal picdesc As Object, ByRef iid As Guid, ByVal fOwn As Boolean) As stdole.IPictureDisp
    End Function

    Private ReadOnly hCollector As New HandleCollector("Icon handles", 1000)


    'PICTDESC is a union in native, so we'll just
    'define different ones for the different types
    'the "unused" fields are there to make it the right
    'size, since the struct in native is as big as the biggest
    'union.
    Private Class PICTDESC

        'Picture Types
        Public Const PICTYPE_UNINITIALIZED As Short = -1
        Public Const PICTYPE_NONE As Short = 0
        Public Const PICTYPE_BITMAP As Short = 1
        Public Const PICTYPE_METAFILE As Short = 2
        Public Const PICTYPE_ICON As Short = 3
        Public Const PICTYPE_ENHMETAFILE As Short = 4

        <StructLayout(LayoutKind.Sequential)> _
        Public Class Icon

            Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Icon))
            Friend picType As Integer = PICTDESC.PICTYPE_ICON
            Friend hicon As IntPtr = IntPtr.Zero
            Friend unused1 As Integer = 0
            Friend unused2 As Integer = 0

            Friend Sub New(ByVal icon As System.Drawing.Icon)
                Me.hicon = icon.ToBitmap().GetHicon()
            End Sub

        End Class

        <StructLayout(LayoutKind.Sequential)> _
        Public Class Bitmap

            Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Bitmap))
            Friend picType As Integer = PICTDESC.PICTYPE_BITMAP
            Friend hbitmap As IntPtr = IntPtr.Zero
            Friend hpal As IntPtr = IntPtr.Zero
            Friend unused As Integer = 0

            Friend Sub New(ByVal bitmap As System.Drawing.Bitmap)
                Me.hbitmap = bitmap.GetHbitmap()
            End Sub
        End Class

    End Class

End Module

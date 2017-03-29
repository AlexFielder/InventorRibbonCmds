Imports System.Windows.Forms
Module VariableRead
    Public FunctionName As String
    Public Function readLine(ByVal FilePath As String) As String()
        FunctionName = "readline"
        Dim counter As Integer = 0
        Dim line As String = String.Empty
        Dim variables As String() = New String(20) {} ' We have to assign a value to the string array otherwise we get an error!
        Try
            Dim file As New System.IO.StreamReader(FilePath)
            While (InlineAssignHelper(line, file.ReadLine())) IsNot Nothing
                variables(counter) = line
                counter += 1
            End While
            file.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Uh oh, something is broken with: " & FunctionName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return variables
    End Function
    Private Function InlineAssignHelper(Of T)(ByRef target As T, ByVal value As T) As T
        target = value
        Return value
    End Function

End Module

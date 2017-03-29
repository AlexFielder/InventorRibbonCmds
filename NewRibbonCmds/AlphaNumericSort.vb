Imports System.Collections
Public Module AlphaNumericSort
    Public Function Main(ByVal revs As String()) As String()
        'Dim highways As String() = New String() {"100F", "50F", "SR100", "SR9"}
        '
        ' We want to sort a string[] array called highways in an
        ' alphanumeric way. Call the static Array.Sort method.
        '
        Array.Sort(revs, New AlphanumComparatorFast())
        '
        ' Display the results
        '
        Main = revs
        For Each h As String In revs
            Console.WriteLine(h)
        Next
        Return Main
    End Function
End Module

Public Class AlphanumComparatorFast
    Implements IComparer
    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
        Dim s1 As String = TryCast(x, String)
        If s1 Is Nothing Then
            Return 0
        End If
        Dim s2 As String = TryCast(y, String)
        If s2 Is Nothing Then
            Return 0
        End If

        Dim len1 As Integer = s1.Length
        Dim len2 As Integer = s2.Length
        Dim marker1 As Integer = 0
        Dim marker2 As Integer = 0

        ' Walk through two the strings with two markers.
        While marker1 < len1 AndAlso marker2 < len2
            Dim ch1 As Char = s1(marker1)
            Dim ch2 As Char = s2(marker2)

            ' Some buffers we can build up characters in for each chunk.
            Dim space1 As Char() = New Char(len1 - 1) {}
            Dim loc1 As Integer = 0
            Dim space2 As Char() = New Char(len2 - 1) {}
            Dim loc2 As Integer = 0

            ' Walk through all following characters that are digits or
            ' characters in BOTH strings starting at the appropriate marker.
            ' Collect char arrays.
            Do
                space1(System.Math.Max(System.Threading.Interlocked.Increment(loc1), loc1 - 1)) = ch1
                marker1 += 1

                If marker1 < len1 Then
                    ch1 = s1(marker1)
                Else
                    Exit Do
                End If
            Loop While Char.IsDigit(ch1) = Char.IsDigit(space1(0))

            Do
                space2(System.Math.Max(System.Threading.Interlocked.Increment(loc2), loc2 - 1)) = ch2
                marker2 += 1

                If marker2 < len2 Then
                    ch2 = s2(marker2)
                Else
                    Exit Do
                End If
            Loop While Char.IsDigit(ch2) = Char.IsDigit(space2(0))

            ' If we have collected numbers, compare them numerically.
            ' Otherwise, if we have strings, compare them alphabetically.
            Dim str1 As New String(space1)
            Dim str2 As New String(space2)

            Dim result As Integer

            If Char.IsDigit(space1(0)) AndAlso Char.IsDigit(space2(0)) Then
                Dim thisNumericChunk As Integer = Integer.Parse(str1)
                Dim thatNumericChunk As Integer = Integer.Parse(str2)
                result = thisNumericChunk.CompareTo(thatNumericChunk)
            Else
                result = str1.CompareTo(str2)
            End If

            If result <> 0 Then
                Return result
            End If
        End While
        Return len1 - len2
    End Function
End Class


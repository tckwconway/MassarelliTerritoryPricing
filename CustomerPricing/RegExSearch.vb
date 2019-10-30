Imports System.Text.RegularExpressions

Module RegExSearch
    Public Enum PadPosition As Integer
        PadStart = 1
        PadEnd = 2
    End Enum
    Public Function RegExStripCharacters(ByVal input As String) As String
        Dim output As String

        output = Regex.Replace(input, "[a-zA-Z]", "")

        Return output
    End Function
    Public Function PadString(input As String, padChars As String, padPos As PadPosition, outputLen As Integer, Optional startPos As Integer = 0) As String
        Dim output As String

        If padPos = PadPosition.PadStart Then
            Dim s As String = (padChars & input)
            Dim l As Integer = s.Length
            output = s.Substring(l - outputLen, outputLen)
        Else
            Dim s As String = (input & padChars)
            Dim l As Integer = s.Length
            output = s.Substring(startPos, outputLen)
        End If

        Return output
    End Function
End Module

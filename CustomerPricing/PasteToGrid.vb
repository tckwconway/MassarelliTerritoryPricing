Module PasteToGrid

    Public Sub PasteClipboard(dgv As DataGridView, cl As Integer, rw As Integer)

        Dim _pasteText As String = Clipboard.GetText()
        Dim _new As String = _pasteText.Replace(vbCr & vbLf, "R")
        'replace escape chars
        _new = _new.Replace(vbTab, "C")

        Dim _splitter As String() = {"R"}
        'splitter for rows
        Dim _splitterC As String() = {"C"}

        Dim _rows As String() = _new.Split(_splitter, StringSplitOptions.RemoveEmptyEntries)
        Dim _r As Integer = _rows.Length

        Dim _cells As String() = _rows(0).Split(_splitterC, StringSplitOptions.None)
        Dim _c As Integer = _cells.Length

        For row As Integer = 0 To _r - 1
            _cells = _rows(row).Split(_splitterC, StringSplitOptions.None)
            'replace escape chars
            For col As Integer = 0 To _c - 1
                dgv.Rows(row + rw).Cells(col + cl).Value = _cells(col)
            Next
        Next

    End Sub
    Public Sub ClearSelection(dgv As DataGridView)

        Dim emptycell_Decimal As Decimal = 0
        Dim emptycell_String As String = ""
        Dim emptycell_Integer As Integer = 0
        Dim emptycell_Date As Date = #1/1/1900#

        With dgv
            For Each cel As DataGridViewCell In dgv.SelectedCells
                Select Case cel.ValueType.FullName
                    Case "System.String"
                        cel.Value = emptycell_String
                    Case "System.Decimal"
                        cel.Value = emptycell_Decimal
                    Case "System.Integer"
                        cel.Value = emptycell_Integer
                    Case "System.Date"
                        cel.Value = emptycell_Date
                End Select

            Next
        End With

    End Sub

End Module

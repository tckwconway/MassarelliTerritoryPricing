Imports System.Text


Module ExportGridInExcel

    Public Function GetSelectedFieldNames(ByVal dg As DataGridView) As String()
        Dim fieldnames() As String
        Dim fldnms As New StringBuilder
        Dim col As DataGridViewColumn
        Try
            For Each col In dg.Columns
                If col.Visible = True Then
                    Try
                        If fldnms.ToString.Length = 0 Then
                            fldnms.Append(col.Name.ToString)
                        Else
                            fldnms.Append(",")
                            fldnms.Append(col.Name.ToString)
                        End If

                    Catch ex As Exception
                        MsgBox("Error occurred at GetSelectedFieldNames Function")
                        MsgBox(ex.Message)
                    End Try

                End If

            Next


        Catch ex As Exception

        End Try
        fieldnames = Split(fldnms.ToString, ",")
        Return fieldnames

    End Function

    Public Function GetFixedFieldNames(FieldList As String) As String()
        Dim fieldnames() As String
        fieldnames = Split(FieldList.ToString, ",")

        Return fieldnames

    End Function
End Module

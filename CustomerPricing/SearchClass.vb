Public Class SearchClass

    Private mDataGridViewRow As Integer
    Public Property DataGridViewRow() As Integer
        Get
            Return mDataGridViewRow
        End Get
        Set(ByVal value As Integer)
            mDataGridViewRow = value
        End Set
    End Property

    Private mDataGridViewCol As Integer
    Public Property DataGridViewCol() As Integer
        Get
            Return mDataGridViewCol
        End Get
        Set(ByVal value As Integer)
            mDataGridViewCol = value
        End Set
    End Property

    Private mSearchPropertyName As String
    Public Property SearchPropertyName() As String
        Get
            Return mSearchPropertyName
        End Get
        Set(ByVal value As String)
            mSearchPropertyName = value
        End Set
    End Property

    Private mSearchPropertyValue As String
    Public Property SearchPropertyValue() As String
        Get
            Return mSearchPropertyValue
        End Get
        Set(ByVal value As String)
            mSearchPropertyValue = value
        End Set
    End Property

End Class

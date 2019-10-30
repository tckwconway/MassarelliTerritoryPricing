Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.ComponentModel
Class SortComparer(Of T)
    Implements IComparer(Of T)
    Private m_SortCollection As ListSortDescriptionCollection = Nothing
    Private m_PropDesc As PropertyDescriptor = Nothing
    Private m_Direction As ListSortDirection = ListSortDirection.Ascending
    Public Sub New(ByVal propDesc As PropertyDescriptor, ByVal direction As ListSortDirection)
        m_PropDesc = propDesc
        m_Direction = direction
    End Sub
    Public Sub New(ByVal sortCollection As ListSortDescriptionCollection)
        m_SortCollection = sortCollection
    End Sub
#Region "IComparer<T> Members"
    Function Compare(ByVal x As T, ByVal y As T) As Integer Implements IComparer(Of T).Compare
        If m_PropDesc IsNot Nothing Then
            ' Simple sort 
            Dim xValue As Object = m_PropDesc.GetValue(x)
            Dim yValue As Object = m_PropDesc.GetValue(y)
            Return CompareValues(xValue, yValue, m_Direction)
        Else
            If m_SortCollection IsNot Nothing AndAlso m_SortCollection.Count > 0 Then
                Return RecursiveCompareInternal(x, y, 0)
            Else
                Return 0
            End If
        End If
    End Function
#End Region
    Private Function CompareValues(ByVal xValue As Object, ByVal yValue As Object, ByVal direction As ListSortDirection) As Integer
        Dim retValue As Integer = 0
        If (xValue Is Nothing And yValue Is Nothing) Then Exit Function
        If TypeOf xValue Is IComparable Then
            ' Can ask the x value
            retValue = (DirectCast(xValue, IComparable)).CompareTo(yValue)
        Else
            If TypeOf yValue Is IComparable Then
                'Can ask the y value
                retValue = (DirectCast(yValue, IComparable)).CompareTo(xValue)
            Else
                If Not xValue.Equals(yValue) Then
                    ' not comparable, compare String representations
                    retValue = xValue.ToString().CompareTo(yValue.ToString())
                End If
            End If
        End If
        If direction = ListSortDirection.Ascending Then
            Return retValue
        Else
            Return retValue * -1
        End If
    End Function
    Private Function RecursiveCompareInternal(ByVal x As T, ByVal y As T, ByVal index As Integer) As Integer
        If index >= m_SortCollection.Count Then
            Return 0
        End If
        ' termination condition
        Dim listSortDesc As ListSortDescription = m_SortCollection(index)
        Dim xValue As Object = listSortDesc.PropertyDescriptor.GetValue(x)
        Dim yValue As Object = listSortDesc.PropertyDescriptor.GetValue(y)
        Dim retValue As Integer = CompareValues(xValue, yValue, listSortDesc.SortDirection)
        If retValue = 0 Then
            Return RecursiveCompareInternal(x, y, System.Threading.Interlocked.Increment(index))
        Else
            Return retValue
        End If
    End Function
End Class

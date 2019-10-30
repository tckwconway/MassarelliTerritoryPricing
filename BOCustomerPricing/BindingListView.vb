Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Text
Public Class BindingListView(Of T)
    Inherits BindingList(Of T)
    Implements IBindingListView
    Implements IRaiseItemChangedEvents
    Private m_Sorted As Boolean = False
    Private m_Filtered As Boolean = False
    Private m_FilterString As String = Nothing
    Private m_SortDirection As ListSortDirection = ListSortDirection.Ascending
    Private m_SortProperty As PropertyDescriptor = Nothing
    Private m_SortDescriptions As ListSortDescriptionCollection = New ListSortDescriptionCollection()
    Private m_OriginalCollection As List(Of T) = New List(Of T)()
    Public Sub New()

        MyBase.New()
    End Sub
    Public Sub New(ByVal list As List(Of T))
        MyBase.New(list)
    End Sub
    Protected Overloads Overrides ReadOnly Property SupportsSearchingCore() As Boolean
        Get
            Return True
        End Get
    End Property
    Protected Overloads Overrides Function FindCore(ByVal [property] As PropertyDescriptor, ByVal key As Object) As Integer
        For i As Integer = 0 To Count - 1
            ' Simple iteration:
            Dim item As T = Me(i)
            If [property].GetValue(item).Equals(key) Then
                Return i
            End If
        Next
        Return -1
   
    End Function
    Protected Overloads Overrides ReadOnly Property SupportsSortingCore() As Boolean
        Get
            Return True
        End Get
    End Property
    Protected Overloads Overrides ReadOnly Property IsSortedCore() As Boolean
        Get
            Return m_Sorted
        End Get
    End Property
    Protected Overloads Overrides ReadOnly Property SortDirectionCore() As ListSortDirection
        Get
            Return m_SortDirection
        End Get
    End Property
    Protected Overloads Overrides ReadOnly Property SortPropertyCore() As PropertyDescriptor
        Get
            Return m_SortProperty
        End Get
    End Property
    Protected Overloads Overrides Sub ApplySortCore(ByVal [property] As PropertyDescriptor, ByVal direction As ListSortDirection)
        m_SortDirection = direction
        m_SortProperty = [property]
        Dim comparer As SortComparer(Of T) = New SortComparer(Of T)([property], direction)
        ApplySortInternal(comparer)
    End Sub
    Private Sub ApplySortInternal(ByVal comparer As SortComparer(Of T))
        If m_OriginalCollection.Count = 0 Then
            m_OriginalCollection.AddRange(Me)
        End If
        Dim listRef As List(Of T) = TryCast(Me.Items, List(Of T))
        If listRef Is Nothing Then
            Return
        End If
        listRef.Sort(comparer)
        m_Sorted = True
        OnListChanged(New ListChangedEventArgs(ListChangedType.Reset, -1))
    End Sub
    Protected Overloads Overrides Sub RemoveSortCore()
        If Not m_Sorted Then
            Return
        End If
        Clear()
        For Each item As T In m_OriginalCollection
            Add(item)
        Next
        m_OriginalCollection.Clear()
        m_SortProperty = Nothing
        m_SortDescriptions = Nothing
        m_Sorted = False
    End Sub
#Region "IBindingListView Members"
    Sub ApplySort(ByVal sorts As ListSortDescriptionCollection) Implements IBindingListView.ApplySort
        m_SortProperty = Nothing
        m_SortDescriptions = sorts
        Dim comparer As SortComparer(Of T) = New SortComparer(Of T)(sorts)
        ApplySortInternal(comparer)
    End Sub
    Property Filter() As String Implements IBindingListView.Filter
        Get
            Return m_FilterString
        End Get
        Set(ByVal value As String)
            m_FilterString = value
            m_Filtered = True
            If value.Contains("|") Then
                UpdateFilterContains()
            Else
                UpdateFilter()
            End If

        End Set
    End Property
   
    Sub RemoveFilter() Implements IBindingListView.RemoveFilter
        If Not m_Filtered Then
            Return
        End If
        m_FilterString = Nothing
        m_Filtered = False
        m_Sorted = False
        m_SortDescriptions = Nothing
        m_SortProperty = Nothing
        Clear()
        For Each item As T In m_OriginalCollection
            Add(item)
        Next
        m_OriginalCollection.Clear()
    End Sub
    ReadOnly Property SortDescriptions() As ListSortDescriptionCollection Implements IBindingListView.SortDescriptions
        Get
            Return m_SortDescriptions
        End Get
    End Property
    ReadOnly Property SupportsAdvancedSorting() As Boolean Implements IBindingListView.SupportsAdvancedSorting
        Get
            Return True
        End Get
    End Property
    ReadOnly Property SupportsFiltering() As Boolean Implements IBindingListView.SupportsFiltering
        Get
            Return True
        End Get
    End Property
#End Region
    Protected Overridable Sub UpdateFilter()
        Dim equalsPos As Integer = m_FilterString.IndexOf("="c)
        ' Get property name
        Dim propName As String = m_FilterString.Substring(0, equalsPos).Trim()
        ' Get Filter criteria
        Dim criteria As String = m_FilterString.Substring(equalsPos + 1, m_FilterString.Length - equalsPos - 1).Trim()
        criteria = criteria.Substring(0, criteria.Length)
        ' string leading and trailing quotes
        ' Get a property descriptor for the filter property
        Dim propDesc As PropertyDescriptor = TypeDescriptor.GetProperties(GetType(T))(propName)
        If m_OriginalCollection.Count = 0 Then
            m_OriginalCollection.AddRange(Me)
        End If
        Dim currentCollection As List(Of T) = New List(Of T)(Me)
        Clear()
        For Each item As T In currentCollection
            Dim value As Object = propDesc.GetValue(item)

            If criteria.Contains("'" & value.ToString().Trim & "'") Then
                'If criteria.ToString.Trim = value.ToString.Trim Then
                Add(item)
            End If
        Next
    End Sub

    Protected Overridable Sub UpdateFilterContains()
        Dim equalsPos As Integer = m_FilterString.IndexOf("="c)
        ' Get property name
        Dim propName As String = m_FilterString.Substring(0, equalsPos).Trim()
        ' Get Filter criteria
        Dim criteria As String = m_FilterString.Substring(equalsPos + 1, m_FilterString.Length - equalsPos - 1).Trim()
        criteria = criteria.Substring(0, criteria.Length)
        ' string leading and trailing quotes
        ' Get a property descriptor for the filter property
        Dim propDesc As PropertyDescriptor = TypeDescriptor.GetProperties(GetType(T))(propName)
        If m_OriginalCollection.Count = 0 Then
            m_OriginalCollection.AddRange(Me)
        End If
        Dim currentCollection As List(Of T) = New List(Of T)(Me)
        Clear()
        For Each item As T In currentCollection
            Dim value As Object = propDesc.GetValue(item)
            If criteria.Contains("'" & value.ToString().Trim & "'") Then
                'If criteria.ToString.Trim = value.ToString.Trim Then
                Add(item)
            End If
        Next
    End Sub

#Region "IBindingList overrides"
   
    Private Function CheckReadOnly() As Boolean
        If m_Sorted OrElse m_Filtered Then
            Return False
        Else
            Return True
        End If
    End Function
#End Region
    Protected Overloads Overrides Sub InsertItem(ByVal index As Integer, ByVal item As T)
        For Each propDesc As PropertyDescriptor In TypeDescriptor.GetProperties(item)
            If propDesc.SupportsChangeEvents Then
                propDesc.AddValueChanged(item, AddressOf OnItemChanged)
            End If
        Next
        MyBase.InsertItem(index, item)
    End Sub
    Protected Overloads Overrides Sub RemoveItem(ByVal index As Integer)
        Dim item As T = Items(index)
        Dim propDescs As PropertyDescriptorCollection = TypeDescriptor.GetProperties(item)
        For Each propDesc As PropertyDescriptor In propDescs
            If propDesc.SupportsChangeEvents Then
                propDesc.RemoveValueChanged(item, AddressOf OnItemChanged)
            End If
        Next
        MyBase.RemoveItem(index)
    End Sub
    Sub OnItemChanged(ByVal sender As Object, ByVal args As EventArgs)
        Dim index As Integer = Items.IndexOf(DirectCast(sender, T))
        OnListChanged(New ListChangedEventArgs(ListChangedType.ItemChanged, index))
    End Sub
    
End Class

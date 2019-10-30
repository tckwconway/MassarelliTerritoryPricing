Imports System.ComponentModel
Imports System.Data.DataRowView
Imports System.Data
Imports System


Public Class ItemLookupObj
    Inherits BaseClass
    Implements INotifyPropertyChanged
    Implements IEditableObject


#Region "   Methods   "
    Public Shared Function NewItemLookup(ByVal ItemNo As String, ByVal ItemDesc As String)
        Dim itmlookup As New ItemLookupObj
        itmlookup.Item_itmLookup = ItemNo
        itmlookup.ItemDesc_itmLookup = ItemDesc
        Return itmlookup

    End Function
    Public Shared Function GetItemLookup(ByVal ItemNo As String, ByVal ItemDesc As String)
        Dim itmlookup As New ItemLookupObj
        itmlookup.Item_itmLookup = ItemNo
        itmlookup.ItemDesc_itmLookup = ItemDesc
        Return itmlookup

    End Function

#End Region

#Region "   Properties   "


    Private mItem_itmLookup As String
    Public Property Item_itmLookup() As String
        Get
            Return mItem_itmLookup
        End Get
        Set(ByVal value As String)
            mItem_itmLookup = value
        End Set
    End Property

    Private mItemDesc_itmLookup As String
    Public Property ItemDesc_itmLookup() As String
        Get
            Return mItemDesc_itmLookup
        End Get
        Set(ByVal value As String)
            mItemDesc_itmLookup = value
        End Set
    End Property




#End Region
    Public Event PropertyChanged(ByVal sender As Object, ByVal e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Public Sub BeginEdit() Implements System.ComponentModel.IEditableObject.BeginEdit

    End Sub

    Public Sub CancelEdit() Implements System.ComponentModel.IEditableObject.CancelEdit

    End Sub

    Public Sub EndEdit() Implements System.ComponentModel.IEditableObject.EndEdit

    End Sub
End Class

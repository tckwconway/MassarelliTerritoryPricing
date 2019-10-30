Imports System.ComponentModel
Imports System.Data.DataRowView
Imports System.Data
Imports System


Public Class CategoryLookupObj
    Inherits BaseClass
    Implements INotifyPropertyChanged
    Implements IEditableObject


#Region "   Methods   "
    Public Shared Function NewCategoryLookup(ByVal CategoryNo As String, ByVal CategoryDesc As String)
        Dim catlookup As New CategoryLookupObj
        catlookup.Category_catlookup = CategoryNo
        catlookup.CategoryDesc_catlookup = CategoryDesc
        Return catlookup

    End Function
    Public Shared Function GetCategoryLookup(ByVal CategoryNo As String, ByVal CategoryDesc As String)
        Dim catlookup As New CategoryLookupObj
        catlookup.Category_catlookup = CategoryNo
        catlookup.CategoryDesc_catlookup = CategoryDesc
        Return catlookup

    End Function

#End Region

#Region "   Properties   "


    Private mCategory_catlookup As String
    Public Property Category_catLookup() As String
        Get
            Return mCategory_catlookup
        End Get
        Set(ByVal value As String)
            mCategory_catlookup = value
        End Set
    End Property

    Private mCategoryDesc_catlookup As String
    Public Property CategoryDesc_catlookup() As String
        Get
            Return mCategoryDesc_catlookup
        End Get
        Set(ByVal value As String)
            mCategoryDesc_catlookup = value
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






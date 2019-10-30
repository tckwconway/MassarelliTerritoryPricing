Imports System.ComponentModel
Imports System.Data.DataRowView
Imports System.String
Imports System.Data.SqlClient
Imports System.Data

'Price list object is used with the Excel Export (Item Pricing Object is used with the main application datagrid)
Public Class PriceListObj

    Inherits BaseClass
    Implements INotifyPropertyChanged
    Implements IEditableObject
#Region "   Methods   "
    Public Shared Function NewPriceListObj(ByVal ItemNo As String, ByVal PageNo As String, ByVal ItemDescription As String,
    ByVal ItemWeight As Decimal, ByVal OriginalPriceColor As Decimal, ByVal PriceFinish As Decimal, ByVal OriginalPriceDetailStain As Decimal, TerritoryCode As String) As PriceListObj
        Dim prclst As New PriceListObj
        prclst.ItemNo = ItemNo
        prclst.PageNo = PageNo
        prclst.ItemDescription = ItemDescription
        prclst.ItemWeight = ItemWeight
        prclst.OriginalPriceColor = OriginalPriceColor
        prclst.PriceFinish = PriceFinish
        prclst.OriginalPriceDetailStain = OriginalPriceDetailStain
        prclst.TerritoryCode = TerritoryCode
        Return prclst
    End Function
    Public Shared Function GetPriceListObj(ByVal ItemNo As String, ByVal PageNo As String, ByVal ItemDescription As String,
    ByVal ItemWeight As Decimal, ByVal OriginalPriceColor As Decimal, ByVal PriceFinish As Decimal, ByVal OriginalPriceDetailStain As Decimal, TerritoryCode As String) As PriceListObj
        Dim prclst As New PriceListObj
        prclst.ItemNo = ItemNo
        prclst.PageNo = PageNo
        prclst.ItemDescription = ItemDescription
        prclst.ItemWeight = ItemWeight
        prclst.OriginalPriceColor = OriginalPriceColor
        prclst.PriceFinish = PriceFinish
        prclst.OriginalPriceDetailStain = OriginalPriceDetailStain
        prclst.TerritoryCode = TerritoryCode
        Return prclst
    End Function

#End Region

#Region "   Properties   "

    Public Shared Function NewPricingObj(ByVal TerCode As String, ByVal ItemNo As String, ByVal ItemDesc As String,
        ByVal ProdCat As String, ByVal ProdCatDesc As String, ByVal OriginalPriceColor As Decimal, ByVal OriginalPriceRococo As Decimal,
        ByVal OriginalPriceDetailStain As Decimal, ByVal ItemLocPriceColor As Decimal,
        ByVal ItemLocPriceRococo As Decimal, ByVal ItemLocPriceDetailStain As Decimal, ByVal LastDate As Date, ByVal ItemWeight As Decimal, TerritoryCode As String) As ItemPricingObj
        Dim prc As New ItemPricingObj
        prc.TerCode = TerCode
        prc.ItemNo = ItemNo
        prc.ItemDesc = ItemDesc
        prc.ProdCat = ProdCat
        prc.ProdCatDesc = ProdCatDesc
        prc.OriginalPriceColor = OriginalPriceColor
        prc.OriginalPriceRococo = OriginalPriceRococo
        prc.OriginalPriceDetailStain = OriginalPriceDetailStain
        prc.ItemLocPriceColor = ItemLocPriceColor
        prc.ItemLocPriceRococo = ItemLocPriceRococo
        prc.ItemLocPriceDetailStain = ItemLocPriceDetailStain
        prc.Selected = False
        prc.LastDate = LastDate
        prc.ItemWeight = ItemWeight
        prc.TerCode = TerritoryCode
        Return prc
    End Function




    Private mItemNo As String
    Public Property ItemNo() As String
        Get
            Return mItemNo
        End Get
        Set(ByVal value As String)
            mItemNo = value
        End Set
    End Property

    Private mPageNo As String
    Public Property PageNo() As String
        Get
            Return mPageNo
        End Get
        Set(ByVal value As String)
            mPageNo = value
        End Set
    End Property

    Private mItemDescription As String
    Public Property ItemDescription() As String
        Get
            Return mItemDescription
        End Get
        Set(ByVal value As String)
            mItemDescription = value
        End Set
    End Property

    Private mItemWeight As Decimal
    Public Property ItemWeight() As Decimal
        Get
            Return mItemWeight
        End Get
        Set(ByVal value As Decimal)
            mItemWeight = value
        End Set
    End Property

    Private mOriginalPriceColor As Decimal
    Public Property OriginalPriceColor() As Decimal
        Get
            Return mOriginalPriceColor
        End Get
        Set(ByVal value As Decimal)
            mOriginalPriceColor = value
        End Set
    End Property

    Private mPriceFinish As Decimal
    Public Property PriceFinish() As Decimal
        Get
            Return mPriceFinish
        End Get
        Set(ByVal value As Decimal)
            mPriceFinish = value
        End Set
    End Property

    Private mOriginalPriceDetailStain As Decimal
    Public Property OriginalPriceDetailStain() As Decimal
        Get
            Return mOriginalPriceDetailStain
        End Get
        Set(ByVal value As Decimal)
            mOriginalPriceDetailStain = value
        End Set
    End Property
    Private mTerritoryCode As String
    Public Property TerritoryCode() As String
        Get
            Return mTerritoryCode
        End Get
        Set(ByVal value As String)
            mTerritoryCode = value
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

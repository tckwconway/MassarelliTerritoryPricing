Imports System.ComponentModel
Imports System.Data.DataRowView
Imports System.String
Imports System.Data.SqlClient
Imports System.Data

'Item Pricing Object is used with the main application datagrid (Price List Object is used with the Excel Export )
Public Class ItemPricingObj
    Inherits BaseClass
    Implements INotifyPropertyChanged
    Implements IEditableObject

#Region "   Methods   "

    Public Shared Function NewPricingObj(ByVal TerCode As String, ByVal ItemNo As String, ByVal ItemDesc As String,
    ByVal ProdCat As String, ByVal ProdCatDesc As String, ByVal OriginalPriceColor As Decimal, ByVal OriginalPriceRococo As Decimal,
    ByVal OriginalPriceDetailStain As Decimal, ByVal ItemLocPriceColor As Decimal, ByVal ItemLocPriceRococo As Decimal,
    ByVal ItemLocPriceDetailStain As Decimal, ByVal ActiveItemLocPriceColor As Decimal, ByVal ActiveItemLocPriceRococo As Decimal,
    ByVal ActiveItemLocPriceDetailStain As Decimal, ByVal CopiedItemLocPriceColor As Decimal, ByVal CopiedItemLocPriceRococo As Decimal,
    ByVal CopiedItemLocPriceDetailStain As Decimal, ByVal TerFrom As String, ByVal TerDescription As String, ByVal LastDate As Date, Selected As Boolean, ByVal ItemWeight As Decimal,
    ByVal Page_n As Integer, ByVal OnPriceList As String, Dimensions As String, A4GLIdentity As Integer) As ItemPricingObj
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
        prc.ActivePriceColor = ActiveItemLocPriceColor
        prc.ActivePriceRococo = ActiveItemLocPriceRococo
        prc.ActivePriceDetailStain = ActiveItemLocPriceDetailStain
        prc.CopiedPriceColor = CopiedItemLocPriceColor
        prc.CopiedPriceRococo = CopiedItemLocPriceRococo
        prc.CopiedPriceDetailStain = CopiedItemLocPriceDetailStain
        prc.TerFrom = TerFrom
        prc.TerDescription = TerDescription
        prc.LastDate = LastDate
        prc.Selected = Selected
        prc.ItemWeight = ItemWeight
        prc.PageNo = Page_n
        prc.OnPriceList = OnPriceList
        prc.Dimensions = Dimensions
        prc.A4GLIdentity = A4GLIdentity
        Return prc
    End Function

    Public Shared Function GetPricingObj(ByVal TerCode As String, ByVal ItemNo As String, ByVal ItemDesc As String,
    ByVal ProdCat As String, ByVal ProdCatDesc As String, ByVal OriginalPriceColor As Decimal, ByVal OriginalPriceRococo As Decimal,
    ByVal OriginalPriceDetailStain As Decimal, ByVal ItemLocPriceColor As Decimal, ByVal ItemLocPriceRococo As Decimal,
    ByVal ItemLocPriceDetailStain As Decimal, ByVal ActiveItemLocPriceColor As Decimal, ByVal ActiveItemLocPriceRococo As Decimal,
    ByVal ActiveItemLocPriceDetailStain As Decimal, ByVal CopiedItemLocPriceColor As Decimal, ByVal CopiedItemLocPriceRococo As Decimal,
    ByVal CopiedItemLocPriceDetailStain As Decimal, ByVal TerFrom As String, ByVal TerDescription As String, ByVal LastDate As Date, Selected As Boolean, ByVal ItemWeight As Decimal,
    ByVal Page_n As Integer, ByVal OnPriceList As String, Dimensions As String, A4GLIdentity As Integer) As ItemPricingObj
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
        prc.ActivePriceColor = ActiveItemLocPriceColor
        prc.ActivePriceRococo = ActiveItemLocPriceRococo
        prc.ActivePriceDetailStain = ActiveItemLocPriceDetailStain
        prc.CopiedPriceColor = CopiedItemLocPriceColor
        prc.CopiedPriceRococo = CopiedItemLocPriceRococo
        prc.CopiedPriceDetailStain = CopiedItemLocPriceDetailStain
        prc.TerFrom = TerFrom
        prc.TerDescription = TerDescription
        prc.LastDate = LastDate
        prc.Selected = Selected
        prc.ItemWeight = ItemWeight
        prc.PageNo = Page_n
        prc.OnPriceList = OnPriceList
        prc.Dimensions = Dimensions
        prc.A4GLIdentity = A4GLIdentity
        Return prc
    End Function


    Public Function SaveItemPricing(ByVal ItemNo As String, ByVal TerCode As String, ByVal OriginalPriceColor As Decimal,
                                    ByVal OriginalPrcColor As Decimal, ByVal OriginalPriceDetailStain As Decimal, ByVal TerFrom As String, ByVal TerDesc As String,
                                    ByVal State As String, ByVal cn As SqlConnection) As Object

        Dim Success As Boolean = False
        Dim Result As Object = 0

        Result = DAC.ExecuteSaveSP(My.Resources.SP_spIMSaveItem_MAS, cn,
                   DAC.Parameter(My.Resources.Param_iItemNo, ItemNo, ParameterDirection.Input),
                   DAC.Parameter(My.Resources.Param_iTerCode, TerCode, ParameterDirection.Input),
                   DAC.Parameter(My.Resources.Param_iPrcColor, OriginalPriceColor, ParameterDirection.Input),
                   DAC.Parameter(My.Resources.Param_iPrcRococo, OriginalPrcColor, ParameterDirection.Input),
                   DAC.Parameter(My.Resources.Param_iPrcDetailStain, OriginalPriceDetailStain, ParameterDirection.Input),
                   DAC.Parameter(My.Resources.Param_iTerFrom, TerFrom, ParameterDirection.Input),
                   DAC.Parameter(My.Resources.Param_iRowState, State, ParameterDirection.Input),
                   DAC.Parameter(My.Resources.Param_iTerDesc, TerDesc, ParameterDirection.Input))
        Me.EntityState = EntityStateEnum.Unchanged
        Return True

    End Function

#End Region



#Region "   Properties   "

    'Private mDirty As Boolean
    'Public Property Dirty() As Boolean
    '    Get
    '        Return mDirty
    '    End Get
    '    Set(ByVal value As Boolean)
    '        mDirty = value
    '    End Set
    'End Property

    Private mDirty As Boolean
    Public Property Dirty() As Boolean
        Get
            Return mDirty
        End Get
        Set(ByVal value As Boolean)
            mDirty = value
        End Set
    End Property

    'Public ReadOnly Property Dirty()
    '    Get
    '        Return Me.isDirty 'Me.EntityState <> EntityStateEnum.Unchanged
    '    End Get
    'End Property
    Private mTerCode As String
    Public Property TerCode() As String
        Get
            Return mTerCode
        End Get
        Set(ByVal value As String)
            If mTerCode <> value Then
                Dim propertyName As String = "TerCode"
                Me.DataStateChanged(EntityStateEnum.Modified)
                mTerCode = value
                Me.Dirty = Me.isDirty

            End If

        End Set
    End Property

    Private mItemNo As String
    Public Property ItemNo() As String
        Get
            Return mItemNo
        End Get
        Set(ByVal value As String)
            mItemNo = value
        End Set
    End Property

    Private mItemDesc As String
    Public Property ItemDesc() As String
        Get
            Return mItemDesc
        End Get
        Set(ByVal value As String)
            mItemDesc = value
        End Set
    End Property

    Private mProdCat As String
    Public Property ProdCat() As String
        Get
            Return mProdCat
        End Get
        Set(ByVal value As String)
            mProdCat = value
        End Set
    End Property

    Private mProdCatDesc As String
    Public Property ProdCatDesc() As String
        Get
            Return mProdCatDesc
        End Get
        Set(ByVal value As String)
            mProdCatDesc = value
        End Set
    End Property

    Private mOriginalPriceColor As Decimal
    Public Property OriginalPriceColor() As Decimal
        Get
            Return mOriginalPriceColor
        End Get
        Set(ByVal value As Decimal)
            If mOriginalPriceColor <> value Then
                Dim propertyName As String = "OriginalPriceColor"
                Me.DataStateChanged(EntityStateEnum.Modified)
                mOriginalPriceColor = value
                Me.Dirty = Me.isDirty

            End If
        End Set
    End Property

    Private mOriginalPriceRococo As Decimal
    Public Property OriginalPriceRococo() As Decimal
        Get
            Return mOriginalPriceRococo
        End Get
        Set(ByVal value As Decimal)
            If mOriginalPriceRococo <> value Then
                Dim propertyName As String = "OriginalPriceRococo"
                Me.DataStateChanged(EntityStateEnum.Modified)
                mOriginalPriceRococo = value
                Me.Dirty = Me.isDirty

            End If
        End Set
    End Property

    Private mOriginalPriceDetailStain As Decimal
    Public Property OriginalPriceDetailStain() As Decimal
        Get
            Return mOriginalPriceDetailStain
        End Get
        Set(ByVal value As Decimal)
            If mOriginalPriceDetailStain <> value Then
                Dim propertyName As String = "OriginalPriceDetailStain"
                Me.DataStateChanged(EntityStateEnum.Modified)
                mOriginalPriceDetailStain = value
                Me.Dirty = Me.isDirty

            End If
        End Set
    End Property


    Private mItemLocPriceColor As Decimal
    Public Property ItemLocPriceColor() As Decimal
        Get
            Return mItemLocPriceColor
        End Get
        Set(ByVal value As Decimal)
            mItemLocPriceColor = value
        End Set
    End Property

    Private mItemLocPriceRococo As Decimal
    Public Property ItemLocPriceRococo() As Decimal
        Get
            Return mItemLocPriceRococo
        End Get
        Set(ByVal value As Decimal)
            mItemLocPriceRococo = value
        End Set
    End Property

    Private mItemLocPriceDetailStain As Decimal
    Public Property ItemLocPriceDetailStain() As Decimal
        Get
            Return mItemLocPriceDetailStain
        End Get
        Set(ByVal value As Decimal)
            mItemLocPriceDetailStain = value
        End Set
    End Property

    Private mSelected As Boolean
    Public Property Selected() As Boolean
        Get
            Return mSelected
        End Get
        Set(ByVal value As Boolean)
            mSelected = value
        End Set
    End Property

    Private mLastDate As Date
    Public Property LastDate() As Date
        Get
            Return mLastDate
        End Get
        Set(ByVal value As Date)
            If value <> CDate("1/1/1900") Then
                mLastDate = value
            Else
                mLastDate = CDate("1/1/1900")
            End If

        End Set
    End Property

    Private mCreateDate As Date
    Public Property CreateDate() As Date
        Get
            Return mCreateDate
        End Get
        Set(ByVal value As Date)
            If value <> CDate("1/1/1900") Then
                mCreateDate = value
            Else
                mCreateDate = CDate("1/1/1900")
            End If

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

    Private mPageNo As Integer
    Public Property PageNo() As Integer
        Get
            Return mPageNo
        End Get
        Set(ByVal value As Integer)
            mPageNo = value
        End Set
    End Property

    Private mDimensions As String
    Public Property Dimensions() As String
        Get
            Return mDimensions
        End Get
        Set(ByVal value As String)
            mDimensions = value
        End Set
    End Property


    Private mOnPriceList As String
    Public Property OnPriceList() As String
        Get
            Return mOnPriceList
        End Get
        Set(ByVal value As String)
            mOnPriceList = value
        End Set
    End Property


    Private mTerFrom As String
    Public Property TerFrom() As String
        Get
            Return mTerFrom
        End Get
        Set(ByVal value As String)
            mTerFrom = value
        End Set
    End Property

    Private mActivePriceColor As Decimal
    Public Property ActivePriceColor() As Decimal
        Get
            Return mActivePriceColor
        End Get
        Set(ByVal value As Decimal)
            mActivePriceColor = value
        End Set
    End Property
    Private mActivePriceRococo As Decimal
    Public Property ActivePriceRococo() As Decimal
        Get
            Return mActivePriceRococo
        End Get
        Set(ByVal value As Decimal)
            mActivePriceRococo = value
        End Set
    End Property
    Private mActivePriceDetailStain As Decimal
    Public Property ActivePriceDetailStain() As Decimal
        Get
            Return mActivePriceDetailStain
        End Get
        Set(ByVal value As Decimal)
            mActivePriceDetailStain = value
        End Set
    End Property

    Private mCopiedPriceColor As Decimal
    Public Property CopiedPriceColor() As Decimal
        Get
            Return mCopiedPriceColor
        End Get
        Set(ByVal value As Decimal)
            mCopiedPriceColor = value
        End Set
    End Property
    Private mCopiedPriceRococo As Decimal
    Public Property CopiedPriceRococo() As Decimal
        Get
            Return mCopiedPriceRococo
        End Get
        Set(ByVal value As Decimal)
            mCopiedPriceRococo = value
        End Set
    End Property
    Private mCopiedPriceDetailStain As Decimal
    Public Property CopiedPriceDetailStain() As Decimal
        Get
            Return mCopiedPriceDetailStain
        End Get
        Set(ByVal value As Decimal)
            mCopiedPriceDetailStain = value
        End Set
    End Property

    Private mTerDescription As String
    Public Property TerDescription() As String
        Get
            Return mTerDescription
        End Get
        Set(ByVal value As String)
            mTerDescription = value
        End Set
    End Property

    Private mA4GLIdentity As Integer
    Public Property A4GLIdentity() As Integer
        Get
            Return mA4GLIdentity
        End Get
        Set(ByVal value As Integer)
            mA4GLIdentity = value
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

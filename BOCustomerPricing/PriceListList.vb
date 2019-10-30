Imports System.Collections.Generic
Imports System.Collections
Imports System.ComponentModel
Public Class PriceListList
    Inherits BindingListView(Of PriceListObj)
    Protected Overrides Function AddNewCore() As Object

        Dim prclst As PriceListObj = Nothing



        Return MyBase.AddNewCore()
    End Function
End Class

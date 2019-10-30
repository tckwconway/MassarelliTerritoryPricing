Imports System.Collections.Generic
Imports System.Collections
Imports System.ComponentModel

Public Class ItemPricingList
    Inherits BindingListView(Of ItemPricingObj)
    Protected Overrides Function AddNewCore() As Object

        'Dim prc As ItemPricingObj

        Return MyBase.AddNewCore()
    End Function
End Class

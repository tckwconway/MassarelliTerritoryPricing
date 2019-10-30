Imports System.Collections.Generic
Imports System.Collections
Imports System.ComponentModel

Public Class ItemLookupList
    Inherits BindingList(Of ItemLookupObj)
    Protected Overrides Function AddNewCore() As Object
        Dim itmlk As ItemLookupObj = ItemLookupObj.NewItemLookup _
            ("", "")
        Add(itmlk)

        Return itmlk

    End Function
End Class

Public Class BaseClass

    Public Enum EntityStateEnum
        Unchanged
        Added
        Deleted
        Modified
    End Enum
    Private mEntityState As EntityStateEnum
    Protected Property EntityState() As EntityStateEnum
        Get
            Return mEntityState
        End Get
        Set(ByVal value As EntityStateEnum)
            mEntityState = value
        End Set
    End Property

    Protected Sub DataStateChanged(ByVal dataState As EntityStateEnum)
        If dataState = EntityStateEnum.Deleted Then
            Me.EntityState = dataState
        End If
        If Me.EntityState = EntityStateEnum.Unchanged OrElse dataState = EntityStateEnum.Unchanged Then
            Me.EntityState = dataState
        End If
    End Sub
    Protected ReadOnly Property isDirty() As Boolean
        Get
            Return Me.EntityState <> EntityStateEnum.Unchanged
        End Get
    End Property
End Class

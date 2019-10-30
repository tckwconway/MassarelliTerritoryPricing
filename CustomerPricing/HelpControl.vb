Public Class HelpControl
    Private appPath As String = ""
    Public helpFile As String = ""
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub lnkHelp_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkHelp.LinkClicked
        Dim Rect As Rectangle
        Dim BodySize As Size
        appPath = System.AppDomain.CurrentDomain.BaseDirectory() & "Help\" & Me.ParentForm.Tag.ToString
        '
        With wbHelp
            'Dock = DockStyle.None
            .Navigate(appPath)
            Rect = .Document.Body.ScrollRectangle
            BodySize = New Size(Rect.Width, Rect.Height)

            'If TypeOf Me.ParentForm Is Lookup Then
            '    CType(Me.ParentForm, Lookup).SplitContainer1.SplitterDistance = 175
            'ElseIf TypeOf Me.ParentForm Is TerrPricing Then

            'End If

            'wbHelp.Size = BodySize
            'wbHelp.Dock = DockStyle.Fill

        End With

    End Sub

    'Private Sub lnkDone_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkDone.LinkClicked
    '    If TypeOf Me.ParentForm Is Lookup Then
    '        CType(Me.ParentForm, Lookup).SplitContainer1.SplitterDistance = Panel1.Height
    '    ElseIf TypeOf Me.ParentForm Is TerrPricing Then

    '    End If
    'End Sub
End Class

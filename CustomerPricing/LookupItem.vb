Imports System.Data.SqlClient
Imports System.Data
Imports System.Text
Imports System.ComponentModel
Imports System.IO
Imports System.Type
Imports System.Linq


Public Class LookupItem
    'Private dtItemsAutoComplete As DataTable = New DataTable
    'Private dvItemsAutoComplete As DataView
    'Private isdirty As IsDirtyTracker

    'Private Sub LookupItem_Load(sender As Object, e As System.EventArgs) Handles Me.Load
    '    LoadForm()
    'End Sub

    'Private Sub LoadForm()
    '    Try

    '        LoadAutoComplete()

    '        isdirty = New IsDirtyTracker(Me)
    '        isdirty.SetAsClean()
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub

    'Private Sub LoadAutoComplete()
    '    Dim sSQL As String
    '    Dim ItmAutoCmplt As New AutoCompleteStringCollection

    '    sSQL = "Select distinct RTrim(lTrim(itm.item_no)) as item_no, RTrim(lTrim(itm.item_desc_1)) as item_desc_1 " & vbCrLf & _
    '           " from IMITMIDX_SQL itm " & vbCrLf & _
    '           " where itm.prod_cat not in ('007', '009', '030', '040', '050', '060', '100', '110', '120', '140', '170', '700', '900', 'CLG', 'DPB', 'FDB', 'FGB', 'HRB', 'MLB', 'NFL', 'TDB', 'TPB') " & vbCrLf & _
    '           " Order by item_no "
    '    Try
    '        dtItemsAutoComplete = BusObj.ExecuteSQLDataTable(sSQL, "ItemsAutoComplete", cn)
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try

    '    'Use Linq to get the Array ...
    '    Dim itms = (From row In dtItemsAutoComplete Select itmno = row(0).ToString).ToArray

    '    ItmAutoCmplt.AddRange(itms)
    '    With txtItem
    '        .AutoCompleteCustomSource = ItmAutoCmplt
    '        .AutoCompleteSource = AutoCompleteSource.CustomSource
    '        .AutoCompleteMode = AutoCompleteMode.SuggestAppend
    '    End With

    '    dvItemsAutoComplete = dtItemsAutoComplete.DefaultView
    '    'dvZoneMarkup.RowFilter = "Trim(ter_zone) = '" & filtertercode & "'"
    'End Sub

    'Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick

    '   Dim tmr As Timer = CType(sender, Timer)

    '    With tmr
    '        .Enabled = False
    '    End With

    '    Try
    '        'Set the Product Description value based on the selected Product No
    '        Dim txt As TextBox = CType(txtItem, TextBox)
    '        dvItemsAutoComplete.RowFilter = "item_no = '" & txt.Text.ToString & "' "
    '        txtItemDescription.Text = dvItemsAutoComplete(0)(1).ToString
    '        ToolTip1.SetToolTip(txtItemDescription, txtItemDescription.Text)
    '    Catch ex As Exception

    '    End Try
    '    isdirty.IsFormPopulated(Me.Controls)
    'End Sub

    'Private Sub txtItem_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtItem.KeyUp
    '    If e.KeyCode = Keys.Enter Then
    '        Dim txt As TextBox = CType(sender, TextBox)

    '        With Timer1
    '            .Interval = 100
    '            .Enabled = True
    '        End With
    '    End If


    'End Sub
End Class


Public Class IsDirtyTracker

    Private _frmTracked As Form
    Private _isDirty As Boolean
    Private _iPopulated As Boolean
    Private btn As Button
    ' property denoting whether the tracked form is clean or dirty
    Public Property IsDirty() As Boolean
        Get
            Return _isDirty
        End Get
        Set(value As Boolean)
            _isDirty = value
        End Set
    End Property

    ' Check that the form is fully populated.  
    Public Property IsPopulated() As Boolean
        Get
            Return _iPopulated
        End Get
        Set(value As Boolean)
            _iPopulated = value
        End Set
    End Property

    ' methods to make dirty or clean
    Public Sub SetAsDirty()
        _isDirty = True
    End Sub

    Public Sub SetAsClean()
        _isDirty = False
        'btn.Enabled = IsDirty
    End Sub

    ' event handlers
    Private Sub IsDirtyTracker_TextChanged(sender As Object, e As EventArgs)
        _isDirty = True
        'btn.Enabled = IsDirty
        'btn.Enabled = IsPopulated
    End Sub

    Private Sub IsDirtyTracker_CheckedChanged(sender As Object, e As EventArgs)
        _isDirty = True
        'btn.Enabled = IsDirty
        'IsFormPopulated(fSpecialItems)
        'btn.Enabled = IsPopulated
    End Sub

    Private Sub IsDirtyTracker_SelectedIndexChanged(sender As Object, e As EventArgs)
        _isDirty = True
        'btn.Enabled = IsDirty
    End Sub

    Public Sub IsFormPopulated(ctls As Control.ControlCollection)
        For Each c As Control In ctls
            If TypeOf c Is TextBox Then
                If c.Text = "" Then
                    IsPopulated = False : Exit For
                Else
                    IsPopulated = True
                End If
            End If
            If TypeOf c Is ComboBox Then
                If c.Text = "" Then
                    IsPopulated = False : Exit For
                Else
                    IsPopulated = True
                End If
            End If

            If c.HasChildren Then
                IsFormPopulated(c.Controls)
            End If
        Next
        'btn.Enabled = IsPopulated
        'btn.Enabled = IsDirty
    End Sub

    Private Sub AssignHandlersForControlCollection(coll As Control.ControlCollection)
        For Each c As Control In coll
            If TypeOf c Is TextBox Then
                AddHandler CType(c, TextBox).TextChanged, AddressOf IsDirtyTracker_CheckedChanged
            End If

            If TypeOf c Is CheckBox Then
                AddHandler CType(c, CheckBox).CheckedChanged, AddressOf IsDirtyTracker_CheckedChanged
                'TryCast(c, CheckBox).CheckedChanged += New EventHandler(IsDirtyTracker_CheckedChanged)
            End If

            If TypeOf c Is ComboBox Then
                AddHandler CType(c, ComboBox).SelectedIndexChanged, AddressOf IsDirtyTracker_SelectedIndexChanged
                'TryCast(c, CheckBox).CheckedChanged += New EventHandler(IsDirtyTracker_CheckedChanged)
            End If

            'If TypeOf c Is Button Then
            '    If c.Name = fSpecialItems.btnSaveItem.Name Then
            '        btn = c
            '    End If
            'End If

            ' recurively apply to inner collections
            If c.HasChildren Then
                AssignHandlersForControlCollection(c.Controls)
            End If
        Next
    End Sub


    Public Sub New(frm As Form)
        _frmTracked = frm
        AssignHandlersForControlCollection(frm.Controls)
    End Sub

End Class

Option Infer On

Imports System.Data.SqlClient
Imports System.Data
Imports System.Text
Imports System.ComponentModel

Public Class Lookup
    Public dtData As DataTable
    Public dtDataAdditional As DataTable
    Public dtTerrCode As DataTable
    Private dtItem As DataTable
    Private sFilter As String
    Private sCTEFilter As String
    Private bClearAll As Boolean
    Private bDataLoaded As Boolean
    Private bIsLoading As Boolean = True
    Private bFirstSelectHighlight As Boolean = True
    Private sFormatType As String

    Public helpFile As String = "Advanced Search.htm"
    Private cOptions As New OptionalCriteriaAdvanced
    Private cItem As New Item
    Public Event SendDataToGrid(dt As DataTable)
    Public Event FormatTerPricing(formatType As String)
    'Contstants ...
    Const BallotBoxWithCheck As Char = ChrW(&H2611)
    Const CheckMark As Char = ChrW(&H2713)
    Const HeavyCheckMark As Char = ChrW(&H2714)
    Const LightCheckMark As Char = ChrW(&H221A)

    Public Sub New(dt As DataTable)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        dtTerrCode = dt
    End Sub
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

    End Sub
    Private Enum FormatType As Integer
        Zone = 1
        Closed = 2
    End Enum
    Private Sub Lookup_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not sFormatType = FormatType.Zone.ToString Then sFormatType = FormatType.Closed.ToString
        RaiseEvent FormatTerPricing(sFormatType)
    End Sub

    Private Sub Lookup_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        LoadTerrCodeDGV()

        With cOptions
            .CheckState = False
            .CheckStateAdditional = False
        End With

        If bIsLoading = True Then
            Timer2.Interval = 50
            Timer2.Enabled = True
            bIsLoading = False
        End If
    End Sub

    Private Sub CollectUIData()

        With cOptions
            .Clear()
            .ItemNo = txtItemNo.Text
            For Each rw As DataGridViewRow In dgvTerrCode.Rows
                With rw
                    If rw.Selected Then
                        If cOptions.TerrCode = "" Then
                            cOptions.TerrCode = "'" & .Cells("prc_level").Value.ToString & "'"
                        Else
                            cOptions.TerrCode = cOptions.TerrCode & ", '" & .Cells("prc_level").Value.ToString & "'"
                        End If

                    End If

                End With
            Next

        End With

    End Sub
    Private Sub CollectGridData()
        cOptions.TerrCode = ""
        With cOptions
            .ItemNo = txtItemNo.Text
            For Each rw As DataGridViewRow In DataGridView1.Rows
                With rw
                    If Convert.ToBoolean(rw.Cells("Selected").Value) = True Then
                        If cOptions.TerrCode = "" Then
                            cOptions.TerrCode = "'" & .Cells("prc_level").Value.ToString & "'"
                        Else
                            cOptions.TerrCode = cOptions.TerrCode & ", '" & .Cells("prc_level").Value.ToString & "'"
                        End If

                    End If

                End With
            Next
            For Each rw As DataGridViewRow In DataGridView2.Rows
                With rw
                    If Convert.ToBoolean(rw.Cells("Selected").Value) = True Then
                        If cOptions.TerrCode = "" Then
                            cOptions.TerrCode = "'" & .Cells("prc_level").Value.ToString & "'"
                        Else
                            cOptions.TerrCode = cOptions.TerrCode & ", '" & .Cells("prc_level").Value.ToString & "'"
                        End If

                    End If

                End With
            Next

        End With

    End Sub

    Private Function BuildFilter() As String
        sFilter = " Where itm.item_no <> '' "
        With cOptions

            Select Case cOptions.IsTerrCode AndAlso cOptions.IsItemNoSet
                Case True
                    sFilter = sFilter & " and itm.item_no = '" & cOptions.ItemNo & "' and prc.prc_level IN (" & cOptions.TerrCode & ")"
                Case False
                    Return ""
            End Select

        End With

        Return sFilter
    End Function
    Private Function BuildCTEFilter() As String
        sCTEFilter = " Where item_no <> '' "
        With cOptions

            Select Case cOptions.IsTerrCode AndAlso cOptions.IsItemNoSet
                Case True
                    sCTEFilter = sCTEFilter & " and item_no = '" & cOptions.ItemNo & "' and prc_level IN (" & cOptions.TerrCode & ")"
                Case False
                    Return ""
            End Select

        End With

        Return sCTEFilter
    End Function
    Private Sub Clear(bClearAll As Boolean)
        'DataGridView1 ...
        If dtData IsNot Nothing Then dtData.Clear()
        With DataGridView1
            .DataSource = Nothing
            .Columns.Clear()

        End With
        If dtDataAdditional IsNot Nothing Then dtDataAdditional.Clear()
        With DataGridView2
            .DataSource = Nothing
            .Columns.Clear()
        End With

        'Clear all controls ...
        If bClearAll Then
            'With dgvTerrCode
            '    For Each rw As DataGridViewRow In .Rows
            '        rw.Cells("Selected").Value = False
            '        rw.Selected = False
            '    Next
            'End With

            txtItemNo.Text = ""
            txtItemDesc.Text = ""
            cOptions.TerrCode = ""
        End If

    End Sub

    Private Sub FillLookupWithItems()

        With DataGridView1
            .AllowUserToResizeRows = False
            .AllowUserToAddRows = False
            .RowHeadersVisible = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells)

            Dim col As New DataGridViewCheckBoxColumn
            col.HeaderText = ""
            col.Name = "Selected"
            col.DataPropertyName = "itm_selected"
            col.Width = 30
            col.Visible = True
            col.DisplayIndex = 0
            col.HeaderText = HeavyCheckMark
            .Columns.Add(col)

            .DataSource = dtData

            With .Columns("item_no")
                .Width = 80
                .HeaderText = "Item No"
                .ReadOnly = True
            End With
            With .Columns("item_desc_1")
                .Width = 120
                .HeaderText = "Description"
                .ReadOnly = True
            End With
            With .Columns("prod_cat")
                .Width = 75
                .HeaderText = "Prod Cat"
                .ReadOnly = True
            End With
            With .Columns("prod_cat_desc")
                .Width = 120
                .HeaderText = "Prod Cat Desc"
                .ReadOnly = True
            End With
            With .Columns("prc_level")
                .Width = 70
                .HeaderText = "Terr Cd"
                .ReadOnly = True
            End With
            With .Columns("prc_natural")
                .Width = 60
                .HeaderText = "Natural"
                .ReadOnly = True
                With .DefaultCellStyle
                    .Format = "N2"
                    .Alignment = DataGridViewContentAlignment.MiddleRight
                End With
            End With
            With .Columns("prc_color")
                .Width = 60
                .HeaderText = "Color"
                With .DefaultCellStyle
                    .Format = "N2"
                    .Alignment = DataGridViewContentAlignment.MiddleRight
                End With
                .ReadOnly = True
            End With
            With .Columns("prc_detail")
                .Width = 60
                .HeaderText = "Detail"
                With .DefaultCellStyle
                    .Format = "N2"
                    .Alignment = DataGridViewContentAlignment.MiddleRight
                End With
                .ReadOnly = True
            End With
            With .Columns("LastDate")
                .Visible = False
            End With

        End With

    End Sub

    Private Sub FillAdditionalItems()
        Dim c As Integer = 0
        Dim tercount As Integer = dgvTerrCode.RowCount - 1
        Dim tercodes(tercount, 2) As String '= Split(cOptions.TerrCode, ",")
        Dim primarykey(2) As String
        For Each row As DataGridViewRow In dgvTerrCode.Rows
            tercodes(c, 0) = row.Cells("prc_level").Value.ToString
            tercodes(c, 1) = row.Cells("ter_from").Value.ToString
            tercodes(c, 2) = row.Cells("ter_desc").Value.ToString
            c += 1
        Next

        'Get the list of TerrCodes the User has selected and were found on a price list
        For Each row As DataGridViewRow In DataGridView1.Rows
            With cOptions
                If .TerrCodeOnPriceList = "" Then
                    .TerrCodeOnPriceList = row.Cells("prc_level").Value.ToString
                Else
                    .TerrCodeOnPriceList = .TerrCodeOnPriceList & " " & row.Cells("prc_level").Value.ToString
                End If

            End With
            
        Next


        With dtDataAdditional
            For i As Integer = 0 To tercodes.GetUpperBound(0)
                primarykey(0) = tercodes(i, 0)
                primarykey(1) = tercodes(i, 1)
                'This line validates that the next Terr Code in the loop is one the user has selected before it's allowed on the list
                If cOptions.TerrCode.IndexOf(primarykey(0).ToString) <> -1 Then

                    If cOptions.TerrCodeOnPriceList Is Nothing OrElse cOptions.TerrCodeOnPriceList.IndexOf(primarykey(0).ToString) = -1 Then
                        .Rows.Add(cItem.NewItemDataRow(primarykey(0), primarykey(1)))
                        .Rows(.Rows.Count - 1)("prc_level") = primarykey(0).Trim  'cItem.prc_level.ToString.Trim
                        .Rows(.Rows.Count - 1)("item_no") = Convert.ToString(IIf(cItem.item_no Is Nothing, "", cItem.item_no.Trim))
                        .Rows(.Rows.Count - 1)("item_desc_1") = Convert.ToString(IIf(cItem.item_desc_1 Is Nothing, "", cItem.item_desc_1.ToString.Trim))
                        .Rows(.Rows.Count - 1)("prod_cat") = Convert.ToString(IIf(cItem.prod_cat Is Nothing, "", cItem.prod_cat.ToString.Trim))
                        .Rows(.Rows.Count - 1)("prod_cat_desc") = Convert.ToString(IIf(cItem.prod_cat_desc Is Nothing, "", cItem.prod_cat_desc))
                        .Rows(.Rows.Count - 1)("prc_natural") = cItem.orig_prc_natural
                        .Rows(.Rows.Count - 1)("prc_color") = cItem.orig_prc_color
                        .Rows(.Rows.Count - 1)("prc_detail") = cItem.orig_prc_detail
                        .Rows(.Rows.Count - 1)("base_prc_nat") = cItem.item_loc_pric_natural
                        .Rows(.Rows.Count - 1)("base_prc_col") = cItem.item_loc_prc_color
                        .Rows(.Rows.Count - 1)("base_prc_det") = cItem.item_loc_prc_detail
                        .Rows(.Rows.Count - 1)("active_natural") = cItem.active_prc_natural
                        .Rows(.Rows.Count - 1)("active_color") = cItem.active_prc_color
                        .Rows(.Rows.Count - 1)("active_detail") = cItem.active_prc_detail
                        .Rows(.Rows.Count - 1)("copiedbas_natural") = cItem.copied_prc_natural
                        .Rows(.Rows.Count - 1)("copiedbas_color") = cItem.copied_prc_color
                        .Rows(.Rows.Count - 1)("copiedbas_detail") = cItem.copied_prc_detail
                        .Rows(.Rows.Count - 1)("ter_from") = primarykey(1).Trim  'cItem.ter_from.ToString.Trim
                        .Rows(.Rows.Count - 1)("ter_desc") = tercodes(i, 2).Trim  ' Convert.ToString(IIf(cItem.ter_desc Is Nothing, "", cItem.ter_desc.ToString.Trim))
                        .Rows(.Rows.Count - 1)("lastdate") = Convert.ToDateTime(IIf(cItem.lastdate.Equals(DBNull.Value), #1/1/1900#, cItem.lastdate))
                        .Rows(.Rows.Count - 1)("itm_selected") = 1  'All Items are to be checked ...   'cItem.itm_selected
                        .Rows(.Rows.Count - 1)("item_weight") = cItem.item_weight
                        .Rows(.Rows.Count - 1)("page_no") = cItem.page_no
                        .Rows(.Rows.Count - 1)("onpricelist") = Convert.ToString(IIf(cItem.onpricelist Is Nothing, "", cItem.onpricelist.ToString.Trim))
                        .Rows(.Rows.Count - 1)("dimensions") = Convert.ToString(IIf(cItem.dimensions Is Nothing, "", cItem.dimensions.ToString.Trim))
                        .Rows(.Rows.Count - 1)("A4GLIdentity") = 0 'Convert.ToInt32(IIf(cItem.A4GLIdentity.ToString Is Nothing, 0, cItem.A4GLIdentity))

                    End If
                End If
            Next

        End With

        With DataGridView2
            .AllowUserToResizeRows = False
            .AllowUserToAddRows = False
            .RowHeadersVisible = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells)

            Dim col As New DataGridViewCheckBoxColumn
            col.HeaderText = ""
            col.Name = "Selected"
            col.DataPropertyName = "itm_selected"
            col.Width = 30
            col.Visible = True
            col.DisplayIndex = 0
            col.HeaderText = HeavyCheckMark
            .Columns.Add(col)

            .DataSource = dtDataAdditional

            With .Columns("item_no")
                .Width = 80
                .HeaderText = "Item No"
                .ReadOnly = True
            End With
            With .Columns("item_desc_1")
                .Width = 120
                .HeaderText = "Description"
                .ReadOnly = True
            End With
            With .Columns("prod_cat")
                .Width = 75
                .HeaderText = "Prod Cat"
                .ReadOnly = True
            End With
            With .Columns("prod_cat_desc")
                .Width = 120
                .HeaderText = "Prod Cat Desc"
                .ReadOnly = True
            End With
            With .Columns("prc_level")
                .Width = 70
                .HeaderText = "Terr Cd"
                .ReadOnly = True
            End With
            With .Columns("prc_natural")
                .Width = 60
                .HeaderText = "Natural"
                With .DefaultCellStyle
                    .Format = "N2"
                    .Alignment = DataGridViewContentAlignment.MiddleRight
                End With
                .ReadOnly = True
            End With
            With .Columns("prc_color")
                .Width = 60
                .HeaderText = "Color"
                With .DefaultCellStyle
                    .Format = "N2"
                    .Alignment = DataGridViewContentAlignment.MiddleRight
                End With
                .ReadOnly = True
            End With
            With .Columns("prc_detail")
                .Width = 60
                .HeaderText = "Detail"
                With .DefaultCellStyle
                    .Format = "N2"
                    .Alignment = DataGridViewContentAlignment.MiddleRight
                End With
                .ReadOnly = True
            End With
            With .Columns("LastDate")
                .Visible = False
            End With

        End With

    End Sub


    Private Sub ButtonSearch_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click

        If txtItemNo.Text = "" Then
            MsgBox("Item Number field is empty.  An item number is required.", MsgBoxStyle.OkOnly, "Item Number")
            Exit Sub
        End If
        Dim itemdesc As String = ValidateItem(txtItemNo)
        If itemdesc = "nothing" Then Exit Sub
        LoadGrids()
        btnOKandClose.Enabled = True

    End Sub

    Private Sub LoadGrids()

        CollectUIData()

        bDataLoaded = False

        GetItem()
        LoadData()

        bDataLoaded = True

    End Sub

    Private Sub GetItem()
        cItem = New Item
        Dim dt As DataTable = BusObj.GetItem(cOptions.ItemNo, cn)
        Try
            With cItem
                .prc_level = Convert.ToString(dt.Rows(0)("prc_level"))
                .item_no = Convert.ToString(dt.Rows(0)("item_no"))
                .item_desc_1 = Convert.ToString(dt.Rows(0)("item_desc_1"))
                .prod_cat = Convert.ToString(dt.Rows(0)("prod_cat"))
                .prod_cat_desc = Convert.ToString(dt.Rows(0)("prod_cat_desc"))
                .orig_prc_natural = Convert.ToDecimal(IIf(dt.Rows(0)("prc_natural").Equals(DBNull.Value), 0, dt.Rows(0)("prc_natural")))
                .orig_prc_color = Convert.ToDecimal(IIf(dt.Rows(0)("prc_color").Equals(DBNull.Value), 0, dt.Rows(0)("prc_color")))
                .orig_prc_detail = Convert.ToDecimal(IIf(dt.Rows(0)("prc_detail").Equals(DBNull.Value), 0, dt.Rows(0)("prc_detail")))

                .item_loc_pric_natural = Convert.ToDecimal(IIf(dt.Rows(0)("base_prc_nat").Equals(DBNull.Value), 0, dt.Rows(0)("base_prc_nat")))
                .item_loc_prc_color = Convert.ToDecimal(IIf(dt.Rows(0)("base_prc_col").Equals(DBNull.Value), 0, dt.Rows(0)("base_prc_col")))
                .item_loc_prc_detail = Convert.ToDecimal(IIf(dt.Rows(0)("base_prc_det").Equals(DBNull.Value), 0, dt.Rows(0)("base_prc_det")))


                .active_prc_natural = 0 'Convert.ToDecimal(IIf(dt.Rows(0)("copiedbas_natural").Equals(DBNull.Value), 0, dt.Rows(0)("copiedbas_natural")))
                .active_prc_color = 0 ' Convert.ToDecimal(IIf(dt.Rows(0)("copiedbas_color").Equals(DBNull.Value), 0, dt.Rows(0)("copiedbas_color")))
                .active_prc_detail = 0 ' Convert.ToDecimal(IIf(dt.Rows(0)("copiedbas_detail").Equals(DBNull.Value), 0, dt.Rows(0)("copiedbas_detail")))

                .copied_prc_natural = Convert.ToDecimal(IIf(dt.Rows(0)("copiedbas_natural").Equals(DBNull.Value), 0, dt.Rows(0)("copiedbas_natural")))
                .copied_prc_color = Convert.ToDecimal(IIf(dt.Rows(0)("copiedbas_color").Equals(DBNull.Value), 0, dt.Rows(0)("copiedbas_color")))
                .copied_prc_detail = Convert.ToDecimal(IIf(dt.Rows(0)("copiedbas_detail").Equals(DBNull.Value), 0, dt.Rows(0)("copiedbas_detail")))

                .ter_from = Convert.ToString(dt.Rows(0)("ter_from"))
                .ter_desc = Convert.ToString(dt.Rows(0)("ter_desc"))
                .lastdate = Convert.ToDateTime(IIf(dt.Rows(0)("lastdate").Equals(DBNull.Value), #1/1/1900#, dt.Rows(0)("lastdate")))
                .item_weight = Convert.ToDouble(IIf(dt.Rows(0)("item_weight").Equals(DBNull.Value), 0, dt.Rows(0)("item_weight")))
                .page_no = Convert.ToInt32(dt.Rows(0)("page_no"))
                .onpricelist = Convert.ToString(dt.Rows(0)("onpricelist"))
                .dimensions = Convert.ToString(dt.Rows(0)("dimensions"))
                .A4GLIdentity = Convert.ToInt32(dt.Rows(0)("A4GLIdentity"))
            End With

        Catch ex As Exception

        End Try
        
    End Sub


    Private Function GetDataTableFromDataGridView(dgv As DataGridView) As DataTable
        Dim dt As New DataTable()

        ' add the columns to the datatable            
        If dgv IsNot Nothing Then
            For i As Integer = 0 To dgv.Columns.Count - 1
                dt.Columns.Add(dgv.Columns(i).Name.ToString)
            Next
            'dt.Columns.Add("OrderNo")
        End If

        '  add each of the data rows to the table
        For Each row As DataGridViewRow In dgv.Rows
            If Convert.ToBoolean(row.Cells("Selected").Value) = True Then
                Dim dr As DataRow
                dr = dt.NewRow()

                For i As Integer = 0 To row.Cells.Count - 1
                    dr(i) = row.Cells(i).Value.ToString.Trim '        .Value.ToString.Replace(" ", "")
                Next
                dt.Rows.Add(dr)
            End If

        Next
        Return dt
    End Function

    Private Function GetDataTableFromDataGridView(dgv As DataGridView, dgv2 As DataGridView) As DataTable
        Dim itms = dtdata.AsEnumerable()




        Dim dt As New DataTable()

        ' add the columns to the datatable            
        If dgv IsNot Nothing Then
            For i As Integer = 0 To dgv.Columns.Count - 1
                dt.Columns.Add(dgv.Columns(i).DataPropertyName.ToString)
            Next

        End If

        '  add each of the data rows to the table
        For Each row As DataGridViewRow In dgv.Rows
            If Convert.ToBoolean(row.Cells("Selected").Value) = True Then
                Dim dr As DataRow
                dr = dt.NewRow()

                For i As Integer = 0 To row.Cells.Count - 1
                    dr(i) = row.Cells(i).Value.ToString.Trim        '.Replace(" ", "")
                Next
                dt.Rows.Add(dr)
            End If

        Next

        For Each row As DataGridViewRow In dgv2.Rows
            If Convert.ToBoolean(row.Cells("Selected").Value) = True Then
                Dim dr As DataRow
                dr = dt.NewRow()

                For i As Integer = 0 To row.Cells.Count - 1
                    dr(i) = row.Cells(i).Value.ToString.Trim        '.Replace(" ", "")
                Next
                dt.Rows.Add(dr)
            End If

        Next
        Return dt
    End Function



    Private Sub LoadData()

        If Not bDataLoaded Then
            bClearAll = False
            Clear(bClearAll)
            sFilter = ""
            sFilter = BuildFilter()
            sCTEFilter = BuildCTEFilter()
            With Timer1
                .Interval = 50
                .Enabled = True
            End With

            dtData = BusObj.GetItemLookupDataTable(sFilter, sCTEFilter, cn)
            'set a primary key to prc_level

            'Clone the table so we don't get Primarykey in dtDataAdditional
            dtDataAdditional = dtData.Clone

            Dim primaryKey(2) As DataColumn
            primaryKey(0) = dtData.Columns("prc_level")
            primaryKey(1) = dtData.Columns("ter_from")
            primaryKey(2) = dtData.Columns("A4GLIdentity")
            dtData.PrimaryKey = primaryKey

        End If

    End Sub

    Private Sub Lookup_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        dtTerrCode.Clear()
        TerrPricing.LookupClosed()
    End Sub

    Public Sub LoadTerrCodeDGV()
        Dim dt As DataTable = dtTerrCode.Copy
        With dt
            .Rows.RemoveAt(0)
            .Columns.RemoveAt(2)
        End With

        With dgvTerrCode
            .DataSource = Nothing
            .Columns.Clear()

            .DataSource = dt
            .Refresh()
            With .Columns(0)
                .Width = 35
                .HeaderText = "Code"
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
           End With
            With .Columns(1)
                .Width = 110
                .HeaderText = "Description"
            End With
            With .Columns(2)
                .Width = 35
                .HeaderText = "From"
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
           End With

            Dim col As New DataGridViewCheckBoxColumn
            With col
                .Width = 21
                .DisplayIndex = 0
                .Name = "Selected"
            End With

            .Columns.Add(col)
            With .Columns(3)
                .HeaderText = HeavyCheckMark
                .Visible = False
            End With

        End With
    End Sub
    Private Function IsAnythingSelected() As Boolean

        Return CBool(IIf(DataGridView1.SelectedRows.Count = 0, False, True))

    End Function

#Region "   Context Menu   "

    Private Sub SelectAll()
        Dim rw As DataGridViewRow
        Me.Validate()
        With DataGridView1
            For Each rw In .Rows
                rw.Cells("Selected").Value = cOptions.CheckState
            Next
            If cOptions.CheckState = True Then
                cOptions.CheckState = False
            Else
                cOptions.CheckState = True
            End If

        End With
    End Sub

    Private Sub SelectAllTerrCodes(dgv As DataGridView)
        Dim rw As DataGridViewRow
        Cursor = Cursors.WaitCursor
        Me.Validate()
        With dgv

            Select Case dgv.Name
                Case "DataGridView1"
                    For Each rw In .Rows
                        rw.Cells("Selected").Value = cOptions.CheckState
                    Next

                    If cOptions.CheckState = True Then
                        cOptions.CheckState = False
                    Else
                        cOptions.CheckState = True
                    End If

                Case "DataGridView2"
                    For Each rw In .Rows
                        rw.Cells("Selected").Value = cOptions.CheckStateAdditional
                    Next

                    If cOptions.CheckStateAdditional = True Then
                        cOptions.CheckStateAdditional = False
                    Else
                        cOptions.CheckStateAdditional = True
                    End If

            End Select
            
        End With
        Cursor = Cursors.Default
    End Sub

    Private Sub SelectHighlightedFromGrid(dgv As DataGridView)
        Dim i As Integer = 0
        ' Dim chkstate As Boolean
        Dim ctr As Integer = 0  ' use as a counter to select the first check value, set the chkstate to the opposite value and retain the new value for the entire loop 

        Dim selectedRowCount As Integer = _
        dgv.Rows.GetRowCount(DataGridViewElementStates.Selected)
        dgv.EndEdit()
        If selectedRowCount > 0 Then

            Dim sb As New System.Text.StringBuilder()
            If bFirstSelectHighlight = True Then
                ctr = 1
                cOptions.CheckState = True
            End If
            For i = 0 To dgv.RowCount - 1

                If dgv.Rows.Item(i).Selected = True Then
                    If ctr = 0 Then
                        ctr = 1
                        If dgv.Item("Selected", i).Value Is DBNull.Value OrElse CBool(dgv.Item("Selected", i).Value) = False Then
                            cOptions.CheckState = True
                        Else
                            cOptions.CheckState = False
                        End If
                    End If

                    dgv.Item("Selected", i).Value = CInt(cOptions.CheckState)

                End If

            Next i

            sb.Append("Total: " + selectedRowCount.ToString())
            dgv.Refresh()

        End If

        bFirstSelectHighlight = False
    End Sub

#End Region

    Private Sub ButtonSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SelectAll()
    End Sub

    Private Sub TextBoxEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim txt As TextBox = CType(sender, TextBox)
        txt.SelectAll()
    End Sub

    Private Sub FillItemsOnKeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            FillLookupWithItems()
        End If
    End Sub

    Private Sub ToolStripMacola_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs)
        Dim ctl As ToolStripButton = TryCast(e.ClickedItem, ToolStripButton)
        If ctl Is Nothing Then Exit Sub

        Select Case ctl.Name

            Case "CloseForm"
                Me.Close()

        End Select
    End Sub
    Private Sub txtItemNo_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtItemNo.KeyDown


        If txtItemNo.Text = "" Then Exit Sub
        Dim txt As TextBox = CType(sender, TextBox)

        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Then
            'ValidateItem(txt)
            e.Handled = True
            With btnOKandClose
                .Enabled = True
                .Focus()
            End With
        End If
    End Sub

    Private Function ValidateItem(txt As TextBox) As String
        Dim itemdesc As String = BusObj.GetItemDesc(txt.Text, cn)

        If itemdesc = "nothing" Then
            MsgBox("Item " & txt.Text & " not found in Item Master", MsgBoxStyle.OkCancel, "Item Not Found")
            txtItemDesc.Text = ""
        Else
            txtItemDesc.Text = itemdesc
            If dgvTerrCode.SelectedRows.Count = 0 Then Return "nothing"
            LoadGrids()

        End If

        Return itemdesc
    End Function

    Private Sub txtItemNo_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtItemNo.KeyUp
        Dim txt As TextBox = CType(sender, TextBox)
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Then
            With Timer1
                .Interval = 500
                .Enabled = True
            End With
        End If
        
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Dim tmr As Timer = CType(sender, Timer)

        tmr.Enabled = False

        FillLookupWithItems()
        FillAdditionalItems()
    End Sub



    Private Sub txtItemNo_Leave(sender As Object, e As System.EventArgs) Handles txtItemNo.Leave

        If txtItemNo.Text = "" Then Exit Sub

        Dim itemdesc As String = BusObj.GetItemDesc(CType(sender, TextBox).Text, cn)
        If itemdesc = "nothing" Then
            MsgBox("Item " & CType(sender, TextBox).Text & " not found in Item Master", MsgBoxStyle.OkCancel, "Item Not Found")
            txtItemDesc.Text = ""
            Exit Sub
        Else
            txtItemDesc.Text = itemdesc
        End If

    End Sub

    Private Sub btnClear_Click(sender As System.Object, e As System.EventArgs) Handles btnClear.Click
        Dim bClearAll As Boolean = True
        Clear(bClearAll)
    End Sub

    Private Sub txtItemDesc_MouseEnter(sender As Object, e As System.EventArgs) Handles txtItemDesc.MouseEnter
        ToolTip1.Show(CType(sender, TextBox).Text, CType(sender, TextBox))
    End Sub

    Private Sub txtItemDesc_MouseLeave(sender As Object, e As System.EventArgs) Handles txtItemDesc.MouseLeave
        ToolTip1.Hide(CType(sender, TextBox))
    End Sub

    Private Sub Timer2_Tick(sender As System.Object, e As System.EventArgs) Handles Timer2.Tick
        Timer2.Enabled = False
        'Dim bSelectAll As Boolean = True
        'Dim dgv As DataGridView = CType(dgvTerrCode, DataGridView)
        'txtItemNo.Focus()
        'SelectDGVRows(dgv, bSelectAll)
        'With btnSelectAll
        '    .Text = "Select None"
        '    .Image = My.Resources.SelectNone1616
        'End With

        Dim bSelectAll As Boolean = False
        'Dim dgv As DataGridView = CType(dgvTerrCode, DataGridView)
        txtItemNo.Focus()
        'SelectDGVRows(dgv, bSelectAll)
        With btnSelectAll
            .Text = "Select All"
            .Image = My.Resources.SelectAll1616
        End With
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) _
                                                     Handles DataGridView2.ColumnHeaderMouseClick, DataGridView1.ColumnHeaderMouseClick

        Dim dgv As DataGridView = CType(sender, DataGridView)
        If e.ColumnIndex = dgv.Columns("Selected").Index Then SelectAllTerrCodes(dgv)

    End Sub

    Private Sub mnuSelectHighlighted_Click(sender As System.Object, e As System.EventArgs) Handles mnuSelectHighlighted.Click

        'Retrieve the datagridview the context menu strip is assocatiated with ...
        Dim mnuitm As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Dim mnustrp As ContextMenuStrip = CType(mnuitm.Owner, ContextMenuStrip)
        Dim dgv As DataGridView = CType(mnustrp.SourceControl, DataGridView)

        SelectHighlightedFromGrid(dgv)

    End Sub

   
    Private Sub btnOKandClose_Click(sender As System.Object, e As System.EventArgs) Handles btnOKandClose.Click
        sFormatType = FormatType.Zone.ToString

        Dim dt As DataTable = GetDataTableFromDataGridView(DataGridView1, DataGridView2)


        bDataLoaded = True
        'send the existing priced items first
        If dt.Rows.Count > 0 Then
            RaiseEvent SendDataToGrid(dt)
        End If

        bDataLoaded = False

        Me.Close()

    End Sub

    Private Sub DataGridView_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseClick, DataGridView2.MouseClick

        If e.Button = Windows.Forms.MouseButtons.Left And My.Computer.Keyboard.ShiftKeyDown Or e.Button = Windows.Forms.MouseButtons.Left And My.Computer.Keyboard.CtrlKeyDown Then

            Dim dgv As DataGridView = CType(sender, DataGridView)
            SelectHighlightedFromGrid(dgv)

        End If

    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        sFormatType = FormatType.Closed.ToString
        Me.Close()
    End Sub

    Private Sub btnSelectAll_Click(sender As System.Object, e As System.EventArgs) Handles btnSelectAll.Click
        Dim dgv As DataGridView = CType(dgvTerrCode, DataGridView)
        Dim btn As Button = CType(btnSelectAll, Button)
        Dim bIsSelectAll As Boolean
        With btn
            bIsSelectAll = Convert.ToBoolean(IIf(.Text = "Select All", True, False))
            If bIsSelectAll Then
                .Text = "Select None"
                .Image = My.Resources.SelectNone1616
            Else
                .Text = "Select All"
                .Image = My.Resources.SelectAll1616
            End If
        End With

        SelectDGVRows(dgv, bIsSelectAll)

    End Sub


    Private Sub SelectDGVRows(dgv As DataGridView, bSelectAll As Boolean)
        With dgv
            For Each rw As DataGridViewRow In .Rows
                rw.Selected = bSelectAll
            Next
        End With
    End Sub

    Private Sub txtItemNo_PreviewKeyDown(sender As Object, e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles txtItemNo.PreviewKeyDown

        If txtItemNo.Text = "" Then Exit Sub
        Dim txt As TextBox = CType(sender, TextBox)

        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Then
            ValidateItem(txt)
        End If
    End Sub

    Private Sub txtItemNo_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtItemNo.TextChanged

    End Sub
End Class

Friend Class OptionalCriteriaAdvanced
    Private mItemNo As String
    Private mTerrCode As String
    Private mTerrCodeOnPriceList As String
    Private mCheckState As Boolean
    Private mCheckStateAdditonal As Boolean

    Public Sub New()
        mItemNo = ""
        mTerrCode = ""
        mCheckState = False
    End Sub

    Public Sub Clear()
        ItemNo = ""
        TerrCode = ""
        TerrCodeOnPriceList = ""
        CheckState = False
        CheckStateAdditional = False
    End Sub

    Public Shared ReadOnly Property IsSet(ByVal oValue As Object) As Boolean
        Get
            Dim value As Boolean = False
            If oValue Is Nothing Then
                value = False
            ElseIf TypeOf oValue Is String Then
                Dim sValue As String = CType(oValue, String)
                sValue = sValue.Trim
                If sValue <> "" And sValue <> "%" Then
                    value = True
                End If
            End If
            Return value
        End Get
    End Property
    Public ReadOnly Property IsItemNoSet() As Boolean
        Get
            Return IsSet(ItemNo)
        End Get
    End Property
    Public ReadOnly Property IsTerrCode() As Boolean
        Get
            Return IsSet(TerrCode)
        End Get
    End Property
    Public Property ItemNo() As String
        Get
            Return mItemNo
        End Get
        Set(ByVal value As String)
            mItemNo = value
        End Set
    End Property
    Public Property TerrCode() As String
        Get
            Return mTerrCode
        End Get
        Set(ByVal value As String)
            mTerrCode = value
        End Set
    End Property
    Public Property TerrCodeOnPriceList() As String
        Get
            Return mTerrCodeOnPriceList
        End Get
        Set(ByVal value As String)
            mTerrCodeOnPriceList = value
        End Set
    End Property
    Public Property CheckState() As Boolean
        Get
            Return mCheckState
        End Get
        Set(ByVal value As Boolean)
            mCheckState = value
        End Set
    End Property
    Public Property CheckStateAdditional() As Boolean
        Get
            Return mCheckStateAdditonal
        End Get
        Set(ByVal value As Boolean)
            mCheckStateAdditonal = value
        End Set
    End Property

End Class

Public Class Item

    Public prc_level As String
    Public item_no As String
    Public item_desc_1 As String
    Public prod_cat As String
    Public prod_cat_desc As String
    Public orig_prc_natural As Decimal
    Public orig_prc_color As Decimal
    Public orig_prc_detail As Decimal
    Public item_loc_pric_natural As Decimal
    Public item_loc_prc_color As Decimal
    Public item_loc_prc_detail As Decimal
    Public active_prc_natural As Decimal
    Public active_prc_color As Decimal
    Public active_prc_detail As Decimal
    Public copied_prc_natural As Decimal
    Public copied_prc_color As Decimal
    Public copied_prc_detail As Decimal
    Public ter_from As String
    Public ter_desc As String
    Public lastdate As Date
    Public itm_selected As Integer
    Public item_weight As Double
    Public page_no As Integer
    Public onpricelist As String
    Public dimensions As String
    Public A4GLIdentity As Integer

    Public Sub clear()

        prc_level = ""
        item_no = ""
        item_desc_1 = ""
        prod_cat = ""
        prod_cat_desc = ""
        orig_prc_natural = 0
        orig_prc_color = 0
        orig_prc_detail = 0
        item_loc_pric_natural = 0
        item_loc_prc_color = 0
        item_loc_prc_detail = 0
        active_prc_natural = 0
        active_prc_color = 0
        active_prc_detail = 0
        ter_from = ""
        ter_desc = ""
        lastdate = #1/1/1900#
        itm_selected = 0
        item_weight = 0
        page_no = 0
        onpricelist = ""
        dimensions = ""

    End Sub

    Public Function NewItemDataRow(prc_level As String, ter_from As String) As Item
        'Dim rw As DataRow
        Dim itm As New Item
        With itm
            .prc_level = prc_level
            .item_no = ""
            .item_desc_1 = ""
            .prod_cat = ""
            .prod_cat_desc = ""
            .orig_prc_natural = 0
            .orig_prc_color = 0
            .orig_prc_detail = 0
            .item_loc_pric_natural = 0
            .item_loc_prc_color = 0
            .item_loc_prc_detail = 0
            .active_prc_natural = 0
            .active_prc_color = 0
            .active_prc_detail = 0
            .ter_from = ter_from
            .ter_desc = ""
            .lastdate = #1/1/1900#
            .itm_selected = 0
            .item_weight = 0
            .page_no = 0
            .onpricelist = ""
            .dimensions = ""
        End With

        Return itm
    End Function

    Private Sub SelectHighlightedFromGrid(dgv As DataGridView)
        Dim i As Integer
        Dim chkstate As Boolean
        
        Dim selectedRowCount As Integer = _
        dgv.Rows.GetRowCount(DataGridViewElementStates.Selected)

        If selectedRowCount > 0 Then

            Dim sb As New System.Text.StringBuilder()
            Dim ctr As Integer = 0  ' use as a counter to select the first check value, set the chkstate to the opposite value and retain the new value for the entire loop 
            For i = 0 To dgv.RowCount - 1

                If dgv.Rows.Item(i).Selected = True Then
                    If ctr = 0 Then
                        ctr = 1
                        If Convert.ToBoolean(dgv.Item(0, i).Value) = False Then
                            chkstate = True
                        Else
                            chkstate = False
                        End If
                    End If

                    dgv.Item(0, i).Value = chkstate

                End If

            Next i

            sb.Append("Total: " + selectedRowCount.ToString())
            dgv.Refresh()

        End If
    End Sub
End Class
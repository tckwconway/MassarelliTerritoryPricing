Imports System.Data.SqlClient
Imports System.Data
Imports System.Text
Imports System.ComponentModel
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices.Marshal
Imports System.Runtime.InteropServices

Public Class TerrPricing
    'Classes ...
    Friend cItemPricingList As New ItemPricingList
    Private cItemPricingObj As New ItemPricingObj
    Private cSearchClass As SearchClass = New SearchClass
    Private cOptionalCriteria As New OptionalCriteria
    Private cOptionalCriteriaSearch As New OptionalCriteria
    Private itmprclst As New ItemPricingList
    'Forms ...
    Private WithEvents fAdvancedLookup As New Lookup

    'Variables ...
    ' - colors 
    Private readOnlyBackColor As Color = Color.FromArgb(225, 225, 211)
    Private readOnlyBackColorDark As Color = Color.FromArgb(201, 201, 188)
    ' - datatables
    Public dtTerrCodes As DataTable
    Public dtCopyFrom As DataTable
    Private dtSave As DataTable
    Private dtDelete As DataTable
    Private dtFilteredItems As New DataTable

    ' - strings
    Private exportExcelFileName As String
    'Friend items As String
    Friend type As String
    Friend prclevel As String
    Friend onprclist As String
    Private sSearchType As String ' Either from Item Master Table or Terr Pricing Table, IMITMIDX_SQL or OEPRCCUS_MAZ
    Private terDescFill As String

    ' - integers
    Private result As Integer
    ' - booleans
    Private bCheckState As Boolean
    Private bIsLoading As Boolean = True
    Private bByPass As Boolean = False
    Private bEnableSaveDB As Boolean = False
    Private bEnableRefreshDB As Boolean = True
    Private bEnableCtls As Boolean = True
    Private bSkipChecked As Boolean
    Private bSetCriteria As Boolean
    Private bLoadPressed As Boolean
    Private bFillPressed As Boolean
    Private bAdvanceSearch As Boolean
    Private bCopyToOption As Boolean = False
    'Controls ...
    Private chkd As CheckBox
    'Hit Test
    Private hitContextMenu As DataGridView.HitTestInfo
    Private ht As DataGridView.HitTestInfo

    'Events from other classes ....
    '   SendDataToGrid from Lookup.vb

    'Contstants ...
    Const BallotBoxWithCheck As Char = ChrW(&H2611)
    Const CheckMark As Char = ChrW(&H2713)
    Const HeavyCheckMark As Char = ChrW(&H2714)
    Const LightCheckMark As Char = ChrW(&H221A)
    Const ArrowRightTriangle As Char = ChrW(&H25B6)
    'Enums ...

    'Price Type is used for Fill Options.  
    ' - Copied: copies a Primary Zone pricing (i.e. 002, 003 etc) directly into the Active Price Column
    ' - CustomPercent: Does a calculation on the 
    Private Enum PrcType As Integer
        Primary = 0
        Copied = 1
        CustomPercent = 2
        CustomDollar = 3
        Advanced = 4
        Zone = 5
        'DirectCopied = 5
        'Update = 6
        'Advanced = 7
        'Copied = 1

        ''Update = 4
        ''Primary = 1
    End Enum

    'Search Type is sent as a parameter to spIMGetItemList_MAS to return either the Primary Price Only or 
    'the Copied Prices (i.e. 143 VAN LIEWS) with the Primary Prices (the Terr_From prices, i.e. 002) so
    'we can compare the Copied to Primary and if differnent, the row turns Yellow on the DataGridView
    Private Enum SearchType As Integer
        CopiedPrice = 1
        PrimaryPrice = 2
        FromMacola = 3
    End Enum
    'Markup Type 
    Private Enum MarkupType As Integer
        Copy = 1
        Custom = 2
        Zone = 3
        Direct = 4
        'Manual = 4
    End Enum

    Public Enum PrcUpdateType As Integer
        Update = 1
        Copied = 2
        DirectCopy = 3
        Primary = 4


        PrimaryUpdate = 11
        CopiedUpdate = 12
        PrimaryNew = 13
        CopiedNew = 14
        PrimaryUpdateManualPrc = 15
        CopiedUpdateManualPrc = 16
        PrimaryNewManualPrc = 17
        CopiedNewManualPrc = 18
        Undetermined = 19
        Advanced = 20
        Delete = 99
    End Enum

    Private Enum FillType As Integer
        Percent = 1
        Amount = 2
    End Enum

    Private Sub CustPricing_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.Validate()
        Me.Refresh()
        Me.ItemPricingObjBindingSource.DataSource = cItemPricingList
        Me.ItemPricingObjBindingSource.ResetBindings(True)
    End Sub

    Private Enum FormatPriceTypeState
        Open = 1
        Update = 2
        Copy = 3
        DirectCopy = 4
        Primary = 5
        Closed = 6
        Advanced = 7
    End Enum

    Private Enum FormatFillTypeState
        Open = 0
        Manual = 1
        Zone = 2
        Custom = 3
        Copy = 4
        Close = 5
    End Enum

#Region "   Startup   "

    Private Sub CustPricing_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        bIsLoading = True
        LoadTerritoryPricing()
    End Sub

    Private Sub LoadTerritoryPricing()
        'Load the default DB
        cOptionalCriteria.DBName = My.Settings.DefaultDB
        Try
            MacStartup(My.Settings.DefaultDB)
        Catch ex As Exception
            MsgBox("MacStartup " & ex.Message)
        End Try

        'Load the list of SQL Databases
        Try
            listSQLDatabases()
        Catch ex As Exception
            MsgBox("ListSQLDatabases " & ex.Message)
        End Try

        cbDBList.Text = My.Settings.DefaultDB
        'Status bar labels ...
        lblCurrentDB.Text = cOptionalCriteria.CurrentDB
        lblDefaultDB.Text = cOptionalCriteria.DefaultDB

        'Load Search Terr List
        FillTerritoryCodeList()

        fillZoneCodes()
        fillCategoryCodes()

        SetupDataGrids()

        'Control Startup Settings...
        rbFromItemMaster.Checked = True
        btnSaveDefaultDB.Enabled = False
        btnRefreshDB.Enabled = False

        'This is to get the focus on the Mujlti Column Combobox
        With Timer2
            .Interval = 200
            .Enabled = True
        End With

        mcboFillTerrCodes.MaxDropDownItems = 20
        mcboTerritoryCodes.MaxDropDownItems = 20

        btnSelectAll.Enabled = False
        btnFillPricing.Enabled = False
        btnRoundPricing.Enabled = False
        lblColor.Enabled = False
        lblDetail.Enabled = False
        lblNatural.Enabled = False
        pnlFillOptions.Enabled = False
        txtColorMarkup.Enabled = False
        txtDetailMarkup.Enabled = False
        txtNaturalMarkup.Enabled = False
        'cboZonePrc.Enabled = False

        bIsLoading = False

        bCheckState = True

        If My.Settings.DefaultDB = "DATA" Then
            lblCompany.Visible = False : lblTest.Visible = False
        Else
            lblCompany.Visible = True : lblTest.Visible = True
        End If

    End Sub

    Private Sub SetupDataGrids()

        bCheckState = True

        'Setup Header for Title Only DataGridVeiw1 
        Dim itmcatwidth As Integer
        Dim natwidth As Integer
        Dim colwidth As Integer
        Dim detwidth As Integer
        With DataGridView1
            .RowHeadersVisible = False

            With .Columns("Selected")
                .HeaderText = HeavyCheckMark
                .Width = 22
            End With

            With .Columns("ItemLocPriceColor")
                .ReadOnly = True
                .DefaultCellStyle.BackColor = readOnlyBackColor
            End With
            With .Columns("ItemLocPriceRococo")
                .ReadOnly = True
                .DefaultCellStyle.BackColor = readOnlyBackColor
            End With
            With .Columns("ItemLocPriceDetailStain")
                .ReadOnly = True
                .DefaultCellStyle.BackColor = readOnlyBackColor
            End With

            With .Columns("OriginalPriceColor")
                .ReadOnly = True
                .DefaultCellStyle.BackColor = readOnlyBackColor
            End With
            With .Columns("OriginalPriceRococo")
                .ReadOnly = True
                .DefaultCellStyle.BackColor = readOnlyBackColor
            End With
            With .Columns("OriginalPriceDetailStain")
                .ReadOnly = True
                .DefaultCellStyle.BackColor = readOnlyBackColor
            End With

            .Columns("TerFrom").Width = 60

            itmcatwidth = .Columns("Selected").Width + .Columns("ItemNumber").Width + .Columns("ItemDescription").Width + .Columns("ProdCategory").Width + .Columns("ProdCatDescription").Width + .Columns("TerCode").Width '+ .RowHeadersWidth
            natwidth = .Columns("ItemLocPriceColor").Width + .Columns("OriginalPriceColor").Width + .Columns("ActivePriceColor").Width
            colwidth = .Columns("ItemLocPriceRococo").Width + .Columns("OriginalPriceRococo").Width + .Columns("ActivePriceRococo").Width
            detwidth = .Columns("ItemLocPriceDetailStain").Width + .Columns("OriginalPriceDetailStain").Width + .Columns("ActivePriceDetailStain").Width
            .Dock = DockStyle.Fill
            .ScrollBars = ScrollBars.Both
        End With

        With dgvHeader
            .Columns("Column1").Width = itmcatwidth '- 1
            .Columns("Column2").Width = natwidth '- 1
            .Columns("Column3").Width = colwidth '- 2
            .Columns("Column4").Width = detwidth '- 3
            .AllowUserToAddRows = False
            .Height = .ColumnHeadersHeight
        End With

    End Sub

    Private Sub listSQLDatabases()
        On Error Resume Next

        Dim cmd As New SqlCommand("", cn)
        Dim rdr As SqlDataReader
        cmd.CommandText = "exec sys.sp_databases"

        rdr = cmd.ExecuteReader()
        With cbDBList
            While (rdr.Read())
                If rdr.GetString(0).Substring(0, 4) = "DATA" Then .Items.Add(rdr.GetString(0))

            End While
        End With
        rdr.Dispose()
        cmd.Dispose()

    End Sub

#End Region

#Region "   Methods   "


    Private Sub SetOptionalCriteria()
        On Error Resume Next
        With cOptionalCriteria
            .Clear()
            .SearchType = IIf(txtFrom.Text.Trim = "", SearchType.PrimaryPrice.ToString, SearchType.CopiedPrice.ToString).ToString
            .PricingType = PrcType.Copied.ToString
            'LOAD PRICE Group
            .TerFromCode = mcboTerritoryCodes.Data.Rows(Me.mcboTerritoryCodes.SelectedIndex)(0).ToString.Trim ' LOAD PRICE: mcboTerritoryCodes
            .TerCopiedFromCode = mcboTerritoryCodes.Data.Rows(Me.mcboTerritoryCodes.SelectedIndex)(3).ToString.Trim ' LOAD PRICE: txtFrom
            .TerFromDesc = mcboTerritoryCodes.Data.Rows(Me.mcboTerritoryCodes.SelectedIndex)(1).ToString.Trim ' LOAD PRICE: txtTerrDesc
            .ProdCat = cboCategoryCodes.Text.Trim 'LOAD PRICE: cboCategoryCodes
            .OnPriceList = cboOnPriceList.Text.Trim 'LOAD PRICE: cboOnPriceList
            .HasTerrPrice = cboHasTerrPrice.Text.Trim
            'FILL OPTIONS Group
            .TerCopyToCode = mcboFillTerrCodes.Text.Trim 'FILL OPTIONS: Copy To Tab
            .TerCopyToDesc = txtFillTerrDesc.Text.Trim  'FILL OPTIONS: Copy To Tab
            If bFillPressed Then .TerCodeSearchTerFromInTable = txtTerFrom.Text.Trim Else .TerCodeSearchTerFromInTable = txtFrom.Text.Trim
            .TerCopyToCode = mcboFillTerrCodes.Text.Trim
            .TerCopyToDesc = txtFillTerrDesc.Text.Trim
        End With
    End Sub

    Private Sub SetOptionalCopyToCriteria()
        On Error Resume Next
        With cOptionalCriteria
            .Clear()
            .SearchType = IIf(txtFrom.Text.Trim = "", SearchType.PrimaryPrice.ToString, SearchType.CopiedPrice.ToString).ToString
            .PricingType = PrcType.Copied.ToString
            'LOAD PRICE Group
            .TerFromCode = mcboTerritoryCodes.Data.Rows(Me.mcboTerritoryCodes.SelectedIndex)(0).ToString.Trim ' LOAD PRICE: mcboTerritoryCodes
            .TerCopiedFromCode = mcboTerritoryCodes.Data.Rows(Me.mcboTerritoryCodes.SelectedIndex)(3).ToString.Trim ' LOAD PRICE: txtFrom
            .TerFromDesc = mcboTerritoryCodes.Data.Rows(Me.mcboTerritoryCodes.SelectedIndex)(1).ToString.Trim ' LOAD PRICE: txtTerrDesc
            .ProdCat = cboCategoryCodes.Text.Trim 'LOAD PRICE: cboCategoryCodes
            .OnPriceList = cboOnPriceList.Text.Trim 'LOAD PRICE: cboOnPriceList
            'FILL OPTIONS Group
            .TerCopyToCode = mcboFillTerrCodes.Text.Trim 'FILL OPTIONS: Copy To Tab
            .TerCopyToDesc = txtFillTerrDesc.Text.Trim  'FILL OPTIONS: Copy To Tab
            .TerCodeSearchTerFromInTable = txtTerFrom.Text.Trim
            .TerCopyToCode = mcboFillTerrCodes.Text.Trim
            .TerCopyToDesc = txtFillTerrDesc.Text.Trim

        End With
    End Sub


    Private Sub SetOptionalCriteriaSearch()
        On Error Resume Next

        With cOptionalCriteriaSearch
            .Clear()

            .ProdCat = cboCategoryCodes.Text.Trim 'LOAD PRICE: cboCategoryCodes
            .OnPriceList = cboOnPriceList.Text.Trim
            .HasTerrPrice = cboHasTerrPrice.Text.Trim

            If tabOptions.SelectedIndex = 1 Then
                .SearchType = SearchType.CopiedPrice.ToString
                .TerFromCode = mcboFillTerrCodes.Text.Trim
                .TerFromDesc = txtFillTerrDesc.Text.Trim
                .TerCodeSearchTerFromInTable = txtTerFrom.Text.Trim
            ElseIf txtFrom.Text.Trim > "" Then
                .SearchType = SearchType.CopiedPrice.ToString
                .TerFromCode = mcboTerritoryCodes.Data.Rows(mcboTerritoryCodes.SelectedIndex)(0).ToString.Trim
                .TerFromDesc = txtTerrDesc.Text.Trim
                .TerCodeSearchTerFromInTable = txtFrom.Text.Trim
            Else
                .SearchType = SearchType.PrimaryPrice.ToString
                .TerFromDesc = txtTerrDesc.Text.Trim
                .TerCodeSearchTerFromInTable = ""
                .TerFromCode = mcboTerritoryCodes.Data.Rows(mcboTerritoryCodes.SelectedIndex)(0).ToString.Trim
                .TerFromDesc = mcboTerritoryCodes.Data.Rows(mcboTerritoryCodes.SelectedIndex)(1).ToString.Trim
            End If

        End With
    End Sub


    Private Sub FillAllItemList(ByVal bSetOptionalCriteria As Boolean, ByVal SearchCriteria As OptionalCriteria)
        'For Loading the data: 
        ' - DataReader holds the data from the SQL database when we make the SQL call
        ' - ItemPricingList is the collection of ItemPricingObj that will hold the data from the DataReader
        Dim rd As SqlDataReader
        Dim itmprclst As New ItemPricingList


        'For the SQL Call we have Parameters that are sent, then the SQL call is made and put into the ItemPricingObjects and List
        ' - Fill the parameters with a value or empty string 
        ' - TerCodeSearchFrom is the Territory Code that use initially selected from the "Load A Price" Territory Code drop down
        ' - TerCodeSearchTerFromInTable, this value is needed for the Query only.  We want the initial TerCodeTo to be the same as TerCodeFrom
        '   That way, when the user presses Fill, it uses this value to populate the Ter column in the datagridview.  But for the query, we 
        '   need "" if the TerCodeFrom is a Primary since Primarys have no Ter_From and sending a value in TerCodeTo would put a value to search for in
        '   the Where Clause, not good.  
        With SearchCriteria
            rd = BusObj.GetSearchItems(.SearchType, .ProdCat, .TerFromCode, .OnPriceList, .TerCodeSearchTerFromInTable, .HasTerrPrice, cn)
            itmprclst = BusObj.PopulateSearchItems(rd)
        End With

        'We will use a Binding Source to bind to the Grid with the ItemPriceList collection
        With ItemPricingObjBindingSource
            .DataSource = Nothing
            .DataSource = itmprclst
        End With

        'Finally, set the BindingSource as the datasource to the grid ...
        DataGridView1.DataSource = ItemPricingObjBindingSource
        cItemPricingList = itmprclst

        'We have a Filter DropDown Control on the UI and below it is bound to the same BindingSource with only ItemNo on the list
        With cboFilter
            .DataSource = ItemPricingObjBindingSource
            .DisplayMember = "ItemNo"
        End With

        'Handles some of the UI Controls:
        'By default, first pricing type is Copied with the TerFromSearch as the Copy To pricing
        ' - Set status bar at the bottom of the form to default 
        ' - Show a count of the number of items we have retrieved 


        With SearchCriteria

            mcboTerritoryCodes.Text = SearchCriteria.TerFromCode
            txtFrom.Text = .TerCodeSearchTerFromInTable
            txtTerrDesc.Text = .TerFromDesc
            lblPricingType.Text = ""    '"Pricing Type: Copied"
            lblCount.Text = itmprclst.Count.ToString & " Items Retrieved  "
            txtTerCode.Text = SearchCriteria.TerFromCode
            txtTerFrom.Text = SearchCriteria.TerCodeSearchTerFromInTable
        End With





        SearchCriteria.PricingType = PrcType.Copied.ToString

        FormatPriceType(SearchCriteria.PricingType.ToString)

    End Sub

    'Private Sub FillAllItemList(cOptionalCriteria As OptionalCriteria)
    '    Dim rd As SqlDataReader
    '    Dim itmprclst As New ItemPricingList
    '    Dim prodcat As String = ""
    '    Dim terrcode As String = ""
    '    Dim terrfrom As String = ""
    '    Dim onpricelist As String = ""

    '    'DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

    '    With cOptionalCriteria
    '        .Clear()
    '        .ProdCat = Me.cboCategoryCodes.Text
    '        .TerCodeSearchFrom = mcboTerritoryCodes.Text
    '        If mcboTerritoryCodes.SelectedIndex = -1 Then
    '            .TerFromCode = ""
    '        Else
    '            .TerFromCode = mcboTerritoryCodes.Data.Rows(Me.mcboTerritoryCodes.SelectedIndex)(3).ToString
    '        End If
    '        .OnPriceList = Me.cboOnPriceList.Text.ToString
    '        rd = BusObj.GetSearchItems(cOptionalCriteria.SearchType, .ProdCat, .TerCodeSearchFrom, .OnPriceList, .TerFromCode.Trim, cn)
    '        itmprclst = BusObj.PopulateSearchItems(rd)
    '    End With

    '    With ItemPricingObjBindingSource
    '        .DataSource = Nothing
    '        .DataSource = itmprclst
    '    End With

    '    DataGridView1.DataSource = ItemPricingObjBindingSource
    '    cItemPricingList = itmprclst

    '    With cboFilter
    '        .DataSource = ItemPricingObjBindingSource
    '        .DisplayMember = "ItemNo"
    '    End With

    '    If cOptionalCriteria.MarkupType = MarkupType.Zone.ToString Then
    '        lblPricingType.Text = "" '"Pricing Type: Add New Items to Pricing"
    '        'ElseIf cOptionalCriteria.PricingType = PrcType.Update.ToString Then
    '        '    lblPricingType.Text = "Pricing Type: Update an Existing Pricing"
    '    ElseIf Not cOptionalCriteria.TerFromCode.Trim = "" Then
    '        rbCopiedTerritory.Checked = True
    '        txtTerCode.Enabled = True
    '        lblPricingType.Text = "" '"Pricing Type: New Copied Pricing"
    '    Else
    '        rbPrimaryTerritory.Checked = True
    '        'txtTerFrom.Enabled = False
    '        lblPricingType.Text = "" '"Pricing Type: New Primary Pricing"
    '    End If

    '    rbUpdate.Checked = True

    '    lblCount.Text = itmprclst.Count.ToString & " Items Retrieved  "
    'End Sub

    Private Function IsAnyRowChecked(ByVal dgv As DataGridView) As Boolean
        For Each row As DataGridViewRow In dgv.Rows
            If CBool(row.Cells("Selected").Value) = True Then
                Return True
                Exit Function
            End If
        Next
        Return False
    End Function

    Private Sub DeleteItemPricing()
        Me.Validate()
        Dim itm As ItemPricingObj
        Dim itempricecount As Integer
        Dim state As String

        For Each itm In cItemPricingList
            If itm.Selected = True Then
                itempricecount = BusObj.CheckCopiedItemState(itm.ItemNo, itm.TerCode, itm.TerFrom, cn)
                If itempricecount > 0 Then
                    state = "Deleted"
                    cItemPricingObj.SaveItemPricing(itm.ItemNo, itm.TerCode, itm.OriginalPriceColor, itm.OriginalPriceRococo, itm.OriginalPriceDetailStain, "", itm.TerDescription, state, cn)
                End If
            End If
        Next
    End Sub

    Private Sub SaveItemPricing(ByVal itmPrcLst As ItemPricingList)
        'Deletes the existing price records, then uses SQLBulkCopy to add them back with the new pricing.  

        Dim selected As Boolean = IsAnyRowChecked(Me.DataGridView1)
        Dim res As Integer
        If selected = False Then
            res = MsgBox("No items have been selected.  Select the items you would like to update.", MsgBoxStyle.OkOnly, "Round Pricing")
            Exit Sub
        End If

        Me.Validate()

        With cOptionalCriteria
            dtDelete = GetItemPrcLevelFromObject(cItemPricingList)
            dtSave = GetSQLBulkCopyDataTableFromObject(cItemPricingList, .TerCopyToCode.Trim, .TerCopyToDesc.Trim, .TerFromCode.Trim)
        End With

        Try
            'BusObj.DeleteA4GLIdentitybyTVP(dtDelete, cn)
            BusObj.DeleteItemsbyTVP(dtDelete, cn)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Using s As SqlBulkCopy = New SqlBulkCopy(cn)
            s.DestinationTableName = "OEPRCCUS_MAZ"
            s.WriteToServer(dtSave)
            s.Close()
        End Using

        Try
            FillTerritoryCodeList()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SaveAdvancedItemPricing(ByVal itmPrcLst As ItemPricingList)
        'Deletes the existing price records, then uses SQLBulkCopy to add them back with the new pricing.  

        Dim selected As Boolean = IsAnyRowChecked(Me.DataGridView1)
        If selected = False Then
            MsgBox("No items have been selected.  Select the items to update.", MsgBoxStyle.OkOnly, "Select Items")
            Exit Sub
        End If

        Me.Validate()

        dtDelete = GetItemA4GLIdentityFromObject(cItemPricingList)
        dtSave = GetSQLBulkCopyDataTableFromObject(cItemPricingList)

        BusObj.DeleteA4GLIdentitybyTVP(dtDelete, cn)

        Using s As SqlBulkCopy = New SqlBulkCopy(cn)
            s.DestinationTableName = "OEPRCCUS_MAZ"
            s.WriteToServer(dtSave)
            s.Close()
        End Using

        Try
            FillTerritoryCodeList()

        Catch ex As Exception

        End Try
    End Sub

    Private Function RetrieveAdvancedItemPricing(ByVal itmPrcLst As ItemPricingList) As DataTable
        Dim dt As DataTable
        Dim prc_level As String = ""
        Dim item_no As String = itmPrcLst(0).ItemNo.ToString.Trim

        For Each itm As ItemPricingObj In itmPrcLst
            If itm.Selected = True Then
                If prc_level = "" Then
                    prc_level = "'" & itm.TerCode & "'"
                Else
                    prc_level = prc_level & ", '" & itm.TerCode & "'"
                End If
            End If
        Next

        dt = BusObj.GetItem(item_no, prc_level, cn)
        Return dt
    End Function
#End Region

#Region "  Fill Pricing   "

    Private Sub btnFillPricing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFillPricing.Click
        bFillPressed = True : bLoadPressed = False
        Dim selected As Boolean = IsAnyRowChecked(Me.DataGridView1)
        If selected = False Then
            MsgBox("No items have been selected.  Select the items to update.", MsgBoxStyle.OkOnly, "Select Items")
            Exit Sub
        End If


        If cItemPricingList.Count = 0 Then Exit Sub
        bByPass = True

        Cursor = Cursors.WaitCursor

        CollectUIFillData()

        'Determine the Pricing Type we are doing.  Either Copied, Custom Percent or Custom Dollar ....
        With cOptionalCriteria
            If .MarkupType = MarkupType.Zone.ToString Then
                DoZonePercentMarkup()
            ElseIf .MarkupType = MarkupType.Direct.ToString Then
                DoDirectMarkup()
            ElseIf .PricingType = PrcType.CustomPercent.ToString OrElse
                   .PricingType = PrcType.CustomDollar.ToString OrElse
                   .PricingType = PrcType.Advanced.ToString And rbCustomFill.Checked Then
                DoFillWithMarkup()
            ElseIf .PricingType = PrcType.Copied.ToString Then
                DoStandardMarkup()
            ElseIf .PricingType = PrcType.Primary.tostring Then
                DoStandardMarkup()
            End If

        End With

        Cursor = Cursors.Default

        cOptionalCriteria.IsFillPressed = True
        btnSave.Enabled = cOptionalCriteria.IsFillPressed
        bByPass = False
    End Sub
    Private Sub DoZonePercentMarkup()
        On Error Resume Next
        'Zone percents are added by reading the table oeprczon_mas and returning the markup percent.    
        'Zone 2 – 18% 
        'Zone 3 – 22%
        'Zone 4 - 33%
        'Zone 5 – 43%
        'Zone 6 – 50%
        'Zone 7 – the rounded Zone 4 price * 1.25
        'Zone 8 – the rounded Zone 6 price * 1.25


        Dim sSQL As String = "Select ter_zone, frt_markup, ter_from, ter_desc from oeprczon_mas "
        Dim dtZoneMarkup As DataTable = BusObj.GetZonePriceList(sSQL, "ZoneMarkupList", cn)
        Dim dvZoneMarkup As DataView = dtZoneMarkup.DataSet.DefaultViewManager.CreateDataView(dtZoneMarkup)

        With DataGridView1
            For Each rw As DataGridViewRow In .Rows
                If Convert.ToBoolean(rw.Cells("Selected").Value) = True Then
                    If cOptionalCriteria.PricingType <> PrcType.Advanced.ToString Then rw.Cells("TerCode").Value = cOptionalCriteria.TerZoneCode
                    Dim tercode As String
                    Dim terfrom As String
                    Dim terdesc As String
                    Dim filtertercode As String
                    Dim frtmarkuppct As Decimal
                    Dim basedon As Decimal

                    If cOptionalCriteria.FillType <> "Ter Code" Then
                        tercode = cOptionalCriteria.TerCopyToCode
                        terfrom = IIf(tercode.Substring(0, 2) = "00", "", cOptionalCriteria.TerFromCode).ToString
                        terdesc = cOptionalCriteria.TerCopyToDesc
                        filtertercode = IIf(terfrom.Trim = "", tercode, terfrom).ToString
                        dvZoneMarkup.RowFilter = "Trim(ter_zone) = '" & filtertercode & "'"
                        frtmarkuppct = CDec(dvZoneMarkup.Item(0).Item(1))
                        ' if zone 007 or 008 use basedon 

                    Else
                        tercode = rw.Cells("TerCode").Value.ToString
                        terfrom = rw.Cells("TerFrom").Value.ToString
                        terdesc = rw.Cells("TerDesc").Value.ToString
                        filtertercode = rw.Cells("TerCode").Value.ToString
                        dvZoneMarkup.RowFilter = "Trim(ter_zone) = '" & filtertercode & "'"
                        frtmarkuppct = CDec(dvZoneMarkup.Item(0).Item(1))
                        If tercode = "007" Then
                            basedon = CDec(0.33)
                        ElseIf tercode = "008" Then
                            basedon = CDec(0.5)
                        End If
                    End If

                    If Not (rw.Cells("ItemLocPriceColor").Value Is Nothing) Then
                        rw.Cells("ActivePriceColor").Value =
                            GetZoneMarkupPrice(frtmarkuppct, Math.Round(Convert.ToDecimal(rw.Cells("ItemLocPriceColor").Value)),
                            Convert.ToDecimal(rw.Cells("ItemLocPriceColor").Value), rw.Cells("ProdCategory").ToString, basedon)
                    Else
                        rw.Cells("OriginalPriceColor").Value = ""
                    End If

                    If Not (rw.Cells("ItemLocPriceRococo").Value Is Nothing) Then
                        rw.Cells("ActivePriceRococo").Value =
                            GetZoneMarkupPrice(frtmarkuppct, Math.Round(Convert.ToDecimal(rw.Cells("ItemLocPriceRococo").Value)),
                            Convert.ToDecimal(rw.Cells("ItemLocPriceRococo").Value), rw.Cells("ProdCategory").ToString, basedon)
                    Else
                        rw.Cells("OriginalPriceRococo").Value = ""
                    End If

                    If Not (rw.Cells("ItemLocPriceDetailStain").Value Is Nothing) Then
                        rw.Cells("ActivePriceDetailStain").Value =
                            GetZoneMarkupPrice(frtmarkuppct, Math.Round(Convert.ToDecimal(rw.Cells("ItemLocPriceDetailStain").Value)),
                            Convert.ToDecimal(rw.Cells("ItemLocPriceDetailStain").Value), rw.Cells("ProdCategory").ToString, basedon)
                    Else
                        rw.Cells("OriginalPriceDetailStain").Value = ""
                    End If

                    rw.Cells("TerCode").Value = tercode
                    rw.Cells("TerFrom").Value = terfrom  '.TercodeSearch  
                    rw.Cells("TerDesc").Value = terdesc
                    rw.Cells("LastDate").Value = Now.ToString("MM/dd/yyyy")

                End If
            Next
        End With

    End Sub

    Private Sub DoStandardMarkup()

        Try
            With DataGridView1
                For Each rw As DataGridViewRow In .Rows

                    If Not (rw.Cells("ItemLocPriceDetailStain").Value Is Nothing) Then
                        rw.Cells("ActivePriceColor").Value = rw.Cells("OriginalPriceColor").Value
                        'rw.Cells("ActivePriceColor").Value = GetStardardMarkupPrice(frtmarkuppct, Convert.ToDecimal(rw.Cells("ItemLocPriceColor").Value), Convert.ToDecimal(rw.Cells("ItemLocPriceRococo").Value))
                    Else
                        rw.Cells("OriginalPriceDetailStain").Value = "" : rw.Cells("ActivePriceColor").Value = ""
                    End If

                    If Not (rw.Cells("ItemLocPriceDetailStain").Value Is Nothing) Then
                        rw.Cells("ActivePriceRococo").Value = rw.Cells("OriginalPriceRococo").Value
                        'rw.Cells("ActivePriceRococo").Value = GetStardardMarkupPrice(frtmarkuppct, Convert.ToDecimal(rw.Cells("ItemLocPriceRococo").Value), Convert.ToDecimal(rw.Cells("ItemLocPriceRococo").Value))
                    Else
                        rw.Cells("OriginalPriceDetailStain").Value = "" : rw.Cells("ActivePriceRococo").Value = ""
                    End If

                    If Not (rw.Cells("ItemLocPriceDetailStain").Value Is Nothing) Then
                        rw.Cells("ActivePriceDetailStain").Value = rw.Cells("OriginalPriceDetailStain").Value
                        'rw.Cells("ActivePriceDetailStain").Value = GetStardardMarkupPrice(frtmarkuppct, Convert.ToDecimal(rw.Cells("ItemLocPriceDetailStain").Value), Convert.ToDecimal(rw.Cells("ItemLocPriceRococo").Value))
                    Else
                        rw.Cells("OriginalPriceDetailStain").Value = "" : rw.Cells("ActivePriceDetailStain").Value = ""
                    End If


                    With cOptionalCriteria


                        'If Not .PricingType.ToString = PrcType.Advanced.ToString Then
                        '    rw.Cells("TerCode").Value = IIf(.TerCopyToCode = "", Me.txtTerCode.Text.Trim, .TerCopyToCode)
                        '    rw.Cells("TerFrom").Value = .TerFromCode   'dvZoneMarkup(0)("ter_from")
                        '    rw.Cells("TerDesc").Value = .TerCopyToDesc
                        '    rw.Cells("LastDate").Value = Now.ToString("MM/dd/yyyy")
                        'End If
                        If .PricingType.ToString = PrcType.Copied.ToString Then
                            rw.Cells("TerCode").Value = IIf(.TerCopyToCode = "", Me.txtTerCode.Text.Trim, .TerCopyToCode)
                            rw.Cells("TerFrom").Value = .TerFromCode   'dvZoneMarkup(0)("ter_from")
                            rw.Cells("TerDesc").Value = .TerCopyToDesc
                            rw.Cells("LastDate").Value = Now.ToString("MM/dd/yyyy")
                        ElseIf .PricingType = PrcType.Primary.tostring Then
                            rw.Cells("TerCode").Value = .TerCode
                            rw.Cells("TerFrom").Value = .TerFromCode
                            rw.Cells("TerDesc").Value = .TerFromDesc
                            rw.Cells("LastDate").Value = Now.ToString("MM/dd/yyyy")
                        End If
                    End With
                Next
            End With

        Catch ex As Exception

        End Try



    End Sub
    Private Sub DoDirectMarkup()
        Try

            With DataGridView1
                ' Populate the values based on the type of Update or New Pricing we are doing ...
                For Each rw As DataGridViewRow In .Rows
                    If CBool(rw.Cells("Selected").Value) = True Then
                        rw.Cells("ActivePriceColor").Value = rw.Cells("ItemLocPriceColor").Value
                        rw.Cells("ActivePriceRococo").Value = rw.Cells("ItemLocPriceRococo").Value
                        rw.Cells("ActivePriceDetailStain").Value = rw.Cells("ItemLocPriceDetailStain").Value
                        If cOptionalCriteria.PricingType <> PrcType.Advanced.ToString Then
                            rw.Cells("TerFrom").Value = cOptionalCriteria.TerCodeSearchFrom 'Me.multicboTerritoryCodes.Text
                            rw.Cells("TerCode").Value = Me.mcboTerritoryCodes.Text.Trim
                            rw.Cells("TerDesc").Value = Me.txtTerrDesc.Text
                        End If
                        rw.Cells("LastDate").Value = CDate(Now)
                    End If
                Next

            End With

        Catch ex As Exception

        End Try

    End Sub
    Private Sub DoCopyFrom()
        Try

            With DataGridView1
                ' Populate the values based on the type of Update or New Pricing we are doing ...
                For Each rw As DataGridViewRow In .Rows
                    If CBool(rw.Cells("Selected").Value) = True Then
                        rw.Cells("ActivePriceColor").Value = rw.Cells("OriginalPriceColor").Value
                        rw.Cells("ActivePriceRococo").Value = rw.Cells("OriginalPriceRococo").Value
                        rw.Cells("ActivePriceDetailStain").Value = rw.Cells("OriginalPriceDetailStain").Value
                        rw.Cells("TerFrom").Value = cOptionalCriteria.TerCodeSearchFrom 'Me.multicboTerritoryCodes.Text
                        rw.Cells("TerCode").Value = Me.mcboFillTerrCodes.Text
                        rw.Cells("TerDesc").Value = Me.txtFillTerrDesc.Text
                        rw.Cells("LastDate").Value = CDate(Now)
                    End If
                Next

            End With

        Catch ex As Exception

        End Try

    End Sub

    Private Sub DoFillWithMarkup()
        ' See if everything is set to proceed, i.e. items checked, options complete etc...
        Dim selected As Boolean = IsAnyRowChecked(Me.DataGridView1)

        If selected = False Then
            MsgBox("No items have been selected.  Select the items to update.", MsgBoxStyle.OkOnly, "Select Items")
            Exit Sub
        End If

        ' Go to the fill ...
        With cOptionalCriteria
            If .PricingType = PrcType.CustomDollar.ToString OrElse .PricingType = PrcType.Advanced.ToString And rbCustomFill.Checked And rbAmountFill.Checked Then
                TryFill(.NatAmount, .ClrAmount, .DtlAmount, .PricingType, .MarkupType, .FillType)
            ElseIf .PricingType = PrcType.CustomPercent.ToString OrElse .PricingType = PrcType.Advanced.ToString And rbCustomFill.Checked And rbPercentFill.Checked Then
                TryFill(.NatPercent, .ClrPercent, .DtlPercent, .PricingType, .MarkupType, .FillType)
            End If
        End With

    End Sub

    Private Function GetZoneMarkupPrice(frtmarkuppct As Decimal, itemlocprice As Decimal, markupprice As Decimal, prod_cat As String, Optional basedon As Decimal = CDec(0)) As Decimal
        Dim frtmarkupamt As Decimal
        Dim retPrice As Decimal
        Dim basedonamt As Decimal

        '' terr 007 and 008 are handled differently.  
        '' zone 007 = item price for 004 * 1.25
        '' zone 008 = item price for 006 * 1.25

        '' See if we have either zone 007 or 008
        'If tercode = "007" Then

        'End If


        Select Case prod_cat
            Case "150", "152", "390" 'Product Categories here are accessories or a Mfg Item that is the same price for all Zones, these items us the Price in MACOLA, itemlocprice.  
                Return itemlocprice
            Case Else
                If basedon = 0 Then
                    frtmarkupamt = itemlocprice * frtmarkuppct
                    retPrice = markupprice + frtmarkupamt
                Else
                    basedonamt = Math.Round(Convert.ToDecimal(markupprice * (1 + basedon))) ' first use this calculation to get the 004 price to calculate 007 or 006 price for 008
                    retPrice = Convert.ToDecimal(basedonamt * (1 + frtmarkuppct))
                    'retPrice = markupprice + frtmarkupamt
                End If

                Return retPrice
        End Select

    End Function

    Private Sub CollectUIFillData()
        With cOptionalCriteria
            If Not .PricingType = PrcType.Advanced.ToString Then

                .Clear()

                'Retrieve the user's selected Pricing Options, default is Copied ...
                .PricingType = PrcType.Copied.ToString ' We're only setting a default value here, if it's a CustomFill, it changes with the next line ...
                If rbCustomFill.Checked Then
                    If rbPercentFill.Checked Then
                        .PricingType = PrcType.CustomPercent.ToString
                    Else
                        .PricingType = PrcType.CustomDollar.ToString
                    End If
                    .TerCopyToCode = txtTerCode.Text.Trim
                    .TerCopyToDesc = txtTerrDesc.Text.Trim
                    .TerFromCode = txtTerFrom.Text.Trim
                ElseIf rbZoneFill.Checked Then
                    .MarkupType = MarkupType.Zone.ToString
                    .TerZoneCode = cboZonePrc.Text.Trim
                    .TerCopyToCode = txtTerCode.Text.Trim
                    .TerCopyToDesc = txtTerrDesc.Text.Trim
                    .TerFromCode = cboZonePrc.Text
                ElseIf rbDirectMacolaPrc.Checked Then
                    .MarkupType = MarkupType.Direct.ToString
                ElseIf txtTerCode.Text = "002" Or txtTerCode.Text = "003" Or txtTerCode.Text = "004" Or txtTerCode.Text = "005" Or txtTerCode.Text = "006" Or txtTerCode.Text = "007" Or txtTerCode.Text = "008" Or txtTerCode.Text = "009" Then
                    .PricingType = PrcType.Primary.ToString
                    .TerCopyToCode = txtTerCode.Text.Trim
                    .TerFromCode = ""
                    .TerCopyToDesc = txtTerrDesc.Text.Trim

                End If

                'once we have the TerCopyToCode (the TerrCode that will be saved) determine if it's a Primary or Copied.
                'then set the TerCodeFrom to "" we are saving a Primary or set the TerCodeFrom to the Primary TerCode if it's Copied 
                'so we have a record of which Primary the Copied Territory pricing originated from.
                'First, set some variables to make things more managable 

                Dim bSaveCodeIsPrimary As Boolean
                Dim bSearchCodeIsPrimary As Boolean

                If txtFrom.Text.Trim = "" Then   '.TerCopyToCode = "002" Or .TerCopyToCode = "003" Or .TerCopyToCode = "004" Or .TerCopyToCode = "005" Or .TerCopyToCode = "006" Then
                    bSaveCodeIsPrimary = True
                Else
                    bSearchCodeIsPrimary = False
                End If

                'set the terr code optional criteria

                '.TerFromCode = txtTerFrom.Text.Trim
                '.TerCopyToCode = txtTerCode.Text.Trim
                '.TerCopyToDesc = txtTerrDesc.Text.Trim

            End If
            'If we are doing a markup by Percent or Dollar Amt, get the Natural, Color and Detail markup values
            If Not (.PricingType.ToString = PrcType.Copied.ToString) Then
                If .PricingType = PrcType.CustomPercent.ToString OrElse (.PricingType = PrcType.Advanced.ToString And rbCustomFill.Checked And rbPercentFill.Checked) Then
                    .NatPercent = CDbl(IIf(txtNaturalMarkup.Text = "", 0, txtNaturalMarkup.Text))
                    .ClrPercent = CDbl(IIf(txtColorMarkup.Text = "", 0, txtColorMarkup.Text))
                    .DtlPercent = CDbl(IIf(txtDetailMarkup.Text = "", 0, txtDetailMarkup.Text))
                ElseIf .PricingType = PrcType.CustomDollar.ToString OrElse (.PricingType = PrcType.Advanced.ToString And rbCustomFill.Checked And rbAmountFill.Checked) Then
                    .NatAmount = CDbl(IIf(txtNaturalMarkup.Text = "", 0, txtNaturalMarkup.Text))
                    .ClrAmount = CDbl(IIf(txtColorMarkup.Text = "", 0, txtColorMarkup.Text))
                    .DtlAmount = CDbl(IIf(txtDetailMarkup.Text = "", 0, txtDetailMarkup.Text))
                ElseIf .PricingType = PrcType.Advanced.ToString And cboZonePrc.Text = "Ter Code" Then
                    .FillType = "Ter Code"
                End If
            End If

            If rbZoneFill.Checked Then
                .MarkupType = MarkupType.Zone.ToString
            ElseIf rbDirectMacolaPrc.Checked Then
                .MarkupType = MarkupType.Direct.ToString
            End If

        End With
    End Sub

    Private Function GetSQLBulkCopyDataTableFromObject(ByVal cItemPricingList As ItemPricingList) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("item_no", GetType(String))
        dt.Columns.Add("prc_level", GetType(String))
        dt.Columns.Add("prc_natural", GetType(Decimal))
        dt.Columns.Add("prc_color", GetType(Decimal))
        dt.Columns.Add("prc_detail", GetType(Decimal))
        dt.Columns.Add("ter_from", GetType(String))
        dt.Columns.Add("ter_desc", GetType(String))
        dt.Columns.Add("year_cd", GetType(Integer))
        dt.Columns.Add("createdate", GetType(Date))
        dt.Columns.Add("lastdate", GetType(Date))

        For Each o As ItemPricingObj In cItemPricingList
            If o.Selected Then
                dt.Rows.Add(o.ItemNo, o.TerCode, o.ActivePriceRococo, o.ActivePriceColor, o.ActivePriceDetailStain, o.TerFrom, o.TerDescription, Year(Now), Now, Now)
            End If
        Next
        Return dt
    End Function

    Private Function GetSQLBulkCopyDataTableFromObject(ByVal cItemPricingList As ItemPricingList, ByVal prc_level As String, ByVal ter_desc As String, ByVal ter_from As String) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("item_no", GetType(String))
        dt.Columns.Add("prc_level", GetType(String))
        dt.Columns.Add("prc_natural", GetType(Decimal)) 'Rococo
        dt.Columns.Add("prc_color", GetType(Decimal)) 'Natural and Color Stain
        dt.Columns.Add("prc_detail", GetType(Decimal)) 'Detail Stain
        dt.Columns.Add("ter_from", GetType(String))
        dt.Columns.Add("ter_desc", GetType(String))
        dt.Columns.Add("year_cd", GetType(Integer))
        dt.Columns.Add("createdate", GetType(Date))
        dt.Columns.Add("lastdate", GetType(Date))

        For Each o As ItemPricingObj In cItemPricingList
            If o.Selected Then
                dt.Rows.Add(o.ItemNo.Trim, prc_level.Trim, o.ActivePriceRococo, o.ActivePriceColor, o.ActivePriceDetailStain, ter_from.Trim, ter_desc.Trim, Year(Now), Now, Now)
            End If
        Next
        Return dt
    End Function
    Private Function GetItemPrcLevelFromObject(ByVal cItemPricingList As ItemPricingList) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("item_no", GetType(String))
        dt.Columns.Add("prc_level", GetType(String))
        dt.Columns.Add("ter_from", GetType(String))
        dt.Columns.Add("A4GLIdentity", GetType(String))
        For Each o As ItemPricingObj In cItemPricingList
            If o.Selected = True Then
                dt.Rows.Add(o.ItemNo, o.TerCode, o.TerFrom, o.A4GLIdentity)
            End If
        Next
        Return dt
    End Function
    Private Function GetItemPrcLevelFromObject(ByVal cItemPricingList As ItemPricingList, ByVal tercode As String, ByVal terfrom As String) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("item_no", GetType(String))
        dt.Columns.Add("prc_level", GetType(String))
        dt.Columns.Add("ter_from", GetType(String))

        For Each o As ItemPricingObj In cItemPricingList
            If o.Selected = True Then
                dt.Rows.Add(o.ItemNo.Trim, tercode.Trim, terfrom.Trim)
            End If
        Next
        Return dt
    End Function


    Private Function GetItemA4GLIdentityFromObject(cItemPricingList As ItemPricingList) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("A4GLIdentity", GetType(Integer))

        For Each o As ItemPricingObj In cItemPricingList
            If o.Selected = True Then
                dt.Rows.Add(o.A4GLIdentity)
            End If
        Next
        Return dt
    End Function

    Private Sub TryFill(AmtNtl As Double, AmtClr As Double, AmtDtl As Double, PricingType As String, MarkupType As String, Optional FillType As String = "")
        Dim rw As New DataGridViewRow
        ' Generate the percentages we'll need to use to calculate the price ...

        Dim prct_ntl As Double = CDbl(AmtNtl * 0.01) + 1.0
        Dim prct_clr As Double = CDbl(AmtClr * 0.01) + 1.0
        Dim prct_dtl As Double = CDbl(AmtDtl * 0.01) + 1.0


        With DataGridView1
            ' Populate the values based on the type of Update or New Pricing we are doing ...
            For Each rw In .Rows
                If CBool(rw.Cells("Selected").Value) = True Then
                    rw.Cells("ActivePriceColor").Value = IIf(rbPercentFill.Checked,
                                                               CDec(rw.Cells("OriginalPriceColor").Value) * CDec(prct_ntl),
                                                               CDec(rw.Cells("OriginalPriceColor").Value) + CDec(cOptionalCriteria.NatAmount))
                    rw.Cells("ActivePriceRococo").Value = IIf(rbPercentFill.Checked,
                                                             CDec(rw.Cells("OriginalPriceRococo").Value) * CDec(prct_clr),
                                                             CDec(rw.Cells("OriginalPriceRococo").Value) + CDec(cOptionalCriteria.ClrAmount))
                    rw.Cells("ActivePriceDetailStain").Value = IIf(rbPercentFill.Checked,
                                                              CDec(rw.Cells("OriginalPriceDetailStain").Value) * CDec(prct_dtl),
                                                              CDec(rw.Cells("OriginalPriceDetailStain").Value) + CDec(cOptionalCriteria.DtlAmount))

                    'Advanced doesn't have these values changed ...
                    If Not cOptionalCriteria.PricingType = PrcType.Advanced.ToString Then
                        rw.Cells("TerFrom").Value = cOptionalCriteria.TerFromCode
                        rw.Cells("TerCode").Value = cOptionalCriteria.TerCopyToCode
                        rw.Cells("TerDesc").Value = cOptionalCriteria.TerCopyToDesc
                        rw.Cells("LastDate").Value = Now.ToString("MM/dd/yyyy")
                    End If

                End If
            Next

        End With

    End Sub

    Private Function FillCalculation(Value As Decimal, pctIncrease As Decimal) As Decimal
        Dim returnAmount As Decimal = 0
        If cOptionalCriteria.FillType = "Percent" Then
            returnAmount = Value * pctIncrease
        Else
            returnAmount = pctIncrease
        End If

        Return returnAmount

    End Function

#End Region

#Region "  Format Methods   "

    Private Sub SetDataGridBackgroundColor()
        Dim dgv As DataGridView = CType(DataGridView1, DataGridView)
        For Each rw As DataGridViewRow In dgv.Rows
            If Convert.ToDecimal(rw.Cells("OriginalPriceColor").Value) <> Convert.ToDecimal(rw.Cells("CopiedPriceColor").Value) OrElse
               Convert.ToDecimal(rw.Cells("OriginalPriceRococo").Value) <> Convert.ToDecimal(rw.Cells("CopiedPriceRococo").Value) OrElse
               Convert.ToDecimal(rw.Cells("OriginalPriceDetailStain").Value) <> Convert.ToDecimal(rw.Cells("CopiedPriceDetailStain").Value) Then
                rw.DefaultCellStyle.BackColor = Color.PaleGoldenrod
                rw.Cells("OriginalPriceColor").Style.BackColor = readOnlyBackColorDark
                rw.Cells("OriginalPriceRococo").Style.BackColor = readOnlyBackColorDark
                rw.Cells("OriginalPriceDetailStain").Style.BackColor = readOnlyBackColorDark
                rw.Cells("ItemLocPriceColor").Style.BackColor = readOnlyBackColorDark
                rw.Cells("ItemLocPriceRococo").Style.BackColor = readOnlyBackColorDark
                rw.Cells("ItemLocPriceDetailStain").Style.BackColor = readOnlyBackColorDark
            End If
        Next


    End Sub

#End Region


    Private Sub FillTerritoryCodeList()
        'Retrieve the territory Codes
        Try
            dtTerrCodes = BusObj.GetTerritoryCodeList(cn)
            dtCopyFrom = dtTerrCodes.Copy
            'Add a blank row to the list
            Dim row As DataRow = dtTerrCodes.NewRow
            row("prc_level") = ""
            row("ter_desc") = ""
            row("copied_from") = ""
            row("ter_from") = ""
            dtTerrCodes.Rows.InsertAt(row, 0)
            Dim PrimaryKeyCopy(1) As DataColumn
            PrimaryKeyCopy(1) = dtCopyFrom.Columns("prc_level")
            dtCopyFrom.PrimaryKey = PrimaryKeyCopy

            Dim PrimaryKeyTer(1) As DataColumn
            PrimaryKeyTer(1) = dtCopyFrom.Columns("prc_level")
            dtCopyFrom.PrimaryKey = PrimaryKeyTer
            'clear the multicolumn combo
            If dtTerrCodes IsNot Nothing Then
                With mcboTerritoryCodes
                    .Data.Clear()
                    .Data.GetChanges()
                    .Data.AcceptChanges()
                    .Data = dtTerrCodes

                    .ViewColumn = 0
                    .Columns(1).Display = True
                    .Columns(2).Display = True
                    .Columns(3).Display = True
                    .Refresh()
                End With
                With mcboFillTerrCodes
                    .Data.Clear()
                    .Data.GetChanges()
                    .Data.AcceptChanges()
                    .Data = dtCopyFrom
                    .ViewColumn = 0
                    .Columns(1).Display = True
                    .Columns(2).Display = True
                    .Columns(3).Display = True
                    .Refresh()
                End With
            End If

        Catch ex As Exception
            MsgBox(My.Settings.DefaultDB & " does not have the correct tables for territory pricing.  Setting the database to DATA and reopening.")
            My.Settings.DefaultDB = "DATA"
            My.Settings.Save()
            LoadTerritoryPricing()
        End Try

    End Sub

    Private Sub fillCategoryCodes()
        cboCategoryCodes.Items.Clear()
        cboCategoryCodes.Items.Add("")
        Dim rd As SqlDataReader
        rd = BusObj.GetCategoryCodes(Convert.ToInt32(chkShowActiveOnly.Checked), cn)
        If rd IsNot Nothing Then
            While rd.Read
                Me.cboCategoryCodes.Items.Add(rd.Item(0).ToString)
            End While
        End If
        rd.Close()

    End Sub

    Private Sub fillZoneCodes()
        Dim ssql As String = "Select ter_zone  from OEPRCZON_MAS where ter_code like 'ZONE%'"
        cboZonePrc.Items.Clear()

        cboZonePrc.Items.Add("Ter Code") ' This means "use the zone that is in the row for each Item in the Data Grid View"

        Dim rd As SqlDataReader
        rd = BusObj.GetFreightMarkups(ssql, cn)
        While rd.Read
            cboZonePrc.Items.Add(rd.Item(0).ToString)
        End While
        rd.Close()

    End Sub

    Private Sub RoundPrices()
        Dim itm As ItemPricingObj

        Dim selected As Boolean = IsAnyRowChecked(Me.DataGridView1)
        If selected = False Then
            MsgBox("No items have been selected.  Select the items to update.", MsgBoxStyle.OkOnly, "Select Items")
            Exit Sub
        End If

        For Each itm In cItemPricingList
            If itm.Selected = True Then
                Dim d As Decimal
                d = itm.ActivePriceColor
                itm.ActivePriceColor = CDec((Math.Round(d).ToString))
                d = itm.ActivePriceRococo
                itm.ActivePriceRococo = CDec((Math.Round(d).ToString))
                d = itm.ActivePriceDetailStain
                itm.ActivePriceDetailStain = CDec((Math.Round(d).ToString))
                If cOptionalCriteria.MarkupType = MarkupType.Zone.ToString Then
                    d = itm.OriginalPriceColor
                    itm.OriginalPriceColor = CDec((Math.Round(d).ToString))
                    d = itm.OriginalPriceRococo
                    itm.OriginalPriceRococo = CDec((Math.Round(d).ToString))
                    d = itm.OriginalPriceDetailStain
                    itm.OriginalPriceDetailStain = CDec((Math.Round(d).ToString))
                End If
            End If
        Next

        Me.ValidateChildren()
        cItemPricingList.ResetBindings()

    End Sub

    Private Sub DataGridView1_BindingContextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.BindingContextChanged

    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick

        Try
            Dim dgv As DataGridView = CType(sender, DataGridView)
            Dim itemno As String = dgv.Rows(e.RowIndex).Cells(1).Value.ToString
            AddGridItemToFilterList(itemno)
            btnApplyFilter.Enabled = True
            btnRemoveFilter.Enabled = True
        Catch ex As Exception

        End Try

    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        If DataGridView1.IsCurrentCellDirty Then
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If bByPass Then Exit Sub
        With btnDelete
            .Enabled = False
            For Each rw As DataGridViewRow In CType(sender, DataGridView).Rows
                If Convert.ToBoolean(rw.Cells("Selected").Value) = True Then
                    .Enabled = True
                    Exit For
                End If
            Next

        End With
        With btnSave
            .Enabled = False
            For Each rw As DataGridViewRow In CType(sender, DataGridView).Rows
                If Convert.ToBoolean(rw.Cells("Selected").Value) = True Then
                    .Enabled = True
                    Exit For
                End If
            Next

        End With
    End Sub

    Private Sub DataGridView1_CurrentCellDirtyStateChanged(sender As Object, e As System.EventArgs) Handles DataGridView1.CurrentCellDirtyStateChanged

        If CType(sender, DataGridView).CurrentCell.ColumnIndex = 0 Then
            If DataGridView1.IsCurrentCellDirty Then
                DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
            End If

        End If
    End Sub


    Private Sub DataGridView1_ColumnHeaderMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        If e.ColumnIndex = 0 Then
            bByPass = True
            SelectAll()
            bByPass = False

            With btnDelete
                .Enabled = False
                For Each rw As DataGridViewRow In CType(sender, DataGridView).Rows
                    If Convert.ToBoolean(rw.Cells("Selected").Value) = True Then
                        .Enabled = True
                        Exit Sub
                    End If
                Next

            End With
        End If

        If cOptionalCriteria.SearchType = SearchType.CopiedPrice.ToString Then SetDataGridBackgroundColor()

    End Sub

    Private Sub DataGridView_ColumnWidthChanged(sender As Object, e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles DataGridView1.ColumnWidthChanged
        If bIsLoading Then Exit Sub
        Dim dgv As DataGridView = Me.DataGridView1

        Dim itmcatwidth As Integer
        Dim natwidth As Integer
        Dim colwidth As Integer
        Dim detwidth As Integer

        With DataGridView1
            itmcatwidth = .Columns("Selected").Width + .Columns("ItemNumber").Width + .Columns("ItemDescription").Width + .Columns("ProdCategory").Width + .Columns("ProdCatDescription").Width + .Columns("TerCode").Width '+ .RowHeadersWidth
            natwidth = .Columns("ItemLocPriceColor").Width + .Columns("OriginalPriceColor").Width + .Columns("ActivePriceColor").Width
            colwidth = .Columns("ItemLocPriceRococo").Width + .Columns("OriginalPriceRococo").Width + .Columns("ActivePriceRococo").Width
            detwidth = .Columns("ItemLocPriceDetailStain").Width + .Columns("OriginalPriceDetailStain").Width + .Columns("ActivePriceDetailStain").Width

        End With
        With dgvHeader
            If dgv.Columns("ActivePriceColor").Width <> 2 Then
                .Columns("Column1").Width = itmcatwidth
                .Columns("Column2").Width = natwidth
                .Columns("Column3").Width = colwidth
                .Columns("Column4").Width = detwidth
            Else
                .Columns("Column1").Width = itmcatwidth - 1
                .Columns("Column2").Width = natwidth - 1
                .Columns("Column3").Width = colwidth - 2
                .Columns("Column4").Width = detwidth - 3
            End If
            .AllowUserToAddRows = False
            .Height = .ColumnHeadersHeight
        End With

    End Sub

    Private Sub ItemPricingObjDataGridView_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs)
        MessageBox.Show("Error happened " _
                & e.Context.ToString())

        If (e.Context = DataGridViewDataErrorContexts.Commit) _
            Then
            MessageBox.Show("Commit error")
        End If
        If (e.Context = DataGridViewDataErrorContexts _
            .CurrentCellChange) Then
            MessageBox.Show("Cell change")
        End If
        If (e.Context = DataGridViewDataErrorContexts.Parsing) _
            Then
            MessageBox.Show("parsing error")
        End If
        If (e.Context =
            DataGridViewDataErrorContexts.LeaveControl) Then
            MessageBox.Show("leave control error")
        End If

        If (TypeOf (e.Exception) Is ConstraintException) Then
            Dim view As DataGridView = CType(sender, DataGridView)
            view.Rows(e.RowIndex).ErrorText = "an error"
            view.Rows(e.RowIndex).Cells(e.ColumnIndex) _
                .ErrorText = "an error"

            e.ThrowException = False
        End If
    End Sub

    Private Sub pasteFromClipboard(ByVal dgv As DataGridView)
        Dim rowSplitter As Char() = {Convert.ToChar(vbCr), Convert.ToChar(vbLf)}
        Dim columnSplitter As Char() = {Convert.ToChar(vbTab)}
        Dim rowsInClipboard As String() = Nothing

        'get the text from clipboard
        Dim dataInClipboard As IDataObject = Clipboard.GetDataObject()
        Dim stringInClipboard As String = CStr(dataInClipboard.GetData(DataFormats.Text))

        If stringInClipboard Is Nothing Then Exit Sub

        'split it into lines
        Try
            rowsInClipboard = stringInClipboard.Split(rowSplitter, StringSplitOptions.RemoveEmptyEntries)
        Catch ex As Exception
            ex.Message.ToString()
        End Try

        Dim iRow As Integer = 0
        While iRow < rowsInClipboard.Length
            Dim itm As New ItemPricingObj

            'split row into cell values
            Dim valuesInRow As String() = rowsInClipboard(iRow).Split(columnSplitter)

            'Try
            If valuesInRow(0) = "" Then itm.Selected = True Else Convert.ToBoolean(valuesInRow(0))
            itm.ItemNo = Convert.ToString(valuesInRow(1))
            itm.ItemDesc = Convert.ToString(valuesInRow(2))
            itm.ProdCat = Convert.ToString(IIf(valuesInRow(3) Is Nothing, "", RegExSearch.PadString(Convert.ToString(valuesInRow(3)), "000", PadPosition.PadStart, 3)))
            itm.ProdCatDesc = Convert.ToString(valuesInRow(4))
            itm.TerCode = Convert.ToString(IIf(valuesInRow(5) Is Nothing, "", RegExSearch.PadString(Convert.ToString(valuesInRow(5)), "000", PadPosition.PadStart, 3)))
            If valuesInRow(6) = "" Then itm.ItemLocPriceColor = 0 Else itm.ItemLocPriceColor = Convert.ToDecimal(valuesInRow(6))
            If valuesInRow(7) = "" Then itm.OriginalPriceColor = 0 Else itm.OriginalPriceColor = Convert.ToDecimal(valuesInRow(7))
            If valuesInRow(8) = "" Then itm.ActivePriceColor = 0 Else itm.ActivePriceColor = Convert.ToDecimal(valuesInRow(8))
            If valuesInRow(9) = "" Then itm.ItemLocPriceRococo = 0 Else itm.ItemLocPriceRococo = Convert.ToDecimal(valuesInRow(9))
            If valuesInRow(10) = "" Then itm.OriginalPriceRococo = 0 Else itm.OriginalPriceRococo = Convert.ToDecimal(valuesInRow(10))
            If valuesInRow(11) = "" Then itm.ActivePriceRococo = 0 Else itm.ActivePriceRococo = Convert.ToDecimal(valuesInRow(11))
            If valuesInRow(12) = "" Then itm.ItemLocPriceDetailStain = 0 Else itm.ItemLocPriceDetailStain = Convert.ToDecimal(valuesInRow(12))
            If valuesInRow(13) = "" Then itm.OriginalPriceDetailStain = 0 Else itm.OriginalPriceDetailStain = Convert.ToDecimal(valuesInRow(13))
            If valuesInRow(14) = "" Then itm.ActivePriceDetailStain = 0 Else itm.ActivePriceDetailStain = Convert.ToDecimal(valuesInRow(14))

            itm.TerFrom = Convert.ToString(IIf(valuesInRow(15) Is Nothing, "", RegExSearch.PadString(Convert.ToString(valuesInRow(15)), "000", PadPosition.PadStart, 3)))
            itm.TerDescription = Convert.ToString(valuesInRow(16))
            itm.OnPriceList = Convert.ToString(valuesInRow(17))
            If valuesInRow(18) = "" Then itm.LastDate = Convert.ToDateTime("01/01/1900") Else itm.LastDate = Convert.ToDateTime(valuesInRow(10))



            iRow += 1
            cItemPricingList.Add(itm)
        End While

        ItemPricingObjBindingSource.DataSource = cItemPricingList
        dgv.DataSource = ItemPricingObjBindingSource

    End Sub


    Private Sub btnLoadItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoadItems.Click
        bSetCriteria = True
        bFillPressed = False : bLoadPressed = True
        SetOptionalCriteria()

        If mcboTerritoryCodes.Text = "" Then

            If MsgBox("A Territory Code has not been selected.  Do you want to load ALL Territory Codes?", MsgBoxStyle.YesNo, "Load Territory Codes") = MsgBoxResult.Yes Then
                With Timer1
                    .Interval = 100
                    .Enabled = True
                End With
            Else
                Exit Sub
            End If
        Else

            With Timer1
                .Interval = 100
                .Enabled = True
            End With

        End If

    End Sub



    Private Sub LoadData(ByVal SearchCriteria As OptionalCriteria)
        Cursor = Cursors.WaitCursor

        ItemPricingObjBindingSource.RemoveFilter()
        cItemPricingList.Clear()

        FillAllItemList(bSetCriteria, SearchCriteria)

        Cursor = Cursors.Default

        bByPass = True
        bCheckState = True

        bByPass = False

        If cOptionalCriteria.SearchType = SearchType.CopiedPrice.ToString Then SetDataGridBackgroundColor()
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Dim tmr As Timer = CType(sender, Timer)

        tmr.Enabled = False
        clear(False)
        LoadData(cOptionalCriteria)
        tabOptions.SelectedIndex = 0
        mcboTerritoryCodes.Focus()

    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Timer2.Enabled = False
        mcboTerritoryCodes.Focus()
    End Sub

    Private Sub Timer3_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        Timer3.Enabled = False
        With rbAdvanced
            .Visible = True
            .Checked = True
        End With
        cboZonePrc.SelectedIndex = 0
        cboZonePrc.SelectedItem = "Ter Code"
        cboZonePrc.Text = "Ter Code"
        cboZonePrc.Enabled = False
        cboZonePrc.Refresh()
        cOptionalCriteria.PricingType = PrcType.Advanced.ToString

    End Sub

    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick
        Dim tmr As Timer = CType(sender, Timer)

        tmr.Enabled = False
        If cOptionalCriteria.PricingType = PrcType.Advanced.ToString Then Exit Sub
        SetOptionalCriteriaSearch()

        LoadData(cOptionalCriteriaSearch)
        FormatPriceType(FormatPriceTypeState.Open.ToString)
        tabOptions.SelectedIndex = 0
        bAdvanceSearch = False
    End Sub

    Private Sub rbPricingOptions_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbUpdate.CheckedChanged, rbCopy.CheckedChanged,
                                                rbPrimary.CheckedChanged, rbDirectCopy.CheckedChanged, rbAdvanced.CheckedChanged
        If bByPass Then Exit Sub

        Dim rb As RadioButton = CType(sender, RadioButton)

        Try
            Select Case rb.Name
                Case rbUpdate.Name
                    FormatPriceType(FormatPriceTypeState.Update.ToString)
                    With cOptionalCriteria
                        '.PricingType = PrcType.Update.ToString
                        .UpdateType = PrcUpdateType.Update.ToString
                    End With

                Case rbCopy.Name
                    FormatPriceType(FormatPriceTypeState.Copy.ToString)
                    With cOptionalCriteria
                        .PricingType = PrcType.Copied.ToString
                        .UpdateType = PrcUpdateType.Copied.ToString
                    End With
                Case rbPrimary.Name
                    FormatPriceType(FormatPriceTypeState.Primary.ToString)
                    With cOptionalCriteria
                        '.PricingType = PrcType.Primary.ToString
                        .UpdateType = PrcUpdateType.Primary.ToString
                    End With
                Case rbDirectCopy.Name
                    FormatPriceType(FormatPriceTypeState.DirectCopy.ToString)
                    With cOptionalCriteria
                        '.PricingType = PrcType.DirectCopied.ToString
                        .UpdateType = PrcUpdateType.DirectCopy.ToString
                    End With
                Case rbAdvanced.Name
                    FormatPriceType(FormatPriceTypeState.Advanced.ToString)
                    With cOptionalCriteria
                        .PricingType = PrcType.Advanced.ToString
                        .UpdateType = PrcUpdateType.Advanced.ToString
                    End With
            End Select

        Catch ex As Exception
        End Try
    End Sub



#Region "   Format Price Types   "

    Private Sub FormatPriceType(FormatType As String)

        'Format based on the Price Type we are using 
        Select Case FormatType

            Case PrcType.Copied.ToString
                'Group Boxes ...
                grpFill.Enabled = True
                grpFilter.Enabled = True
                'Disable Update Options

                'With txtTerCode
                '    .Enabled = False
                '    .ReadOnly = True
                'End With
                'With txtTerFrom
                '    .Enabled = False
                '    .ReadOnly = True
                'End With

                'Enable Buttons
                btnSelectAll.Enabled = True
                btnFillPricing.Enabled = True
                btnRoundPricing.Enabled = True
                btnDelete.Enabled = True
                btnSave.Enabled = cOptionalCriteria.IsFillPressed

                rbZoneFill.Enabled = True '***no longer used
                rbCustomFill.Enabled = True
                pnlFillOptions.Enabled = True
                rbPercentFill.Enabled = False
                rbAmountFill.Enabled = False
                lblNatural.Enabled = False
                lblColor.Enabled = False
                lblDetail.Enabled = False
                rbPercentFill.Enabled = False
                rbAmountFill.Enabled = False
                txtNaturalMarkup.Enabled = False
                txtColorMarkup.Enabled = False
                txtDetailMarkup.Enabled = False

                txtFillTerrDesc.Enabled = True
                mcboFillTerrCodes.Enabled = True

            Case PrcType.CustomPercent.ToString
                'Group Boxes ...
                grpFill.Enabled = True
                grpFilter.Enabled = True
                'Disable Update Options
                cboZonePrc.Enabled = False
                With txtTerCode
                    .Enabled = False
                    .ReadOnly = True
                End With

                'Enable Buttons
                btnSelectAll.Enabled = True
                btnFillPricing.Enabled =
                btnRoundPricing.Enabled = True
                btnDelete.Enabled = True
                btnSave.Enabled = cOptionalCriteria.IsFillPressed

                rbZoneFill.Enabled = True '***no longer used
                pnlFillOptions.Enabled = True
                rbCustomFill.Enabled = True
                rbPercentFill.Enabled = True
                rbAmountFill.Enabled = True
                lblNatural.Enabled = True
                lblColor.Enabled = True
                lblDetail.Enabled = True
                rbPercentFill.Enabled = True
                rbAmountFill.Enabled = True
                txtNaturalMarkup.Enabled = True
                txtColorMarkup.Enabled = True
                txtDetailMarkup.Enabled = True

                txtFillTerrDesc.Enabled = True
                mcboFillTerrCodes.Enabled = True
                tabOptions.SelectedIndex = 0

            Case PrcType.Advanced.ToString

                'Enable Section
                grpFill.Enabled = True : grpPriceType.Enabled = True
                rbZoneFill.Enabled = True : rbCustomFill.Enabled = True
                btnDelete.Enabled = True

                'Disable Section
                lblCurrent.Enabled = False : lblNew.Enabled = False
                txtFillTerrDesc.Enabled = False : txtTerCode.Enabled = False
                mcboFillTerrCodes.Enabled = False
                rbUpdate.Enabled = False : rbCopy.Enabled = False : rbDirectCopy.Enabled = False : rbPrimary.Enabled = False
                btnSelectAll.Enabled = True : btnFillPricing.Enabled = True : btnRoundPricing.Enabled = True
                grpFilter.Enabled = False

            Case MarkupType.Zone.ToString
                rbCustomFill.Enabled = True
                pnlFillOptions.Enabled = True
                rbPercentFill.Enabled = False
                rbAmountFill.Enabled = False
                lblNatural.Enabled = False
                lblColor.Enabled = False
                lblDetail.Enabled = False
                rbPercentFill.Enabled = False
                rbAmountFill.Enabled = False
                txtNaturalMarkup.Enabled = False
                txtColorMarkup.Enabled = False
                txtDetailMarkup.Enabled = False
                cboZonePrc.Enabled = True
                btnSelectAll.Enabled = True
                btnFillPricing.Enabled = True
                btnRoundPricing.Enabled = True
                btnDelete.Enabled = True
                btnSave.Enabled = cOptionalCriteria.IsFillPressed
            Case FormatPriceTypeState.Closed.ToString 'Used in Clear method
                For Each ctl As Control In GroupBox2.Controls
                    ctl.Enabled = True
                Next
                For Each ctl As Control In grpPriceType.Controls
                    If ctl.Name = rbUpdate.Name _
                        Or ctl.Name = rbCopy.Name _
                        Or ctl.Name = rbPrimary.Name _
                        Or ctl.Name = rbDirectCopy.Name Then
                        ctl.Enabled = True
                        Dim rb As RadioButton = CType(ctl, RadioButton)
                        rb.Checked = False
                    Else
                        ctl.Enabled = False
                    End If
                Next

                grpPriceType.Enabled = False
                grpFill.Enabled = False
                btnSelectAll.Enabled = False
                btnFillPricing.Enabled = False
                btnRoundPricing.Enabled = False
                btnDelete.Enabled = False
                btnSave.Enabled = cOptionalCriteria.IsFillPressed
        End Select


    End Sub

    Private Sub FormatFillType(FillType As String)
        grpFill.Enabled = True

        Select Case FillType
            Case FormatFillTypeState.Open.ToString
                btnSelectAll.Enabled = False
                btnFillPricing.Enabled = False
                btnRoundPricing.Enabled = False

            Case FormatFillTypeState.Copy.ToString

            Case FormatFillTypeState.Zone.ToString

            Case FormatFillTypeState.Custom.ToString

            Case FormatFillTypeState.Close.ToString

        End Select

    End Sub

    'Private Sub FormatManualFilterSave()
    '    btnSave.Enabled = True
    'End Sub

#End Region
    Private Sub btnRoundPricing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRoundPricing.Click
        If cItemPricingList.Count = 0 Then Exit Sub
        Me.RoundPrices()
    End Sub

    Private Sub setSaveButtonState()
        Dim dgv As DataGridView = Me.DataGridView1
        Dim rw As DataGridViewRow

        For Each rw In dgv.Rows

            If CBool(rw.Cells("Selected").Value) = True Then
                cOptionalCriteria.EnableFillButtons = True
                cOptionalCriteria.IsFillPressed = True
                Exit For
            Else
                cOptionalCriteria.EnableFillButtons = False
                btnSave.Enabled = cOptionalCriteria.IsFillPressed
            End If

        Next

    End Sub

    Public Function getDataFromExcel(ByVal FileName As String,
                                  ByVal SheetName As String, ByVal RangeName As String,
                                  ByVal ImportType As String, ByVal WhseID As String) As DataSet

        Try
            Dim strConn As String =
                "Provider=Microsoft.Jet.OLEDB.4.0;" &
                "Data Source=" & FileName & "; Extended Properties=Excel 8.0;"
            Dim oCn _
                As New System.Data.OleDb.OleDbConnection(strConn)
            oCn.Open()
            ' Create oects ready to grab data
            Dim oCmd As New System.Data.OleDb.OleDbCommand(
                "SELECT * FROM [" & SheetName & RangeName & "]", oCn)
            Dim oDA As New System.Data.OleDb.OleDbDataAdapter()
            oDA.SelectCommand = oCmd

            ' Fill DataSet
            Dim oDS As New DataSet()
            oDA.Fill(oDS)
            Return oDS

        Catch

            MsgBox(Err.Description)
            'MsgBox("STOP")
            Return Nothing
            Exit Function
        End Try

    End Function
    Private Function CheckZeroPrices() As Boolean
        Dim bZeroPriceFound As Boolean = True
        Dim sMsg As String
        'Dim iNonZeroPrices As Integer = 0
        'Dim iZeroPrices As Integer = 0
        For Each itm As ItemPricingObj In cItemPricingList
            If itm.Selected Then
                If ((itm.ActivePriceColor = 0 And itm.OriginalPriceColor > 0) _
                    OrElse (itm.ActivePriceRococo = 0 And itm.OriginalPriceRococo > 0) _
                    OrElse (itm.ActivePriceDetailStain = 0 And itm.OriginalPriceDetailStain > 0)) Then
                    sMsg = "SET PRICE TO ZERO ITEMS FOUND!" & vbCrLf & vbCrLf &
                           "Setting prices to ZERO is not permitted.  " & vbCrLf & vbCrLf &
                                        "Uncheck this item - OR - Change the ZERO Price in the 'New Prc' Column: " & vbCrLf & vbCrLf &
                                        "Item: " & itm.ItemNo & ", " & itm.ItemDesc & vbCrLf & vbCrLf &
                                        "NATRURAL:" & vbTab & "Current Price - " & itm.OriginalPriceColor & " / New Price - " & itm.ActivePriceColor & vbCrLf &
                                        "COLOR:" & vbTab & vbTab & "Current Price - " & itm.OriginalPriceRococo & " / New Price - " & itm.ActivePriceRococo & vbCrLf &
                                        "DETAIL:" & vbTab & vbTab & "Current Price - " & itm.OriginalPriceDetailStain & " / New Price - " & itm.ActivePriceDetailStain & vbCrLf & vbCrLf &
                                        "Message copied to Clipboard.  Paste to Notepad for reference. "

                    MsgBox(sMsg, MsgBoxStyle.OkOnly, "SET PRICE TO ZERO PRICE FOUND!")

                    Clipboard.SetText(sMsg)


                    sMsg = "Save Aborted"
                    MsgBox(sMsg, MsgBoxStyle.OkOnly, "SET PRICE TO ZERO PRICE FOUND!")
                    bZeroPriceFound = True
                    Return bZeroPriceFound
                End If
            End If

        Next
        bZeroPriceFound = False

        Return bZeroPriceFound
    End Function
    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        'check for zeros in all New Prc Columns to see if Fill was pressed.  Stop if it was not.  

        'Dim iZeroPricesCount As Integer = 
        If CheckZeroPrices() Then
            Cursor = Cursors.Default
            Exit Sub
        End If

        bSetCriteria = False
        Dim dt As DataTable
        Dim saveMessage As New StringBuilder


        'check if this is a manual update before proceeding, bLoadPressed = True
        If bLoadPressed = True Then
            With cOptionalCriteria

                .TerCopyToCode = txtTerCode.Text.Trim
                .TerCopyToDesc = IIf(bCopyToOption = True, txtFillTerrDesc.Text.Trim, txtTerrDesc.Text.Trim).ToString
                .TerFromCode = IIf(bCopyToOption = True, txtTerFrom.Text.Trim, txtTerrDesc.Text.Trim).ToString 'txtTerFrom.Text.Trim 'IIf(.TerCopyToCode.Substring(0, 2) = "00", "", txtTerCode.Text.Trim).ToString

            End With
        End If


        Cursor = Cursors.WaitCursor
        'Direct Pricing
        If cOptionalCriteria.PricingType = PrcType.Copied.ToString And cOptionalCriteria.MarkupType = MarkupType.Direct.ToString Then
            ' handle the Ter From first - if it's a primary price we are saving, empty cOptionalCriteria.TerFromCode.  TerFrom is only used in non-primary pricing.
            With cOptionalCriteria
                .TerCopyToCode = mcboTerritoryCodes.Text.Trim
                .TerCopyToDesc = txtTerrDesc.Text.Trim
                .TerFromCode = IIf(.TerCopyToCode.Substring(0, 2) = "00", "", txtTerCode.Text.Trim).ToString
            End With

            'if its a copied price list, handle here ... 
            With saveMessage
                .Append(vbCrLf)
                If cOptionalCriteria.TerCopyToCode > "" Then .Append("Save Territory To: ") : .Append(cOptionalCriteria.TerCopyToCode & vbCrLf)
                If cOptionalCriteria.TerCopyToDesc > "" Then .Append("Save Territory To Description: ") : .Append(cOptionalCriteria.TerCopyToDesc & vbCrLf)
                If cOptionalCriteria.TerFromCode > "" Then .Append("Territory From: ") : .Append(cOptionalCriteria.TerFromCode & vbCrLf)
                .Append(vbCrLf)
            End With


            'Double Check before we make changes to the pricing 
            If MsgBox("Ready to save price list: " & vbCrLf & saveMessage.ToString & "Proceed?",
                                            MsgBoxStyle.YesNo, "Save Pricing") = MsgBoxResult.No Then
                Cursor = Cursors.Default
                Exit Sub

            End If

            SaveItemPricing(cItemPricingList)

            Cursor = Cursors.Default

            With Timer4
                .Interval = 100
                .Enabled = True
            End With

            Cursor = Cursors.Default
            cOptionalCriteria.IsFillPressed = False

            Exit Sub

        ElseIf cOptionalCriteria.PricingType = PrcType.Copied.ToString Then
            ' handle the Ter From first - if it's a primary price we are saving, empty cOptionalCriteria.TerFromCode.  TerFrom is only used in non-primary pricing.
            With cOptionalCriteria
                If .TerFromCode > "" Then
                    If .TerCopiedFromCode > "" Then
                        .TerFromCode = IIf(.TerCopyToCode.Substring(0, 2) = "00", "", .TerCopiedFromCode.Trim).ToString
                    End If

                End If

                If bCopyToOption = True Then
                    .TerFromCode = txtTerFrom.Text.Trim
                Else
                    .TerFromCode = txtFrom.Text
                End If


            End With

            'if its a copied price list, handle here ... 
            With saveMessage
                .Append(vbCrLf)
                If cOptionalCriteria.TerCopyToCode > "" Then .Append("Save Territory To: ") : .Append(cOptionalCriteria.TerCopyToCode & vbCrLf)
                If cOptionalCriteria.TerCopyToDesc > "" Then .Append("Save Territory To Description: ") : .Append(cOptionalCriteria.TerCopyToDesc & vbCrLf)
                If cOptionalCriteria.TerFromCode > "" Then .Append("Territory From: ") : .Append(cOptionalCriteria.TerFromCode & vbCrLf)
                .Append(vbCrLf)
            End With


            'Double Check before we make changes to the pricing 
            If MsgBox("Ready to save price list: " & vbCrLf & saveMessage.ToString & "Proceed?",
                                            MsgBoxStyle.YesNo, "Save Pricing") = MsgBoxResult.No Then
                Cursor = Cursors.Default
                Exit Sub
            Else
                bCopyToOption = False
            End If
            'cOptionalCriteria.SearchType = SearchType.PrimaryPrice.ToString

            SaveItemPricing(cItemPricingList)

            Cursor = Cursors.Default

            With Timer4
                .Interval = 100
                .Enabled = True
            End With

            Cursor = Cursors.Default
            cOptionalCriteria.IsFillPressed = False

            Exit Sub

        ElseIf cOptionalCriteria.PricingType = PrcType.Advanced.ToString Then
            'Advanced Pricing is from the Lookup - One item with the price adjusted across multiple territories 
            cOptionalCriteria.PricingType = PrcType.Advanced.ToString

            With saveMessage
                .Append(vbCrLf)
                .Append("Territory From: ") : .Append(" - Keep existing markup for each territory.  Manual price changes will also be saved." & vbCrLf)
                .Append(vbCrLf)
            End With


            'Double Check before we make changes to the pricing 
            If MsgBox("Ready to save price list: " & vbCrLf & saveMessage.ToString & "Proceed?",
                                            MsgBoxStyle.YesNo, "Save Pricing") = MsgBoxResult.No Then
                Cursor = Cursors.Default
                Exit Sub
            End If

            SaveAdvancedItemPricing(cItemPricingList)
            dt = RetrieveAdvancedItemPricing(cItemPricingList)


            Dim rd As DataTableReader = dt.CreateDataReader
            'empty all objects and refresh the grid
            clear()

            itmprclst = New ItemPricingList
            itmprclst = BusObj.PopulateSearchItems(rd)
            cItemPricingList = itmprclst

            With ItemPricingObjBindingSource
                .DataSource = Nothing
                .DataSource = itmprclst
            End With

            DataGridView1.DataSource = Nothing
            DataGridView1.Refresh()
            DataGridView1.DataSource = ItemPricingObjBindingSource
            DataGridView1.Refresh()

            Cursor = Cursors.Default
            cOptionalCriteria.IsFillPressed = False

            With Timer4
                .Interval = 100
                .Enabled = True
            End With
            Exit Sub

        Else

            'All other Pricing types go here 
            CollectUIFillData()

            With saveMessage
                .Append(vbCrLf)
                If cOptionalCriteria.TerCopyToCode > "" Then .Append("Save Territory To: ") : .Append(cOptionalCriteria.TerCopyToCode & vbCrLf)
                If cOptionalCriteria.TerCopyToDesc > "" Then .Append("Save Territory To Description: ") : .Append(cOptionalCriteria.TerCopyToDesc & vbCrLf)
                If cOptionalCriteria.TerFromCode > "" Then .Append("Territory From: ") : .Append(cOptionalCriteria.TerFromCode & vbCrLf)
                If cOptionalCriteria.PricingType > "" Then .Append("Pricing Type: ") : .Append(cOptionalCriteria.PricingType & vbCrLf)
                .Append(vbCrLf)
            End With


            'Double Check before we make changes to the pricing 
            If MsgBox("Ready to save price list: " & vbCrLf & saveMessage.ToString & "Proceed?",
                                            MsgBoxStyle.YesNo, "Save Pricing") = MsgBoxResult.No Then
                Cursor = Cursors.Default
                Exit Sub
            End If

            SaveItemPricing(cItemPricingList)

            With Timer4
                .Interval = 100
                .Enabled = True
            End With

            Cursor = Cursors.Default
            cOptionalCriteria.IsFillPressed = False
        End If
    End Sub



#Region "   Quick Filter    "
    Private Sub btnApplyFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApplyFilter.Click
        ApplyFilter()
    End Sub
    Private Sub ApplyFilter()
        'dtFilteredItems holds the items filtered out so when we save and refresh, we can reuse 
        'the items in the LoadData(cOptionalCriteria) function and set the user's filtered values
        dtFilteredItems = New DataTable
        dtFilteredItems.Columns.Add("item_no", GetType(String))

        Dim s As String = ""
        For Each itm As Object In ListBox1.Items
            If s = "" Then
                s = "'" & itm.ToString & "'"
            Else
                s = "'" & s & "'|'" & itm.ToString & "'"
            End If
            dtFilteredItems.Rows.Add(itm.ToString.Trim)
        Next
        ItemPricingObjBindingSource.Filter = "ItemNo=" & s

        With cboFilter
            .Enabled = False
            .DataSource = Nothing
            .DisplayMember = "ItemNo"
        End With

    End Sub

    Private Sub btnRemoveFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveFilter.Click
        RemoveFilter()
    End Sub

    Private Sub RemoveFilter()
        cItemPricingList.RemoveFilter()

        With cSearchClass
            .SearchPropertyValue = ""
        End With
        Reset_MassarelliOutboundBindingSource()

        With cboFilter
            .DataSource = ItemPricingObjBindingSource
            .Enabled = True
        End With
        'clear dtFilteredItems so it will be skipped when we save and want to reset the UI to what it was
        dtFilteredItems = New DataTable
    End Sub
    Private Sub DataGridView_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs)

        ' User can cold the ShiftKey and select a group of rows ...
        If e.Button = Windows.Forms.MouseButtons.Left And Control.ModifierKeys = Keys.Shift Then

            If e.ColumnIndex = 0 Then
                Me.Validate()
                For Each rw As DataGridViewRow In CType(sender, DataGridView).Rows
                    If rw.Cells("Selected").Selected Then
                        rw.Cells(0).Value = True
                    End If
                Next
            End If
        End If

    End Sub

    Private Sub Reset_MassarelliOutboundBindingSource()
        SearchBindingSource.DataSource = cSearchClass
    End Sub


    Private Sub AddItemToFilterList()
        If cboFilter.SelectedItem Is Nothing Then
            MsgBox("Item " & cboFilter.Text & " was not found on the grid.  It may still be a valid item, but to set pricing it must be retrieved to the grid.  Check the Item or change the Search Criteria and try again", MsgBoxStyle.OkOnly, "Item Not On Grid")
            Exit Sub
        End If
        If Not ListBox1.Items.Contains(DirectCast(cboFilter.SelectedItem, ItemPricingObj).ItemNo) Then ListBox1.Items.Add(DirectCast(cboFilter.SelectedItem, ItemPricingObj).ItemNo)

        DirectCast(cboFilter.SelectedItem, ItemPricingObj).Selected = True
        With DataGridView1
            .EndEdit()
            .RefreshEdit()
            .Refresh()
        End With
    End Sub


    Private Sub AddGridItemToFilterList(ByVal itemno As String)

        'If dgv.SelectedItem Is Nothing Then
        '    MsgBox("Item " & cboFilter.Text & " was not found on the grid.  It may still be a valid item, but to set pricing it must be retrieved to the grid.  Check the Item or change the Search Criteria and try again", MsgBoxStyle.OkOnly, "Item Not On Grid")
        '    Exit Sub
        'End If
        If Not ListBox1.Items.Contains(itemno) Then ListBox1.Items.Add(itemno)

        'DirectCast(cboFilter.SelectedItem, ItemPricingObj).Selected = True
        With DataGridView1
            .EndEdit()
            .RefreshEdit()
            .Refresh()
        End With
    End Sub


#End Region


    Private Sub btnAdd_Click(sender As System.Object, e As System.EventArgs) Handles btnAdd.Click
        AddItemToFilterList()
        Dim lst As ListBox = CType(ListBox1, ListBox)

        If lst.Items.Count = 0 Then
            btnApplyFilter.Enabled = False
            btnRemoveFilter.Enabled = False
        Else
            btnApplyFilter.Enabled = True
            btnRemoveFilter.Enabled = True
        End If

    End Sub

#Region "   Context Menu   "

    'Private Sub btnCheckAll_Click(sender As System.Object, e As System.EventArgs)
    '    SelectAll()
    '    setSaveButtonState()
    'End Sub

    'Private Sub btnCheckHighlighted_Click(sender As System.Object, e As System.EventArgs)
    '    SelectHighlightedFromGrid()
    'End Sub

    'Private Sub ClearAllToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)
    '    cItemPricingList.Clear()
    '    Me.ItemPricingObjBindingSource.Clear()
    'End Sub

    ''Private Sub btnClearChecked_Click(sender As System.Object, e As System.EventArgs)
    '    Me.RemoveSelectedFromGrid()
    'End Sub

    Private Sub SelectAll()
        Dim rw As DataGridViewRow
        Me.SuspendLayout()
        Cursor = Cursors.WaitCursor
        With DataGridView1
            For Each rw In .Rows
                If rw.ReadOnly = False Then
                    rw.Cells("Selected").Value = bCheckState
                    'rw.Selected = False
                End If

            Next
            If bCheckState = True Then
                bCheckState = False
            Else
                bCheckState = True
            End If


        End With
        Cursor = Cursors.Default
        Me.ResumeLayout()
    End Sub
    Private Sub FormatTerritoryPricing(FormatType As String) Handles fAdvancedLookup.FormatTerPricing

        If FormatType = MarkupType.Zone.ToString Then
            Dim dgv As DataGridView = CType(DataGridView1, DataGridView)
            For Each rw As DataGridViewRow In dgv.Rows
                rw.Cells("Selected").Value = 1
            Next
        End If
        rbZoneFill.Checked = Convert.ToBoolean(IIf(FormatType = MarkupType.Zone.ToString, True, False))
        FormatPriceType(FormatType)
    End Sub


    Private Sub GetDataFromAdvancedSearch(dtAdvancedData As DataTable) Handles fAdvancedLookup.SendDataToGrid

        bAdvanceSearch = True
        Dim rd As DataTableReader = dtAdvancedData.CreateDataReader
        Dim itmprclsttmp As New ItemPricingList

        itmprclsttmp = BusObj.PopulateSearchItems(rd)

        For Each itm As ItemPricingObj In itmprclsttmp
            itmprclst.Add(itm)
        Next

        With ItemPricingObjBindingSource
            .DataSource = Nothing
            .DataSource = itmprclst
        End With
        DataGridView1.DataSource = ItemPricingObjBindingSource
        cItemPricingList = itmprclst
        FillTerritoryCodeList()
        itmprclsttmp.Clear()

        With Timer3
            .Interval = 500
            .Enabled = True
        End With

    End Sub

    Private Sub SelectHighlightedFromGrid()
        Dim i As Integer
        Dim chkstate As Boolean
        Dim dgv As DataGridView = Me.DataGridView1

        Dim selectedRowCount As Integer =
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
            Me.DataGridView1.Refresh()

        End If
    End Sub

    Private Sub RemoveSelectedFromGrid()
        Dim i As Integer
        Me.Validate()
        For i = cItemPricingList.Count - 1 To 0 Step -1
            If cItemPricingList.Item(i).Selected = True Then
                cItemPricingList.Remove(cItemPricingList.Item(i))
            End If
        Next

        i = 0

    End Sub

#End Region

    Private Sub cbDBList_DropDownClosed(sender As Object, e As System.EventArgs) Handles cbDBList.DropDownClosed
        RefreshDATA()
    End Sub

    Private Sub cbDBList_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbDBList.SelectedIndexChanged
        If bIsLoading Then Exit Sub
        Dim cbo As ComboBox = CType(sender, ComboBox)
        Dim db As String = cbo.SelectedItem.ToString

        If Not db = "DATA" Then
            Me.DataGridView1.BackgroundColor = Color.DimGray
            Me.DataGridView1.DefaultCellStyle.BackColor = SystemColors.Info
            lblTest.Visible = True : lblCompany.Visible = True
        Else
            Me.DataGridView1.BackgroundColor = SystemColors.AppWorkspace
            Me.DataGridView1.DefaultCellStyle.BackColor = SystemColors.Window
            lblTest.Visible = False : lblCompany.Visible = False
        End If
        bEnableRefreshDB = True 'CBool(IIf(cbo.SelectedItem.ToString = My.Settings.DefaultDB, False, True))
        bEnableSaveDB = True ' CBool(IIf(cbo.SelectedItem.ToString = My.Settings.DefaultDB, False, True))
        bEnableCtls = True ' CBool(IIf(cbo.SelectedItem.ToString = My.Settings.DefaultDB, True, False))

        SetControlState(bEnableRefreshDB, bEnableSaveDB, bEnableCtls)
    End Sub

    Private Sub btnSaveDefaultDB_Click(sender As System.Object, e As System.EventArgs) Handles btnSaveDefaultDB.Click
        Dim cbo As ComboBox = CType(cbDBList, ComboBox)
        Try
            With cOptionalCriteria
                .DBName = cbo.Text
                MacStartup(.DBName)
                lblCurrentDB.Text = .CurrentDB
                lblDefaultDB.Text = .DefaultDB

                'save the default db
                My.Settings.DefaultDB = .DBName
                My.Settings.Save()
            End With

            LoadControls()

            bEnableRefreshDB = False
            bEnableSaveDB = False
            bEnableCtls = True
            SetControlState(bEnableRefreshDB, bEnableSaveDB, bEnableCtls)

        Catch ex As Exception
            MsgBox(My.Settings.DefaultDB & " does not have required data for Territory Pricing.  Setting default to DATA")
            My.Settings.DefaultDB = "DATA"
            My.Settings.Save()
        End Try


    End Sub

    Private Sub btnRefreshDB_Click(sender As System.Object, e As System.EventArgs) Handles btnRefreshDB.Click
        RefreshDATA()
    End Sub
    Private Sub RefreshDATA()
        Dim cbo As ComboBox = CType(cbDBList, ComboBox)
        cOptionalCriteria.DBName = cbo.SelectedItem.ToString
        MacStartup(cOptionalCriteria.DBName)
        lblCurrentDB.Text = cOptionalCriteria.CurrentDB

        LoadControls()
        bEnableRefreshDB = False
        bEnableSaveDB = True
        bEnableCtls = True
        SetControlState(bEnableRefreshDB, bEnableSaveDB, bEnableCtls)
    End Sub
    Private Sub LoadControls()
        Try
            FillTerritoryCodeList()
        Catch ex As Exception
            ' dtTerrCodes.Clear()
        End Try
        Try
            fillCategoryCodes()
        Catch ex As Exception
            'MsgBox("FillCategoryCodes " & ex.Message)
        End Try

    End Sub

    Private Sub SetControlState(enableRefresh As Boolean, enableSave As Boolean, enablectls As Boolean)
        Dim ctl As Control
        btnSaveDefaultDB.Enabled = enableSave
        btnRefreshDB.Enabled = enableRefresh
        Try

            For Each ctl In GroupBox2.Controls
                ctl.Enabled = enablectls
            Next


            For Each ctl In grpFilter.Controls
                ctl.Enabled = enablectls
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub btnRemove_Click(sender As System.Object, e As System.EventArgs) Handles btnRemove.Click
        Dim itm As Integer
        Dim sitm As String

        With ListBox1
            If .SelectedIndex = -1 Then Exit Sub
            itm = .SelectedIndex
            sitm = .SelectedItem.ToString

            .Items.Remove(ListBox1.SelectedItem)
            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.Cells("ItemNumber").Value.ToString = sitm Then
                    row.Cells("Selected").Value = False
                    Exit For
                End If
            Next
            If .Items.Count >= itm Then
                .SelectedIndex = itm - 1
            Else
                .SelectedIndex = .Items.Count - 1
            End If
        End With

        Dim lst As ListBox = CType(ListBox1, ListBox)

        If lst.Items.Count = 0 Then
            btnApplyFilter.Enabled = False
            btnRemoveFilter.Enabled = False
        Else
            btnApplyFilter.Enabled = True
            btnRemoveFilter.Enabled = True
        End If
    End Sub

    Private Sub btnRemoveAll_Click(sender As System.Object, e As System.EventArgs) Handles btnRemoveAll.Click
        ListBox1.Items.Clear()
        For Each row As DataGridViewRow In DataGridView1.Rows
            row.Cells("Selected").Value = False
        Next

        Dim lst As ListBox = CType(ListBox1, ListBox)

        If lst.Items.Count = 0 Then
            btnApplyFilter.Enabled = False
            btnRemoveFilter.Enabled = False
        Else
            btnApplyFilter.Enabled = True
            btnRemoveFilter.Enabled = True
        End If
    End Sub

    Private Sub btnClearAll_Click(sender As System.Object, e As System.EventArgs) Handles btnClearAll.Click
        clear()
        cItemPricingList.RemoveFilter()
    End Sub

    Private Sub clear(Optional clearAll As Boolean = True)
        bByPass = True

        'Grid and Binding Source and ItemPriceList
        If clearAll Then
            mcboTerritoryCodes.Text = ""
            mcboFillTerrCodes.Text = ""
            txtFillTerrDesc.Text = ""
            cboCategoryCodes.SelectedIndex = 0
            cboOnPriceList.SelectedIndex = 0
            cboHasTerrPrice.SelectedIndex = 0
            chkShowActiveOnly.Checked = True
            txtTerrDesc.Text = ""
            txtTerCode.Text = ""
            txtTerFrom.Text = ""
            FormatPriceType(FormatPriceTypeState.Closed.ToString)

            cOptionalCriteria.Clear()
            grpFill.Enabled = False
            grpFilter.Enabled = False
            mcboTerritoryCodes.Enabled = True
            txtTerrDesc.Enabled = True
            cboCategoryCodes.Enabled = True
            cboOnPriceList.Enabled = True
            cboHasTerrPrice.Enabled = True
            chkShowActiveOnly.Enabled = True
            rbAdvanced.Checked = False
            cboZonePrc.Text = ""
            rbZoneFill.Checked = False
            rbDirectMacolaPrc.Checked = False
            bCopyToOption = False
        End If

        DataGridView1.DataSource = Nothing
        With ItemPricingObjBindingSource
            .RemoveFilter()
            .DataSource = Nothing
        End With

        cItemPricingList.Clear()
        'Reset the multicolumncombo, otherwise it creates array size problems next time you use it ...
        FillTerritoryCodeList()

        'Fill Group ...
        txtNaturalMarkup.Text = ""
        txtColorMarkup.Text = ""
        txtDetailMarkup.Text = ""
        cboZonePrc.Text = ""
        cboZonePrc.Enabled = False
        rbZoneFill.Checked = False
        rbCustomFill.Checked = False
        'rbPercentFill.Checked = False

        'Filter Group
        cboFilter.DataSource = Nothing
        cboFilter.Enabled = True
        ListBox1.Items.Clear()

        'Status Strip 
        lblCount.Text = ""

        'buttons & chkboxes ...
        btnSave.Enabled = False
        btnSelectAll.Enabled = False
        btnFillPricing.Enabled = False
        btnRoundPricing.Enabled = False
        rbCustomFill.Checked = False
        mcboTerritoryCodes.Focus()
        bByPass = False
        tabOptions.SelectedIndex = 0
        bCopyToOption = False
    End Sub

    Private Sub AdvancedSearch()

        rbPrimaryTerritory.Checked = False

        mcboTerritoryCodes.Enabled = False
        txtTerrDesc.Enabled = False
        cboCategoryCodes.Enabled = False
        cboOnPriceList.Enabled = False
        chkShowActiveOnly.Enabled = False
        cOptionalCriteria.UpdateType = PrcUpdateType.Undetermined.ToString
        itmprclst.Clear()
        With DataGridView1
            Try
                .Rows.Clear()
            Catch ex As Exception

            End Try
            .Columns("ActivePriceColor").Visible = True
            .Columns("ActivePriceRococo").Visible = True
            .Columns("ActivePriceDetailStain").Visible = True
            .Columns("TerFrom").Visible = True
            .Columns("ActivePriceColor").Width = 60
            .Columns("ActivePriceRococo").Width = 60
            .Columns("ActivePriceDetailStain").Width = 60
            .Columns("TerFrom").Width = 60
            .Columns("OriginalPriceColor").ReadOnly = True
            .Columns("OriginalPriceColor").DefaultCellStyle.BackColor = Color.FromArgb(235, 235, 221)
            .Columns("OriginalPriceRococo").DefaultCellStyle.BackColor = Color.FromArgb(235, 235, 221)
            .Columns("OriginalPriceRococo").ReadOnly = True
            .Columns("OriginalPriceDetailStain").DefaultCellStyle.BackColor = Color.FromArgb(235, 235, 221)
            .Columns("OriginalPriceDetailStain").ReadOnly = True
        End With

        If dtTerrCodes.Rows.Count = 0 Or dtTerrCodes.Columns.Count < 4 Then
            FillTerritoryCodeList()
        End If
        fAdvancedLookup = New Lookup(dtTerrCodes)
        fAdvancedLookup.Show()
    End Sub

    'Private Sub LinkLabel1_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)
    '    AdvancedSearch()
    'End Sub

    Public Sub LookupClosed()
        FillTerritoryCodeList()
    End Sub

    'Private Sub mcboTerritoryCodes_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles mcboTerritoryCodes.KeyDown
    '    If e.KeyCode = Keys.Enter Then
    '        With mcboTerritoryCodes
    '            .SelectedItem = .Text
    '            '.ResetText()
    '            '.Refresh()
    '        End With

    '        'MsgBox("Stop")
    '        With Timer1
    '            .Interval = 100
    '            .Enabled = True
    '        End With
    '    End If
    'End Sub

    Private Sub mcboTerritoryCodes_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles mcboTerritoryCodes.SelectedIndexChanged
        If bIsLoading Then Exit Sub
        On Error Resume Next
        'cOptionalCriteria.TerCodeSearchFrom = mcboTerritoryCodes.Data.Rows(Me.mcboTerritoryCodes.SelectedIndex)(0).ToString
        txtTerrDesc.Text = dtTerrCodes.Rows(CType(sender, JTG.ColumnComboBox).SelectedIndex)(1).ToString
        txtFrom.Text = dtTerrCodes.Rows(CType(sender, JTG.ColumnComboBox).SelectedIndex)(3).ToString
        frmSetZonePercentage.Focus()

    End Sub


    Private Sub rbCustomFill_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbCustomFill.CheckedChanged,
                                                                                                        rbZoneFill.CheckedChanged, rbDirectMacolaPrc.CheckedChanged

        Me.Refresh()
        Me.ValidateChildren()

        Dim rb As RadioButton = CType(sender, RadioButton)
        'If bSkipChecked Then Exit Sub
        Select Case rb.Name
            Case rbCustomFill.Name
                'bSkipChecked = True
                If rb.Checked Then
                    FormatPriceType(PrcType.CustomPercent.ToString)
                    cboZonePrc.Text = ""
                Else
                    FormatPriceType(PrcType.Copied.ToString)
                    txtNaturalMarkup.Text = ""
                    txtColorMarkup.Text = ""
                    txtDetailMarkup.Text = ""
                End If

            Case rbZoneFill.Name

                'bSkipChecked = True
                If rb.Checked Then
                    FormatPriceType(MarkupType.Zone.ToString)
                Else
                    FormatPriceType(PrcType.Copied.ToString)
                End If

                If bAdvanceSearch = True Then
                    cboZonePrc.SelectedIndex = 0
                    cboZonePrc.SelectedItem = "Ter Code"
                    cboZonePrc.Text = "Ter Code"
                    cboZonePrc.Enabled = False
                    cboZonePrc.Refresh()

                End If
            Case rbDirectMacolaPrc.Name
                If rb.Checked Then
                    FormatPriceType(MarkupType.Direct.ToString)
                End If
        End Select
        bSkipChecked = False
    End Sub


    Private Sub btnExportExcel_Click(sender As System.Object, e As System.EventArgs) Handles btnExportExcel.Click, btnExportHeader.Click         'btnExcelDataHeader.Click, btnExcelHeaderOnly.Click

        'exportCurrentGridToExcel()
        Try
            Dim dg As DataGridView = DirectCast(Me.DataGridView1, DataGridView)
            Dim FldNames() As String
            FldNames = ExportGridInExcel.GetFixedFieldNames("ItemNumber, ItemDescription, ProdCategory, ProdCatDescription, " &
                                                            "TerCode, ItemLocPriceRococo, OriginalPriceRococo, ActivePriceRococo, " &
                                                            "ItemLocPriceColor, OriginalPriceColor, ActivePriceColor, ItemLocPriceDetailStain, OriginalPriceDetailStain, ActivePriceDetailStain, " &
                                                            "TerFrom, TerDesc, Active, ItemWeight, Dimensions, PageNo")
            Dim rowsCounter As Int32 = 0
            'Dim colsVisible As Int32 = 0
            Dim rows As Int32 = cItemPricingList.Count
            Dim cols As Int32 = UBound(FldNames)
            Dim DataArr(rows, cols) As Object
            Dim rng As Excel.Range

            Dim btn As Button = CType(sender, Button)

            ''Object Data Export: 
            If btn.Name = btnExportExcel.Name Then
                For Each itmprc As ItemPricingObj In cItemPricingList
                    With itmprc

                        DataArr(rowsCounter, 0) = .ItemNo
                        DataArr(rowsCounter, 1) = .ItemDesc
                        DataArr(rowsCounter, 2) = .ProdCat
                        DataArr(rowsCounter, 3) = .ProdCatDesc
                        DataArr(rowsCounter, 4) = .TerCode
                        DataArr(rowsCounter, 5) = .ItemLocPriceRococo
                        DataArr(rowsCounter, 6) = .OriginalPriceRococo
                        DataArr(rowsCounter, 7) = .ActivePriceRococo
                        DataArr(rowsCounter, 8) = .ItemLocPriceColor
                        DataArr(rowsCounter, 9) = .OriginalPriceColor
                        DataArr(rowsCounter, 10) = .ActivePriceColor
                        DataArr(rowsCounter, 11) = .ItemLocPriceDetailStain
                        DataArr(rowsCounter, 12) = .OriginalPriceDetailStain
                        DataArr(rowsCounter, 13) = .ActivePriceDetailStain
                        DataArr(rowsCounter, 14) = .TerFrom
                        DataArr(rowsCounter, 15) = .TerDescription
                        DataArr(rowsCounter, 16) = .A4GLIdentity
                        DataArr(rowsCounter, 17) = BusObj.GetScalarValue("Select itm.item_weight FROM IMITMIDX_SQL itm where item_no = '" & .ItemNo & "'", cn).ToString.Trim
                        DataArr(rowsCounter, 18) = BusObj.GetScalarValue("Select IsNull (itm.user_def_fld_2, '''') as dimensions  FROM IMITMIDX_SQL itm where item_no = '" & .ItemNo & "'", cn).ToString.Trim
                        DataArr(rowsCounter, 19) = BusObj.GetScalarValue("Select dbo.fnCatalogPage (itm.user_def_fld_1) FROM IMITMIDX_SQL itm where item_no = '" & .ItemNo & "'", cn).ToString.Trim

                        rowsCounter += 1
                    End With

                Next
            End If

            'Data Grid Export: If Data with Header Pressed, get the data, otherwise skip ...
            'If btn.Name = btnExportExcel.Name Then
            '    For rowsCounter As Int32 = 0 To rows
            '        For Each cell As DataGridViewCell In dg.Rows(rowsCounter).Cells

            '            If dg.Columns(cell.ColumnIndex).Visible = True Then
            '                DataArr(rowsCounter, colsVisible) = cell.FormattedValue
            '                Debug.Print(cell.FormattedValueType.ToString & " - " & cell.FormattedValue.ToString)
            '                colsVisible = colsVisible + 1

            '            End If
            '            colsCounter = colsCounter + 1

            '        Next
            '        colsVisible = 0
            '        colsCounter = 0
            '    Next
            'End If

            'Excel Variables
            Dim xlapp As New Excel.Application
            Dim xlwbook As Excel.Workbook = xlapp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet)
            Dim xlwsheet As Excel.Worksheet = CType(xlwbook.Worksheets(1), Excel.Worksheet)
            Dim xlcalc As Excel.XlCalculation

            With xlapp
                xlcalc = .Calculation
                .Calculation = Excel.XlCalculation.xlCalculationManual
            End With

            With xlwsheet
                .Range(.Cells(1, 1), .Cells(1, cols + 1)).Value = FldNames
                .Range(.Cells(2, 1), .Cells(rows + 2, cols + 1)).Value = DataArr
                .UsedRange.Columns.AutoFit()
                rng = .Range(.Cells(1, 1), .Cells(rows + 2, cols + 1))
                With rng
                    .HorizontalAlignment = Excel.Constants.xlLeft
                End With
            End With

            With xlapp
                .Visible = True
                .UserControl = True
                .Calculation = xlcalc
            End With

            xlwsheet = Nothing
            xlwbook = Nothing
            xlapp = Nothing
            GC.Collect()

        Catch ex As Exception
            MsgBox("Export Failed")
            MsgBox(ex.Message)

        End Try

    End Sub

    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click

        If MsgBox("Delete the checked Items", MsgBoxStyle.YesNo, "Delete Items") = MsgBoxResult.No Then Exit Sub

        Dim doRefresh As Boolean = True
        DeleteItems(doRefresh)
    End Sub

    Private Sub DeleteItems(doRefresh As Boolean)
        Dim dt As DataTable = GetItemPrcLevelFromObject(cItemPricingList)

        BusObj.DeleteItemsbyTVP(dt, cn)

        clear()
        cItemPricingList.RemoveFilter()

        With Timer1
            .Interval = 1000
            .Enabled = True
        End With

    End Sub

    'Private Sub CloseForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 
    '    Me.Close()
    'End Sub

    Private Sub rbPercentFill_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbPercentFill.CheckedChanged, rbAmountFill.CheckedChanged
        If rbAmountFill.Checked Then
            cOptionalCriteria.FillType = FillType.Amount.ToString
            lblColor.Text = "$ Color"
            lblNatural.Text = "$ Natural"
            lblDetail.Text = "$ Detail"
        Else
            cOptionalCriteria.FillType = FillType.Percent.ToString
            lblColor.Text = "% Color"
            lblNatural.Text = "% Natural"
            lblDetail.Text = "% Detail"
        End If
    End Sub




    'Private Sub rbPricingType_CheckedChanged(sender As Object, e As System.EventArgs)
    '    With cOptionalCriteria
    '        If rbCopiedTerritory.Checked Then
    '            .PricingType = PrcType.Copied.ToString
    '        ElseIf rbPrimaryTerritory.Checked Then
    '            '.PricingType = PrcType.Primary.ToString
    '        ElseIf rbAdvancedSearchPricing.Checked Then
    '            '.PricingType = PrcType.Advanced.ToString
    '        End If
    '    End With


    'End Sub

    'Private Sub cboFillTerrCodes_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboFillTerrCodes.SelectedIndexChanged
    '    Dim cbo As ComboBox = CType(sender, ComboBox)
    '    Dim terDescription As String = ""
    '    Dim terNumber As String = cbo.Text.Trim
    '    Dim sSQL As String = "select distinct MAX(ter_desc) as ter_desc from OEPRCCUS_MAZ where prc_level = '" & terNumber & "'"

    '    terDescription = BusObj.GetScalarValue(sSQL, cn).ToString
    '    Me.txtFillTerrDesc.Text = terDescription
    'End Sub

    Private Sub btnAdvancedSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdvancedSearch.Click
        AdvancedSearch()
    End Sub

    'Private Sub mcboFillTerrCodes_DropDownClosed(sender As Object, e As System.EventArgs) Handles mcboFillTerrCodes.DropDownClosed



    '    With Timer6
    '        .Interval = 500
    '        .Enabled = True
    '    End With



    'End Sub
    'Private Sub Timer6_Tick(sender As System.Object, e As System.EventArgs) Handles Timer6.Tick

    '    Dim tmr As Timer = DirectCast(sender, Timer)
    '    With tmr
    '        .Enabled = False

    '    End With

    '    txtTerCode.Text = dtCopyFrom.Rows(CType(sender, JTG.ColumnComboBox).SelectedIndex)(1).ToString
    '    txtTerFrom.Text = dtTerrCodes.Rows(CType(sender, JTG.ColumnComboBox).SelectedIndex)(1).ToString


    'End Sub
    Private Sub mcboFillTerrCodes_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mcboFillTerrCodes.SelectedIndexChanged
        If bIsLoading Or bByPass Then Exit Sub
        ''On Error Resume Next

        txtFillTerrDesc.Text = dtCopyFrom.Rows(CType(sender, JTG.ColumnComboBox).SelectedIndex)(1).ToString()
        txtTerCode.Text = dtCopyFrom.Rows(CType(sender, JTG.ColumnComboBox).SelectedIndex)(0).ToString
        txtTerFrom.Text = mcboTerritoryCodes.Text.Trim


    End Sub

    'Private Sub Timer5_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer5.Tick
    '    Dim tmr As Timer = CType(sender, Timer)
    '    tmr.Enabled = False
    '    txtFillTerrDesc.Text = terDescFill
    '    btnSave.Enabled = False
    'End Sub

    Private Sub cboZonePrc_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboZonePrc.SelectedIndexChanged
        'Conditional Section
        If rbZoneFill.Checked = True Then
            btnRoundPricing.Enabled = CBool(IIf(cboZonePrc.Text > "", True, False))
            btnSelectAll.Enabled = CBool(IIf(cboZonePrc.Text > "", True, False))
            btnFillPricing.Enabled = CBool(IIf(cboZonePrc.Text > "", True, False))
        End If


        If bAdvanceSearch = True Then
            cboZonePrc.SelectedIndex = 0
            cboZonePrc.SelectedItem = "Ter Code"
            cboZonePrc.Text = "Ter Code"
            cboZonePrc.Enabled = False
            cboZonePrc.Refresh()
        End If

    End Sub


    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub


    Private Sub cboFilter_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboFilter.KeyDown
        If e.KeyCode = Keys.Enter Then
            AddItemToFilterList()
        End If
    End Sub



    Private Sub chkCopiedFromVislble_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCopiedFromVislble.CheckedChanged
        Dim dgv As DataGridView = CType(DataGridView1, DataGridView)
        Dim chk As CheckBox = CType(sender, CheckBox)
        With dgv
            .Columns("CopiedPriceColor").Visible = chk.Checked
            .Columns("CopiedPriceRococo").Visible = chk.Checked
            .Columns("CopiedPriceDetailStain").Visible = chk.Checked
        End With
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim dgv As DataGridView = CType(sender, DataGridView)
        If dgv.Columns(e.ColumnIndex).Name = "Selected" Then
            For Each rw As DataGridViewRow In dgv.Rows
                If Convert.ToBoolean(rw.Cells("Selected").Value) = True Then
                    btnDelete.Enabled = True
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub btnPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPaste.Click
        Dim dgv As DataGridView = CType(DataGridView1, DataGridView)
        Dim cel As DataGridViewCell = dgv.CurrentCell
        Dim rw As Integer = cel.RowIndex
        Dim cl As Integer = cel.ColumnIndex
        'Be sure they have not selected multiple cells to past to ...
        Dim cellsselected As Integer = dgv.SelectedCells.Count
        If cellsselected > 1 Then
            Dim topcell As Integer = dgv.SelectedCells(cellsselected - 1).RowIndex
            rw = topcell
        End If




        Try
            PasteToGrid.PasteClipboard(DataGridView1, cl, rw)
        Catch ex As Exception
            MsgBox("Paste Failed.  Be sure to select only the top cell when pasting to multiple cells.", MsgBoxStyle.OkOnly, "Paste Failed")
        End Try

    End Sub




    Private Sub btnClearSelection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearSelection.Click
        Dim dgv As DataGridView = CType(DataGridView1, DataGridView)
        PasteToGrid.ClearSelection(dgv)


        'Dim emptycell_Decimal As Decimal = 0
        'Dim emptycell_String As String = ""
        'Dim emptycell_Integer As Integer = 0
        'Dim emptycell_Date As Date = #1/1/1900#

        'With dgv
        '    For Each cel As DataGridViewCell In dgv.SelectedCells
        '        Select Case cel.ValueType.FullName
        '            Case "System.String"
        '                cel.Value = emptycell_String
        '            Case "System.Decimal"
        '                cel.Value = emptycell_Decimal
        '            Case "System.Integer"
        '                cel.Value = emptycell_Integer
        '            Case "System.Date"
        '                cel.Value = emptycell_Date
        '        End Select

        '    Next
        'End With

    End Sub

    Private Sub btnCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopy.Click
        Try
            Clipboard.SetDataObject( _
                    Me.DataGridView1.GetClipboardContent())
        Catch ex As Exception

        End Try
    End Sub

    'Private Sub btnByItems_Click(sender As System.Object, e As System.EventArgs) Handles btnByItems.Click
    '    LookupItem.Show()
    'End Sub

    'Private Sub DataGridView1_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseDown
    '    If e.Button = Windows.Forms.MouseButtons.Left Then
    '        Exit Sub
    '    Else

    '        Dim dgv As DataGridView = CType(sender, DataGridView)
    '        'Dim ht As DataGridView.HitTestInfo
    '        ht = dgv.HitTest(e.X, e.Y)
    '        hitContextMenu = dgv.HitTest(e.X, e.Y)

    '        'to set current row programatically, Because current row can be multiple rows
    '        'when row selection is set to MultiSelect, it cannot be set where the little
    '        'black arrow will move to that row.  But setting the Cell to CurrentCell, there is only
    '        'one currentcell, not multiples.  So, set Cell first, then row, and the little black 
    '        'arrow will move. 
    '        Try
    '            With dgv
    '                .ClearSelection()
    '                .CurrentCell = .Rows(ht.RowIndex).Cells(ht.ColumnIndex)
    '                '.Rows(ht.RowIndex).Selected = True
    '            End With
    '        Catch ex As Exception

    '        End Try

    '    End If
    'End Sub

    Private Sub mcboFillTerrCodes_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles mcboFillTerrCodes.TextChanged
        btnSave.Enabled = False
    End Sub

    Private Sub tabOptions_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabOptions.SelectedIndexChanged
        Dim tb As TabControl = CType(sender, TabControl)

        If tb.SelectedIndex = 0 Then
            btnSelectAll.Enabled = True
            btnFillPricing.Enabled = True
            btnRoundPricing.Enabled = True
            btnDelete.Enabled = True
            btnSave.Enabled = False
        Else
            btnSelectAll.Enabled = True
            btnFillPricing.Enabled = False
            btnRoundPricing.Enabled = False
            btnDelete.Enabled = False
            btnSave.Enabled = True
        End If
        rbCustomFill.Checked = False
        rbZoneFill.Checked = False
        'rbPercentFill.Checked = False
        txtColorMarkup.Text = ""
        txtNaturalMarkup.Text = ""
        txtDetailMarkup.Text = ""
    End Sub

    Private Sub btnCopyTo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopyTo.Click
        Dim selected As Boolean = IsAnyRowChecked(Me.DataGridView1)
        If selected = False Then
            MsgBox("No items have been selected.  Select the items to update.", MsgBoxStyle.OkOnly, "Select Items")
            Exit Sub
        End If
        
        If cItemPricingList.Count = 0 Then Exit Sub
        bByPass = True
        bCopyToOption = True
        Cursor = Cursors.WaitCursor
        SetOptionalCopyToCriteria()
        CollectUIFillData()

        DoStandardMarkup()

        Cursor = Cursors.Default

        cOptionalCriteria.IsFillPressed = True
        btnSave.Enabled = cOptionalCriteria.IsFillPressed
        bByPass = False
    End Sub


    Private Sub btnSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectAll.Click
        bByPass = True
        SelectAll()
        bByPass = False
        Dim dgv As DataGridView = CType(DataGridView1, DataGridView)

        With btnDelete
            .Enabled = False
            For Each rw As DataGridViewRow In dgv.Rows
                If Convert.ToBoolean(rw.Cells("Selected").Value) = True Then
                    .Enabled = True
                    Exit Sub
                End If
            Next

        End With

        If cOptionalCriteria.SearchType = SearchType.CopiedPrice.ToString Then SetDataGridBackgroundColor()

    End Sub

    Private Sub btnBackup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBackup.Click
        Dim datestring As String = Now.Year.ToString & "_" & Now.Month.ToString & "_" & Now.Day.ToString & "_" & Now.Hour.ToString & Now.Minute.ToString

        Dim sql As String = "select * into zzOEPRCCUS_MAX_" & datestring & " from dbo.OEPRCCUS_MAZ "
        Try
            BusObj.BackupPriceList(sql, cn)
        Catch ex As Exception
            MsgBox("Backup Failed: Error is " & ex.Message, MsgBoxStyle.OkOnly, "Backup Failed")
            Exit Try
        Finally
            MsgBox("Backup Successful: Table name zzOEPRCCUS_MAX_" & datestring & " created.", MsgBoxStyle.OkCancel, "Backup Successful")
        End Try



    End Sub

   
    Private Sub btnPasteDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPasteDown.Click
        Dim dgv As DataGridView = CType(DataGridView1, DataGridView)
        Dim cel As DataGridViewCell = dgv.CurrentCell
        Dim rw As Integer = cel.RowIndex
        Dim cl As Integer = cel.ColumnIndex
        'Be sure they have not selected multiple cells to past to ...
        Dim cellsselected As Integer = dgv.SelectedCells.Count
        If cellsselected > 1 Then
            Dim topcell As Integer = dgv.SelectedCells(cellsselected - 1).RowIndex
            rw = topcell
        End If




        Try
            Cursor = Cursors.WaitCursor
            For Each r As DataGridViewRow In dgv.Rows
                PasteToGrid.PasteClipboard(DataGridView1, cl, r.Index)
            Next
            Cursor = Cursors.Default
            btnSave.Enabled = True
        Catch ex As Exception
            MsgBox("Paste Failed.  Be sure to select only the top cell when pasting to multiple cells.", MsgBoxStyle.OkOnly, "Paste Failed")
            Cursor = Cursors.Default
        End Try
    End Sub


    'Private Sub cboHasTerrPrice_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboHasTerrPrice.SelectedIndexChanged
    '    Dim cbo As ComboBox = CType(sender, ComboBox)
    '    If cbo.SelectedIndex = 1 Then
    '        txtFrom.Text = ""
    '    End If
    'End Sub

    Private Sub btnSetZonePercentage_Click(sender As System.Object, e As System.EventArgs) Handles btnSetZonePercentage.Click
        frmSetZonePercentage.Show()
    End Sub


    Private Sub rbCustomFill_Click(sender As Object, e As System.EventArgs) Handles rbCustomFill.Click
        Me.ValidateChildren()
        Me.Refresh()
    End Sub

    
    
End Class

Friend Class OptionalCriteria
    Public DBName As String
    Public EnableFillButtons As Boolean
    Public IsPrimaryPricingEnabled As Boolean
    Public IsManualPricingChecked As Boolean
    Public IsFillWithPricingChecked As Boolean
    Public IsFillCopyTowithCopyFromChecked As Boolean

    Public IsStandardMarkupPricingChecked As Boolean
    Public IsUpdatePriceList As Boolean
    Public IsFillPressed As Boolean
    Public IsSkipTerCodeEnabled As Boolean

    Private mCurrentDB As String
    Private mDefaultDB As String
    Private mMarkupType As String

    Public TerCode As String
    Public TerCodeSearchFrom As String
    Public TerCodeSearchTerFromInTable As String

    Public TerCodeFill As String
    Public TerFromCode As String
    Public TerCopiedFromCode As String
    Public TerFromDesc As String
    Public TerCopyToCode As String
    Public TerCopyToDesc As String
    Public TerZoneCode As String

    Public SearchType As String

    Public NatPercent As Double
    Public ClrPercent As Double
    Public DtlPercent As Double

    Public NatAmount As Double
    Public ClrAmount As Double
    Public DtlAmount As Double

    Public ProdCat As String
    Public OnPriceList As String
    Public HasTerrPrice As String
    Public PricingType As String
    Public UpdateType As String
    Public FillType As String

    Public Sub Clear()

        DBName = ""
        mCurrentDB = ""
        mDefaultDB = ""

        IsPrimaryPricingEnabled = True
        IsManualPricingChecked = False
        IsUpdatePriceList = False
        IsFillPressed = False

        NatPercent = 0
        ClrPercent = 0
        DtlPercent = 0

        NatAmount = 0
        ClrAmount = 0
        DtlAmount = 0

        MarkupType = ""

        UpdateType = ""

        FillType = ""
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
            ElseIf TypeOf oValue Is Boolean Then
                value = CBool(oValue)
            End If
            Return value
        End Get
    End Property

    Public ReadOnly Property IsNewPrimaryPricing As Boolean
        Get
            If IsTercodeFillSet AndAlso IsTerDescSet Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

    Public ReadOnly Property IsTercodeFillSet() As Boolean
        Get
            If TerCodeFill Is Nothing Then
                Return False
            Else
                Return IsSet(TerCodeFill.Trim)
            End If

        End Get
    End Property

    Public ReadOnly Property IsTerDescSet() As Boolean
        Get
            Return IsSet(TerFromDesc.Trim)
        End Get
    End Property

    Public ReadOnly Property IsUpdatePriceListSet() As Boolean
        Get
            Return IsSet(IsUpdatePriceList)
        End Get
    End Property

    Public ReadOnly Property IsFillPressedSet() As Boolean
        Get
            Return IsSet(IsFillPressed)
        End Get
    End Property

    Public ReadOnly Property CurrentDB() As String
        Get
            mCurrentDB = "Current DB: " & DBName
            Return mCurrentDB
        End Get
    End Property

    Public ReadOnly Property DefaultDB() As String
        Get
            mDefaultDB = "Default DB: " & DBName
            Return mDefaultDB
        End Get
    End Property

    Public Property MarkupType() As String
        Get
            Return mMarkupType
        End Get
        Set(ByVal value As String)
            mMarkupType = value
        End Set
    End Property

End Class
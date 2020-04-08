Imports System.ComponentModel
Imports System.Data.DataRowView
Imports System.Data
Imports System.Text
Imports System.Data.SqlClient


Public Class BusObj
    Inherits BaseClass

    Public Enum PrcUpdateType
        PrimaryUpdate = 1
        CopiedUpdate = 2
        PrimaryNew = 3
        CopiedNew = 4
        PrimaryUpdateManualPrc = 5
        CopiedUpdateManualPrc = 6
        PrimaryNewManualPrc = 7
        CopiedNewManualPrc = 8
        Undetermined = 9
        Advanced = 10
        Delete = 99
    End Enum

#Region "   Methods   "
#Region "   Search Items for Data Grid View   "

    Public Shared Function GetSearchItems(ByVal items As String, ByVal ProdCat As String, _
                                          ByVal TerCode As String, OnPriceList As String, TerFrom As String, _
                                          HasTerrPrice As String, ByVal cn As SqlConnection) As SqlDataReader

        Dim rd As SqlDataReader

        rd = DAC.ExecuteSP_Reader(My.Resources.SP_spIMGetItemList_MAS, cn, _
        DAC.Parameter(My.Resources.Param_iItems, items, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iProdCategory, ProdCat, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iTerCode, TerCode, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iOnPriceList, OnPriceList, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iTerFrom, TerFrom, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iHasTerrPrice, HasTerrPrice, ParameterDirection.Input))

        Return rd

    End Function


    Public Shared Function ExecuteSQLDataTable(ByVal sql As String, ByVal tablename As String, ByVal cn As SqlConnection) As DataTable
        Dim dt As DataTable

        dt = DAC.ExecuteSQL_DataTable(sql, cn, tablename)
        Return dt

    End Function

    Public Shared Function GetPriceListItems(ByVal sSQL As String, ByVal cn As SqlConnection) As SqlDataReader
        
        Dim rd As SqlDataReader
        rd = DAC.ExecuteSQL_Reader(sSQL, cn)
        Return rd

    End Function
    Public Shared Function GetItemDesc(ItemNo As String, cn As SqlConnection) As String
        Dim ItemDesc As Object = ""

        ItemDesc = DAC.Execute_Scalar("Select item_desc_1 from IMITMIDX_SQL Where item_no = '" & ItemNo & "'", cn)
        If ItemDesc Is Nothing Then
            Return "nothing"
        Else
            Return ItemDesc.ToString.Trim
        End If

    End Function

    Public Shared Sub DeleteItemsbyTVP(dt As DataTable, cn As SqlConnection)
        'the SP has been altered to include only the A4GLIdentity column to join on the TVP
        DAC.Execute_SP_DeleteItemsByTVP(My.Resources.SP_spDeleteItemsByItem_noTVP, cn, _
                                            My.Resources.Param_tvpItem_no, dt)

    End Sub
    Public Shared Sub DeleteA4GLIdentitybyTVP(dt As DataTable, cn As SqlConnection)
        DAC.Execute_SP_DeleteItemsByTVP(My.Resources.SP_spDeleteItemsByA4GLIdentity, cn, _
                                            My.Resources.Param_tvpA4GLIdentity, dt)

    End Sub


    Public Shared Function PopulateSearchItems(ByVal rd As SqlDataReader) As ItemPricingList
        On Error Resume Next
        Dim oItemPricingList(25) As Object
        Dim itmPricingList As New ItemPricingList
        Dim schematable As DataTable
        schematable = rd.GetSchemaTable
        While rd.Read
            If rd(0).Equals(DBNull.Value) Then oItemPricingList(0) = "" Else oItemPricingList(0) = CStr(rd(0)).Trim 'TerCode
            If rd(1).Equals(DBNull.Value) Then oItemPricingList(1) = "" Else oItemPricingList(1) = CStr(rd(1)).Trim 'ItemNo
            If rd(2).Equals(DBNull.Value) Then oItemPricingList(2) = "" Else oItemPricingList(2) = CStr(rd(2)).Trim 'ItemDesc
            If rd(3).Equals(DBNull.Value) Then oItemPricingList(3) = "" Else oItemPricingList(3) = CStr(rd(3)).Trim 'ProdCat
            If rd(4).Equals(DBNull.Value) Then oItemPricingList(4) = "" Else oItemPricingList(4) = CStr(rd(4)).Trim 'ProdDesc
            If rd(5).Equals(DBNull.Value) Then oItemPricingList(5) = 0 Else oItemPricingList(5) = CDec(rd(5)) 'Terr Price Rococo
            If rd(6).Equals(DBNull.Value) Then oItemPricingList(6) = 0 Else oItemPricingList(6) = CDec(rd(6)) 'Terr Price Color
            If rd(7).Equals(DBNull.Value) Then oItemPricingList(7) = 0 Else oItemPricingList(7) = CDec(rd(7)) 'Terr Price Detail
            If rd(8).Equals(DBNull.Value) Then oItemPricingList(8) = 0 Else oItemPricingList(8) = CDec(rd(8)) 'Base Price Rococo
            If rd(9).Equals(DBNull.Value) Then oItemPricingList(9) = 0 Else oItemPricingList(9) = CDec(rd(9)) 'Base Price Color
            If rd(10).Equals(DBNull.Value) Then oItemPricingList(10) = 0 Else oItemPricingList(10) = CDec(rd(10)) 'Base Price Detail

            If rd(11).Equals(DBNull.Value) Then oItemPricingList(11) = 0 Else oItemPricingList(11) = CDec(rd(11)) 'Active Price Rococo - Always 0 so User can set the new price
            If rd(12).Equals(DBNull.Value) Then oItemPricingList(12) = 0 Else oItemPricingList(12) = CDec(rd(12)) 'Active Price Color - Always 0 so User can set the new price
            If rd(13).Equals(DBNull.Value) Then oItemPricingList(13) = 0 Else oItemPricingList(13) = CDec(rd(13)) 'Active Price Detail - Always 0 so User can set the new price

            If rd(14).Equals(DBNull.Value) Then oItemPricingList(14) = 0 Else oItemPricingList(14) = CDec(rd(14)) 'Copied Price Rococo - Copied Prices are the Copied From Prices so we can compare with New Copied Terr Price to see if it's changed and set the DataGridView row to Yellow if it has changed.  
            If rd(15).Equals(DBNull.Value) Then oItemPricingList(15) = 0 Else oItemPricingList(15) = CDec(rd(15)) 'Copied Price Color
            If rd(16).Equals(DBNull.Value) Then oItemPricingList(16) = 0 Else oItemPricingList(16) = CDec(rd(16)) 'Copied Price Detail

            If rd(17).Equals(DBNull.Value) Then oItemPricingList(17) = "" Else oItemPricingList(17) = CStr(rd(17)).Trim 'TerFrom
            If rd(18).Equals(DBNull.Value) Then oItemPricingList(18) = "" Else oItemPricingList(18) = CStr(rd(18)).Trim 'TerDesc

            If rd(19).Equals(DBNull.Value) Then oItemPricingList(19) = CDate("01/01/1900") Else oItemPricingList(19) = CDate(rd(19)) 'LastDate
            oItemPricingList(20) = False                                                                                             'Selected  
            If rd(21).Equals(DBNull.Value) Then oItemPricingList(21) = 0 Else oItemPricingList(21) = CDec(rd(21)) 'ItemWeight
            If rd(22).Equals(DBNull.Value) Then oItemPricingList(22) = 0 Else oItemPricingList(22) = CInt(rd(22)) 'PageNo
            If rd(23).Equals(DBNull.Value) Then oItemPricingList(23) = "" Else oItemPricingList(23) = CStr(rd(23)).Trim 'OnPriceList
            If rd(24).Equals(DBNull.Value) Then oItemPricingList(24) = "" Else oItemPricingList(24) = CStr(rd(24)).Trim 'Dimensions
            If rd(25).Equals(DBNull.Value) Then oItemPricingList(25) = 0 Else oItemPricingList(25) = CInt(rd(25)) 'A4GLIdentity
            Dim itmobj As ItemPricingObj
            Dim o As Object

            itmPricingList.Add(ItemPricingObj.GetPricingObj(CStr(oItemPricingList(0)), CStr(oItemPricingList(1)),
                                 CStr(oItemPricingList(2)), CStr(oItemPricingList(3)), CStr(oItemPricingList(4)),
                                 CDec(oItemPricingList(6)), CDec(oItemPricingList(5)), CDec(oItemPricingList(7)),
                                 CDec(oItemPricingList(9)), CDec(oItemPricingList(8)), CDec(oItemPricingList(10)),
                                 CDec(oItemPricingList(12)), CDec(oItemPricingList(11)), CDec(oItemPricingList(13)),
                                 CDec(oItemPricingList(15)), CDec(oItemPricingList(14)), CDec(oItemPricingList(16)),
                                 CStr(oItemPricingList(17)), CStr(oItemPricingList(18)), CDate(oItemPricingList(19)),
                                 CBool(oItemPricingList(20)), CDec(oItemPricingList(21)),
                                 CInt(oItemPricingList(22)), CStr(oItemPricingList(23)), CStr(oItemPricingList(24)), CInt(oItemPricingList(25))))

        End While
        rd.Close()
        Return itmPricingList

    End Function

    Public Shared Function PopulateSearchItems(ByVal rd As DataTableReader) As ItemPricingList
        On Error Resume Next
        Dim oItemPricingList(25) As Object
        Dim itmPricingList As New ItemPricingList
        Dim schematable As DataTable
        schematable = rd.GetSchemaTable
        While rd.Read
            If (Not rd(17).Equals(DBNull.Value)) AndAlso Convert.ToBoolean(rd(17)) = True Then
                If rd(0).Equals(DBNull.Value) Then oItemPricingList(0) = "" Else oItemPricingList(0) = CStr(rd(0)).Trim 'TerCode
                If rd(1).Equals(DBNull.Value) Then oItemPricingList(1) = "" Else oItemPricingList(1) = CStr(rd(1)).Trim 'ItemNo
                If rd(2).Equals(DBNull.Value) Then oItemPricingList(2) = "" Else oItemPricingList(2) = CStr(rd(2)).Trim 'ItemDesc
                If rd(3).Equals(DBNull.Value) Then oItemPricingList(3) = "" Else oItemPricingList(3) = CStr(rd(3)).Trim 'ProdCat
                If rd(4).Equals(DBNull.Value) Then oItemPricingList(4) = "" Else oItemPricingList(4) = CStr(rd(4)).Trim 'ProdDesc
                If rd(5).Equals(DBNull.Value) Then oItemPricingList(5) = 0 Else oItemPricingList(5) = CDec(rd(5)) 'Terr Price Rococo
                If rd(6).Equals(DBNull.Value) Then oItemPricingList(6) = 0 Else oItemPricingList(6) = CDec(rd(6)) 'Terr Price Color
                If rd(7).Equals(DBNull.Value) Then oItemPricingList(7) = 0 Else oItemPricingList(7) = CDec(rd(7)) 'Terr Price Detail
                If rd(8).Equals(DBNull.Value) Then oItemPricingList(8) = 0 Else oItemPricingList(8) = CDec(rd(8)) 'Base Price Rococo
                If rd(9).Equals(DBNull.Value) Then oItemPricingList(9) = 0 Else oItemPricingList(9) = CDec(rd(9)) 'Base Price Color
                If rd(10).Equals(DBNull.Value) Then oItemPricingList(10) = 0 Else oItemPricingList(10) = CDec(rd(10)) 'Base Price Detail

                If rd(11).Equals(DBNull.Value) Then oItemPricingList(11) = 0 Else oItemPricingList(11) = CDec(rd(11)) 'Active Price Rococo - Always 0 so User can set the new price
                If rd(12).Equals(DBNull.Value) Then oItemPricingList(12) = 0 Else oItemPricingList(12) = CDec(rd(12)) 'Active Price Color - Always 0 so User can set the new price
                If rd(13).Equals(DBNull.Value) Then oItemPricingList(13) = 0 Else oItemPricingList(13) = CDec(rd(13)) 'Active Price Detail - Always 0 so User can set the new price

                If rd(14).Equals(DBNull.Value) Then oItemPricingList(14) = 0 Else oItemPricingList(14) = CDec(rd(14)) 'Copied Price Rococo - Copied Prices are the Copied From Prices so we can compare with New Copied Terr Price to see if it's changed and set the DataGridView row to Yellow if it has changed.  
                If rd(15).Equals(DBNull.Value) Then oItemPricingList(15) = 0 Else oItemPricingList(15) = CDec(rd(15)) 'Copied Price Color
                If rd(16).Equals(DBNull.Value) Then oItemPricingList(16) = 0 Else oItemPricingList(16) = CDec(rd(16)) 'Copied Price Detail

                If rd(17).Equals(DBNull.Value) Then oItemPricingList(17) = "" Else oItemPricingList(17) = CStr(rd(17)).Trim 'TerFrom
                If rd(18).Equals(DBNull.Value) Then oItemPricingList(18) = "" Else oItemPricingList(18) = CStr(rd(18)).Trim 'TerDesc

                If rd(19).Equals(DBNull.Value) Then oItemPricingList(19) = CDate("01/01/1900") Else oItemPricingList(19) = CDate(rd(19)) 'LastDate
                oItemPricingList(20) = False                                                                                             'Selected  
                If rd(21).Equals(DBNull.Value) Then oItemPricingList(21) = 0 Else oItemPricingList(21) = CDec(rd(21)) 'ItemWeight
                If rd(22).Equals(DBNull.Value) Then oItemPricingList(22) = 0 Else oItemPricingList(22) = CInt(rd(22)) 'PageNo
                If rd(23).Equals(DBNull.Value) Then oItemPricingList(23) = "" Else oItemPricingList(23) = CStr(rd(23)).Trim 'OnPriceList
                If rd(24).Equals(DBNull.Value) Then oItemPricingList(24) = "" Else oItemPricingList(24) = CStr(rd(24)).Trim 'Dimensions
                If rd(25).Equals(DBNull.Value) Then oItemPricingList(25) = 0 Else oItemPricingList(25) = CInt(rd(25)) 'A4GLIdentity
                Dim itmobj As ItemPricingObj
                Dim o As Object

                itmPricingList.Add(ItemPricingObj.GetPricingObj(CStr(oItemPricingList(0)), CStr(oItemPricingList(1)),
                                     CStr(oItemPricingList(2)), CStr(oItemPricingList(3)), CStr(oItemPricingList(4)),
                                     CDec(oItemPricingList(6)), CDec(oItemPricingList(5)), CDec(oItemPricingList(7)),
                                     CDec(oItemPricingList(9)), CDec(oItemPricingList(8)), CDec(oItemPricingList(10)),
                                     CDec(oItemPricingList(12)), CDec(oItemPricingList(11)), CDec(oItemPricingList(13)),
                                     CDec(oItemPricingList(15)), CDec(oItemPricingList(14)), CDec(oItemPricingList(16)),
                                     CStr(oItemPricingList(17)), CStr(oItemPricingList(18)), CDate(oItemPricingList(19)),
                                     CBool(oItemPricingList(20)), CDec(oItemPricingList(21)),
                                     CInt(oItemPricingList(22)), CStr(oItemPricingList(23)), CStr(oItemPricingList(24)), CInt(oItemPricingList(25))))

            End If

        End While
        rd.Close()
        Return itmPricingList

    End Function

    'TODO FIX THIS
    'Public Shared Function PopulateSearchItems(ByVal ds As DataSet) As ItemPricingList

    '    'Called from frmExcelImport CheckList Control, chklstExcelSheetNames_ItemCheck Event ... 
    '    ' Field Order of Excel Import:  0-ItemNo   1-ItemDesc  2-Cat  3-CatDesc  4-Terr  5-B-Nat  6-B-Color  7-B-Detail
    '    Dim oItemPricingList(15) As Object
    '    Dim itmPricingList As New ItemPricingList
    '    Dim dt As DataTable = ds.Tables(0)
    '    Dim rw As DataRow
    '    For Each rw In dt.Rows

    '        If rw(4).Equals(DBNull.Value) Then oItemPricingList(0) = "" Else oItemPricingList(0) = CStr(rw(4)) 'TerCode 
    '        If rw(0).Equals(DBNull.Value) Then oItemPricingList(1) = "" Else oItemPricingList(1) = CStr(rw(0)) 'ItemNo
    '        If rw(1).Equals(DBNull.Value) Then oItemPricingList(2) = "" Else oItemPricingList(2) = CStr(rw(1)) 'ItemDesc
    '        If rw(2).Equals(DBNull.Value) Then oItemPricingList(3) = "" Else oItemPricingList(3) = CStr(rw(2)) 'ProdCat
    '        If rw(3).Equals(DBNull.Value) Then oItemPricingList(4) = "" Else oItemPricingList(4) = CStr(rw(3)) 'ProdDesc
    '        oItemPricingList(5) = 0  'Terr Price Nat
    '        oItemPricingList(6) = 0  'Terr Price Color
    '        oItemPricingList(7) = 0 'Terr Price Detail
    '        If rw(5).Equals(DBNull.Value) Then oItemPricingList(8) = 0 Else oItemPricingList(8) = CDec(rw(5)) 'Base Price Nat
    '        If rw(6).Equals(DBNull.Value) Then oItemPricingList(9) = 0 Else oItemPricingList(9) = CDec(rw(6)) 'Base Price Color
    '        If rw(7).Equals(DBNull.Value) Then oItemPricingList(10) = 0 Else oItemPricingList(10) = CDec(rw(7)) 'Base Price Detail
    '        oItemPricingList(11) = CDate("01/01/1900") 'LastDate
    '        oItemPricingList(12) = False 'itm_selected
    '        oItemPricingList(13) = 0 'ItemWeight
    '        oItemPricingList(14) = 0 'Page_no
    '        oItemPricingList(15) = "" 'ActiveItem (not used for import, only read from database)

    '        itmPricingList.Add(ItemPricingObj.GetPricingObj(CStr(oItemPricingList(0)), CStr(oItemPricingList(1)), _
    '                             CStr(oItemPricingList(2)), CStr(oItemPricingList(3)), CStr(oItemPricingList(4)), _
    '                             CDec(oItemPricingList(5)), CDec(oItemPricingList(6)), CDec(oItemPricingList(7)), _
    '                             CDec(oItemPricingList(8)), CDec(oItemPricingList(9)), CDec(oItemPricingList(10)), _
    '                             CDate(oItemPricingList(11)), CBool(oItemPricingList(12)), CDec(oItemPricingList(13)), CInt(oItemPricingList(14)), CStr(oItemPricingList(15))))

    '    Next
    '    ds.Dispose()
    '    dt.Dispose()
    '    Return itmPricingList

    'End Function
    Public Shared Function AppendSearchItems(ByVal rd As SqlDataReader, ByVal itmprclst As ItemPricingList) As ItemPricingList
        Dim oItemPricingList(24) As Object
        Dim itmPricingList As New ItemPricingList
        While rd.Read
            If rd(0).Equals(DBNull.Value) Then oItemPricingList(0) = "" Else oItemPricingList(0) = CStr(rd(0)).Trim 'TerCode
            If rd(1).Equals(DBNull.Value) Then oItemPricingList(1) = "" Else oItemPricingList(1) = CStr(rd(1)).Trim 'ItemNo
            If rd(2).Equals(DBNull.Value) Then oItemPricingList(2) = "" Else oItemPricingList(2) = CStr(rd(2)).Trim 'ItemDesc
            If rd(3).Equals(DBNull.Value) Then oItemPricingList(3) = "" Else oItemPricingList(3) = CStr(rd(3)).Trim 'ProdCat
            If rd(4).Equals(DBNull.Value) Then oItemPricingList(4) = "" Else oItemPricingList(4) = CStr(rd(4)).Trim 'ProdDesc
            If rd(5).Equals(DBNull.Value) Then oItemPricingList(5) = 0 Else oItemPricingList(5) = CDec(rd(5)) 'Terr Price Nat
            If rd(6).Equals(DBNull.Value) Then oItemPricingList(6) = 0 Else oItemPricingList(6) = CDec(rd(6)) 'Terr Price Color
            If rd(7).Equals(DBNull.Value) Then oItemPricingList(7) = 0 Else oItemPricingList(7) = CDec(rd(7)) 'Terr Price Detail
            If rd(8).Equals(DBNull.Value) Then oItemPricingList(8) = 0 Else oItemPricingList(8) = CDec(rd(8)) 'Base Price Nat
            If rd(9).Equals(DBNull.Value) Then oItemPricingList(9) = 0 Else oItemPricingList(9) = CDec(rd(9)) 'Base Price Color
            If rd(10).Equals(DBNull.Value) Then oItemPricingList(10) = 0 Else oItemPricingList(10) = CDec(rd(10)) 'Base Price Detail
            If rd(11).Equals(DBNull.Value) Then oItemPricingList(11) = 0 Else oItemPricingList(11) = CDec(rd(11)) 'Active Price Nat - Always 0 so User can set the new price
            If rd(12).Equals(DBNull.Value) Then oItemPricingList(12) = 0 Else oItemPricingList(12) = CDec(rd(12)) 'Active Price Color - Always 0 so User can set the new price
            If rd(13).Equals(DBNull.Value) Then oItemPricingList(13) = 0 Else oItemPricingList(13) = CDec(rd(13)) 'Active Price Detail - Always 0 so User can set the new price

            If rd(14).Equals(DBNull.Value) Then oItemPricingList(14) = 0 Else oItemPricingList(14) = CDec(rd(14)) 'Copied Price Nat - Copied Prices are the Copied From Prices so we can compare with New Copied Terr Price to see if it's changed and set the DataGridView row to Yellow if it has changed.  
            If rd(15).Equals(DBNull.Value) Then oItemPricingList(15) = 0 Else oItemPricingList(15) = CDec(rd(15)) 'Copied Price Color
            If rd(16).Equals(DBNull.Value) Then oItemPricingList(16) = 0 Else oItemPricingList(16) = CDec(rd(16)) 'Copied Price Detail

            If rd(17).Equals(DBNull.Value) Then oItemPricingList(17) = "" Else oItemPricingList(17) = CStr(rd(17)).Trim 'TerFrom
            If rd(18).Equals(DBNull.Value) Then oItemPricingList(18) = "" Else oItemPricingList(18) = CStr(rd(18)).Trim 'TerDesc

            If rd(19).Equals(DBNull.Value) Then oItemPricingList(19) = CDate("01/01/1900") Else oItemPricingList(19) = CDate(rd(19)) 'LastDate
            oItemPricingList(20) = False                                                                                             'Selected  
            If rd(21).Equals(DBNull.Value) Then oItemPricingList(21) = 0 Else oItemPricingList(21) = CDec(rd(21)) 'ItemWeight
            If rd(22).Equals(DBNull.Value) Then oItemPricingList(22) = 0 Else oItemPricingList(22) = CInt(rd(22)) 'PageNo
            If rd(23).Equals(DBNull.Value) Then oItemPricingList(23) = "" Else oItemPricingList(23) = CStr(rd(23)).Trim 'OnPriceList
            If rd(24).Equals(DBNull.Value) Then oItemPricingList(24) = "" Else oItemPricingList(24) = CStr(rd(24)).Trim 'Dimensions
            If rd(25).Equals(DBNull.Value) Then oItemPricingList(25) = 0 Else oItemPricingList(25) = CInt(rd(25)) 'A4GLIdentity

            itmPricingList.Add(ItemPricingObj.GetPricingObj(CStr(oItemPricingList(0)), CStr(oItemPricingList(1)),
                                 CStr(oItemPricingList(2)), CStr(oItemPricingList(3)), CStr(oItemPricingList(4)),
                                 CDec(oItemPricingList(6)), CDec(oItemPricingList(5)), CDec(oItemPricingList(7)),
                                 CDec(oItemPricingList(9)), CDec(oItemPricingList(8)), CDec(oItemPricingList(10)),
                                 CDec(oItemPricingList(12)), CDec(oItemPricingList(11)), CDec(oItemPricingList(13)),
                                 CDec(oItemPricingList(15)), CDec(oItemPricingList(14)), CDec(oItemPricingList(16)),
                                 CStr(oItemPricingList(17)), CStr(oItemPricingList(18)), CDate(oItemPricingList(19)),
                                 CBool(oItemPricingList(20)), CDec(oItemPricingList(21)),
                                 CInt(oItemPricingList(22)), CStr(oItemPricingList(23)), CStr(oItemPricingList(24)), CInt(oItemPricingList(25))))


        End While
        rd.Close()
        Return itmprclst

    End Function

    Public Shared Function PopulatePriceListItems(ByVal rd As SqlDataReader) As PriceListList
        Dim oPriceList(7) As Object
        Dim PriceList As New PriceListList
        While rd.Read
            If rd(0).Equals(DBNull.Value) Then oPriceList(0) = "" Else oPriceList(0) = CStr(rd(0)) 'TerCode
            If rd(1).Equals(DBNull.Value) Then oPriceList(1) = 0 Else oPriceList(1) = CStr(rd(1)) 'ItemNo
            If rd(2).Equals(DBNull.Value) Then oPriceList(2) = "" Else oPriceList(2) = CStr(rd(2)) 'ItemDesc
            If rd(3).Equals(DBNull.Value) Then oPriceList(3) = 0 Else oPriceList(3) = CDec(rd(3)) 'ProdCat
            If rd(4).Equals(DBNull.Value) Then oPriceList(4) = 0 Else oPriceList(4) = CDec(rd(4)) 'ProdDesc
            If rd(5).Equals(DBNull.Value) Then oPriceList(5) = 0 Else oPriceList(5) = CDec(rd(5)) 'Terr Price Nat
            If rd(6).Equals(DBNull.Value) Then oPriceList(6) = 0 Else oPriceList(6) = CDec(rd(6)) 'ItemWeight
            If rd(7).Equals(DBNull.Value) Then oPriceList(7) = 0 Else oPriceList(7) = CStr(rd(7)) 'TerritoryCode

            PriceList.Add(PriceListObj.GetPriceListObj(CStr(oPriceList(0)), CStr(oPriceList(1)), _
                                 CStr(oPriceList(2)), CDec(oPriceList(3)), CDec(oPriceList(4)), _
                                 CDec(oPriceList(5)), CDec(oPriceList(6)), CStr(oPriceList(7))))

        End While
        rd.Close()
        Return PriceList

    End Function
#End Region

#Region "   Item Numbers Lookup   "

    'Public Shared Function GetItemLookupDataTable(ByVal WhereClause As String, ByVal IncludeTerritory As Boolean, ByVal cn As SqlConnection) As DataTable

    '    Dim dt As DataTable
    '    Dim sp As String
    '    If IncludeTerritory = True Then
    '        sp = My.Resources.SP_spIMGetItemLookup_NoTerritory_MAS
    '    Else
    '        sp = My.Resources.SP_spIMGetItemLookup_MAS
    '    End If
    '    dt = DAC.ExecuteSP_DataTable(sp, cn, _
    '    DAC.Parameter(My.Resources.Param_iWhere, WhereClause, ParameterDirection.Input))

    '    Return dt

    'End Function
    Public Shared Function GetItem(item_no As String, cn As SqlConnection) As DataTable
        Dim dt As DataTable
        Dim sSQL As String = "With ctePrice (item_no, prc_natural, prc_color, prc_detail, prc_level) " & vbCrLf _
                           & "as" & vbCrLf _
                           & "(" & vbCrLf _
                           & "Select Distinct item_no, prc_natural, prc_color, prc_detail, prc_level  " & vbCrLf _
                           & "from OEPRCCUS_MAZ " & vbCrLf _
                           & "where item_no = '" & item_no & "' " & vbCrLf _
                           & "and prc_level not in ('002', '003', '004', '005', '006') " & vbCrLf _
                           & ")" & vbCrLf _
                           & "SELECT C.prc_level, A.item_no, A.item_desc_1, B.prod_cat, " & vbCrLf _
                           & "B.prod_cat_desc, C.prc_natural, C.prc_color, C.prc_detail," & vbCrLf _
                           & "D.price AS base_prc_nat, D.sls_price AS base_prc_col, " & vbCrLf _
                           & "Case " & vbCrLf _
                           & "  When D.user_def_fld_1 is not null " & vbCrLf _
                           & "  Then Cast((Cast(D.user_def_fld_1 as bigint) * .000001) as Decimal(12, 5))" & vbCrLf _
                           & "  Else null End as base_prc_det, " & vbCrLf _
                           & "0 as active_natural , 0 as active_color , 0 as active_detail , " & vbCrLf _
                           & "ISNULL(cte.prc_natural, 0) as copiedbas_natural , ISNULL(cte.prc_color, 0) as copiedbas_color , ISNULL(cte.prc_detail, 0) as copiedbas_detail , " & vbCrLf _
                           & "C.ter_from , C.ter_desc , C.lastdate, " & vbCrLf _
                           & "0 as itm_selected, A.item_weight, 0 as page_no, A.user_def_cd as onpricelist, IsNull (A.user_def_fld_2, '''') as dimensions, C.A4GLIdentity " & vbCrLf _
                           & "FROM IMITMIDX_SQL A " & vbCrLf _
                           & "Left Join IMCATFIL_SQL B " & vbCrLf _
                           & "	On A.prod_cat = B.prod_cat " & vbCrLf _
                           & "Left Join OEPRCCUS_MAZ C " & vbCrLf _
                           & "	On A.item_no = C.item_no " & vbCrLf _
                           & "Join IMINVLOC_SQL D " & vbCrLf _
                           & "	On A.item_no  = D.item_no " & vbCrLf _
                           & "Left Join ctePrice cte " & vbCrLf _
                           & "	on cte.item_no = C.item_no " & vbCrLf _
                           & "	and cte.prc_level = c.prc_level " & vbCrLf _
                           & "Where D.loc = '001'" _
                           & "  And A.item_no = '" & item_no & "'"

        dt = DAC.ExecuteSQL_DataSet(sSQL, cn, "Item")
        Return dt
    End Function

    Public Shared Function GetItem(item_no As String, prc_level As String, cn As SqlConnection) As DataTable
        Dim dt As DataTable
        Dim sSQL As String = "With ctePrice (item_no, prc_natural, prc_color, prc_detail, prc_level) " & vbCrLf _
                           & "as" & vbCrLf _
                           & "(" & vbCrLf _
                           & "Select Distinct item_no, prc_natural, prc_color, prc_detail, prc_level  " & vbCrLf _
                           & "from OEPRCCUS_MAZ " & vbCrLf _
                           & "where item_no = '" & item_no & "' " & vbCrLf _
                           & "and prc_level not in ('002', '003', '004', '005', '006') " & vbCrLf _
                           & ")" & vbCrLf _
                           & "SELECT C.prc_level, A.item_no, A.item_desc_1, B.prod_cat, " & vbCrLf _
                           & "B.prod_cat_desc, C.prc_natural, C.prc_color, C.prc_detail," & vbCrLf _
                           & "D.price AS base_prc_nat, D.sls_price AS base_prc_col, " & vbCrLf _
                           & "Case " & vbCrLf _
                           & "  When D.user_def_fld_1 is not null " & vbCrLf _
                           & "  Then Cast((Cast(D.user_def_fld_1 as bigint) * .000001) as Decimal(12, 5))" & vbCrLf _
                           & "  Else null End as base_prc_det, " & vbCrLf _
                           & "0 as active_natural , 0 as active_color , 0 as active_detail , " & vbCrLf _
                           & "ISNULL(cte.prc_natural, 0) as copiedbas_natural , ISNULL(cte.prc_color, 0) as copiedbas_color , ISNULL(cte.prc_detail, 0) as copiedbas_detail , " & vbCrLf _
                           & "C.ter_from , C.ter_desc , C.lastdate, " & vbCrLf _
                           & "0 as itm_selected, A.item_weight, 0 as page_no, A.user_def_cd as onpricelist, IsNull (A.user_def_fld_2, '''') as dimensions, C.A4GLIdentity " & vbCrLf _
                           & "FROM IMITMIDX_SQL A " & vbCrLf _
                           & "Left Join IMCATFIL_SQL B " & vbCrLf _
                           & "	On A.prod_cat = B.prod_cat " & vbCrLf _
                           & "Left Join OEPRCCUS_MAZ C " & vbCrLf _
                           & "	On A.item_no = C.item_no " & vbCrLf _
                           & "Join IMINVLOC_SQL D " & vbCrLf _
                           & "	On A.item_no  = D.item_no " & vbCrLf _
                           & "Left Join ctePrice cte " & vbCrLf _
                           & "	on cte.item_no = C.item_no " & vbCrLf _
                           & "	and cte.prc_level = c.prc_level " & vbCrLf _
                           & "Where D.loc = '001'" _
                           & "  And A.item_no = '" & item_no & "'" _
                           & "  And C.prc_level in (" & prc_level & ") "

        dt = DAC.ExecuteSQL_DataSet(sSQL, cn, "Item")
        Return dt
    End Function


    Public Shared Function GetItemLookupDataTable(ByVal WhereClause As String, CTEWhereClause As String, ByVal cn As SqlConnection) As DataTable

        Dim dt As DataTable
        Dim sp As String

        sp = My.Resources.SP_spIMGetItemLookup_MAS
        'Select Case IsTerrCodeSet
        '    Case True

        '    Case False
        '        sp = My.Resources.SP_spIMGetItemLookup_NoTerritory_MAS
        'End Select

        dt = DAC.ExecuteSP_DataTable(sp, cn, _
        DAC.Parameter(My.Resources.Param_iWhere, WhereClause, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iCTEWhere, CTEWhereClause, ParameterDirection.Input))

        Return dt

    End Function

    'Public Shared Function CheckItemState(ByVal ItemNo As String, ByVal TerCode As String, ByVal cn As SqlConnection) As Integer
    '    Dim sSQL As String
    '    Dim obj As Object
    '    sSQL = "Select Count(*) from OEPRCCUS_MAZ Where item_no = '" & ItemNo.Trim & "' and prc_level = '" & TerCode.Trim & "'" & "' and ter_from = '" & Terfrom.Trim & "'"
    '    obj = DAC.Execute_Scalar(sSQL, cn)

    '    If obj Is Nothing Then
    '        Return 0
    '    Else
    '        Return CInt(obj)
    '    End If

    'End Function

    Public Shared Function CheckCopiedItemState(ByVal ItemNo As String, ByVal TerCode As String, TerFrom As String, ByVal cn As SqlConnection) As Integer
        Dim sSQL As String = ""
        Dim obj As Object

        sSQL = "Select Count(*) from OEPRCCUS_MAZ Where IsNull(item_no, '') = '" & ItemNo.Trim & "' and IsNull(prc_level, '') = '" _
             & TerCode.Trim & "' and IsNull(ter_from, '') = '" & TerFrom & "'"

        obj = DAC.Execute_Scalar(sSQL, cn)

        If obj Is Nothing Then
            Return 0
        Else
            Return CInt(obj)
        End If

    End Function

    Public Shared Function GetTerritoryCodes(ByVal cn As SqlConnection) As SqlDataReader
        Dim sSQL As String

        sSQL = "Select distinct prc_level from OEPRCCUS_MAZ Order By prc_level"
        GetTerritoryCodes = DAC.ExecuteSQL_Reader(sSQL, cn)

    End Function
    Public Shared Function GetTerritoryCodeList(ByVal cn As SqlConnection) As DataTable
        Dim sSQL As String

        sSQL = "Delete from OEPRCCUS_MAZ where prc_level = ''"
        DAC.Execute_SQL(sSQL, cn)

        sSQL = "Select distinct lTrim(rTrim(prc_level)) as prc_level, IsNull(ter_desc, '') as ter_desc, Case when (ter_from is not null and lTrim(rTrim(ter_from)) > '') then 'copied from: ' else '' end as copied_from, IsNull(ter_from, '') as ter_from from OEPRCCUS_MAZ where prc_level <> '' Order by prc_level    "
        GetTerritoryCodeList = DAC.ExecuteSQL_DataSet(sSQL, cn, "TerritoryCodes")
        'GetTerritoryCodeList = DAC.ExecuteSQL_Reader(sSQL, cn)

    End Function
    Public Shared Function GetCategoryCodes(ByVal ShowActive As Integer, ByVal cn As SqlConnection) As SqlDataReader

        'GetCategoryCodes = DAC.ExecuteSP_Reader(My.Resources.SP_spIMGetProductCategories_MAS, cn)
        GetCategoryCodes = DAC.ExecuteSP_Reader(My.Resources.SP_spIMGetProdCatActiveItems_MAS, cn, _
                                                DAC.Parameter(My.Resources.Param_iShowOnlyActive, ShowActive, ParameterDirection.Input))
    End Function
    Public Shared Function GetFreightMarkups(sSQL As String, ByVal cn As SqlConnection) As SqlDataReader

        GetFreightMarkups = DAC.ExecuteSQL_Reader(sSQL, cn)

    End Function

    Public Shared Function GetZonePriceList(sSQL As String, tablename As String, cn As SqlConnection) As DataTable

        GetZonePriceList = DAC.ExecuteSQL_DataTable(sSQL, cn, tablename)

    End Function

    Public Shared Function GetScalarValue(sSQL As String, cn As SqlConnection) As Object
        Dim obj As Object
        obj = DAC.Execute_Scalar(sSQL, cn)
        Return obj
    End Function

    Public Shared Sub BackupPriceList(ByVal sql As String, ByVal cn As SqlConnection)
        DAC.Execute_NonSQL(sql, cn)

    End Sub

    Public Shared Sub ItemObjectToSQLExport(ByVal itemno As String, ByVal itemdesc As String, ByVal prodcat As String, _
                                      ByVal prodcatdesc As String, ByVal tercode As String, _
                                      ByVal macprcnatural As Decimal, ByVal prcnatural As Decimal, ByVal newprcnatural As Decimal, _
                                      ByVal macprccolor As Decimal, ByVal prccolor As Decimal, ByVal newprccolor As Decimal, _
                                      ByVal macprcdetail As Decimal, ByVal prcdetail As Decimal, ByVal newprcdetail As Decimal, _
                                      ByVal terfrom As String, terdesc As String, activeitm As String, _
                                      ByVal lastdate As Date, ByVal cn As SqlConnection)

        DAC.ExecuteSaveSP(My.Resources.SP_spIMGetItemList_MAS, cn, _
        DAC.Parameter(My.Resources.Param_iItemNo, itemno, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iItemDesc, itemdesc, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iProdCategory, prodcat, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iProdCategory, prodcatdesc, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iTerCode, tercode, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iOnPriceList, macprcnatural, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iOnPriceList, prcnatural, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iOnPriceList, newprcnatural, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iOnPriceList, macprccolor, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iOnPriceList, prccolor, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iOnPriceList, newprccolor, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iOnPriceList, macprcdetail, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iOnPriceList, prcdetail, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iOnPriceList, newprcdetail, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iTerFrom, terfrom, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iTerFrom, terdesc, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iTerFrom, activeitm, ParameterDirection.Input), _
        DAC.Parameter(My.Resources.Param_iHasTerrPrice, lastdate, ParameterDirection.Input))

    End Sub


#End Region
#End Region
End Class

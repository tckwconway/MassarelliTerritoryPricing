Imports System.Data.SqlClient
Imports System.Data
Imports System.Text


Public Class frmSetZonePercentage


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click

    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmSetZonePercentage_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Dim sql As String = "Select ter_zone [Terr Code], frt_markup [Pct Markup], ter_from [Terr From], ter_desc [Terr Description] from oeprczon_mas "
        Dim dt As DataTable
        dt = BusObj.ExecuteSQLDataTable(sql, "ZonePct", cn)

        With DataGridView1
            .DataSource = dt
        End With
    End Sub
End Class
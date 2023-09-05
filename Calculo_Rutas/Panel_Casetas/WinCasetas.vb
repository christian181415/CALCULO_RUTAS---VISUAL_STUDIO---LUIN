Imports System.ComponentModel

Public Class WinCasetas
    Dim NewConfig As New ClassCasetasData
    Dim Activador As Integer = 1
    Private Sub FormCasetasConfig_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim r As Rectangle = My.Computer.Screen.WorkingArea
        Dim Largo = (r.Width / 2) - 350
        Dim Alto = (r.Height / 2) - 210
        Location = New Point(Largo, Alto)

        If LNewCaseta.Text = "NuevaCasetaRuta" Then
            Me.Size = New Size(707, 444)
            P_CasetaRuta.Location = New Point(0, 0)
            Dim dtDestinos As DataTable = NewConfig.MostrarDestinosCR(CmbCRDestino)
            CmbCRDestino.DataSource = dtDestinos
            CmbCRDestino.DisplayMember = "Nombre"
            CmbCRDestino.SelectedIndex = -1
            LIDRutaC.Text = "LIDRuta"
            LIDVehiculo.Text = "LIDVehiculo"

            Dim dtCasetas As DataTable = NewConfig.MostrarCasetasCR(DTGCasetaExists)
            DTGCasetaExists.Columns.Clear()
            DTGCasetaExists.DataSource = Nothing
            DTGCasetaExists.DataSource = dtCasetas
        End If

        If LUpdateCaseta.Text = "UpdateCasetaRuta" Then
            Me.Size = New Size(707, 444)
            P_UpCasetaRuta.Location = New Point(0, 0)
            NewConfig.ObtenerDestino_Vehiculo(LIDRutaUp, LNombreDestino, LIDVehiculoUp, LNombreVehiculo)

            Dim dtCasetas As DataTable = NewConfig.MostrarCasetasCR(DTGCasetaExistsUp)
            DTGCasetaExistsUp.Columns.Clear()
            DTGCasetaExistsUp.DataSource = Nothing
            DTGCasetaExistsUp.DataSource = dtCasetas

            Dim dtMyCasetas As DataTable = NewConfig.ObtenerCasetasRuta(LIDRutaUp, LIDVehiculoUp)
            DTGCasetaSelectUp.DataSource = Nothing
            DTGCasetaSelectUp.Columns.Clear()
            DTGCasetaSelectUp.DataSource = dtMyCasetas
            DTGCasetaSelectUp.Columns("Casetas").ReadOnly = True
            DTGCasetaSelectUp.Columns("Casetas").Width = 271
            DTGCasetaSelectUp.Columns("Importe").Width = 53
            DTGCasetaExistsUp.Sort(DTGCasetaExistsUp.Columns(0), ListSortDirection.Ascending)
        End If
    End Sub
    Public Function CloseForm()
        Return Activador
    End Function

#Region "ACCIONES REGISTER (CASETA - RUTA)"
    'ACCION PARA OBTENER LAS CASETAS SELECCIONADAS DEL DTG1
    Private Sub DTGCasetaExists_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DTGCasetaExists.CellContentDoubleClick
        If e.ColumnIndex = 0 And CmbCRDestino.Text <> String.Empty And CmbCRVehiculo.Text <> String.Empty Then
            Dim Caseta = Convert.ToString(DTGCasetaExists.Rows(e.RowIndex).Cells(0).Value)
            DTGCasetaSelect.Rows.Add(Caseta)
        Else
            MsgBox("Selecciona un destino y unidad", MsgBoxStyle.Information, "Información | Corporativo LUIN | DTGCasetaExists")
        End If
    End Sub
    'ACCION PARA ELIMINAR LAS CASETAS SELECCIONADAS DEL DTG2
    Private Sub DTGCasetaSelect_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DTGCasetaSelect.CellContentDoubleClick
        If e.RowIndex > -1 And CmbCRDestino.Text <> String.Empty And CmbCRVehiculo.Text <> String.Empty Then
            DTGCasetaSelect.Rows.RemoveAt(e.RowIndex)
        End If
    End Sub
    'AL SELECCIONAR UN DATO OBTIENE EL ID DEL DTG1

    Private Sub CmbCRVehiculo_MouseClick(sender As Object, e As MouseEventArgs) Handles CmbCRVehiculo.MouseClick
        If CmbCRDestino.Text <> "" Then
            NewConfig.ObtenerIDDestinoCR(CmbCRDestino, LIDRutaC)

            Dim dtVehiculos As DataTable = NewConfig.MostrarVehiculosCR(CmbCRVehiculo, LIDRutaC)
            CmbCRVehiculo.DataSource = dtVehiculos
            CmbCRVehiculo.DisplayMember = "Vehiculo"
            CmbCRVehiculo.SelectedIndex = -1
        End If
    End Sub
    'AL SELECCIONAR UN DATO OBTIENE EL ID DEL DTG2
    Private Sub CmbCRVehiculo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCRVehiculo.SelectedIndexChanged
        If CmbCRDestino.Text <> String.Empty Then
            NewConfig.ObtenerIDVehiculoCR(CmbCRVehiculo, LIDVehiculo)
        End If
    End Sub
    'BOTON PARA REGISTRAR LA INFORMACION
    Private Sub BtnNewCasetaRuta_Click(sender As Object, e As EventArgs) Handles BtnNewCasetaRuta.Click
        NewConfig.RegistrarCasetaCR(CmbCRDestino, CmbCRVehiculo, DTGCasetaSelect, Me, LIDRutaC, LIDVehiculo, P_CasetaRuta)
        Activador = 0
        LNewCaseta.Text = ""
    End Sub
    'BOTON PARA CERRAR EL FORMULARIO ACTUAL
    Private Sub BtnCasetaRutaClose_Click(sender As Object, e As EventArgs) Handles BtnCasetaRutaClose.Click
        Activador = 0
        Me.Close()
        Me.Dispose()
        WinPrincipal.Opacity = 1
    End Sub
    Private Sub PBoxWeb_Click(sender As Object, e As EventArgs) Handles PBoxWeb.Click
        System.Diagnostics.Process.Start("http://app.sct.gob.mx/sibuac_internet/ControllerUI?action=cmdEscogeRuta")
    End Sub
    Private Sub PBoxWeb_MouseHover(sender As Object, e As EventArgs) Handles PBoxWeb.MouseHover
        Dim Tip As ToolTip = New ToolTip
        Tip.ToolTipTitle = "¿Buscas las casetas a tu destino?"
        Tip.SetToolTip(PBoxWeb, "Al dar click en este icono seras redireccionado a un sitio web" & Chr(10) & "donde proporcionaras tu origen y destino, mostrandote una lista de casetas.")
    End Sub
    Private Sub PBoxInfoCasetas_MouseHover(sender As Object, e As EventArgs) Handles PBoxInfoCasetas.MouseHover
        Dim Tip As ToolTip = New ToolTip
        Tip.ToolTipTitle = "Como asignar o eliminar casetas."
        Tip.SetToolTip(PBoxInfoCasetas, "Da clic para obtener mas información.")
    End Sub
    Private Sub PBoxInfoCasetas_Click(sender As Object, e As EventArgs) Handles PBoxInfoCasetas.Click
        Dim Texto As String = "• Selecciona el destino y el chofer para poder asignar casetas." & Chr(10) & "• Para seleccionar una caseta da doble clic en el nombre de la caseta a elegir." & Chr(10) & "• Para descartar una caseta da doble clic en el nombre de la caseta elegida"
        ClassToolTip.Show("5", "SeleccionCaseta.png", Texto, 250, 250)
    End Sub
#End Region

#Region "ACCIONES UPDATE"
    Private Sub DTGCasetaExistsUp_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DTGCasetaExistsUp.CellContentDoubleClick
        If e.ColumnIndex = 0 And LNombreDestino.Text <> String.Empty And LNombreVehiculo.Text <> String.Empty Then
            Dim Caseta = Convert.ToString(DTGCasetaExistsUp.Rows(e.RowIndex).Cells(0).Value)
            Dim dt2 As DataTable = New DataTable
            dt2 = DTGCasetaSelectUp.DataSource
            Dim dataRow As DataRow
            dataRow = dt2.NewRow
            dataRow("Casetas") = Caseta
            dt2.Rows.Add(dataRow)
        Else
            MsgBox("Selecciona un destino y unidad", MsgBoxStyle.Information, "Información | Corporativo LUIN | DTGCasetaExists")
        End If
    End Sub
    Private Sub DTGCasetaSelectUp_CellContentDoubleClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DTGCasetaSelectUp.CellContentDoubleClick
        If e.RowIndex > -1 And LNombreDestino.Text <> String.Empty And LNombreVehiculo.Text <> String.Empty Then
            DTGCasetaSelectUp.Rows.RemoveAt(e.RowIndex)
        End If
    End Sub
    Private Sub BtnUpCasetaRuta_Click(sender As Object, e As EventArgs) Handles BtnUpCasetaRuta.Click
        NewConfig.ActualizarCaseta_Ruta(LIDRutaUp, LIDVehiculoUp, DTGCasetaSelectUp, Me, P_UpCasetaRuta)
        Activador = 0
        LUpdateCaseta.Text = ""
    End Sub
    Private Sub BtnCasetaRutaCloseUp_Click(sender As Object, e As EventArgs) Handles BtnCasetaRutaCloseUp.Click
        Activador = 0
        Me.Close()
        Me.Dispose()
        WinPrincipal.Opacity = 1
    End Sub


    Private Sub PBoxWebUp_Click(sender As Object, e As EventArgs) Handles PBoxWebUp.Click
        System.Diagnostics.Process.Start("http://app.sct.gob.mx/sibuac_internet/ControllerUI?action=cmdEscogeRuta")
    End Sub
    Private Sub PBoxWebUp_MouseHover(sender As Object, e As EventArgs) Handles PBoxWebUp.MouseHover
        Dim Tip As ToolTip = New ToolTip
        Tip.ToolTipTitle = "¿Buscas las casetas a tu destino?"
        Tip.SetToolTip(PBoxWebUp, "Al dar click en este icono seras redireccionado a un sitio web" & Chr(10) & "donde proporcionaras tu origen y destino, mostrandote una lista de casetas.")
    End Sub
    Private Sub PBoxInfoCasetasUp_MouseHover(sender As Object, e As EventArgs) Handles PBoxInfoCasetasUp.MouseHover
        Dim Tip As ToolTip = New ToolTip
        Tip.ToolTipTitle = "Como asignar o eliminar casetas."
        Tip.SetToolTip(PBoxInfoCasetasUp, "Da clic para obtener mas información.")
    End Sub
    Private Sub PBoxInfoCasetasUp_Click(sender As Object, e As EventArgs) Handles PBoxInfoCasetasUp.Click
        Dim Texto As String = "• Selecciona el destino y el chofer para poder asignar casetas." & Chr(10) & "• Para seleccionar una caseta da doble clic en el nombre de la caseta a elegir." & Chr(10) & "• Para descartar una caseta da doble clic en el nombre de la caseta elegida"
        ClassToolTip.Show("5", "SeleccionCaseta.png", Texto, 250, 250)
    End Sub


    Private Sub DTGCasetaSelect_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DTGCasetaSelect.EditingControlShowing
        Try
            RemoveHandler e.Control.KeyPress, AddressOf KeyNumber
            AddHandler e.Control.KeyPress, AddressOf KeyNumber

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR | Corporativo LUIN")
        End Try
    End Sub
    Private Sub DTGCasetaSelectUp_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DTGCasetaSelectUp.EditingControlShowing
        Try
            RemoveHandler e.Control.KeyPress, AddressOf KeyNumber
            AddHandler e.Control.KeyPress, AddressOf KeyNumber

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR | Corporativo LUIN")
        End Try
    End Sub

    Sub KeyNumber(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Not (Char.IsControl(e.KeyChar) OrElse Char.IsDigit(e.KeyChar)) AndAlso (Not e.KeyChar = ".") Then
            e.Handled = True
        End If
    End Sub

#End Region
End Class
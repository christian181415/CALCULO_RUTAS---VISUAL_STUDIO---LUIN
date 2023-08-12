Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class WinRegistros
    Dim NewRegistroDta As New ClassRegistrosData
    Private Sub FormActionsConfig_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TimerGlobalRegistros.Start()


        Dim r As Rectangle = My.Computer.Screen.WorkingArea
        Dim Largo = (r.Width / 2) - 100
        Dim Alto = (r.Height / 2) - 130
        Location = New Point(Largo, Alto)

        If LInfoTabla.Text = "Clientes" Then
            Me.Size = New Size(241, 322)
            P_NewCliente.Location = New Point(0, 0)
        ElseIf LInfoTabla.Text = "Choferes" Then
            Me.Size = New Size(241, 322)
            P_NewChofer.Location = New Point(0, 0)
        ElseIf LInfoTabla.Text = "Unidades" Then
            Me.Size = New Size(241, 322)
            P_NewUnidad.Location = New Point(0, 0)
        ElseIf LInfoTabla.Text = "Casetas" Then
            Me.Size = New Size(241, 322)
            P_NewCaseta.Location = New Point(0, 0)
        End If
        If LUpTabla.Text = "Clientes" Then
            Me.Size = New Size(241, 322)
            P_UpdateCliente.Location = New Point(0, 0)
        ElseIf LUpTabla.Text = "Choferes" Then
            Me.Size = New Size(241, 322)
            P_UpdateChofer.Location = New Point(0, 0)
        ElseIf LUpTabla.Text = "Unidades" Then
            Me.Size = New Size(241, 322)
            P_UpdateUnidad.Location = New Point(0, 0)
        ElseIf LUpTabla.Text = "Casetas" Then
            Me.Size = New Size(241, 322)
            P_UpdateCaseta.Location = New Point(0, 0)
        End If


        If LNewRuta.Text = "NuevaRuta" Then
            Me.Size = New Size(241, 322)
            P_NewRuta.Location = New Point(0, 0)
            Dim dtDestinos As DataTable = NewRegistroDta.MostrarDestinos(CmbDestino)
            CmbDestino.DataSource = dtDestinos
            CmbDestino.DisplayMember = "Nombre"
            CmbDestino.SelectedIndex = -1
        End If
        If LUpRuta.Text = "ActualizarRuta" Then
            Me.Size = New Size(241, 322)
            P_UpRuta.Location = New Point(0, 0)
            NewRegistroDta.ObtenerDomicilioRuta(LIDCliente, LDestinoUp)
        End If


    End Sub

#Region "---------------------------------------------------------------ACCIONES REGISTER CATALOGO----------------------------------------------------------------"
#Region "---------------------------------------------------------------REGISTRAR CLIENTE----------------------------------------------------------------"
    'BOTON PARA REGISTRAR UN CLIENTE
    Private Sub BtnNewCliente_Click(sender As Object, e As EventArgs) Handles BtnNewCliente.Click
        NewRegistroDta.RegistrarCliente(TxTNombreC, TxTDomicilioC, Me, P_NewCliente)
        LInfoTabla.Text = ""
        WinPrincipal.Opacity = 1
    End Sub
    'BOTON PARA CERRAR LA VENTANA
    Private Sub BtnClienteClose_Click(sender As Object, e As EventArgs) Handles BtnClienteClose.Click
        WinPrincipal.Opacity = 1
        LInfoTabla.Text = ""
        P_NewCliente.Location = New Point(260, 2)
        Me.Close()
    End Sub
    Private Sub TxTNombreC_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxTNombreC.KeyPress
        If Char.IsLetter(e.KeyChar) Or Char.IsSeparator(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub
#End Region

#Region "---------------------------------------------------------------REGISTRAR CHOFER----------------------------------------------------------------"
    'BOTON PARA REGISTRAR UN CHOFER
    Private Sub BtnNewChofer_Click(sender As Object, e As EventArgs) Handles BtnNewChofer.Click
        NewRegistroDta.RegistrarChofer(TxTNombreCH, TxTTelefonoCH, Me, P_NewChofer)
        LInfoTabla.Text = ""
        WinPrincipal.Opacity = 1
    End Sub
    'BOTON PARA CERRRAR LA VENTANA
    Private Sub BtnChoferClose_Click(sender As Object, e As EventArgs) Handles BtnChoferClose.Click
        LInfoTabla.Text = ""
        P_NewChofer.Location = New Point(517, 2)
        Me.Close()
        WinPrincipal.Opacity = 1
    End Sub
    Private Sub TxTNombreCH_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxTNombreCH.KeyPress
        If Char.IsLetter(e.KeyChar) Or Char.IsSeparator(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub
#End Region

#Region "---------------------------------------------------------------REGISTRAR UNIDAD----------------------------------------------------------------"
    Private Sub BtnNewUnidad_Click(sender As Object, e As EventArgs) Handles BtnNewUnidad.Click
        NewRegistroDta.RegistrarUnidad(TxTVehiculoU, TxTPlacasU, CmbDescripcionU, Me, P_NewUnidad)
        LInfoTabla.Text = ""
        WinPrincipal.Opacity = 1
    End Sub
    Private Sub BtnUnidadClose_Click(sender As Object, e As EventArgs) Handles BtnUnidadClose.Click
        WinPrincipal.Opacity = 1
        LInfoTabla.Text = ""
        P_NewUnidad.Location = New Point(774, 2)
        Me.Close()
        WinPrincipal.Opacity = 1
    End Sub
#End Region

#Region "---------------------------------------------------------------REGISTRAR CASETA----------------------------------------------------------------"
    Private Sub BtnNewCaseta_Click(sender As Object, e As EventArgs) Handles BtnNewCaseta.Click
        NewRegistroDta.RegistrarCaseta(TxTCaseta, Me, P_NewCaseta)
        LInfoTabla.Text = ""
        WinPrincipal.Opacity = 1
    End Sub
    Private Sub BtnCasetaClose_Click(sender As Object, e As EventArgs) Handles BtnCasetaClose.Click
        LInfoTabla.Text = ""
        P_NewCaseta.Location = New Point(1034, 2)
        Me.Close()
        WinPrincipal.Opacity = 1
    End Sub

#End Region
#End Region



#Region "---------------------------------------------------------------ACCIONES UPDATE CATALOGO----------------------------------------------------------------"
#Region "---------------------------------------------------------------ACTUALIZAR CLIENTE----------------------------------------------------------------"
    Private Sub BtnActualizarC_Click(sender As Object, e As EventArgs) Handles BtnActualizarC.Click
        NewRegistroDta.ActualizarCliente(TxTNombreCUp, TxTDomicilioCUp, Me, CBoxStatusCUp, LIdTabla)
        WinPrincipal.Opacity = 1
    End Sub
    Private Sub BtnCCloseUp_Click(sender As Object, e As EventArgs) Handles BtnCCloseUp.Click
        LUpTabla.Text = ""
        P_UpdateCliente.Location = New Point(260, 341)
        Me.Close()
        WinPrincipal.Opacity = 1
    End Sub
    Private Sub TxTNombreCUp_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxTNombreCUp.KeyPress
        If Char.IsLetter(e.KeyChar) Or Char.IsSeparator(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub
#End Region

#Region "---------------------------------------------------------------ACTUALIZAR CHOFER----------------------------------------------------------------"
    Private Sub BtnActualizarCH_Click(sender As Object, e As EventArgs) Handles BtnActualizarCH.Click
        NewRegistroDta.ActualizarChofer(TxTNombreCHUp, TxTTelefonoCHUp, Me, CBoxStatusCHUp, LIdTabla)
        WinPrincipal.Opacity = 1
    End Sub
    Private Sub BtnCHCloseUp_Click(sender As Object, e As EventArgs) Handles BtnCHCloseUp.Click
        LUpTabla.Text = ""
        P_UpdateChofer.Location = New Point(517, 341)
        Me.Close()
        WinPrincipal.Opacity = 1
    End Sub
    Private Sub TxTNombreCHUp_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxTNombreCHUp.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub
#End Region

#Region "---------------------------------------------------------------ACTUALIZAR UNIDAD----------------------------------------------------------------"
    Private Sub BtnActualizarU_Click(sender As Object, e As EventArgs) Handles BtnActualizarU.Click
        NewRegistroDta.ActualizarUnidad(TxTVehiculoUUp, TxTPlacasUUp, Me, CmbDescripcionUUp, LIdTabla)
        WinPrincipal.Opacity = 1
    End Sub
    Private Sub BtnUCloseUp_Click(sender As Object, e As EventArgs) Handles BtnUCloseUp.Click
        LUpTabla.Text = ""
        P_UpdateUnidad.Location = New Point(774, 344)
        Me.Close()
        WinPrincipal.Opacity = 1
    End Sub
#End Region

#Region "---------------------------------------------------------------ACTUALIZAR CASETA----------------------------------------------------------------"
    Private Sub BtnActualizarCaseta_Click(sender As Object, e As EventArgs) Handles BtnActualizarCaseta.Click
        NewRegistroDta.ActualizarCaseta(TxTCasetaUp, Me, LIdTabla)
        WinPrincipal.Opacity = 1
    End Sub
    Private Sub BtnCasetaCloseUp_Click(sender As Object, e As EventArgs) Handles BtnCasetaCloseUp.Click
        LUpTabla.Text = ""
        P_UpdateCaseta.Location = New Point(1034, 344)
        Me.Close()
        WinPrincipal.Opacity = 1
    End Sub
#End Region
#End Region



#Region "---------------------------------------------------------------ACCIONES REGISTER RUTAS----------------------------------------------------------------"
    Private Sub CmbDestino_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbDestino.SelectedIndexChanged
        If CmbDestino.Text <> String.Empty Then
            NewRegistroDta.ObtenerIDCliente(CmbDestino, LIDRuta)
        End If
    End Sub
    Private Sub BtnNewRuta_Click(sender As Object, e As EventArgs) Handles BtnNewRuta.Click
        NewRegistroDta.RegistrarRuta(TxTOrigen, CmbDestino, TxtKilometros, TxtTTrayecto, TxtToka, TxtFegali, LCombustible, LIDRuta, Me, P_NewRuta)
        LNewRuta.Text = ""
        WinPrincipal.Opacity = 1
    End Sub
    Private Sub BtnRutaNClose_Click(sender As Object, e As EventArgs) Handles BtnRutaNClose.Click
        LNewRuta.Text = ""
        P_NewRuta.Location = New Point(260, 681)
        Me.Close()
        WinPrincipal.Opacity = 1
    End Sub
    Private Sub TxtToka_TextChanged(sender As Object, e As EventArgs) Handles TxtToka.TextChanged
        If TxtToka.Text <> String.Empty And TxtFegali.Text <> String.Empty Then
            LCombustible.Text = Convert.ToDouble(TxtToka.Text) + Convert.ToDouble(TxtFegali.Text)
        End If
    End Sub
    Private Sub TxtFegali_TextChanged(sender As Object, e As EventArgs) Handles TxtFegali.TextChanged
        If TxtToka.Text <> String.Empty And TxtFegali.Text <> String.Empty Then
            LCombustible.Text = Convert.ToDouble(TxtToka.Text) + Convert.ToDouble(TxtFegali.Text)
        End If
    End Sub
    Private Sub TxtToka_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtToka.KeyPress
        If Not (Char.IsControl(e.KeyChar) OrElse Char.IsDigit(e.KeyChar)) _
            AndAlso (Not e.KeyChar = "." Or TxtToka.Text.Contains(".")) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxtFegali_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtFegali.KeyPress
        If Not (Char.IsControl(e.KeyChar) OrElse Char.IsDigit(e.KeyChar)) _
            AndAlso (Not e.KeyChar = "." Or TxtFegali.Text.Contains(".")) Then
            e.Handled = True
        End If
    End Sub
#End Region

#Region "---------------------------------------------------------------ACCIONES ACTUALIZAR RUTAS----------------------------------------------------------------"
    Private Sub BtnActualizarRuta_Click(sender As Object, e As EventArgs) Handles BtnActualizarRuta.Click
        If LDestinoUp.Text <> String.Empty Then
            NewRegistroDta.ObtenerIDClienteUp(LIDClienteUp, LDestinoUp)
            NewRegistroDta.ActualizarRuta(TxTOrigenUp, LDestinoUp, TxTKilometrosUp, TxTTTrayectoUp, TxTTOKAUp, TxTFegaliUp, LCombustibleUp, LIDClienteUp, Me, LUpIDRuta)
            LNewRuta.Text = ""
            P_NewRuta.Location = New Point(260, 681)
            WinPrincipal.Opacity = 1
        End If
    End Sub
    Private Sub BtnRutaCloseUp_Click(sender As Object, e As EventArgs) Handles BtnRutaCloseUp.Click
        LUpRuta.Text = ""
        P_UpRuta.Location = New Point(517, 681)
        Me.Close()
        WinPrincipal.Opacity = 1
    End Sub
    Private Sub TxTKilometrosUp_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxTKilometrosUp.KeyPress
        If Not (Char.IsControl(e.KeyChar) OrElse Char.IsDigit(e.KeyChar)) _
            AndAlso (Not e.KeyChar = "." Or TxTKilometrosUp.Text.Contains(".")) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxTTOKAUp_TextChanged(sender As Object, e As EventArgs) Handles TxTTOKAUp.TextChanged
        If TxTTOKAUp.Text <> String.Empty And TxTFegaliUp.Text <> String.Empty Then
            LCombustibleUp.Text = Convert.ToDouble(TxTTOKAUp.Text) + Convert.ToDouble(TxTFegaliUp.Text)
        End If
    End Sub
    Private Sub TxTFegaliUp_TextChanged(sender As Object, e As EventArgs) Handles TxTFegaliUp.TextChanged
        If TxTTOKAUp.Text <> String.Empty And TxTFegaliUp.Text <> String.Empty Then
            LCombustibleUp.Text = Convert.ToDouble(TxTTOKAUp.Text) + Convert.ToDouble(TxTFegaliUp.Text)
        End If
    End Sub

    Private Sub TxTTOKAUp_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxTTOKAUp.KeyPress
        If Not (Char.IsControl(e.KeyChar) OrElse Char.IsDigit(e.KeyChar)) _
            AndAlso (Not e.KeyChar = "." Or TxTTOKAUp.Text.Contains(".")) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxTFegaliUp_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxTFegaliUp.KeyPress
        If Not (Char.IsControl(e.KeyChar) OrElse Char.IsDigit(e.KeyChar)) _
            AndAlso (Not e.KeyChar = "." Or TxTFegaliUp.Text.Contains(".")) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TxtKilometros_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtKilometros.KeyPress
        If Not (Char.IsControl(e.KeyChar) OrElse Char.IsDigit(e.KeyChar)) _
            AndAlso (Not e.KeyChar = "." Or TxtToka.Text.Contains(".")) Then
            e.Handled = True
        End If
    End Sub



    Private Sub PBoxToka_Click(sender As Object, e As EventArgs) Handles PBoxToka.Click
        Dim Texto As String = "• Tarjeta para recargar combustible fuera de la zona corporativa."
        ClassToolTip.Show("5", "TOKA.png", Texto, 205, 290)
    End Sub
    Private Sub PBoxFegali_Click(sender As Object, e As EventArgs) Handles PBoxFegali.Click
        Dim Texto As String = "• Proveedor de gasolina en la zona corporativa."
        ClassToolTip.Show("5", "FEGALI.png", Texto, 145, 290)
    End Sub




    Private Sub BtnGenerarHPDF_Click(sender As Object, e As EventArgs) Handles BtnGenerarHPDF.Click
        'MsgBox(CalendarLastPDF.SelectionStart & " " & LLastHoras.SelectedItem)
        NewRegistroDta.GetLastPDF(CalendarLastPDF, LLastHoras)
    End Sub
    Private Sub BtnLastPDFClose_Click(sender As Object, e As EventArgs) Handles BtnLastPDFClose.Click
        LLastPDF.Text = ""
        PLastPDF.Location = New Point(774, 681)
        Me.Close()
        WinPrincipal.Opacity = 1
    End Sub

    Private Sub TimerGlobalRegistros_Tick(sender As Object, e As EventArgs) Handles TimerGlobalRegistros.Tick
        If LLastPDF.Text = "ConsultarPDF" Then
            Me.Size = New Size(345, 322)
            NewRegistroDta.ShowDatePDF(CalendarLastPDF)
            CalendarLastPDF.SelectionStart = DateValue("10/12/2001")
            CalendarLastPDF.SelectionStart = Date.Now
            LLastPDF.Text = ""
            PLastPDF.Location = New Point(0, 0)
        End If
    End Sub

    Private Sub CalendarLastPDF_DateChanged(sender As Object, e As DateRangeEventArgs) Handles CalendarLastPDF.DateChanged

    End Sub

    Private Sub CalendarLastPDF_DateSelected(sender As Object, e As DateRangeEventArgs) Handles CalendarLastPDF.DateSelected
        NewRegistroDta.ShowHoursPDF(CalendarLastPDF, LLastHoras)
    End Sub
#End Region
End Class
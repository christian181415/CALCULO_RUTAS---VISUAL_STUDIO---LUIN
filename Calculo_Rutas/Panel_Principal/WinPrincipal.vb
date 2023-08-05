Public Class WinPrincipal
    'VARIABLES DE ACCESO A LA INFORMACION DE OTRAS CLASES
    Dim NewConsultaP As New ClassPrincipalData
    Dim NewConsultaC As New ClassPrincipalConfig
    Dim NewGIF As GIF = Nothing


#Region "---------------------------------------------------------------LOAD PRINCIPAL (RUTAS - SELECCIONAR DATOS)---------------------------------------------------------------"
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'CONFIGURA EL SISTEMA EN LA CARGA INICIAL
        Me.Size = New Size(556, 599)
        NewGIF = New GIF(My.Resources.IconAlert)
        NewGIF.ReverseAtEnd = False
        TimerErrorAlert.Enabled = True
        'VALIDA LA CONEXION A BASE DE DATOS
        NewConsultaP.ValidarConexionP(PBoxAlertIcon, P_Chofer, P_GastosDestino, P_Cofiguracion, PBoxConfiguracion, DTP_Fecha, LFecha, TimerErrorAlert)
        'MUESTRA LA INFORMACION
        Dim dtRutas As DataTable = NewConsultaC.MostrarRutas()
        CMB_Rutas.DataSource = dtRutas
        CMB_Rutas.DisplayMember = "Nombre"
        CMB_Rutas.SelectedIndex = -1
        Dim dtRutasPCasetas As DataTable = NewConsultaC.MostrarRutasPCasetas()
        CmbRutaImporte.DataSource = dtRutasPCasetas
        CmbRutaImporte.DisplayMember = "Nombre"
        CmbRutaImporte.SelectedIndex = -1
        'GUARDAR CONFIG DE COSTO COMBUSTIBLE
        TxTCostoCombustible.Text = My.Settings.CostoCombustible
        'TIMERS START
        TimerAcciones.Start()
        TimerInfoDB.Start()
        'LIMPIA LOS COMPONENTES PARA INICIAR EN BLANCO
        NewConsultaP.LimpiarInfo(CMB_Cliente, CMB_Chofer, CMB_Vehiculo, L_Ruta_Destino, E_Alimentos, TxTViaticos, TXT_Notas, L_Desgloce_Casetas, Total_Casetas, Total_Combustible, CBox_Alimentos)
    End Sub

    Private Sub PBoxAlertIcon_MouseHover_1(sender As Object, e As EventArgs) Handles PBoxAlertIcon.MouseHover
        'TOOLTIP PARA EL COMPONENTE PICTURE_BOX DE ALERTA
        Dim Tip As ToolTip = New ToolTip
        Tip.SetToolTip(PBoxAlertIcon, "Error de conexion a base de datos")
    End Sub
    Private Sub TimerAcciones_Tick(sender As Object, e As EventArgs) Handles TimerAcciones.Tick
        'TIMER PARA SABER SI CAMBIO ALGO EN LOS COMBO_BOX
        If CMB_Cliente.Text <> String.Empty And CMB_Chofer.Text <> String.Empty And CMB_Vehiculo.Text <> String.Empty Then
            P_Destino.Enabled = True
            P_InfoDestino.Enabled = True
            P_GastosDestino.Enabled = True
            P_OpcionesExtra.Enabled = True
        Else
            P_Destino.Enabled = False
            P_InfoDestino.Enabled = False
            P_GastosDestino.Enabled = False
            P_OpcionesExtra.Enabled = False
            L_Desgloce_Casetas.SelectionMode = SelectionMode.One
            L_Desgloce_Casetas.DataSource = Nothing
            L_Desgloce_Casetas.Items.Clear()
            L_Desgloce_Casetas.SelectionMode = SelectionMode.None
        End If
        If E_Alimentos.Text <> String.Empty And TxTViaticos.Text <> String.Empty Then
            LEfectivoTotal.Text = Convert.ToDouble(E_Alimentos.Text) + Convert.ToDouble(TxTViaticos.Text)
        Else
            LEfectivoTotal.Text = "0.00"
        End If
    End Sub
    Private Sub TimerErrorAlert_Tick(sender As Object, e As EventArgs) Handles TimerErrorAlert.Tick
        PBoxAlertIcon.BackgroundImage = NewGIF.GetNextFrame()
    End Sub
#End Region




#Region "----------------------------------------------------------PANEL PRINCIPAL (RUTAS - MOSTRAR INFO DB)-----------------------------------------------------------"
    'FUCION VALIDA SE SE HABILITA OPCION ALIMENTOS
    Private Sub CBox_Alimentos_CheckedChanged(sender As Object, e As EventArgs) Handles CBox_Alimentos.CheckedChanged
        If CBox_Alimentos.Checked = True Then
            E_Alimentos.Enabled = True
            TxTViaticos.Enabled = True
        Else
            If CBox_Alimentos.Checked = False Then
                E_Alimentos.Text = "0.00"
                E_Alimentos.Enabled = False
                TxTViaticos.Text = "0.00"
                TxTViaticos.Enabled = False
            End If
        End If
    End Sub
    Private Sub E_Alimentos_KeyPress(sender As Object, e As KeyPressEventArgs) Handles E_Alimentos.KeyPress
        If Not (Char.IsControl(e.KeyChar) OrElse Char.IsDigit(e.KeyChar)) _
            AndAlso (Not e.KeyChar = "." Or TxTTOKA.Text.Contains(".")) Then
            e.Handled = True
        End If
    End Sub
    'FUCION LIMPIA CAMPO ALIMENTOS AL HACER CLICK
    Private Sub E_Alimentos_Leave(sender As Object, e As EventArgs) Handles E_Alimentos.Leave
        If E_Alimentos.Text = String.Empty Then
            E_Alimentos.Text = "0.00"
        End If
    End Sub
    Private Sub E_Alimentos_MouseClick(sender As Object, e As MouseEventArgs) Handles E_Alimentos.MouseClick
        E_Alimentos.Text = ""
    End Sub
    Private Sub TxTViaticos_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxTViaticos.KeyPress
        If Not (Char.IsControl(e.KeyChar) OrElse Char.IsDigit(e.KeyChar)) _
            AndAlso (Not e.KeyChar = "." Or TxTTOKA.Text.Contains(".")) Then
            e.Handled = True
        End If
    End Sub
    Private Sub TxTViaticos_Leave(sender As Object, e As EventArgs) Handles TxTViaticos.Leave
        If TxTViaticos.Text = String.Empty Then
            TxTViaticos.Text = "0.00"
        End If
    End Sub
    Private Sub TxTViaticos_MouseClick(sender As Object, e As MouseEventArgs) Handles TxTViaticos.MouseClick
        TxTViaticos.Text = ""
    End Sub
    'ACCION DE MOSTRAR LOS DATOS DB RELACIONADOS AL COMBO_BOX
    Private Sub CMB_Cliente_MouseClick(sender As Object, e As MouseEventArgs) Handles CMB_Cliente.MouseClick
        Dim dtCliente As DataTable = NewConsultaP.MostrarClientes(LCliente, CMB_Cliente)
        CMB_Cliente.DataSource = dtCliente
        CMB_Cliente.DisplayMember = "Nombre"
    End Sub
    Private Sub CMB_Chofer_MouseClick(sender As Object, e As MouseEventArgs) Handles CMB_Chofer.MouseClick
        'MOSTRAR CHOFERES
        Dim dtChofer As DataTable = NewConsultaP.MostrarChofer(LChofer, CMB_Chofer)
        CMB_Chofer.DataSource = dtChofer
        CMB_Chofer.DisplayMember = "Nombre"
    End Sub
    Private Sub CMB_Cliente_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CMB_Cliente.SelectedIndexChanged
        If CMB_Cliente.Text <> String.Empty And CMB_Chofer.Text <> String.Empty Then
            Dim dtUnidades As DataTable = NewConsultaP.MostrarUnidades(LUnidad, CMB_Vehiculo, CMB_Cliente)
            CMB_Vehiculo.DataSource = dtUnidades
            CMB_Vehiculo.DisplayMember = "Vehiculo"

            NewConsultaP.MostrarRutas(LRuta, L_Ruta_Destino, CMB_Cliente)
            NewConsultaP.MostrarDatosBusqueda(Total_Casetas, Total_Combustible, L_Ruta_Destino, CMB_Vehiculo, L_Desgloce_Casetas, TxTCostoCombustible, LKilometrosPDF, LTiempoTrayectoPDF)
        End If
    End Sub
    Private Sub CMB_Chofer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CMB_Chofer.SelectedIndexChanged
        If CMB_Cliente.Text <> String.Empty And CMB_Chofer.Text <> String.Empty Then
            Dim dtUnidades As DataTable = NewConsultaP.MostrarUnidades(LUnidad, CMB_Vehiculo, CMB_Cliente)
            CMB_Vehiculo.DataSource = dtUnidades
            CMB_Vehiculo.DisplayMember = "Vehiculo"

            NewConsultaP.MostrarRutas(LRuta, L_Ruta_Destino, CMB_Cliente)
            NewConsultaP.MostrarDatosBusqueda(Total_Casetas, Total_Combustible, L_Ruta_Destino, CMB_Vehiculo, L_Desgloce_Casetas, TxTCostoCombustible, LKilometrosPDF, LTiempoTrayectoPDF)
        End If
    End Sub
    Private Sub CMB_Vehiculo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CMB_Vehiculo.SelectedIndexChanged
        NewConsultaP.MostrarDatosBusqueda(Total_Casetas, Total_Combustible, L_Ruta_Destino, CMB_Vehiculo, L_Desgloce_Casetas, TxTCostoCombustible, LKilometrosPDF, LTiempoTrayectoPDF)
    End Sub
    'FUNCION DEL COMBUSTIBLE
    Private Sub TxTCostoCombustible_TextChanged(sender As Object, e As EventArgs) Handles TxTCostoCombustible.TextChanged
        If L_Ruta_Destino.Text <> String.Empty And CMB_Vehiculo.Text <> String.Empty And TxTCostoCombustible.Text <> String.Empty Then
            NewConsultaP.MostrarCombustible(CMB_Vehiculo, L_Ruta_Destino, Total_Combustible, TxTCostoCombustible)
        End If
    End Sub
    Private Sub TxTCostoCombustible_MouseClick(sender As Object, e As MouseEventArgs) Handles TxTCostoCombustible.MouseClick
        TxTCostoCombustible.Text = ""
    End Sub
    Private Sub TxTCostoCombustible_Leave(sender As Object, e As EventArgs) Handles TxTCostoCombustible.Leave
        If TxTCostoCombustible.Text = String.Empty Then
            Total_Combustible.Text = "$ 0.00"
            TxTCostoCombustible.Text = My.Settings.CostoCombustible
            My.Settings.Save()
        End If
    End Sub
    Private Sub TxTCostoCombustible_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxTCostoCombustible.KeyPress
        If Not (Char.IsControl(e.KeyChar) OrElse Char.IsDigit(e.KeyChar)) _
            AndAlso (Not e.KeyChar = "." Or TxTTOKA.Text.Contains(".")) Then
            e.Handled = True
        End If
    End Sub
#End Region




#Region "--------------------------------------------------------PANEL PRINCIPAL (PICTURE_BOX - CONFIGURACION)----------------------------------------------------------"
    'AL DAR CLICK ABRE EL PANEL DE CONFIGURACION
    Private Sub PBoxConfiguracion_Click(sender As Object, e As EventArgs) Handles PBoxConfiguracion.Click
        Me.Size = New Size(751, 599)
        P_Chofer.Enabled = False
        P_Destino.Enabled = False
        P_GastosDestino.Enabled = False
        P_OpcionesExtra.Enabled = False
        P_InfoDestino.Enabled = False
        DTP_Fecha.Enabled = False
        LFecha.Enabled = False

        TimerAcciones.Stop()
    End Sub
    Private Sub PBoxConfiguracion_MouseHover(sender As Object, e As EventArgs) Handles PBoxConfiguracion.MouseHover
        Dim Tip As ToolTip = New ToolTip
        Dim ImgActive As System.Drawing.Bitmap = Bitmap.FromFile(System.AppDomain.CurrentDomain.BaseDirectory() + "\Assets\IMG\BtnConfig_Active.png")
        Tip.SetToolTip(PBoxConfiguracion, "Opciones de configuración")
        PBoxConfiguracion.BackgroundImage = ImgActive
    End Sub
    Private Sub PBoxConfiguracion_MouseLeave(sender As Object, e As EventArgs) Handles PBoxConfiguracion.MouseLeave
        Dim ImgDisable As System.Drawing.Bitmap = Bitmap.FromFile(System.AppDomain.CurrentDomain.BaseDirectory() + "\Assets\IMG\BtnConfig_Disable.png")
        PBoxConfiguracion.BackgroundImage = ImgDisable
    End Sub
#End Region




#Region "-------------------------------------------------------PANEL PRINCIPAL (OPCIONES EXTRA)-----------------------------------------------------"
    'BOTON PARA LIMPIAR LOS COMPONENTES 
    Private Sub BTN_Limpiar_Click(sender As Object, e As EventArgs) Handles BTN_Limpiar.Click
        NewConsultaP.LimpiarInfo(CMB_Cliente, CMB_Chofer, CMB_Vehiculo, L_Ruta_Destino, E_Alimentos, TxTViaticos, TXT_Notas, L_Desgloce_Casetas, Total_Casetas, Total_Combustible, CBox_Alimentos)
    End Sub
    'BOTON PARA SALIR DEL SISTEMA
    Private Sub BTN_Salir_Click(sender As Object, e As EventArgs) Handles BTN_Salir.Click
        Application.Exit()
        My.Settings.CostoCombustible = TxTCostoCombustible.Text
    End Sub
    'BOTON PARA GENERAR UN PDF
    Private Sub PBoxPDF_Click(sender As Object, e As EventArgs) Handles PBoxPDF.Click
        Try
            If CMB_Cliente.Text <> String.Empty And CMB_Chofer.Text <> String.Empty And CMB_Vehiculo.Text <> String.Empty And L_Ruta_Destino.Text <> "Corporativo LUIN" Then
                NewConsultaP.ObtenerIDsBitcaora(L_Ruta_Destino, LRutaIDBit, LOrigenBit, LLitroCombustibleBit, CMB_Chofer, LChoferIDBit)
                If LRutaIDBit.Text <> "LRutaIDBit" And LChoferIDBit.Text <> "LChoferIDBit" And LOrigenBit.Text <> "LOrigenBit" And L_Ruta_Destino.Text <> "Corporativo LUIN" And LKilometrosPDF.Text <> "LKilometrosPDF" And LTiempoTrayectoPDF.Text <> "LTiempoTrayectoPDF" And TxTViaticos.Text <> String.Empty And LLitroCombustibleBit.Text <> "LLitroCombustibleBit" And TxTCostoCombustible.Text <> String.Empty And Total_Casetas.Text <> String.Empty And E_Alimentos.Text <> String.Empty Then
                    Dim NuevoPDF As PDF = New PDF()
                    NewConsultaP.RegistrarBitacora(DTP_Fecha.Text, LRutaIDBit, LChoferIDBit, LOrigenBit, L_Ruta_Destino, LKilometrosPDF, LTiempoTrayectoPDF, TxTViaticos, LLitroCombustibleBit, TxTCostoCombustible, Total_Casetas, TXT_Notas, E_Alimentos, CMB_Chofer)
                    NuevoPDF.GenerarPDF(SFDialogPDF, DTP_Fecha.Text, CMB_Cliente.Text, CMB_Vehiculo.Text, L_Ruta_Destino.Text, TXT_Notas.Text, Total_Combustible.Text, LEfectivoTotal.Text, Total_Casetas.Text, L_Desgloce_Casetas, LKilometrosPDF.Text, LTiempoTrayectoPDF.Text)
                    NewConsultaP.LimpiarInfo(CMB_Cliente, CMB_Chofer, CMB_Vehiculo, L_Ruta_Destino, E_Alimentos, TxTViaticos, TXT_Notas, L_Desgloce_Casetas, Total_Casetas, Total_Combustible, CBox_Alimentos)
                Else
                    MsgBox("Error al generar PDF.", MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | Generar PDF")
                End If
            Else
                MsgBox("Faltan campos por rellenar.", MsgBoxStyle.Information, "Información | Corporativo LUIN | Generar PDF")
            End If
        Catch ex As Exception
            MsgBox("Error al generar PDF.")
        End Try
    End Sub
    Private Sub PBoxPDF_MouseHover(sender As Object, e As EventArgs) Handles PBoxPDF.MouseHover
        Dim Tip As ToolTip = New ToolTip
        Tip.SetToolTip(PBoxPDF, "Generar PDF")
        Dim ImgActive As System.Drawing.Bitmap = Bitmap.FromFile(System.AppDomain.CurrentDomain.BaseDirectory() + "\Assets\IMG\BtnPDF_Active.png")
        PBoxPDF.BackgroundImage = ImgActive
    End Sub
    Private Sub PBoxPDF_MouseLeave(sender As Object, e As EventArgs) Handles PBoxPDF.MouseLeave
        Dim ImgDisable As System.Drawing.Bitmap = Bitmap.FromFile(System.AppDomain.CurrentDomain.BaseDirectory() + "\Assets\IMG\BtnPDF_Disable.png")
        PBoxPDF.BackgroundImage = ImgDisable
    End Sub
    Private Sub PBoxCasetas_MouseHover(sender As Object, e As EventArgs) Handles PBoxCasetas.MouseHover
        Dim Tip As ToolTip = New ToolTip
        Tip.ToolTipTitle = "Presupuesto para los gastos de casetas"
        Tip.SetToolTip(PBoxCasetas, "El presupuesto toma en cuenta la ida y regreso del viaje.")
    End Sub
    Private Sub PBoxEfectivo_MouseHover(sender As Object, e As EventArgs) Handles PBoxEfectivo.MouseHover
        Dim Tip As ToolTip = New ToolTip
        Tip.ToolTipTitle = "Efectivo para el viaje           *Campo rellenado con 'Efectivo del viaje'"
        Tip.SetToolTip(PBoxEfectivo, "Puede ser usado para situaciones de emergencia, material necesario, etc...")
    End Sub
    Private Sub PBoxCombustible_MouseHover(sender As Object, e As EventArgs) Handles PBoxCombustible.MouseHover
        Dim Tip As ToolTip = New ToolTip
        Tip.ToolTipTitle = "Proyeccción para los gastos de combustible"
        Tip.SetToolTip(PBoxCombustible, "Esta proyeccion es una estimacion sobre los gastos del combustible.")
    End Sub
    Private Sub PBoxPrecioCombustible_MouseHover(sender As Object, e As EventArgs) Handles PBoxPrecioCombustible.MouseHover
        Dim Tip As ToolTip = New ToolTip
        Tip.ToolTipTitle = "Costo de combustible por litro              *Campo rellenado manualmente"
        Tip.SetToolTip(PBoxPrecioCombustible, "Este campo puede ser modificado cuando cambie el precio del combustible.")
    End Sub
#End Region




#Region "-------------------------------------------------------PANEL CONFIGURACION-----------------------------------------------------"
    'FUCION CIERRA OPCIONES DE CONFIGURACION
    Private Sub BTN_Cerrar_Click(sender As Object, e As EventArgs) Handles BTN_Cerrar.Click
        Me.Size = New Size(556, 599)
        P_Chofer.Enabled = True
        P_Destino.Enabled = True
        P_GastosDestino.Enabled = True
        P_InfoDestino.Enabled = True
        P_OpcionesExtra.Enabled = True
        DTP_Fecha.Enabled = True
        LFecha.Enabled = True

        TimerAcciones.Start()
    End Sub
#Region "-------------------------------------------------------FUNCIONES PARA CREAR, ACTUALIZAR Y ELIMINAR CATALOGOS-----------------------------------------------------"
    'FUCION ASIGNA NOMBRE A LISTA CATALOGOS
    Private Sub CMB_Directorio_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CMB_Directorio.SelectedIndexChanged
        NewConsultaC.MostrarActualizarCatalogo(L_Directorio, CMB_Directorio, Lista_Catalago)
    End Sub
    'FUNCION CREA UN NUEVO CATALOGO
    Private Sub BtnNewCatalogo_Click(sender As Object, e As EventArgs) Handles BtnNewCatalogo.Click
        If CMB_Directorio.Text <> String.Empty Then
            NewConsultaC.NewCatalogo(CMB_Directorio, Me)
            NewConsultaC.MostrarActualizarCatalogo(L_Directorio, CMB_Directorio, Lista_Catalago)
        End If
    End Sub
    Private Sub BtnUpdateCatalogo_Click(sender As Object, e As EventArgs) Handles BtnUpdateCatalogo.Click
        If CMB_Directorio.Text <> String.Empty Then
            Lista_Catalago.Enabled = True
            NewConsultaC.ObtenerCatalogo(CMB_Directorio, Lista_Catalago, Me)
            NewConsultaC.MostrarActualizarCatalogo(L_Directorio, CMB_Directorio, Lista_Catalago)
        Else
            Lista_Catalago.Enabled = False
        End If
    End Sub
    Private Sub TimerInfoDB_Tick(sender As Object, e As EventArgs) Handles TimerInfoDB.Tick
        If CMB_Directorio.Text <> String.Empty Then
            BtnNewCatalogo.Enabled = True
        Else
            BtnNewCatalogo.Enabled = False
        End If

        If CMB_Directorio.Text <> String.Empty And Lista_Catalago.SelectedIndex <> -1 Then
            BtnUpdateCatalogo.Enabled = True
            BtnEliminarCatalogo.Enabled = True
        Else
            BtnUpdateCatalogo.Enabled = False
            BtnEliminarCatalogo.Enabled = False
        End If


        If CMB_Rutas.Text <> String.Empty Then
            BtnEliminarRuta.Enabled = True
            BtnUpdateRuta.Enabled = True
        Else
            BtnEliminarRuta.Enabled = False
            BtnUpdateRuta.Enabled = False
        End If

        If LRutaID.Text <> "RutaID" And LVehiculoID.Text <> "VehiculoID" And CmbRutaImporte.Text <> String.Empty And CmbVehiculoImporte.Text <> String.Empty Then
            BtnEliminarCasetaRuta.Enabled = True
            BtnUpdateImporteC.Enabled = True
        Else
            BtnEliminarCasetaRuta.Enabled = False
            BtnUpdateImporteC.Enabled = False
        End If
    End Sub

#Region "Eliminar Catalogo"
    Private Sub BtnEliminarCatalogo_Click(sender As Object, e As EventArgs) Handles BtnEliminarCatalogo.Click
        If CMB_Directorio.Text <> String.Empty And Lista_Catalago.SelectedIndex <> -1 Then
            NewConsultaC.EliminarCatalogo(CMB_Directorio, Lista_Catalago)
            NewConsultaC.MostrarActualizarCatalogo(L_Directorio, CMB_Directorio, Lista_Catalago)
        End If
    End Sub
#End Region
#End Region


#Region "-------------------------------------------------------FUNCIONES PARA CREAR, ACTUALIZAR Y ELIMINAR LAS RUTAS-----------------------------------------------------"
    Private Sub CMB_Rutas_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CMB_Rutas.SelectedIndexChanged
        If CMB_Rutas.Text <> String.Empty Then
            NewConsultaC.MostrarDatosRuta(CMB_Rutas, TxTOrigen, TxTDestino, TxTKilometros, TxTTiempoTrayecto, TxTTOKA, TxTFegali, TxTCombustible, LIDRuta)
        End If
    End Sub
    Private Sub CMB_Rutas_MouseClick(sender As Object, e As MouseEventArgs) Handles CMB_Rutas.MouseClick
        Dim dtRutasC As DataTable = NewConsultaC.MostrarRutas()
        CMB_Rutas.DataSource = dtRutasC
        CMB_Rutas.DisplayMember = "Nombre"
    End Sub

    Private Sub BtnNewRuta_Click(sender As Object, e As EventArgs) Handles BtnNewRuta.Click
        NewConsultaC.NewRuta(Me)
        Dim dtRutasC As DataTable = NewConsultaC.MostrarRutas()
        CMB_Rutas.DataSource = dtRutasC
        CMB_Rutas.DisplayMember = "Nombre"
    End Sub

    Private Sub BtnUpdateRuta_Click(sender As Object, e As EventArgs) Handles BtnUpdateRuta.Click
        If CMB_Rutas.Text <> String.Empty Then
            BtnUpdateRuta.Enabled = True
            NewConsultaC.ObtenerInfoRuta(LIDRuta, Me)
            Dim dtRutasC As DataTable = NewConsultaC.MostrarRutas()
            CMB_Rutas.DataSource = dtRutasC
            CMB_Rutas.DisplayMember = "Nombre"
        Else
            BtnUpdateRuta.Enabled = False
        End If
    End Sub
    Private Sub BtnEliminarRuta_Click(sender As Object, e As EventArgs) Handles BtnEliminarRuta.Click
        NewConsultaC.EliminarCasetasRuta(LIDRuta)
        NewConsultaC.EliminarRuta(LIDRuta)
        Dim dtRutasC As DataTable = NewConsultaC.MostrarRutas()
        CMB_Rutas.DataSource = dtRutasC
        CMB_Rutas.DisplayMember = "Nombre"
    End Sub
#End Region




#Region "-------------------------------------------------------FUNCIONES PARA CREAR Y ACTUALIZAR LA UNION DE CASETAS-----------------------------------------------------"
    Private Sub CmbRutaImporte_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbRutaImporte.SelectedIndexChanged
        Try
            If CmbRutaImporte.Text <> String.Empty Then
                NewConsultaC.ObtenerIDDomicilioIDVehiculo(CmbRutaImporte, CmbVehiculoImporte, LRutaID, LVehiculoID)
            End If
            If LRutaID.Text <> "RutaID" And LRutaID.Text <> String.Empty Then
                Dim dtUnidadesC As DataTable = NewConsultaC.MostrarUnidadesC(LRutaID)
                CmbVehiculoImporte.DataSource = Nothing
                CmbVehiculoImporte.DataSource = dtUnidadesC
                CmbVehiculoImporte.DisplayMember = "Vehiculo"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error | Corporativo LUIN | CMBRutaImporte")
        End Try
    End Sub
    Private Sub CmbRutaImporte_MouseClick(sender As Object, e As MouseEventArgs) Handles CmbRutaImporte.MouseClick
        Dim dtRutas As DataTable = NewConsultaC.MostrarRutasPCasetas()
        CmbRutaImporte.DataSource = dtRutas
        CmbRutaImporte.DisplayMember = "Nombre"
    End Sub

    Private Sub CmbVehiculoImporte_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbVehiculoImporte.SelectedIndexChanged
        If CmbVehiculoImporte.Text <> String.Empty Then
            NewConsultaC.ObtenerIDDomicilioIDVehiculo(CmbRutaImporte, CmbVehiculoImporte, LRutaID, LVehiculoID)
        End If
        If CmbRutaImporte.Text <> String.Empty And CmbVehiculoImporte.Text <> String.Empty Then
            Dim dtCasetas As DataTable = NewConsultaC.MostrarCasetas(CmbRutaImporte, CmbVehiculoImporte)
            Lista_Casetas.DataSource = dtCasetas
            Lista_Casetas.DisplayMember = "CasetaImporte"
        End If
    End Sub

    Private Sub BtnNewImporteC_Click(sender As Object, e As EventArgs) Handles BtnNewImporteC.Click
        NewConsultaC.NewCasetaRuta(Me)
        Dim dtRutas As DataTable = NewConsultaC.MostrarRutasPCasetas()
        CmbRutaImporte.DataSource = dtRutas
        CmbRutaImporte.DisplayMember = "Nombre"
        CmbRutaImporte.SelectedIndex = -1
        CmbVehiculoImporte.SelectedIndex = -1
        LRutaID.Text = "IDRuta"
        LVehiculoID.Text = "VehiculoID"
        Lista_Casetas.SelectionMode = SelectionMode.One
        Lista_Casetas.DataSource = Nothing
        Lista_Casetas.Items.Clear()
        Lista_Casetas.SelectionMode = SelectionMode.None
    End Sub
    Private Sub BtnUpdateImporteC_Click(sender As Object, e As EventArgs) Handles BtnUpdateImporteC.Click
        If LRutaID.Text <> String.Empty And LVehiculoID.Text <> String.Empty Then
            NewConsultaC.ActualizarCasetaRuta(LRutaID, LVehiculoID, Me)
            Dim dtRutas As DataTable = NewConsultaC.MostrarRutasPCasetas()
            CmbRutaImporte.DataSource = dtRutas
            CmbRutaImporte.DisplayMember = "Nombre"
            CmbRutaImporte.SelectedIndex = -1
            CmbVehiculoImporte.SelectedIndex = -1
            LRutaID.Text = "IDRuta"
            LVehiculoID.Text = "VehiculoID"
            Lista_Casetas.SelectionMode = SelectionMode.One
            Lista_Casetas.DataSource = Nothing
            Lista_Casetas.Items.Clear()
            Lista_Casetas.SelectionMode = SelectionMode.None
        End If
    End Sub
    Private Sub BtnEliminarCasetaRuta_Click(sender As Object, e As EventArgs) Handles BtnEliminarCasetaRuta.Click
        If LRutaID.Text <> String.Empty And LVehiculoID.Text <> String.Empty Then
            NewConsultaC.EliminarCasetaRuta(LRutaID, LVehiculoID)
            Dim dtRutas As DataTable = NewConsultaC.MostrarRutasPCasetas()
            CmbRutaImporte.DataSource = dtRutas
            CmbRutaImporte.DisplayMember = "Nombre"
            CmbRutaImporte.SelectedIndex = -1
            CmbVehiculoImporte.SelectedIndex = -1
            LRutaID.Text = "IDRuta"
            LVehiculoID.Text = "VehiculoID"
            Lista_Casetas.SelectionMode = SelectionMode.One
            Lista_Casetas.DataSource = Nothing
            Lista_Casetas.Items.Clear()
            Lista_Casetas.SelectionMode = SelectionMode.None
        End If
    End Sub

    Private Sub PBoxLastPDF_Click(sender As Object, e As EventArgs) Handles PBoxLastPDF.Click
        NewConsultaC.LastPDF("ConsultarPDF", Me)
    End Sub
    Private Sub PBoxLastPDF_MouseHover(sender As Object, e As EventArgs) Handles PBoxLastPDF.MouseHover
        Dim Tip As ToolTip = New ToolTip
        Tip.SetToolTip(PBoxLastPDF, "¿Deseas obtener una ruta anterior?" & Chr(10) & "Da clic a continuación.")
        Dim ImgActive As System.Drawing.Bitmap = Bitmap.FromFile(System.AppDomain.CurrentDomain.BaseDirectory() + "\Assets\IMG\LastPDF_Active.png")
        PBoxLastPDF.BackgroundImage = ImgActive
    End Sub
    Private Sub PBoxLastPDF_MouseLeave(sender As Object, e As EventArgs) Handles PBoxLastPDF.MouseLeave
        Dim ImgDisable As System.Drawing.Bitmap = Bitmap.FromFile(System.AppDomain.CurrentDomain.BaseDirectory() + "\Assets\IMG\LastPDF_Disable.png")
        PBoxLastPDF.BackgroundImage = ImgDisable
    End Sub

#End Region
#End Region
End Class

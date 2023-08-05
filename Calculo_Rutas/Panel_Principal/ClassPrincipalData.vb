Imports System.Configuration
Imports System.Data.OleDb

Public Class ClassPrincipalData
#Region "---------------------------------------------------------------CONEXION A DB----------------------------------------------------------------"
    Dim CadenaConexion As String = ConfigurationManager.ConnectionStrings("ConexionDB").ConnectionString
#End Region

#Region "---------------------------------------------------------------LOAD PRINCIPAL---------------------------------------------------------------"
    Public Function ValidarConexionP(AlertaIcon As PictureBox, P_Chofer As Panel, P_CalculoRuta As Panel, Panel_Cofiguracion As Panel, PBoxConfiguracion As PictureBox, DTP_Fecha As DateTimePicker, LFecha As Label, TimerErrorAlert As Timer)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            If conexionDB.State = ConnectionState.Closed Then
                conexionDB.Open()
                AlertaIcon.Visible = False
                TimerErrorAlert.Stop()
                conexionDB.Close()
            End If
        Catch ex As Exception
            TimerErrorAlert.Start()
            P_Chofer.Enabled = False
            P_CalculoRuta.Enabled = False
            Panel_Cofiguracion.Enabled = False
            PBoxConfiguracion.Enabled = False
            DTP_Fecha.Enabled = False
            LFecha.Enabled = False
            MsgBox(ex.Message)
        End Try
    End Function
#End Region

#Region "--------------------------------------------------------PANEL PRINCIPAL (COMBOBOX)----------------------------------------------------------"
    Public Function MostrarClientes(LCliente As Label, CMB_Cliente As ComboBox)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT Nombre FROM (Clientes " &
                                    "INNER JOIN Rutas ON Clientes.IdCliente = Rutas.Cliente_ID) INNER JOIN InfoRutas ON Rutas.IdRuta = InfoRutas.Ruta_ID " &
                                    "WHERE Status = True " &
                                    "GROUP BY Nombre;"
            conexionDB.Open()
            Dim adap As OleDbDataAdapter = New OleDbDataAdapter(consulta, conexionDB)
            Dim dsDatos As DataTable = New DataTable()
            adap.Fill(dsDatos)
            Return dsDatos
            conexionDB.Close()
        Catch ex As Exception
            LCliente.Enabled = False
            CMB_Cliente.Enabled = False
        End Try
    End Function
    Public Function MostrarChofer(LChofer As Label, CMB_Chofer As ComboBox)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT * FROM Choferes WHERE Status = True"
            conexionDB.Open()
            Dim adap As OleDbDataAdapter = New OleDbDataAdapter(consulta, conexionDB)
            Dim dsDatos As DataTable = New DataTable()
            adap.Fill(dsDatos)
            Return dsDatos
            conexionDB.Close()
        Catch ex As Exception
            LChofer.Enabled = False
            CMB_Chofer.Enabled = False
        End Try
    End Function
    Public Function MostrarUnidades(LUnidad As Label, CMB_Vehiculo As ComboBox, CMB_Cliente As ComboBox)
        Try
            Dim conexionDB As New OleDbConnection(CadenaConexion)
            Dim consulta As String = "SELECT Vehiculo FROM Unidades " &
                                    "INNER JOIN ((Clientes INNER JOIN Rutas ON Clientes.IdCliente = Rutas.Cliente_ID) INNER JOIN InfoRutas ON Rutas.IdRuta = InfoRutas.Ruta_ID) ON Unidades.IdUnidad = InfoRutas.Unidad_ID " &
                                    "WHERE Nombre = '" & CMB_Cliente.Text & "' " &
                                    "GROUP BY Vehiculo;"
            conexionDB.Open()
            Dim adap As OleDbDataAdapter = New OleDbDataAdapter(consulta, conexionDB)
            Dim dsDatos As DataTable = New DataTable()
            adap.Fill(dsDatos)
            Return dsDatos
            conexionDB.Close()
        Catch ex As Exception
            LUnidad.Enabled = False
            CMB_Vehiculo.Enabled = False
        End Try
    End Function
#End Region


#Region "----------------------------------------------------------PANEL PRINCIPAL (RUTAS)-----------------------------------------------------------"
    Public Function MostrarRutas(LRuta As Label, L_Ruta_Destino As Label, CMB_Cliente As ComboBox)
        Try
            Dim conexionDB As New OleDbConnection(CadenaConexion)
            Dim command As OleDbCommand
            Dim consulta As String = "SELECT Domicilio FROM Unidades " &
                                    "INNER JOIN ((Clientes INNER JOIN Rutas ON Clientes.IdCliente = Rutas.Cliente_ID) INNER JOIN InfoRutas ON Rutas.IdRuta = InfoRutas.Ruta_ID) ON Unidades.IdUnidad = InfoRutas.Unidad_ID " &
                                    "WHERE Nombre = '" & CMB_Cliente.Text & "' AND Status = True " &
                                    "GROUP BY Domicilio;"
            command = New OleDbCommand(consulta, conexionDB)
            conexionDB.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            If reader.Read() Then
                L_Ruta_Destino.Text = reader.GetValue(0)
            End If
            reader.Close()
            conexionDB.Close()
        Catch ex As Exception
            LRuta.Enabled = False
            L_Ruta_Destino.Enabled = False
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | MostrarRutas")
        End Try
    End Function
    Public Function MostrarCasetas(CMB_Vehiculo As ComboBox, L_Ruta_Destino As Label)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT ('$ '& InfoRutas.Importe_Caseta &'      '& Casetas.Caseta) AS CasetaImporte FROM Casetas " &
                                     "INNER JOIN (((Clientes " &
                                        "INNER JOIN Rutas ON Clientes.IdCliente = Rutas.Cliente_ID) " &
                                        "INNER JOIN InfoRutas ON Rutas.IdRuta = InfoRutas.Ruta_ID) " &
                                        "INNER JOIN Unidades ON InfoRutas.Unidad_ID = Unidades.IdUnidad) " &
                                     "ON Casetas.IdCaseta = InfoRutas.Caseta_ID " &
                                     "WHERE Clientes.Domicilio = '" & L_Ruta_Destino.Text & "' " &
                                     "AND Unidades.Vehiculo = '" & CMB_Vehiculo.Text & "';"
            conexionDB.Open()
            Dim adap As OleDbDataAdapter = New OleDbDataAdapter(consulta, conexionDB)
            Dim dsDatos As DataTable = New DataTable()
            adap.Fill(dsDatos)
            Return dsDatos
            conexionDB.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
    Public Function MostrarTotalCasetas(CMB_Vehiculo As ComboBox, L_Ruta_Destino As Label, Total_Casetas As Label)
        Total_Casetas.Text = "0.00"
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Dim command As OleDbCommand
        Try
            Dim consulta As String = "SELECT SUM(Importe_Caseta) FROM Casetas " &
                                     "INNER JOIN (((Clientes " &
                                        "INNER JOIN Rutas ON Clientes.IdCliente = Rutas.Cliente_ID) " &
                                        "INNER JOIN InfoRutas ON Rutas.IdRuta = InfoRutas.Ruta_ID) " &
                                        "INNER JOIN Unidades ON InfoRutas.Unidad_ID = Unidades.IdUnidad) " &
                                     "ON Casetas.IdCaseta = InfoRutas.Caseta_ID " &
                                     "WHERE Clientes.Domicilio = '" & L_Ruta_Destino.Text & "' " &
                                     "AND Unidades.Vehiculo = '" & CMB_Vehiculo.Text & "'  AND Status = True;"
            command = New OleDbCommand(consulta, conexionDB)
            conexionDB.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            If reader.Read() Then
                Total_Casetas.Text = reader.GetValue(0)
            End If
            reader.Close()
            conexionDB.Close()
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | MostrarTotalCasetas")
        End Try
    End Function
    Public Function MostrarCombustible(CMB_Vehiculo As ComboBox, L_Ruta_Destino As Label, Total_Combustible As Label, TxTCostoCombustible As TextBox)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Dim command As OleDbCommand
        Try
            Dim consulta As String = "SELECT Litros_Combustible FROM Casetas " &
                                     "INNER JOIN (((Clientes " &
                                        "INNER JOIN Rutas ON Clientes.IdCliente = Rutas.Cliente_ID) " &
                                        "INNER JOIN InfoRutas ON Rutas.IdRuta = InfoRutas.Ruta_ID) " &
                                        "INNER JOIN Unidades ON InfoRutas.Unidad_ID = Unidades.IdUnidad) " &
                                     "ON Casetas.IdCaseta = InfoRutas.Caseta_ID " &
                                     "WHERE Clientes.Domicilio = '" & L_Ruta_Destino.Text & "' " &
                                     "AND Unidades.Vehiculo = '" & CMB_Vehiculo.Text & "' AND Status = True " &
                                     "GROUP BY Litros_Combustible;"
            command = New OleDbCommand(consulta, conexionDB)
            conexionDB.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            If reader.Read() Then
                Dim Combustible, PrecioCombustible As Double
                Combustible = reader.GetDouble(0)
                PrecioCombustible = TruncateDecimal((Combustible * TxTCostoCombustible.Text) * 2, 2)
                Total_Combustible.Text = PrecioCombustible
            End If
            reader.Close()
            conexionDB.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | MostrarCombustible")
        End Try
    End Function
    Public Function MostrarKlmTimeT(CMB_Vehiculo As ComboBox, L_Ruta_Destino As Label, LKilometrosPDF As Label, LTiempoTrayectoPDF As Label)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Dim command As OleDbCommand
        Try
            Dim consulta As String = "SELECT Kilometros, Tiempo_Trayecto FROM Casetas " &
                                     "INNER JOIN (((Clientes " &
                                        "INNER JOIN Rutas ON Clientes.IdCliente = Rutas.Cliente_ID) " &
                                        "INNER JOIN InfoRutas ON Rutas.IdRuta = InfoRutas.Ruta_ID) " &
                                        "INNER JOIN Unidades ON InfoRutas.Unidad_ID = Unidades.IdUnidad) " &
                                     "ON Casetas.IdCaseta = InfoRutas.Caseta_ID " &
                                     "WHERE Clientes.Domicilio = '" & L_Ruta_Destino.Text & "' " &
                                     "AND Unidades.Vehiculo = '" & CMB_Vehiculo.Text & "' AND Status = True " &
                                     "GROUP BY Kilometros, Tiempo_Trayecto;"
            command = New OleDbCommand(consulta, conexionDB)
            conexionDB.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            If reader.Read() Then
                LKilometrosPDF.Text = reader.GetDouble(0)
                LTiempoTrayectoPDF.Text = reader.GetValue(1)
            End If
            reader.Close()
            conexionDB.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
    Public Function MostrarDatosBusqueda(Total_Casetas As Label, Total_Combustible As Label, L_Ruta_Destino As Label, CMB_Vehiculo As ComboBox, L_Desgloce_Casetas As ListBox, TxTCostoCombustible As TextBox, LKilometrosPDF As Label, LTiempoTrayectoPDF As Label)
        Total_Casetas.Text = "0.00"
        Total_Combustible.Text = "0.00"

        If L_Ruta_Destino.Text <> "Corporativo LUIN" And CMB_Vehiculo.Text <> String.Empty Then
            Dim dtCasetas As DataTable = MostrarCasetas(CMB_Vehiculo, L_Ruta_Destino)
            L_Desgloce_Casetas.DataSource = dtCasetas
            L_Desgloce_Casetas.DisplayMember = "CasetaImporte"
        End If

        If L_Ruta_Destino.Text <> "Corporativo LUIN" And CMB_Vehiculo.Text <> String.Empty And TxTCostoCombustible.Text <> String.Empty Then
            MostrarTotalCasetas(CMB_Vehiculo, L_Ruta_Destino, Total_Casetas)
            MostrarCombustible(CMB_Vehiculo, L_Ruta_Destino, Total_Combustible, TxTCostoCombustible)
            MostrarKlmTimeT(CMB_Vehiculo, L_Ruta_Destino, LKilometrosPDF, LTiempoTrayectoPDF)
        Else
            Total_Casetas.Text = "0.00"
            Total_Combustible.Text = "0.00"
        End If
    End Function
#End Region


#Region "-------------------------------------------------------PANEL PRINCIPAL (OPCIONES EXTRA)-----------------------------------------------------"
    Public Function ObtenerIDsBitcaora(L_Ruta_Destino As Label, LRutaIDBit As Label, LOrigenBit As Label, LLitroCombustibleBit As Label, CMB_Chofer As ComboBox, LChoferIDBit As Label)
        If L_Ruta_Destino.Text <> "Corporativo LUIN" Then
            Try
                Dim conexionDB As New OleDbConnection(CadenaConexion)
                Dim consulta As String = "SELECT idRuta, Origen, Litros_Combustible FROM Clientes " &
                                        "INNER JOIN Rutas ON Clientes.IdCliente = Rutas.Cliente_ID " &
                                        "WHERE Domicilio = '" & L_Ruta_Destino.Text & "' AND Status = True " &
                                        "GROUP BY idRuta, Origen, Litros_Combustible;"
                Dim comando As OleDbCommand = New OleDbCommand(consulta)
                comando.Connection = conexionDB
                conexionDB.Open()
                Dim reader As OleDbDataReader = comando.ExecuteReader
                While reader.Read
                    LRutaIDBit.Text = reader.GetInt32(0)
                    LOrigenBit.Text = reader.GetString(1)
                    LLitroCombustibleBit.Text = reader.GetDouble(2)
                End While
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | RegistrarBitacora01")
            End Try
        End If

        If CMB_Chofer.Text <> String.Empty Then
            Try
                Dim conexionDB As New OleDbConnection(CadenaConexion)
                Dim consulta As String = "SELECT IdChofer FROM Choferes " &
                                        "WHERE Nombre = '" & CMB_Chofer.Text & "' AND Status = True ;"
                Dim comando As OleDbCommand = New OleDbCommand(consulta)
                comando.Connection = conexionDB
                conexionDB.Open()
                Dim reader As OleDbDataReader = comando.ExecuteReader
                While reader.Read
                    LChoferIDBit.Text = reader.GetInt32(0)
                End While
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | RegistrarBitacora02")
            End Try
        End If
    End Function
    Public Function RegistrarBitacora(DTP_Fecha As DateTime, LRutaIDBit As Label, LChoferIDBit As Label, LOrigenBit As Label, L_Ruta_Destino As Label, LKilometrosPDF As Label, LTiempoTrayectoPDF As Label, TxTViaticos As TextBox, LLitroCombustibleBit As Label, TxTCostoCombustible As TextBox, Total_Casetas As Label, TXT_Notas As TextBox, E_Alimentos As TextBox, CMB_Chofer As ComboBox)
        Try
            If LRutaIDBit.Text <> "LRutaIDBit" And LChoferIDBit.Text <> "LChoferIDBit" And LOrigenBit.Text <> "LOrigenBit" And L_Ruta_Destino.Text <> "Corporativo LUIN" And LKilometrosPDF.Text <> "LKilometrosPDF" And LTiempoTrayectoPDF.Text <> "LTiempoTrayectoPDF" And TxTViaticos.Text <> String.Empty And LLitroCombustibleBit.Text <> "LLitroCombustibleBit" And TxTCostoCombustible.Text <> String.Empty And Total_Casetas.Text <> String.Empty And E_Alimentos.Text <> String.Empty Then
                Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
                Dim consulta = "INSERT INTO Bitacoras(Fecha_Registro_Ruta, Ruta_ID, Chofer_ID, Origen, Destino, Kilometros, Tiempo_Trayecto, Viaticos, Litros_Combustible, PL_Combustible, Total_Importe_Casetas, Descripcion, Gasto_Alimentacion, Fecha_Registro) " &
                                "VALUES('" & DTP_Fecha.Date & "', " & LRutaIDBit.Text & ", " & LChoferIDBit.Text & ", '" & LOrigenBit.Text & "', '" & L_Ruta_Destino.Text & "', " & LKilometrosPDF.Text & ", '" & LTiempoTrayectoPDF.Text & "', " & TxTViaticos.Text & ", " & LLitroCombustibleBit.Text & ", " & TxTCostoCombustible.Text & ", " & Total_Casetas.Text & ", '" & TXT_Notas.Text & "', " & E_Alimentos.Text & ", '" & Date.Now() & "')"
                conexionDB.Open()
                Dim comando As OleDbCommand = New OleDbCommand(consulta, conexionDB)
                comando.ExecuteNonQuery()
            Else
                MsgBox("Rellene todos los campos", MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | RegistrarBitacora03")
        End Try
    End Function
    Public Function LimpiarInfo(CMB_Cliente As ComboBox, CMB_Chofer As ComboBox, CMB_Vehiculo As ComboBox, L_Ruta_Destino As Label, E_Alimentos As TextBox, TxTViaticos As TextBox, TXT_Notas As TextBox, L_Desgloce_Casetas As ListBox, Total_Casetas As Label, Total_Combustible As Label, CBox_Alimentos As CheckBox)
        CMB_Cliente.SelectedIndex = -1
        CMB_Chofer.SelectedIndex = -1
        CMB_Vehiculo.SelectedIndex = -1
        L_Ruta_Destino.Text = "Corporativo LUIN"
        E_Alimentos.Text = "0.00"
        TxTViaticos.Text = "0.00"
        CBox_Alimentos.Checked = False
        TXT_Notas.Text = ""
        L_Desgloce_Casetas.SelectionMode = SelectionMode.One
        L_Desgloce_Casetas.DataSource = Nothing
        L_Desgloce_Casetas.Items.Clear()
        L_Desgloce_Casetas.SelectionMode = SelectionMode.None
        Total_Casetas.Text = "0.00"
        Total_Combustible.Text = "0.00"
    End Function
#End Region




    Public Shared Function TruncateDecimal(valor As Decimal, decimales As Integer) As Decimal
        Dim stepper As Decimal = Math.Pow(10, decimales)
        Dim tmp As Integer = Math.Truncate(stepper * valor)
        Return tmp / stepper
    End Function
End Class

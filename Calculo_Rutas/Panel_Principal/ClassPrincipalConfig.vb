Imports System.Configuration
Imports System.Data.OleDb
Imports System.Windows.Interop

Public Class ClassPrincipalConfig
#Region "---------------------------------------------------------------VARIABLES GLOBALES----------------------------------------------------------------"
    'INSTANCIA A FORM_ACTION_CONFIGURATION
    Dim FormAC As New WinRegistros
    Dim FormCasetas As WinCasetas = New WinCasetas
#End Region
#Region "---------------------------------------------------------------CONEXION A DB----------------------------------------------------------------"
    Dim CadenaConexion As String = ConfigurationManager.ConnectionStrings("ConexionDB").ConnectionString
#End Region
#Region "---------------------------------------------------------------PANEL PRINCIPAL (PDF)---------------------------------------------------------------"
    Public Function LastPDF(Panel As String, Form As Form)
        Try
            Form.Opacity = 0.6
            FormAC.LLastPDF.Text = Panel
            FormAC.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
#End Region




#Region "---------------------------------------------------------------PANEL CONFIGURACION (CATALOGO)---------------------------------------------------------------"
    'MOSTRAR TODOS LOS REGISTROS DE UN CATALOGO (SHOW)
    Public Function MostrarCatalogo(CMB_Directorio As ComboBox)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT * FROM " + CMB_Directorio.Text
            conexionDB.Open()
            Dim adap As OleDbDataAdapter = New OleDbDataAdapter(consulta, conexionDB)
            Dim dsDatos As DataTable = New DataTable()
            adap.Fill(dsDatos)
            Return dsDatos
            conexionDB.Close()
        Catch ex As Exception

        End Try
    End Function
    'CREAR UN NUEVO REGISTRO EN ALGUN CATALOGO (REGISTER)
    Public Function NewCatalogo(CMB_Directorio As ComboBox, Form As Form)
        Try
            Form.Opacity = 0.6
            FormAC.LInfoTabla.Text = CMB_Directorio.Text
            FormAC.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
    'OBTENER EL REGISTRO DE UN CATALOGO PARA MODIFICAR (UPDATE)
    Public Function ObtenerCatalogo(CMB_Directorio As ComboBox, Lista_Catalago As ListBox, Form As Form)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Form.Opacity = 0.6
            Dim Item As DataRowView = Lista_Catalago.SelectedItem
            Dim CatalogoSelect As DataRow = Item.Row

            If CMB_Directorio.Text = "Clientes" Then
                Dim consulta As String = "SELECT * FROM Clientes WHERE Nombre = '" & CatalogoSelect(1) & "';"
                Dim comando As OleDbCommand = New OleDbCommand(consulta)
                comando.Connection = conexionDB
                conexionDB.Open()
                Dim reader As OleDbDataReader = comando.ExecuteReader
                While reader.Read
                    FormAC.LIdTabla.Text = reader.GetInt32(0)
                    FormAC.TxTNombreCUp.Text = reader.GetString(1)
                    FormAC.TxTDomicilioCUp.Text = reader.GetString(2)
                    FormAC.CBoxStatusCUp.Checked = reader.GetBoolean(3)
                End While
            ElseIf CMB_Directorio.Text = "Choferes" Then
                Dim consulta As String = "SELECT * FROM Choferes WHERE Nombre = '" & CatalogoSelect(1) & "';"
                Dim comando As OleDbCommand = New OleDbCommand(consulta)
                comando.Connection = conexionDB
                conexionDB.Open()
                Dim reader As OleDbDataReader = comando.ExecuteReader
                While reader.Read
                    FormAC.LIdTabla.Text = reader.GetInt32(0)
                    FormAC.TxTNombreCHUp.Text = reader.GetString(1)
                    FormAC.TxTTelefonoCHUp.Text = reader.GetString(2)
                    FormAC.CBoxStatusCHUp.Checked = reader.GetBoolean(3)
                End While
            ElseIf CMB_Directorio.Text = "Unidades" Then
                Dim consulta As String = "SELECT * FROM Unidades WHERE Vehiculo = '" & CatalogoSelect(1) & "';"
                Dim comando As OleDbCommand = New OleDbCommand(consulta)
                comando.Connection = conexionDB
                conexionDB.Open()
                Dim reader As OleDbDataReader = comando.ExecuteReader
                While reader.Read
                    FormAC.LIdTabla.Text = reader.GetInt32(0)
                    FormAC.TxTVehiculoUUp.Text = reader.GetString(1)
                    FormAC.TxTPlacasUUp.Text = reader.GetString(2)
                    FormAC.CmbDescripcionUUp.Text = reader.GetString(3)
                End While
            ElseIf CMB_Directorio.Text = "Casetas" Then
                Dim consulta As String = "SELECT * FROM Casetas WHERE Caseta = '" & CatalogoSelect(1) & "';"
                Dim comando As OleDbCommand = New OleDbCommand(consulta)
                comando.Connection = conexionDB
                conexionDB.Open()
                Dim reader As OleDbDataReader = comando.ExecuteReader
                While reader.Read
                    FormAC.LIdTabla.Text = reader.GetInt32(0)
                    FormAC.TxTCasetaUp.Text = reader.GetString(1)
                End While
            End If
            FormAC.LUpTabla.Text = CMB_Directorio.Text
            FormAC.ShowDialog()
            conexionDB.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
    Public Function EliminarCatalogo(CMB_Directorio As ComboBox, Lista_Catalago As ListBox)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim Item As DataRowView = Lista_Catalago.SelectedItem
            Dim CatalogoSelect As DataRow = Item.Row

            If CMB_Directorio.Text = "Clientes" Then
                Dim ID_Cliente, ID_Ruta As Integer
                'HACER LOS SELECT PARA OBTENER TODOS LOS ID's
                Dim consultaID01 As String = "SELECT IdCliente FROM Clientes WHERE Nombre = '" & CatalogoSelect(1) & "';"
                Dim comandoID01 As OleDbCommand = New OleDbCommand(consultaID01)
                comandoID01.Connection = conexionDB
                conexionDB.Open()
                Dim reader01 As OleDbDataReader = comandoID01.ExecuteReader
                While reader01.Read
                    ID_Cliente = reader01.GetInt32(0)
                End While
                '---------------------------------------------
                Dim consultaID02 As String = "SELECT IdRuta FROM Rutas WHERE Cliente_ID = " & ID_Cliente & ";"
                Dim comandoID02 As OleDbCommand = New OleDbCommand(consultaID02)
                comandoID02.Connection = conexionDB
                Dim reader02 As OleDbDataReader = comandoID02.ExecuteReader
                While reader02.Read
                    ID_Ruta = reader02.GetInt32(0)
                End While
                'HACER LOS DELETE PARA ELIMINAR TODAS LAS UNIONES DE ID's A OTRAS TABLAS
                Dim consulta01 As String = "DELETE * FROM InfoRutas WHERE Ruta_ID = " & ID_Ruta & ";"
                Dim comando01 As OleDbCommand = New OleDbCommand(consulta01)
                comando01.Connection = conexionDB
                comando01.ExecuteReader()
                '---------------------------------------------
                Dim consulta02 As String = "DELETE * FROM Rutas WHERE Cliente_ID = " & ID_Cliente & ";"
                Dim comando02 As OleDbCommand = New OleDbCommand(consulta02)
                comando02.Connection = conexionDB
                comando02.ExecuteReader()
                '---------------------------------------------
                Dim consulta03 As String = "DELETE * FROM Clientes WHERE Nombre = '" & CatalogoSelect(1) & "';"
                Dim comando03 As OleDbCommand = New OleDbCommand(consulta03)
                comando03.Connection = conexionDB
                comando03.ExecuteReader()
                MsgBox("Cliente eliminado", MsgBoxStyle.Information, "Información | Corporativo LUIN")
                conexionDB.Close()

            ElseIf CMB_Directorio.Text = "Choferes" Then
                Dim consulta01 As String = "DELETE * FROM Choferes WHERE Nombre = '" & CatalogoSelect(1) & "';"
                Dim comando01 As OleDbCommand = New OleDbCommand(consulta01)
                comando01.Connection = conexionDB
                conexionDB.Open()
                comando01.ExecuteReader()
                MsgBox("Chofer eliminado", MsgBoxStyle.Information, "Información | Corporativo LUIN")
                conexionDB.Close()

            ElseIf CMB_Directorio.Text = "Unidades" Then
                Dim ID_Unidad As Integer
                'HACER LOS SELECT PARA OBTENER TODOS LOS ID's
                Dim consultaID01 As String = "SELECT IdUnidad FROM Unidades WHERE Vehiculo = '" & CatalogoSelect(1) & "';"
                Dim comandoID01 As OleDbCommand = New OleDbCommand(consultaID01)
                comandoID01.Connection = conexionDB
                conexionDB.Open()
                Dim reader01 As OleDbDataReader = comandoID01.ExecuteReader
                While reader01.Read
                    ID_Unidad = reader01.GetInt32(0)
                End While
                'HACER LOS DELETE PARA ELIMINAR TODAS LAS UNIONES DE ID's A OTRAS TABLAS
                Dim consulta01 As String = "DELETE * FROM InfoRutas WHERE Unidad_ID = " & ID_Unidad & ";"
                Dim comando01 As OleDbCommand = New OleDbCommand(consulta01)
                comando01.Connection = conexionDB
                comando01.ExecuteReader()
                '---------------------------------------------
                Dim consulta02 As String = "DELETE * FROM Unidades WHERE Vehiculo = '" & CatalogoSelect(1) & "';"
                Dim comando02 As OleDbCommand = New OleDbCommand(consulta02)
                comando02.Connection = conexionDB
                comando02.ExecuteReader()
                '---------------------------------------------
                MsgBox("Unidad eliminada", MsgBoxStyle.Information, "Información | Corporativo LUIN")
                conexionDB.Close()

            ElseIf CMB_Directorio.Text = "Casetas" Then
                Dim ID_Caseta As Integer
                'HACER LOS SELECT PARA OBTENER TODOS LOS ID's
                Dim consultaID01 As String = "SELECT IdCaseta FROM Casetas WHERE Caseta = '" & CatalogoSelect(1) & "';"
                Dim comandoID01 As OleDbCommand = New OleDbCommand(consultaID01)
                comandoID01.Connection = conexionDB
                conexionDB.Open()
                Dim reader01 As OleDbDataReader = comandoID01.ExecuteReader
                While reader01.Read
                    ID_Caseta = reader01.GetInt32(0)
                End While
                'HACER LOS DELETE PARA ELIMINAR TODAS LAS UNIONES DE ID's A OTRAS TABLAS
                Dim consulta01 As String = "DELETE * FROM InfoRutas WHERE Caseta_ID = " & ID_Caseta & ";"
                Dim comando01 As OleDbCommand = New OleDbCommand(consulta01)
                comando01.Connection = conexionDB
                comando01.ExecuteReader()
                '---------------------------------------------
                Dim consulta02 As String = "DELETE * FROM Casetas WHERE Caseta = '" & CatalogoSelect(1) & "';"
                Dim comando02 As OleDbCommand = New OleDbCommand(consulta02)
                comando02.Connection = conexionDB
                comando02.ExecuteReader()
                '---------------------------------------------
                MsgBox("Caseta eliminada", MsgBoxStyle.Information, "Información | Corporativo LUIN")
                conexionDB.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | EliminarCatalogo")
        End Try
    End Function
#End Region


#Region "---------------------------------------------------------------PANEL CONFIGURACION (RUTAS)---------------------------------------------------------------"
    'MOSTRAR TODOS LOS REGISTROS DE UNA RUTA
    Public Function MostrarRutas()
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT Nombre " &
                                    "FROM Clientes INNER JOIN Rutas ON Clientes.IdCliente = Rutas.Cliente_ID " &
                                    "WHERE Status = True " &
                                    "GROUP BY Nombre;"
            conexionDB.Open()
            Dim adap As OleDbDataAdapter = New OleDbDataAdapter(consulta, conexionDB)
            Dim dsDatos As DataTable = New DataTable()
            adap.Fill(dsDatos)
            Return dsDatos
            conexionDB.Close()
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | MostrarRutas")
        End Try
    End Function
    Public Function MostrarDatosRuta(CMB_Rutas As ComboBox, TxTOrigen As Label, TxTDestino As Label, TxTKilometros As Label, TxTTiempoTrayecto As Label, TxTTOKA As Label, TxTFegali As Label, LCombustible As Label, LIDRuta As Label)
        If CMB_Rutas.Text <> String.Empty Then
            Dim conexionDB As New OleDbConnection(CadenaConexion)
            Try
                Dim consulta As String = "SELECT Origen, Domicilio, Kilometros, Tiempo_Trayecto, TOKA, Fegali, Litros_Combustible, IdRuta
                                         From Clientes INNER JOIN Rutas ON Clientes.IdCliente = Rutas.Cliente_ID
                                         WHERE Nombre = '" & CMB_Rutas.Text & "' AND Status = True ;"
                Dim comando As OleDbCommand = New OleDbCommand(consulta)
                comando.Connection = conexionDB
                conexionDB.Open()
                Dim reader As OleDbDataReader = comando.ExecuteReader
                While reader.Read
                    TxTOrigen.Text = reader.GetString(0)
                    TxTDestino.Text = reader.GetString(1)
                    TxTKilometros.Text = reader.GetDouble(2) & " KM"
                    TxTTiempoTrayecto.Text = reader.GetString(3)
                    TxTTOKA.Text = reader.GetDouble(4) & " LTS"
                    TxTFegali.Text = reader.GetDouble(5) & " LTS"
                    LCombustible.Text = reader.GetDouble(6) & " LTS"
                    LIDRuta.Text = reader.GetInt32(7)
                End While
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | MostrarDatosRuta")
            End Try
        End If
    End Function
    'CREAR UN NUEVO REGISTRO DE RUTA
    Public Function NewRuta(Form As Form)
        Try
            Form.Opacity = 0.6
            FormAC.LNewRuta.Text = "NuevaRuta"
            FormAC.StartPosition = FormStartPosition.CenterScreen
            FormAC.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | NewRuta")
        End Try
    End Function
    'ACTUALIZAR LA INFO DE UNA RUTA
    Public Function ObtenerInfoRuta(LIDRuta As Label, Form As Form)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Form.Opacity = 0.6
            Dim consulta As String = "SELECT * FROM Rutas WHERE IdRuta = " & LIDRuta.Text & ";"
            Dim comando As OleDbCommand = New OleDbCommand(consulta)
            comando.Connection = conexionDB
            conexionDB.Open()
            Dim reader As OleDbDataReader = comando.ExecuteReader
            While reader.Read
                FormAC.LUpIDRuta.Text = reader.GetInt32(0)
                FormAC.TxTOrigenUp.Text = reader.GetString(1)
                FormAC.TxTKilometrosUp.Text = reader.GetDouble(2)
                FormAC.TxTTTrayectoUp.Text = reader.GetString(3)
                FormAC.TxTTOKAUp.Text = reader.GetDouble(4)
                FormAC.TxTFegaliUp.Text = reader.GetDouble(5)
                FormAC.LCombustibleUp.Text = reader.GetDouble(6)
                FormAC.LIDCliente.Text = reader.GetInt32(7)
            End While
            FormAC.LUpRuta.Text = "ActualizarRuta"
            FormAC.ShowDialog()
            conexionDB.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | ObtenerInfoRuta")
        End Try
    End Function
    Public Function EliminarRuta(LIDRuta As Label)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "DELETE * FROM Rutas WHERE IdRuta = " & LIDRuta.Text
            Dim comando As OleDbCommand = New OleDbCommand(consulta)
            comando.Connection = conexionDB
            conexionDB.Open()
            Dim reader As OleDbDataReader = comando.ExecuteReader
            MsgBox("Ruta eliminada", MsgBoxStyle.Information, "Información | Corporativo LUIN")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
        End Try
    End Function
    Public Function EliminarCasetasRuta(LIDRuta As Label)
        Try
            Dim conexionDB As New OleDbConnection(CadenaConexion)
            Dim consulta As String = "DELETE * FROM InfoRutas WHERE Ruta_ID = " & LIDRuta.Text
            Dim comando As OleDbCommand = New OleDbCommand(consulta)
            comando.Connection = conexionDB
            conexionDB.Open()
            Dim reader As OleDbDataReader = comando.ExecuteReader
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | EliminarCasetaRuta")
        End Try
    End Function
#End Region


#Region "---------------------------------------------------------------PANEL CONFIGURACION (CASETAS)---------------------------------------------------------------"
#Region "MOSTRAR RUTAS Y UNIDADES"
    Public Function MostrarRutasPCasetas()
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
            'MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | MostrarRutas")
        End Try
    End Function
    Public Function MostrarRutasC()
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT Nombre FROM Clientes WHERE Status = True "
            conexionDB.Open()
            Dim adap As OleDbDataAdapter = New OleDbDataAdapter(consulta, conexionDB)
            Dim dsDatos As DataTable = New DataTable()
            adap.Fill(dsDatos)
            Return dsDatos
            conexionDB.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | MostrarRutasC")
        End Try
    End Function
    Public Function ObtenerIDDomicilioIDVehiculo(CmbRutaImporte As ComboBox, CmbVehiculoImporte As ComboBox, LRutaID As Label, LVehiculoID As Label)
        If CmbRutaImporte.Text <> String.Empty Then
            Try
                Dim conexionDB As New OleDbConnection(CadenaConexion)
                Dim command As OleDbCommand
                Dim consulta As String = "SELECT IdRuta FROM Clientes INNER JOIN Rutas ON Clientes.IdCliente = Rutas.Cliente_ID WHERE Nombre = '" & CmbRutaImporte.Text & "' AND Status = True ;"
                command = New OleDbCommand(consulta, conexionDB)
                conexionDB.Open()
                Dim reader As OleDbDataReader = command.ExecuteReader()
                If reader.Read() Then
                    LRutaID.Text = reader.GetValue(0)
                End If
                reader.Close()
                conexionDB.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

        If CmbVehiculoImporte.Text <> String.Empty Then
            Try
                Dim conexionDB As New OleDbConnection(CadenaConexion)
                Dim command As OleDbCommand
                Dim consulta As String = "SELECT IdUnidad FROM Unidades WHERE Vehiculo = '" & CmbVehiculoImporte.Text & "';"
                command = New OleDbCommand(consulta, conexionDB)
                conexionDB.Open()
                Dim reader As OleDbDataReader = command.ExecuteReader()
                If reader.Read() Then
                    LVehiculoID.Text = reader.GetValue(0)
                End If
                reader.Close()
                conexionDB.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Function
    Public Function MostrarUnidadesC(LRutaID As Label)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT Vehiculo FROM Clientes INNER JOIN (Rutas INNER JOIN (Unidades INNER Join InfoRutas ON Unidades.IdUnidad = InfoRutas.Unidad_ID) ON Rutas.IdRuta = InfoRutas.Ruta_ID) On Clientes.IdCliente = Rutas.Cliente_ID WHERE IdRuta = " & LRutaID.Text & " GROUP BY Vehiculo;"
            conexionDB.Open()
            Dim adap As OleDbDataAdapter = New OleDbDataAdapter(consulta, conexionDB)
            Dim dsDatos As DataTable = New DataTable()
            adap.Fill(dsDatos)
            Return dsDatos
            conexionDB.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | Mostrar UnidadesC")
        End Try
    End Function
    Public Function MostrarCasetas(CmbRutaImporte As ComboBox, CmbVehiculoImporte As ComboBox)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT ('$ '& InfoRutas.Importe_Caseta &'      '& Casetas.Caseta) AS CasetaImporte FROM Casetas " &
                                     "INNER JOIN (((Clientes " &
                                        "INNER JOIN Rutas ON Clientes.IdCliente = Rutas.Cliente_ID) " &
                                        "INNER JOIN InfoRutas ON Rutas.IdRuta = InfoRutas.Ruta_ID) " &
                                        "INNER JOIN Unidades ON InfoRutas.Unidad_ID = Unidades.IdUnidad) " &
                                     "ON Casetas.IdCaseta = InfoRutas.Caseta_ID " &
                                     "WHERE Clientes.Nombre = '" & CmbRutaImporte.Text & "' " &
                                     "AND Unidades.Vehiculo = '" & CmbVehiculoImporte.Text & "' AND Status = True ;"
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
#End Region
#Region "CREAR UNA NUEVA RUTA-CASETA"
    Public Function NewCasetaRuta(Form As Form)
        Try
            Form.Opacity = 0.6
            Dim Activador As Integer = 1
            Do
                Dim FormCasetas01 As WinCasetas = New WinCasetas
                FormCasetas01.LNewCaseta.Text = "NuevaCasetaRuta"
                FormCasetas01.ShowDialog()
                Activador = FormCasetas01.CloseForm()
            Loop While Activador <> 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | NewCasetaRuta")
        End Try
    End Function
    Public Function ActualizarCasetaRuta(LRutaID As Label, LVehiculoID As Label, Form As Form)
        Try
            Form.Opacity = 0.6
            Dim ActivadorUp As Integer = 1
            Do
                Dim FormCasetas02 As WinCasetas = New WinCasetas
                FormCasetas02.LUpdateCaseta.Text = "UpdateCasetaRuta"
                FormCasetas02.LIDRutaUp.Text = LRutaID.Text
                FormCasetas02.LIDVehiculoUp.Text = LVehiculoID.Text
                FormCasetas02.ShowDialog()
                ActivadorUp = FormCasetas02.CloseForm()
            Loop While ActivadorUp <> 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | NewCasetaRuta")
        End Try
    End Function
    Public Function EliminarCasetaRuta(LRutaID As Label, LVehiculoID As Label)
        Try
            Dim conexionDB As New OleDbConnection(CadenaConexion)
            Dim consulta As String = "DELETE * FROM InfoRutas WHERE Ruta_ID = " & LRutaID.Text & " AND Unidad_ID = " & LVehiculoID.Text
            Dim comando As OleDbCommand = New OleDbCommand(consulta)
            comando.Connection = conexionDB
            conexionDB.Open()
            Dim reader As OleDbDataReader = comando.ExecuteReader
            MsgBox("Casetas eliminadas de la ruta", MsgBoxStyle.Information, "Información | Corporativo LUIN")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | EliminarCasetaRuta")
        End Try
    End Function
#End Region
#End Region


#Region "-------------------------------------------------------PANEL PRINCIPAL (FUNCIONES DE ACTUALIZACION)-----------------------------------------------------"
    Public Function MostrarActualizarCatalogo(L_Directorio As Label, CMB_Directorio As ComboBox, Lista_Catalago As ListBox)
        L_Directorio.Text = CMB_Directorio.Text
        If CMB_Directorio.Text = "Clientes" Then
            Dim dtRutas As DataTable = MostrarCatalogo(CMB_Directorio)
            Lista_Catalago.DataSource = dtRutas
            Lista_Catalago.DisplayMember = "Nombre"
            Lista_Catalago.SelectedIndex = -1
        ElseIf CMB_Directorio.Text = "Choferes" Then
            Dim dtRutas As DataTable = MostrarCatalogo(CMB_Directorio)
            Lista_Catalago.DataSource = dtRutas
            Lista_Catalago.DisplayMember = "Nombre"
            Lista_Catalago.SelectedIndex = -1
        ElseIf CMB_Directorio.Text = "Unidades" Then
            Dim dtRutas As DataTable = MostrarCatalogo(CMB_Directorio)
            Lista_Catalago.DataSource = dtRutas
            Lista_Catalago.DisplayMember = "Vehiculo"
            Lista_Catalago.SelectedIndex = -1
        ElseIf CMB_Directorio.Text = "Casetas" Then
            Dim dtRutas As DataTable = MostrarCatalogo(CMB_Directorio)
            Lista_Catalago.DataSource = dtRutas
            Lista_Catalago.DisplayMember = "Caseta"
            Lista_Catalago.SelectedIndex = -1
        End If
    End Function
#End Region
End Class

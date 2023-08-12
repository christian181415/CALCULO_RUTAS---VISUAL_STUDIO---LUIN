Imports System.Configuration
Imports System.Data.OleDb

Public Class ClassRegistrosData
#Region "---------------------------------------------------------------CONEXION A DB----------------------------------------------------------------"
    Dim CadenaConexion As String = ConfigurationManager.ConnectionStrings("ConexionDB").ConnectionString
#End Region

#Region "---------------------------------------------------------------PANEL CONFIGURACION (CATALOGO)----------------------------------------------------------------"
#Region "---------------------------------------------------------------ACCIONES REGISTER----------------------------------------------------------------"
#Region "---------------------------------------------------------------REGISTRAR CLIENTES----------------------------------------------------------------"
    Public Function RegistrarCliente(TxTNombreC As TextBox, TxTDomicilioC As TextBox, Window As Form, P_NewCliente As Panel)
        Try
            If TxTNombreC.Text <> String.Empty And TxTDomicilioC.Text <> String.Empty Then
                Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
                Dim consulta = "INSERT INTO Clientes(Nombre, Domicilio, Status) VALUES('" & TxTNombreC.Text & "', '" & TxTDomicilioC.Text & "', True)"
                conexionDB.Open()
                Dim comando As OleDbCommand = New OleDbCommand(consulta, conexionDB)
                comando.ExecuteNonQuery()
                MsgBox("Cliente Registrado", MsgBoxStyle.Information, "Exito | Corporativo LUIN")
                TxTNombreC.Text = ""
                TxTDomicilioC.Text = ""
                Window.Close()
                P_NewCliente.Location = New Point(260, 2)
            Else
                MsgBox("Rellene todos los campos", MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
        End Try
    End Function
#End Region



#Region "---------------------------------------------------------------REGISTRAR CHOFERES----------------------------------------------------------------"
    Public Function RegistrarChofer(TxTNombreCH As TextBox, TxTTelefonoCH As MaskedTextBox, Window As Form, P_NewChofer As Panel)
        Try
            If TxTNombreCH.Text <> String.Empty And TxTTelefonoCH.Text <> String.Empty Then
                Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
                Dim consulta = "INSERT INTO Choferes(Nombre, Telefono, Status) VALUES('" & TxTNombreCH.Text & "', '" & TxTTelefonoCH.Text & "', True)"
                conexionDB.Open()
                Dim comando As OleDbCommand = New OleDbCommand(consulta, conexionDB)
                comando.ExecuteNonQuery()
                MsgBox("Chofer Registrado", MsgBoxStyle.Information, "Exito | Corporativo LUIN")
                TxTNombreCH.Text = ""
                TxTTelefonoCH.Text = ""
                Window.Close()
                P_NewChofer.Location = New Point(517, 2)
            Else
                MsgBox("Rellene todos los campos", MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
        End Try
    End Function
#End Region



#Region "---------------------------------------------------------------REGISTRAR UNIDADES----------------------------------------------------------------"
    Public Function RegistrarUnidad(TxTVehiculoU As TextBox, TxTPlacasU As TextBox, CmbDescripcionU As ComboBox, Window As Form, P_NewUnidad As Panel)
        Try
            If TxTVehiculoU.Text <> String.Empty And TxTPlacasU.Text <> String.Empty And CmbDescripcionU.Text <> String.Empty Then
                Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
                Dim consulta = "INSERT INTO Unidades(Vehiculo, Placas, Descripcion) VALUES('" & TxTVehiculoU.Text & "', '" & TxTPlacasU.Text & "', '" & CmbDescripcionU.Text & "')"
                conexionDB.Open()
                Dim comando As OleDbCommand = New OleDbCommand(consulta, conexionDB)
                comando.ExecuteNonQuery()
                MsgBox("Unidad Registrada", MsgBoxStyle.Information, "Exito | Corporativo LUIN")
                TxTVehiculoU.Text = ""
                TxTPlacasU.Text = ""
                CmbDescripcionU.SelectedIndex = -1
                Window.Close()
                P_NewUnidad.Location = New Point(774, 2)
            Else
                MsgBox("Rellene todos los campos", MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
        End Try
    End Function
#End Region



#Region "---------------------------------------------------------------REGISTRAR CASETAS----------------------------------------------------------------"
    Public Function RegistrarCaseta(TxTCaseta As TextBox, Window As Form, P_NewCaseta As Panel)
        Try
            If TxTCaseta.Text <> String.Empty Then
                Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
                Dim consulta = "INSERT INTO Casetas(Caseta) VALUES('" & TxTCaseta.Text & "')"
                conexionDB.Open()
                Dim comando As OleDbCommand = New OleDbCommand(consulta, conexionDB)
                comando.ExecuteNonQuery()
                MsgBox("Caseta Registrada", MsgBoxStyle.Information, "Exito | Corporativo LUIN")
                TxTCaseta.Text = ""
                Window.Close()
                P_NewCaseta.Location = New Point(1034, 2)
            Else
                MsgBox("Rellene todos los campos", MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
        End Try
    End Function
#End Region
#End Region


#Region "---------------------------------------------------------------ACCIONES UPDATE----------------------------------------------------------------"
#Region "---------------------------------------------------------------ACTUALIZAR CLIENTE----------------------------------------------------------------"
    Public Function ActualizarCliente(TxTNombreCUp As TextBox, TxTDomicilioCUp As TextBox, Window As Form, CBoxStatusCUp As CheckBox, LIdTabla As Label)
        Try
            If TxTNombreCUp.Text <> String.Empty And TxTDomicilioCUp.Text <> String.Empty Then
                Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
                Dim consulta = "UPDATE Clientes SET Nombre='" & TxTNombreCUp.Text & "', Domicilio='" & TxTDomicilioCUp.Text & "', Status=" & CBoxStatusCUp.Checked.ToString & " WHERE IdCliente = " & LIdTabla.Text
                conexionDB.Open()
                Dim comando As OleDbCommand = New OleDbCommand(consulta, conexionDB)
                comando.ExecuteNonQuery()
                MsgBox("Cliente Actualizado", MsgBoxStyle.Information, "Exito | Corporativo LUIN")
                TxTNombreCUp.Text = ""
                TxTDomicilioCUp.Text = ""
                Window.Close()
            Else
                MsgBox("Rellene todos los campos", MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
        End Try
    End Function
#End Region


#Region "---------------------------------------------------------------ACTUALIZAR CHOFER----------------------------------------------------------------"
    Public Function ActualizarChofer(TxTNombreCHUp As TextBox, TxTTelefonoCHUp As MaskedTextBox, Window As Form, CBoxStatusCHUp As CheckBox, LIdTabla As Label)
        Try
            If TxTNombreCHUp.Text <> String.Empty And TxTTelefonoCHUp.Text <> String.Empty Then
                Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
                Dim consulta = "UPDATE Choferes SET Nombre='" & TxTNombreCHUp.Text & "', Telefono='" & TxTTelefonoCHUp.Text & "', Status=" & CBoxStatusCHUp.Checked.ToString & " WHERE IdChofer = " & LIdTabla.Text
                conexionDB.Open()
                Dim comando As OleDbCommand = New OleDbCommand(consulta, conexionDB)
                comando.ExecuteNonQuery()
                MsgBox("Chofer Actualizado", MsgBoxStyle.Information, "Exito | Corporativo LUIN")
                TxTNombreCHUp.Text = ""
                TxTTelefonoCHUp.Text = ""
                Window.Close()
            Else
                MsgBox("Rellene todos los campos", MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
        End Try
    End Function
#End Region


#Region "---------------------------------------------------------------ACTUALIZAR UNIDAD----------------------------------------------------------------"
    Public Function ActualizarUnidad(TxTVehiculoUUp As TextBox, TxTPlacasUUp As TextBox, Window As Form, CmbDescripcionUUp As ComboBox, LIdTabla As Label)
        Try
            If TxTVehiculoUUp.Text <> String.Empty And TxTPlacasUUp.Text <> String.Empty And CmbDescripcionUUp.Text <> String.Empty Then
                Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
                Dim consulta = "UPDATE Unidades SET Vehiculo='" & TxTVehiculoUUp.Text & "', Placas='" & TxTPlacasUUp.Text & "', Descripcion='" & CmbDescripcionUUp.Text & "' WHERE IdUnidad = " & LIdTabla.Text
                conexionDB.Open()
                Dim comando As OleDbCommand = New OleDbCommand(consulta, conexionDB)
                comando.ExecuteNonQuery()
                MsgBox("Chofer Actualizado", MsgBoxStyle.Information, "Exito | Corporativo LUIN")
                TxTVehiculoUUp.Text = ""
                TxTPlacasUUp.Text = ""
                CmbDescripcionUUp.SelectedIndex = -1
                CmbDescripcionUUp.Text = ""
                Window.Close()
            Else
                MsgBox("Rellene todos los campos", MsgBoxStyle.Exclamation, "Error01 | Corporativo LUIN")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error02 | Corporativo LUIN")
        End Try
    End Function
#End Region


#Region "---------------------------------------------------------------ACTUALIZAR CASETA----------------------------------------------------------------"
    Public Function ActualizarCaseta(TxTCasetaUp As TextBox, Window As Form, LIdTabla As Label)
        Try
            If TxTCasetaUp.Text <> String.Empty Then
                Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
                Dim consulta = "UPDATE Casetas SET Caseta='" & TxTCasetaUp.Text & "' WHERE IdCaseta = " & LIdTabla.Text
                conexionDB.Open()
                Dim comando As OleDbCommand = New OleDbCommand(consulta, conexionDB)
                comando.ExecuteNonQuery()
                MsgBox("Caseta Actualizada", MsgBoxStyle.Information, "Exito | Corporativo LUIN")
                TxTCasetaUp.Text = ""
                Window.Close()
            Else
                MsgBox("Rellene todos los campos", MsgBoxStyle.Exclamation, "Error01 | Corporativo LUIN")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error02 | Corporativo LUIN")
        End Try
    End Function
#End Region
#End Region
#End Region

#Region "---------------------------------------------------------------PANEL CONFIGURACION (RUTAS)----------------------------------------------------------------"
#Region "---------------------------------------------------------------ACCIONES REGISTER----------------------------------------------------------------"
    Public Function MostrarDestinos(CmbDestino As ComboBox)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT * " &
                                    "FROM Clientes " &
                                    "WHERE IdCliente NOT IN (SELECT Cliente_ID FROM Rutas);"
            conexionDB.Open()
            Dim adap As OleDbDataAdapter = New OleDbDataAdapter(consulta, conexionDB)
            Dim dsDatos As DataTable = New DataTable()
            adap.Fill(dsDatos)
            Return dsDatos
            conexionDB.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | MostrarDestinos")
        End Try
    End Function
    Public Function ObtenerIDCliente(CmbDestino As ComboBox, LIDCliente As Label)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT IdCliente FROM Clientes WHERE Nombre = '" & CmbDestino.Text & "' AND Status = True ;"
            Dim comando As OleDbCommand = New OleDbCommand(consulta)
            comando.Connection = conexionDB
            conexionDB.Open()
            Dim reader As OleDbDataReader = comando.ExecuteReader
            While reader.Read
                LIDCliente.Text = reader.GetInt32(0)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | ObtenerIDCliente")
        End Try
    End Function
    Public Function RegistrarRuta(TxTOrigen As TextBox, CmbDestino As ComboBox, TxTKilometros As TextBox, TxTTTrayecto As MaskedTextBox, TxTTOKA As TextBox, TxTFegali As TextBox, LCombustible As Label, LIDCliente As Label, Window As Form, P_NewRuta As Panel)
        Try
            If TxTOrigen.Text <> String.Empty And CmbDestino.Text <> String.Empty And TxTKilometros.Text <> String.Empty And TxTTTrayecto.Text <> String.Empty And LCombustible.Text <> String.Empty Then
                Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
                Dim consulta = "INSERT INTO Rutas(Origen, Kilometros, Tiempo_Trayecto, TOKA, Fegali, Litros_Combustible, Cliente_ID) VALUES('" & TxTOrigen.Text & "', '" & TxTKilometros.Text & "', '" & TxTTTrayecto.Text & "', '" & TxTTOKA.Text & "', '" & TxTFegali.Text & "', '" & LCombustible.Text & "', '" & LIDCliente.Text & "')"
                conexionDB.Open()
                Dim comando As OleDbCommand = New OleDbCommand(consulta, conexionDB)
                comando.ExecuteNonQuery()
                MsgBox("Ruta Registrada", MsgBoxStyle.Information, "Exito | Corporativo LUIN")
                TxTOrigen.Text = "Corporativo LUIN"
                TxTKilometros.Text = ""
                TxTTTrayecto.Text = ""
                TxTTOKA.Text = ""
                TxTFegali.Text = ""
                LCombustible.Text = ""
                Window.Close()
                P_NewRuta.Location = New Point(260, 681)
            Else
                MsgBox("Rellene todos los campos", MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | RegistrarRuta")
        End Try
    End Function
#End Region


#Region "---------------------------------------------------------------ACCIONES UPDATE----------------------------------------------------------------"
    Public Function ObtenerDomicilioRuta(LIDCliente As Label, LDestinoUp As Label)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT Nombre FROM Clientes WHERE IdCliente = " & LIDCliente.Text & " AND Status = True ;"
            Dim comando As OleDbCommand = New OleDbCommand(consulta)
            comando.Connection = conexionDB
            conexionDB.Open()
            Dim reader As OleDbDataReader = comando.ExecuteReader
            While reader.Read
                LDestinoUp.Text = reader.GetString(0)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | ObtenerIDCliente")
        End Try
    End Function
    Public Function ObtenerIDClienteUp(LIDClienteUp As Label, LDestinoUp As Label)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT IdCliente FROM Clientes WHERE Nombre = '" & LDestinoUp.Text & "' AND Status = True ;"
            Dim comando As OleDbCommand = New OleDbCommand(consulta)
            comando.Connection = conexionDB
            conexionDB.Open()
            Dim reader As OleDbDataReader = comando.ExecuteReader
            While reader.Read
                LIDClienteUp.Text = reader.GetInt32(0)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | ObtenerIDClienteUp")
        End Try
    End Function
    Public Function ActualizarRuta(TxTOrigenUp As TextBox, LDestinoUp As Label, TxTKilometrosUp As TextBox, TxTTTrayectoUp As MaskedTextBox, TxTTOKAUp As TextBox, TxTFegaliUp As TextBox, LCombustibleUp As Label, LIDClienteUp As Label, Window As Form, LUpIDRuta As Label)
        Try
            If TxTOrigenUp.Text <> String.Empty And LDestinoUp.Text <> String.Empty And TxTKilometrosUp.Text <> String.Empty And TxTTTrayectoUp.Text <> String.Empty And LCombustibleUp.Text <> String.Empty Then
                Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
                Dim consulta = "UPDATE Rutas SET Origen='" & TxTOrigenUp.Text & "', Kilometros='" & TxTKilometrosUp.Text & "', Tiempo_Trayecto='" & TxTTTrayectoUp.Text & "', TOKA='" & TxTTOKAUp.Text & "', Fegali='" & TxTFegaliUp.Text & "', Litros_Combustible='" & LCombustibleUp.Text & "', Cliente_ID=" & LIDClienteUp.Text & " WHERE IdRuta = " & LUpIDRuta.Text
                conexionDB.Open()
                Dim comando As OleDbCommand = New OleDbCommand(consulta, conexionDB)
                comando.ExecuteNonQuery()
                MsgBox("Ruta Actualizada", MsgBoxStyle.Information, "Exito | Corporativo LUIN")
                TxTOrigenUp.Text = "Corporativo LUIN"
                TxTKilometrosUp.Text = ""
                TxTTTrayectoUp.Text = ""
                TxTTOKAUp.Text = ""
                TxTFegaliUp.Text = ""
                LCombustibleUp.Text = ""
                Window.Close()
            Else
                MsgBox("Rellene todos los campos", MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | ActualizarRuta")
        End Try
    End Function
#End Region
#End Region


#Region "HISTORIAL PDF"
    Public Function ShowDatePDF(Componente As Object)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT  FORMAT(Fecha_Registro, 'dd/MM/yyyy') FROM Bitacoras" &
                                    " GROUP BY FORMAT(Fecha_Registro, 'dd/MM/yyyy');"
            Dim comando As OleDbCommand = New OleDbCommand(consulta)
            comando.Connection = conexionDB
            conexionDB.Open()
            Dim reader As OleDbDataReader = comando.ExecuteReader
            While reader.Read
                Componente.AddBoldedDate(reader.GetString(0))
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | ObtenerIDClienteUp")
        End Try
    End Function
    Public Function ShowHoursPDF(Componente As Object, ListHoras As ListBox)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT  FORMAT(Fecha_Registro, 'hh:nn:ss am/pm') FROM Bitacoras" &
                                    " WHERE FORMAT(Fecha_Registro, 'dd/MM/yyyy') = '" & Componente.SelectionStart & "'" &
                                    " GROUP BY FORMAT(Fecha_Registro, 'hh:nn:ss am/pm');"
            Dim comando As OleDbCommand = New OleDbCommand(consulta)
            comando.Connection = conexionDB
            conexionDB.Open()
            Dim reader As OleDbDataReader = comando.ExecuteReader
            ListHoras.Items.Clear()
            ListHoras.DataSource = Nothing
            While reader.Read
                ListHoras.Items.Add(reader.GetString(0))
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | ObtenerIDClienteUp")
        End Try
    End Function
    Public Function GetLastPDF(Componente As MonthCalendar, ListHoras As ListBox)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT * FROM Bitacoras" &
                                    " WHERE (((Bitacoras.[Fecha_Registro])=DateValue('" & Componente.SelectionStart & "')+TimeValue('" & ListHoras.SelectedItem & "')));"
            Dim comando As OleDbCommand = New OleDbCommand(consulta)
            comando.Connection = conexionDB
            conexionDB.Open()
            Dim reader As OleDbDataReader = comando.ExecuteReader
            While reader.Read
                MsgBox(reader.GetString(5))
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | ObtenerIDClienteUp")
        End Try
    End Function
#End Region
End Class

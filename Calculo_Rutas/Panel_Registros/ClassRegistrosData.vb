Imports System.Configuration
Imports System.Data.OleDb
Imports System.IO

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
            conexionDB.Dispose()
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
            conexionDB.Close()
            conexionDB.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | ObtenerIDCliente")
        End Try
    End Function
    Public Function RegistrarRuta(TxTOrigen As TextBox, CmbDestino As ComboBox, TxTKilometros As TextBox, NDHoras As NumericUpDown, NDMinutos As NumericUpDown, TxTTOKA As TextBox, TxTFegali As TextBox, LIDCliente As Label, Window As Form, P_NewRuta As Panel)
        Try
            If TxTOrigen.Text <> String.Empty And CmbDestino.Text <> String.Empty And TxTKilometros.Text <> String.Empty And NDHoras.Value.ToString <> String.Empty And NDMinutos.Value.ToString <> String.Empty Then
                Dim TiempoTrayecto As String = NDHoras.Value.ToString & " H " & NDMinutos.Value.ToString & " MIN"
                Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
                Dim consulta = "INSERT INTO Rutas(Origen, Kilometros, Tiempo_Trayecto, TOKA, Fegali, Cliente_ID) VALUES('" & TxTOrigen.Text & "', '" & TxTKilometros.Text & "', '" & TiempoTrayecto & "', '" & TxTTOKA.Text & "', '" & TxTFegali.Text & "', '" & LIDCliente.Text & "')"
                conexionDB.Open()
                Dim comando As OleDbCommand = New OleDbCommand(consulta, conexionDB)
                comando.ExecuteNonQuery()
                MsgBox("Ruta Registrada", MsgBoxStyle.Information, "Exito | Corporativo LUIN")
                TxTOrigen.Text = "Corporativo LUIN"
                TxTKilometros.Text = ""
                NDHoras.Value = 0
                NDMinutos.Value = 0
                TxTTOKA.Text = ""
                TxTFegali.Text = ""
                Window.Close()
                conexionDB.Close()
                conexionDB.Dispose()
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
            conexionDB.Close()
            conexionDB.Dispose()
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
            conexionDB.Close()
            conexionDB.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | ObtenerIDClienteUp")
        End Try
    End Function
    Public Function ActualizarRuta(TxTOrigenUp As TextBox, LDestinoUp As Label, TxTKilometrosUp As TextBox, NDHorasUp As NumericUpDown, NDMinutosUp As NumericUpDown, TxTTOKAUp As TextBox, TxTFegaliUp As TextBox, LIDClienteUp As Label, Window As Form, LUpIDRuta As Label, P_UpRuta As Panel)
        Try
            If TxTOrigenUp.Text <> String.Empty And LDestinoUp.Text <> String.Empty And TxTKilometrosUp.Text <> String.Empty And NDHorasUp.Value.ToString <> String.Empty And NDMinutosUp.Value.ToString <> String.Empty Then
                Dim TiempoTrayecto As String = NDHorasUp.Value.ToString & " H " & NDMinutosUp.Value.ToString & " MIN"
                Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
                Dim consulta = "UPDATE Rutas SET Origen='" & TxTOrigenUp.Text & "', Kilometros='" & TxTKilometrosUp.Text & "', Tiempo_Trayecto='" & TiempoTrayecto & "', TOKA='" & TxTTOKAUp.Text & "', Fegali='" & TxTFegaliUp.Text & "', Cliente_ID=" & LIDClienteUp.Text & " WHERE IdRuta = " & LUpIDRuta.Text
                conexionDB.Open()
                Dim comando As OleDbCommand = New OleDbCommand(consulta, conexionDB)
                comando.ExecuteNonQuery()
                MsgBox("Ruta Actualizada", MsgBoxStyle.Information, "Exito | Corporativo LUIN")
                TxTOrigenUp.Text = "Corporativo LUIN"
                TxTKilometrosUp.Text = ""
                NDHorasUp.Value = 0
                NDMinutosUp.Value = 0
                TxTTOKAUp.Text = ""
                TxTFegaliUp.Text = ""
                Window.Close()
                conexionDB.Close()
                conexionDB.Dispose()
                P_UpRuta.Location = New Point(260, 681)
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
            Dim consulta As String = "SELECT  FORMAT(FechaRuta, 'dd/MM/yyyy') FROM Bitacoras" &
                                    " GROUP BY FORMAT(FechaRuta, 'dd/MM/yyyy');"
            Dim comando As OleDbCommand = New OleDbCommand(consulta)
            comando.Connection = conexionDB
            conexionDB.Open()
            Dim reader As OleDbDataReader = comando.ExecuteReader
            While reader.Read
                Componente.AddBoldedDate(reader.GetString(0))
            End While
            conexionDB.Close()
            conexionDB.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | ShowDatePDF")
        End Try
    End Function
    Public Function ShowHoursPDF(Componente As Object)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT Cliente, FORMAT(FechaRuta, 'hh:nn:ss am/pm') FROM Bitacoras" &
                                    " WHERE FORMAT(FechaRuta, 'dd/MM/yyyy') = '" & Componente.SelectionStart & "'" &
                                    " GROUP BY Cliente, FORMAT(FechaRuta, 'hh:nn:ss am/pm');"
            conexionDB.Open()
            Dim adap As OleDbDataAdapter = New OleDbDataAdapter(consulta, conexionDB)
            Dim dsDatos As DataTable = New DataTable()
            adap.Fill(dsDatos)
            Return dsDatos
            conexionDB.Close()
            conexionDB.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | ShowHoursPDF")
        End Try
    End Function
    Public Function ConvertToPDFStream(BinaryStr As Stream) As Byte()
        Dim Bytes(BinaryStr.Length) As Byte
        BinaryStr.Read(Bytes, 0, BinaryStr.Length) 'Leo el archivo y lo convierto a binario
        Return Bytes
        BinaryStr.Close() 'Cierro el FileStream
        BinaryStr.Dispose()
    End Function
    Public Function GetLastPDF(NombreCliente As String, FechaRuta As String, SFDRutaPDF As SaveFileDialog)
        Try
            Dim Nombre As String = Nothing
            Dim PDF As Byte() = Nothing

            Dim ConexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
            Dim Consulta As String = "SELECT Nombre, PDF FROM Bitacoras " &
                                    "WHERE Cliente = @Cliente " &
                                    "AND FechaRuta = @FechaRuta"
            Dim Comando As OleDbCommand = New OleDbCommand(Consulta, ConexionDB)
            Comando.Parameters.AddWithValue("@Cliente", OleDbType.VarChar).Value = NombreCliente
            Comando.Parameters.AddWithValue("@FechaRegistro", OleDbType.VarChar).Value = FechaRuta
            ConexionDB.Open()
            Dim Reader As OleDbDataReader = Comando.ExecuteReader
            If Reader.Read Then
                Nombre = Reader.GetString(0)
                PDF = ConvertToPDFStream(Reader.GetStream(1))
            End If
            SFDRutaPDF.FileName = "COPY_" & NombreCliente & "_" & Format(CDate(FechaRuta), "ddMMyy")
            If SFDRutaPDF.ShowDialog = DialogResult.OK Then
                Dim RutaArchivo As String = SFDRutaPDF.FileName
                File.WriteAllBytes(RutaArchivo, PDF)
            End If
            ConexionDB.Close()
            ConexionDB.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR | Corporativo LUIN | OBTENER")
        End Try
    End Function
#End Region
End Class

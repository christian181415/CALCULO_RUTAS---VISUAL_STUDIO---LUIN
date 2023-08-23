Imports System.ComponentModel
Imports System.Configuration
Imports System.Data.OleDb

Public Class ClassCasetasData
#Region "---------------------------------------------------------------CONEXION A DB----------------------------------------------------------------"
    Dim CadenaConexion As String = ConfigurationManager.ConnectionStrings("ConexionDB").ConnectionString
#End Region



#Region "---------------------------------------------------------------PANEL CONFIGURACION (RUTAS - CASETAS)----------------------------------------------------------------"
#Region "---------------------------------------------------------------ACCIONES REGISTER----------------------------------------------------------------"
    Public Function MostrarDestinosCR(CmbCRDetino As ComboBox)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT * FROM Clientes " '&
            '"INNER JOIN Rutas ON Rutas.Cliente_ID = Clientes.IdCliente " &
            '"WHERE IdRuta NOT IN (SELECT Ruta_ID FROM InfoRutas);"
            Dim adap As OleDbDataAdapter = New OleDbDataAdapter(consulta, conexionDB)
            Dim dsDatos As DataTable = New DataTable()
            conexionDB.Open()
            adap.Fill(dsDatos)
            conexionDB.Close()
            conexionDB.Dispose()
            Return dsDatos
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | MostrarDestinosCR")
        End Try
    End Function
    Public Function MostrarVehiculosCR(CmbCRVehiculo As ComboBox, LIDRutaC As Label)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            'Dim consulta As String = "SELECT * FROM Unidades"
            Dim consulta As String = "SELECT * FROM Unidades AS U" &
                                    " WHERE NOT EXISTS (" &
                                    " SELECT * FROM InfoRutas AS IR" &
                                    " INNER JOIN Rutas AS R ON R.IdRuta = IR.Ruta_ID" &
                                    " WHERE IR.Unidad_ID = U.IdUnidad" &
                                    " AND IR.Ruta_ID = " & LIDRutaC.Text & ")"
            Dim adap As OleDbDataAdapter = New OleDbDataAdapter(consulta, conexionDB)
            Dim dsDatos As DataTable = New DataTable()
            conexionDB.Open()
            adap.Fill(dsDatos)
            conexionDB.Close()
            conexionDB.Dispose()
            Return dsDatos
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | MostrarVehiculosCR")
        End Try
    End Function
    Public Function MostrarCasetasCR(DTGCasetaExists As DataGridView)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT Caseta FROM Casetas"
            Dim adap As OleDbDataAdapter = New OleDbDataAdapter(consulta, conexionDB)
            Dim dsDatos As DataTable = New DataTable()
            conexionDB.Open()
            adap.Fill(dsDatos)
            conexionDB.Close()
            conexionDB.Dispose()
            Return dsDatos
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | MostrarCasetasCR")
        End Try
    End Function

    Public Function ObtenerIDDestinoCR(CmbCRDestino As ComboBox, LIDDestino As Label)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT IdRuta FROM Clientes INNER JOIN Rutas ON Clientes.IdCliente = Rutas.Cliente_ID WHERE Nombre = '" & CmbCRDestino.Text & "' AND Status = True ;"
            Dim comando As OleDbCommand = New OleDbCommand(consulta)
            comando.Connection = conexionDB
            conexionDB.Open()
            Dim reader As OleDbDataReader = comando.ExecuteReader
            While reader.Read
                LIDDestino.Text = reader.GetInt32(0)
            End While
            conexionDB.Close()
            conexionDB.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | ObtenerIDCliente")
        End Try
    End Function
    Public Function ObtenerIDVehiculoCR(CmbCRVehiculo As ComboBox, LIDVehiculo As Label)
        Dim conexionDB As New OleDbConnection(CadenaConexion)
        Try
            Dim consulta As String = "SELECT IdUnidad FROM Unidades WHERE Vehiculo = '" & CmbCRVehiculo.Text & "';"
            Dim comando As OleDbCommand = New OleDbCommand(consulta)
            comando.Connection = conexionDB
            conexionDB.Open()
            Dim reader As OleDbDataReader = comando.ExecuteReader
            While reader.Read
                LIDVehiculo.Text = reader.GetInt32(0)
            End While
            conexionDB.Close()
            conexionDB.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | ObtenerIDCliente")
        End Try
    End Function
    Public Function RegistrarCasetaCR(CmbCRDestino As ComboBox, CmbCRVehiculo As ComboBox, DTGCasetaSelect As DataGridView, Window As Form, LImporte As Label, LIDRuta As Label, LIDCaseta As Label, LIDVehiculo As Label, P_CasetaRuta As Panel)
        Try
            For Fila As Integer = 0 To DTGCasetaSelect.Rows.Count - 1
                If CmbCRDestino.Text <> String.Empty And CmbCRVehiculo.Text <> String.Empty And DTGCasetaSelect.Rows(Fila).Cells(1).Value IsNot Nothing Then
                    Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
                    Dim consultaIDRuta As String = "SELECT IdCaseta FROM Casetas WHERE Caseta = '" & DTGCasetaSelect.Rows(Fila).Cells(0).Value & "';"
                    Dim comandoIDRuta As OleDbCommand = New OleDbCommand(consultaIDRuta)
                    comandoIDRuta.Connection = conexionDB
                    conexionDB.Open()
                    Dim reader As OleDbDataReader = comandoIDRuta.ExecuteReader
                    While reader.Read
                        LIDCaseta.Text = reader.GetInt32(0)
                        LImporte.Text = DTGCasetaSelect.Rows(Fila).Cells(1).Value
                    End While
                    conexionDB.Close()
                    conexionDB.Dispose()
                Else
                    If DTGCasetaSelect.Rows(Fila).Cells(1).Value Is Nothing Then
                        MsgBox("Complete el importe de la celda: " & Fila, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
                    Else
                        MsgBox("Rellene todos los campos", MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
                    End If
                End If

                If CmbCRDestino.Text <> String.Empty And CmbCRVehiculo.Text <> String.Empty And DTGCasetaSelect.Rows(Fila).Cells(1).Value IsNot Nothing Then
                    Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
                    Dim consulta = "INSERT INTO InfoRutas(Importe_Caseta, Ruta_ID, Caseta_ID, Unidad_ID) VALUES('" & LImporte.Text & "', " & LIDRuta.Text & ", " & LIDCaseta.Text & ", " & LIDVehiculo.Text & ")"
                    conexionDB.Open()
                    Dim comando As OleDbCommand = New OleDbCommand(consulta, conexionDB)
                    comando.ExecuteNonQuery()
                    conexionDB.Close()
                    conexionDB.Dispose()
                    Window.Close()
                    Window.Dispose()
                    P_CasetaRuta.Location = New Point(726, 0)
                    MsgBox("Casetas asignadas", MsgBoxStyle.Information, "Exito | Corporativo LUIN")
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | RegistrarRuta")
        End Try
    End Function
#End Region



#Region "ACCIONES UPDATE"
    Public Function ObtenerDestino_Vehiculo(LIDRutaUp As Label, LNombreDestino As Label, LIDVehiculoUp As Label, LNombreVehiculo As Label)
        Try
            Dim conexionDB As New OleDbConnection(CadenaConexion)
            Dim consulta As String = "SELECT Nombre, Vehiculo FROM Unidades " &
                                    "INNER JOIN ((Clientes INNER JOIN Rutas ON Clientes.IdCliente = Rutas.Cliente_ID) INNER JOIN InfoRutas ON Rutas.IdRuta = InfoRutas.Ruta_ID) ON Unidades.IdUnidad = InfoRutas.Unidad_ID " &
                                    "WHERE Ruta_ID = " & LIDRutaUp.Text & " AND Unidad_ID = " & LIDVehiculoUp.Text & " AND Status = True " &
                                    "GROUP BY Nombre, Vehiculo;"
            Dim comando As OleDbCommand = New OleDbCommand(consulta)
            comando.Connection = conexionDB
            conexionDB.Open()
            Dim reader As OleDbDataReader = comando.ExecuteReader
            While reader.Read
                LNombreDestino.Text = reader.GetString(0)
                LNombreVehiculo.Text = reader.GetString(1)
            End While
            conexionDB.Close()
            conexionDB.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | ObtenerDestino")
        End Try
    End Function
    Public Function ObtenerCasetasRuta(LIDRutaUp As Label, LIDVehiculoUp As Label)
        Try
            Dim conexionDB As New OleDbConnection(CadenaConexion)
            Dim consulta As String = "SELECT (Caseta)AS Casetas, (Importe_Caseta)AS Importe FROM Casetas 
                                    INNER JOIN InfoRutas ON Casetas.IdCaseta = InfoRutas.Caseta_ID
                                    WHERE Ruta_ID = " & LIDRutaUp.Text & " AND Unidad_ID = " & LIDVehiculoUp.Text & ";"
            Dim comando As OleDbCommand = New OleDbCommand(consulta)
            comando.Connection = conexionDB
            conexionDB.Open()
            Dim DTAdapter As New OleDbDataAdapter(comando)
            Dim table As New DataTable
            DTAdapter.Fill(table)
            Return table
            conexionDB.Close()
            conexionDB.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | MostrarCasetasCR")
        End Try
    End Function
    Public Function ActualizarCaseta_Ruta(LIDRutaUp As Label, LIDVehiculoUp As Label, DTGCasetaSelectUp As DataGridView, LIDCasetaUp As Label, LImporteUp As Label, Window As Form)
        Try
            If LIDRutaUp.Text <> "LIDRutaUp" And LIDVehiculoUp.Text <> "LIDVehiculoUp" And LIDRutaUp.Text <> String.Empty And LIDVehiculoUp.Text <> String.Empty Then
                Dim conexionDB As New OleDbConnection(CadenaConexion)
                Dim consulta As String = "DELETE * FROM InfoRutas WHERE Ruta_ID = " & LIDRutaUp.Text & " AND Unidad_ID = " & LIDVehiculoUp.Text
                Dim comando As OleDbCommand = New OleDbCommand(consulta)
                comando.Connection = conexionDB
                conexionDB.Open()
                Dim reader As OleDbDataReader = comando.ExecuteReader
                conexionDB.Close()
                conexionDB.Dispose()
            End If

            For Fila As Integer = 0 To DTGCasetaSelectUp.Rows.Count - 1
                If LIDRutaUp.Text <> "LIDRutaUp" And LIDVehiculoUp.Text <> "LIDVehiculoUp" And LIDRutaUp.Text <> String.Empty And LIDVehiculoUp.Text <> String.Empty Then
                    Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
                    Dim consultaIDRuta As String = "SELECT IdCaseta FROM Casetas WHERE Caseta = '" & DTGCasetaSelectUp.Rows(Fila).Cells(0).Value & "';"
                    Dim comandoIDRuta As OleDbCommand = New OleDbCommand(consultaIDRuta)
                    comandoIDRuta.Connection = conexionDB
                    conexionDB.Open()
                    Dim reader As OleDbDataReader = comandoIDRuta.ExecuteReader
                    While reader.Read
                        LIDCasetaUp.Text = reader.GetInt32(0)
                        LImporteUp.Text = DTGCasetaSelectUp.Rows(Fila).Cells(1).Value
                    End While
                    conexionDB.Close()
                    conexionDB.Dispose()
                End If

                If LIDRutaUp.Text <> "LIDRutaUp" And LIDVehiculoUp.Text <> "LIDVehiculoUp" And LIDRutaUp.Text <> String.Empty And LIDVehiculoUp.Text <> String.Empty And LImporteUp.Text <> "LImporteUp" And LIDCasetaUp.Text <> "LIDCasetaUp" Then
                    Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
                    Dim consulta = "INSERT INTO InfoRutas(Importe_Caseta, Ruta_ID, Caseta_ID, Unidad_ID) VALUES('" & LImporteUp.Text & "', " & LIDRutaUp.Text & ", " & LIDCasetaUp.Text & ", " & LIDVehiculoUp.Text & ")"
                    conexionDB.Open()
                    Dim comando As OleDbCommand = New OleDbCommand(consulta, conexionDB)
                    comando.ExecuteNonQuery()
                    conexionDB.Close()
                    conexionDB.Dispose()
                    Window.Close()
                    Window.Dispose()
                Else
                    MsgBox("Asigne las casetas correspondientes a la ruta para poder continuar", MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
                End If
            Next
            MsgBox("Casetas actualizadas", MsgBoxStyle.Information, "Exito | Corporativo LUIN")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | ActualizarCaseta_Ruta")
        End Try
    End Function
#End Region
#End Region
End Class

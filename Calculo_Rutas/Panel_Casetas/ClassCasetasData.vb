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
            'Dim consulta As String = "SELECT * FROM Clientes " '&
            '"INNER JOIN Rutas ON Rutas.Cliente_ID = Clientes.IdCliente " &
            '"WHERE IdRuta NOT IN (SELECT Ruta_ID FROM InfoRutas);"
            Dim consulta As String = "SELECT Nombre " &
                                    "FROM Clientes INNER JOIN Rutas ON Clientes.IdCliente = Rutas.Cliente_ID " &
                                    "WHERE Status = True " &
                                    "GROUP BY Nombre;"
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
    Public Function RegistrarCasetaCR(CmbCRDestino As ComboBox, CmbCRVehiculo As ComboBox, DTGCasetaSelect As DataGridView, Window As Form, LIDRuta As Label, LIDVehiculo As Label, P_CasetaRuta As Panel)
        Try
            Dim Limite As Integer = DTGCasetaSelect.Rows.Count - 1
            Dim DTGComplete As Boolean = Nothing
            If CmbCRDestino.Text <> Nothing And CmbCRVehiculo.Text <> Nothing And Limite > -1 Then
                For Fila As Integer = 0 To DTGCasetaSelect.Rows.Count - 1
                    If DTGCasetaSelect.Rows(Fila).Cells(1).Value IsNot Nothing Then
                        DTGComplete = True
                    Else
                        DTGComplete = False
                        MsgBox("Complete los importes faltantes", MsgBoxStyle.Exclamation, "CAMPOS FALTANTES | Corporativo LUIN")
                        Exit For
                    End If
                Next

                If DTGComplete = True Then
                    Dim IDCaseta As Integer
                    Dim Importe As Integer
                    Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
                    conexionDB.Open()
                    For Fila As Integer = 0 To DTGCasetaSelect.Rows.Count - 1
                        Dim consultaIDRuta As String = "SELECT IdCaseta FROM Casetas WHERE Caseta = '" & DTGCasetaSelect.Rows(Fila).Cells(0).Value & "';"
                        Dim comandoIDRuta As OleDbCommand = New OleDbCommand(consultaIDRuta, conexionDB)
                        Dim reader As OleDbDataReader = comandoIDRuta.ExecuteReader
                        While reader.Read
                            IDCaseta = reader.GetInt32(0)
                            Importe = DTGCasetaSelect.Rows(Fila).Cells(1).Value
                        End While


                        Dim consulta = "INSERT INTO InfoRutas(Importe_Caseta, Ruta_ID, Caseta_ID, Unidad_ID) VALUES('" & Importe & "', " & LIDRuta.Text & ", " & IDCaseta & ", " & LIDVehiculo.Text & ")"
                        Dim comando As OleDbCommand = New OleDbCommand(consulta, conexionDB)
                        comando.ExecuteNonQuery()
                    Next
                    P_CasetaRuta.Location = New Point(726, 0)
                    conexionDB.Close()
                    conexionDB.Dispose()
                    Window.Close()
                    MsgBox("Casetas asignadas", MsgBoxStyle.Information, "EXITO | Corporativo LUIN")
                    WinPrincipal.Opacity = 1
                End If
            Else
                MsgBox("Para poder asignar casetas." & Chr(10) & "• Seleccione un destino" & Chr(10) & "• Seleccione un vehiculo" & Chr(10) & "• Asigne las casetas", MsgBoxStyle.Exclamation, "CAMPOS FALTANTES | Corporativo LUIN")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR | Corporativo LUIN")
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
    Public Function ActualizarCaseta_Ruta(LIDRutaUp As Label, LIDVehiculoUp As Label, DTGCasetaSelectUp As DataGridView, Window As Form, P_UpCasetaRuta As Panel)
        'Try
        '    If LIDRutaUp.Text <> "LIDRutaUp" And LIDVehiculoUp.Text <> "LIDVehiculoUp" And LIDRutaUp.Text <> String.Empty And LIDVehiculoUp.Text <> String.Empty Then
        '        Dim conexionDB As New OleDbConnection(CadenaConexion)
        '        Dim consulta As String = "DELETE * FROM InfoRutas WHERE Ruta_ID = " & LIDRutaUp.Text & " AND Unidad_ID = " & LIDVehiculoUp.Text
        '        Dim comando As OleDbCommand = New OleDbCommand(consulta)
        '        comando.Connection = conexionDB
        '        conexionDB.Open()
        '        Dim reader As OleDbDataReader = comando.ExecuteReader
        '        conexionDB.Close()
        '        conexionDB.Dispose()
        '    End If

        '    For Fila As Integer = 0 To DTGCasetaSelectUp.Rows.Count - 1
        '        If LIDRutaUp.Text <> "LIDRutaUp" And LIDVehiculoUp.Text <> "LIDVehiculoUp" And LIDRutaUp.Text <> String.Empty And LIDVehiculoUp.Text <> String.Empty Then
        '            Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
        '            Dim consultaIDRuta As String = "SELECT IdCaseta FROM Casetas WHERE Caseta = '" & DTGCasetaSelectUp.Rows(Fila).Cells(0).Value & "';"
        '            Dim comandoIDRuta As OleDbCommand = New OleDbCommand(consultaIDRuta)
        '            comandoIDRuta.Connection = conexionDB
        '            conexionDB.Open()
        '            Dim reader As OleDbDataReader = comandoIDRuta.ExecuteReader
        '            While reader.Read
        '                LIDCasetaUp.Text = reader.GetInt32(0)
        '                LImporteUp.Text = DTGCasetaSelectUp.Rows(Fila).Cells(1).Value
        '            End While
        '            conexionDB.Close()
        '            conexionDB.Dispose()
        '        End If

        '        If LIDRutaUp.Text <> "LIDRutaUp" And LIDVehiculoUp.Text <> "LIDVehiculoUp" And LIDRutaUp.Text <> String.Empty And LIDVehiculoUp.Text <> String.Empty And LImporteUp.Text <> "LImporteUp" And LIDCasetaUp.Text <> "LIDCasetaUp" Then
        '            Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
        '            Dim consulta = "INSERT INTO InfoRutas(Importe_Caseta, Ruta_ID, Caseta_ID, Unidad_ID) VALUES('" & LImporteUp.Text & "', " & LIDRutaUp.Text & ", " & LIDCasetaUp.Text & ", " & LIDVehiculoUp.Text & ")"
        '            conexionDB.Open()
        '            Dim comando As OleDbCommand = New OleDbCommand(consulta, conexionDB)
        '            comando.ExecuteNonQuery()
        '            conexionDB.Close()
        '            conexionDB.Dispose()
        '            Window.Close()
        '        Else
        '            MsgBox("Asigne las casetas correspondientes a la ruta para poder continuar", MsgBoxStyle.Exclamation, "Error | Corporativo LUIN")
        '        End If
        '    Next
        '    MsgBox("Casetas actualizadas", MsgBoxStyle.Information, "Exito | Corporativo LUIN")
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error | Corporativo LUIN | ActualizarCaseta_Ruta")
        'End Try



        Try
            Dim Limite As Integer = DTGCasetaSelectUp.Rows.Count - 1
            Dim DTGComplete As Boolean = Nothing
            If LIDRutaUp.Text <> Nothing And LIDRutaUp.Text <> "LIDRutaUp" And LIDVehiculoUp.Text <> Nothing And LIDVehiculoUp.Text <> "LIDVehiculoUp" And Limite > -1 Then
                For Fila As Integer = 0 To DTGCasetaSelectUp.Rows.Count - 1
                    If DTGCasetaSelectUp.Rows(Fila).Cells(1).Value.ToString <> String.Empty Then
                        DTGComplete = True
                    Else
                        DTGComplete = False
                        MsgBox("Complete los importes faltantes", MsgBoxStyle.Exclamation, "CAMPOS FALTANTES | Corporativo LUIN")
                        Exit For
                    End If
                Next

                If DTGComplete = True Then
                    Dim IDCaseta As Integer
                    Dim Importe As Integer
                    Dim conexionDB As OleDbConnection = New OleDbConnection(CadenaConexion)
                    Dim consultaDelete As String = "DELETE * FROM InfoRutas WHERE Ruta_ID = " & LIDRutaUp.Text & " AND Unidad_ID = " & LIDVehiculoUp.Text
                    Dim comandoDelete As OleDbCommand = New OleDbCommand(consultaDelete, conexionDB)
                    conexionDB.Open()
                    comandoDelete.ExecuteReader()

                    For Fila As Integer = 0 To DTGCasetaSelectUp.Rows.Count - 1
                        Dim consultaIDRuta As String = "SELECT IdCaseta FROM Casetas WHERE Caseta = '" & DTGCasetaSelectUp.Rows(Fila).Cells(0).Value & "';"
                        Dim comandoIDRuta As OleDbCommand = New OleDbCommand(consultaIDRuta, conexionDB)
                        Dim reader As OleDbDataReader = comandoIDRuta.ExecuteReader
                        While reader.Read
                            IDCaseta = reader.GetInt32(0)
                            Importe = DTGCasetaSelectUp.Rows(Fila).Cells(1).Value
                        End While


                        Dim consulta = "INSERT INTO InfoRutas(Importe_Caseta, Ruta_ID, Caseta_ID, Unidad_ID) VALUES('" & Importe & "', " & LIDRutaUp.Text & ", " & IDCaseta & ", " & LIDVehiculoUp.Text & ")"
                        Dim comando As OleDbCommand = New OleDbCommand(consulta, conexionDB)
                        comando.ExecuteNonQuery()
                    Next
                    P_UpCasetaRuta.Location = New Point(726, 470)
                    conexionDB.Close()
                    conexionDB.Dispose()
                    Window.Close()
                    MsgBox("Casetas actualizadas", MsgBoxStyle.Information, "EXITO | Corporativo LUIN")
                    WinPrincipal.Opacity = 1
                End If
            Else
                MsgBox("Debe asignar casetas", MsgBoxStyle.Exclamation, "CAMPOS FALTANTES | Corporativo LUIN")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR | Corporativo LUIN")
        End Try
    End Function
#End Region
#End Region
End Class

Imports System.Configuration
Imports System.Data.OleDb

Public Class WinDropDB
#Region "---------------------------------------------------------------CONEXION A DB----------------------------------------------------------------"
    Dim CadenaConexion As String = ConfigurationManager.ConnectionStrings("ConexionDB").ConnectionString
#End Region

    Private Sub BtnEliminar_MouseHover(sender As Object, e As EventArgs) Handles BtnEliminar.MouseHover
        BtnEliminar.ForeColor = Color.White
    End Sub

    Private Sub BtnEliminar_MouseLeave(sender As Object, e As EventArgs) Handles BtnEliminar.MouseLeave
        BtnEliminar.ForeColor = Color.Red
    End Sub
    Private Sub BtnCCloseUp_Click(sender As Object, e As EventArgs) Handles BtnCCloseUp.Click
        Me.Close()
        WinPrincipal.Opacity = 1
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click
        Try
            If txtPassword.Text = "@SITLiberateDBLu1n." Then
                Dim conexionDB As New OleDbConnection(CadenaConexion)
                Dim consulta As String = "DELETE * FROM Bitacoras"
                Dim comando As OleDbCommand = New OleDbCommand(consulta)
                comando.Connection = conexionDB
                conexionDB.Open()
                Dim reader As OleDbDataReader = comando.ExecuteReader
                conexionDB.Close()
                conexionDB.Dispose()
                MsgBox("Información eliminada correctamente.", MsgBoxStyle.Information, "EXITO | Corporativo LUIN")
            ElseIf txtPassword.Text = Nothing Then
                MsgBox("Introdusca su contraseña.", MsgBoxStyle.Information, "Corporativo LUIN")
            Else
                MsgBox("Su contraseña no es valida.", MsgBoxStyle.Exclamation, "Corporativo LUIN")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "ERROR | Corporativo LUIN")
        End Try
    End Sub
End Class
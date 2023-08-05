Public MustInherit Class ClassToolTip
    Public Shared Function Show(Tiempo As String, picture As String, text As String, VWidth As Integer, VHeight As Integer)
        Try
            Dim Form As New WinToolTip
            Dim Ancho, Alto As Integer

            If picture = Nothing Then
                Form.LWidth.Text = VWidth
                Form.LHeigth.Text = VHeight
                Form.LMilisegundos.Text = Tiempo * 1000
                Form.PBoxPicture.Visible = False
                Form.LText.Text = text
                Ancho = Form.LText.Width
                Alto = Form.PTop.Height + Form.LText.Height
                Form.Size = New Size(Ancho, Alto)
                Form.ShowDialog()
            Else

                Dim RutaImagen As System.Drawing.Bitmap = Bitmap.FromFile(System.AppDomain.CurrentDomain.BaseDirectory() + "\Assets\IMG\ToolTip\" & picture)
                Form.LWidth.Text = VWidth
                Form.LHeigth.Text = VHeight
                Form.LMilisegundos.Text = Tiempo * 1000
                Form.PBoxPicture.BackgroundImage = RutaImagen
                Form.LText.Text = text
                Ancho = Form.LText.Width
                Alto = Form.PTop.Height + Form.PBoxPicture.Height + Form.LText.Height
                Form.Size = New Size(Ancho, Alto)
                Form.ShowDialog()
            End If


        Catch ex As Exception
            MsgBox("Error al mostrar el ToolTip" & Chr(10) & ex.Message, MsgBoxStyle.Critical, "Error | Corporativo LUIN")
        End Try
    End Function
End Class

Public Class WinToolTip
    Dim Contador As Integer = 0

    Private Sub WinToolTip_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim r As Rectangle = My.Computer.Screen.WorkingArea
        Dim Largo = (r.Width / 2) - LWidth.Text
        Dim Alto = (r.Height / 1) - LHeigth.Text
        Location = New Point(Largo, Alto)

        Timer1.Interval = LMilisegundos.Text
        Timer1.Start()
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        While Contador < 100
            Contador = Contador + 1
            Me.Close()
        End While
    End Sub
End Class
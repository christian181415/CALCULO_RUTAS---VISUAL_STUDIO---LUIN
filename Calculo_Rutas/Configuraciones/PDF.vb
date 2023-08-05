Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf

Public Class PDF
    Public Function GenerarPDF(SFDialogPDF As SaveFileDialog, TXTFecha As String, CMBCliente As String, CMBVehiculo As String, L_Ruta_Destino As String, TXT_Notas As String, Total_Combustible As String, LEfectivoTotal As String, Total_Casetas As String, L_Desgloce_Casetas As ListBox, LKilometrosPDF As String, LTiempoTrayectoPDF As String)
        Try
            If SFDialogPDF.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                Dim oDoc As New iTextSharp.text.Document(PageSize.LETTER, 55, 60, 305, 0)
                Dim pdfw As iTextSharp.text.pdf.PdfWriter
                Dim cb As PdfContentByte
                Dim fuente As iTextSharp.text.pdf.BaseFont
                'DATOS PARA EL OpenFileDialog PDF
                SFDialogPDF.Title = "Escriba el nombre del archivo a guardar."
                SFDialogPDF.Filter = "Archivos PDF (*.pdf)|*.pdf|Todos los archivos (*.*)|*.*"
                Dim NombreArchivo As String = SFDialogPDF.FileName


                Try
                    pdfw = PdfWriter.GetInstance(oDoc, New FileStream(NombreArchivo,
                    FileMode.Create, FileAccess.Write, FileShare.None))
                    oDoc.Open()
                    cb = pdfw.DirectContent
                    oDoc.NewPage()
                    cb.BeginText()
                    fuente = FontFactory.GetFont(FontFactory.HELVETICA, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL).BaseFont
                    Dim FontStandar As BaseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1250, True)
                    Dim TxTNormal As Font = New Font(FontStandar, 10.0F, Font.NORMAL, BaseColor.BLACK)
                    cb.SetFontAndSize(fuente, 10)
                    cb.SetColorFill(iTextSharp.text.BaseColor.BLACK)

                    Dim logo As String = System.AppDomain.CurrentDomain.BaseDirectory() + "\Assets\PDF\logo.jpg"
                    Dim img As Image = Image.GetInstance(logo)
                    img.Alignment = iTextSharp.text.Image.ALIGN_LEFT
                    img.ScalePercent(40)
                    img.SetAbsolutePosition(50, 690)
                    cb.AddImage(img)


                    cb.ShowTextAligned(Element.ALIGN_LEFT, "FORMATO DE GASTOS", 240, 725, 0)
                    cb.ShowTextAligned(Element.ALIGN_LEFT, "RUTAS FORÁNEAS", 250, 705, 0)
                    cb.ShowTextAligned(Element.ALIGN_RIGHT, "Fecha", 555, 720, 0)
                    cb.ShowTextAligned(Element.ALIGN_RIGHT, TXTFecha, 555, 710, 0)

                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Registro de kilometraje", 50, 670, 0)
                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Nombre:", 50, 650, 0)
                    cb.ShowTextAligned(Element.ALIGN_CENTER, CMBCliente, 235, 651, 0)
                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Vehículo:", 50, 630, 0)
                    cb.ShowTextAligned(Element.ALIGN_CENTER, CMBVehiculo, 235, 631, 0)

                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Fecha y Hora Salida: ", 350, 650, 0)
                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Fecha y Hora Llegada: ", 350, 630, 0)

                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Destino: ", 50, 595, 0)
                    cb.ShowTextAligned(Element.ALIGN_CENTER, L_Ruta_Destino, 350, 596, 0)
                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Kilometros: ", 50, 575, 0)
                    cb.ShowTextAligned(Element.ALIGN_CENTER, LKilometrosPDF & " Km", 215, 576, 0)
                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Tiempo trayecto: ", 50, 555, 0)
                    cb.ShowTextAligned(Element.ALIGN_CENTER, LTiempoTrayectoPDF, 215, 556, 0)
                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Kilometraje inicial: ", 50, 535, 0)
                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Kilometraje final: ", 50, 515, 0)
                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Litros fegali: ", 50, 495, 0)
                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Importe fegali: ", 50, 475, 0)
                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Litros TOKA: ", 50, 455, 0)
                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Importe TOKA: ", 50, 435, 0)

                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Efectivo total: ", 310, 575, 0)
                    cb.ShowTextAligned(Element.ALIGN_CENTER, "$ " & LEfectivoTotal, 510, 576, 0)
                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Combustible Proyectado: ", 310, 555, 0)
                    cb.ShowTextAligned(Element.ALIGN_CENTER, "$ " & Total_Combustible, 510, 556, 0)
                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Costo Casetas: ", 310, 535, 0)
                    cb.ShowTextAligned(Element.ALIGN_CENTER, "$ " & Total_Casetas, 510, 536, 0)
                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Importes ", 315, 515, 0)
                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Casetas ", 430, 515, 0)

                    Dim nextData As Integer = 500
                    For Each item As DataRowView In L_Desgloce_Casetas.Items
                        Dim row As DataRow = item.Row
                        For n As Integer = 0 To 0
                            cb.ShowTextAligned(Element.ALIGN_LEFT, CStr(row(0)), 315, nextData, 0)
                            nextData = nextData - 15
                        Next
                    Next

                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Incidencias", 50, 380, 0)
                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Detalle y Requerimientos del viaje:", 55, 360, 0)
                    'cb.ShowTextAligned(Element.ALIGN_LEFT, TXT_Notas, 55, 340, 0)


                    oDoc.Add(New Phrase(""))
                    For cont As Integer = 0 To 7
                        oDoc.Add(Chunk.NEWLINE)
                    Next

                    Dim Parrafo As Paragraph = New Paragraph(TXT_Notas, TxTNormal)
                    Parrafo.Alignment = Element.ALIGN_LEFT
                    oDoc.Add(Parrafo)

                    cb.ShowTextAligned(Element.ALIGN_LEFT, "Gastos en alimentos", 50, 292, 0)


                    'LINEA DE NOMBRE
                    cb.MoveTo(140, 649)
                    cb.LineTo(330, 649)
                    cb.ClosePathStroke()
                    'LINEA DE VEHICULO
                    cb.MoveTo(140, 629)
                    cb.LineTo(330, 629)
                    cb.ClosePathStroke()
                    'LINEA DE FECHA SALIDA
                    cb.MoveTo(465, 649)
                    cb.LineTo(555, 649)
                    cb.ClosePathStroke()
                    'LINEA DE FECHA LLEGADA
                    cb.MoveTo(465, 629)
                    cb.LineTo(555, 629)
                    cb.ClosePathStroke()
                    'LINEA DE RUTA DESTINO
                    cb.MoveTo(140, 594)
                    cb.LineTo(555, 594)
                    cb.ClosePathStroke()
                    'LINEA DE KILOMETROS
                    cb.MoveTo(140, 574)
                    cb.LineTo(290, 574)
                    cb.ClosePathStroke()
                    'LINEA DE TIEMPO TRAYECTO
                    cb.MoveTo(140, 554)
                    cb.LineTo(290, 554)
                    cb.ClosePathStroke()
                    'LINEA DE KILOMETRO INICIAL
                    cb.MoveTo(140, 534)
                    cb.LineTo(290, 534)
                    cb.ClosePathStroke()
                    'LINEA DE KILOMETRO FINAL
                    cb.MoveTo(140, 514)
                    cb.LineTo(290, 514)
                    cb.ClosePathStroke()
                    'LINEA DE LITROS FEGALI
                    cb.MoveTo(140, 494)
                    cb.LineTo(290, 494)
                    cb.ClosePathStroke()
                    'LINEA DE IMPORTE FEGALI
                    cb.MoveTo(140, 474)
                    cb.LineTo(290, 474)
                    cb.ClosePathStroke()
                    'LINEA DE LITROS TOKA
                    cb.MoveTo(140, 454)
                    cb.LineTo(290, 454)
                    cb.ClosePathStroke()
                    'LINEA DE IMPORTE TOKA
                    cb.MoveTo(140, 434)
                    cb.LineTo(290, 434)
                    cb.ClosePathStroke()
                    'LINEA DE EFECTIVO
                    cb.MoveTo(465, 574)
                    cb.LineTo(555, 574)
                    cb.ClosePathStroke()
                    'LINEA DE COMBUSTIBLE PROYECTADO
                    cb.MoveTo(465, 554)
                    cb.LineTo(555, 554)
                    cb.ClosePathStroke()
                    'LINEA DE IMPORTE COSTO CASETAS
                    cb.MoveTo(465, 534)
                    cb.LineTo(555, 534)
                    cb.ClosePathStroke()
                    'LINEA DE NOMBRES CASETAS
                    'UP
                    cb.MoveTo(310, 525)
                    cb.LineTo(555, 525)
                    cb.ClosePathStroke()
                    'DOWN
                    cb.MoveTo(310, 390)
                    cb.LineTo(555, 390)
                    cb.ClosePathStroke()
                    'LEFT
                    cb.MoveTo(310, 525)
                    cb.LineTo(310, 390)
                    cb.ClosePathStroke()
                    'RIGHT
                    cb.MoveTo(555, 525)
                    cb.LineTo(555, 390)
                    cb.ClosePathStroke()
                    'LINEA DE DETALLE Y REQUERIMIENTO
                    'UP
                    cb.MoveTo(50, 372)
                    cb.LineTo(555, 372)
                    cb.ClosePathStroke()
                    'DOWN
                    cb.MoveTo(50, 305)
                    cb.LineTo(555, 305)
                    cb.ClosePathStroke()
                    'LEFT
                    cb.MoveTo(50, 372)
                    cb.LineTo(50, 305)
                    cb.ClosePathStroke()
                    'RIGHT
                    cb.MoveTo(555, 372)
                    cb.LineTo(555, 305)
                    cb.ClosePathStroke()


                    'BORDES DE LA PAGINA
                    'LEFT
                    cb.MoveTo(40, 40)
                    cb.LineTo(40, 750)
                    cb.ClosePathStroke()
                    'UP
                    cb.MoveTo(40, 750)
                    cb.LineTo(570, 750)
                    cb.ClosePathStroke()
                    'RIGHT
                    cb.MoveTo(570, 750)
                    cb.LineTo(570, 40)
                    cb.ClosePathStroke()
                    'DOWN
                    cb.MoveTo(570, 40)
                    cb.LineTo(40, 40)
                    cb.ClosePathStroke()


                    'FORMATO DE RECIBO EN IMAGEN
                    Dim RPagologo As String = System.AppDomain.CurrentDomain.BaseDirectory() + "\Assets\PDF\RPagoLogo.jpg"
                    Dim RImg As Image = Image.GetInstance(RPagologo)
                    RImg.Alignment = iTextSharp.text.Image.ALIGN_LEFT
                    RImg.ScalePercent(23.7)
                    RImg.SetAbsolutePosition(50, 42)
                    cb.AddImage(RImg)

                    cb.EndText()
                    pdfw.Flush()
                    oDoc.Close()
                Catch ex As Exception
                    If File.Exists(NombreArchivo) Then
                        MsgBox(ex.Message)
                        If oDoc.IsOpen Then oDoc.Close()
                        File.Delete(NombreArchivo)
                    End If
                    Throw New Exception("Error al generar archivo PDF (" + ex.Message + ")")
                Finally
                    cb = Nothing
                    pdfw = Nothing
                    oDoc = Nothing
                End Try
            End If
        Catch ex As Exception

        End Try
    End Function
End Class

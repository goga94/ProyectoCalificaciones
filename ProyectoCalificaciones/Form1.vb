Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Drawing.Text
Imports System.Security.Cryptography
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO


Public Class Form1
    Private Sub btncalcular_Click(sender As Object, e As EventArgs) Handles btncalcular.Click
        Dim matricula, español, matematicas, biologia, historia, ingles, promedio As Decimal
        Dim nombre, estado As String

        matricula = txtmatricula.Text
        nombre = txtnombre.Text
        español = txtespañol.Text
        matematicas = txtmatematicas.Text
        biologia = txtbiologia.Text
        historia = txthistoria.Text
        ingles = txtingles.Text

        promedio = (español + matematicas + biologia + historia + ingles) / 5

        If promedio <= 5 Then
            estado = "REPROBADO"
            txtestado.Text = estado
            txtpromedio.Text = promedio
        ElseIf promedio >= 6 And promedio <= 10 Then
            estado = "APROBADO"
            txtestado.Text = estado
            txtpromedio.Text = promedio
        Else
            estado = "ERROR"
            txtestado.Text = estado
            txtpromedio.Text = promedio
        End If

    End Sub

    Private Sub btnlimpiar_Click(sender As Object, e As EventArgs) Handles btnlimpiar.Click
        txtmatricula.Clear()
        txtnombre.Clear()
        txtespañol.Clear()
        txtmatematicas.Clear()
        txtbiologia.Clear()
        txthistoria.Clear()
        txtingles.Clear()
        txtestado.Clear()
        txtpromedio.Clear()
    End Sub

    Private Sub btnsalir_Click(sender As Object, e As EventArgs) Handles btnsalir.Click
        If MsgBox("¿Estas seguro de salir del programa?", vbQuestion + vbYesNo, "Alerta") = vbYes Then
            End
        End If
    End Sub

    Private Sub btnguardar_Click(sender As Object, e As EventArgs) Handles btnguardar.Click
        Dim matricula As Integer = txtmatricula.Text
        Dim nombre As String = txtnombre.Text
        Dim español As Integer = txtespañol.Text
        Dim matematicas As Integer = txtmatematicas.Text
        Dim biologia As Integer = txtbiologia.Text
        Dim historia As Integer = txthistoria.Text
        Dim ingles As Integer = txtingles.Text
        Dim estado As String = txtestado.Text
        Dim promedio As Decimal = txtpromedio.Text

        Dim query As String = "Insert into Calificaciones values (@matricula, @nombre, @español, @matematicas, @biologia, @historia, @ingles, @estado, @promedio)"
        Using con As SqlConnection = New SqlConnection("Data Source=LAPTOP-U6IJCDNR;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False")
            Using cmd As SqlCommand = New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@matricula", matricula)
                cmd.Parameters.AddWithValue("@nombre", nombre)
                cmd.Parameters.AddWithValue("@español", español)
                cmd.Parameters.AddWithValue("@matematicas", matematicas)
                cmd.Parameters.AddWithValue("@biologia", biologia)
                cmd.Parameters.AddWithValue("@historia", historia)
                cmd.Parameters.AddWithValue("@ingles", ingles)
                cmd.Parameters.AddWithValue("@estado", estado)
                cmd.Parameters.AddWithValue("@promedio", promedio)
                con.Open()
                cmd.ExecuteNonQuery()
                con.Close()
                MessageBox.Show("Alumno registrado correctamente.")
            End Using

        End Using


    End Sub

    Private Sub btnimprimir_Click(sender As Object, e As EventArgs) Handles btnimprimir.Click
        Dim SaveFileDialog As New SaveFileDialog
        Dim Ruta As String
        With SaveFileDialog
            .Title = "Guardar"
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            .Filter = "Archivos PDF (*.pdf)|*.pdf"
            .FileName = "Boleta"
            .OverwritePrompt = True
            .CheckPathExists = True
        End With

        If SaveFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Ruta = SaveFileDialog.FileName

        Else
            Ruta = String.Empty
            Exit Sub
        End If

        Try
            Dim Document As New iTextSharp.text.Document(PageSize.LETTER)
            Document.PageSize.Rotate()
            Document.AddAuthor("Ing. Alexander")
            Document.AddTitle("Creacion de PDF")

            Dim writer As PdfWriter = PdfWriter.GetInstance(Document, New System.IO.FileStream(Ruta, System.IO.FileMode.Create))
            writer.ViewerPreferences = PdfWriter.PageLayoutSinglePage
            Document.Open()

            Dim matricula As PdfContentByte = writer.DirectContent
            Dim nombre As PdfContentByte = writer.DirectContent
            Dim español As PdfContentByte = writer.DirectContent
            Dim matematicas As PdfContentByte = writer.DirectContent
            Dim biologia As PdfContentByte = writer.DirectContent
            Dim historia As PdfContentByte = writer.DirectContent
            Dim ingles As PdfContentByte = writer.DirectContent
            Dim estado As PdfContentByte = writer.DirectContent
            Dim promedio As PdfContentByte = writer.DirectContent
            Dim bf As BaseFont = BaseFont.CreateFont(BaseFont.COURIER_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED)
            matricula.SetFontAndSize(bf, 20)
            matricula.BeginText()


            'datos
            matricula.SetTextMatrix(50, 790)
            matricula.ShowText("Matrícula:" & Me.txtmatricula.Text)
            nombre.SetTextMatrix(50, 890)
            nombre.ShowText("Nombre:" & Me.txtnombre.Text)
            español.SetTextMatrix(50, 990)
            español.ShowText("Español:" & Me.txtespañol.Text)
            matematicas.SetTextMatrix(50, 1090)
            matematicas.ShowText("Matemáticas:" & Me.txtmatematicas.Text)
            biologia.SetTextMatrix(50, 1190)
            biologia.ShowText("Biología:" & Me.txtbiologia.Text)
            historia.SetTextMatrix(50, 1290)
            historia.ShowText("Historia:" & Me.txthistoria.Text)
            ingles.SetTextMatrix(50, 1390)
            ingles.ShowText("Inglés:" & Me.txtingles.Text)
            estado.SetTextMatrix(50, 1490)
            estado.ShowText("Estadp:" & Me.txtestado.Text)
            promedio.SetTextMatrix(50, 1590)
            promedio.ShowText("Promedio:" & Me.txtpromedio.Text)

            matricula.EndText()
            nombre.EndText()
            español.EndText()
            matematicas.EndText()
            biologia.EndText()
            historia.EndText()
            ingles.EndText()
            estado.EndText()
            promedio.EndText()

            Document.Close()


        Catch ex As Exception
            MessageBox.Show("Error de creación", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
End Class

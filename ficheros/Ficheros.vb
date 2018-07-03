Imports System.IO
Imports System.Windows.Forms
Imports System.Drawing
Module Ficheros

    Sub escribeFichero(ByVal rutaFichero As String, ByVal texto As String, ByVal anadir As Boolean)

        Dim sw As New StreamWriter(rutaFichero, anadir)

        sw.WriteLine(texto)

        sw.Close()

    End Sub

    Function leerTodoElFichero(ByVal rutaFichero As String) As String

        Dim sr As New StreamReader(rutaFichero)

        Dim contenido As String = sr.ReadToEnd()

        sr.Close()

        Return contenido
    End Function

    Sub escribirFicheroDGV(ByVal dgv As DataGridView, ByVal rutaFichero As String)

        Dim sw As New StreamWriter(rutaFichero, True)

        For Each row As DataGridViewRow In dgv.Rows
            For Each cell As DataGridViewCell In row.Cells
                If Not String.IsNullOrEmpty(cell.Value) Then
                    ' Check if the string only consists of spaces
                    sw.WriteLine(cell.Value)

                End If

            Next
        Next

        sw.Close()

    End Sub

End Module

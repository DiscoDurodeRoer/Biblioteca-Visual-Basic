Imports System.Windows.Forms
Imports System.Drawing
Module ValidacionesFormularios


    'Incluye positivos y negativos
    Sub escribirSoloNumeros(ByVal t As TextBox, ByVal label As Label, ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Not IsNumeric(e.KeyChar) And Convert.ToInt32(e.KeyChar) <> 8 And Convert.ToInt32(e.KeyChar) <> 45 Then
            e.Handled = True
            label.Text = "Debes teclear valores numericos"
        Else
            label.Text = ""
        End If
    End Sub


    Sub escribirSoloNumerosPositivos(ByVal t As TextBox, ByVal label As Label, ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Not IsNumeric(e.KeyChar) And Convert.ToInt32(e.KeyChar) <> 8 Then
            e.Handled = True
            label.Text = "Debes teclear valores numericos"
        Else
            label.Text = ""
        End If
    End Sub

    Sub escribirSoloNumerosNegativos(ByVal t As TextBox, ByVal label As Label, ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Not IsNumeric(e.KeyChar) And Convert.ToInt32(e.KeyChar) <> 8 And Convert.ToInt32(e.KeyChar) <> 45 Then
            e.Handled = True
            label.Text = "Debes teclear valores numericos"
        Else
            label.Text = ""
        End If
    End Sub

    Sub escribirSoloLetras(ByVal t As TextBox, ByVal label As Label, ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Not Char.IsLetter(e.KeyChar) And Convert.ToInt32(e.KeyChar) <> 8 And Convert.ToInt32(e.KeyChar) <> 164 Then
            e.Handled = True
            label.Text = "Debes teclear letras"
        Else
            label.Text = ""

        End If
    End Sub

    Sub escribirSoloLetrasYespacios(ByVal t As TextBox, ByVal label As Label, ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Not Char.IsLetter(e.KeyChar) And Asc(e.KeyChar) <> 8 And Asc(e.KeyChar) <> 164 And Asc(e.KeyChar) <> 32 Then
            e.Handled = True
            label.Text = "Debes teclear letras o espacios"
        Else
            label.Text = ""

        End If
    End Sub

    Sub soloDoubles(ByVal txt As TextBox, ByVal etiquetaError As Label, ByVal numeroEnteros As Integer, ByVal numeroDecimales As Integer, ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

        Dim text As String = txt.Text
        Dim caracter As Char = e.KeyChar

        Dim punto As Boolean = False
        If text.IndexOf(".") = -1 Then
            punto = False
        Else
            punto = True
        End If

        If txt.Text(0) = "0" And Not punto Then
            etiquetaError.Text = "No puedes añadir mas ceros a la izquierda"
        Else

            If (txt.TextLength = 0 And (Not Char.IsDigit(caracter))) Then
                e.Handled = True
                etiquetaError.Text = "El primer digito no puede ser un punto"
            Else
                If (Not Char.IsDigit(caracter) And (text.Contains(".") Or Asc(caracter) <> 46) And Asc(caracter) <> 8) Then
                    e.Handled = True
                    etiquetaError.Text = "Solo numeros y puntos"
                Else
                    If (punto) Then
                        If text.Substring(text.IndexOf(".")).Length() > numeroDecimales And Asc(caracter) <> 8 Then
                            e.Handled = True
                            etiquetaError.Text = "No puedes escribir mas decimales"
                        Else
                            etiquetaError.Text = ""
                        End If
                    End If
                End If
            End If
        End If

    End Sub

    Sub soloDoubles(ByVal txt As TextBox, ByVal etiquetaError As Label, ByVal numeroDecimales As Integer, ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

        Dim text As String = txt.Text
        Dim caracter As Char = e.KeyChar

        Dim punto As Boolean = False
        If text.IndexOf(".") = -1 Then
            punto = False
        Else
            punto = True
        End If

        If (txt.TextLength = 0 And (Not Char.IsDigit(caracter) Or caracter = "0")) Then
            e.Handled = True
            etiquetaError.Text = "El primer digito no puede ser un punto o un 0"
        Else
            If (Not Char.IsDigit(caracter) And (text.Contains(".") Or Asc(caracter) <> 46) And Asc(caracter) <> 8) Then
                e.Handled = True
                etiquetaError.Text = "Solo numeros y puntos"
            Else
                If (punto) Then
                    If text.Substring(text.IndexOf(".")).Length() > numeroDecimales And Asc(caracter) <> 8 Then
                        e.Handled = True
                        etiquetaError.Text = "No puedes escribir mas decimales"
                    Else
                        etiquetaError.Text = ""
                    End If
                End If
            End If
        End If

    End Sub


    Function comprobarVacios(ByVal txt() As TextBox, ByVal nombresCampos() As String) As String

        Dim vacios As String = ""

        For i = 0 To txt.Length - 1

            If txt(i).Text = "" Then
                vacios &= "- The field " & nombresCampos(i) & " can't be empty " & vbCrLf
            End If

        Next

        Return vacios

    End Function

End Module

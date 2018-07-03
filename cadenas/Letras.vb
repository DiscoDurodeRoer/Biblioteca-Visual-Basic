Module Letras

    ''' <summary>
    ''' Genera un numero aleatorio, minimo y maximo incluido(se incluye en letras porque me es necesario para otras funciones)
    ''' </summary>
    ''' <param name="minimo"></param>
    ''' <param name="maximo"></param>
    ''' <returns>Numero entre el minimo y el maximo incluido</returns>
    ''' <remarks></remarks>
    Function numAleatorioEntre(ByVal minimo As Integer, ByVal maximo As Integer) As Integer
        Randomize() ' inicializar la semilla  
        Return CLng((minimo - maximo) * Rnd() + maximo)
    End Function



    Function generaMayus() As Char

        Return Convert.ToChar(numAleatorioEntre(65, 90))

    End Function

    Function generaMinus() As Char

        Return Convert.ToChar(numAleatorioEntre(97, 122))

    End Function

    Function generaCaracteresEspeciales() As Char
        Dim caracteresEspeciales() As Char = {"!", "%", "&", "(", ")", "*", "/", "-", "#", "@"}

        Return caracteresEspeciales(numAleatorioEntre(0, caracteresEspeciales.Length - 1))

    End Function

    Function comprobarDNI(ByVal DNI As String) As Boolean

        Dim correcto As Boolean = False

        Try
            If DNI.Length() >= 8 And DNI.Length() <= 9 Then

                Const DIVISOR As Integer = 23
                Dim letrasNIF() As Char = {"T", "R", "W", "A", "G", "M", "Y",
                "F", "P", "D", "X", "B", "N", "J", "Z",
                "S", "Q", "V", "H", "L", "C", "K", "E"}


                Dim numDNI As Integer = Convert.ToInt32(DNI.Substring(0, DNI.Length - 1))

                Dim res As Integer = numDNI - (Math.Truncate(numDNI / DIVISOR) * DIVISOR)

                'Console.WriteLine(">> " & 

                Dim DNICalculado As String = Convert.ToString(numDNI) & letrasNIF(res)

                If DNICalculado.Equals(DNI) Then
                    correcto = True
                End If

            End If
        Catch ex As FormatException
            MsgBox("La cadena pasada es incorrecta")
        End Try

        Return correcto

    End Function

End Module

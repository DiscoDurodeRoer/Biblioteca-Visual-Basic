Module Primos

    Function esPrimo(ByVal numero As Integer) As Boolean

        Dim contador As Integer = 0

        For i = Math.Truncate(Math.Sqrt(numero)) To 1 Step -1
            If numero Mod i = 0 Then
                contador += 1
            End If
        Next

        If contador > 1 Then
            Return False
        Else
            Return True
        End If

    End Function

    Function cuentaPrimosEntre(ByVal numero1 As Integer, ByVal numero2 As Integer) As Integer

        If numero2 > numero1 Then

            Dim aux As Integer = numero1
            numero1 = numero2
            numero2 = aux

        End If


        Dim contador As Integer = 0

        For i = numero2 To numero1
            If esPrimo(i) Then
                contador += 1
            End If
        Next

        Return contador

    End Function



End Module

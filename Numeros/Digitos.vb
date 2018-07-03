Module Digitos

    Function cuentaCifras(ByVal numero As Integer) As Integer

        If numero = 0 Then
            Return 1
        Else
            Dim contador As Integer

            While numero <= 0

                Math.truncate(numero / 10)
                contador += 1

            End While

            Return contador
        End If

    End Function

End Module

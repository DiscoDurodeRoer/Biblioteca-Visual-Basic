Module Conversion

    'True = byte, si unidadActual es mayor que unidadAPasar
    'false= bit, si unidadAPaar es mayor que unidadActual
    'Mejorar
    Function convertirUnidadInformacion(ByVal unidadActual As String, ByVal unidadAPasar As String, ByVal cantidad As Double, ByVal unidadActEsBit As Boolean, ByVal unidadAPasEsBit As Boolean) As Double

        Dim cantidadPasada As Double = 0

        Dim unidades() As String = {"B", "KB", "MB", "GB", "TB", "PB"}

        Dim posUnidadActual As Integer = 0
        Dim posUnidadAPasar As Integer = 0

        For i = 0 To unidades.Length - 1

            If unidades(i) = unidadActual Then
                posUnidadActual = i
            End If

            If unidades(i) = unidadAPasar Then
                posUnidadAPasar = i
            End If

        Next

        Dim bitAct As Integer = 0
        Dim bitPas As Integer = 0

        If unidadActEsBit Then
            bitAct = 3
        End If

        If unidadAPasEsBit Then
            bitPas = -3
        End If

        ''Si los dos estan activos, no es necesario sumar
        'If unidadActEsBit And unidadAPasEsBit Then
        '    bitAct = 0
        '    bitPas = 0
        'End If

        If posUnidadActual > posUnidadAPasar Then
            cantidadPasada = cantidad * Math.Pow(2, (10 * (posUnidadActual - posUnidadAPasar)) + (bitAct + bitPas))
        Else
            cantidadPasada = cantidad / Math.Pow(2, (10 * (posUnidadAPasar - posUnidadActual)) + (bitAct + bitPas))
        End If

        Return cantidadPasada

    End Function

End Module

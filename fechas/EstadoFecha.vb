Module EstadoFecha


    'Funciones necesarias

    Function deNumeroAFecha(ByVal numero As Integer) As Date

        Try
            Dim anio As Integer = Math.Truncate(numero / 10000)

            Dim mes As Integer = Math.Truncate(((numero Mod 10000) / 100))

            Dim dia As Integer = numero Mod 100

            Dim fecha As New Date(anio, mes, dia)

            Return fecha

        Catch ex As ArgumentOutOfRangeException
            MsgBox("Error, la fecha no es correcta")
            Return Nothing
        End Try

    End Function


    Function deNumeroAHora(ByVal numero As Integer) As Date

        Dim hora As Integer = Math.Truncate(numero / 10000)


        Dim minuto As Integer = Math.Truncate((numero Mod 10000) / 100)


        Dim segundo As Integer = numero Mod 100


        Dim fecha As New Date(1, 1, 1, hora, minuto, segundo)

        Return fecha.ToString("HH:mm:ss")

    End Function



    'Funciones estado fecha

    '0 = fecha1 igual que fecha2
    '1 = fecha1 mayor que fecha2
    '2 = fecha2 mayor que fecha1
    Function fechaMayor(ByVal fecha1 As Date, ByVal fecha2 As Date) As Integer

        If fecha1 = fecha2 Then
            fechaMayor = 0
        Else
            If fecha1 > fecha2 Then
                fechaMayor = 1
            Else
                fechaMayor = 2
            End If
        End If

    End Function

    Function fechaCompletaDate(ByVal fecha As Date, ByVal hora As Date) As Date

        Dim fechaCompleta As New Date(fecha.Year, fecha.Month, fecha.Day, hora.Hour, hora.Minute, hora.Second)

        Return fechaCompleta

    End Function

    Function fechaCompletaDate(ByVal fechaNum As Integer, ByVal horaNum As Integer) As Date

        Dim fecha As Date = deNumeroAFecha(fechaNum)

        Dim hora As Date = deNumeroAHora(horaNum)

        Dim fechaCompleta As New Date(fecha.Year, fecha.Month, fecha.Day, hora.Hour, hora.Minute, hora.Second)

        Return fechaCompleta

    End Function

End Module

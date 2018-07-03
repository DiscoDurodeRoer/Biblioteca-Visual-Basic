Module ConversionFechas

    Function deFechaANumero(ByVal fecha As Date) As Integer

        Dim f As String = fecha.Year

        If fecha.Month < 10 Then
            f &= "0" & fecha.Month
        Else
            f &= fecha.Month
        End If

        If fecha.Day < 10 Then
            f &= "0" & fecha.Day
        Else
            f &= fecha.Day
        End If

        Return Convert.ToInt32(f)

    End Function

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

    Function deHoraANumero(ByVal fecha As Date) As Integer
        Dim f As String = fecha.ToString("HH:mm:ss")

        If fecha.Hour < 10 Then
            f = "0" & fecha.Hour
        Else
            f = fecha.Hour
        End If

        If fecha.Minute < 10 Then
            f &= "0" & fecha.Minute
        Else
            f &= fecha.Minute
        End If

        If fecha.Second < 10 Then
            f &= "0" & fecha.Second
        Else
            f &= fecha.Second
        End If

        Return Convert.ToInt32(f)

    End Function

    Function deNumeroAHora(ByVal numero As Integer) As Date

        Dim hora As Integer = Math.Truncate(numero / 10000)


        Dim minuto As Integer = Math.Truncate((numero Mod 10000) / 100)


        Dim segundo As Integer = numero Mod 100


        Dim fecha As New Date(1, 1, 1, hora, minuto, segundo)

        Return fecha.ToString("HH:mm:ss")

    End Function

End Module

Module Fechas

    Function esBisiesto(ByVal anio As Integer)

        Dim bisiesto As Boolean = False

        If (anio Mod 4 = 0) And Not (anio Mod 100 = 0) Or (anio Mod 400 = 0) Then
            bisiesto = True
        End If

        Return bisiesto

    End Function

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

    Function validarFecha(ByVal anio As Integer, ByVal mes As Integer, ByVal dia As Integer) As Boolean

        Dim valido As Boolean = False

        If anio >= 0 And (mes >= 1 And mes <= 12) Then
            If numeroDeDiasMes(mes) >= dia Then
                valido = True
            End If
        End If

        Return valido
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

    Function validaHora(ByVal hora As Integer, ByVal minuto As Integer, ByVal segundo As Integer) As Boolean

        Dim valido As Boolean = False

        If (hora >= 0 And hora <= 23) And (minuto >= 0 And minuto <= 59) And (segundo >= 0 And segundo <= 59) Then
            valido = True
        End If

        Return valido

    End Function

    Function numeroDeDiasMes(ByVal mes As Integer) As Integer

        Dim numeroDias As Integer = -1

        Select Case (mes)
            Case 1
                numeroDias = 31
            Case 3
                numeroDias = 31
            Case 5
                numeroDias = 31
            Case 7
                numeroDias = 31
            Case 8
                numeroDias = 31
            Case 10
                numeroDias = 31
            Case 12
                numeroDias = 31
            Case 4
                numeroDias = 30
            Case 6
                numeroDias = 30
            Case 9
                numeroDias = 30
            Case 11
                numeroDias = 30
            Case 2S

                Dim anioActual As Date = Date.Now

                If esBisiesto(anioActual.Year) Then
                    numeroDias = 29
                Else
                    numeroDias = 28
                End If
        End Select


        Return numeroDias

    End Function

    Function numeroDeDiasMes(ByVal mes As Integer, ByVal anio As Integer) As Integer

        Dim numeroDias As Integer = -1

        Select Case (mes)
            Case 1
                numeroDias = 31
            Case 3
                numeroDias = 31
            Case 5
                numeroDias = 31
            Case 7
                numeroDias = 31
            Case 8
                numeroDias = 31
            Case 10
                numeroDias = 31
            Case 12
                numeroDias = 31
            Case 4
                numeroDias = 30
            Case 6
                numeroDias = 30
            Case 9
                numeroDias = 30
            Case 11
                numeroDias = 30
            Case 2S

                Dim anioActual As Date = Date.Now

                If esBisiesto(anio) Then
                    numeroDias = 29
                Else
                    numeroDias = 28
                End If
        End Select


        Return numeroDias

    End Function

    Function numeroDiasEntre(ByVal fecha1 As Date, ByVal fecha2 As Date) As Integer

        Dim dias As Integer

        If fecha1 > fecha2 Then
            dias = (fecha1 - fecha2).TotalDays
        Else
            dias = (fecha2 - fecha1).TotalDays
        End If

        Return dias

    End Function

    Function numeroAniosCienticoEntre(ByVal fechaCientifica1 As Integer, ByVal fechaCientifica2 As Integer) As Integer

        Dim anio1 As Integer = Math.Truncate(fechaCientifica1 / 10000)
        Dim anio2 As Integer = Math.Truncate(fechaCientifica2 / 10000)

        Dim resta As Integer = anio1 - anio2

        If resta < 0 Then
            Return Math.Abs(resta)
        Else

            Return resta
        End If

    End Function

    Function numeroAniosEntre(ByVal fecha1 As Date, ByVal fecha2 As Date) As Integer

        Dim anio1 As Integer = fecha1.Year
        Dim anio2 As Integer = fecha2.Year

        Dim resta As Integer = anio1 - anio2

        If resta < 0 Then
            Return Math.Abs(resta)
        Else

            Return resta
        End If

    End Function

    Function numeroDiasCientificosEntre(ByVal fechaCientifica1 As Integer, ByVal fechaCientifica2 As Integer) As Integer

        Dim diaFecha1 As Integer = Math.Truncate(fechaCientifica1 / 1000000)
        Dim diaFecha2 As Integer = Math.Truncate(fechaCientifica2 / 1000000)

        Dim resta As Integer = diaFecha1 - diaFecha2

        If resta < 0 Then
            Return Math.Abs(resta)
        Else
            Return resta
        End If

    End Function

    Sub establecerIntervaloDTP(ByVal dtps() As DateTimePicker, ByVal mesOrigenIntervalo As Integer, ByVal diaMesInicio As Integer, ByVal diaMesFin As Integer)

        If Date.Now.Month = mesOrigenIntervalo Then

            For i = 0 To dtps.Length

                dtps(i).MinDate = "#" & diaMesInicio & Date.Now.Year - 1 & "#"
                dtps(i).MaxDate = "#" & diaMesFin & Date.Now.Year & "#"

            Next

        Else

            For i = 0 To dtps.Length

                dtps(i).MinDate = "#" & diaMesInicio & Date.Now.Year & "#"
                dtps(i).MaxDate = "#" & diaMesFin & Date.Now.Year + 1 & "#"

            Next

        End If

    End Sub

    Sub establecerIntervaloDTP(ByVal dtp As DateTimePicker, ByVal mesOrigenIntervalo As Integer, ByVal diaMesInicio As Integer, ByVal diaMesFin As Integer)

        If Date.Now.Month = mesOrigenIntervalo Then

            dtps(i).MinDate = "#" & diaMesInicio & Date.Now.Year - 1 & "#"
            dtps(i).MaxDate = "#" & diaMesFin & Date.Now.Year & "#"

        Else

            dtps(i).MinDate = "#" & diaMesInicio & Date.Now.Year & "#"
            dtps(i).MaxDate = "#" & diaMesFin & Date.Now.Year + 1 & "#"

        End If

    End Sub

    Function fechaEntreIntervalo(ByVal fechaEvaluada As Date, ByVal fechaMinima As Date, ByVal fechaMaxima As Date) As Boolean

        If fechaEvaluada.ToString("dd/MM/yyyy") > fechaMinima.ToString("dd/MM/yyyy") And fechaEvaluada.ToString("dd/MM/yyyy") < fechaMaxima.ToString("dd/MM/yyyy") Then
            Return True
        End If

        Return False

    End Function

    Sub anadirDiasLaborables(ByVal dtp As DateTimePicker, ByVal fechaInicio As DateTime, ByVal numeroDias As Integer)

        Dim fecha As Date = fechaInicio

        Dim contador As Integer

        While (numeroDias > contador)

            Select Case fecha.DayOfWeek
                Case 1 To 5
                    fecha = fecha.AddDays(1)
                    dtpHasta.Value = fecha
                    contador += 1
                Case 0, 6
                    fecha = fecha.AddDays(1)

                Case Else
            End Select



        End While

        fecha = fecha.AddDays(-1)
        dtpHasta.Value = fecha

    End Sub

    'Cuenta los dias entre una fecha y otra sin contar los fines de semana
    Function calcularDiasLaborables(ByVal fechaInicial As Date, ByVal fechaFinal As Date) As Integer

        Dim contador As Integer = 0

        Dim fecha As Date = fechainicial

        While (fecha < fechaFinal)

            Select Case fecha.DayOfWeek
                Case 1 To 5
                    contador += 1
                Case Else
            End Select

            fecha = fecha.AddDays(1)

        End While

        Return contador - 1

    End Function

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


    ''' <summary>
    ''' Devuelve un array con los meses del año
    ''' </summary>
    ''' <returns>Un array con todos los meses del año</returns>
    ''' <remarks></remarks>
    Function mesesAño() As String()
        Dim meses() As String = {"Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"}
        Return meses
    End Function

End Module

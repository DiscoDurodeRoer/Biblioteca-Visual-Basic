Imports System.Windows.Forms

Module Dias

    'Funciones necesarias

    Function esBisiesto(ByVal anio As Integer)

        Dim bisiesto As Boolean = False

        If (anio Mod 4 = 0) And Not (anio Mod 100 = 0) Or (anio Mod 400 = 0) Then
            bisiesto = True
        End If

        Return bisiesto

    End Function

    'Funciones de dias

    Function numeroDeDiasMes(ByVal mes As Integer) As Integer

        Dim numeroDias As Integer = -1

        Select Case (mes)
            Case 1, 3, 5, 7, 8, 10, 12
                numeroDias = 31

            Case 4, 6, 9, 11
                numeroDias = 30
            Case 2

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
            Case 1, 3, 5, 7, 8, 10, 12
                numeroDias = 31

            Case 4, 6, 9, 11
                numeroDias = 30

            Case 2

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

    'Cuenta los dias entre una fecha y otra sin contar los fines de semana
    Function calcularDiasLaborables(ByVal fechaInicial As Date, ByVal fechaFinal As Date) As Integer

        Dim contador As Integer = 0

        Dim fecha As Date = fechaInicial

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

    Sub anadirDiasLaborables(ByVal dtp As DateTimePicker, ByVal fechaInicio As DateTime, ByVal numeroDias As Integer)

        Dim fecha As Date = fechaInicio

        Dim contador As Integer

        While (numeroDias > contador)

            Select Case fecha.DayOfWeek
                Case 1 To 5
                    fecha = fecha.AddDays(1)
                    dtp.Value = fecha
                    contador += 1
                Case 0, 6
                    fecha = fecha.AddDays(1)

                Case Else
            End Select



        End While

        fecha = fecha.AddDays(-1)
        dtp.Value = fecha

    End Sub
End Module

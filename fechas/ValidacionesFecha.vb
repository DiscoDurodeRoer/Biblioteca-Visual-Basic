Module ValidacionesFecha

    'Funciones necesarias

    Function esBisiesto(ByVal anio As Integer)

        Dim bisiesto As Boolean = False

        If (anio Mod 4 = 0) And Not (anio Mod 100 = 0) Or (anio Mod 400 = 0) Then
            bisiesto = True
        End If

        Return bisiesto

    End Function

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


    'Funciones de validacion de fechas

    Function validarFecha(ByVal anio As Integer, ByVal mes As Integer, ByVal dia As Integer) As Boolean

        Dim valido As Boolean = False

        If anio >= 0 And (mes >= 1 And mes <= 12) Then
            If numeroDeDiasMes(mes) >= dia Then
                valido = True
            End If
        End If

        Return valido
    End Function

    Function validaHora(ByVal hora As Integer, ByVal minuto As Integer, ByVal segundo As Integer) As Boolean

        Dim valido As Boolean = False

        If (hora >= 0 And hora <= 23) And (minuto >= 0 And minuto <= 59) And (segundo >= 0 And segundo <= 59) Then
            valido = True
        End If

        Return valido

    End Function

End Module

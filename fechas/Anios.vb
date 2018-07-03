Module Anios

    Function esBisiesto(ByVal anio As Integer)

        Dim bisiesto As Boolean = False

        If (anio Mod 4 = 0) And Not (anio Mod 100 = 0) Or (anio Mod 400 = 0) Then
            bisiesto = True
        End If

        Return bisiesto

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

    'Sin Probar
    'Hace falta la funcion deFechaANumero
    Function sacarAnioDe(ByVal fecha As Date) As Integer

        Dim numCientificoFecha = deFechaANumero(fecha)

        Return Math.Truncate(numCientificoFecha) / 10000

    End Function

End Module

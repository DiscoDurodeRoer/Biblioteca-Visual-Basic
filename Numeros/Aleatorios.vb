Module Aleatorios

    'Funciones necesarias

    Sub rellenaArray(ByRef vector As Array, ByVal num As Integer)
        For i = 0 To vector.Length - 1
            vector(i) = num
        Next

    End Sub

    'Funciones de aleatorio

    ''' Genera un numero aleatorio, minimo y maximo incluido
    ''' </summary>
    ''' <param name="minimo"></param>
    ''' <param name="maximo"></param>
    ''' <returns>Numero entre el minimo y el maximo incluido</returns>
    ''' <remarks></remarks>
    Function numAleatorioEntre(ByVal minimo As Integer, ByVal maximo As Integer) As Integer
        Randomize() ' inicializar la semilla  
        Return CLng((minimo - maximo) * Rnd() + maximo)
    End Function

    ''' <summary>
    ''' Crea un array con la longitud pasada, con numeros aleatorios entre el minimo y el maximo incluidos
    ''' </summary>
    ''' <param name="longitud"></param>
    ''' <param name="minimo"></param>
    ''' <param name="maximo"></param>
    ''' <returns>Array de numeros entre el minimo y el maximo incluidos</returns>
    ''' <remarks></remarks>
    Function numerosAleatoriosEntre(ByVal longitud As Integer, ByVal minimo As Integer, ByVal maximo As Integer) As Integer()

        Dim numeros(longitud - 1) As Integer

        For i = 0 To numeros.Length - 1
            numeros(i) = numAleatorioEntre(minimo, maximo)
        Next

        Return numeros

    End Function

    ''' <summary>
    ''' Genera numeros aleatorios no repetidos entre los numeros indicados. ATENCION: la longitud no puede ser mayor que la diferencia entre los numeros indicados
    ''' Por ejemplo (10, 0, 9) Es valido (un array de 10 posiciones con numeros del 0 al 9 incluidos)
    '''             (11, 0, 9) No es valido (un array de 11 posiciones con numeros del 0 al 9 incluidos(faltaria un numero mas))
    ''' </summary>
    ''' <param name="longitud"></param>
    ''' <param name="minimo"></param>
    ''' <param name="maximo"></param>
    ''' <returns>Array de numeros entre el minimo y el maximo incluidos no repetidos</returns>
    ''' <remarks></remarks>
    Function numerosAleatoriosNoRepetidosEntre(ByVal longitud As Integer, ByVal minimo As Integer, ByVal maximo As Integer) As Integer()

        If ((longitud - 1) - (maximo - minimo)) <= 0 Then

            Dim numeros(longitud - 1) As Integer

            rellenaArray(numeros, minimo - 1)

            Dim i As Integer = 0
            Dim num As Integer = 0
            While i < numeros.Length

                Dim repetido As Boolean
                Do
                    repetido = False
                    num = numAleatorioEntre(minimo, maximo)

                    For j = 0 To numeros.Length - 1 And Not repetido
                        If numeros(j) = num Then
                            repetido = True
                        End If
                    Next
                Loop While repetido

                numeros(i) = num
                i += 1
            End While

            Return numeros
        Else
            MsgBox("Error, la diferencia entre el maximo y el minimo debe ser mayor que la longitud")
            Return Nothing
        End If

    End Function

    Function numerosAleatoriosNoRepetidosEntreV2(ByVal longitud As Integer, ByVal minimo As Integer, ByVal maximo As Integer) As Integer()


        Dim numeros(longitud - 1) As Integer

        rellenaArray(numeros, minimo - 1)

        Dim i As Integer = 0
        Dim num As Integer = 0
        While i < numeros.Length

            Dim repetido As Boolean
            Do
                repetido = False
                num = numAleatorioEntre(minimo, maximo)

                For j = 0 To numeros.Length - 1 And Not repetido
                    If numeros(j) = num Then
                        repetido = True
                    End If
                Next
            Loop While repetido

            numeros(i) = num
            i += 1
        End While

        Return numeros

    End Function

    Function generarNumerosPrimosAzarEntre(ByVal minimo As Integer, ByVal maximo As Integer) As Integer()

        Dim numerosPrimos(cuentaPrimosEntre(minimo, maximo) - 1) As Integer

        If numerosPrimos.Length = 0 Then
            MsgBox("No hay primos entre " & minimo & " y " & maximo)
            Return Nothing
        Else

            Dim indice As Integer = 0

            For i = minimo To maximo
                If esPrimo(i) Then
                    numerosPrimos(indice) = i
                    indice += 1
                End If
            Next

            Return numerosPrimos

        End If

    End Function

    Function primoAlAzarEntre(ByVal minimo As Integer, ByVal maximo As Integer) As Integer

        Dim numPrimo As Integer = -1
        If cuentaPrimosEntre(minimo, maximo) = 0 Then
            MsgBox("No hay primos entre " & minimo & " y " & maximo)
        Else
            Do
                numPrimo = numAleatorioEntre(minimo, maximo)
            Loop While Not esPrimo(numPrimo)
        End If

        Return numPrimo

    End Function


End Module

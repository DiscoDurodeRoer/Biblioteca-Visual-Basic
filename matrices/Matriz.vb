Module Matriz

    'Funciones necesarias


    ''' <summary>
    ''' Pide un numero y hasta que no se introduzca no se devuelve el numero
    ''' </summary>
    ''' <returns>numero validado</returns>
    ''' <remarks></remarks>
    Function validaNumero() As Integer

        Dim interruptor As Boolean
        interruptor = True
        Dim num As Integer

        Do While interruptor
            Try
                num = Convert.ToInt32(Console.ReadLine())
                interruptor = False
            Catch ex As FormatException
                Console.WriteLine("No es un numero, intentalo de nuevo")
            End Try
        Loop

        Return num

    End Function

    ''' <summary>
    ''' Pide un numero entre el minimo y el maximo, hasta que se introduzca un numero entre estos numeros se devolvera el numero
    ''' </summary>
    ''' <param name="minimo"></param>
    ''' <param name="maximo"></param>
    ''' <returns>Numero validado</returns>
    ''' <remarks></remarks>
    Function validaNumeroEntre(ByVal minimo As Integer, ByVal maximo As Integer) As Integer
        Dim interruptor As Boolean
        interruptor = True
        Dim num As Integer

        Do
            Try
                num = Convert.ToInt32(Console.ReadLine())
                If (num < minimo Or num > maximo) Then
                    Console.WriteLine("Debes introducir un numero entre " & minimo & " y " & maximo)
                Else
                    interruptor = False
                End If
            Catch ex As FormatException
                Console.WriteLine("No es un numero, intentalo de nuevo")
            End Try
        Loop While interruptor And (num < minimo Or num > maximo)

        Return num

    End Function

    ''' <summary>
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


    'Aqui empiezan las funciones de la matriz


    ''' <summary>
    ''' Rellena una matriz con valores que nosotros le pasemos
    ''' </summary>
    ''' <param name="matriz"></param>
    ''' <remarks></remarks>
    Sub rellenaMatriz(ByRef matriz)

        For i = 0 To matriz.GetUpperBound(0) - 1
            For j = 0 To matriz.GetUpperBound(1) - 1
                Console.WriteLine("Escribe un valor para la posicion " & i & " " & j)
                matriz(i, j) = Console.ReadLine()
            Next
        Next

    End Sub

    ''' <summary>
    ''' Rellena una matriz con valores que nosotros le pasemos
    ''' </summary>
    ''' <param name="matriz"></param>
    ''' <remarks></remarks>
    Sub rellenaMatrizValidado(ByRef matriz)

        For i = 0 To matriz.GetUpperBound(0) - 1
            For j = 0 To matriz.GetUpperBound(1) - 1
                Console.WriteLine("Escribe un valor para la posicion " & i & " " & j)
                If IsNumeric(matriz(i, j)) Then
                    matriz(i, j) = validaNumero()
                Else
                    matriz(i, j) = Console.ReadLine
                End If
            Next
        Next

    End Sub

    ''' <summary>
    ''' Muestra el contenido de una matriz
    ''' </summary>
    ''' <param name="matriz"></param>
    ''' <remarks></remarks>
    Sub muestraMatriz(ByRef matriz)

        For i = 0 To matriz.GetUpperBound(0)
            For j = 0 To matriz.GetUpperBound(1)
                Console.Write(matriz(i, j))
            Next
            Console.WriteLine()
        Next

    End Sub

    Sub muestraMatrizObj(ByRef matriz As Object)

        For i = 0 To matriz.GetUpperBound(0)
            For j = 0 To matriz.GetUpperBound(1)
                Console.WriteLine(matriz(i, j).ToString)
            Next
        Next

    End Sub

    'adyacentes
    ''' <summary>
    ''' Devuelve un array con las posiciones adyacentes al pasado por parametros
    ''' </summary>
    ''' <param name="matriz"></param>
    ''' <param name="i"></param>
    ''' <param name="j"></param>
    ''' <returns>Array con posiciones adyancentes</returns>
    ''' <remarks></remarks>
    Function adyacentesAV1(ByRef matriz As Object, ByVal i As Integer, ByVal j As Integer) As Integer()
        Dim tamanio As Integer = matriz.GetUpperBound(0) + 1
        Dim dirs() = {1, -1, 1, 0, 1, 1, 0, 1, -1, 1, -1, 0, -1, -1, 0, -1}
        Dim posicionesValidas(15) As Integer

        Dim indice2 = 0
        For indice = 0 To dirs.Length - 2 Step 2
            If ((i + dirs(indice)) >= 0 And (i + dirs(indice)) < tamanio) And ((j + dirs(indice + 1)) >= 0 And (j + dirs(indice + 1)) < tamanio) Then
                posicionesValidas(indice2) = i + dirs(indice)
                indice2 += 1
                posicionesValidas(indice2) = j + dirs(indice + 1)
                indice2 += 1

            End If

        Next

        ReDim Preserve posicionesValidas(indice2 - 1)

        adyacentesA = posicionesValidas
    End Function

    ''' <summary>
    ''' Posicionar un elemento en una matriz pasandoles la posicion
    ''' </summary>
    ''' <param name="matriz"></param>
    ''' <param name="posicion"></param>
    ''' <param name="objeto"></param>
    ''' <remarks></remarks>
    Sub posicionarElementoMatriz(ByRef matriz As Object, ByVal posicion As Integer, ByRef objeto As Object)

        Dim tamanio As Integer = matriz.GetUpperBound(0) + 1

        matriz(Math.Truncate(posicion / tamanio), posicion Mod tamanio) = objeto

    End Sub

    '
    ''' <summary>
    ''' Posiciona un objeto y devuelve un array indicando donde lo ha posicionado
    ''' </summary>
    ''' <param name="matriz"></param>
    ''' <param name="posicion"></param>
    ''' <param name="objeto"></param>
    ''' <returns>Array con posiciones donde se ha insertado el objeto</returns>
    ''' <remarks></remarks>
    Function devolverPosicionMatriz(ByRef matriz As Object(,), ByVal posicion As Integer, ByRef objeto As Object) As Integer()

        Dim posiciones(1) As Integer

        Dim tamanio As Integer = matriz.GetUpperBound(0) + 1

        matriz(Math.Truncate(posicion / tamanio), posicion Mod tamanio) = objeto

        posiciones(0) = Math.Truncate(posicion / tamanio)
        posiciones(1) = posicion Mod tamanio

        Return posiciones

    End Function

    ''' <summary>
    ''' Posiciona un objeto en una matriz aleatoriamente
    ''' </summary>
    ''' <param name="matriz"></param>
    ''' <param name="objeto"></param>
    ''' <remarks></remarks>
    Sub posicionarElementoMatriz(ByRef matriz As Object, ByRef objeto As Object)

        Dim tamanio As Integer = matriz.GetUpperBound(0) + 1
        Dim posicion As Integer = numAleatorioEntre(0, (tamanio * tamanio) - 1)

        matriz(Math.Truncate(posicion / tamanio), posicion Mod tamanio) = objeto

    End Sub

    Function devolverPosiciones(ByVal matriz As Object(,), ByVal posicion As Integer) As Integer()

        Dim posiciones(1) As Integer

        Dim tamanio As Integer = matriz.GetUpperBound(0)

        posiciones(0) = Math.Floor(posicion / tamanio)
        posiciones(1) = posicion Mod tamanio

        Return posiciones

    End Function

    Function existePosicion(ByVal matriz As Object(,), ByVal fila As Integer, ByVal columna As Integer) As Boolean

        Dim existe As Boolean = False

        If (IsNothing(matriz(fila, columna))) Then
            existe = True
        End If

        Return existe

    End Function


    Function posicionCentroMatriz(ByVal matriz(,) As Object) As Integer()

        Dim posiciones(1) As Integer

        posiciones(0) = Math.Truncate(matriz.GetUpperBound(0) / 2) 'Fila
        posiciones(1) = Math.Truncate(matriz.GetUpperBound(1) / 2) 'Columna

        Return posiciones

    End Function

    Function posicionesEnCruz(ByVal posI As Integer, ByVal posJ As Integer) As Integer(,)

        Dim posiciones(1, 3) As Integer

        Dim contador As Integer = 0

        Dim indice As Integer = 0

        For i = -1 To 1
            For j = -1 To 1

                If Not (((posI + i < 0) Or (posI + i > matriz.GetUpperBound(0))) Or (posJ + j < 0 Or (posJ + j > matriz.GetUpperBound(1)))) And (Math.Abs(i) + Math.Abs(j) = 1) Then
                    posiciones(0, indice) = posI + i
                    posiciones(1, indice) = posJ + j
                    indice += 1
                    contador += 1
                End If

            Next
        Next

        If contador <> matriz.GetUpperBound(0) Then
            ReDim Preserve posiciones(1, contador - 1)
        End If

        Return posiciones


    End Function

    Function AdyacentesAV2(ByVal posI As Integer, ByVal posJ As Integer) As Integer(,)

        Dim posiciones(1, 7) As Integer

        Dim contador As Integer = 0

        Dim indice As Integer = 0

        For i = -1 To 1
            For j = -1 To 1

                If Not (((posI + i < 0) Or (posI + i > matriz.GetUpperBound(0))) Or (posJ + j < 0 Or (posJ + j > matriz.GetUpperBound(1)))) And (i <> 0 Or j <> 0) Then
                    posiciones(0, indice) = posI + i
                    posiciones(1, indice) = posJ + j
                    indice += 1
                    contador += 1
                End If

            Next
        Next

        If contador <> matriz.GetUpperBound(0) Then
            ReDim Preserve posiciones(1, contador - 1)
        End If

        Return posiciones

    End Function


    Function posicionesDiagonales(ByVal posI As Integer, ByVal posJ As Integer) As Integer(,)

        Dim posiciones(1, 3) As Integer

        Dim contador As Integer = 0

        Dim indice As Integer = 0

        For i = -1 To 1
            For j = -1 To 1

                'If Not (((posI + i < 0) Or (posI + i > matriz.GetUpperBound(0))) Or (posJ + j < 0 Or (posJ + j > matriz.GetUpperBound(1)))) And ((posJ + j) <> posJ And (posI + i) <> posI) Then
                '    posiciones(0, indice) = posI + i
                '    posiciones(1, indice) = posJ + j
                '    indice += 1
                '    contador += 1
                'End If

                If posicionDentroMatriz(matriz, posI + i, posJ + j) And ((posJ + j) <> posJ And (posI + i) <> posI) Then
                    posiciones(0, indice) = posI + i
                    posiciones(1, indice) = posJ + j
                    indice += 1
                    contador += 1
                End If

            Next
        Next

        If contador <> matriz.GetUpperBound(0) Then
            ReDim Preserve posiciones(1, contador - 1)
        End If

        Return posiciones

    End Function

    Function posicionFueraMatriz(ByVal matriz As Object, ByVal posI As Integer, ByVal posJ As Integer) As Boolean

        Return ((posI < 0) Or (posI > matriz.GetUpperBound(0))) Or (posJ < 0 Or (posJ > matriz.GetUpperBound(1)))

    End Function

    Function posicionDentroMatriz(ByVal matriz As Object, ByVal posI As Integer, ByVal posJ As Integer) As Boolean

        Return ((posI >= 0) And posI <= matriz.GetUpperBound(0)) And (posJ >= 0 And posJ <= matriz.GetUpperBound(1))

    End Function

End Module

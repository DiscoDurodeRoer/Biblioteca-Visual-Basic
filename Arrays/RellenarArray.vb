Module RellenarArray

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

    'Funciones

    ''' <summary>
    ''' Rellena un array con los valores que vamos metiendo, segun el primer valor determina el tipo de valores que se deben añadir
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <remarks></remarks>
    Sub rellenaArray(ByRef vector As Array)

        For i = 0 To vector.Length - 1
            Console.WriteLine("Escribe un valor para la posicion " & i)
            If IsNumeric(vector(i)) Then
                vector(i) = validaNumero()
            Else
                vector(i) = Console.ReadLine
            End If

        Next

    End Sub

    ''' <summary>
    ''' Rellena el array con un valor
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <param name="num"></param>
    ''' <remarks></remarks>
    Sub rellenaArray(ByRef vector As Array, ByVal num As Integer)
        For i = 0 To vector.Length - 1
            vector(i) = num
        Next

    End Sub

    ''' <summary>
    ''' Rellena un array con numeros introducidos por nosotros, debe ser entre el minimo y el maximo
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <param name="minimo"></param>
    ''' <param name="maximo"></param>
    ''' <remarks></remarks>
    Sub rellenaArray(ByRef vector() As Integer, ByVal minimo As Integer, ByVal maximo As Integer)

        For i = 0 To vector.Length - 1
            Console.WriteLine("Escribe un valor para la posicion " & i & " debe ser mayor o igual que " & minimo & " y menor o igual que " & maximo)
            Dim num As Integer

            num = validaNumeroEntre(minimo, maximo)

            vector(i) = num
        Next

    End Sub

    ''' <summary>
    ''' Rellena un array con numero aleatorios entre el minimo y el maximo incluidos
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <param name="minimo"></param>
    ''' <param name="maximo"></param>
    ''' <remarks></remarks>
    Sub rellenaArrayAleatorios(ByRef vector() As Integer, ByVal minimo As Integer, ByVal maximo As Integer)

        For i = 0 To vector.Length - 1
            vector(i) = numAleatorioEntre(minimo, maximo)
        Next

    End Sub



End Module

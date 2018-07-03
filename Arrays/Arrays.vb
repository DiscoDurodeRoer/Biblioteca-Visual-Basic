Module Arrays

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



    'Aqui empiezan las funciones de arrays






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

    ''' <summary>
    ''' Muestra todos los valores del array
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <remarks></remarks>
    Sub mostrarArray(ByRef vector As Array)

        For i = 0 To vector.Length - 1
            Console.WriteLine(vector(i))
        Next

    End Sub

    ''' <summary>
    ''' Muestra todos los valores hasta donde le digamos
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <param name="limite"></param>
    ''' <remarks></remarks>
    Sub mostrarArray(ByRef vector As Array, ByVal limite As Integer)
        For i = 0 To limite - 1
            Console.WriteLine(vector(i))
        Next
    End Sub

    ''' <summary>
    ''' Copia el contenido de un array a otro
    ''' </summary>
    ''' <param name="vectorOriginal"></param>
    ''' <param name="vectorACopiar"></param>
    ''' <remarks></remarks>
    Sub copiaArray(ByRef vectorOriginal As Array, ByRef vectorACopiar As Array)

        Dim limite As Integer = 0
        If vectorOriginal.Length = vectorACopiar.Length Then
            limite = vectorOriginal.Length - 1
        Else
            If vectorOriginal.Length > vectorACopiar.Length Then
                limite = vectorACopiar.Length - 1
            Else
                limite = vectorOriginal.Length - 1
            End If
        End If

        For i = 0 To limite
            vectorACopiar(i) = vectorOriginal(i)
        Next
    End Sub

    ''' <summary>
    ''' Suma los valores del array pasado
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <returns>Suma de todos los valores</returns>
    ''' <remarks></remarks>
    Function sumaArray(ByRef vector() As Integer) As Integer

        Dim suma As Integer = 0
        For i = 0 To vector.Length - 1
            suma += vector(i)
        Next

        Return suma

    End Function

    ''' <summary>
    ''' Hace la media de un array de numeros
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <returns>La media de todos los valores del array</returns>
    ''' <remarks></remarks>
    Function mediaArray(ByRef vector() As Integer) As Double

        Return sumaArray(vector) / vector.Length

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

    Function CopiarFila_A_Array(ByRef matriz(,) As String, ByVal fila As Integer) As String()

        If fila < 0 Or fila > matriz.GetUpperBound(0) Then
            MsgBox("la fila seleccionada no existe")
            Return Nothing
        Else
            Dim aux(matriz.GetUpperBound(0) + 1) As String

            For j = 0 To matriz.GetUpperBound(1)
                aux(j) = matriz(fila, j)
            Next

            Return aux
        End If

    End Function

    Function CopiarFila_A_Array(ByRef matriz(,) As Integer, ByVal fila As Integer) As Integer()

        If fila < 0 Or fila > matriz.GetUpperBound(0) Then
            MsgBox("la fila seleccionada no existe")
            Return Nothing
        Else
            Dim aux(matriz.GetUpperBound(0) + 1) As Integer

            For j = 0 To matriz.GetUpperBound(1)
                aux(j) = matriz(fila, j)
            Next

            Return aux

        End If

    End Function

    Function CopiarFila_A_Array(ByRef matriz(,) As Double, ByVal fila As Integer) As Double()

        If fila < 0 Or fila > matriz.GetUpperBound(0) Then
            MsgBox("la fila seleccionada no existe")
            Return Nothing
        Else
            Dim aux(matriz.GetUpperBound(0) + 1) As Double

            For j = 0 To matriz.GetUpperBound(1)
                aux(j) = matriz(fila, j)
            Next

            Return aux

        End If

    End Function

    Function CopiarColumna_A_Array(ByRef matriz(,) As String, ByVal columna As Integer) As String()

        If columna < 0 Or columna > matriz.GetUpperBound(1) Then
            MsgBox("La columna seleccionada no existe")
            Return Nothing
        Else
            Dim aux(matriz.GetUpperBound(1) + 1) As String

            For j = 0 To matriz.GetUpperBound(0)
                aux(j) = matriz(j, columna)
            Next

            Return aux

        End If

    End Function

    Function CopiarColumna_A_Array(ByRef matriz(,) As Integer, ByVal columna As Integer) As Integer()

        If columna < 0 Or columna > matriz.GetUpperBound(1) Then
            MsgBox("La columna seleccionada no existe")
            Return Nothing
        Else
            Dim aux(matriz.GetUpperBound(1) + 1) As Integer

            For j = 0 To matriz.GetUpperBound(0)
                aux(j) = matriz(j, columna)
            Next

            Return aux

        End If

    End Function

    Function CopiarColumna_A_Array(ByRef matriz(,) As Double, ByVal columna As Integer) As Double()

        If columna < 0 Or columna > matriz.GetUpperBound(1) Then
            MsgBox("La columna seleccionada no existe")
            Return Nothing
        Else
            Dim aux(matriz.GetUpperBound(1) + 1) As Double

            For j = 0 To matriz.GetUpperBound(0)
                aux(j) = matriz(j, columna)
            Next

            Return aux

        End If

    End Function

    Public Function devuelvePlanetas() As String()

        Dim planetas() As String = {"Mercurio", "Venus", "Tierra", "Marte", "Jupiter",
                          "Saturno", "Urano", "Neptuno", "Plutón"}

        Return planetas

    End Function


End Module

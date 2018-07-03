Module Numeros

    'Funciones necesarias

    Sub rellenaArray(ByRef vector As Array, ByVal num As Integer)
        For i = 0 To vector.Length - 1
            vector(i) = num
        Next

    End Sub


    'Aqui empiezan las funciones de numeros

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
    ''' Pide un numero positivo, hasta que el usuario no ingresa un numero positivo no devuelve el numero
    ''' </summary>
    ''' <returns>Numero positivo validado</returns>
    ''' <remarks></remarks>
    Function validaNumeroPositivo() As Integer
        Dim interruptor As Boolean
        interruptor = True
        Dim num As Integer

        Do
            Try
                num = Convert.ToInt32(Console.ReadLine())
                If (num < 0) Then
                    Console.WriteLine("Debes introducir un numero mayor que 0")
                Else
                    interruptor = False
                End If
            Catch ex As FormatException
                Console.WriteLine("No es un numero, intentalo de nuevo")
            End Try
        Loop While interruptor And (num < 0)

        Return num
    End Function

    ''' <summary>
    ''' Pide un numero negativo, hasta que el usuario no ingresa un numero negativo no devuelve el numero
    ''' </summary>
    ''' <returns>Numero negativo validado</returns>
    ''' <remarks></remarks>
    Function validaNumeroNegativo() As Integer
        Dim interruptor As Boolean
        interruptor = True
        Dim num As Integer

        Do
            Try
                num = Convert.ToInt32(Console.ReadLine())
                If (num >= 0) Then
                    Console.WriteLine("Debes introducir un numero menor que 0")
                Else
                    interruptor = False
                End If
            Catch ex As FormatException
                Console.WriteLine("No es un numero, intentalo de nuevo")
            End Try
        Loop While interruptor And (num >= 0)

        Return num
    End Function

    ''' <summary>
    ''' Indica si un numero es multiplo de otro
    ''' </summary>
    ''' <param name="num"></param>
    ''' <param name="multiplo"></param>
    ''' <returns>Booleano indicando si es multiplo o no</returns>
    ''' <remarks></remarks>
    Function esMultiploDe(ByVal num As Integer, ByVal multiplo As Integer) As Boolean

        Dim esMultiplo As Boolean = False

        If num Mod multiplo = 0 Then
            esMultiplo = True
        End If

        Return esMultiplo

    End Function

    ''' <summary>
    ''' Calcula el factorial del numero indicado
    ''' </summary>
    ''' <param name="numero"></param>
    ''' <returns>Factorial del numero pasado</returns>
    ''' <remarks></remarks>
    Function factorial(ByVal numero As Integer) As Integer

        Dim producto = numero
        For i = numero - 1 To 1 Step -1
            producto *= i
        Next

        Return producto

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

    Function areaTriangulo(ByVal base As Double, ByVal altura As Double) As Double

        Return (base * altura) / 2

    End Function

    Function areaCuadrado(ByVal lado As Double) As Double
        Return lado * lado
    End Function

    Function areaRectangulo(ByVal base As Double, ByVal altura As Double) As Double
        Return base * altura
    End Function

    Function areaCirculo(ByVal radio As Double) As Double
        Return Math.PI * Math.Pow(radio, 2)
    End Function

    Function areaHexagono(ByVal lado As Double) As Double

        Dim perimetro As Double = (lado * 6)
        Dim apotema As Double = Math.Sqrt(Math.Pow(lado, 2) - (Math.Pow(lado / 2, 2)))

        Return (perimetro * apotema) / 2

    End Function

    Function ecuacion2Grado(ByVal a As Integer, ByVal b As Integer, ByVal c As Integer) As Double()
        If (Math.Pow(b, 2) - (4 * a * c)) >= 0 Then

            Dim soluciones(1) As Double

            soluciones(0) = ((-b) + Math.Sqrt(Math.Pow(b, 2) - (4 * a * c))) / (2 * a)

            soluciones(1) = ((-b) - Math.Sqrt(Math.Pow(b, 2) - (4 * a * c))) / (2 * a)

            Return soluciones
        Else
            MsgBox("La ecuacion no se puede resolver")
            Return Nothing
        End If
    End Function

    Function cuentaCifras(ByVal numero As Integer) As Integer

        If numero = 0 Then
            Return 1
        Else
            Dim contador As Integer

            While numero <= 0

                Math.truncate(numero / 10)
                contador += 1

            End While

            Return contador
        End If

    End Function

    Function esPrimo(ByVal numero As Integer) As Boolean

        Dim contador As Integer = 0

        For i = Math.Truncate(Math.Sqrt(numero)) To 1 Step -1
            If numero Mod i = 0 Then
                contador += 1
            End If
        Next

        If contador > 1 Then
            Return False
        Else
            Return True
        End If

    End Function

    Function cuentaPrimosEntre(ByVal numero1 As Integer, ByVal numero2 As Integer) As Integer

        If numero2 > numero1 Then

            Dim aux As Integer = numero1
            numero1 = numero2
            numero2 = aux

        End If


        Dim contador As Integer = 0

        For i = numero2 To numero1
            If esPrimo(i) Then
                contador += 1
            End If
        Next

        Return contador

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

    'True = byte, si unidadActual es mayor que unidadAPasar
    'false= bit, si unidadAPaar es mayor que unidadActual
    'Mejorar
    Function convertirUnidadInformacion(ByVal unidadActual As String, ByVal unidadAPasar As String, ByVal cantidad As Double, ByVal unidadActEsBit As Boolean, ByVal unidadAPasEsBit As Boolean) As Double

        Dim cantidadPasada As Double = 0

        Dim unidades() As String = {"B", "KB", "MB", "GB", "TB", "PB"}

        Dim posUnidadActual As Integer = 0
        Dim posUnidadAPasar As Integer = 0

        For i = 0 To unidades.Length - 1

            If unidades(i) = unidadActual Then
                posUnidadActual = i
            End If

            If unidades(i) = unidadAPasar Then
                posUnidadAPasar = i
            End If

        Next

        Dim bitAct As Integer = 0
        Dim bitPas As Integer = 0

        If unidadActEsBit Then
            bitAct = 3
        End If

        If unidadAPasEsBit Then
            bitPas = -3
        End If

        ''Si los dos estan activos, no es necesario sumar
        'If unidadActEsBit And unidadAPasEsBit Then
        '    bitAct = 0
        '    bitPas = 0
        'End If

        If posUnidadActual > posUnidadAPasar Then
            cantidadPasada = cantidad * Math.Pow(2, (10 * (posUnidadActual - posUnidadAPasar)) + (bitAct + bitPas))
        Else
            cantidadPasada = cantidad / Math.Pow(2, (10 * (posUnidadAPasar - posUnidadActual)) + (bitAct + bitPas))
        End If

        Return cantidadPasada

    End Function

    Function esPotenciaDe(ByVal pot, ByVal num) As Boolean

        Dim contador As Integer = -1

        Dim aux As Integer = num

        While num <> 0

            contador += 1
            num = Math.Truncate(num / pot)
        End While

        If Math.Pow(pot, contador) = aux Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function esNumeroEntero(ByVal num As String) As Boolean
        Dim n As Integer
        Dim esNum As Boolean = True
        Try
            n = Convert.ToInt32(num)
        Catch ex As Exception
            esNum = False
        End Try
        Return esNum
    End Function

    Function deStringADouble(ByVal cadena As String) As Double

        Return CDbl(Val(Replace(cadena, ",", ".")))

    End Function

End Module

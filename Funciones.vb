Module ModuleFunciones

    Sub main()

        'Console.WriteLine(esBisiesto(2013))



        Console.ReadLine()
    End Sub


    '''''''''''
    ''Numeros''
    '''''''''''

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

    ''''''''''
    ''Fechas''
    ''''''''''

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

            Return fecha.ToShortDateString

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
        Return Nothing
    End Function

    ''''''''''
    ''Letras''
    ''''''''''

    Function generaMayus() As Char

        Return Convert.ToChar(numAleatorioEntre(65, 90))

    End Function

    Function generaMinus() As Char

        Return Convert.ToChar(numAleatorioEntre(97, 122))

    End Function

    Function generaCaracteresEspeciales() As Char
        Dim caracteresEspeciales() As Char = {"!", "%", "&", "(", ")", "*", "/", "-", "#", "@"}

        Return caracteresEspeciales(numAleatorioEntre(0, caracteresEspeciales.Length - 1))

    End Function

    Function comprobarDNI(ByVal DNI As String) As Boolean

        Dim correcto As Boolean = False

        Try
            If DNI.Length() >= 8 And DNI.Length() <= 9 Then

                Const DIVISOR As Integer = 23
                Dim letrasNIF() As Char = {"T", "R", "W", "A", "G", "M", "Y",
                "F", "P", "D", "X", "B", "N", "J", "Z",
                "S", "Q", "V", "H", "L", "C", "K", "E"}


                Dim numDNI As Integer = Convert.ToInt32(DNI.Substring(0, DNI.Length - 1))

                Dim res As Integer = numDNI - (Math.Truncate(numDNI / DIVISOR) * DIVISOR)

                'Console.WriteLine(">> " & 

                Dim DNICalculado As String = Convert.ToString(numDNI) & letrasNIF(res)

                If DNICalculado.Equals(DNI) Then
                    correcto = True
                End If

            End If
        Catch ex As FormatException
            MsgBox("La cadena pasada es incorrecta")
        End Try

        Return correcto

    End Function

    '''''''''
    ''Array''
    '''''''''

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

    ''''''''''
    ''Matriz''
    ''''''''''

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

        For i = 0 To matriz.GetUpperBound(0) - 1
            For j = 0 To matriz.GetUpperBound(1) - 1
                Console.WriteLine(matriz(i, j))
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
    Function adyacentesA(ByRef matriz As Object, ByVal i As Integer, ByVal j As Integer) As Integer()
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
    Function devolverPosicionMatriz(ByRef matriz As Object, ByVal posicion As Integer, ByRef objeto As Object) As Integer()

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


    'graficos

    'Incluye positivos y negativos
    Sub escribirSoloNumeros(ByVal t As textbox, ByVal label As label, ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Not IsNumeric(e.KeyChar) And Convert.ToInt32(e.KeyChar) <> 8 And Convert.ToInt32(e.KeyChar) <> 45 Then
            e.Handled = True
            label.Text = "Debes teclear valores numericos"
        Else
            label.Text = ""
        End If
    End Sub

    Sub escribirSoloLetras(ByVal t As textbox, ByVal label As label, ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Not Char.IsLetter(e.KeyChar) And Convert.ToInt32(e.KeyChar) <> 8 And Convert.ToInt32(e.KeyChar) <> 164 Then
            e.Handled = True
            label.Text = "Debes teclear letras"
        Else
            label.Text = ""

        End If
    End Sub

    Sub limpiarCampos(ByRef campos() As textbox)

        For i = 0 To campos.Length - 1
            campos(i).Text = ""
        Next

    End Sub

    Sub limpiarCamposValorNulo(ByRef campos() As textbox, ByVal valorNulo As Integer)

        For i = 0 To campos.Length - 1
            campos(i).Text = Convert.ToString(valorNulo)
        Next

    End Sub

    Sub limpiarRadioButton(ByRef radios() As RadioButton)

        For i = 0 To radios.Length - 1
            radios(i).Checked = False
        Next

    End Sub

    Sub limpiarCheckButton(ByRef check() As CheckBox)

        For i = 0 To check.Length - 1
            check(i).checked = False
        Next

    End Sub

    Sub limpiarDataGridView(ByVal dgv As DataGridView)
        
        While dgv.ColumnCount <> 0
            dgv.Columns.RemoveAt(0)
        End While

    End Sub

    ' La propiedad AutoSizeRowsMode no debe estar activada
    Sub ajustaContenidoDataGridView(ByVal dgv As DataGridView)
        For i = 0 To dgv.RowCount
            dgv.Rows.Item(i).Height = (dgv.Size.Height / dgv.RowCount)
        Next
    End Sub

    'Rellena datos de una matriz a una tabla, debe poder convertirse en String
    'Tambien alinea los datos
    Sub rellenarDataGridView(ByVal dgv As DataGridView, ByRef matriz As Object)

        Dim aux(matriz.GetUpperBound(1)) As String

        dgv.ColumnCount = matriz.GetUpperBound(1) + 1

        For i = 0 To matriz.GetUpperBound(0)
            For j = 0 To matriz.GetUpperBound(1)
                aux(j) = Convert.ToString(matriz(i, j))
                'Alinea los datos
                dgv.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            Next
            dgv.Rows.Add(aux)
        Next

    End Sub

    Sub rellenaCabeceraDataGridView(ByVal dgv As DataGridView, ByRef nombres() As String)

        If nombres.Length = dgv.ColumnCount Then

            For i = 0 To dgv.ColumnCount - 1
                dgv.Columns(i).Name = nombres(i)
            Next
        Else
            MsgBox("La longitud de arrays de nombres es mayor que el numero de columnas")
        End If


    End Sub

    Sub estilodgv(ByVal dgv As DataGridView)

        dgv.ReadOnly = True ' celdas de solo lectura
        dgv.AllowUserToAddRows = False ' para que no se muestre la ultima fila en blanco
        dgv.ColumnHeadersVisible = False ' para que no se muestre el titulo de la columna
        dgv.RowHeadersVisible = False ' para que no se muestre la columna vacia de cada fila

        With dgv.ColumnHeadersDefaultCellStyle  ' formato por defecto de las cabeceras de columna
            .Font = New Font("Tahoma", 8, FontStyle.Bold)
            .ForeColor = Color.Black
            .BackColor = Color.Beige
            .SelectionForeColor = Color.Yellow
            .SelectionBackColor = Color.Blue
        End With

        dgv.RowTemplate.Height = 100   'altura de las filas
        dgv.EditMode = DataGridViewEditMode.EditProgrammatically
        dgv.SelectionMode = DataGridViewSelectionMode.CellSelect  'modos de seleccion muy importante FullRowSelect
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill ' para ajustarse el tamaño de las columnas a las dimensiones del datagridview
        dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells  ' para ajustar el tamaño de las filas a las dimensiones del datagridview
        dgv.AllowUserToResizeColumns = False ' no permitir redimensionar columnas
        dgv.AllowUserToResizeRows = False ' no permitir redimensionar filas

        With (dgv) 'para alternar dos colores entre las filas del datagridview
            .RowsDefaultCellStyle.BackColor = Color.Linen
            .AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise
        End With

        With dgv.DefaultCellStyle ' formato por defecto del datagridview
            .Font = New Font("Tahoma", 8)
            .ForeColor = Color.Black
            .BackColor = Color.Beige
            .SelectionForeColor = Color.Yellow
            .SelectionBackColor = Color.Silver
        End With
    End Sub

    Sub cargarYADaptarImagenPB(ByVal ptb As PictureBox, ByVal rutaImagen As String)
        ptb.Load(rutaImagen)
        ptb.SizeMode = PictureBoxSizeMode.StretchImage
    End Sub

    Sub rellenaListBox(ByVal ltb As ListBox, ByRef datos() As String)

        If ltb.Items.Count > 0 Then
            ltb.Items.Clear()
        End If

        For i = 0 To datos.Length - 1
            ltb.Items.Add(datos(i))
        Next

        ltb.SelectedIndex = 0

    End Sub

    Sub rellenaListBox(ByVal ltb As ListBox, ByRef datos() As Integer)

        If ltb.Items.Count > 0 Then
            ltb.Items.Clear()
        End If

        For i = 0 To datos.Length - 1
            ltb.Items.Add(datos(i))
        Next

        ltb.SelectedIndex = 0

    End Sub

    Sub rellenaListBox(ByVal ltb As ListBox, ByRef datos() As Double)

        If ltb.Items.Count > 0 Then
            ltb.Items.Clear()
        End If

        For i = 0 To datos.Length - 1
            ltb.Items.Add(datos(i))
        Next

        ltb.SelectedIndex = 0

    End Sub

    Sub rellenaListBox(ByVal ltb As ListBox, ByRef datos() As Object)

        If ltb.Items.Count > 0 Then
            ltb.Items.Clear()
        End If

        For i = 0 To datos.Length - 1
            ltb.Items.Add(datos(i))
        Next

        ltb.SelectedIndex = 0

    End Sub

    Sub rellenaComboBox(ByVal cmb As ComboBox, ByRef contenido() As String)

        cmb.Items.Clear()

        For i = 0 To contenido.Length - 1
            cmb.Items.Add(contenido(i))
        Next

        cmb.SelectedIndex = 0

    End Sub

End Module

Module Calculo


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

    'IVA=21
    Function precioSinIVA(ByVal valor As Double, ByVal IVA As Integer)

        Dim total As Double = (valor * 100) / (IVA + 100)

        Return total

    End Function

End Module

Module ValidacionesNumeros


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

End Module

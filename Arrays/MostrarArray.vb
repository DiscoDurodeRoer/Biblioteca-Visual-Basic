Module MostrarArray


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



End Module

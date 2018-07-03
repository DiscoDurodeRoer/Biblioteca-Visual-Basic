Module OperacionesArray

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

End Module

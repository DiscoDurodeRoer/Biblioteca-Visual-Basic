Module DivisionMatrizAArray

    ' ''' <summary>
    ' ''' Devuelve un array con los meses del año
    ' ''' </summary>
    ' ''' <returns>Un array con todos los meses del año</returns>
    ' ''' <remarks></remarks>
    'Function mesesAño() As String()
    '    Dim meses() As String = {"Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"}
    '    Return meses
    'End Function

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

End Module

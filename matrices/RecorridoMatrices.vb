Module RecorridoMatrices


    'Solo columnas
    Sub recorrerMatrizColumnas(ByVal matriz(,) As Object)

        For j = 0 To matriz.GetUpperBound(0) - 1

            For i = 0 To matriz.GetUpperBound(1) - 1

                'Ejecutar acciones

            Next

        Next

    End Sub

    'Solo columnas, desde el final hasta el principio
    Sub recorrerMatrizColumnasFinal(ByVal matriz(,) As Object)


        For j = matriz.GetUpperBound(0) - 1 To 0 Step -1


            For i = 0 To matriz.GetUpperBound(0) - 1


                'Acciones

            Next

        Next


    End Sub

    'En diagonal desde 0,0
    Sub recorrerMatrizDiagonalIzquierda(ByVal matriz(,) As Object)


        Dim aux As Integer

        'Recorrido desde 0,0 hasta 4,0, pasando por las diagonales
        For j = 0 To matriz.GetUpperBound(0) - 1

            aux = j

            For i = 0 To j

                'Acciones

                aux -= 1

            Next

        Next

        'Recorrido desde el limite,limite hasta el limite,1 , pasando por las diagonales
        For j = matriz.GetUpperBound(0) - 1 To 1 Step -1

            aux = j

            For i = matriz.GetUpperBound(0) - 1 To j Step -1

                'Acciones

                aux += 1

            Next

        Next

    End Sub

    'Igual que el anterior pero inverso
    Sub recorrerMatrizDiagonalIzquierdaInverso(ByVal matriz(,) As Object)


        Dim aux As Integer

        'Recorrido desde 0,0 hasta 4,0, pasando por las diagonales
        For j = matriz.GetUpperBound(0) - 1 To 0 Step -1

            aux = j

            For i = 0 To j

                'Acciones

                aux -= 1

            Next

        Next

        ''Recorrido desde el limite,limite hasta el limite,1 , pasando por las diagonales
        For j = 1 To matriz.GetUpperBound(0) - 1

            aux = j

            For i = matriz.GetUpperBound(0) - 1 To j Step -1

                'Acciones

                aux += 1

            Next

        Next

    End Sub

    Sub recorrerMatrizDiagonalDerecha(ByVal matriz(,) As Object)


        Dim aux As Integer
        Dim i As Integer

        For j = matriz.GetUpperBound(0) - 1 To 0 Step -1

            aux = j
            i = 0
            While Not aux = matriz.GetUpperBound(0)

                'Acciones

                aux += 1
                i += 1
            End While

        Next

        For j = matriz.GetUpperBound(0) - 1 To 1 Step -1

            aux = j
            i = 0

            While Not aux = matriz.GetUpperBound(0)

                'Acciones

                aux += 1
                i += 1
            End While

        Next

    End Sub


End Module

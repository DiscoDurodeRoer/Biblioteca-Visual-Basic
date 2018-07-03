Module Formulas

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

End Module

Public Class DNI

    'Atributos
    Private numeroDNI As Integer
    Private LetraDNI As Char
    Private DNIcompleto As String

    'Constantes
    Private LETRAS_NIF() As Char = {"T", "R", "W", "A", "G", "M", "Y",
    "F", "P", "D", "X", "B", "N", "J", "Z",
    "S", "Q", "V", "H", "L", "C", "K", "E"}

    Private Const DIVISOR As Integer = 23

    Sub New()
        generarDNIAleatorio()
    End Sub

    Sub New(ByVal numeroDNI As Integer)

        If (cuentaCifrasDNI(numeroDNI) >= 7 And cuentaCifrasDNI(numeroDNI) <= 8) Then

            Me.numeroDNI = numeroDNI
            letraNIF()
            completarDNI()
        Else

            numeroDNI = 0
            LetraDNI = ""
            DNIcompleto = ""

        End If

    End Sub

    Sub New(ByVal DNIcompleto As String)

        If (comprobarDNI(DNIcompleto)) Then

            numeroDNI = Convert.ToInt32(DNIcompleto.Substring(0, DNIcompleto.Length() - 1))
            LetraDNI = DNIcompleto.Chars(DNIcompleto.Length() - 1)
            completarDNI()

        Else

            numeroDNI = 0
            LetraDNI = ""
            DNIcompleto = ""

        End If


    End Sub

    Private Sub generarDNIAleatorio()

        numeroDNI = numAleatorioEntre_DNI(1000000, 99999999)

        letraNIF()

        completarDNI()

    End Sub


    Private Sub completarDNI()

        DNIcompleto = Convert.ToString(numeroDNI) + LetraDNI

    End Sub


    Private Sub letraNIF()

        If (cuentaCifrasDNI(numeroDNI) >= 7 And cuentaCifrasDNI(numeroDNI) <= 8) Then

            Dim res As Integer = numeroDNI - (numeroDNI / DIVISOR * DIVISOR)

            LetraDNI = LETRAS_NIF(res)

        Else
            MsgBox("El DNI debe ser de 7 u 8 cifras")
        End If
    End Sub

    Private Sub letraNIF(ByVal DNI As String)

        If (DNI.Length() >= 7 And DNI.Length() <= 8) Then

            Try

                Dim numDNI As Integer = Convert.ToInt32(DNI)
                Dim res As Integer = numDNI - (Math.Truncate(numDNI / DIVISOR) * DIVISOR)

                LetraDNI = LETRAS_NIF(res)
            Catch e As FormatException
                MsgBox("La cadena pasada es incorrecta")
            End Try
        Else
            MsgBox("El DNI debe ser de 7 u 8 cifras")
        End If

    End Sub

    Private Function comprobarDNI(ByVal DNI As String) As Boolean

        If (DNI.Length() >= 8 And DNI.Length() <= 9) Then
            Try

                Dim correcto As Boolean = False

                Dim numDNI As Integer = Convert.ToInt32(DNI.Substring(0, DNI.Length() - 1))
                Dim res As Integer = numDNI - (Math.Truncate(numDNI / DIVISOR) * DIVISOR)

                Dim DNICalculado As String = Convert.ToInt32(numDNI) & LETRAS_NIF(res)

                If (Replace(DNICalculado, "0", "").Equals(Replace(DNI, "0", ""))) Then
                    correcto = True
                End If

                Return correcto
            Catch e As FormatException
                MsgBox("La cadena pasada es incorrecta")
                Return False
            End Try
        Else
            MsgBox("El DNI debe ser de 8 u 9 cifras")
            Return False
        End If

    End Function


    Private Shared Function numAleatorioEntre_DNI(ByVal op1 As Integer, ByVal op2 As Integer) As Integer
        Randomize() ' inicializar la semilla  
        Return CLng((op1 - op2) * Rnd() + op2)
    End Function


    'Metodos estaticos

    Private Shared Function cuentaCifrasDNI(ByVal num As Integer) As Integer
        Dim contador As Integer = 0
        If num = 0 Then
            Return 1
        ElseIf num > 0 Then

            While (num > 0)
                num = Math.Floor(num / 10)
                contador += 1
            End While

            Return contador
        Else
            While (num < 0)
                num = Math.Floor(num / 10)
                contador += 1
            End While

            Return contador
        End If

    End Function

    Public Shared Function letraNIF(ByVal DNI As Integer) As Char

        If (cuentaCifrasDNI(DNI) >= 7 And cuentaCifrasDNI(DNI) <= 8) Then

            Dim res As Integer = DNI - (DNI / DIVISOR * DIVISOR)

            Dim letrasNIF() As Char = {"T", "R", "W", "A", "G", "M", "Y",
            "F", "P", "D", "X", "B", "N", "J", "Z",
            "S", "Q", "V", "H", "L", "C", "K", "E"}

            Return letrasNIF(res)
        Else
            MsgBox("El DNI debe ser de 7 u 8 cifras")
            Return ""
        End If


    End Function

    Public Shared Function devolverLetraNIF(ByVal DNI As String) As String

        If (DNI.Length() >= 7 And DNI.Length() <= 8) Then
            Try

                Dim numDNI As Integer = Convert.ToInt32(DNI)
                Dim res As Integer = numDNI - (Math.Truncate(numDNI / DIVISOR) * DIVISOR)

                Dim letrasNIF() As Char = {"T", "R", "W", "A", "G", "M", "Y",
                "F", "P", "D", "X", "B", "N", "J", "Z",
                "S", "Q", "V", "H", "L", "C", "K", "E"}

                Return DNI + letrasNIF(res)
            Catch e As FormatException
                MsgBox("La cadena pasada es incorrecta")
                Return Nothing
            End Try
        Else
            MsgBox("El DNI debe ser de 7 u 8 cifras")
            Return Nothing
        End If

    End Function

    Public Shared Function validarDNI(ByVal DNI As String) As Boolean

        If (DNI.Length() >= 8 And DNI.Length() <= 9) Then
            Try

                Const DIVISOR As Integer = 23

                Dim correcto As Boolean = False

                Dim numDNI As Integer = Convert.ToInt32(DNI.Substring(0, DNI.Length() - 1))
                Dim res As Integer = numDNI - (Math.Truncate(numDNI / DIVISOR) * DIVISOR)

                Dim letrasNIF() As Char = {"T", "R", "W", "A", "G", "M", "Y",
                                           "F", "P", "D", "X", "B", "N", "J", "Z",
                                           "S", "Q", "V", "H", "L", "C", "K", "E"}

                Dim DNICalculado As String = Convert.ToString(numDNI) + letrasNIF(res)
                If (Replace(DNICalculado, "0", "") = Replace(DNI.ToUpper, "0", "")) Then
                    correcto = True
                End If

                Return correcto
            Catch e As FormatException
                Return False
            End Try
        Else
            Return False
        End If

    End Function

    Public Shared Function generaDNIAleatorio() As String

        Dim numero As Integer = numAleatorioEntre_DNI(1000000, 99999999)

        Dim letra As Char = letraNIF(numero)

        Return Convert.ToString(numero) + letra

    End Function

    Public Shared Function generaDNIsAletorios(ByVal numeroDNIs As Integer) As String()

        Dim DNIs(numeroDNIs) As String

        Dim dni

        For i = 0 To DNIs.Length - 1

            Do

                dni = generaDNIAleatorio()

            Loop While (existeDNI(DNIs, dni))

            DNIs(i) = dni

        Next

        Return DNIs

    End Function

    Private Shared Function existeDNI(ByVal DNIs As String(), ByVal DNI As String) As Boolean

        Dim encontrado As Boolean = False

        Dim i As Integer = 0

        While (i < DNIs.Length And Not encontrado)

            If Not IsNothing(DNIs(i)) Then
                If (DNIs(i) = DNI) Then
                    encontrado = True
                End If
            End If

        End While

        Return encontrado

    End Function

    Public Shared Function Valida_CIF(ByVal valor As String) As Boolean
        Dim A As Integer
        Dim B As Integer
        Dim C As Integer
        Dim CIF As String
        Dim CIFDIGITO As String
        Dim i As Integer

        A = 0
        B = 0

        Valida_CIF = False

        If Len(valor) <> 9 Then 'el CIF debe tener 9 cifras
            Exit Function
        End If

        CIF = Mid(valor, 2, 7) 'se obtienen los dígitos centrales

        CIFDIGITO = Right(valor, 1) 'dígito de control

        For i = 1 To 6 Step 2
            A = A + Mid(CIF, i + 1, 1)   'Suma de posiciones pares
            C = 2 * Mid(CIF, i, 1)       'Doble de posiciones impares
            B = B + (C Mod 10) + Int(C / 10)   'Suma de digitos de doble de pares
        Next i
        'para obtener el cálculo de la cifra de la séptima posición que no se trata
        'en el bucle
        B = B + ((2 * Mid(CIF, 7, 1)) Mod 10) + Int((2 * Mid(CIF, 7, 1)) / 10)

        'se obtiene la unidad de la cifra total
        C = (10 - ((A + B) Mod 10)) Mod 10

        Dim Digito As String
        Dim letras() As String
        letras = {"J", "A", "B", "C", "D", "E", "F", "G", "H", "I"}
        Select Case (Left(valor, 1))
            'los cif que comienzan por estas letras deben terminar en una letra
            'concreta de la lista anterior
            Case "K", "P", "R", "Q", "S", "W" : Digito = letras(C)

                'los cif que comienzan por estas letras deben terminar en un dígito
            Case "A", "B", "E", "H", "J", "U", "V" : Digito = C

            Case "X", "Y", "Z"
                ' Error: es un NIE

                'para el resto de cif, la terminación puede ser un número o una letra
            Case Else
                If IsNumeric(CIFDIGITO) Then
                    Digito = C
                Else
                    Digito = letras(C)
                End If
        End Select
        Valida_CIF = (CIFDIGITO = Digito)

    End Function

End Class

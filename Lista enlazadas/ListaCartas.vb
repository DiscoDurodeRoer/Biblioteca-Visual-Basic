Public Class ListaCartas

    Public Class Nodo

        'Atributos
        Public dato As Carta
        Public sig As Nodo

        'Métodos constructores
        Public Sub New(ByVal palo As Char, ByVal numero As Integer)
            sig = Nothing
            dato = New Carta(palo, numero)
        End Sub

        Public Sub New(ByVal m As Carta)
            sig = Nothing
            dato = m
        End Sub
    End Class


    '*********************************************************
    '*********************************************************
    Private primero As Nodo
    Public palos() As Char = {"E", "O", "B", "C"}
    Public numeros() As Integer = {1, 2, 3, 4, 5, 6, 7, 10, 11, 12}

    '---------------------------------------------
    Public Sub New()

        Lista_Vacia()
        rellenaCartas()

    End Sub

    '---------------------------------------------
    Public Sub Lista_Vacia()

        primero = Nothing

    End Sub


    '---------------------------------------------
    Public Function Esta_Vacia() As Boolean

        Return primero Is Nothing

    End Function

    '---------------------------------------------
    Public Sub Insertar_Primero(ByVal m As Carta)

        Dim nuevo As New Nodo(m)

        If (primero Is Nothing) Then
            primero = nuevo
        Else
            nuevo.sig = primero
            primero = nuevo
        End If

    End Sub


    '---------------------------------------------
    Public Sub Insertar_Ultimo(ByVal m As Carta)

        Dim aux As New Nodo(m)
        Dim rec_aux As Nodo

        If (primero Is Nothing) Then
            aux.sig = primero
            primero = aux
        Else
            rec_aux = primero
            While (Not rec_aux.sig Is Nothing)
                rec_aux = rec_aux.sig
            End While
            rec_aux.sig = aux
        End If

    End Sub


    '---------------------------------------------
    Public Sub Quitar_Primero()

        Dim aux As Nodo

        If (Not Esta_Vacia()) Then
            aux = primero
            primero = primero.sig
            aux = Nothing  'Lo marcamos para el recolector de basura   
        End If

    End Sub

    '---------------------------------------------
    Public Sub Quitar_Ultimo()

        Dim aux As Nodo = primero
        If (aux.sig Is Nothing) Then
            Lista_Vacia()
        End If
        If (Not Esta_Vacia()) Then
            aux = primero
            While (Not aux.sig.sig Is Nothing)
                aux = aux.sig
            End While
            aux.sig = Nothing
        End If

    End Sub

    '---------------------------------------------
    Public Function Dev_Ultimo() As Carta

        Dim elemen As Carta = Nothing
        Dim aux As Nodo = Nothing
        If (Not Esta_Vacia()) Then
            aux = primero
            While (Not aux.sig Is Nothing)
                aux = aux.sig
            End While
            elemen = aux.dato
        End If
        Return elemen

    End Function

    '---------------------------------------------
    Public Function Dev_Primero() As Carta

        Dim elemen As Carta = Nothing
        If (Not Esta_Vacia()) Then
            elemen = primero.dato
        End If
        Return elemen

    End Function


    '---------------------------------------------
    Public Function Cuantos_Elementos() As Integer

        Dim aux As Nodo = Nothing
        Dim i As Integer = 0
        aux = primero

        While (Not aux Is Nothing)
            aux = aux.sig
            i += 1
        End While
        Return i

    End Function

    Public Sub Insertar_posicion(ByVal pos As Integer, ByVal m As Carta)

        If (primero Is Nothing) Then ' (Esta_vacio()) si no tiene ningúna muestra:
            primero = New Nodo(m)
        Else

            Dim nuevo As New Nodo(m) 'creamos un nuevo nodo con la muestra
            ' hay que tener en cuenta que la posicion sea la primera u otra distinta
            If (pos = 0) Then
                'si es la primera sólo hay que añadirle el nuevo nodo a primero
                nuevo.sig = primero 'Insertar_Primero(muestra)
                primero = nuevo
            Else
                Dim cont As Integer = 0
                Dim aux As Nodo = primero
                Dim ant As Nodo = Nothing
                'para llegar a la posición y añadir la nueva muestra:
                While ((Not aux Is Nothing) And (cont <> pos)) 'mientras que aux no apunte a null
                    ant = aux
                    aux = aux.sig
                    cont += 1
                End While
                ant.sig = nuevo
                nuevo.sig = aux
            End If
        End If
    End Sub

    '------------------------------------------
    Public Sub Borra_Posicion(ByVal pos As Integer)
        ' Las posiciones empiezan desde 1.
        Dim aux, ant As Nodo
        Dim cont As Integer = 1

        aux = primero
        ant = Nothing
        While (Not aux Is Nothing)
            If (pos = cont) Then
                If (ant Is Nothing) Then 'Primero
                    primero = primero.sig
                    aux = Nothing
                Else
                    ant.sig = aux.sig
                    aux = Nothing  'Para que borre efectivamente el nodo.
                End If
            Else
                ant = aux
                aux = aux.sig
                cont += 1
            End If
        End While

    End Sub

    Public Function devuelveCartaPosicion(ByVal posicion As Integer) As Carta


        Dim elemen As Carta = Nothing
        Dim aux As Nodo = Nothing
        Dim contador As Integer = 1
        If (Not Esta_Vacia()) Then
            aux = primero
            'Si es el primero, lo devuelve
            If posicion = 1 Then
                Return Dev_Primero()
            Else
                'Si es el ultimo lo devuelve
                If posicion = Cuantos_Elementos() Then
                    Return Dev_Ultimo()
                Else
                    While (Not aux.sig Is Nothing And contador <> posicion)
                        aux = aux.sig
                        contador += 1
                    End While
                End If
            End If

            elemen = aux.dato
        End If

        'Si ha llegado al final y no ha encontrado la posicion, es que no se encuentra en la lista
        If IsNothing(aux.sig) Then
            Return Nothing
        Else
            Return elemen
        End If


    End Function

    Public Function buscaYEliminaCarta(ByVal palo As Char, ByVal numero As Integer) As Carta

        Dim aux As Nodo = primero
        If Esta_Vacia() Then
            Return Nothing
        Else

            'Busco la carta, si la encuentra la borra
            Dim encontrado As Boolean = False
            Dim contador As Integer = 1
            While Not IsNothing(aux.sig) And Not encontrado

                If aux.dato.PropNumero = numero And aux.dato.PropPalo = palo Then
                    encontrado = True
                Else
                    contador += 1
                    aux = aux.sig
                End If
            End While

            'Puede pasar, que la carta no se encuentre
            If Not IsNothing(aux) Then
                Console.WriteLine("Con " & contador)
                Borra_Posicion(contador)
                Return aux.dato
            Else
                Return Nothing
            End If

        End If

    End Function


    Public Function devYBorrarPrimero() As Carta

        Dim m As Carta = Dev_Primero()
        Quitar_Primero()

        Return m

    End Function


    Public Sub rellenaCartas()


        For i = 0 To palos.Length - 1
            For j = 0 To numeros.Length - 1
                Insertar_Ultimo(New Carta(palos(i), numeros(j)))
            Next
        Next

    End Sub

End Class

Public Class ListaCaracteres

    Public Class Nodo

        'Atributos
        Public dato As Char
        Public sig As Nodo

        'Métodos constructores       

        Public Sub New(ByVal m As Char)
            sig = Nothing
            dato = m
        End Sub
    End Class


    '*********************************************************
    '*********************************************************
    Private primero As Nodo

    '---------------------------------------------
    Public Sub New()

        Lista_Vacia()

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
    Public Sub Insertar_Primero(ByVal m As Char)

        Dim nuevo As New Nodo(m)

        If (primero Is Nothing) Then
            primero = nuevo
        Else
            nuevo.sig = primero
            primero = nuevo
        End If

    End Sub


    '---------------------------------------------
    Public Sub Insertar_Ultimo(ByVal m As Char)

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
    Public Function Dev_Ultimo() As Char

        Dim elemen As Char = Nothing
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
    Public Function Dev_Primero() As Char

        Dim elemen As Char = Nothing
        If (Not Esta_Vacia()) Then
            elemen = primero.dato
        End If
        Return elemen

    End Function

    Public Sub meterTexto(ByVal texto As String)

        Dim elemen As Char = Nothing

        Dim caracteres() As Char = texto.ToCharArray()


        For i = 0 To caracteres.Length - 1

            Insertar_Ultimo(caracteres(i))

        Next

    End Sub
    

End Class

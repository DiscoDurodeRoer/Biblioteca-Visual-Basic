Public Class Carta

    'Atributos
    Private palo As Char
    Private numero As Integer

    'Properties
    Public Property PropPalo As Char
        Get
            Return palo
        End Get
        Set(ByVal value As Char)
            palo = value
        End Set
    End Property

    Public Property PropNumero As Integer
        Get
            Return numero
        End Get
        Set(ByVal value As Integer)
            numero = value
        End Set
    End Property

    'Constructores
    Sub New(ByVal palo As Char, ByVal numero As Integer)
        PropPalo = palo
        PropNumero = numero
    End Sub


End Class

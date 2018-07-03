Module EnlaceFichero

    Sub abrirEnlace(ByVal enlace As String)

        Process.Start(enlace)

    End Sub

    Sub abrirFichero(ByVal fichero As String)

        Process.Start(fichero)

    End Sub

End Module

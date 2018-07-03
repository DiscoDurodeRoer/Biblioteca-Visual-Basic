Imports System.Windows.Forms
Module Imagenes

    Sub cargarYADaptarImagenPB(ByVal ptb As PictureBox, ByVal rutaImagen As String)
        ptb.Load(rutaImagen)
        ptb.SizeMode = PictureBoxSizeMode.StretchImage
    End Sub

End Module

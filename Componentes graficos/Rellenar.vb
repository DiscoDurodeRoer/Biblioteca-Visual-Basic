Imports System.Windows.Forms
Imports System.Drawing
Module Rellenar

    'Funciones necesarias

    ''' <summary>
    ''' Devuelve un array con los meses del año
    ''' </summary>
    ''' <returns>Un array con todos los meses del año</returns>
    ''' <remarks></remarks>
    Function mesesAño() As String()
        Dim meses() As String = {"Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"}
        Return meses
    End Function


    'Funciones Rellenar

    Sub rellenaListBox(ByVal ltb As ListBox, ByRef datos() As String)

        If ltb.Items.Count > 0 Then
            ltb.Items.Clear()
        End If

        For i = 0 To datos.Length - 1
            ltb.Items.Add(datos(i))
        Next

        ltb.SelectedIndex = 0

    End Sub

    Sub rellenaListBox(ByVal ltb As ListBox, ByRef datos() As Integer)

        If ltb.Items.Count > 0 Then
            ltb.Items.Clear()
        End If

        For i = 0 To datos.Length - 1
            ltb.Items.Add(datos(i))
        Next

        ltb.SelectedIndex = 0

    End Sub

    Sub rellenaListBox(ByVal ltb As ListBox, ByRef datos() As Double)

        If ltb.Items.Count > 0 Then
            ltb.Items.Clear()
        End If

        For i = 0 To datos.Length - 1
            ltb.Items.Add(datos(i))
        Next

        ltb.SelectedIndex = 0

    End Sub

    Sub rellenaListBox(ByVal ltb As ListBox, ByRef datos() As Object)

        If ltb.Items.Count > 0 Then
            ltb.Items.Clear()
        End If

        For i = 0 To datos.Length - 1
            ltb.Items.Add(datos(i))
        Next

        ltb.SelectedIndex = 0

    End Sub

    Sub rellenaComboBox(ByVal cmb As ComboBox, ByRef contenido() As String)

        cmb.Items.Clear()

        For i = 0 To contenido.Length - 1
            cmb.Items.Add(contenido(i))
        Next

        cmb.SelectedIndex = 0

    End Sub

    Sub rellenaComboBoxDias(ByVal cmb As ComboBox, ByVal maximo As Integer)

        cmb.Items.Clear()

        cmb.Items.Add("--Elije dia--")

        Dim meses() As String = mesesAño()

        For i = 1 To maximo
            cmb.Items.Add(i)
        Next

        cmb.SelectedIndex = 0

    End Sub

    Sub rellenaComboBoxMeses(ByVal cmb As ComboBox)

        cmb.Items.Clear()

        cmb.Items.Add("--Elije mes--")

        Dim meses() As String = mesesAño()

        For i = 0 To mesesAño.Length - 1
            cmb.Items.Add(meses(i))
        Next

        cmb.SelectedIndex = 0

    End Sub

    Sub rellenaComboBoxAnios(ByVal cmb As ComboBox, ByVal anioMinimo As Integer, ByVal anioMaximo As Integer)

        cmb.Items.Clear()

        cmb.Items.Add("--Elije Año--")

        For i = anioMinimo To anioMaximo
            cmb.Items.Add(i)

        Next

        cmb.SelectedIndex = 0

    End Sub

    Sub rellenaComboBoxHora(ByVal cmb As ComboBox)

        cmb.Items.Clear()

        cmb.Items.Add("--Elije Hora--")

        For i = 0 To 23
            cmb.Items.Add(i)

        Next

        cmb.SelectedIndex = 0


    End Sub

    Sub rellenaComboBoxMinutos(ByVal cmb As ComboBox)

        cmb.Items.Clear()

        cmb.Items.Add("--Elije Minuto--")

        For i = 0 To 59
            cmb.Items.Add(i)

        Next

        cmb.SelectedIndex = 0

    End Sub

    Sub rellenaComboBoxSegundos(ByVal cmb As ComboBox)

        cmb.Items.Clear()

        cmb.Items.Add("--Elije Segundo--")

        For i = 0 To 59
            cmb.Items.Add(i)

        Next

        cmb.SelectedIndex = 0

    End Sub

End Module

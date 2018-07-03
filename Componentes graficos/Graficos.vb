Imports System.Windows.Forms
Imports System.Drawing
Module Graficos

    'Incluye positivos y negativos
    Sub escribirSoloNumeros(ByVal t As TextBox, ByVal label As Label, ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Not IsNumeric(e.KeyChar) And Convert.ToInt32(e.KeyChar) <> 8 And Convert.ToInt32(e.KeyChar) <> 45 Then
            e.Handled = True
            label.Text = "Debes teclear valores numericos"
        Else
            label.Text = ""
        End If
    End Sub


    Sub escribirSoloNumerosPositivos(ByVal t As TextBox, ByVal label As Label, ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Not IsNumeric(e.KeyChar) And Convert.ToInt32(e.KeyChar) <> 8 Then
            e.Handled = True
            label.Text = "Debes teclear valores numericos"
        Else
            label.Text = ""
        End If
    End Sub

    Sub escribirSoloNumerosNegativos(ByVal t As TextBox, ByVal label As Label, ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Not IsNumeric(e.KeyChar) And Convert.ToInt32(e.KeyChar) <> 8 And Convert.ToInt32(e.KeyChar) <> 45 Then
            e.Handled = True
            label.Text = "Debes teclear valores numericos"
        Else
            label.Text = ""
        End If
    End Sub

    Sub escribirSoloLetras(ByVal t As TextBox, ByVal label As Label, ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Not Char.IsLetter(e.KeyChar) And Convert.ToInt32(e.KeyChar) <> 8 And Convert.ToInt32(e.KeyChar) <> 164 Then
            e.Handled = True
            label.Text = "Debes teclear letras"
        Else
            label.Text = ""

        End If
    End Sub

    Sub escribirSoloLetrasYespacios(ByVal t As TextBox, ByVal label As Label, ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Not Char.IsLetter(e.KeyChar) And Asc(e.KeyChar) <> 8 And Asc(e.KeyChar) <> 164 And Asc(e.KeyChar) <> 32 Then
            e.Handled = True
            label.Text = "Debes teclear letras o espacios"
        Else
            label.Text = ""

        End If
    End Sub

    Sub soloDoubles(ByVal txt As TextBox, ByVal etiquetaError As Label, ByVal numeroEnteros As Integer, ByVal numeroDecimales As Integer, ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

        Dim text As String = txt.Text
        Dim caracter As Char = e.KeyChar

        Dim punto As Boolean = False
        If text.IndexOf(".") = -1 Then
            punto = False
        Else
            punto = True
        End If

        If txt.Text(0) = "0" And Not punto Then
            etiquetaError.Text = "No puedes añadir mas ceros a la izquierda"
        Else

            If (txt.TextLength = 0 And (Not Char.IsDigit(caracter))) Then
                e.Handled = True
                etiquetaError.Text = "El primer digito no puede ser un punto"
            Else
                If (Not Char.IsDigit(caracter) And (text.Contains(".") Or Asc(caracter) <> 46) And Asc(caracter) <> 8) Then
                    e.Handled = True
                    etiquetaError.Text = "Solo numeros y puntos"
                Else
                    If (punto) Then
                        If text.Substring(text.IndexOf(".")).Length() > numeroDecimales And Asc(caracter) <> 8 Then
                            e.Handled = True
                            etiquetaError.Text = "No puedes escribir mas decimales"
                        Else
                            etiquetaError.Text = ""
                        End If
                    End If
                End If
            End If
        End If

    End Sub

    Sub soloDoubles(ByVal txt As TextBox, ByVal etiquetaError As Label, ByVal numeroDecimales As Integer, ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

        Dim text As String = txt.Text
        Dim caracter As Char = e.KeyChar

        Dim punto As Boolean = False
        If text.IndexOf(".") = -1 Then
            punto = False
        Else
            punto = True
        End If

        If (txt.TextLength = 0 And (Not Char.IsDigit(caracter) Or caracter = "0")) Then
            e.Handled = True
            etiquetaError.Text = "El primer digito no puede ser un punto o un 0"
        Else
            If (Not Char.IsDigit(caracter) And (text.Contains(".") Or Asc(caracter) <> 46) And Asc(caracter) <> 8) Then
                e.Handled = True
                etiquetaError.Text = "Solo numeros y puntos"
            Else
                If (punto) Then
                    If text.Substring(text.IndexOf(".")).Length() > numeroDecimales And Asc(caracter) <> 8 Then
                        e.Handled = True
                        etiquetaError.Text = "No puedes escribir mas decimales"
                    Else
                        etiquetaError.Text = ""
                    End If
                End If
            End If
        End If

    End Sub


    Sub limpiarCampos(ByRef campos() As TextBox)

        For i = 0 To campos.Length - 1
            campos(i).Text = ""
        Next

    End Sub

    Sub limpiarCamposValorNulo(ByRef campos() As TextBox, ByVal valorNulo As Integer)

        For i = 0 To campos.Length - 1
            campos(i).Text = Convert.ToString(valorNulo)
        Next

    End Sub

    Sub limpiarRadioButton(ByRef radios() As RadioButton)

        For i = 0 To radios.Length - 1
            radios(i).Checked = False
        Next

    End Sub

    Sub limpiarCheckButton(ByRef check() As CheckBox)

        For i = 0 To check.Length - 1
            check(i).Checked = False
        Next

    End Sub

    Sub limpiarDataGridView(ByVal dgv As DataGridView)

        While dgv.ColumnCount <> 0
            dgv.Columns.RemoveAt(0)
        End While

    End Sub

    ' La propiedad AutoSizeRowsMode no debe estar activada
    Sub ajustaContenidoDataGridView(ByVal dgv As DataGridView)
        For i = 0 To dgv.RowCount
            dgv.Rows.Item(i).Height = (dgv.Size.Height / dgv.RowCount)
        Next
    End Sub

    'Rellena datos de una matriz a una tabla, debe poder convertirse en String
    'Tambien alinea los datos
    Sub rellenarDataGridView(ByVal dgv As DataGridView, ByRef matriz As Object)

        Dim aux(matriz.GetUpperBound(1)) As String

        dgv.ColumnCount = matriz.GetUpperBound(1) + 1

        For i = 0 To matriz.GetUpperBound(0)
            For j = 0 To matriz.GetUpperBound(1)
                aux(j) = Convert.ToString(matriz(i, j))
                'Alinea los datos
                dgv.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            Next
            dgv.Rows.Add(aux)
        Next

    End Sub

    Function rellenaMatrizDGV(ByVal dgv As DataGridView) As Object(,)

        Dim matriz(dgv.RowCount - 1, dgv.ColumnCount - 1) As Object


        Dim i As Integer = 0
        Dim j As Integer = 0


        For Each row As DataGridViewRow In dgv.Rows
            For Each cell As DataGridViewCell In row.Cells
                If Not String.IsNullOrEmpty(cell.Value) Then
                    ' Check if the string only consists of spaces
                    matriz(i, j) = cell.Value

                End If

                If dgv.ColumnCount - 1 = j Then
                    j = 0
                    i += 1
                Else
                    j += 1
                End If

            Next
        Next

        Return matriz

    End Function

    Sub rellenarDataGridViewBD(ByVal dgv As DataGridView, ByVal dataset As DataSet, ByVal nombreTabla As String)

        dgv.ColumnCount = dataset.Tables(nombreTabla).Columns.Count

        For Each row As DataRow In dataset.Tables(nombreTabla).Rows

            dgv.Rows.Add(row.ItemArray)

        Next

    End Sub

    Sub sacarDatoDataGrigView(ByVal t As TextBox, ByVal dgv As DataGridView)
        Try
            Dim celda As String
            celda = dgv.CurrentCell.Value.ToString
            t.Text = celda
        Catch ex As NullReferenceException
            MsgBox("Celda no Valida")
        End Try

    End Sub

    'Debe usarse despues de rellenar la tabla
    Sub rellenaCabeceraDataGridView(ByVal dgv As DataGridView, ByRef nombres() As String)

        If nombres.Length = dgv.ColumnCount Then

            For i = 0 To dgv.ColumnCount - 1
                dgv.Columns(i).Name = nombres(i)
                dgv.Columns(i).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            Next
        Else
            MsgBox("La longitud de arrays de nombres es mayor que el numero de columnas")
        End If


    End Sub

    Function IsDataGridViewEmpty(ByRef dataGridView As DataGridView) As Boolean
        Dim isEmpty As Boolean
        isEmpty = True
        'For Each row As DataGridViewRow In dataGridView.Rows
        '    For Each cell As DataGridViewCell In row.Cells
        '        If Not String.IsNullOrEmpty(cell.Value) Then
        '            ' Check if the string only consists of spaces
        '            If Not String.IsNullOrEmpty(Trim(cell.Value.ToString())) Then
        '                isEmpty = False
        '                Exit For
        '            End If
        '        End If
        '    Next
        'Next

        'dataGridView.RowCount

        Return isEmpty
    End Function

    Sub estilodgv(ByVal dgv As DataGridView)

        dgv.ReadOnly = True ' celdas de solo lectura
        dgv.AllowUserToAddRows = False ' para que no se muestre la ultima fila en blanco
        dgv.ColumnHeadersVisible = False ' para que no se muestre el titulo de la columna
        dgv.RowHeadersVisible = False ' para que no se muestre la columna vacia de cada fila

        With dgv.ColumnHeadersDefaultCellStyle  ' formato por defecto de las cabeceras de columna
            .Font = New Font("Tahoma", 8, FontStyle.Bold)
            .ForeColor = Color.Black
            .BackColor = Color.Beige
            .SelectionForeColor = Color.Yellow
            .SelectionBackColor = Color.Blue
        End With

        dgv.RowTemplate.Height = 100   'altura de las filas
        dgv.EditMode = DataGridViewEditMode.EditProgrammatically
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect  'modos de seleccion muy importante FullRowSelect
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill ' para ajustarse el tamaño de las columnas a las dimensiones del datagridview
        dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells  ' para ajustar el tamaño de las filas a las dimensiones del datagridview
        dgv.AllowUserToResizeColumns = False ' no permitir redimensionar columnas
        dgv.AllowUserToResizeRows = False ' no permitir redimensionar filas

        With (dgv) 'para alternar dos colores entre las filas del datagridview
            .RowsDefaultCellStyle.BackColor = Color.Linen
            .AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise
        End With

        With dgv.DefaultCellStyle ' formato por defecto del datagridview
            .Font = New Font("Tahoma", 8)
            .ForeColor = Color.Black
            .BackColor = Color.Beige
            .SelectionForeColor = Color.Yellow
            .SelectionBackColor = Color.Silver
        End With
    End Sub

    Sub cargarYADaptarImagenPB(ByVal ptb As PictureBox, ByVal rutaImagen As String)
        ptb.Load(rutaImagen)
        ptb.SizeMode = PictureBoxSizeMode.StretchImage
    End Sub

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

    Sub centrarDatosDataGridView(ByVal dgv As DataGridView)

        For i = 0 To dgv.ColumnCount - 1
            For j = 0 To dgv.RowCount - 1
                dgv.Rows(j).Cells(i).Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            Next
        Next

    End Sub

    'Calcula el total de cantidad de una linea
    Function sumaColumnas(ByVal dgv As DataGridView, ByVal columna As Integer) As Double

        Dim suma As Double = 0.0

        Dim valorActual As Double

        For i = 0 To dgv.RowCount - 1

            valorActual = CDbl(Val(dgv.Rows(i).Cells(columna).Value))

            suma += valorActual

        Next

        Return suma

    End Function

    Sub cambiaFechasCientificas(ByVal dgv As DataGridView, ByVal columna As Integer)

        For i = 0 To dgv.RowCount - 1

            Dim fecha As Date = deNumeroAFecha(dgv.Rows(i).Cells(columna).Value)

            dgv.Rows(i).Cells(columna).Value = fecha

        Next

    End Sub

    Sub cambiaHorasCientificas(ByVal dgv As DataGridView, ByVal columna As Integer)

        For i = 0 To dgv.RowCount - 1

            Dim hora As Date = deNumeroAHora(dgv.Rows(i).Cells(columna).Value)

            dgv.Rows(i).Cells(columna).Value = hora.ToString("HH:mm:ss")

        Next

    End Sub

    'En la tabla, no en la base de datos
    Function existeElemento(ByVal dgv As DataGridView, ByVal id As Integer, ByVal indiceColumna As Integer) As Boolean

        Dim existe As Boolean = False

        For i = 0 To dgv.RowCount - 1

            If dgv.Rows(i).Cells(indiceColumna).Value = id Then
                existe = True
            End If

        Next

        Return existe

    End Function

    Public Function DataGridViewToDataTable(ByVal dtg As DataGridView,
        Optional ByVal DataTableName As String = "myDataTable") As DataTable
        Try
            Dim dt As New DataTable(DataTableName)
            Dim row As DataRow
            Dim TotalDatagridviewColumns As Integer = dtg.ColumnCount - 1
            'Add Datacolumn
            For Each c As DataGridViewColumn In dtg.Columns
                Dim idColumn As DataColumn = New DataColumn()
                idColumn.ColumnName = c.Name
                dt.Columns.Add(idColumn)
            Next
            'Now Iterate thru Datagrid and create the data row
            For Each dr As DataGridViewRow In dtg.Rows
                'Iterate thru datagrid
                row = dt.NewRow 'Create new row
                'Iterate thru Column 1 up to the total number of columns
                For cn As Integer = 0 To TotalDatagridviewColumns
                    row.Item(cn) = IfNullObj(dr.Cells(cn).Value) ' This Will handle error datagridviewcell on NULL Values
                Next
                'Now add the row to Datarow Collection
                dt.Rows.Add(row)
            Next
            'Now return the data table
            Return dt
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Function IfNullObj(ByVal o As Object, Optional ByVal DefaultValue As String = "") As String
        Dim ret As String = ""
        Try
            If o Is DBNull.Value Then
                ret = DefaultValue
            Else
                ret = o.ToString
            End If
            Return ret
        Catch ex As Exception
            Return ret
        End Try
    End Function


    Sub abrirEnlace(ByVal enlace As String)

        Process.Start(enlace)

    End Sub

    Sub abrirFichero(ByVal fichero As String)

        Process.Start(fichero)

    End Sub

    Sub reproducirFicheroWav(ByVal rutaFicheroWav As String)

        My.Computer.Audio.Play(ruta)

    End Sub
    
End Module

Imports System.Windows.Forms
Imports System.Drawing
Module DGV

    'Funciones necesarias

    Function deNumeroAFecha(ByVal numero As Integer) As Date

        Try
            Dim anio As Integer = Math.Truncate(numero / 10000)

            Dim mes As Integer = Math.Truncate(((numero Mod 10000) / 100))

            Dim dia As Integer = numero Mod 100

            Dim fecha As New Date(anio, mes, dia)

            Return fecha

        Catch ex As ArgumentOutOfRangeException
            MsgBox("Error, la fecha no es correcta")
            Return Nothing
        End Try

    End Function

    Function deNumeroAHora(ByVal numero As Integer) As Date

        Dim hora As Integer = Math.Truncate(numero / 10000)


        Dim minuto As Integer = Math.Truncate((numero Mod 10000) / 100)


        Dim segundo As Integer = numero Mod 100


        Dim fecha As New Date(1, 1, 1, hora, minuto, segundo)

        Return fecha.ToString("HH:mm:ss")

    End Function

    'Funciones DGV

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

    Sub centrarCabeceraDataGridView(ByVal dgv As DataGridView)

        For i = 0 To dgv.ColumnCount - 1
            dgv.Columns(i).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next

    End Sub

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
        dgv.SelectionMode = DataGridViewSelectionMode.CellSelect  'modos de seleccion muy importante FullRowSelect
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

    Sub seleccionar1ºElemento(ByVal dgv As DataGridView)


        If dgv.RowCount <> 0 Then

            dgv.CurrentCell = dgv.Item(2, 0)

        End If

    End Sub






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



End Module

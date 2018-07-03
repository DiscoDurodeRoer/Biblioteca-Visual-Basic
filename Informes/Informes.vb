Module Informes

    'columnas = nombre de las columnas
    'Tipos (String, int32, etc.)
    Sub rellenaColumnasDataTable(ByRef dt As DataTable, ByVal columnas() As String, ByVal tipos() As String)

        For i = 0 To columnas.Length - 1

            dt.Columns.Add(columnas(i), Type.GetType("System." & tipos(i)))

        Next

    End Sub


    'Debe existir el objeto CrystalReport (renombrar el fichero .rpt con cr)
    'El dataset debe contener los cmpos que quieren ser visualizamos, no mas
    'El nombre del elemento es CrystalReportViewer
    'Recuerda que debes tener el resto de elementos creados
    'nombreTablaDS debe tener el mismo nombre que el datatable creado en el dataset.xsd
    Sub informe(ByVal ds As DataSet, ByVal dt As DataTable, ByVal nombreTablaDS As String)

        Dim tmp As DataTable = ds.Tables(nombreTablaDS)

        For Each row As DataRow In tmp.Rows
            dt.Rows.Add(row.ItemArray)
        Next

        Dim miReporte As New cr
        miReporte.Database.Tables(nombreTablaDS).SetDataSource(dt)

        CrystalReportViewer.ReportSource = miReporte

    End Sub

    'Debe existir el objeto CrystalReport (renombrar el fichero .rpt con cr)
    'El dataset debe contener los cmpos que quieren ser visualizamos, no mas
    'El nombre del elemento es CrystalReportViewer
    'Recuerda que debes tener el resto de elementos creados
    'Los valores que se pasan deben ser creados en el cr (campos por parametro) deben coincidir con el tipo y en orden
    'nombreTablaDS debe tener el mismo nombre que el datatable creado en el dataset.xsd
    Sub informe(ByVal ds As DataSet, ByVal dt As DataTable, ByVal nombreTablaDS As String, ByVal valores() As Object)

        Dim tmp As DataTable = ds.Tables(nombreTablaDS)

        Dim fecha As Date
        Dim hora As Date

        For Each row As DataRow In tmp.Rows

            fecha = deNumeroAFecha(row("fecha"))
            hora = deNumeroAHora(row("hora"))

            dt.Rows.Add(fecha.ToString("dd/MM/yyyy"), hora.ToString("HH:mm:ss"), row("saldo_total"), row("nombre"))
        Next

        Dim miReporte As New crCaja



        miReporte.Database.Tables(nombreTablaDS).SetDataSource(dt)

        crvCaja.ReportSource = miReporte

        For i = 0 To valores.Length - 1

            miReporte.SetParameterValue(i, valores(i))

        Next


    End Sub

End Module

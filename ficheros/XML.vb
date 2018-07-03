Imports System.Xml
Imports System.IO
Imports System.Xml.Serialization
Imports System.Windows.Forms
Imports System.Drawing

Module XML

    'Muestra el contenido de los nodos especificados en msgbox
    Sub leerXMLMensajes(ByVal ficheroXML As String, ByVal nombreNodo As String)

        Dim xmldoc As New XmlDataDocument()
        Dim xmlnode As XmlNodeList

        Dim str As String
        Dim fs As New FileStream(ficheroXML, FileMode.Open, FileAccess.Read)
        xmldoc.Load(fs)
        xmlnode = xmldoc.GetElementsByTagName(nombreNodo)

        For i = 0 To xmlnode.Count - 1

            str = xmlnode(i).ChildNodes.Item(0).InnerText.Trim() & " "

            MsgBox(str)
        Next

    End Sub

    'Devuelve el resultado en un combobox
    Sub leerXMLCMB(ByVal cmb As ComboBox, ByVal ficheroXML As String, ByVal nombreNodo As String, ByVal inicio As String)

        Dim xmldoc As New XmlDataDocument()
        Dim xmlnode As XmlNodeList

        Dim str As String
        Dim fs As New FileStream(ficheroXML, FileMode.Open, FileAccess.Read)
        xmldoc.Load(fs)
        xmlnode = xmldoc.GetElementsByTagName(nombreNodo)
        cmb.Items.Add(inicio)

        For i = 0 To xmlnode.Count - 1
            xmlnode(i).ChildNodes.Item(0).InnerText.Trim()
            str = xmlnode(i).ChildNodes.Item(0).InnerText.Trim()
            cmb.Items.Add(str)
        Next

        cmb.SelectedIndex = 0

    End Sub

    'etiquetasNodos: contiene el nombre de los nodos con valores
    'valores: contiene los datos que se van a introducir
    Sub crearXML(ByVal ficheroXML As String,
                 ByVal nombreNodoRaiz As String,
                 ByVal nombreNodo As String,
                 ByVal etiquetasNodos As String(),
                 ByVal valores As Object(,))

        Dim writer As New XmlTextWriter(ficheroXML, System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        writer.WriteStartElement(nombreNodoRaiz)

        For i = 0 To valores.GetUpperBound(0)

            writer.WriteStartElement(nombreNodo)

            crearNodo(etiquetasNodos, valores, i, writer)

            writer.WriteEndElement()

        Next

        writer.WriteEndElement()
        writer.WriteEndDocument()
        writer.Close()

    End Sub

    Private Sub crearNodo(ByVal etiquetas As String(),
                          ByVal valores As Object(,),
                          ByVal indiceActual As Integer,
                          ByVal writer As XmlTextWriter)

        For i = 0 To valores.GetUpperBound(1)

            writer.WriteStartElement(etiquetas(i))

            writer.WriteString(valores(indiceActual, i))

            writer.WriteEndElement()
        Next

    End Sub

    'Crea un XML a partir de un DataSet
    'etiquetasNodos: contiene el nombre de los nodos con valores
    'valores: contiene los datos que se van a introducir
    Sub crearXMLDataSet(ByVal ficheroXML As String,
                        ByVal nombreNodo As String,
                        ByVal etiquetas As String(),
                        ByVal valores As Object(,))

        Dim dt As New DataTable()
        Dim ds As New DataSet

        Dim tiposSystem As System.Type() = tiposDatosSystem(CopiarFila_A_Array(valores, 0))

        For i = 0 To etiquetas.Length - 1

            dt.Columns.Add(New DataColumn(etiquetas(i), tiposSystem(i)))

        Next

        For i = 0 To valores.GetUpperBound(0)

            llenarFilas(dt, etiquetas, i, valores)

        Next

        ds.Tables.Add(dt)
        ds.Tables(0).TableName = nombreNodo
        ds.WriteXml(ficheroXML)

    End Sub

    'Crea un XML a partir de un DataSet
    'etiquetasNodos: contiene el nombre de los nodos con valores
    'valores: contiene los datos que se van a introducir
    Sub crearXMLDataSet(ByVal ficheroXML As String,
                        ByVal nombreNodo As String,
                        ByVal dt As DataTable)

        Dim ds As New Dataset

        ds.Tables.Add(dt)
        ds.Tables(0).TableName = nombreNodo
        ds.WriteXml(ficheroXML)

    End Sub

    Sub xmlSerializado(ByVal ficheroXML As String,
                        ByVal nombreNodo As String,
                        ByVal etiquetas As String(),
                        ByVal valores As Object(,))

        Dim ds As New DataSet
        Dim dt As New DataTable()

        Dim tiposSystem As System.Type() = tiposDatosSystem(CopiarFila_A_Array(valores, 0))

        For i = 0 To etiquetas.Length - 1

            dt.Columns.Add(New DataColumn(etiquetas(i), tiposSystem(i)))

        Next

        For i = 0 To valores.GetUpperBound(0)

            llenarFilas(dt, etiquetas, i, valores)

        Next

        ds.Tables.Add(dt)
        ds.Tables(0).TableName = nombreNodo

        Dim serialWriter As StreamWriter
        serialWriter = New StreamWriter(ficheroXML)
        Dim xmlWriter As New XmlSerializer(ds.GetType())
        xmlWriter.Serialize(serialWriter, ds)
        serialWriter.Close()
        ds.Clear()

    End Sub


    Private Sub llenarFilas(ByVal dt As DataTable,
                            ByVal etiquetas As String(),
                            ByVal indiceActual As Integer,
                            ByVal valores As Object(,))
        Dim dr As DataRow
        dr = dt.NewRow()

        For i = 0 To valores.GetUpperBound(1)
            dr(etiquetas(i)) = valores(indiceActual, i)

        Next
        dt.Rows.Add(dr)

    End Sub

    'Muestra el contenido de un nodo en MSGBOx, usando dataset
    Sub leerXMLDataSetMensajes(ByVal ficheroXML As String,
                               ByVal campo As Integer)

        Dim xmlFile As XmlReader
        xmlFile = XmlReader.Create(ficheroXML, New XmlReaderSettings())
        Dim ds As New DataSet
        ds.ReadXml(xmlFile)
        Dim i As Integer
        For i = 0 To ds.Tables(0).Rows.Count - 1
            MsgBox(ds.Tables(0).Rows(i).Item(campo))
        Next

    End Sub

    'Devuelve el contenido de un nodo, usando dataset
    Function leerXMLDataSet(ByVal ficheroXML As String,
                            ByVal campo As Integer) As String

        Dim xmlFile As XmlReader
        xmlFile = XmlReader.Create(ficheroXML, New XmlReaderSettings())
        Dim ds As New DataSet
        ds.ReadXml(xmlFile)

        Dim mensaje As String = ""
        For i = 0 To ds.Tables(0).Rows.Count - 1
            mensaje &= ds.Tables(0).Rows(i).Item(campo) & " "
        Next

        Return mensaje

    End Function

    'Muestra en un DGV el contenido de un fichero XML
    Sub leerXMLEnDGV(ByVal ficheroXML As String,
                     ByVal dgv As DataGridView)

        Dim xmlFile As XmlReader
        xmlFile = XmlReader.Create(ficheroXML, New XmlReaderSettings())
        Dim ds As New DataSet
        ds.ReadXml(xmlFile)
        dgv.DataSource = ds.Tables(0)

    End Sub

    'Muestra en un DGV el contenido de un fichero XML en un treeview
    Sub leerXMLEnTreeView(ByVal ficheroXML As String, ByVal tv As TreeView)

        Dim xmldoc As New XmlDataDocument()
        Dim xmlnode As XmlNode
        Dim fs As New FileStream(ficheroXML, FileMode.Open, FileAccess.Read)

        xmldoc.Load(fs)
        xmlnode = xmldoc.ChildNodes(1)
        tv.Nodes.Clear()
        tv.Nodes.Add(New TreeNode(xmldoc.DocumentElement.Name))

        Dim tNode As TreeNode
        tNode = tv.Nodes(0)
        AddNode(xmlnode, tNode)


    End Sub

    Private Sub AddNode(ByVal inXmlNode As XmlNode, ByVal inTreeNode As TreeNode)

        Dim xNode As XmlNode
        Dim tNode As TreeNode
        Dim nodeList As XmlNodeList
        Dim i As Integer
        If inXmlNode.HasChildNodes Then
            nodeList = inXmlNode.ChildNodes
            For i = 0 To nodeList.Count - 1
                xNode = inXmlNode.ChildNodes(i)
                inTreeNode.Nodes.Add(New TreeNode(xNode.Name))
                tNode = inTreeNode.Nodes(i)
                AddNode(xNode, tNode)
            Next
        Else
            inTreeNode.Text = inXmlNode.InnerText.ToString
        End If
    End Sub

    'BuscarEn = campo donde queremos buscar el dato
    'ValorBuscado = valor que queremos encontrar
    'ValorMostrado = valor que se mostrara
    'Lo devuelve en un MSGBOX
    Sub buscarEnXMLMensaje(ByVal ficheroXML As String,
                           ByVal buscarEn As String,
                           ByVal valorBuscado As String,
                           ByVal valorMostrado As String)

        Dim xmlFile As XmlReader
        xmlFile = XmlReader.Create(ficheroXML, New XmlReaderSettings())
        Dim ds As New DataSet
        Dim dv As DataView
        ds.ReadXml(xmlFile)

        dv = New DataView(ds.Tables(0))
        dv.Sort = buscarEn
        Dim index As Integer = dv.Find(valorBuscado)

        If index = -1 Then
            MsgBox("Elemento no encontrado")
        Else
            MsgBox(dv(index)(valorMostrado).ToString())
        End If

    End Sub

    'BuscarEn = campo donde queremos buscar el dato
    'ValorBuscado = valor que queremos encontrar
    'ValorMostrado = valor que se mostrara
    'Lo devuelve en un MSGBOX
    Function buscarEnXML(ByVal ficheroXML As String,
                         ByVal buscarEn As String,
                         ByVal valorBuscado As String,
                         ByVal valorMostrado As String) As String

        Dim xmlFile As XmlReader
        xmlFile = XmlReader.Create(ficheroXML, New XmlReaderSettings())
        Dim ds As New DataSet
        Dim dv As DataView
        ds.ReadXml(xmlFile)

        dv = New DataView(ds.Tables(0))
        dv.Sort = buscarEn
        Dim index As Integer = dv.Find(valorBuscado)

        If index = -1 Then
            MsgBox("Elemento no encontrado")
            Return Nothing
        Else
            Return dv(index)(valorMostrado).ToString()
        End If

    End Function

    'OrdenadoPor = campo por el el queremos ordenar
    Sub filtrarXML(ByVal ficheroXMLOrigen As String,
                   ByVal ficheroXMLDestino As String,
                   ByVal condicion As String,
                   ByVal ordenadoPor As String)

        Dim xmlFile As XmlReader
        xmlFile = XmlReader.Create(ficheroXMLOrigen, New XmlReaderSettings())
        Dim ds As New DataSet
        Dim dv As DataView
        ds.ReadXml(xmlFile)
        dv = New DataView(ds.Tables(0), condicion, ordenadoPor, DataViewRowState.CurrentRows)
        dv.ToTable().WriteXml(ficheroXMLDestino)

    End Sub

    'Pasa a un XML el resultado de un dataset
    Sub deBDaXML(ByVal ficheroXMLDestino As String, ByVal data As DataSet, ByVal nombreNodo As String)

        data.Tables(0).TableName = nombreNodo
        data.WriteXml(ficheroXMLDestino)

    End Sub

    'Pasa de un XML a una base de datos, si existe el registro lo actualiza y si no lo inserta
    'ATENCION: De momento solo funciona para tablas unicas, es decir, los valores del XML deben pertenecer a la misma tabla
    Sub deXMLaBD(ByVal ficheroXMLOrigen As String,
                 ByVal tabla As String,
                 ByVal conexion As MiConnectDB,
                 ByVal tipos() As String,
                 ByVal campos() As String,
                 ByVal posID As Integer)

        Dim xmlFile As XmlReader
        xmlFile = XmlReader.Create(ficheroXMLOrigen, New XmlReaderSettings())
        Dim ds As New DataSet
        ds.ReadXml(xmlFile)

        Dim datos As String = ""
        For Each row As DataRow In ds.Tables(0).Rows

            Dim registroActual() As Object = row.ItemArray

            'Indico si existe o no el campo (no existe = -1)
            If conexion.DLookUp(campos(posID), tabla, campos(posID) & "=" & registroActual(posID)) = -1 Then

                datos = datosInsert(registroActual, tipos)

                conexion.setData("insert into " & tabla & " values (" & datos & ")")

            Else

                datos = datosUpdate(registroActual, tipos, campos, posID)

                Dim condicion As String = campos(posID) & "=" & registroActual(posID)

                conexion.setData("update " & tabla & " set " & datos & " where " & condicion & "")

            End If

            datos = ""

        Next

    End Sub

    Private Function datosInsert(ByVal registro() As Object, ByVal tipos() As String) As String

        Dim datos As String = ""

        For i = 0 To registro.Length - 1

            If (tipos(i) = "String" Or tipos(i) = "DateTime") And registro(i) <> "null" Then

                datos &= "'" & registro(i) & "'"

            Else
                datos &= registro(i)
            End If

            If i <> (registro.Length - 1) Then
                datos &= ", "
            End If

        Next

        Return datos

    End Function

    Private Function datosUpdate(ByVal registro() As Object, ByVal tipos() As String, ByVal campos() As String, ByVal posID As Integer) As String

        Dim datos As String = ""

        For i = 0 To registro.Length - 1

            'Si la posicion i es distinta de la posicion de Id, añade datos al string
            If i <> posID Then

                If (tipos(i) = "String" Or tipos(i) = "DateTime") And registro(i) <> "null" Then

                    datos &= campos(i) & "='" & registro(i) & "'"

                Else
                    datos &= campos(i) & "=" & registro(i)
                End If

                If i <> (registro.Length - 1) Then
                    datos &= ", "
                End If

            End If
            
        Next

        Return datos

    End Function

    Function tiposDatos(ByVal datos As Object()) As String()

        Dim tipos(datos.Length) As String

        For i = 0 To datos.Length - 1

            tipos(i) = TypeName(datos(i))

        Next

        Return tipos

    End Function

    Function tiposDatosSystem(ByVal datos As Object()) As System.Type()

        Dim tipos() As System.Type = Type.GetTypeArray(datos)

        Return tipos

    End Function

    Function CopiarFila_A_Array(ByRef matriz(,) As Object, ByVal fila As Integer) As Object()

        If fila < 0 Or fila > matriz.GetUpperBound(0) Then
            MsgBox("la fila seleccionada no existe")
            Return Nothing
        Else
            Dim aux(matriz.GetUpperBound(1)) As Object

            For j = 0 To matriz.GetUpperBound(1)
                aux(j) = matriz(fila, j)
            Next

            Return aux

        End If

    End Function

    

End Module

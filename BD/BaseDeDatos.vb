Imports System.Windows.Forms
Imports System.Drawing
Module BaseDeDatos

    Function tablaVacia(ByVal tabla As String) As Boolean

        Dim vacia As Boolean = False

        Dim t As DataSet = conexion.getData("select * from " + tabla, tabla)

        If t.Tables(0).Rows.Count = 0 Then
            vacia = True
        End If

        Return vacia

    End Function

    Function masDeunValor(ByVal tabla As String) As Boolean

        Dim valores As Boolean = False

        Dim t As DataSet = conexion.getData("select * from " + tabla, tabla)

        If t.Tables(0).Rows.Count > 1 Then
            valores = True
        End If

        Return valores

    End Function

    'Averigua el id disponible
    Function ultimoIDInt(ByVal columna As String, ByVal tabla As String) As Integer

        Dim maximo As Integer = conexion.DLookUp(columna, tabla, columna & "=(select max(" & columna & ") from " & tabla & ")")

        If maximo = -1 Then
            Return 1
        Else
            Return maximo + 1
        End If

    End Function

    Function ultimoIDString(ByVal columna As String, ByVal tabla As String) As Integer

        Dim maximo As Integer = conexion.DLookUp(columna, tabla, columna & "=(select max(TO_NUMBER(" & columna & ")) from " & tabla & ")")

        If maximo = -1 Then
            Return 1
        Else
            Return maximo + 1
        End If

    End Function

    Function sacarIntFilaBD(ByVal data As DataSet, ByVal columna As String, ByVal nombreTabla As String) As Integer

        Dim fila As DataRow

        fila = data.Tables(nombreTabla).Rows(0)

        Dim num As Integer = Convert.ToInt32(fila(columna))

        Return num

    End Function

    Function sacarStringFilaBD(ByVal data As DataSet, ByVal columna As String, ByVal nombreTabla As String) As String

        Dim fila As DataRow

        fila = data.Tables(nombreTabla).Rows(0)

        Dim cadena As String = fila(columna).ToString

        Return cadena

    End Function

    Function sacarDoubleFilaBD(ByVal data As DataSet, ByVal columna As String, ByVal nombreTabla As String) As Double

        Dim fila As DataRow

        fila = data.Tables(nombreTabla).Rows(0)

        Dim num As Double = CDbl(Val(fila(columna)))

        Return num

    End Function

    'debe ser que devuelva un tipo de valor comun
    Sub rellenaComboBoxBD(ByVal cmb As ComboBox, ByVal dataset As DataSet, ByVal columna As String)

        cmb.Items.Clear()

        cmb.Items.Add("--Elije " & columna & "--")

        For Each row As DataRow In dataset.Tables(0).Rows

            cmb.Items.Add(row(columna))

        Next

        cmb.SelectedIndex = 0

    End Sub

    Sub rellenaListBoxBD(ByVal ltb As ListBox, ByVal dataset As DataSet, ByVal columna As String)

        ltb.Items.Clear()

        ltb.Items.Add("--Elije " & columna & "--")

        For Each row As DataRow In dataset.Tables(0).Rows

            ltb.Items.Add(row(columna))

        Next

    End Sub

    'el textbox debe tener la propiedad multiline activada
    Sub rellenaTextAreaBD(ByVal txt As TextBox, ByVal dataset As DataSet, ByVal columna As String)

        txt.Multiline = True

        txt.Text = ""

        For Each row As DataRow In DataSet.Tables(0).Rows

            txt.Text = txt.Text & row(columna) & vbCrLf

        Next

    End Sub

    'el dataset debes llamarlo con conexion.getData(consulta)
    'LLamar a la tabla de dataset tdata
    Sub rellenaDGVDataTable(ByVal data As DataSet, ByVal dgv As DataGridView)

        Dim tabla As DataTable = data.Tables("tdata")

        dgv.DataSource = tabla

    End Sub

    'Se debe realizar una consulta con dataset antes (getData)
    Sub rellenaCmbBD2Cloumnas(ByVal combo As ComboBox, ByVal data As DataSet, ByVal inicio As String)

        Dim ttablas As DataTable = data.Tables("tabla")
        If Not inicio = "" Then
            Dim newrow2 As DataRow = ttablas.NewRow()
            newrow2(0) = -1
            newrow2(1) = inicio
            ttablas.Rows.InsertAt(newrow2, 0)
        End If
        combo.DataSource = ttablas
        combo.DisplayMember = ttablas.Columns(1).ToString
        combo.ValueMember = ttablas.Columns(0).ToString
    End Sub

    Sub asignarItemCmb2Columnas(ByVal cmb As ComboBox, ByVal codigo As Integer)

        Dim encontrado As Boolean = False


        For i = 1 To cmb.Items.Count - 1

            cmb.SelectedIndex = i

            Dim codigoActual As Integer = cmb.SelectedValue

            If codigo = codigoActual Then
                encontrado = True
                Exit For
            End If

        Next

        If Not encontrado Then
            cmb.SelectedIndex = 0
        End If


    End Sub

    Sub rellenaLB_BD2Columnas(ByVal lst As ListBox, ByVal data As DataSet, ByVal inicio As String)

        Dim ttablas As DataTable = data.Tables("tabla")
        If Not inicio = "" Then
            Dim newrow2 As DataRow = ttablas.NewRow()
            newrow2(0) = -1
            newrow2(1) = inicio
            ttablas.Rows.InsertAt(newrow2, 0)
        End If
        lst.DataSource = ttablas
        lst.DisplayMember = ttablas.Columns(1).ToString
        lst.ValueMember = ttablas.Columns(0).ToString
    End Sub

    Sub asignarItemLB2Columnas(ByVal lb As ListBox, ByVal codigo As Integer)

        Dim encontrado As Boolean = False

        For i = 1 To lb.Items.Count - 1

            lb.SelectedIndex = i

            Dim codigoActual As Integer = lb.SelectedValue

            If codigo = codigoActual Then
                encontrado = True
                Exit For
            End If

        Next

        If Not encontrado Then
            lb.SelectedIndex = 0
        End If


    End Sub


    Function existeValor(ByVal valor As String, ByVal columna As String, ByVal tabla As String) As Boolean

        Dim existe As Boolean = False

        Dim t As DataSet = conexion.getData("select * from " & tabla & " where " & columna & "='" & valor & "'", tabla)

        If t.Tables(0).Rows.Count > 0 Then
            existe = True
        End If

        Return existe

    End Function

    Function existeValorDS(ByVal data As DataSet) As Boolean

        Dim existe As Boolean = False

        If data.Tables(0).Rows.Count > 0 Then
            existe = True
        End If

        Return existe

    End Function

    'En caso de error, devuelve -1
    Function sumaDe(ByVal columna As String, ByVal tabla As String, ByVal columnaID As String, ByVal idRegistro As String) As Integer

        Dim ds As DataSet = conexion.getData("select sum(" & columna & ") as suma from " & tabla & " where eliminado=0 and " & columnaID & "=" & idRegistro & "", "tdata")

        Dim fila As DataRow = ds.Tables("tdata").Rows(0)


        If IsDBNull(fila("suma")) Then

            Return -1

        Else

            Dim suma As Integer = fila("suma")

            Return suma
        End If

    End Function

    'En caso de error, devuelve -1
    Function sumaDe(ByVal columna As String, ByVal tabla As String) As Integer

        Dim ds As DataSet = conexion.getData("select sum(" & columna & ") as suma from " & tabla & " where eliminado=0", "tdata")

        Dim fila As DataRow = ds.Tables("tdata").Rows(0)


        If IsDBNull(fila("suma")) Then

            Return -1

        Else

            Dim suma As Integer = fila("suma")

            Return suma
        End If

    End Function

    Function numeroDe(ByVal columna As String, ByVal tabla As String, ByVal columnaID As String, ByVal idRegistro As String) As Integer

        Dim ds As DataSet = conexion.getData("select count(" & columna & ") as cuenta from " & tabla & " where eliminado=0 and " & columnaID & "=" & idRegistro & "", "tdata")

        Dim fila As DataRow = ds.Tables("tdata").Rows(0)

        If IsDBNull(fila("cuenta")) Then

            Return -1

        Else

            Dim cuenta As Integer = fila("cuenta")

            Return cuenta
        End If

    End Function

    Function numeroDe(ByVal columna As String, ByVal tabla As String) As Integer

        Dim ds As DataSet = conexion.getData("select count(" & columna & ") as cuenta from " & tabla & " where eliminado=0", "tdata")

        Dim fila As DataRow = ds.Tables("tdata").Rows(0)

        If IsDBNull(fila("cuenta")) Then

            Return -1

        Else

            Dim cuenta As Integer = fila("cuenta")

            Return cuenta
        End If

    End Function

    Function numeroDeCP(ByVal idPob As Integer) As Integer

        Dim ds As DataSet = conexion.getData("select count(*) as cuenta from codigospostales where idcodigopostal in (select refcodigopostal from codigospostalespoblaciones where refpoblacion=" & idPob & ")", "tdata")

        Dim fila As DataRow = ds.Tables("tdata").Rows(0)

        If IsDBNull(fila("cuenta")) Then

            Return -1

        Else

            Dim cuenta As Integer = fila("cuenta")

            Return cuenta
        End If

    End Function

    Sub rellenaComunidades(ByVal cmbCom As ComboBox)

        Dim ds As DataSet = conexion.getData("select idcomunidad, comunidad from comunidades", "tabla")

        rellenaCmbBD2Cloumnas(cmbCom, ds, "")

    End Sub

    Sub rellenaProvincias(ByVal cmbCom As ComboBox, ByVal cmbPro As ComboBox)

        Dim idcom As Integer = cmbCom.SelectedValue

        Dim ds As DataSet = conexion.getData("select idprovincia, provincia from provincias where refcomunidad=" & idcom & "", "tabla")

        rellenaCmbBD2Cloumnas(cmbPro, ds, "--Elije provinvcia--")

    End Sub

    Sub rellenaPoblaciones(ByVal cmbPro, ByVal cmbPob)

        Dim idprovincia As Integer = cmbPro.SelectedValue

        Dim ds As DataSet = conexion.getData("select distinct idpoblacion, poblacion from poblaciones p, CODIGOSPOSTALESPOBLACIONES cpp where p.idpoblacion=cpp.refpoblacion and refprovincia= " & idprovincia & "", "tabla")

        rellenaCmbBD2Cloumnas(cmbPob, ds, "--Elige poblacion--")

    End Sub

    Sub rellenaCodigoPostales(ByVal cmbPob As ComboBox, ByVal cmbCP As ComboBox)

        Dim poblacion As Integer = cmbPob.SelectedValue

        Dim t As DataSet = conexion.getData("select idcodigopostal,codigopostal from codigospostales where idcodigopostal =(select refcodigopostal from codigospostalespoblaciones where refpoblacion=" & poblacion & ")", "tabla")

        rellenaCmbBD2Cloumnas(cmbCP, t, "Elige CP")

    End Sub

    'Devuelve el cp, pob, pro y com a partir de un idcodigopostalpob
    Function devolverDatosPoblacion(ByVal idcpp As Integer) As String()


        Dim valores(3) As String

        Dim ds As DataSet = conexion.getData("select cp.codigopostal, p.poblacion,pro.provincia, com.comunidad " & _
                                            "from codigospostalespoblaciones cpp, poblaciones p, provincias pro, comunidades com, codigospostales cp " & _
                                            "where cpp.refpoblacion= p.idpoblacion " & _
                                            "and cpp.refcodigopostal= cp.idcodigopostal " & _
                                            "and cpp.refprovincia= pro.idprovincia " & _
                                            "and com.idcomunidad= pro.refcomunidad " & _
                                            "and cpp.idcodigopostalpob = " & idcpp, "tdata")

        Dim fila As DataRow = ds.Tables(0).Rows(0)

        valores(0) = fila("codigopostal")
        valores(1) = fila("poblacion")
        valores(2) = fila("provincia")
        valores(3) = fila("comunidad")

        Return valores

    End Function


    'cmb(0)=ComboBoxCP
    'cmb(1)=ComboBoxPob
    'cmb(2)=ComboBoxPro
    'cmb(3)=ComboBoxCom

    Sub rellenarCodigoPostalInverso(ByVal idCPP As Integer, ByVal cmb() As ComboBox)


        Dim Sql As DataSet

        'sacar el idcodigopostalpob, con el idcp, idpob y idpob
        'Dim idcpp As String = conexion.DLookUp("idcodigopostalpob", "codigospostalespoblaciones", " refcodigopostal=" & cp & " and refpoblacion=" & cboPoblacion.SelectedValue & " and refprovincia=" & cboProvincia.SelectedValue & "")

        'Saco todos los datos con el idcodigopostalpob
        Dim codigoCP As Integer = conexion.DLookUp("refcodigopostal",
                                                   "codigospostalespoblaciones",
                                                   "idcodigopostalpob = " & idCPP)

        Dim codigoPob As Integer = conexion.DLookUp("idpoblacion",
                                                   "poblaciones",
                                                   "idpoblacion = (select refpoblacion from codigospostalespoblaciones where IDCODIGOPOSTALPOB=" & idCPP & ")")

        Dim codigoPro As Integer = conexion.DLookUp("idprovincia",
                                                    "provincias",
                                                    "idprovincia = (select refprovincia from codigospostalespoblaciones where refpoblacion=" & codigoPob & " and IDCODIGOPOSTALPOB=" & idCPP & ")")


        Dim codigoCom As Integer = conexion.DLookUp("idcomunidad",
                                                    "comunidades",
                                                    "idcomunidad = (select refcomunidad from provincias where  idprovincia=" & codigoPro & ")")



        'Comunidad
        Sql = conexion.getData("select idcomunidad, comunidad from comunidades", "tabla")

        rellenaCmbBD2Cloumnas(cmb(3),
                            Sql,
                            "")


        'Provincia
        Sql = conexion.getData("select idprovincia, provincia from provincias where refcomunidad=" & codigoCom & "", "tabla")

        rellenaCmbBD2Cloumnas(cmb(2),
                                Sql,
                                "-Provincia-")


        'Poblacion
        Sql = conexion.getData("select distinct idpoblacion, poblacion from poblaciones p, CODIGOSPOSTALESPOBLACIONES cpp where p.idpoblacion=cpp.refpoblacion and refprovincia= " & codigoPro & "", "tabla")

        rellenaCmbBD2Cloumnas(cmb(1),
                                Sql,
                                "-Poblacion-")

        'codigopostal

        If numeroDeCP(codigoPob) = 1 Then

            Sql = conexion.getData("select idcodigopostal,codigopostal from codigospostales where idcodigopostal =(select refcodigopostal from codigospostalespoblaciones where refpoblacion=" & codigoPob & ")", "tabla")

            rellenaCmbBD2Cloumnas(cmb(0),
                                Sql,
                                "")

        Else

            Sql = conexion.getData("select idcodigopostal,codigopostal from codigospostales where idcodigopostal in (select refcodigopostal from codigospostalespoblaciones where refpoblacion=" & codigoPob & ")", "tabla")


            rellenaCmbBD2Cloumnas(cmb(0),
                                Sql,
                                "-Codigo Postal-")
        End If


        'Asignamos los valores
        asignarItemCmb2Columnas(cmb(3), codigoCom)

        asignarItemCmb2Columnas(cmb(2), codigoPro)

        asignarItemCmb2Columnas(cmb(1), codigoPob)

        asignarItemCmb2Columnas(cmb(0), codigoCP)

    End Sub

End Module

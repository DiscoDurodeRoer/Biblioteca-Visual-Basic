Imports Oracle.DataAccess.Client

Public Class MiConnectDB

    Private host As String
    Private user As String
    Private passWord As String
    Private serviceName As String

    Private driver As String

    '=======================================================================================================
    ReadOnly Property getData(ByVal query As String, ByVal table As String) As DataSet
        Get
            Dim objConexion As OracleConnection
            Dim objComando As OracleDataAdapter
            Dim requestquery As New DataSet

            objConexion = New OracleConnection(driver)
            objConexion.Open()
            objComando = New OracleDataAdapter(query, objConexion)
            objComando.Fill(requestquery, table)
            objConexion.Close()
            Return requestquery
        End Get
    End Property
    Sub setData(ByVal statement As String)

        Dim objconexion As OracleConnection
        Dim objcomando As OracleCommand

        objconexion = New OracleConnection(driver)
        objconexion.Open()
        objcomando = New OracleCommand(statement, objconexion)

        objcomando.ExecuteNonQuery()

        objcomando.Connection.Close()

    End Sub
    ReadOnly Property DLookUp(ByVal column As String, ByVal table As String, ByVal condition As String) As Object
        Get

            Dim objConexion As OracleConnection
            Dim objComando As OracleDataAdapter
            Dim requestquery As New DataSet
            Dim result As Object

            objConexion = New OracleConnection(driver)
            objConexion.Open()
            If condition = "" Then
                objComando = New OracleDataAdapter("Select " & column & " from " & table, objConexion)
            Else
                objComando = New OracleDataAdapter("Select " & column & " from " & table & " where " & condition, objConexion)
            End If

            objComando.Fill(requestquery)
            Try
                result = requestquery.Tables(0).Rows(0)(requestquery.Tables(0).Columns.IndexOf(column))
            Catch ex As Exception
                result = -1

            End Try
            objConexion.Close()
            Return result
        End Get
    End Property

    Sub New(ByVal host As String, ByVal user As String)

        Me.host = host
        Me.user = user
        Me.passWord = "oradam02"
        Me.serviceName = "ORADAM2"


        Me.driver = "Data Source=(DESCRIPTION =" _
    + "(ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST =" & host & ")(PORT = 1521)))" _
    + "(CONNECT_DATA = (SERVICE_NAME = " & serviceName & "))); " _
    + "User Id=" & user & "; Password=" & passWord & ";"
    End Sub

    Sub New(ByVal host As String, ByVal user As String, ByVal password As String, ByVal serviceName As String)

        Me.host = host
        Me.user = user
        Me.passWord = password
        Me.serviceName = serviceName


        Me.driver = "Data Source=(DESCRIPTION =" _
    + "(ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST =" & host & ")(PORT = 1521)))" _
    + "(CONNECT_DATA = (SERVICE_NAME = " & serviceName & "))); " _
    + "User Id=" & user & "; Password=" & password & ";"


    End Sub

End Class


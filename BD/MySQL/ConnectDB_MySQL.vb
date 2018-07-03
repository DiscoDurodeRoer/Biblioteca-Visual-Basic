Imports MySql.Data.MySqlClient

Public Class ConnectDB

    Private host As String
    Private user As String
    Private password As String
    Private database As String
    Private conexion As MySqlConnection

    Private cadena_conexion As String

    Const driver As String = "Data Source=localhost;Database=galaxy_system;Uid=root ;Password=123456"

    
    '=======================================================================================================
    ReadOnly Property getData(ByVal query As String, ByVal table As String) As DataSet
        Get
            Dim objConexion As MySqlConnection
            Dim objComando As MySqlDataAdapter
            Dim requestquery As New DataSet

            objConexion = New MySqlConnection(cadena_conexion)
            objConexion.Open()
            objComando = New MySqlDataAdapter(query, objConexion)
            objComando.Fill(requestquery, table)
            objConexion.Close()
            Return requestquery
        End Get
    End Property
    Sub setData(ByVal statement As String)

        Dim objconexion As MySqlConnection
        Dim objcomando As MySqlCommand

        objconexion = New MySqlConnection(cadena_conexion)
        objconexion.Open()
        objcomando = New MySqlCommand(statement, objconexion)

        objcomando.ExecuteNonQuery()

        objcomando.Connection.Close()

    End Sub
    ReadOnly Property DLookUp(ByVal column As String, ByVal table As String, ByVal condition As String) As Object
        Get

            Dim objConexion As MySqlConnection
            Dim objComando As MySqlDataAdapter
            Dim requestquery As New DataSet
            Dim result As Object

            objConexion = New MySqlConnection(cadena_conexion)
            objConexion.Open()
            If condition = "" Then
                objComando = New MySqlDataAdapter("Select " & column & " from " & table, objConexion)
            Else
                objComando = New MySqlDataAdapter("Select " & column & " from " & table & " where " & condition, objConexion)
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

End Class


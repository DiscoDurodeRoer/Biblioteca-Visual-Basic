Imports System.Text.RegularExpressions
Module ExpresionesRegulares

    Function validaNumeroEntero_Exp(ByVal texto As String)
        Return Regex.IsMatch(texto, "^[0-9]+$")
    End Function

    Function validaNumeroRealNegativo_Exp(ByVal texto As String) As Boolean
        Return Regex.IsMatch(texto, "^-[0-9]+([\.,][0-9]{1,2})$")
    End Function

    Function validaNumeroReal_Exp(ByVal texto As String) As Boolean
        Return Regex.IsMatch(texto, "^[-]?[0-9]+([\.,][0-9]{1,2})?$")
    End Function

    Function validaNumeroRealPositivo_Exp(ByVal texto As String) As Boolean
        Return Regex.IsMatch(texto, "^[0-9]+([\.,][0-9]{1,2})?$")
    End Function

    'Formato = dd/mm/aaaa
    Function validaNumeroFecha_Exp(ByVal texto As String) As Boolean
        Return Regex.IsMatch(texto, "^(0?[1-9]|[12][0-9]|3[01])[\/](0?[1-9]|1[012])[/\\/](19|20)\d{2}$")
    End Function


    Public Function validar_Mail_Exp(ByVal sMail As String) As Boolean
        ' retorna true o false   
        Return Regex.IsMatch(sMail, _
                "^([\w-]+\.)*?[\w-]+@[\w-]+\.([\w-]+\.)*?[\w]+$")
    End Function

    Function validaIP_Exp(ByVal IP As String) As Boolean

        Return Regex.IsMatch(IP, "^(([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5]).){3}([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])$")

    End Function

    Function validaURL_Exp(ByVal URL As String) As Boolean

        Return Regex.IsMatch(URL, "^(https?:\/\/)?([\da-z\.-]+)\.([a-z\.]{2,6})([\/\w \?=.-]*)*\/?$")

    End Function


End Module

Imports System.Windows.Forms

Module IntervaloFechas

    'Varios
    Sub establecerIntervaloDTPS(ByVal dtps() As DateTimePicker, ByVal mesOrigenIntervalo As Integer, ByVal diaMesInicio As Integer, ByVal diaMesFin As Integer)

        If Date.Now.Month = mesOrigenIntervalo Then

            For i = 0 To dtps.Length

                dtps(i).MinDate = "#" & diaMesInicio & Date.Now.Year - 1 & "#"
                dtps(i).MaxDate = "#" & diaMesFin & Date.Now.Year & "#"

            Next

        Else

            For i = 0 To dtps.Length

                dtps(i).MinDate = "#" & diaMesInicio & Date.Now.Year & "#"
                dtps(i).MaxDate = "#" & diaMesFin & Date.Now.Year + 1 & "#"

            Next

        End If

    End Sub

    'Solo uno
    Sub establecerIntervaloDTP(ByVal dtp As DateTimePicker, ByVal mesOrigenIntervalo As Integer, ByVal diaMesInicio As Integer, ByVal diaMesFin As Integer)

        If Date.Now.Month = mesOrigenIntervalo Then

            dtp.MinDate = "#" & diaMesInicio & Date.Now.Year - 1 & "#"
            dtp.MaxDate = "#" & diaMesFin & Date.Now.Year & "#"

        Else

            dtp.MinDate = "#" & diaMesInicio & Date.Now.Year & "#"
            dtp.MaxDate = "#" & diaMesFin & Date.Now.Year + 1 & "#"

        End If

    End Sub

    Function fechaEntreIntervalo(ByVal fechaEvaluada As Date, ByVal fechaMinima As Date, ByVal fechaMaxima As Date) As Boolean

        If fechaEvaluada.ToString("dd/MM/yyyy") > fechaMinima.ToString("dd/MM/yyyy") And fechaEvaluada.ToString("dd/MM/yyyy") < fechaMaxima.ToString("dd/MM/yyyy") Then
            Return True
        End If

        Return False

    End Function


End Module

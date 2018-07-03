Imports System.Windows.Forms
Imports System.Drawing
Module LimpiarComponentes

    Sub limpiarCampos(ByRef campos() As TextBox)

        For i = 0 To campos.Length - 1
            campos(i).Text = ""
        Next

    End Sub

    Sub limpiarTodoDGVS(ByVal dgvs() As DataGridView)

        For i = 0 To dgvs.Length - 1

            limpiarDataGridView(dgvs(i))

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

End Module

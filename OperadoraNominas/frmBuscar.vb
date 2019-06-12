Public Class frmBuscar
    Public gNombre As String

    Private Sub frmBuscart_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtbuscar.TabIndex = 1

    End Sub

    Private Sub cmdCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub cmdAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAceptar.Click
        Try
            gNombre = txtbuscar.Text
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()

        Catch ex As Exception

        End Try
    End Sub


    
   
    Private Sub txtbuscar_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtbuscar.KeyDown
        Select Case e.KeyData
            Case Keys.Enter
                cmdAceptar_Click(sender, e)
        End Select
    End Sub
End Class
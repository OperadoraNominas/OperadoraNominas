Imports ClosedXML.Excel
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Net.Mime.MediaTypeNames
Imports Microsoft.Office.Interop

Public Class frmNominaFinal
    Dim sheetIndex As Integer = -1
    Dim SQL As String
    Dim contacolumna As Integer
    Dim ini, fin As String
    Dim rutita As String
    Dim fechadepago As String

    Public dsReporte As New DataSet

    Private Sub frmNominaFinal_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        'MostrarEmpresasC()
        Dim moment As Date = Date.Now()
        cargarperiodos()

    End Sub


    Private Sub cargarperiodos()
        'Verificar si se tienen permisos
        Dim sql As String
        Try
            sql = "Select (CONVERT(nvarchar(12),dFechaInicio,103) + ' - ' + CONVERT(nvarchar(12),dFechaFin,103)) as dFechaInicio,iIdPeriodo  from periodos where iEstatus=1 order by iEjercicio,iNumeroPeriodo"
            nCargaCBO(cboperiodo, sql, "dFechainicio", "iIdPeriodo")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub tsbNuevo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles tsbNuevo.Click
        tsbNuevo.Enabled = False
        tsbImportar.Enabled = True
        tsbImportar_Click(sender, e)
    End Sub

    Private Sub tsbImportar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles tsbImportar.Click
        Dim dialogo As New OpenFileDialog
        lblRuta.Text = ""
        With dialogo
            .Title = "Búsqueda de archivos de saldos."
            .Filter = "Hoja de cálculo de excel (xlsx)|*.xlsx;"
            .CheckFileExists = True
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                lblRuta.Text = .FileName
            End If
        End With
        tsbProcesar.Enabled = lblRuta.Text.Length > 0
        If tsbProcesar.Enabled Then
            tsbProcesar_Click(sender, e)
        End If
    End Sub

    Private Sub tsbProcesar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbProcesar.Click
        lsvLista.Items.Clear()
        lsvLista.Columns.Clear()
        lsvLista.Clear()

        pnlCatalogo.Enabled = False
        tsbGuardar.Enabled = False
        tsbCancelar.Enabled = False
        tsbEnviar.Enabled = False
        lsvLista.Visible = False
        tsbImportar.Enabled = False
        Me.cmdCerrar.Enabled = False
        Me.Cursor = Cursors.WaitCursor
        Me.Enabled = False
        ' Application.DoEvents()

        Try
            If File.Exists(lblRuta.Text) Then
                Dim Archivo As String = lblRuta.Text
                Dim Hoja As String


                Dim book As New ClosedXML.Excel.XLWorkbook(Archivo)
                If book.Worksheets.Count >= 1 Then
                    sheetIndex = 1
                    If book.Worksheets.Count >= 1 Then
                        Dim Forma As New frmHojasNomina
                        Dim Hojas As String = ""
                        For i As Integer = 0 To book.Worksheets.Count - 1
                            Hojas &= book.Worksheets(i).Name & IIf(i < (book.Worksheets.Count - 1), "|", "")
                        Next
                        Forma.Hojas = Hojas
                        If Forma.ShowDialog = Windows.Forms.DialogResult.OK Then
                            sheetIndex = Forma.selectedIndex + 1
                        Else
                            Exit Sub
                        End If
                    End If
                    Hoja = book.Worksheet(sheetIndex).Name
                    Dim sheet As IXLWorksheet = book.Worksheet(sheetIndex)

                    Dim colIni As Integer = sheet.FirstColumnUsed().ColumnNumber() - 1
                    Dim colFin As Integer = sheet.LastColumnUsed().ColumnNumber()
                    Dim Columna As String
                    Dim numerocolumna As Integer = 0


                    lsvLista.Columns.Add("#")
                    For c As Integer = colIni To colFin

                        lsvLista.Columns.Add("CODIGO")
                        lsvLista.Columns.Add("NOMBRE")
                        lsvLista.Columns.Add("ISTATUS")
                        lsvLista.Columns.Add("RFC")
                        lsvLista.Columns.Add("CURP")
                        lsvLista.Columns.Add("NSS")
                        lsvLista.Columns.Add("FECHA_NAC")
                        lsvLista.Columns.Add("EDAD")
                        lsvLista.Columns.Add("PUESTO")
                        lsvLista.Columns.Add("BUQUE")
                        lsvLista.Columns.Add("TIPO_INFONAVIT")
                        lsvLista.Columns.Add("VALOR_INFONAVIT")
                        lsvLista.Columns.Add("SALARIO_DIARIO")
                        lsvLista.Columns.Add("SDI")
                        lsvLista.Columns.Add("DIAS_TRABAJADOS")
                        lsvLista.Columns.Add("TIPO_INCAPACIDAD")
                        lsvLista.Columns.Add("NUMERO_DIAS")
                        lsvLista.Columns.Add("SUELDO_BASE")
                        lsvLista.Columns.Add("TIEMPO_EXTRA_FIJO_GRAVADO")
                        lsvLista.Columns.Add("TIEMPO_EXTRA_FIJO_EXENTO")
                        lsvLista.Columns.Add("TIEMPO_EXTRA_OCASIONAL")
                        lsvLista.Columns.Add("DESC_SEM_OBLIGATORIO")
                        lsvLista.Columns.Add("VACACIONES_PROPORCIONALES")
                        lsvLista.Columns.Add("AGUINALDO_GRAVADO")
                        lsvLista.Columns.Add("AGUINALDO_EXENTO")
                        lsvLista.Columns.Add("TOTAL_AGUINALDO")
                        lsvLista.Columns.Add("P_VAC_GRAVADO")
                        lsvLista.Columns.Add("P_VAC_EXENTO")
                        lsvLista.Columns.Add("TOTAL_P_VAC")
                        lsvLista.Columns.Add("TOTAL_PERCEPCIONES")
                        lsvLista.Columns.Add("TOTAL_PERCEPCIONES_P_ISR")
                        lsvLista.Columns.Add("INCAPACIDAD")
                        lsvLista.Columns.Add("ISR")
                        lsvLista.Columns.Add("IMSS")
                        lsvLista.Columns.Add("INFONAVIT")
                        lsvLista.Columns.Add("INFONAVIT_ANT")
                        lsvLista.Columns.Add("INFONAVIT_BIM_ANT")
                        lsvLista.Columns.Add("PENSION_ALIMENTICIA")
                        lsvLista.Columns.Add("SUBSIDIO")
                        lsvLista.Columns.Add("PRESTAMO")
                        lsvLista.Columns.Add("FONACOT")
                        lsvLista.Columns.Add("NETO_PAGAR")
                        lsvLista.Columns.Add("IMSS_CS")
                        lsvLista.Columns.Add("RCV_CS")
                        lsvLista.Columns.Add("INFONAVIT_CS")
                        lsvLista.Columns.Add("ISNOM_CS")
                        lsvLista.Columns.Add("TOTAL_CS")
                        lsvLista.Columns.Add("COSTO_SOCIAL")
                        lsvLista.Columns.Add("Prestamo_Personal_Asimilado")
                        lsvLista.Columns.Add("Adeudo_Infonavit_Asimilado")
                        lsvLista.Columns.Add("Difencia_infonavit_Asimilado")
                        lsvLista.Columns.Add("ASIMILADOS")




                        numerocolumna = numerocolumna + 1

                    Next



                    lsvLista.Columns(1).Width = 400 'Empleado
                    lsvLista.Columns(2).Width = 100  'ISR
                    lsvLista.Columns(3).Width = 50 '#Control
                    lsvLista.Columns(4).Width = 100 'ap
                    lsvLista.Columns(5).Width = 100 'am
                    lsvLista.Columns(6).Width = 100 'nombre
                    lsvLista.Columns(7).Width = 100 'isr
                    lsvLista.Columns(8).Width = 200 'imss
                    lsvLista.Columns(9).Width = 50 'dias
                    lsvLista.Columns(10).Width = 100 'banco
                    lsvLista.Columns(11).Width = 150 'clabe
                    lsvLista.Columns(12).Width = 150 'cuenta
                    lsvLista.Columns(13).Width = 150 'curp
                    lsvLista.Columns(14).Width = 350 'rfc
                    lsvLista.Columns(15).Width = 350


                    Dim Filas As Long = sheet.RowsUsed().Count()
                    For f As Integer = 1 To Filas
                        Dim item As ListViewItem = lsvLista.Items.Add(f.ToString())
                        For c As Integer = colIni To colFin
                            Try

                                Dim Valor As String = ""
                                If (sheet.Cell(f, c).ValueCached Is Nothing) Then
                                    Valor = sheet.Cell(f, c).Value.ToString()
                                Else
                                    Valor = sheet.Cell(f, c).ValueCached.ToString()
                                End If
                                Valor = Valor.Trim()
                                item.SubItems.Add(Valor)


                                If f = 6 And c >= 12 Then


                                    item.SubItems(item.SubItems.Count - 1).Text = Valor
                                End If



                            Catch ex As Exception
                                MessageBox.Show("DENTRO DEL FOR " & f & "-" & c & ex.Message.ToString)
                            End Try

                        Next
                    Next

                    book.Dispose()
                    book = Nothing
                    GC.Collect()

                    pnlCatalogo.Enabled = True
                    If lsvLista.Items.Count = 0 Then
                        MessageBox.Show("El catálogo no puso ser importado o no contiene registros." & vbCrLf & "¿Por favor verifique?", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Else
                        MessageBox.Show("Se han encontrado " & FormatNumber(lsvLista.Items.Count, 0) & " registros en el archivo.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                        tsbGuardar.Enabled = True
                        tsbCancelar.Enabled = True
                        lblRuta.Text = FormatNumber(lsvLista.Items.Count, 0) & " registros en el archivo."
                        Me.Enabled = True
                        Me.cmdCerrar.Enabled = True
                        Me.Cursor = Cursors.Default
                        tsbImportar.Enabled = True
                        tsbEnviar.Enabled = True
                        lsvLista.Visible = True
                    End If




                ElseIf book.Worksheets.Count = 0 Then
                    MessageBox.Show("El archivo no contiene hojas.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Else
                MessageBox.Show("El archivo ya no se encuentra en la ruta indicada.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            tsbCancelar_Click(sender, e)
            'tsbImportar.Enabled = False
            MessageBox.Show(ex.Message.ToString)


        End Try
    End Sub
    Private Sub frmImportarEmpladosAlta_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub chkAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAll.CheckedChanged
        For Each item As ListViewItem In lsvLista.Items
            item.Checked = chkAll.Checked
        Next
        chkAll.Text = IIf(chkAll.Checked, "Desmarcar todos", "Marcar todos")
    End Sub


    Private Sub tsbCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbCancelar.Click
        pnlCatalogo.Enabled = False
        lsvLista.Items.Clear()
        chkAll.Checked = False
        lblRuta.Text = ""
        tsbImportar.Enabled = False
        tsbEnviar.Enabled = False
        tsbCancelar.Enabled = False
        tsbNuevo.Enabled = True
    End Sub

    Private Sub tsbGuardar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbGuardar.Click
        Dim SQL As String, nombresistema As String = ""
        Try
            Dim resultado As Integer = MessageBox.Show("Solo se agregaran los registros seleccionados y en color verde, ¿Desea continuar?", "Pregunta", MessageBoxButtons.YesNo)
            If resultado = DialogResult.Yes Then

                If lsvLista.CheckedItems.Count > 0 Then


                    dsReporte.Tables.Add("Tabla")
                    dsReporte.Tables("Tabla").Columns.Add("Id_empleado")
                    dsReporte.Tables("Tabla").Columns.Add("CodigoEmpleado")
                    dsReporte.Tables("Tabla").Columns.Add("dias")
                    dsReporte.Tables("Tabla").Columns.Add("Salario")
                    dsReporte.Tables("Tabla").Columns.Add("Bono")
                    dsReporte.Tables("Tabla").Columns.Add("Refrendo")
                    dsReporte.Tables("Tabla").Columns.Add("SalarioTMM")
                    dsReporte.Tables("Tabla").Columns.Add("CodigoPuesto")
                    dsReporte.Tables("Tabla").Columns.Add("CodigoBuque")

                    dsReporte.Tables("Tabla").Columns.Add("Fechainicio")
                    dsReporte.Tables("Tabla").Columns.Add("Fechafin")
                    Dim mensaje As String

                    pnlProgreso.Visible = True
                    pnlCatalogo.Enabled = False
                    '     Application.DoEvents()
                    '
                    Dim IdProducto As Long
                    Dim i As Integer = 0
                    Dim conta As Integer = 0

                    pgbProgreso.Minimum = 0
                    pgbProgreso.Value = 0
                    pgbProgreso.Maximum = lsvLista.CheckedItems.Count

                    For Each producto As ListViewItem In lsvLista.CheckedItems
                        SQL = "select * from empleadosC where cCodigoEmpleado = " & Trim(producto.SubItems(1).Text).Substring(2, 4)
                        Dim rwFilas As DataRow() = nConsulta(SQL)

                        If rwFilas Is Nothing = False Then
                            If rwFilas.Length = 1 Then
                                producto.BackColor = Color.Green
                                Dim fila As DataRow = dsReporte.Tables("Tabla").NewRow

                                fila.Item("Id_empleado") = rwFilas(0)("iIdEmpleadoC")
                                fila.Item("CodigoEmpleado") = Trim(producto.SubItems(1).Text).Substring(2, 4)
                                fila.Item("dias") = Trim(producto.SubItems(9).Text)
                                fila.Item("Salario") = Trim(producto.SubItems(17).Text)
                                fila.Item("Bono") = Trim(producto.SubItems(17).Text)
                                fila.Item("Refrendo") = Trim(producto.SubItems(17).Text)
                                fila.Item("SalarioTMM") = Trim(producto.SubItems(17).Text)
                                fila.Item("CodigoPuesto") = Trim(producto.SubItems(4).Text)
                                fila.Item("CodigoBuque") = Trim(producto.SubItems(10).Text)
                                fila.Item("Fechainicio") = (Date.Parse(Trim(producto.SubItems(7).Text))).ToShortDateString
                                fila.Item("Fechafin") = (Date.Parse(Trim(producto.SubItems(8).Text))).ToShortDateString
                                dsReporte.Tables("Tabla").Rows.Add(fila)

                            End If

                        End If
                        pgbProgreso.Value += 1

                    Next

                    'Enviar correo
                    'Enviar_Mail(GenerarCorreoFlujo("Importación Flujo-Conceptos", "Área Facturación", "Se importo un flujo con los conceptos necesarios"), "g.gomez@mbcgroup.mx", "Importación")





                    'tsbCancelar_Click(sender, e)
                    pnlProgreso.Visible = False
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    Me.Close()
                    'MessageBox.Show("Proceso terminado", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                Else

                    MessageBox.Show("Por favor seleccione al menos una registro para importar.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
                pnlCatalogo.Enabled = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub tsbEnviar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbEnviar.Click
        Try
            Dim resultado As Integer = MessageBox.Show("Solo se agregaran los registros seleccionados y en color verde, ¿Desea continuar?", "Pregunta", MessageBoxButtons.YesNo)
            If resultado = DialogResult.Yes Then

                If lsvLista.CheckedItems.Count > 0 Then


                    dsReporte.Tables.Add("Tabla")
                    'dsReporte.Tables("Tabla").Columns.Add("iIdNominaFinal")
                    dsReporte.Tables("Tabla").Columns.Add("CODIGO")
                    dsReporte.Tables("Tabla").Columns.Add("NOMBRE")
                    dsReporte.Tables("Tabla").Columns.Add("ISTATUS")
                    dsReporte.Tables("Tabla").Columns.Add("RFC")
                    dsReporte.Tables("Tabla").Columns.Add("CURP")
                    dsReporte.Tables("Tabla").Columns.Add("NSS")
                    dsReporte.Tables("Tabla").Columns.Add("FECHA_NAC")
                    dsReporte.Tables("Tabla").Columns.Add("EDAD")
                    dsReporte.Tables("Tabla").Columns.Add("PUESTO")
                    dsReporte.Tables("Tabla").Columns.Add("BUQUE")
                    dsReporte.Tables("Tabla").Columns.Add("TIPO_INFONAVIT")
                    dsReporte.Tables("Tabla").Columns.Add("VALOR_INFONAVIT")
                    dsReporte.Tables("Tabla").Columns.Add("SALARIO_DIARIO")
                    dsReporte.Tables("Tabla").Columns.Add("SDI")
                    dsReporte.Tables("Tabla").Columns.Add("DIAS_TRABAJADOS")
                    dsReporte.Tables("Tabla").Columns.Add("TIPO_INCAPACIDAD")
                    dsReporte.Tables("Tabla").Columns.Add("NUMERO_DIAS")
                    dsReporte.Tables("Tabla").Columns.Add("SUELDO_BASE")
                    dsReporte.Tables("Tabla").Columns.Add("TIEMPO_EXTRA_FIJO_GRAVADO")
                    dsReporte.Tables("Tabla").Columns.Add("TIEMPO_EXTRA_FIJO_EXENTO")
                    dsReporte.Tables("Tabla").Columns.Add("TIEMPO_EXTRA_OCASIONAL")
                    dsReporte.Tables("Tabla").Columns.Add("DESC_SEM_OBLIGATORIO")
                    dsReporte.Tables("Tabla").Columns.Add("VACACIONES_PROPORCIONALES")
                    dsReporte.Tables("Tabla").Columns.Add("AGUINALDO_GRAVADO")
                    dsReporte.Tables("Tabla").Columns.Add("AGUINALDO_EXENTO")
                    dsReporte.Tables("Tabla").Columns.Add("TOTAL_AGUINALDO")
                    dsReporte.Tables("Tabla").Columns.Add("P_VAC_GRAVADO")
                    dsReporte.Tables("Tabla").Columns.Add("P_VAC_EXENTO")
                    dsReporte.Tables("Tabla").Columns.Add("TOTAL_P_VAC")
                    dsReporte.Tables("Tabla").Columns.Add("TOTAL_PERCEPCIONES")
                    dsReporte.Tables("Tabla").Columns.Add("TOTAL_PERCEPCIONES_P_ISR")
                    dsReporte.Tables("Tabla").Columns.Add("INCAPACIDAD")
                    dsReporte.Tables("Tabla").Columns.Add("ISR")
                    dsReporte.Tables("Tabla").Columns.Add("IMSS")
                    dsReporte.Tables("Tabla").Columns.Add("INFONAVIT")
                    dsReporte.Tables("Tabla").Columns.Add("INFONAVIT_ANT")
                    dsReporte.Tables("Tabla").Columns.Add("INFONAVIT_BIM_ANT")
                    dsReporte.Tables("Tabla").Columns.Add("PENSION_ALIMENTICIA")
                    dsReporte.Tables("Tabla").Columns.Add("SUBSIDIO")
                    dsReporte.Tables("Tabla").Columns.Add("PRESTAMO")
                    dsReporte.Tables("Tabla").Columns.Add("FONACOT")
                    dsReporte.Tables("Tabla").Columns.Add("NETO_PAGAR")
                    dsReporte.Tables("Tabla").Columns.Add("IMSS_CS")
                    dsReporte.Tables("Tabla").Columns.Add("RCV_CS")
                    dsReporte.Tables("Tabla").Columns.Add("INFONAVIT_CS")
                    dsReporte.Tables("Tabla").Columns.Add("ISNOM_CS")
                    dsReporte.Tables("Tabla").Columns.Add("TOTAL_CS")
                    dsReporte.Tables("Tabla").Columns.Add("COSTO_SOCIAL")
                    dsReporte.Tables("Tabla").Columns.Add("Prestamo_Personal_Asimilado")
                    dsReporte.Tables("Tabla").Columns.Add("Adeudo_Infonavit_Asimilado")
                    dsReporte.Tables("Tabla").Columns.Add("Difencia_infonavit_Asimilado")
                    dsReporte.Tables("Tabla").Columns.Add("ASIMILADOS")
                    dsReporte.Tables("Tabla").Columns.Add("iTipoNomina")
                    dsReporte.Tables("Tabla").Columns.Add("iSerie")
                    dsReporte.Tables("Tabla").Columns.Add("fKiIdPeriodo")

                    Dim mensaje As String

                    pnlProgreso.Visible = True
                    pnlCatalogo.Enabled = False
                    'Application.DoEvents()

                    Dim IdProducto As Long
                    Dim i As Integer = 0
                    Dim conta As Integer = 0

                    pgbProgreso.Minimum = 0
                    pgbProgreso.Value = 0
                    pgbProgreso.Maximum = lsvLista.CheckedItems.Count

                    For Each producto As ListViewItem In lsvLista.CheckedItems
                        SQL = "select * from empleadosC where cCodigoEmpleado = " & Trim(producto.SubItems(1).Text)
                        Dim rwFilas As DataRow() = nConsulta(SQL)

                        If rwFilas Is Nothing = False Then
                            If rwFilas.Length = 1 Then
                                producto.BackColor = Color.Green
              
                                SQL = "EXEC setNominaInsertar 0"
                                SQL &= ",'" & producto.SubItems(1).Text & "'"
                                SQL &= ",'" & producto.SubItems(2).Text & "'"
                                SQL &= ",'" & producto.SubItems(3).Text & "'"
                                SQL &= ",'" & producto.SubItems(4).Text & "'"
                                SQL &= ",'" & producto.SubItems(5).Text & "'"
                                SQL &= ",'" & producto.SubItems(6).Text & "'"
                                SQL &= ",'" & producto.SubItems(7).Text & "'"
                                SQL &= ",'" & producto.SubItems(8).Text & "'"
                                SQL &= ",'" & producto.SubItems(9).Text & "'"
                                SQL &= ",'" & producto.SubItems(10).Text & "'"
                                SQL &= ",'" & producto.SubItems(11).Text & "'"
                                SQL &= ",'" & producto.SubItems(12).Text & "'"
                                SQL &= "," & producto.SubItems(13).Text
                                SQL &= "," & producto.SubItems(14).Text
                                SQL &= ",'" & producto.SubItems(15).Text & "'"
                                SQL &= ",'" & producto.SubItems(16).Text & "'"
                                SQL &= "," & producto.SubItems(17).Text
                                SQL &= "," & producto.SubItems(18).Text
                                SQL &= "," & producto.SubItems(19).Text
                                SQL &= "," & producto.SubItems(20).Text
                                SQL &= "," & producto.SubItems(21).Text
                                SQL &= "," & producto.SubItems(22).Text
                                SQL &= "," & producto.SubItems(23).Text
                                SQL &= "," & producto.SubItems(24).Text
                                SQL &= "," & producto.SubItems(25).Text
                                SQL &= "," & producto.SubItems(26).Text
                                SQL &= "," & producto.SubItems(27).Text
                                SQL &= "," & producto.SubItems(28).Text
                                SQL &= "," & producto.SubItems(29).Text
                                SQL &= "," & producto.SubItems(30).Text
                                SQL &= "," & producto.SubItems(31).Text
                                SQL &= "," & producto.SubItems(32).Text
                                SQL &= "," & producto.SubItems(33).Text
                                SQL &= "," & producto.SubItems(34).Text
                                SQL &= "," & producto.SubItems(35).Text
                                SQL &= "," & producto.SubItems(36).Text
                                SQL &= "," & producto.SubItems(37).Text
                                SQL &= "," & producto.SubItems(38).Text
                                SQL &= "," & producto.SubItems(39).Text
                                SQL &= "," & producto.SubItems(40).Text
                                SQL &= "," & producto.SubItems(41).Text
                                SQL &= "," & producto.SubItems(42).Text
                                SQL &= "," & producto.SubItems(43).Text
                                SQL &= "," & producto.SubItems(44).Text
                                SQL &= "," & producto.SubItems(45).Text
                                SQL &= "," & producto.SubItems(46).Text
                                SQL &= "," & producto.SubItems(47).Text
                                SQL &= "," & producto.SubItems(48).Text
                                SQL &= "," & producto.SubItems(49).Text
                                SQL &= "," & producto.SubItems(50).Text
                                SQL &= "," & producto.SubItems(51).Text
                                SQL &= "," & producto.SubItems(52).Text
                                SQL &= "," & cboTipoNomina.SelectedIndex
                                SQL &= "," & cboserie.SelectedIndex
                                SQL &= "," & cboperiodo.SelectedValue

                            End If
                            If nExecute(SQL) = False Then
                                MessageBox.Show("Ocurrio un error " & producto.SubItems(2).Text, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                pnlProgreso.Visible = False
                                Me.Close()
                                Exit Sub
                            End If

                        End If
                        pgbProgreso.Value += 1

                    Next

                    'Enviar correo
                    'Enviar_Mail(GenerarCorreoFlujo("Importación Flujo-Conceptos", "Área Facturación", "Se importo un flujo con los conceptos necesarios"), "g.gomez@mbcgroup.mx", "Importación")





                    'tsbCancelar_Click(sender, e)
                    pnlProgreso.Visible = False
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    Me.Close()
                    'MessageBox.Show("Proceso terminado", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                Else

                    MessageBox.Show("Por favor seleccione al menos una registro para importar.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Me.Close()
                End If
                pnlCatalogo.Enabled = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
    End Sub

    Public Sub recorrerFilasColumnas(ByRef hoja As IXLWorksheet, ByRef filainicio As Integer, ByRef filafinal As Integer, ByRef colTotal As Integer, ByRef tipo As String, Optional ByVal inicioCol As Integer = 1)

        For f As Integer = filainicio To filafinal


            For c As Integer = IIf(inicioCol = Nothing, 1, inicioCol) To colTotal

                Select Case tipo
                    Case "bold"
                        hoja.Cell(f, c).Style.Font.SetFontColor(XLColor.Black)
                    Case "bold false"
                        hoja.Cell(f, c).Style.Font.SetBold(False)
                    Case "clear"
                        hoja.Cell(f, c).Clear()
                    Case "sin relleno"
                        hoja.Cell(f, c).Style.Fill.BackgroundColor = XLColor.NoColor
                    Case "text black"
                        hoja.Cell(f, c).Style.Font.SetFontColor(XLColor.Black)




                End Select
            Next
        Next

    End Sub
End Class
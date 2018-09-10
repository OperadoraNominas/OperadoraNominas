Imports ClosedXML.Excel
Imports System.IO

Public Class frmnominasmarinos
    Private m_currentControl As Control = Nothing
    Public gIdEmpresa As String
    Public gIdTipoPeriodo As String
    Public gNombrePeriodo As String
    Dim Ruta As String
    Dim nombre As String
    Dim cargado As Boolean = False
    Dim diasperiodo As Integer
    Dim aniocostosocial As Integer
    Dim dgvCombo As DataGridViewComboBoxEditingControl
    Dim campoordenamiento As String
    Dim TipoNomina As Boolean
    Dim IDCalculoInfonavit As Integer

    Private Sub dvgCombo_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            '
            ' se recupera el valor del combo
            ' a modo de ejemplo se escribe en consola el valor seleccionado
            '



            Dim combo As ComboBox = TryCast(sender, ComboBox)

            If dgvCombo IsNot Nothing Then
                Dim sql As String
                'Console.WriteLine(combo.SelectedValue)
                'MessageBox.Show(combo.Text)
                '
                ' se accede a la fila actual, para trabajr con otor de sus campos
                ' en este caso se marca el check si se cambia la seleccion
                '
                Dim row As DataGridViewRow = dtgDatos.CurrentRow

                'Dim cell As DataGridViewCheckBoxCell = TryCast(row.Cells("Seleccionado"), DataGridViewCheckBoxCell)
                'cell.Value = True

                'Poner los datos necesarios para poner el nuevo sueldo diario y el integrado


                sql = "Select salariod,sbc,salariodTopado,sbcTopado from costosocial "
                sql &= " where fkiIdPuesto = " & combo.SelectedValue & " and anio=" & aniocostosocial

                Dim rwDatosSalario As DataRow() = nConsulta(sql)

                If rwDatosSalario Is Nothing = False Then
                    If row.Cells(10).Value >= 55 Then
                        row.Cells(16).Value = rwDatosSalario(0)("salariodTopado")
                        row.Cells(17).Value = rwDatosSalario(0)("sbcTopado")
                    Else
                        row.Cells(16).Value = rwDatosSalario(0)("salariod")
                        row.Cells(17).Value = rwDatosSalario(0)("sbc")
                    End If

                Else
                    MessageBox.Show("No se encontraron datos")
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        


    End Sub

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub frmcontpaqnominas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim sql As String
            cargarperiodos()
            Me.dtgDatos.ContextMenuStrip = Me.cMenu
            cboserie.SelectedIndex = 0
            cboTipoNomina.SelectedIndex = 0
            Sql = "select * from periodos where iIdPeriodo= " & cboperiodo.SelectedValue
            Dim rwPeriodo As DataRow() = nConsulta(Sql)
            If rwPeriodo Is Nothing = False Then

                aniocostosocial = Date.Parse(rwPeriodo(0)("dFechaInicio").ToString).Year

            End If

            campoordenamiento = "Nomina.Buque,cNombreLargo"
            TipoNomina = False


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try



    End Sub

    Private Sub cargarbancosasociados()
        Dim sql As String
        Try
            sql = "select * from bancos inner join ( select distinct(fkiidBanco) from DatosBanco where fkiIdEmpresa=" & gIdEmpresa & ") bancos2 on bancos.iIdBanco=bancos2.fkiidBanco order by cBanco"
            nCargaCBO(cbobancos, sql, "cBanco", "iIdBanco")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub cargarperiodos()
        'Verificar si se tienen permisos
        Dim sql As String
        Try
            sql = "Select (CONVERT(nvarchar(12),dFechaInicio,103) + ' - ' + CONVERT(nvarchar(12),dFechaFin,103)) as dFechaInicio,iIdPeriodo  from periodos order by iEjercicio,iNumeroPeriodo"
            nCargaCBO(cboperiodo, sql, "dFechainicio", "iIdPeriodo")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    
    

    Private Sub cmdverdatos_Click(sender As Object, e As EventArgs) Handles cmdverdatos.Click
        Try
            'If cargado Then



            '    dtgDatos.DataSource = Nothing
            '    llenargrid()
            'Else
            '    cargado = True
            '    llenargrid()
            'End If
            If dtgDatos.RowCount > 0 Then
                Dim resultado As Integer = MessageBox.Show("ya se tienen empleados cargados en la lista, si continua estos se borraran,¿Desea continuar?", "Pregunta", MessageBoxButtons.YesNo)
                If resultado = DialogResult.Yes Then

                    dtgDatos.Columns.Clear()
                    llenargrid()

                End If
            Else
                dtgDatos.Columns.Clear()
                llenargrid()

            End If
            



        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub llenargrid()

        Try
            Dim sql As String
            Dim sql2 As String
            Dim infonavit As Double
            Dim prestamo As Double
            Dim incidencia As Double
            Dim bCalcular As Boolean
            Dim PrimaSA As Double
            Dim cadenabanco As String
            dtgDatos.Columns.Clear()
            dtgDatos.DataSource = Nothing


            dtgDatos.DefaultCellStyle.Font = New Font("Calibri", 8)
            dtgDatos.ColumnHeadersDefaultCellStyle.Font = New Font("Calibri", 9)
            Dim chk As New DataGridViewCheckBoxColumn()
            dtgDatos.Columns.Add(chk)
            chk.HeaderText = ""
            chk.Name = "chk"
            'dtgDatos.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

            'dtgDatos.Columns("chk").SortMode = DataGridViewColumnSortMode.NotSortable

            'dtgDatos.Columns.Add("idempleado", "idempleado")
            'dtgDatos.Columns(0).Width = 30
            'dtgDatos.Columns(0).ReadOnly = True
            ''dtgDatos.Columns(0).DataPropertyName("idempleado")

            'dtgDatos.Columns.Add("departamento", "Departamento")
            'dtgDatos.Columns(1).Width = 100
            'dtgDatos.Columns(1).ReadOnly = True
            'dtgDatos.Columns.Add("nombre", "Trabajador")
            'dtgDatos.Columns(2).Width = 250
            'dtgDatos.Columns(2).ReadOnly = True
            'dtgDatos.Columns.Add("sueldo", "Sueldo Ordinario")
            'dtgDatos.Columns(3).Width = 75
            'dtgDatos.Columns.Add("neto", "Neto")
            'dtgDatos.Columns(4).Width = 75
            'dtgDatos.Columns.Add("infonavit", "Infonavit")
            'dtgDatos.Columns(5).Width = 75
            'dtgDatos.Columns.Add("descuento", "Descuento")
            'dtgDatos.Columns(6).Width = 75
            'dtgDatos.Columns.Add("prestamo", "Prestamo")
            'dtgDatos.Columns(7).Width = 75
            'dtgDatos.Columns.Add("sindicato", "Sindicato")
            'dtgDatos.Columns(8).Width = 75
            'dtgDatos.Columns.Add("neto", "Sueldo Neto")
            'dtgDatos.Columns(9).Width = 75
            'dtgDatos.Columns.Add("imss", "Retención IMSS")
            'dtgDatos.Columns(10).Width = 75
            'dtgDatos.Columns.Add("subsidiado", "Retenciones")
            'dtgDatos.Columns(11).Width = 75
            'dtgDatos.Columns.Add("costosocial", "Costo Social")
            'dtgDatos.Columns(12).Width = 75
            'dtgDatos.Columns.Add("comision", "Comisión")
            'dtgDatos.Columns(13).Width = 75
            'dtgDatos.Columns.Add("subtotal", "Subtotal")
            'dtgDatos.Columns(14).Width = 75
            'dtgDatos.Columns.Add("iva", "IVA")
            'dtgDatos.Columns(15).Width = 75
            'dtgDatos.Columns.Add("total", "Total")
            'dtgDatos.Columns(16).Width = 75


            Dim dsPeriodo As New DataSet
            dsPeriodo.Tables.Add("Tabla")
            dsPeriodo.Tables("Tabla").Columns.Add("Consecutivo")
            dsPeriodo.Tables("Tabla").Columns.Add("Id_empleado")
            dsPeriodo.Tables("Tabla").Columns.Add("CodigoEmpleado")
            dsPeriodo.Tables("Tabla").Columns.Add("Nombre")
            dsPeriodo.Tables("Tabla").Columns.Add("Status")
            dsPeriodo.Tables("Tabla").Columns.Add("RFC")
            dsPeriodo.Tables("Tabla").Columns.Add("CURP")
            dsPeriodo.Tables("Tabla").Columns.Add("Num_IMSS")
            dsPeriodo.Tables("Tabla").Columns.Add("Fecha_Nac")
            dsPeriodo.Tables("Tabla").Columns.Add("Edad")
            dsPeriodo.Tables("Tabla").Columns.Add("Puesto")
            dsPeriodo.Tables("Tabla").Columns.Add("Buque")
            dsPeriodo.Tables("Tabla").Columns.Add("Tipo_Infonavit")
            dsPeriodo.Tables("Tabla").Columns.Add("Valor_Infonavit")
            dsPeriodo.Tables("Tabla").Columns.Add("Sueldo_Base")
            dsPeriodo.Tables("Tabla").Columns.Add("Salario_Diario")
            dsPeriodo.Tables("Tabla").Columns.Add("Salario_Cotización")
            dsPeriodo.Tables("Tabla").Columns.Add("Dias_Trabajados")
            dsPeriodo.Tables("Tabla").Columns.Add("Tipo_Incapacidad")
            dsPeriodo.Tables("Tabla").Columns.Add("Número_días")
            dsPeriodo.Tables("Tabla").Columns.Add("Sueldo_Bruto")
            dsPeriodo.Tables("Tabla").Columns.Add("Tiempo_Extra_Fijo_Gravado")
            dsPeriodo.Tables("Tabla").Columns.Add("Tiempo_Extra_Fijo_Exento")
            dsPeriodo.Tables("Tabla").Columns.Add("Tiempo_Extra_Ocasional")
            dsPeriodo.Tables("Tabla").Columns.Add("Desc_Sem_Obligatorio")
            dsPeriodo.Tables("Tabla").Columns.Add("Vacaciones_proporcionales")
            dsPeriodo.Tables("Tabla").Columns.Add("Aguinaldo_gravado")
            dsPeriodo.Tables("Tabla").Columns.Add("Aguinaldo_exento")
            dsPeriodo.Tables("Tabla").Columns.Add("Total_Aguinaldo")
            dsPeriodo.Tables("Tabla").Columns.Add("Prima_vac_gravado")
            dsPeriodo.Tables("Tabla").Columns.Add("Prima_vac_exento")
            dsPeriodo.Tables("Tabla").Columns.Add("Total_Prima_vac")
            dsPeriodo.Tables("Tabla").Columns.Add("Total_percepciones")
            dsPeriodo.Tables("Tabla").Columns.Add("Total_percepciones_p/isr")
            dsPeriodo.Tables("Tabla").Columns.Add("Incapacidad")
            dsPeriodo.Tables("Tabla").Columns.Add("ISR")
            dsPeriodo.Tables("Tabla").Columns.Add("IMSS")
            dsPeriodo.Tables("Tabla").Columns.Add("Infonavit")
            dsPeriodo.Tables("Tabla").Columns.Add("Infonavit_bim_anterior")
            dsPeriodo.Tables("Tabla").Columns.Add("Ajuste_infonavit")
            dsPeriodo.Tables("Tabla").Columns.Add("Pension_Alimenticia")
            dsPeriodo.Tables("Tabla").Columns.Add("Prestamo")
            dsPeriodo.Tables("Tabla").Columns.Add("Fonacot")
            dsPeriodo.Tables("Tabla").Columns.Add("Subsidio_Generado")
            dsPeriodo.Tables("Tabla").Columns.Add("Subsidio_Aplicado")
            dsPeriodo.Tables("Tabla").Columns.Add("Operadora")
            dsPeriodo.Tables("Tabla").Columns.Add("Prestamo_Personal_A")
            dsPeriodo.Tables("Tabla").Columns.Add("Adeudo_Infonavit_A")
            dsPeriodo.Tables("Tabla").Columns.Add("Diferencia_Infonavit_A")
            dsPeriodo.Tables("Tabla").Columns.Add("Asimilados")
            dsPeriodo.Tables("Tabla").Columns.Add("Retenciones_Operadora")
            dsPeriodo.Tables("Tabla").Columns.Add("%_Comisión")
            dsPeriodo.Tables("Tabla").Columns.Add("Comisión_Operadora")
            dsPeriodo.Tables("Tabla").Columns.Add("Comisión_Asimilados")
            dsPeriodo.Tables("Tabla").Columns.Add("IMSS_CS")
            dsPeriodo.Tables("Tabla").Columns.Add("RCV_CS")
            dsPeriodo.Tables("Tabla").Columns.Add("Infonavit_CS")
            dsPeriodo.Tables("Tabla").Columns.Add("ISN_CS")
            dsPeriodo.Tables("Tabla").Columns.Add("Total_Costo_Social")
            dsPeriodo.Tables("Tabla").Columns.Add("Subtotal")
            dsPeriodo.Tables("Tabla").Columns.Add("IVA")
            dsPeriodo.Tables("Tabla").Columns.Add("TOTAL_DEPOSITO")

           

            'verificamos que no sea una nomina ya guardada como final
            sql = "select * from Nomina inner join EmpleadosC on fkiIdEmpleadoC=iIdEmpleadoC"
            sql &= " where Nomina.fkiIdEmpresa = 1 And fkiIdPeriodo = " & cboperiodo.SelectedValue
            sql &= " and Nomina.iEstatus=1 and iEstatusEmpleado=" & cboserie.SelectedIndex
            sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex
            sql &= " order by " & campoordenamiento 'cNombreLargo"
            'sql = "EXEC getNominaXEmpresaXPeriodo " & gIdEmpresa & "," & cboperiodo.SelectedValue & ",1"

            bCalcular = True
            Dim rwNominaGuardada As DataRow() = nConsulta(sql)

            'If rwNominaGuardadaFinal Is Nothing = False Then
            If rwNominaGuardada Is Nothing = False Then
                'Cargamos los datos de guardados como final
                For x As Integer = 0 To rwNominaGuardada.Count - 1

                    Dim fila As DataRow = dsPeriodo.Tables("Tabla").NewRow

                    fila.Item("Consecutivo") = (x + 1).ToString
                    fila.Item("Id_empleado") = rwNominaGuardada(x)("fkiIdEmpleadoC").ToString





                    fila.Item("CodigoEmpleado") = rwNominaGuardada(x)("cCodigoEmpleado").ToString
                    fila.Item("Nombre") = rwNominaGuardada(x)("cNombreLargo").ToString.ToUpper()
                    fila.Item("Status") = IIf(rwNominaGuardada(x)("iOrigen").ToString = "1", "INTERINO", "PLANTA")
                    fila.Item("RFC") = rwNominaGuardada(x)("cRFC").ToString
                    fila.Item("CURP") = rwNominaGuardada(x)("cCURP").ToString
                    fila.Item("Num_IMSS") = rwNominaGuardada(x)("cIMSS").ToString

                    fila.Item("Fecha_Nac") = Date.Parse(rwNominaGuardada(x)("dFechaNac").ToString).ToShortDateString()
                    'Dim tiempo As TimeSpan = Date.Now - Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString)

                    fila.Item("Edad") = CalcularEdad(Date.Parse(rwNominaGuardada(x)("dFechaNac").ToString).Day, Date.Parse(rwNominaGuardada(x)("dFechaNac").ToString).Month, Date.Parse(rwNominaGuardada(x)("dFechaNac").ToString).Year)
                    fila.Item("Puesto") = rwNominaGuardada(x)("Puesto").ToString
                    fila.Item("Buque") = rwNominaGuardada(x)("Buque").ToString

                    fila.Item("Tipo_Infonavit") = rwNominaGuardada(x)("TipoInfonavit").ToString
                    fila.Item("Valor_Infonavit") = rwNominaGuardada(x)("fValor").ToString
                    '
                    fila.Item("Sueldo_Base") = rwNominaGuardada(x)("fSalarioBase").ToString
                    fila.Item("Salario_Diario") = rwNominaGuardada(x)("fSalarioDiario").ToString
                    fila.Item("Salario_Cotización") = rwNominaGuardada(x)("fSalarioBC").ToString


                    fila.Item("Dias_Trabajados") = rwNominaGuardada(x)("iDiasTrabajados").ToString
                    fila.Item("Tipo_Incapacidad") = rwNominaGuardada(x)("TipoIncapacidad").ToString
                    fila.Item("Número_días") = rwNominaGuardada(x)("iNumeroDias").ToString
                    fila.Item("Sueldo_Bruto") = rwNominaGuardada(x)("fSueldoBruto").ToString
                    fila.Item("Tiempo_Extra_Fijo_Gravado") = rwNominaGuardada(x)("fTExtraFijoGravado").ToString
                    fila.Item("Tiempo_Extra_Fijo_Exento") = rwNominaGuardada(x)("fTExtraFijoExento").ToString
                    fila.Item("Tiempo_Extra_Ocasional") = rwNominaGuardada(x)("fTExtraOcasional").ToString
                    fila.Item("Desc_Sem_Obligatorio") = rwNominaGuardada(x)("fDescSemObligatorio").ToString
                    fila.Item("Vacaciones_proporcionales") = rwNominaGuardada(x)("fVacacionesProporcionales").ToString
                    fila.Item("Aguinaldo_gravado") = rwNominaGuardada(x)("fAguinaldoGravado").ToString
                    fila.Item("Aguinaldo_exento") = rwNominaGuardada(x)("fAguinaldoExento").ToString
                    fila.Item("Total_Aguinaldo") = rwNominaGuardada(x)("fAguinaldoGravado").ToString + rwNominaGuardada(x)("fAguinaldoExento").ToString
                    fila.Item("Prima_vac_gravado") = rwNominaGuardada(x)("fPrimaVacacionalGravado").ToString
                    fila.Item("Prima_vac_exento") = rwNominaGuardada(x)("fPrimaVacacionalExento").ToString

                    fila.Item("Total_Prima_vac") = rwNominaGuardada(x)("fPrimaVacacionalGravado").ToString + rwNominaGuardada(x)("fPrimaVacacionalExento").ToString
                    fila.Item("Total_percepciones") = rwNominaGuardada(x)("fTotalPercepciones").ToString
                    fila.Item("Total_percepciones_p/isr") = rwNominaGuardada(x)("fTotalPercepcionesISR").ToString
                    fila.Item("Incapacidad") = rwNominaGuardada(x)("fIncapacidad").ToString
                    fila.Item("ISR") = rwNominaGuardada(x)("fIsr").ToString
                    fila.Item("IMSS") = rwNominaGuardada(x)("fImss").ToString
                    fila.Item("Infonavit") = rwNominaGuardada(x)("fInfonavit").ToString
                    fila.Item("Infonavit_bim_anterior") = rwNominaGuardada(x)("fInfonavitBanterior").ToString
                    fila.Item("Ajuste_infonavit") = rwNominaGuardada(x)("fAjusteInfonavit").ToString
                    fila.Item("Pension_Alimenticia") = rwNominaGuardada(x)("fPensionAlimenticia").ToString
                    fila.Item("Prestamo") = rwNominaGuardada(x)("fPrestamo").ToString
                    fila.Item("Fonacot") = rwNominaGuardada(x)("fFonacot").ToString
                    fila.Item("Subsidio_Generado") = rwNominaGuardada(x)("fSubsidioGenerado").ToString
                    fila.Item("Subsidio_Aplicado") = rwNominaGuardada(x)("fSubsidioAplicado").ToString
                    fila.Item("Operadora") = rwNominaGuardada(x)("fOperadora").ToString
                    fila.Item("Prestamo_Personal_A") = rwNominaGuardada(x)("fPrestamoPerA").ToString
                    fila.Item("Adeudo_Infonavit_A") = rwNominaGuardada(x)("fAdeudoInfonavitA").ToString
                    fila.Item("Diferencia_Infonavit_A") = rwNominaGuardada(x)("fDiferenciaInfonavitA").ToString
                    fila.Item("Asimilados") = rwNominaGuardada(x)("fAsimilados").ToString
                    fila.Item("Retenciones_Operadora") = rwNominaGuardada(x)("fRetencionOperadora").ToString
                    fila.Item("%_Comisión") = rwNominaGuardada(x)("fPorComision").ToString
                    fila.Item("Comisión_Operadora") = rwNominaGuardada(x)("fComisionOperadora").ToString
                    fila.Item("Comisión_Asimilados") = rwNominaGuardada(x)("fComisionAsimilados").ToString
                    fila.Item("IMSS_CS") = rwNominaGuardada(x)("fImssCS").ToString
                    fila.Item("RCV_CS") = rwNominaGuardada(x)("fRcvCS").ToString
                    fila.Item("Infonavit_CS") = rwNominaGuardada(x)("fInfonavitCS").ToString
                    fila.Item("ISN_CS") = rwNominaGuardada(x)("fInsCS").ToString
                    fila.Item("Total_Costo_Social") = rwNominaGuardada(x)("fTotalCostoSocial").ToString
                    fila.Item("Subtotal") = rwNominaGuardada(x)("fSubtotal").ToString
                    fila.Item("IVA") = rwNominaGuardada(x)("fIVA").ToString
                    fila.Item("TOTAL_DEPOSITO") = rwNominaGuardada(x)("fTotalDeposito").ToString


                    dsPeriodo.Tables("Tabla").Rows.Add(fila)
                Next

                dtgDatos.DataSource = dsPeriodo.Tables("Tabla")

                dtgDatos.Columns(0).Width = 30
                dtgDatos.Columns(0).ReadOnly = True
                dtgDatos.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                'consecutivo
                dtgDatos.Columns(1).Width = 60
                dtgDatos.Columns(1).ReadOnly = True
                dtgDatos.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'idempleado
                dtgDatos.Columns(2).Width = 100
                dtgDatos.Columns(2).ReadOnly = True
                dtgDatos.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'codigo empleado
                dtgDatos.Columns(3).Width = 100
                dtgDatos.Columns(3).ReadOnly = True
                dtgDatos.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'Nombre
                dtgDatos.Columns(4).Width = 250
                dtgDatos.Columns(4).ReadOnly = True
                'Estatus
                dtgDatos.Columns(5).Width = 100
                dtgDatos.Columns(5).ReadOnly = True
                'RFC
                dtgDatos.Columns(6).Width = 100
                dtgDatos.Columns(6).ReadOnly = True
                'dtgDatos.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                'CURP
                dtgDatos.Columns(7).Width = 150
                dtgDatos.Columns(7).ReadOnly = True
                'IMSS 

                dtgDatos.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(8).ReadOnly = True
                'Fecha_Nac
                dtgDatos.Columns(9).Width = 150
                dtgDatos.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(9).ReadOnly = True

                'Edad
                dtgDatos.Columns(10).ReadOnly = True
                dtgDatos.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                'Puesto
                dtgDatos.Columns(11).ReadOnly = True
                dtgDatos.Columns(11).Width = 200
                dtgDatos.Columns.Remove("Puesto")

                Dim combo As New DataGridViewComboBoxColumn

                sql = "select * from puestos where iTipo=1 order by cNombre"

                'Dim rwPuestos As DataRow() = nConsulta(sql)
                'If rwPuestos Is Nothing = False Then
                '    combo.Items.Add("uno")
                '    combo.Items.Add("dos")
                '    combo.Items.Add("tres")
                'End If

                nCargaCBO(combo, sql, "cNombre", "iIdPuesto")

                combo.HeaderText = "Puesto"

                combo.Width = 150
                dtgDatos.Columns.Insert(11, combo)
                'DirectCast(dtgDatos.Columns(11), DataGridViewComboBoxColumn).Sorted = True
                'Dim combo2 As New DataGridViewComboBoxCell
                'combo2 = CType(Me.dtgDatos.Rows(2).Cells(11), DataGridViewComboBoxCell)
                'combo2.Value = combo.Items(11)



                'dtgDatos.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                'Buque
                'dtgDatos.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(12).ReadOnly = True
                dtgDatos.Columns(12).Width = 150
                dtgDatos.Columns.Remove("Buque")

                Dim combo2 As New DataGridViewComboBoxColumn

                sql = "select * from departamentos where iEstatus=1 order by cNombre"

                'Dim rwPuestos As DataRow() = nConsulta(sql)
                'If rwPuestos Is Nothing = False Then
                '    combo.Items.Add("uno")
                '    combo.Items.Add("dos")
                '    combo.Items.Add("tres")
                'End If

                nCargaCBO(combo2, sql, "cNombre", "iIdDepartamento")

                combo2.HeaderText = "Buque"
                combo2.Width = 150
                dtgDatos.Columns.Insert(12, combo2)

                'Tipo_Infonavit
                dtgDatos.Columns(13).ReadOnly = True
                dtgDatos.Columns(13).Width = 150
                'dtgDatos.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight



                'Valor_Infonavit
                dtgDatos.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(14).ReadOnly = True
                dtgDatos.Columns(14).Width = 150
                'Sueldo_Base
                dtgDatos.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(15).ReadOnly = True
                dtgDatos.Columns(15).Width = 150
                'Salario_Diario
                dtgDatos.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(16).ReadOnly = True
                dtgDatos.Columns(16).Width = 150
                'Salario_Cotización
                dtgDatos.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(17).ReadOnly = True
                dtgDatos.Columns(17).Width = 150
                'Dias_Trabajados
                dtgDatos.Columns(18).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(18).Width = 150
                'Tipo_Incapacidad
                dtgDatos.Columns(19).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(19).ReadOnly = True
                dtgDatos.Columns(19).Width = 150
                'Número_días
                dtgDatos.Columns(20).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(20).ReadOnly = True
                dtgDatos.Columns(20).Width = 150
                'Sueldo_Bruto
                dtgDatos.Columns(21).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(21).ReadOnly = True
                dtgDatos.Columns(21).Width = 150
                'Tiempo_Extra_Fijo_Gravado
                dtgDatos.Columns(22).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(22).ReadOnly = True
                dtgDatos.Columns(22).Width = 150

                'Tiempo_Extra_Fijo_Exento
                dtgDatos.Columns(23).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(23).ReadOnly = True
                dtgDatos.Columns(23).Width = 150

                'Tiempo_Extra_Ocasional
                dtgDatos.Columns(24).Width = 150
                dtgDatos.Columns(24).ReadOnly = True
                dtgDatos.Columns(24).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'Desc_Sem_Obligatorio
                dtgDatos.Columns(25).Width = 150
                dtgDatos.Columns(25).ReadOnly = True
                dtgDatos.Columns(25).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'Vacaciones_proporcionales
                dtgDatos.Columns(26).Width = 150
                dtgDatos.Columns(26).ReadOnly = True
                dtgDatos.Columns(26).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'Aguinaldo_gravado
                dtgDatos.Columns(27).Width = 150
                dtgDatos.Columns(27).ReadOnly = True
                dtgDatos.Columns(27).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'Aguinaldo_exento
                dtgDatos.Columns(28).Width = 150
                dtgDatos.Columns(28).ReadOnly = True
                dtgDatos.Columns(28).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'Total_Aguinaldo
                dtgDatos.Columns(29).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(29).Width = 150
                dtgDatos.Columns(29).ReadOnly = True

                'Prima_vac_gravado
                dtgDatos.Columns(30).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(30).ReadOnly = True
                dtgDatos.Columns(30).Width = 150
                'Prima_vac_exento 
                dtgDatos.Columns(31).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(31).ReadOnly = True
                dtgDatos.Columns(31).Width = 150

                'Total_Prima_vac
                dtgDatos.Columns(32).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(32).ReadOnly = True
                dtgDatos.Columns(32).Width = 150


                'Total_percepciones
                dtgDatos.Columns(33).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(33).ReadOnly = True
                dtgDatos.Columns(33).Width = 150
                'Total_percepciones_p/isr
                dtgDatos.Columns(34).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(34).ReadOnly = True
                dtgDatos.Columns(34).Width = 150

                'Incapacidad
                dtgDatos.Columns(35).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(35).ReadOnly = True
                dtgDatos.Columns(35).Width = 150

                'ISR
                dtgDatos.Columns(36).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(36).ReadOnly = True
                dtgDatos.Columns(36).Width = 150


                'IMSS
                dtgDatos.Columns(37).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(37).ReadOnly = True
                dtgDatos.Columns(37).Width = 150

                'Infonavit
                dtgDatos.Columns(38).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(38).ReadOnly = True
                dtgDatos.Columns(38).Width = 150
                'Infonavit_bim_anterior
                dtgDatos.Columns(39).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(39).ReadOnly = True
                dtgDatos.Columns(39).Width = 150
                'Ajuste_infonavit
                dtgDatos.Columns(40).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(40).ReadOnly = True
                dtgDatos.Columns(40).Width = 150
                'Pension_Alimenticia
                dtgDatos.Columns(41).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(40).ReadOnly = True
                dtgDatos.Columns(41).Width = 150
                'Prestamo
                dtgDatos.Columns(42).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(42).ReadOnly = True
                dtgDatos.Columns(42).Width = 150
                'Fonacot
                dtgDatos.Columns(43).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(43).ReadOnly = True
                dtgDatos.Columns(43).Width = 150
                'Subsidio_Generado
                dtgDatos.Columns(44).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(44).ReadOnly = True
                dtgDatos.Columns(44).Width = 150
                'Subsidio_Aplicado
                dtgDatos.Columns(45).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(45).ReadOnly = True
                dtgDatos.Columns(45).Width = 150
                'Operadora
                dtgDatos.Columns(46).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(46).ReadOnly = True
                dtgDatos.Columns(46).Width = 150

                'Prestamo Personal Asimilado
                dtgDatos.Columns(47).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(48).ReadOnly = True
                dtgDatos.Columns(47).Width = 150

                'Adeudo_Infonavit_Asimilado
                dtgDatos.Columns(48).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(49).ReadOnly = True
                dtgDatos.Columns(48).Width = 150

                'Difencia infonavit Asimilado
                dtgDatos.Columns(49).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(50).ReadOnly = True
                dtgDatos.Columns(49).Width = 150

                'Complemento Asimilado
                dtgDatos.Columns(50).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(50).ReadOnly = True
                dtgDatos.Columns(50).Width = 150

                'Retenciones_Operadora
                dtgDatos.Columns(51).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(51).ReadOnly = True
                dtgDatos.Columns(51).Width = 150

                '% Comision
                dtgDatos.Columns(52).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(52).ReadOnly = True
                dtgDatos.Columns(52).Width = 150

                'Comision_Operadora
                dtgDatos.Columns(53).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(53).ReadOnly = True
                dtgDatos.Columns(53).Width = 150

                'Comision asimilados
                dtgDatos.Columns(54).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(54).ReadOnly = True
                dtgDatos.Columns(54).Width = 150

                'IMSS_CS
                dtgDatos.Columns(55).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(55).ReadOnly = True
                dtgDatos.Columns(55).Width = 150

                'RCV_CS
                dtgDatos.Columns(56).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(56).ReadOnly = True
                dtgDatos.Columns(56).Width = 150

                'Infonavit_CS
                dtgDatos.Columns(57).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(57).ReadOnly = True
                dtgDatos.Columns(57).Width = 150

                'ISN_CS
                dtgDatos.Columns(58).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(58).ReadOnly = True
                dtgDatos.Columns(58).Width = 150

                'Total Costo Social
                dtgDatos.Columns(59).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(59).ReadOnly = True
                dtgDatos.Columns(59).Width = 150

                'Subtotal
                dtgDatos.Columns(60).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(60).ReadOnly = True
                dtgDatos.Columns(60).Width = 150

                'IVA
                dtgDatos.Columns(61).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(61).ReadOnly = True
                dtgDatos.Columns(61).Width = 150

                'TOTAL DEPOSITO
                dtgDatos.Columns(62).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(62).ReadOnly = True
                dtgDatos.Columns(62).Width = 150

                'calcular()

                'Cambiamos index del combo en el grid

                For x As Integer = 0 To dtgDatos.Rows.Count - 1

                    sql = "select * from nomina where fkiIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                    sql &= " and fkiIdPeriodo=" & cboperiodo.SelectedValue
                    sql &= " and iEstatusEmpleado=" & cboserie.SelectedIndex
                    sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex
                    Dim rwFila As DataRow() = nConsulta(sql)



                    CType(Me.dtgDatos.Rows(x).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("Puesto").ToString()
                    CType(Me.dtgDatos.Rows(x).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("Buque").ToString()
                Next


                'Cambiamos el index del combro de departamentos

                'For x As Integer = 0 To dtgDatos.Rows.Count - 1

                '    sql = "select * from empleadosC where iIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                '    Dim rwFila As DataRow() = nConsulta(sql)




                'Next

                MessageBox.Show("Datos cargados", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)


            Else

                If cboTipoNomina.SelectedIndex = 0 Then
                    If cboserie.SelectedIndex = 0 Then
                        'Buscamos los datos de sindicato solamente
                        sql = "select  * from empleadosC where fkiIdClienteInter=-1"
                        'sql = "select iIdEmpleadoC,NumCuenta, (cApellidoP + ' ' + cApellidoM + ' ' + cNombre) as nombre, fkiIdEmpresa,fSueldoOrd,fCosto from empleadosC"
                        'sql &= " where empleadosC.iOrigen=2 and empleadosC.iEstatus=1"
                        'sql &= " and empleadosC.fkiIdEmpresa =" & gIdEmpresa
                        sql &= " order by cFuncionesPuesto,cNombreLargo"

                    ElseIf cboserie.SelectedIndex > 0 Or cboserie.SelectedIndex - 1 Then
                        sql = "select * from Nomina inner join EmpleadosC on fkiIdEmpleadoC=iIdEmpleadoC"
                        sql &= " where Nomina.fkiIdEmpresa = 1 And fkiIdPeriodo = " & cboperiodo.SelectedValue
                        sql &= " and Nomina.iEstatus=1 and iEstatusEmpleado=20"
                        sql &= " order by cNombreLargo"

                    End If


                    Dim rwDatosEmpleados As DataRow() = nConsulta(sql)
                    If rwDatosEmpleados Is Nothing = False Then
                        For x As Integer = 0 To rwDatosEmpleados.Length - 1


                            Dim fila As DataRow = dsPeriodo.Tables("Tabla").NewRow

                            fila.Item("Consecutivo") = (x + 1).ToString
                            fila.Item("Id_empleado") = rwDatosEmpleados(x)("iIdEmpleadoC").ToString
                            fila.Item("CodigoEmpleado") = rwDatosEmpleados(x)("cCodigoEmpleado").ToString
                            fila.Item("Nombre") = rwDatosEmpleados(x)("cNombreLargo").ToString.ToUpper()
                            fila.Item("Status") = IIf(rwDatosEmpleados(x)("iOrigen").ToString = "1", "INTERINO", "PLANTA")
                            fila.Item("RFC") = rwDatosEmpleados(x)("cRFC").ToString
                            fila.Item("CURP") = rwDatosEmpleados(x)("cCURP").ToString
                            fila.Item("Num_IMSS") = rwDatosEmpleados(x)("cIMSS").ToString

                            fila.Item("Fecha_Nac") = Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString).ToShortDateString()
                            'Dim tiempo As TimeSpan = Date.Now - Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString)
                            fila.Item("Edad") = CalcularEdad(Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString).Day, Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString).Month, Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString).Year)
                            fila.Item("Puesto") = rwDatosEmpleados(x)("cPuesto").ToString
                            fila.Item("Buque") = "ECO III"

                            fila.Item("Tipo_Infonavit") = rwDatosEmpleados(x)("cTipoFactor").ToString
                            fila.Item("Valor_Infonavit") = rwDatosEmpleados(x)("fFactor").ToString
                            fila.Item("Sueldo_Base") = "0.00"
                            fila.Item("Salario_Diario") = rwDatosEmpleados(x)("fSueldoBase").ToString
                            fila.Item("Salario_Cotización") = rwDatosEmpleados(x)("fSueldoIntegrado").ToString
                            fila.Item("Dias_Trabajados") = "30"
                            fila.Item("Tipo_Incapacidad") = TipoIncapacidad(rwDatosEmpleados(x)("iIdEmpleadoC").ToString, cboperiodo.SelectedValue)
                            fila.Item("Número_días") = NumDiasIncapacidad(rwDatosEmpleados(x)("iIdEmpleadoC").ToString, cboperiodo.SelectedValue)
                            fila.Item("Sueldo_Bruto") = ""
                            fila.Item("Tiempo_Extra_Fijo_Gravado") = ""
                            fila.Item("Tiempo_Extra_Fijo_Exento") = ""
                            fila.Item("Tiempo_Extra_Ocasional") = ""
                            fila.Item("Desc_Sem_Obligatorio") = ""
                            fila.Item("Vacaciones_proporcionales") = ""
                            fila.Item("Aguinaldo_gravado") = ""
                            fila.Item("Aguinaldo_exento") = ""
                            fila.Item("Total_Aguinaldo") = ""
                            fila.Item("Prima_vac_gravado") = ""
                            fila.Item("Prima_vac_exento") = ""

                            fila.Item("Total_Prima_vac") = ""
                            fila.Item("Total_percepciones") = ""
                            fila.Item("Total_percepciones_p/isr") = ""
                            fila.Item("Incapacidad") = ""
                            fila.Item("ISR") = ""
                            fila.Item("IMSS") = ""
                            fila.Item("Infonavit") = ""
                            fila.Item("Infonavit_bim_anterior") = ""
                            fila.Item("Ajuste_infonavit") = ""
                            fila.Item("Pension_Alimenticia") = ""
                            fila.Item("Prestamo") = ""
                            fila.Item("Fonacot") = ""
                            fila.Item("Subsidio_Generado") = ""
                            fila.Item("Subsidio_Aplicado") = ""
                            fila.Item("Operadora") = ""
                            fila.Item("Prestamo_Personal_A") = ""
                            fila.Item("Adeudo_Infonavit_A") = ""
                            fila.Item("Diferencia_Infonavit_A") = ""
                            fila.Item("Asimilados") = ""
                            fila.Item("Retenciones_Operadora") = ""
                            fila.Item("%_Comisión") = ""
                            fila.Item("Comisión_Operadora") = ""
                            fila.Item("Comisión_Asimilados") = ""
                            fila.Item("IMSS_CS") = ""
                            fila.Item("RCV_CS") = ""
                            fila.Item("Infonavit_CS") = ""
                            fila.Item("ISN_CS") = ""
                            fila.Item("Total_Costo_Social") = ""
                            fila.Item("Subtotal") = ""
                            fila.Item("IVA") = ""
                            fila.Item("TOTAL_DEPOSITO") = ""


                            dsPeriodo.Tables("Tabla").Rows.Add(fila)




                        Next




                        dtgDatos.DataSource = dsPeriodo.Tables("Tabla")

                        dtgDatos.Columns(0).Width = 30
                        dtgDatos.Columns(0).ReadOnly = True
                        dtgDatos.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                        'consecutivo
                        dtgDatos.Columns(1).Width = 60
                        dtgDatos.Columns(1).ReadOnly = True
                        dtgDatos.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'idempleado
                        dtgDatos.Columns(2).Width = 100
                        dtgDatos.Columns(2).ReadOnly = True
                        dtgDatos.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'codigo empleado
                        dtgDatos.Columns(3).Width = 100
                        dtgDatos.Columns(3).ReadOnly = True
                        dtgDatos.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Nombre
                        dtgDatos.Columns(4).Width = 250
                        dtgDatos.Columns(4).ReadOnly = True
                        'Estatus
                        dtgDatos.Columns(5).Width = 100
                        dtgDatos.Columns(5).ReadOnly = True
                        'RFC
                        dtgDatos.Columns(6).Width = 100
                        dtgDatos.Columns(6).ReadOnly = True
                        'dtgDatos.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                        'CURP
                        dtgDatos.Columns(7).Width = 150
                        dtgDatos.Columns(7).ReadOnly = True
                        'IMSS 

                        dtgDatos.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(8).ReadOnly = True
                        'Fecha_Nac
                        dtgDatos.Columns(9).Width = 150
                        dtgDatos.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(9).ReadOnly = True

                        'Edad
                        dtgDatos.Columns(10).ReadOnly = True
                        dtgDatos.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                        'Puesto
                        dtgDatos.Columns(11).ReadOnly = True
                        dtgDatos.Columns(11).Width = 200
                        dtgDatos.Columns.Remove("Puesto")

                        Dim combo As New DataGridViewComboBoxColumn

                        sql = "select * from puestos where iTipo=1 order by cNombre"

                        'Dim rwPuestos As DataRow() = nConsulta(sql)
                        'If rwPuestos Is Nothing = False Then
                        '    combo.Items.Add("uno")
                        '    combo.Items.Add("dos")
                        '    combo.Items.Add("tres")
                        'End If

                        nCargaCBO(combo, sql, "cNombre", "iIdPuesto")

                        combo.HeaderText = "Puesto"

                        combo.Width = 150
                        dtgDatos.Columns.Insert(11, combo)
                        'DirectCast(dtgDatos.Columns(11), DataGridViewComboBoxColumn).Sorted = True
                        'Dim combo2 As New DataGridViewComboBoxCell
                        'combo2 = CType(Me.dtgDatos.Rows(2).Cells(11), DataGridViewComboBoxCell)
                        'combo2.Value = combo.Items(11)



                        'dtgDatos.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                        'Buque
                        'dtgDatos.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(12).ReadOnly = True
                        dtgDatos.Columns(12).Width = 150
                        dtgDatos.Columns.Remove("Buque")

                        Dim combo2 As New DataGridViewComboBoxColumn

                        sql = "select * from departamentos where iEstatus=1 order by cNombre"

                        'Dim rwPuestos As DataRow() = nConsulta(sql)
                        'If rwPuestos Is Nothing = False Then
                        '    combo.Items.Add("uno")
                        '    combo.Items.Add("dos")
                        '    combo.Items.Add("tres")
                        'End If

                        nCargaCBO(combo2, sql, "cNombre", "iIdDepartamento")

                        combo2.HeaderText = "Buque"
                        combo2.Width = 150
                        dtgDatos.Columns.Insert(12, combo2)

                        'Tipo_Infonavit
                        dtgDatos.Columns(13).ReadOnly = True
                        dtgDatos.Columns(13).Width = 150
                        'dtgDatos.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight



                        'Valor_Infonavit
                        dtgDatos.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(14).ReadOnly = True
                        dtgDatos.Columns(14).Width = 150
                        'Sueldo_Base
                        dtgDatos.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(15).ReadOnly = True
                        dtgDatos.Columns(15).Width = 150
                        'Salario_Diario
                        dtgDatos.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(16).ReadOnly = True
                        dtgDatos.Columns(16).Width = 150
                        'Salario_Cotización
                        dtgDatos.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(17).ReadOnly = True
                        dtgDatos.Columns(17).Width = 150
                        'Dias_Trabajados
                        dtgDatos.Columns(18).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(18).Width = 150
                        'Tipo_Incapacidad
                        dtgDatos.Columns(19).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(19).ReadOnly = True
                        dtgDatos.Columns(19).Width = 150
                        'Número_días
                        dtgDatos.Columns(20).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(20).ReadOnly = True
                        dtgDatos.Columns(20).Width = 150
                        'Sueldo_Bruto
                        dtgDatos.Columns(21).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(21).ReadOnly = True
                        dtgDatos.Columns(21).Width = 150
                        'Tiempo_Extra_Fijo_Gravado
                        dtgDatos.Columns(22).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(22).ReadOnly = True
                        dtgDatos.Columns(22).Width = 150

                        'Tiempo_Extra_Fijo_Exento
                        dtgDatos.Columns(23).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(23).ReadOnly = True
                        dtgDatos.Columns(23).Width = 150

                        'Tiempo_Extra_Ocasional
                        dtgDatos.Columns(24).Width = 150
                        dtgDatos.Columns(24).ReadOnly = True
                        dtgDatos.Columns(24).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Desc_Sem_Obligatorio
                        dtgDatos.Columns(25).Width = 150
                        dtgDatos.Columns(25).ReadOnly = True
                        dtgDatos.Columns(25).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Vacaciones_proporcionales
                        dtgDatos.Columns(26).Width = 150
                        dtgDatos.Columns(26).ReadOnly = True
                        dtgDatos.Columns(26).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Aguinaldo_gravado
                        dtgDatos.Columns(27).Width = 150
                        dtgDatos.Columns(27).ReadOnly = True
                        dtgDatos.Columns(27).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Aguinaldo_exento
                        dtgDatos.Columns(28).Width = 150
                        dtgDatos.Columns(28).ReadOnly = True
                        dtgDatos.Columns(28).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Total_Aguinaldo
                        dtgDatos.Columns(29).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(29).Width = 150
                        dtgDatos.Columns(29).ReadOnly = True

                        'Prima_vac_gravado
                        dtgDatos.Columns(30).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(30).ReadOnly = True
                        dtgDatos.Columns(30).Width = 150
                        'Prima_vac_exento 
                        dtgDatos.Columns(31).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(31).ReadOnly = True
                        dtgDatos.Columns(31).Width = 150

                        'Total_Prima_vac
                        dtgDatos.Columns(32).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(32).ReadOnly = True
                        dtgDatos.Columns(32).Width = 150


                        'Total_percepciones
                        dtgDatos.Columns(33).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(33).ReadOnly = True
                        dtgDatos.Columns(33).Width = 150
                        'Total_percepciones_p/isr
                        dtgDatos.Columns(34).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(34).ReadOnly = True
                        dtgDatos.Columns(34).Width = 150

                        'Incapacidad
                        dtgDatos.Columns(35).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(35).ReadOnly = True
                        dtgDatos.Columns(35).Width = 150

                        'ISR
                        dtgDatos.Columns(36).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(36).ReadOnly = True
                        dtgDatos.Columns(36).Width = 150


                        'IMSS
                        dtgDatos.Columns(37).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(37).ReadOnly = True
                        dtgDatos.Columns(37).Width = 150

                        'Infonavit
                        dtgDatos.Columns(38).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(38).ReadOnly = True
                        dtgDatos.Columns(38).Width = 150
                        'Infonavit_bim_anterior
                        dtgDatos.Columns(39).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(39).ReadOnly = True
                        dtgDatos.Columns(39).Width = 150
                        'Ajuste_infonavit
                        dtgDatos.Columns(40).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(40).ReadOnly = True
                        dtgDatos.Columns(40).Width = 150
                        'Pension_Alimenticia
                        dtgDatos.Columns(41).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(40).ReadOnly = True
                        dtgDatos.Columns(41).Width = 150
                        'Prestamo
                        dtgDatos.Columns(42).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(42).ReadOnly = True
                        dtgDatos.Columns(42).Width = 150
                        'Fonacot
                        dtgDatos.Columns(43).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(43).ReadOnly = True
                        dtgDatos.Columns(43).Width = 150
                        'Subsidio_Generado
                        dtgDatos.Columns(44).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(44).ReadOnly = True
                        dtgDatos.Columns(44).Width = 150
                        'Subsidio_Aplicado
                        dtgDatos.Columns(45).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(45).ReadOnly = True
                        dtgDatos.Columns(45).Width = 150
                        'Operadora
                        dtgDatos.Columns(46).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(46).ReadOnly = True
                        dtgDatos.Columns(46).Width = 150

                        'Prestamo Personal Asimilado
                        dtgDatos.Columns(47).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(48).ReadOnly = True
                        dtgDatos.Columns(47).Width = 150

                        'Adeudo_Infonavit_Asimilado
                        dtgDatos.Columns(48).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(49).ReadOnly = True
                        dtgDatos.Columns(48).Width = 150

                        'Difencia infonavit Asimilado
                        dtgDatos.Columns(49).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(50).ReadOnly = True
                        dtgDatos.Columns(49).Width = 150

                        'Complemento Asimilado
                        dtgDatos.Columns(50).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(50).ReadOnly = True
                        dtgDatos.Columns(50).Width = 150

                        'Retenciones_Operadora
                        dtgDatos.Columns(51).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(51).ReadOnly = True
                        dtgDatos.Columns(51).Width = 150

                        '% Comision
                        dtgDatos.Columns(52).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(52).ReadOnly = True
                        dtgDatos.Columns(52).Width = 150

                        'Comision_Operadora
                        dtgDatos.Columns(53).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(53).ReadOnly = True
                        dtgDatos.Columns(53).Width = 150

                        'Comision asimilados
                        dtgDatos.Columns(54).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(54).ReadOnly = True
                        dtgDatos.Columns(54).Width = 150

                        'IMSS_CS
                        dtgDatos.Columns(55).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(55).ReadOnly = True
                        dtgDatos.Columns(55).Width = 150

                        'RCV_CS
                        dtgDatos.Columns(56).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(56).ReadOnly = True
                        dtgDatos.Columns(56).Width = 150

                        'Infonavit_CS
                        dtgDatos.Columns(57).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(57).ReadOnly = True
                        dtgDatos.Columns(57).Width = 150

                        'ISN_CS
                        dtgDatos.Columns(58).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(58).ReadOnly = True
                        dtgDatos.Columns(58).Width = 150

                        'Total Costo Social
                        dtgDatos.Columns(59).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(59).ReadOnly = True
                        dtgDatos.Columns(59).Width = 150

                        'Subtotal
                        dtgDatos.Columns(60).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(60).ReadOnly = True
                        dtgDatos.Columns(60).Width = 150

                        'IVA
                        dtgDatos.Columns(61).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(61).ReadOnly = True
                        dtgDatos.Columns(61).Width = 150

                        'TOTAL DEPOSITO
                        dtgDatos.Columns(62).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(62).ReadOnly = True
                        dtgDatos.Columns(62).Width = 150
                        'calcular()

                        'Cambiamos index del combo en el grid

                        For x As Integer = 0 To dtgDatos.Rows.Count - 1

                            sql = "select * from empleadosC where iIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                            Dim rwFila As DataRow() = nConsulta(sql)



                            CType(Me.dtgDatos.Rows(x).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("cPuesto").ToString()
                            CType(Me.dtgDatos.Rows(x).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("cFuncionesPuesto").ToString()
                        Next


                        'Cambiamos el index del combro de departamentos

                        'For x As Integer = 0 To dtgDatos.Rows.Count - 1

                        '    sql = "select * from empleadosC where iIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                        '    Dim rwFila As DataRow() = nConsulta(sql)




                        'Next


                        MessageBox.Show("Datos cargados", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("No hay datos en este período", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If




                    'No hay datos en este período
                Else
                    MessageBox.Show("Para la nomina Descanso, solo se mostraran datos guardados, no se podrá calcular de 0", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If


                

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Function TipoIncapacidad(idempleado As String, periodo As Integer) As String
        Dim sql As String
        Dim cadena As String = "Ninguno"

        Try
            sql = "select * from periodos where iIdPeriodo= " & periodo
            Dim rwPeriodo As DataRow() = nConsulta(Sql)

            If rwPeriodo Is Nothing = False Then

                sql = "select * from incapacidad where iIdIncapacidad= "
                sql &= " (select Max(iIdIncapacidad) from incapacidad where iEstatus=1 and fkiIdEmpleado=" & idempleado & ") "
                Dim rwIncapacidad As DataRow() = nConsulta(sql)

                If rwIncapacidad Is Nothing = False Then
                    Dim FechaBuscar As Date = Date.Parse(rwIncapacidad(0)("FechaInicio"))
                    Dim FechaInicial As Date = Date.Parse(rwPeriodo(0)("dFechaInicio"))
                    Dim FechaFinal As Date = Date.Parse(rwPeriodo(0)("dFechaFin"))
                    'Dim FechaAntiguedad As Date = Date.Parse(rwDatosBanco(0)("dFechaAntiguedad"))

                    If FechaBuscar.CompareTo(FechaInicial) >= 0 And FechaBuscar.CompareTo(FechaFinal) <= 0 Then
                        'Estamos dentro del rango inicial
                        Return Identificadorincapacidad(rwIncapacidad(0)("RamoRiesgo"))

                    ElseIf FechaBuscar.CompareTo(FechaInicial) <= 0 Then
                        FechaBuscar = Date.Parse(rwIncapacidad(0)("fechafin"))
                        If FechaBuscar.CompareTo(FechaFinal) <= 0 Then
                            Return Identificadorincapacidad(rwIncapacidad(0)("RamoRiesgo"))
                        End If

                    End If

                Else
                    cadena = "Ninguno"
                    Return cadena
                End If

                
            Else
                Return "Ninguno"

            End If
            Return "Ninguno"
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        
    End Function

    Private Function NumDiasIncapacidad(idempleado As String, periodo As Integer) As String
        Dim sql As String
        Dim cadena As String

        Try
            sql = "select * from periodos where iIdPeriodo= " & periodo
            Dim rwPeriodo As DataRow() = nConsulta(sql)

            If rwPeriodo Is Nothing = False Then

                sql = "select * from incapacidad where iIdIncapacidad= "
                sql &= " (select Max(iIdIncapacidad) from incapacidad where iEstatus=1 and fkiIdEmpleado=" & idempleado & ") "
                Dim rwIncapacidad As DataRow() = nConsulta(sql)

                If rwIncapacidad Is Nothing = False Then
                    Dim FechaBuscar As Date = Date.Parse(rwIncapacidad(0)("FechaInicio"))
                    Dim FechaInicial As Date = Date.Parse(rwPeriodo(0)("dFechaInicio"))
                    Dim FechaFinal As Date = Date.Parse(rwPeriodo(0)("dFechaFin"))
                    'Dim FechaAntiguedad As Date = Date.Parse(rwDatosBanco(0)("dFechaAntiguedad"))

                    If FechaBuscar.CompareTo(FechaInicial) >= 0 And FechaBuscar.CompareTo(FechaFinal) <= 0 Then
                        'Estamos dentro del rango inicial
                        FechaBuscar = Date.Parse(rwIncapacidad(0)("fechafin"))
                        If FechaBuscar.CompareTo(FechaFinal) <= 0 Then
                            'Restamos entre final incapacidad menos la inicial incapacidad
                            Return (DateDiff(DateInterval.Day, Date.Parse(rwIncapacidad(0)("FechaInicio")), Date.Parse(rwIncapacidad(0)("fechafin"))) + 1).ToString
                        Else
                            'restamos final del periodo menos inicial incapacidad
                            Return (DateDiff(DateInterval.Day, Date.Parse(rwIncapacidad(0)("FechaInicio")), Date.Parse(rwPeriodo(0)("dFechaFin"))) + 1).ToString


                        End If

                    ElseIf FechaBuscar.CompareTo(FechaInicial) <= 0 Then
                        FechaBuscar = Date.Parse(rwIncapacidad(0)("fechafin"))
                        If FechaBuscar.CompareTo(FechaFinal) <= 0 Then
                            'Restamos fecha final incapacidad menos la fechainicial  periodo
                            Return (DateDiff(DateInterval.Day, Date.Parse(rwPeriodo(0)("dFechaInicio")), Date.Parse(rwIncapacidad(0)("fechafin"))) + 1).ToString
                        Else
                            'todos los dias del periodo tiene incapaciddad
                            Return (DateDiff(DateInterval.Day, Date.Parse(rwPeriodo(0)("dFechaInicio")), Date.Parse(rwPeriodo(0)("dFechaFin"))) + 1).ToString
                        End If

                    End If
                Else
                    cadena = "0"
                    Return cadena
                End If


            Else
                Return "0"

            End If
            Return "0"
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Function

    Private Function Identificadorincapacidad(identificador As String) As String
        Try
            Dim TipoIncidencia As String = ""

            If identificador = "0" Then
                TipoIncidencia = "Riesgo de trabajo"
            ElseIf identificador = "1" Then
                TipoIncidencia = "Enfermedad general"
            ElseIf identificador = "2" Then
                TipoIncidencia = "Maternidad"

            End If

            Return TipoIncidencia
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        
    End Function


    Private Function CalcularEdad(ByVal DiaNacimiento As Integer, ByVal MesNacimiento As Integer, ByVal AñoNacimiento As Integer)
        ' SE DEFINEN LAS FECHAS ACTUALES
        Dim AñoActual As Integer = Year(Now)
        Dim MesActual As Integer = Month(Now)
        Dim DiaActual As Integer = Now.Day
        Dim Cumplidos As Boolean = False
        ' SE COMPRUEBA CUANDO FUE EL ULTIMOS CUMPLEAÑOS
        ' FORMULA:
        '   Años cumplidos = (Año del ultimo cumpleaños - Año de nacimiento)
        If (MesNacimiento <= MesActual) Then
            If (DiaNacimiento <= DiaActual) Then
                If (DiaNacimiento = DiaActual And MesNacimiento = MesActual) Then
                    'MsgBox("Feliz Cumpleaños!")
                End If
                ' MsgBox("Ya cumplio")
                Cumplidos = True
            End If
        End If

        If (Cumplidos = False) Then
            AñoActual = (AñoActual - 1)
            'MsgBox("Ultimo cumpleaños: " & AñoActual)
        End If
        ' Se realiza la resta de años para definir los años cumplidos
        Dim EdadAños As Integer = (AñoActual - AñoNacimiento)
        ' DEFINICION DE LOS MESES LUEGO DEL ULTIMO CUMPLEAÑOS
        Dim EdadMes As Integer
        If Not (AñoActual = Now.Year) Then
            EdadMes = (12 - MesNacimiento)
            EdadMes = EdadMes + Now.Month
        Else
            EdadMes = Math.Abs(Now.Month - MesNacimiento)
        End If
        'SACAMOS LA CANTIDAD DE DIAS EXACTOS
        Dim EdadDia As Integer = (DiaActual - DiaNacimiento)

        'RETORNAMOS LOS VALORES EN UNA CADENA STRING
        Return (EdadAños)


    End Function


    Private Sub cmdguardarnomina_Click(sender As Object, e As EventArgs) Handles cmdguardarnomina.Click

        Try
            Dim sql As String
            Dim sql2 As String
            sql = "select * from Nomina where fkiIdEmpresa=1 and fkiIdPeriodo=" & cboperiodo.SelectedValue
            sql &= " and iEstatusNomina=1 and iEstatus=1 and iEstatusEmpleado=" & cboserie.SelectedIndex
            sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex
            'Dim sueldobase, salariodiario, salariointegrado, sueldobruto, TiempoExtraFijoGravado, TiempoExtraFijoExento As Double
            'Dim TiempoExtraOcasional, DesSemObligatorio, VacacionesProporcionales, AguinaldoGravado, AguinaldoExento As Double
            'Dim PrimaVacGravada, PrimaVacExenta, TotalPercepciones, TotalPercepcionesISR As Double
            'Dim incapacidad, ISR, IMSS, Infonavit, InfonavitAnterior, InfonavitAjuste, PensionAlimenticia As Double
            'Dim Prestamo, Fonacot, NetoaPagar, Excedente, Total, ImssCS, RCVCS, InfonavitCS, ISNCS
            'sql = "EXEC getNominaXEmpresaXPeriodo " & gIdEmpresa & "," & cboperiodo.SelectedValue & ",1"

            Dim rwNominaGuardadaFinal As DataRow() = nConsulta(sql)

            If rwNominaGuardadaFinal Is Nothing = False Then
                MessageBox.Show("La nomina ya esta marcada como final, no  se pueden guardar cambios", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else


                sql = "delete from Nomina"
                sql &= " where fkiIdEmpresa=1 and fkiIdPeriodo=" & cboperiodo.SelectedValue
                sql &= " and iEstatusNomina=0 and iEstatus=1 and iEstatusEmpleado=" & cboserie.SelectedIndex
                sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex
                If nExecute(sql) = False Then
                    MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    'pnlProgreso.Visible = False
                    Exit Sub
                End If

                sql = "delete from DetalleDescInfonavit"
                sql &= " where fkiIdPeriodo=" & cboperiodo.SelectedValue
                sql &= " and iSerie=" & cboserie.SelectedIndex
                'sql &= " and iSerie=" & cboserie.SelectedIndex
                sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex

                If nExecute(sql) = False Then
                    MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    'pnlProgreso.Visible = False
                    Exit Sub
                End If

                pnlProgreso.Visible = True

                Application.DoEvents()
                pnlCatalogo.Enabled = False
                pgbProgreso.Minimum = 0
                pgbProgreso.Value = 0
                pgbProgreso.Maximum = dtgDatos.Rows.Count

                For x As Integer = 0 To dtgDatos.Rows.Count - 1



                    sql = "EXEC [setNominaInsertar ] 0"
                    'periodo
                    sql &= "," & cboperiodo.SelectedValue
                    'idempleado
                    sql &= "," & dtgDatos.Rows(x).Cells(2).Value
                    'idempresa
                    sql &= ",1"
                    'Puesto
                    'buscamos el valor en la tabla
                    sql2 = "select * from puestos where cNombre='" & dtgDatos.Rows(x).Cells(11).FormattedValue & "'"

                    Dim rwPuesto As DataRow() = nConsulta(sql2)

                    sql &= "," & rwPuesto(0)("iIdPuesto")


                    'departamento
                    'buscamos el valor en la tabla
                    sql2 = "select * from departamentos where cNombre='" & dtgDatos.Rows(x).Cells(12).FormattedValue & "'"

                    Dim rwDepto As DataRow() = nConsulta(sql2)

                    sql &= "," & rwDepto(0)("iIdDepartamento")

                    'estatus empleado
                    sql &= "," & cboserie.SelectedIndex
                    'edad
                    sql &= "," & dtgDatos.Rows(x).Cells(10).Value
                    'puesto
                    sql &= ",'" & dtgDatos.Rows(x).Cells(11).FormattedValue & "'"
                    'buque
                    sql &= ",'" & dtgDatos.Rows(x).Cells(12).FormattedValue & "'"
                    'iTipo Infonavit
                    sql &= ",'" & dtgDatos.Rows(x).Cells(13).Value & "'"
                    'valor infonavit
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(14).Value = "", "0", dtgDatos.Rows(x).Cells(14).Value.ToString.Replace(",", ""))
                    'salario base
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(15).Value = "", "0", dtgDatos.Rows(x).Cells(15).Value.ToString.Replace(",", ""))
                    'salario diario
                    sql &= "," & dtgDatos.Rows(x).Cells(16).Value
                    'salario integrado
                    sql &= "," & dtgDatos.Rows(x).Cells(17).Value
                    'Dias trabajados
                    sql &= "," & dtgDatos.Rows(x).Cells(18).Value
                    'tipo incapacidad

                    sql &= ",'" & dtgDatos.Rows(x).Cells(19).Value & "'"
                    'numero dias incapacidad
                    sql &= "," & dtgDatos.Rows(x).Cells(20).Value
                    'sueldobruto
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(21).Value = "", "0", dtgDatos.Rows(x).Cells(21).Value.ToString.Replace(",", ""))
                    'tiempo extra fijo gravado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(22).Value = "", "0", dtgDatos.Rows(x).Cells(22).Value.ToString.Replace(",", ""))
                    'tiempo extra fijo exento
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(23).Value = "", "0", dtgDatos.Rows(x).Cells(23).Value.ToString.Replace(",", ""))
                    'Tiempo extra ocasional
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(24).Value = "", "0", dtgDatos.Rows(x).Cells(24).Value.ToString.Replace(",", ""))
                    'descanso semanal obligatorio
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(25).Value = "", "0", dtgDatos.Rows(x).Cells(25).Value.ToString.Replace(",", ""))
                    'vacaciones proporcionales
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(26).Value = "", "0", dtgDatos.Rows(x).Cells(26).Value.ToString.Replace(",", ""))
                    'aguinaldo gravado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(27).Value = "", "0", dtgDatos.Rows(x).Cells(27).Value.ToString.Replace(",", ""))
                    'aguinaldo exento
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(28).Value = "", "0", dtgDatos.Rows(x).Cells(28).Value.ToString.Replace(",", ""))
                    'prima vacacional gravado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(30).Value = "", "0", dtgDatos.Rows(x).Cells(30).Value.ToString.Replace(",", ""))
                    'prima vacacional exento
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(31).Value = "", "0", dtgDatos.Rows(x).Cells(31).Value.ToString.Replace(",", ""))

                    'totalpercepciones
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(33).Value = "", "0", dtgDatos.Rows(x).Cells(33).Value.ToString.Replace(",", ""))
                    'totalpercepcionesISR
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(34).Value = "", "0", dtgDatos.Rows(x).Cells(34).Value.ToString.Replace(",", ""))
                    'Incapacidad
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(35).Value = "", "0", dtgDatos.Rows(x).Cells(35).Value.ToString.Replace(",", ""))
                    'isr
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(36).Value = "", "0", dtgDatos.Rows(x).Cells(36).Value.ToString.Replace(",", ""))
                    'imss
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(37).Value = "", "0", dtgDatos.Rows(x).Cells(37).Value.ToString.Replace(",", ""))
                    'infonavit
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(38).Value = "", "0", dtgDatos.Rows(x).Cells(38).Value.ToString.Replace(",", ""))
                    'infonavit anterior
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(39).Value = "", "0", dtgDatos.Rows(x).Cells(39).Value.ToString.Replace(",", ""))
                    'ajuste infonavit
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(40).Value = "", "0", dtgDatos.Rows(x).Cells(40).Value.ToString.Replace(",", ""))
                    'Pension alimenticia
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(41).Value = "", "0", dtgDatos.Rows(x).Cells(41).Value.ToString.Replace(",", ""))
                    'Prestamo
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(42).Value = "", "0", dtgDatos.Rows(x).Cells(42).Value.ToString.Replace(",", ""))
                    'Fonacot
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(43).Value = "", "0", dtgDatos.Rows(x).Cells(43).Value.ToString.Replace(",", ""))
                    'Subsidio Generado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(44).Value = "", "0", dtgDatos.Rows(x).Cells(44).Value.ToString.Replace(",", ""))
                    'Subsidio Aplicado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(45).Value = "", "0", dtgDatos.Rows(x).Cells(45).Value.ToString.Replace(",", ""))
                    'Operadora
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(46).Value = "", "0", dtgDatos.Rows(x).Cells(46).Value.ToString.Replace(",", ""))
                    'Prestamo Personal Asimilado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(47).Value = "", "0", dtgDatos.Rows(x).Cells(47).Value.ToString.Replace(",", ""))
                    'Adeudo_Infonavit_Asimilado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(48).Value = "", "0", dtgDatos.Rows(x).Cells(48).Value.ToString.Replace(",", ""))
                    'Difencia infonavit Asimilado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(49).Value = "", "0", dtgDatos.Rows(x).Cells(49).Value.ToString.Replace(",", ""))
                    'Complemento Asimilado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(50).Value = "", "0", dtgDatos.Rows(x).Cells(50).Value.ToString.Replace(",", ""))
                    'Retenciones_Operadora
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(51).Value = "", "0", dtgDatos.Rows(x).Cells(51).Value.ToString.Replace(",", ""))
                    '% Comision
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(52).Value = "", "0", dtgDatos.Rows(x).Cells(52).Value.ToString.Replace(",", ""))
                    'Comision_Operadora
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(53).Value = "", "0", dtgDatos.Rows(x).Cells(53).Value.ToString.Replace(",", ""))
                    'Comision asimilados
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(54).Value = "", "0", dtgDatos.Rows(x).Cells(54).Value.ToString.Replace(",", ""))
                    'IMSS_CS
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(55).Value = "", "0", dtgDatos.Rows(x).Cells(55).Value.ToString.Replace(",", ""))
                    'RCV_CS
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(56).Value = "", "0", dtgDatos.Rows(x).Cells(56).Value.ToString.Replace(",", ""))
                    'Infonavit_CS
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(57).Value = "", "0", dtgDatos.Rows(x).Cells(57).Value.ToString.Replace(",", ""))
                    'ISN_CS
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(58).Value = "", "0", dtgDatos.Rows(x).Cells(58).Value.ToString.Replace(",", ""))
                    'Total Costo Social
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(59).Value = "", "0", dtgDatos.Rows(x).Cells(59).Value.ToString.Replace(",", ""))
                    'Subtotal
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(60).Value = "", "0", dtgDatos.Rows(x).Cells(60).Value.ToString.Replace(",", ""))
                    'IVA
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(61).Value = "", "0", dtgDatos.Rows(x).Cells(61).Value.ToString.Replace(",", ""))
                    'TOTAL DEPOSITO
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(62).Value = "", "0", dtgDatos.Rows(x).Cells(62).Value.ToString.Replace(",", ""))
                    'Estatus
                    sql &= ",1"
                    'Estatus Nomina
                    sql &= ",0"
                    'Tipo Nomina
                    sql &= "," & cboTipoNomina.SelectedIndex






                    If nExecute(sql) = False Then
                        MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        'pnlProgreso.Visible = False
                        Exit Sub
                    End If

                    '########GUARDAR INFONAVIT

                    If Double.Parse(dtgDatos.Rows(x).Cells(38).Value) Then

                        Dim MontoInfonavit As Double = MontoInfonavitF(cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value))

                        sql = "EXEC setDetalleDescInfonavitInsertar  0"
                        'fk Calculo infonavit
                        sql &= "," & IIf(MontoInfonavit > 0, IDCalculoInfonavit, 0)
                        'Cantidad
                        sql &= "," & dtgDatos.Rows(x).Cells(38).Value
                        ' fk Empleado
                        sql &= ",'" & dtgDatos.Rows(x).Cells(2).Value
                        'fk Periodo
                        sql &= "'," & cboperiodo.SelectedValue
                        'Serie
                        sql &= "," & cboserie.SelectedIndex
                        'Tipo Nomina
                        sql &= "," & cboTipoNomina.SelectedIndex
                        'iEstatu
                        sql &= ",1"

                        If nExecute(sql) = False Then
                            MessageBox.Show("Ocurrio un error insertar pago prestamo ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            'pnlProgreso.Visible = False
                            Exit Sub
                        End If
                    End If

                    




                    'sql = "update empleadosC set fSueldoOrd=" & dtgDatos.Rows(x).Cells(6).Value & ", fCosto =" & dtgDatos.Rows(x).Cells(18).Value
                    'sql &= " where iIdEmpleadoC = " & dtgDatos.Rows(x).Cells(2).Value

                    'If nExecute(sql) = False Then
                    '    MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    '    'pnlProgreso.Visible = False
                    '    Exit Sub
                    'End If

                    pgbProgreso.Value += 1
                    Application.DoEvents()
                Next
                pnlProgreso.Visible = False
                pnlCatalogo.Enabled = True

                If cboTipoNomina.SelectedIndex = 0 Then
                    MessageBox.Show("Datos guardados correctamente, se generara la nomina descanso", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    NominaB()
                    MessageBox.Show("Nomina Descanso generado, si no hay cambios proceda a guardar", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("Datos guardados correctamente", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
                

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub
    Private Sub NominaB()
        cboTipoNomina.SelectedIndex = 1
        For x As Integer = 0 To dtgDatos.Rows.Count - 1
            'Dim cadena As String = dgvCombo.Text
            If dtgDatos.Rows(x).Cells(11).FormattedValue = "OFICIALES EN PRACTICAS: PILOTIN / ASPIRANTE" Then
                dtgDatos.Rows(x).Cells(15).Value = "0.00"
                dtgDatos.Rows(x).Cells(18).Value = "0.00"
                dtgDatos.Rows(x).Cells(21).Value = "0.00"
                dtgDatos.Rows(x).Cells(22).Value = "0.00"
                dtgDatos.Rows(x).Cells(23).Value = "0.00"
                dtgDatos.Rows(x).Cells(24).Value = "0.00"
                dtgDatos.Rows(x).Cells(25).Value = "0.00"
                dtgDatos.Rows(x).Cells(26).Value = "0.00"
                dtgDatos.Rows(x).Cells(27).Value = "0.00"
                dtgDatos.Rows(x).Cells(28).Value = "0.00"
                dtgDatos.Rows(x).Cells(29).Value = "0.00"
                dtgDatos.Rows(x).Cells(30).Value = "0.00"
                dtgDatos.Rows(x).Cells(31).Value = "0.00"
                dtgDatos.Rows(x).Cells(32).Value = "0.00"
                dtgDatos.Rows(x).Cells(33).Value = "0.00"
                dtgDatos.Rows(x).Cells(34).Value = "0.00"
                dtgDatos.Rows(x).Cells(35).Value = "0.00"
                'ISR
                dtgDatos.Rows(x).Cells(36).Value = "0.00"
                'IMSS
                dtgDatos.Rows(x).Cells(37).Value = "0.00"
                'INFONAVIT
                dtgDatos.Rows(x).Cells(38).Value = "0.00"
                'INFONAVIT BIMESTRE ANTERIOR
                dtgDatos.Rows(x).Cells(38).Value = "0.00"
                'AJUSTE INFONAVIT
                dtgDatos.Rows(x).Cells(38).Value = "0.00"
                'PENSION
                dtgDatos.Rows(x).Cells(38).Value = "0.00"
                'PRESTAMO
                dtgDatos.Rows(x).Cells(38).Value = "0.00"
                'FONACOT
                dtgDatos.Rows(x).Cells(38).Value = "0.00"
                'SUBSIDIO GENERADO
                dtgDatos.Rows(x).Cells(44).Value = "0.00"
                'SUBSIDIO APLICADO
                dtgDatos.Rows(x).Cells(45).Value = "0.00"
                'NETO
                dtgDatos.Rows(x).Cells(46).Value = "0.00"
                'Prestamo Personal Asimilado
                dtgDatos.Rows(x).Cells(47).Value = "0.00"
                'Adeudo_Infonavit_Asimilado
                dtgDatos.Rows(x).Cells(48).Value = "0.00"
                'Difencia infonavit Asimilado
                dtgDatos.Rows(x).Cells(49).Value = "0.00"
                'Complemento Asimilado
                dtgDatos.Rows(x).Cells(50).Value = "0.00"
                'Retenciones_Operadora
                dtgDatos.Rows(x).Cells(51).Value = "0.00"
                '% Comision
                dtgDatos.Rows(x).Cells(52).Value = "0.00"
                'Comision_Operadora
                dtgDatos.Rows(x).Cells(53).Value = "0.00"
                'Comision asimilados
                dtgDatos.Rows(x).Cells(54).Value = "0.00"
                'IMSS_CS
                dtgDatos.Rows(x).Cells(55).Value = "0.00"
                'RCV_CS
                dtgDatos.Rows(x).Cells(56).Value = "0.00"
                'Infonavit_CS
                dtgDatos.Rows(x).Cells(57).Value = "0.00"
                'ISN_CS
                dtgDatos.Rows(x).Cells(58).Value = "0.00"
                'Total Costo Social
                dtgDatos.Rows(x).Cells(59).Value = "0.00"
                'Subtotal
                dtgDatos.Rows(x).Cells(60).Value = "0.00"
                'IVA
                dtgDatos.Rows(x).Cells(61).Value = "0.00"
                'TOTAL DEPOSITO
                dtgDatos.Rows(x).Cells(62).Value = "0.00"

            Else
                dtgDatos.Rows(x).Cells(15).Value = "0.00"

            End If



            
        Next


    End Sub
    Private Sub cmdcalcular_Click(sender As Object, e As EventArgs) Handles cmdcalcular.Click
        Try
            calcular()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

    Private Sub calcular()
        Dim Sueldo As Double
        Dim SueldoBase As Double
        Dim ValorIncapacidad As Double
        Dim TotalPercepciones As Double
        Dim Incapacidad As Double
        Dim isr As Double
        Dim imss As Double
        Dim infonavitvalor As Double
        Dim infonavitanterior As Double
        Dim ajusteinfonavit As Double
        Dim pension As Double
        Dim prestamo As Double
        Dim fonacot As Double
        Dim subsidiogenerado As Double
        Dim subsidioaplicado As Double
        Dim sql As String
        Dim ValorUMA As Double
        Dim primavacacionesgravada As Double
        Dim primavacacionesexenta As Double
        Dim diastrabajados As Double
        Dim Sueldobruto As Double
        Dim TEFG As Double
        Dim TEFE As Double
        Dim TEO As Double
        Dim DSO As Double
        Dim VACAPRO As Double

        Try
            'verificamos que tenga dias a calcular
            For x As Integer = 0 To dtgDatos.Rows.Count - 1
                If Double.Parse(IIf(dtgDatos.Rows(x).Cells(18).Value = "", "0", dtgDatos.Rows(x).Cells(18).Value)) <= 0 Then
                    MessageBox.Show("Existen trabajadores que no tiene dias trabajados, favor de verificar", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
            Next

            

            sql = "select * from Salario "
            sql &= " where Anio=" & aniocostosocial
            sql &= " and iEstatus=1"
            Dim rwValorUMA As DataRow() = nConsulta(sql)
            If rwValorUMA Is Nothing = False Then
                ValorUMA = Double.Parse(rwValorUMA(0)("uma").ToString)
            Else
                ValorUMA = 0
                MessageBox.Show("No se encontro valor para UMA en el año: " & aniocostosocial)
            End If


            pnlProgreso.Visible = True

            Application.DoEvents()
            pnlCatalogo.Enabled = False
            pgbProgreso.Minimum = 0
            pgbProgreso.Value = 0
            pgbProgreso.Maximum = dtgDatos.Rows.Count


            For x As Integer = 0 To dtgDatos.Rows.Count - 1
                'Dim cadena As String = dgvCombo.Text
                If dtgDatos.Rows(x).Cells(11).FormattedValue = "OFICIALES EN PRACTICAS: PILOTIN / ASPIRANTE" Then
                    Sueldo = Double.Parse(dtgDatos.Rows(x).Cells(17).Value) * Double.Parse(IIf(dtgDatos.Rows(x).Cells(18).Value = "", "0", dtgDatos.Rows(x).Cells(18).Value))
                    dtgDatos.Rows(x).Cells(21).Value = Math.Round(Sueldo, 2).ToString("###,##0.00")
                    dtgDatos.Rows(x).Cells(22).Value = "0.00"
                    dtgDatos.Rows(x).Cells(23).Value = "0.00"
                    dtgDatos.Rows(x).Cells(24).Value = "0.00"
                    dtgDatos.Rows(x).Cells(25).Value = "0.00"
                    dtgDatos.Rows(x).Cells(26).Value = "0.00"
                    dtgDatos.Rows(x).Cells(27).Value = "0.00"
                    dtgDatos.Rows(x).Cells(28).Value = "0.00"
                    dtgDatos.Rows(x).Cells(29).Value = "0.00"
                    dtgDatos.Rows(x).Cells(30).Value = "0.00"
                    dtgDatos.Rows(x).Cells(31).Value = "0.00"
                    dtgDatos.Rows(x).Cells(32).Value = "0.00"
                    dtgDatos.Rows(x).Cells(33).Value = Math.Round(Sueldo, 2).ToString("###,##0.00")
                    dtgDatos.Rows(x).Cells(34).Value = Math.Round(Sueldo, 2).ToString("###,##0.00")
                    'Incapacidad
                    ValorIncapacidad = 0.0
                    If dtgDatos.Rows(x).Cells(19).Value <> "Ninguno" Then

                        ValorIncapacidad = Incapacidades(dtgDatos.Rows(x).Cells(19).Value, dtgDatos.Rows(x).Cells(20).Value, dtgDatos.Rows(x).Cells(16).Value)

                    End If
                    dtgDatos.Rows(x).Cells(35).Value = Math.Round(ValorIncapacidad, 2).ToString("###,##0.00")
                    'ISR
                    dtgDatos.Rows(x).Cells(36).Value = Math.Round(Double.Parse((baseisrtotal(dtgDatos.Rows(x).Cells(11).Value, 30, dtgDatos.Rows(x).Cells(17).Value, ValorIncapacidad)) / 30 * dtgDatos.Rows(x).Cells(18).Value), 2).ToString("###,##0.00")
                    'IMSS
                    dtgDatos.Rows(x).Cells(37).Value = "0.00"
                    'INFONAVIT
                    '##### VERIFICAR SI ESTA YA CALCULADO EL INFONAVIT DEL BIMESTRE

                    Dim CalculoInfonavit As Integer = VerificarCalculoInfonavit(cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value))

                    Select Case CalculoInfonavit
                        Case 0
                            'No es necesario calcular
                            dtgDatos.Rows(x).Cells(38).Value = "0.00"
                        Case 1
                            'Ya esta Calculado
                            'Verificar cuanto le toca para el pago
                            Dim MontoInfonavit As Double = MontoInfonavitF(cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value))

                            If MontoInfonavit > 0 Then

                                sql = "select isnull(sum(Cantidad),0) as monto from DetalleDescInfonavit where fkiIdCalculoInfonavit=" & IDCalculoInfonavit
                                Dim rwMontoInfonavit As DataRow() = nConsulta(sql)
                                If rwMontoInfonavit Is Nothing = False Then

                                    If Double.Parse(rwMontoInfonavit(0)("monto").ToString) < MontoInfonavit Then
                                        'Diferencia
                                        Dim FaltanteInfonavit As Double = MontoInfonavit - Double.Parse(rwMontoInfonavit(0)("monto").ToString)

                                        TotalPercepciones = Double.Parse(IIf(dtgDatos.Rows(x).Cells(33).Value = "", "0", dtgDatos.Rows(x).Cells(33).Value.ToString.Replace(",", "")))
                                        Incapacidad = Double.Parse(IIf(dtgDatos.Rows(x).Cells(35).Value = "", "0", dtgDatos.Rows(x).Cells(35).Value))
                                        isr = Double.Parse(IIf(dtgDatos.Rows(x).Cells(36).Value = "", "0", dtgDatos.Rows(x).Cells(36).Value))
                                        imss = Double.Parse(IIf(dtgDatos.Rows(x).Cells(37).Value = "", "0", dtgDatos.Rows(x).Cells(37).Value))

                                        Dim SubtotalAntesInfonavit As Double = TotalPercepciones - Incapacidad - isr - imss

                                        If SubtotalAntesInfonavit > (FaltanteInfonavit / 2) Then
                                            dtgDatos.Rows(x).Cells(38).Value = Math.Round((FaltanteInfonavit / 2), 2)

                                        Else
                                            dtgDatos.Rows(x).Cells(38).Value = Math.Round((SubtotalAntesInfonavit - 1), 2)
                                        End If



                                    Else
                                        dtgDatos.Rows(x).Cells(38).Value = "0.00"
                                    End If


                                End If
                            Else
                                dtgDatos.Rows(x).Cells(38).Value = "0.00"

                            End If
                        Case 2
                            'No esta calculado
                            If CalcularInfonavit(dtgDatos.Rows(x).Cells(13).Value, Double.Parse(dtgDatos.Rows(x).Cells(14).Value), Double.Parse(dtgDatos.Rows(x).Cells(17).Value), Date.Parse("01/01/1900"), cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value)) Then
                                'Verificar cuanto le toca para el pago
                                Dim MontoInfonavit As Double = MontoInfonavitF(cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value))

                                If MontoInfonavit > 0 Then
                                    sql = "select isnull(sum(Cantidad),0) as monto from DetalleDescInfonavit where fkiIdCalculoInfonavit=" & IDCalculoInfonavit
                                    Dim rwMontoInfonavit As DataRow() = nConsulta(sql)
                                    If rwMontoInfonavit Is Nothing = False Then

                                        If Double.Parse(rwMontoInfonavit(0)("monto").ToString) < MontoInfonavit Then
                                            'Diferencia
                                            Dim FaltanteInfonavit As Double = MontoInfonavit - Double.Parse(rwMontoInfonavit(0)("monto").ToString)

                                            TotalPercepciones = Double.Parse(IIf(dtgDatos.Rows(x).Cells(33).Value = "", "0", dtgDatos.Rows(x).Cells(33).Value.ToString.Replace(",", "")))
                                            Incapacidad = Double.Parse(IIf(dtgDatos.Rows(x).Cells(35).Value = "", "0", dtgDatos.Rows(x).Cells(35).Value))
                                            isr = Double.Parse(IIf(dtgDatos.Rows(x).Cells(36).Value = "", "0", dtgDatos.Rows(x).Cells(36).Value))
                                            imss = Double.Parse(IIf(dtgDatos.Rows(x).Cells(37).Value = "", "0", dtgDatos.Rows(x).Cells(37).Value))

                                            Dim SubtotalAntesInfonavit As Double = TotalPercepciones - Incapacidad - isr - imss

                                            If SubtotalAntesInfonavit > (FaltanteInfonavit / 2) Then
                                                dtgDatos.Rows(x).Cells(38).Value = Math.Round((FaltanteInfonavit / 2), 2)

                                            Else
                                                dtgDatos.Rows(x).Cells(38).Value = Math.Round((SubtotalAntesInfonavit - 1), 2)
                                            End If



                                        Else
                                            dtgDatos.Rows(x).Cells(38).Value = "0.00"
                                        End If


                                    End If
                                Else
                                    dtgDatos.Rows(x).Cells(38).Value = "0.00"

                                End If


                            End If
                    End Select

                    '############# CALCULO POR DIAS INFONAVIT

                    'dtgDatos.Rows(x).Cells(38).Value = Math.Round(infonavit(dtgDatos.Rows(x).Cells(13).Value, Double.Parse(dtgDatos.Rows(x).Cells(14).Value), Double.Parse(dtgDatos.Rows(x).Cells(17).Value), Date.Parse("01/01/1900"), cboperiodo.SelectedValue, Double.Parse(dtgDatos.Rows(x).Cells(18).Value), Integer.Parse(dtgDatos.Rows(x).Cells(2).Value)), 2).ToString("###,##0.00")
                    '#############

                    'INFONAVIT BIMESTRE ANTERIOR
                    'AJUSTE INFONAVIT
                    'PENSION
                    'PRESTAMO
                    'FONACOT
                    'SUBSIDIO GENERADO
                    dtgDatos.Rows(x).Cells(44).Value = baseSubsidio(dtgDatos.Rows(x).Cells(11).FormattedValue, 30, Double.Parse(dtgDatos.Rows(x).Cells(17).Value), ValorIncapacidad)
                    'SUBSIDIO APLICADO
                    dtgDatos.Rows(x).Cells(45).Value = Math.Round(baseSubsidiototal(dtgDatos.Rows(x).Cells(11).FormattedValue, 30, Double.Parse(dtgDatos.Rows(x).Cells(17).Value), ValorIncapacidad) / 30 * Double.Parse(dtgDatos.Rows(x).Cells(18).Value), 2)
                    'NETO


                    TotalPercepciones = Double.Parse(IIf(dtgDatos.Rows(x).Cells(33).Value = "", "0", dtgDatos.Rows(x).Cells(33).Value.ToString.Replace(",", "")))
                    Incapacidad = Double.Parse(IIf(dtgDatos.Rows(x).Cells(35).Value = "", "0", dtgDatos.Rows(x).Cells(35).Value))
                    isr = Double.Parse(IIf(dtgDatos.Rows(x).Cells(36).Value = "", "0", dtgDatos.Rows(x).Cells(36).Value))
                    imss = Double.Parse(IIf(dtgDatos.Rows(x).Cells(37).Value = "", "0", dtgDatos.Rows(x).Cells(37).Value))
                    infonavitvalor = Double.Parse(IIf(dtgDatos.Rows(x).Cells(38).Value = "", "0", dtgDatos.Rows(x).Cells(38).Value))
                    infonavitanterior = Double.Parse(IIf(dtgDatos.Rows(x).Cells(39).Value = "", "0", dtgDatos.Rows(x).Cells(39).Value))
                    ajusteinfonavit = Double.Parse(IIf(dtgDatos.Rows(x).Cells(40).Value = "", "0", dtgDatos.Rows(x).Cells(40).Value))
                    pension = Double.Parse(IIf(dtgDatos.Rows(x).Cells(41).Value = "", "0", dtgDatos.Rows(x).Cells(41).Value))
                    prestamo = Double.Parse(IIf(dtgDatos.Rows(x).Cells(42).Value = "", "0", dtgDatos.Rows(x).Cells(42).Value))
                    fonacot = Double.Parse(IIf(dtgDatos.Rows(x).Cells(43).Value = "", "0", dtgDatos.Rows(x).Cells(43).Value))
                    subsidiogenerado = Double.Parse(IIf(dtgDatos.Rows(x).Cells(44).Value = "", "0", dtgDatos.Rows(x).Cells(44).Value))
                    subsidioaplicado = Double.Parse(IIf(dtgDatos.Rows(x).Cells(45).Value = "", "0", dtgDatos.Rows(x).Cells(45).Value))
                    dtgDatos.Rows(x).Cells(46).Value = Math.Round(TotalPercepciones - Incapacidad - isr - imss - infonavitvalor - infonavitanterior - ajusteinfonavit - pension - prestamo - fonacot + subsidioaplicado, 2)


                    'Prestamo Personal Asimilado
                    'Adeudo_Infonavit_Asimilado
                    'Difencia infonavit Asimilado
                    'Complemento Asimilado
                    'Retenciones_Operadora
                    '% Comision
                    'Comision_Operadora
                    'Comision asimilados


                Else
                    diastrabajados = Double.Parse(IIf(dtgDatos.Rows(x).Cells(18).Value = "", "0", dtgDatos.Rows(x).Cells(18).Value))

                    Sueldo = Double.Parse(dtgDatos.Rows(x).Cells(17).Value) * diastrabajados
                    dtgDatos.Rows(x).Cells(21).Value = Math.Round(Sueldo * (26.19568006 / 100), 2).ToString("###,##0.00")
                    Sueldobruto = Math.Round(Sueldo * (26.19568006 / 100), 2)
                    dtgDatos.Rows(x).Cells(22).Value = Math.Round((Sueldo * (8.5070471 / 100)) / 2, 2).ToString("###,##0.00")
                    TEFG = Math.Round((Sueldo * (8.5070471 / 100)) / 2, 2)
                    dtgDatos.Rows(x).Cells(23).Value = Math.Round((Sueldo * (8.5070471 / 100)) / 2, 2).ToString("###,##0.00")
                    TEFE = Math.Round((Sueldo * (8.5070471 / 100)) / 2, 2)
                    dtgDatos.Rows(x).Cells(24).Value = Math.Round(Sueldo * (42.89215164 / 100), 2).ToString("###,##0.00")
                    TEO = Math.Round(Sueldo * (42.89215164 / 100), 2)
                    dtgDatos.Rows(x).Cells(25).Value = Math.Round(Sueldo * (9.677848468 / 100), 2).ToString("###,##0.00")
                    DSO = Math.Round(Sueldo * (9.677848468 / 100), 2)
                    dtgDatos.Rows(x).Cells(26).Value = Math.Round(Sueldo * (7.272727273 / 100), 2).ToString("###,##0.00")
                    VACAPRO = Math.Round(Sueldo * (7.272727273 / 100), 2)
                    SueldoBase = Sueldobruto + TEFG + TEFE + TEO + DSO


                    'Aguinaldo gravado 

                    If ((SueldoBase / diastrabajados) * 15 / 12 * (diastrabajados / 30)) > ((ValorUMA * 30 / 12) * (diastrabajados / 30)) Then
                        'Aguinaldo gravado
                        dtgDatos.Rows(x).Cells(27).Value = Math.Round(((SueldoBase / diastrabajados) * 15 / 12 * (diastrabajados / 30)) - ((ValorUMA * 30 / 12) * (diastrabajados / 30)), 2)
                        'Aguinaldo exento
                        dtgDatos.Rows(x).Cells(28).Value = Math.Round(((ValorUMA * 30 / 12) * (diastrabajados / 30)), 2)


                    Else
                        'Aguinaldo gravado

                        dtgDatos.Rows(x).Cells(27).Value = "0.00"
                        'Aguinaldo exento
                        dtgDatos.Rows(x).Cells(28).Value = Math.Round(((SueldoBase / diastrabajados) * 15 / 12 * (diastrabajados / 30)), 2)

                    End If


                    'Aguinaldo total
                    dtgDatos.Rows(x).Cells(29).Value = Math.Round(Double.Parse(dtgDatos.Rows(x).Cells(27).Value) + Double.Parse(dtgDatos.Rows(x).Cells(28).Value), 2)

                    'Prima de vacaciones

                    'Calculos prima

                    primavacacionesgravada = (SueldoBase * 0.25 / 12 * (diastrabajados / 30)) - ((ValorUMA * 15 / 12) * (diastrabajados / 30))
                    primavacacionesexenta = ((ValorUMA * 15 / 12) * (diastrabajados / 30))

                    If primavacacionesgravada > 0 Then
                        dtgDatos.Rows(x).Cells(30).Value = Math.Round(primavacacionesgravada, 2)
                        dtgDatos.Rows(x).Cells(31).Value = Math.Round(primavacacionesexenta, 2)
                    Else
                        primavacacionesexenta = (SueldoBase * 0.25 / 12 * (diastrabajados / 30))
                        dtgDatos.Rows(x).Cells(30).Value = "0.00"
                        dtgDatos.Rows(x).Cells(31).Value = Math.Round(primavacacionesexenta, 2)
                    End If

                    'Total Prima de vacaciones                    
                    dtgDatos.Rows(x).Cells(32).Value = Math.Round(IIf(primavacacionesgravada > 0, primavacacionesgravada, 0) + primavacacionesexenta, 2)
                    'Total percepciones
                    dtgDatos.Rows(x).Cells(33).Value = Math.Round(SueldoBase + VACAPRO + dtgDatos.Rows(x).Cells(29).Value + dtgDatos.Rows(x).Cells(32).Value, 2)
                    'Total percepsiones para isr
                    dtgDatos.Rows(x).Cells(34).Value = Math.Round(SueldoBase - TEFE + VACAPRO + dtgDatos.Rows(x).Cells(27).Value + dtgDatos.Rows(x).Cells(30).Value, 2)
                    'Incapacidad
                    ValorIncapacidad = 0.0
                    If dtgDatos.Rows(x).Cells(19).Value <> "Ninguno" Then

                        ValorIncapacidad = Incapacidades(dtgDatos.Rows(x).Cells(19).Value, dtgDatos.Rows(x).Cells(20).Value, dtgDatos.Rows(x).Cells(16).Value)

                    End If
                    dtgDatos.Rows(x).Cells(35).Value = Math.Round(ValorIncapacidad, 2).ToString("###,##0.00")
                    'ISR
                    dtgDatos.Rows(x).Cells(36).Value = Math.Round(Double.Parse((baseisrtotal(dtgDatos.Rows(x).Cells(11).FormattedValue, 30, dtgDatos.Rows(x).Cells(17).Value, ValorIncapacidad)) / 30 * dtgDatos.Rows(x).Cells(18).Value), 2).ToString("###,##0.00")
                    'IMSS
                    dtgDatos.Rows(x).Cells(37).Value = "0.00"
                    'INFONAVIT
                    '##### VERIFICAR SI ESTA YA CALCULADO EL INFONAVIT DEL BIMESTRE

                    Dim CalculoInfonavit As Integer = VerificarCalculoInfonavit(cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value))

                    Select Case CalculoInfonavit
                        Case 0
                            'No es necesario calcular
                            dtgDatos.Rows(x).Cells(38).Value = "0.00"
                        Case 1
                            'Ya esta Calculado
                            'Verificar cuanto le toca para el pago
                            Dim MontoInfonavit As Double = MontoInfonavitF(cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value))

                            If MontoInfonavit > 0 Then

                                sql = "select isnull(sum(Cantidad),0) as monto from DetalleDescInfonavit where fkiIdCalculoInfonavit=" & IDCalculoInfonavit
                                Dim rwMontoInfonavit As DataRow() = nConsulta(sql)
                                If rwMontoInfonavit Is Nothing = False Then

                                    If Double.Parse(rwMontoInfonavit(0)("monto").ToString) < MontoInfonavit Then
                                        'Diferencia
                                        Dim FaltanteInfonavit As Double = MontoInfonavit - Double.Parse(rwMontoInfonavit(0)("monto").ToString)

                                        TotalPercepciones = Double.Parse(IIf(dtgDatos.Rows(x).Cells(33).Value = "", "0", dtgDatos.Rows(x).Cells(33).Value.ToString.Replace(",", "")))
                                        Incapacidad = Double.Parse(IIf(dtgDatos.Rows(x).Cells(35).Value = "", "0", dtgDatos.Rows(x).Cells(35).Value))
                                        isr = Double.Parse(IIf(dtgDatos.Rows(x).Cells(36).Value = "", "0", dtgDatos.Rows(x).Cells(36).Value))
                                        imss = Double.Parse(IIf(dtgDatos.Rows(x).Cells(37).Value = "", "0", dtgDatos.Rows(x).Cells(37).Value))

                                        Dim SubtotalAntesInfonavit As Double = TotalPercepciones - Incapacidad - isr - imss

                                        If SubtotalAntesInfonavit > (FaltanteInfonavit / 2) Then
                                            dtgDatos.Rows(x).Cells(38).Value = Math.Round((FaltanteInfonavit / 2), 2)

                                        Else
                                            dtgDatos.Rows(x).Cells(38).Value = Math.Round((SubtotalAntesInfonavit - 1), 2)
                                        End If



                                    Else
                                        dtgDatos.Rows(x).Cells(38).Value = "0.00"
                                    End If


                                End If
                            Else
                                dtgDatos.Rows(x).Cells(38).Value = "0.00"

                            End If
                        Case 2
                            'No esta calculado
                            If CalcularInfonavit(dtgDatos.Rows(x).Cells(13).Value, Double.Parse(dtgDatos.Rows(x).Cells(14).Value), Double.Parse(dtgDatos.Rows(x).Cells(17).Value), Date.Parse("01/01/1900"), cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value)) Then
                                'Verificar cuanto le toca para el pago
                                Dim MontoInfonavit As Double = MontoInfonavitF(cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value))

                                If MontoInfonavit > 0 Then
                                    sql = "select isnull(sum(Cantidad),0) as monto from DetalleDescInfonavit where fkiIdCalculoInfonavit=" & IDCalculoInfonavit
                                    Dim rwMontoInfonavit As DataRow() = nConsulta(sql)
                                    If rwMontoInfonavit Is Nothing = False Then

                                        If Double.Parse(rwMontoInfonavit(0)("monto").ToString) < MontoInfonavit Then
                                            'Diferencia
                                            Dim FaltanteInfonavit As Double = MontoInfonavit - Double.Parse(rwMontoInfonavit(0)("monto").ToString)

                                            TotalPercepciones = Double.Parse(IIf(dtgDatos.Rows(x).Cells(33).Value = "", "0", dtgDatos.Rows(x).Cells(33).Value.ToString.Replace(",", "")))
                                            Incapacidad = Double.Parse(IIf(dtgDatos.Rows(x).Cells(35).Value = "", "0", dtgDatos.Rows(x).Cells(35).Value))
                                            isr = Double.Parse(IIf(dtgDatos.Rows(x).Cells(36).Value = "", "0", dtgDatos.Rows(x).Cells(36).Value))
                                            imss = Double.Parse(IIf(dtgDatos.Rows(x).Cells(37).Value = "", "0", dtgDatos.Rows(x).Cells(37).Value))

                                            Dim SubtotalAntesInfonavit As Double = TotalPercepciones - Incapacidad - isr - imss

                                            If SubtotalAntesInfonavit > (FaltanteInfonavit / 2) Then
                                                dtgDatos.Rows(x).Cells(38).Value = Math.Round((FaltanteInfonavit / 2), 2)

                                            Else
                                                dtgDatos.Rows(x).Cells(38).Value = Math.Round((SubtotalAntesInfonavit - 1), 2)
                                            End If



                                        Else
                                            dtgDatos.Rows(x).Cells(38).Value = "0.00"
                                        End If


                                    End If
                                Else
                                    dtgDatos.Rows(x).Cells(38).Value = "0.00"

                                End If


                            End If
                    End Select


                    '############# CALCULO POR DIAS INFONAVIT
                    'dtgDatos.Rows(x).Cells(38).Value = Math.Round(infonavit(dtgDatos.Rows(x).Cells(13).Value, Double.Parse(dtgDatos.Rows(x).Cells(14).Value), Double.Parse(dtgDatos.Rows(x).Cells(17).Value), Date.Parse("01/01/1900"), cboperiodo.SelectedValue, Double.Parse(dtgDatos.Rows(x).Cells(18).Value), Integer.Parse(dtgDatos.Rows(x).Cells(2).Value)), 2).ToString("###,##0.00")
                    '############# CALCULO POR DIAS INFONAVIT

                    'INFONAVIT BIMESTRE ANTERIOR
                    'AJUSTE INFONAVIT
                    'PENSION
                    'PRESTAMO
                    'FONACOT
                    'SUBSIDIO GENERADO
                    dtgDatos.Rows(x).Cells(44).Value = Math.Round((baseSubsidio(dtgDatos.Rows(x).Cells(11).FormattedValue, 30, Double.Parse(dtgDatos.Rows(x).Cells(17).Value), ValorIncapacidad)), 2).ToString("###,##0.00")
                    'SUBSIDIO APLICADO
                    dtgDatos.Rows(x).Cells(45).Value = Math.Round((baseSubsidiototal(dtgDatos.Rows(x).Cells(11).FormattedValue, 30, Double.Parse(dtgDatos.Rows(x).Cells(17).Value), ValorIncapacidad)) / 30 * Double.Parse(dtgDatos.Rows(x).Cells(18).Value), 2).ToString("###,##0.00")
                    'NETO


                    TotalPercepciones = Double.Parse(IIf(dtgDatos.Rows(x).Cells(33).Value = "", "0", dtgDatos.Rows(x).Cells(33).Value.ToString.Replace(",", "")))
                    Incapacidad = Double.Parse(IIf(dtgDatos.Rows(x).Cells(35).Value = "", "0", dtgDatos.Rows(x).Cells(35).Value))
                    isr = Double.Parse(IIf(dtgDatos.Rows(x).Cells(36).Value = "", "0", dtgDatos.Rows(x).Cells(36).Value))
                    imss = Double.Parse(IIf(dtgDatos.Rows(x).Cells(37).Value = "", "0", dtgDatos.Rows(x).Cells(37).Value))
                    infonavitvalor = Double.Parse(IIf(dtgDatos.Rows(x).Cells(38).Value = "", "0", dtgDatos.Rows(x).Cells(38).Value))
                    infonavitanterior = Double.Parse(IIf(dtgDatos.Rows(x).Cells(39).Value = "", "0", dtgDatos.Rows(x).Cells(39).Value))
                    ajusteinfonavit = Double.Parse(IIf(dtgDatos.Rows(x).Cells(40).Value = "", "0", dtgDatos.Rows(x).Cells(40).Value))
                    pension = Double.Parse(IIf(dtgDatos.Rows(x).Cells(41).Value = "", "0", dtgDatos.Rows(x).Cells(41).Value))
                    prestamo = Double.Parse(IIf(dtgDatos.Rows(x).Cells(42).Value = "", "0", dtgDatos.Rows(x).Cells(42).Value))
                    fonacot = Double.Parse(IIf(dtgDatos.Rows(x).Cells(43).Value = "", "0", dtgDatos.Rows(x).Cells(43).Value))
                    subsidiogenerado = Double.Parse(IIf(dtgDatos.Rows(x).Cells(44).Value = "", "0", dtgDatos.Rows(x).Cells(44).Value))
                    subsidioaplicado = Double.Parse(IIf(dtgDatos.Rows(x).Cells(45).Value = "", "0", dtgDatos.Rows(x).Cells(45).Value))
                    dtgDatos.Rows(x).Cells(46).Value = Math.Round(TotalPercepciones - Incapacidad - isr - imss - infonavitvalor - infonavitanterior - ajusteinfonavit - pension - prestamo - fonacot + subsidioaplicado, 2)


                    'Prestamo Personal Asimilado
                    'Adeudo_Infonavit_Asimilado
                    'Difencia infonavit Asimilado
                    'Complemento Asimilado
                    'Retenciones_Operadora
                    '% Comision
                    'Comision_Operadora
                    'Comision asimilados



                End If


                'Calcular el costo social

                'Obtenemos los datos del empleado,id puesto
                'de acuerdo a la edad y el status


                sql = "select * from empleadosC where iIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                Dim rwEmpleado As DataRow() = nConsulta(sql)
                If rwEmpleado Is Nothing = False Then
                    sql = "select * from costosocial where fkiIdPuesto=" & rwEmpleado(0)("fkiIdPuesto").ToString & " and anio=" & aniocostosocial
                    Dim rwCostoSocial As DataRow() = nConsulta(sql)
                    If rwCostoSocial Is Nothing = False Then
                        If dtgDatos.Rows(x).Cells(10).Value >= 55 Then
                            If dtgDatos.Rows(x).Cells(5).Value = "PLANTA" Then
                                dtgDatos.Rows(x).Cells(55).Value = rwCostoSocial(0)("imsstopado")
                                dtgDatos.Rows(x).Cells(56).Value = rwCostoSocial(0)("RCVtopado")
                                dtgDatos.Rows(x).Cells(57).Value = rwCostoSocial(0)("infonavittopado")
                                dtgDatos.Rows(x).Cells(58).Value = rwCostoSocial(0)("ISNtopado")
                                dtgDatos.Rows(x).Cells(59).Value = Math.Round(Double.Parse(dtgDatos.Rows(x).Cells(55).Value) + Double.Parse(dtgDatos.Rows(x).Cells(56).Value) + Double.Parse(dtgDatos.Rows(x).Cells(57).Value) + Double.Parse(dtgDatos.Rows(x).Cells(58).Value), 2)
                            Else
                                dtgDatos.Rows(x).Cells(55).Value = Math.Round(Double.Parse(rwCostoSocial(0)("imsstopado")) / 30 * dtgDatos.Rows(x).Cells(18).Value, 2)
                                dtgDatos.Rows(x).Cells(56).Value = Math.Round(Double.Parse(rwCostoSocial(0)("RCVtopado")) / 30 * dtgDatos.Rows(x).Cells(18).Value, 2)
                                dtgDatos.Rows(x).Cells(57).Value = Math.Round(Double.Parse(rwCostoSocial(0)("infonavittopado")) / 30 * dtgDatos.Rows(x).Cells(18).Value, 2)
                                dtgDatos.Rows(x).Cells(58).Value = Math.Round(Double.Parse(rwCostoSocial(0)("ISNtopado")) / 30 * dtgDatos.Rows(x).Cells(18).Value, 2)
                                dtgDatos.Rows(x).Cells(59).Value = Math.Round(Double.Parse(dtgDatos.Rows(x).Cells(55).Value) + Double.Parse(dtgDatos.Rows(x).Cells(56).Value) + Double.Parse(dtgDatos.Rows(x).Cells(57).Value) + Double.Parse(dtgDatos.Rows(x).Cells(58).Value), 2)
                            End If

                        Else
                            If dtgDatos.Rows(x).Cells(5).Value = "PLANTA" Then
                                dtgDatos.Rows(x).Cells(55).Value = rwCostoSocial(0)("imss")
                                dtgDatos.Rows(x).Cells(56).Value = rwCostoSocial(0)("RCV")
                                dtgDatos.Rows(x).Cells(57).Value = rwCostoSocial(0)("Infonavit")
                                dtgDatos.Rows(x).Cells(58).Value = rwCostoSocial(0)("ISN")
                                dtgDatos.Rows(x).Cells(59).Value = Math.Round(Double.Parse(dtgDatos.Rows(x).Cells(55).Value) + Double.Parse(dtgDatos.Rows(x).Cells(56).Value) + Double.Parse(dtgDatos.Rows(x).Cells(57).Value) + Double.Parse(dtgDatos.Rows(x).Cells(58).Value), 2)
                            Else
                                dtgDatos.Rows(x).Cells(55).Value = Math.Round(Double.Parse(rwCostoSocial(0)("imss")) / 30 * dtgDatos.Rows(x).Cells(18).Value, 2)
                                dtgDatos.Rows(x).Cells(56).Value = Math.Round(Double.Parse(rwCostoSocial(0)("RCV")) / 30 * dtgDatos.Rows(x).Cells(18).Value, 2)
                                dtgDatos.Rows(x).Cells(57).Value = Math.Round(Double.Parse(rwCostoSocial(0)("Infonavit")) / 30 * dtgDatos.Rows(x).Cells(18).Value, 2)
                                dtgDatos.Rows(x).Cells(58).Value = Math.Round(Double.Parse(rwCostoSocial(0)("ISN")) / 30 * dtgDatos.Rows(x).Cells(18).Value, 2)
                                dtgDatos.Rows(x).Cells(59).Value = Math.Round(Double.Parse(dtgDatos.Rows(x).Cells(55).Value) + Double.Parse(dtgDatos.Rows(x).Cells(56).Value) + Double.Parse(dtgDatos.Rows(x).Cells(57).Value) + Double.Parse(dtgDatos.Rows(x).Cells(58).Value), 2)
                            End If
                        End If
                    End If



                End If

                'Subtotal
                'IVA
                'TOTAL DEPOSITO



                pgbProgreso.Value += 1
                Application.DoEvents()
            Next
            pnlProgreso.Visible = False
            pnlCatalogo.Enabled = True
            MessageBox.Show("Datos calculados ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Function Bisiesto(Num As Integer) As Boolean
        If Num Mod 4 = 0 And (Num Mod 100 Or Num Mod 400 = 0) Then
            Bisiesto = True
        Else
            Bisiesto = False
        End If
    End Function


    Private Function infonavit(tipo As String, valor As Double, sdi As Double, fechapago As Date, periodo As String, diastrabajados As Integer, idempleado As Integer) As Double
        Try
            Dim numbimestre As Integer
            Dim numbimestre2 As Integer
            Dim numdias As Integer
            Dim numdias2 As Integer
            Dim DiasCadaPeriodo As Integer
            Dim DiasCadaPeriodo2 As Integer
            Dim diasfebrero As Integer
            Dim valorinfonavit As Double
            Dim sql As String
            Dim FechaInicioPeriodo1 As Date
            Dim FechaFinPeriodo1 As Date
            Dim FechaInicioPeriodo2 As Date
            Dim FechaFinPeriodo2 As Date
            Dim Seguro1 As Double
            Dim Seguro2 As Double
            Dim ValorInfonavitTabla As Double

            'Validamos si el trabajador tiene o no activo el infonavit
            sql = "select iPermanente from empleadosC where iIdEmpleadoC=" & idempleado
            Dim rwCalcularInfonavit As DataRow() = nConsulta(sql)
            If rwCalcularInfonavit Is Nothing = False Then
                If rwCalcularInfonavit(0)("iPermanente") = "1" Then
                    sql = "select * from periodos where iIdPeriodo= " & periodo
                    Dim rwPeriodo As DataRow() = nConsulta(sql)

                    If rwPeriodo Is Nothing = False Then

                        If diastrabajados = 30 Then
                            FechaInicioPeriodo1 = Date.Parse(rwPeriodo(0)("dFechaInicio"))
                            FechaFinPeriodo1 = Date.Parse("01/" & FechaInicioPeriodo1.Month & "/" & FechaInicioPeriodo1.Year).AddMonths(1).AddDays(-1)
                            FechaFinPeriodo2 = Date.Parse(rwPeriodo(0)("dFechaFin"))
                            FechaInicioPeriodo2 = Date.Parse("01/" & FechaFinPeriodo2.Month & "/" & FechaFinPeriodo2.Year)
                            If (FechaInicioPeriodo1 = FechaInicioPeriodo2) Then
                                FechaInicioPeriodo2 = Date.Parse("01/01/1900")
                            End If

                            If (FechaFinPeriodo1 = FechaFinPeriodo2) Then
                                FechaFinPeriodo2 = Date.Parse("01/01/1900")
                            End If
                        Else
                            'Verificamos si tiene un embarque dentro de periodo
                            sql = "select * from DatosEmbarque where FechaEmbarque Between '" & Date.Parse(rwPeriodo(0)("dFechaInicio")).ToShortDateString & "' and '" & Date.Parse(rwPeriodo(0)("dFechaFin")).ToShortDateString & "'"
                            Dim rwDatosEmbarque As DataRow() = nConsulta(sql)
                            If rwDatosEmbarque Is Nothing = False Then
                                FechaInicioPeriodo1 = rwDatosEmbarque(0)("FechaEmbarque")
                                FechaFinPeriodo2 = FechaInicioPeriodo1.AddDays(diastrabajados)
                                FechaFinPeriodo2 = FechaFinPeriodo2.AddDays(-1)

                                If FechaInicioPeriodo1.Month = FechaFinPeriodo2.Month Then
                                    FechaFinPeriodo1 = FechaInicioPeriodo1.AddDays(diastrabajados - 1)
                                    FechaInicioPeriodo2 = Date.Parse("01/01/1900")
                                    FechaFinPeriodo2 = Date.Parse("01/01/1900")

                                Else

                                    FechaFinPeriodo1 = Date.Parse("01/" & FechaFinPeriodo1.Month & "/" & FechaInicioPeriodo1.Year).AddMonths(1).AddDays(-1)
                                    FechaInicioPeriodo2 = Date.Parse("01/" & FechaFinPeriodo2.Month & "/" & FechaFinPeriodo2.Year)
                                End If


                            Else
                                'Si no lo tiene sumamos de inicio del periodo hasta el numero de dias
                                'Verificamos si esta dentro del mismo mes
                                FechaInicioPeriodo1 = Date.Parse(rwPeriodo(0)("dFechaInicio"))
                                FechaFinPeriodo2 = FechaInicioPeriodo1.AddDays(diastrabajados)
                                FechaFinPeriodo2 = FechaFinPeriodo2.AddDays(-1)
                                If FechaInicioPeriodo1.Month = FechaFinPeriodo2.Month Then
                                    FechaFinPeriodo1 = FechaInicioPeriodo1.AddDays(diastrabajados - 1)
                                    FechaInicioPeriodo2 = Date.Parse("01/01/1900")
                                    FechaFinPeriodo2 = Date.Parse("01/01/1900")

                                Else
                                    FechaFinPeriodo1 = Date.Parse("01/" & FechaFinPeriodo1.Month & "/" & FechaInicioPeriodo1.Year).AddMonths(1).AddDays(-1)
                                    FechaInicioPeriodo2 = Date.Parse("01/" & FechaFinPeriodo2.Month & "/" & FechaFinPeriodo2.Year)
                                End If
                            End If
                        End If





                        If Month(FechaInicioPeriodo1) Mod 2 = 0 Then
                            numbimestre = Month(FechaInicioPeriodo1) / 2
                        Else
                            numbimestre = (Month(FechaInicioPeriodo1) + 1) / 2
                        End If

                        If numbimestre = 1 Then
                            If Bisiesto(Year(FechaInicioPeriodo1)) = True Then
                                diasfebrero = 29
                            Else
                                diasfebrero = 28
                            End If
                            'diasfebrero = Day(DateSerial(Year(fechapago), 3, 0))
                            numdias = 31 + diasfebrero
                        End If

                        If numbimestre = 2 Then
                            numdias = 61
                        End If

                        If numbimestre = 3 Then
                            numdias = 61
                        End If

                        If numbimestre = 4 Then
                            numdias = 62
                        End If

                        If numbimestre = 5 Then
                            numdias = 61
                        End If

                        If numbimestre = 6 Then
                            numdias = 61
                        End If



                        If Month(FechaInicioPeriodo2) Mod 2 = 0 Then
                            numbimestre2 = Month(FechaInicioPeriodo2) / 2
                        Else
                            numbimestre2 = (Month(FechaInicioPeriodo2) + 1) / 2
                        End If

                        If numbimestre2 = 1 Then
                            If Bisiesto(Year(FechaInicioPeriodo1)) = True Then
                                diasfebrero = 29
                            Else
                                diasfebrero = 28
                            End If
                            'diasfebrero = Day(DateSerial(Year(fechapago), 3, 0))
                            numdias2 = 31 + diasfebrero
                        End If

                        If numbimestre2 = 2 Then
                            numdias2 = 61
                        End If

                        If numbimestre2 = 3 Then
                            numdias2 = 61
                        End If

                        If numbimestre2 = 4 Then
                            numdias2 = 62
                        End If

                        If numbimestre2 = 5 Then
                            numdias2 = 61
                        End If

                        If numbimestre2 = 6 Then
                            numdias2 = 61
                        End If



                        DiasCadaPeriodo = DateDiff(DateInterval.Day, FechaInicioPeriodo1, FechaFinPeriodo1) + 1

                        'Verificamos si ya existe el seguro en ese bimestre

                        sql = "select * from PagoSeguroInfonavit where fkiIdEmpleadoC= " & idempleado
                        sql &= " And NumBimestre= " & numbimestre & " And Anio=" & FechaInicioPeriodo1.Year.ToString
                        Dim rwSeguro1 As DataRow() = nConsulta(sql)

                        If rwSeguro1 Is Nothing = False Then
                            Seguro1 = 0
                        Else
                            Seguro1 = 15
                        End If

                        If FechaInicioPeriodo2 = Date.Parse("01/01/1900") Then
                            DiasCadaPeriodo2 = 0
                            Seguro2 = 0

                        Else
                            DiasCadaPeriodo2 = DateDiff(DateInterval.Day, FechaInicioPeriodo2, FechaFinPeriodo2) + 1
                            sql = "select * from PagoSeguroInfonavit where fkiIdEmpleadoC= " & idempleado
                            sql &= " And NumBimestre= " & numbimestre2 & " And Anio=" & FechaInicioPeriodo2.Year.ToString
                            Dim rwSeguro2 As DataRow() = nConsulta(sql)

                            If rwSeguro2 Is Nothing = False Then
                                Seguro2 = 0
                            Else
                                Seguro2 = 15
                            End If

                        End If


                        'Obtener el valor para VSM segun tabla
                        If FechaInicioPeriodo2 = Date.Parse("01/01/1900") Then

                        Else

                        End If


                        sql = "select * from Salario "
                        sql &= " where Anio=" & IIf(FechaFinPeriodo2 = Date.Parse("01/01/1900"), FechaFinPeriodo1.Year.ToString, FechaInicioPeriodo2.Year.ToString)
                        sql &= " and iEstatus=1"
                        Dim rwValorInfonavit As DataRow() = nConsulta(sql)

                        If rwValorInfonavit Is Nothing = False Then
                            ValorInfonavitTabla = rwValorInfonavit(0)("infonavit")
                        Else
                            sql = "select * from Salario "
                            sql &= " where Anio=" & IIf(FechaFinPeriodo2 = Date.Parse("01/01/1900"), FechaFinPeriodo1.Year.ToString, FechaInicioPeriodo2.Year.ToString)
                            sql &= " and iEstatus=1"
                            Dim rwValorInfonavitAntes As DataRow() = nConsulta(sql)
                            If rwValorInfonavitAntes Is Nothing = False Then
                                ValorInfonavitTabla = rwValorInfonavit(0)("infonavit")
                            End If
                        End If



                        If tipo = "VSM" And valor > 0 Then
                            valorinfonavit = (((ValorInfonavitTabla * valor * 2) / numdias) * DiasCadaPeriodo) + Seguro1
                            valorinfonavit = valorinfonavit + ((((ValorInfonavitTabla * valor * 2) / numdias2) * DiasCadaPeriodo2) + IIf(DiasCadaPeriodo2 = 0, 0, Seguro2))
                        End If

                        If tipo = "CUOTA FIJA" And valor > 0 Then


                            valorinfonavit = (((valor * 2) / numdias) * DiasCadaPeriodo) + Seguro1
                            valorinfonavit = valorinfonavit + ((((valor * 2) / numdias2) * DiasCadaPeriodo2) + IIf(DiasCadaPeriodo2 = 0, 0, Seguro2))

                        End If

                        If tipo = "PORCENTAJE" And valor > 0 Then

                            valorinfonavit = ((sdi * (valor / 100) * numdias) + 15) / numdias
                        End If


                        Return valorinfonavit

                    End If

                End If

            End If


            Return 0



        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 0
        End Try
    End Function

    Private Function CalcularInfonavit(tipo As String, valor As Double, sdi As Double, fechapago As Date, periodo As String, idempleado As Integer) As Boolean
        Try
            Dim numbimestre As Integer

            Dim numdias As Integer

            Dim DiasCadaPeriodo As Integer

            Dim diasfebrero As Integer
            Dim valorinfonavit As Double
            Dim sql As String
            Dim FechaInicioPeriodo1 As Date
            Dim FechaFinPeriodo1 As Date
            Dim FechaInicioPeriodo2 As Date
            Dim FechaFinPeriodo2 As Date
            
            Dim ValorInfonavitTabla As Double

            'Validamos si el trabajador tiene o no activo el infonavit
            sql = "select iPermanente from empleadosC where iIdEmpleadoC=" & idempleado
            Dim rwCalcularInfonavit As DataRow() = nConsulta(sql)
            If rwCalcularInfonavit Is Nothing = False Then
                If rwCalcularInfonavit(0)("iPermanente") = "1" Then
                    sql = "select * from periodos where iIdPeriodo= " & periodo
                    Dim rwPeriodo As DataRow() = nConsulta(sql)

                    If rwPeriodo Is Nothing = False Then
                        FechaInicioPeriodo1 = Date.Parse(rwPeriodo(0)("dFechaInicio"))
                        



                        If Month(FechaInicioPeriodo1) Mod 2 = 0 Then
                            numbimestre = Month(FechaInicioPeriodo1) / 2
                        Else
                            numbimestre = (Month(FechaInicioPeriodo1) + 1) / 2
                        End If

                        If numbimestre = 1 Then
                            If Bisiesto(Year(FechaInicioPeriodo1)) = True Then
                                diasfebrero = 29
                            Else
                                diasfebrero = 28
                            End If
                            'diasfebrero = Day(DateSerial(Year(fechapago), 3, 0))
                            numdias = 31 + diasfebrero
                        End If

                        If numbimestre = 2 Then
                            numdias = 61
                        End If

                        If numbimestre = 3 Then
                            numdias = 61
                        End If

                        If numbimestre = 4 Then
                            numdias = 62
                        End If

                        If numbimestre = 5 Then
                            numdias = 61
                        End If

                        If numbimestre = 6 Then
                            numdias = 61
                        End If



                        sql = "select * from Salario "
                        sql &= " where Anio=" & IIf(FechaInicioPeriodo1 = Date.Parse("01/01/1900"), FechaInicioPeriodo1.Year.ToString, FechaInicioPeriodo1.Year.ToString)
                        sql &= " and iEstatus=1"
                        Dim rwValorInfonavit As DataRow() = nConsulta(sql)

                        If rwValorInfonavit Is Nothing = False Then
                            ValorInfonavitTabla = rwValorInfonavit(0)("infonavit")
                        Else
                            
                        End If



                        If tipo = "VSM" And valor > 0 Then
                            valorinfonavit = (((ValorInfonavitTabla * valor * 2) / numdias) * numdias) + 15

                        End If

                        If tipo = "CUOTA FIJA" And valor > 0 Then


                            valorinfonavit = (((valor * 2) / numdias) * numdias) + 15


                        End If

                        If tipo = "PORCENTAJE" And valor > 0 Then

                            valorinfonavit = ((sdi * (valor / 100) * numdias) + 15) / numdias
                        End If


                        'Insertamos los datos

                        sql = "EXEC [setCalculoInfonavitInsertar  ] 0"
                        'Bimestre
                        sql &= "," & numbimestre
                        'Anio
                        sql &= "," & Year(FechaInicioPeriodo1)
                        'TipoFactor
                        sql &= ",'" & tipo
                        'Factor
                        sql &= "'," & valor
                        'idEmpleado
                        sql &= "," & idempleado
                        'Monto
                        sql &= "," & valorinfonavit
                        'Estatus
                        sql &= ",1"
                        





                        If nExecute(sql) = False Then
                            MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            Return False

                        End If

                        Return True
                    End If

                End If

            End If


            Return False



        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try
    End Function

    Private Function VerificarCalculoInfonavit(periodo As String, idempleado As Integer) As Integer

        Try
            Dim numbimestre As Integer

            Dim numdias As Integer

            Dim diasfebrero As Integer

            Dim sql As String
            Dim FechaInicioPeriodo1 As Date


            'Validamos si el trabajador tiene o no activo el infonavit
            sql = "select iPermanente from empleadosC where iIdEmpleadoC=" & idempleado
            Dim rwCalcularInfonavit As DataRow() = nConsulta(sql)
            If rwCalcularInfonavit Is Nothing = False Then
                If rwCalcularInfonavit(0)("iPermanente") = "1" Then
                    sql = "select * from periodos where iIdPeriodo= " & periodo
                    Dim rwPeriodo As DataRow() = nConsulta(sql)

                    If rwPeriodo Is Nothing = False Then
                        FechaInicioPeriodo1 = Date.Parse(rwPeriodo(0)("dFechaInicio"))

                        If Month(FechaInicioPeriodo1) Mod 2 = 0 Then
                            numbimestre = Month(FechaInicioPeriodo1) / 2
                        Else
                            numbimestre = (Month(FechaInicioPeriodo1) + 1) / 2
                        End If

                        If numbimestre = 1 Then
                            If Bisiesto(Year(FechaInicioPeriodo1)) = True Then
                                diasfebrero = 29
                            Else
                                diasfebrero = 28
                            End If
                            'diasfebrero = Day(DateSerial(Year(fechapago), 3, 0))
                            numdias = 31 + diasfebrero
                        End If

                        If numbimestre = 2 Then
                            numdias = 61
                        End If

                        If numbimestre = 3 Then
                            numdias = 61
                        End If

                        If numbimestre = 4 Then
                            numdias = 62
                        End If

                        If numbimestre = 5 Then
                            numdias = 61
                        End If

                        If numbimestre = 6 Then
                            numdias = 61
                        End If





                        'Realizamos la busqueda

                        sql = "select * from CalculoInfonavit where iBimestre=" & numbimestre
                        sql &= " And iAnio= " & Year(FechaInicioPeriodo1) & " And fkiIdEmpleadoC=" & idempleado
                        Dim rwCalculoInfonavit As DataRow() = nConsulta(sql)
                        If rwCalculoInfonavit Is Nothing = False Then
                            Return 1
                        Else
                            Return 2
                        End If

                    Else
                        Return 0
                    End If
                Else
                    Return 0
                End If
            Else
                Return 0
            End If


            Return 0



        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 0
        End Try
    End Function

    Private Function MontoInfonavitF(periodo As String, idempleado As Integer) As Double

        Try
            Dim numbimestre As Integer
            Dim sql As String
            Dim FechaInicioPeriodo1 As Date


            'Validamos si el trabajador tiene o no activo el infonavit
            sql = "select iPermanente from empleadosC where iIdEmpleadoC=" & idempleado
            Dim rwCalcularInfonavit As DataRow() = nConsulta(sql)
            If rwCalcularInfonavit Is Nothing = False Then
                If rwCalcularInfonavit(0)("iPermanente") = "1" Then
                    sql = "select * from periodos where iIdPeriodo= " & periodo
                    Dim rwPeriodo As DataRow() = nConsulta(sql)

                    If rwPeriodo Is Nothing = False Then
                        FechaInicioPeriodo1 = Date.Parse(rwPeriodo(0)("dFechaInicio"))

                        If Month(FechaInicioPeriodo1) Mod 2 = 0 Then
                            numbimestre = Month(FechaInicioPeriodo1) / 2
                        Else
                            numbimestre = (Month(FechaInicioPeriodo1) + 1) / 2
                        End If


                        'Realizamos la busqueda

                        sql = "select * from CalculoInfonavit where iBimestre=" & numbimestre
                        sql &= " And iAnio= " & Year(FechaInicioPeriodo1) & " And fkiIdEmpleadoC=" & idempleado
                        Dim rwCalculoInfonavit As DataRow() = nConsulta(sql)
                        If rwCalculoInfonavit Is Nothing = False Then
                            Return Double.Parse(rwCalculoInfonavit(0)("Monto"))
                            IDCalculoInfonavit = rwCalculoInfonavit(0)("iIdCalculoInfonavit")
                        Else
                            Return 0
                        End If

                    Else
                        Return 0
                    End If
                Else
                    Return 0
                End If
            Else
                Return 0
            End If


            Return 0



        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 0
        End Try
    End Function

    Private Function baseisrtotal(puesto As String, dias As Integer, sdi As Double, incapacidad As Double) As Double
        Dim sueldo As Double
        Dim sueldobase As Double
        Dim baseisr As Double
        Dim isrcalculado As Double
        Dim aguinaldog As Double
        Dim primag As Double
        Dim sql As String
        Dim ValorUMA As Double
        Try

            Sql = "select * from Salario "
            Sql &= " where Anio=" & aniocostosocial
            Sql &= " and iEstatus=1"
            Dim rwValorUMA As DataRow() = nConsulta(Sql)
            If rwValorUMA Is Nothing = False Then
                ValorUMA = Double.Parse(rwValorUMA(0)("uma").ToString)
            Else
                ValorUMA = 0
                MessageBox.Show("No se encontro valor para UMA en el año: " & aniocostosocial)
            End If

            If puesto = "OFICIALES EN PRACTICAS: PILOTIN / ASPIRANTE" Then
                sueldo = sdi * dias
                sueldobase = sueldo
                baseisr = sueldobase - incapacidad
                isrcalculado = isrmensual(baseisr)
            Else
                sueldo = sdi * dias
                sueldobase = (sueldo * (26.19568006 / 100)) + ((sueldo * (8.5070471 / 100)) / 2) + ((sueldo * (8.5070471 / 100)) / 2) + (sueldo * (42.89215164 / 100)) + (sueldo * (9.677848468 / 100))

                ''Aguinaldo gravado
                'aguinaldog = Math.Round(((sueldobase / dias) * 15 / 12 * (dias / 30)) - ((ValorUMA * 30 / 12) * (dias / 30)), 2)


                'primag = (sueldobase * 0.25 / 12 * (dias / 30)) - ((ValorUMA * 15 / 12) * (dias / 30))


                'Aguinaldo gravado 

                If ((sueldobase / dias) * 15 / 12 * (dias / 30)) > ((ValorUMA * 30 / 12) * (dias / 30)) Then
                    'Aguinaldo gravado
                    aguinaldog = Math.Round(((sueldobase / dias) * 15 / 12 * (dias / 30)) - ((ValorUMA * 30 / 12) * (dias / 30)), 2)
                Else
                    'Aguinaldo gravado
                    aguinaldog = "0.00"
                End If

                'Prima de vacaciones

                'Calculos prima
                Dim primavacacionesgravada As Double
                Dim primavacacionesexenta As Double

                primavacacionesgravada = (sueldobase * 0.25 / 12 * (dias / 30)) - ((ValorUMA * 15 / 12) * (dias / 30))
                primavacacionesexenta = ((ValorUMA * 15 / 12) * (dias / 30))

                If primavacacionesgravada > 0 Then
                    primag = primavacacionesgravada

                Else
                    primag = 0
                End If


                baseisr = (sueldobase - ((sueldo * (8.5070471 / 100)) / 2)) + (sueldo * (7.272727273 / 100)) + aguinaldog + primag - incapacidad
                isrcalculado = isrmensual(baseisr)

            End If
            Return isrcalculado
        Catch ex As Exception

        End Try
    End Function


    Private Function isrmensual(monto As Double) As Double

        Dim excendente As Double
        Dim isr As Double
        Dim subsidio As Double



        Dim SQL As String

        Try


            'calculos

            'Calculamos isr

            '1.- buscamos datos para el calculo
            isr = 0
            SQL = "select * from isr where ((" & monto & ">=isr.limiteinf and " & monto & "<=isr.limitesup)"
            SQL &= " or (" & monto & ">=isr.limiteinf and isr.limitesup=0)) and fkiIdTipoPeriodo2=1"


            Dim rwISRCALCULO As DataRow() = nConsulta(SQL)
            If rwISRCALCULO Is Nothing = False Then
                excendente = monto - Double.Parse(rwISRCALCULO(0)("limiteinf").ToString)
                isr = (excendente * (Double.Parse(rwISRCALCULO(0)("porcentaje").ToString) / 100)) + Double.Parse(rwISRCALCULO(0)("cuotafija").ToString)

            End If
            subsidio = 0
            SQL = "select * from subsidio where ((" & monto & ">=subsidio.limiteinf and " & monto & "<=subsidio.limitesup)"
            SQL &= " or (" & monto & ">=subsidio.limiteinf and subsidio.limitesup=0)) and fkiIdTipoPeriodo2=1"


            Dim rwSubsidio As DataRow() = nConsulta(SQL)
            If rwSubsidio Is Nothing = False Then
                subsidio = Double.Parse(rwSubsidio(0)("credito").ToString)

            End If
            If isr > subsidio Then
                Return isr - subsidio
            Else
                Return 0
            End If


        Catch ex As Exception

        End Try
    End Function

    Function subsidiomensual(monto As Double) As Double
        Dim excendente As Double
        Dim isr As Double
        Dim subsidio As Double



        Dim SQL As String

        Try


            'calculos

            'Calculamos isr

            '1.- buscamos datos para el calculo
            isr = 0
            SQL = "select * from isr where ((" & monto & ">=isr.limiteinf and " & monto & "<=isr.limitesup)"
            SQL &= " or (" & monto & ">=isr.limiteinf and isr.limitesup=0)) and fkiIdTipoPeriodo2=1"


            Dim rwISRCALCULO As DataRow() = nConsulta(SQL)
            If rwISRCALCULO Is Nothing = False Then
                excendente = monto - Double.Parse(rwISRCALCULO(0)("limiteinf").ToString)
                isr = (excendente * (Double.Parse(rwISRCALCULO(0)("porcentaje").ToString) / 100)) + Double.Parse(rwISRCALCULO(0)("cuotafija").ToString)

            End If
            subsidio = 0
            SQL = "select * from subsidio where ((" & monto & ">=subsidio.limiteinf and " & monto & "<=subsidio.limitesup)"
            SQL &= " or (" & monto & ">=subsidio.limiteinf and subsidio.limitesup=0)) and fkiIdTipoPeriodo2=1"


            Dim rwSubsidio As DataRow() = nConsulta(SQL)
            If rwSubsidio Is Nothing = False Then
                subsidio = Double.Parse(rwSubsidio(0)("credito").ToString)

            End If

            If isr >= subsidio Then
                subsidiomensual = 0
            Else
                subsidiomensual = subsidio - isr
            End If


        Catch ex As Exception

        End Try



    End Function

    Private Function baseSubsidiototal(puesto As String, dias As Double, sdi As Double, incapacidad As Double) As Double



        Dim sueldo As Double
        Dim sueldobase As Double
        Dim baseisr As Double
        Dim isrcalculado As Double
        Dim aguinaldog As Double
        Dim primag As Double
        Dim sql As String
        Dim ValorUMA As Double
        Try

            sql = "select * from Salario "
            sql &= " where Anio=" & aniocostosocial
            sql &= " and iEstatus=1"
            Dim rwValorUMA As DataRow() = nConsulta(sql)
            If rwValorUMA Is Nothing = False Then
                ValorUMA = Double.Parse(rwValorUMA(0)("uma").ToString)
            Else
                ValorUMA = 0
                MessageBox.Show("No se encontro valor para UMA en el año: " & aniocostosocial)
            End If

            If puesto = "OFICIALES EN PRACTICAS: PILOTIN / ASPIRANTE" Then
                sueldo = sdi * dias
                sueldobase = sueldo
                baseisr = sueldobase - incapacidad
                baseSubsidiototal = subsidiomensual(baseisr)
            Else
                sueldo = sdi * dias
                sueldobase = (sueldo * (26.19568006 / 100)) + ((sueldo * (8.5070471 / 100)) / 2) + ((sueldo * (8.5070471 / 100)) / 2) + (sueldo * (42.89215164 / 100)) + (sueldo * (9.677848468 / 100))

                'Aguinaldo gravado 

                If ((sueldobase / dias) * 15 / 12 * (dias / 30)) > ((ValorUMA * 30 / 12) * (dias / 30)) Then
                    'Aguinaldo gravado
                    aguinaldog = Math.Round(((sueldobase / dias) * 15 / 12 * (dias / 30)) - ((ValorUMA * 30 / 12) * (dias / 30)), 2)
                Else
                    'Aguinaldo gravado
                    aguinaldog = "0.00"
                End If

                'Prima de vacaciones

                'Calculos prima
                Dim primavacacionesgravada As Double
                Dim primavacacionesexenta As Double

                primavacacionesgravada = (sueldobase * 0.25 / 12 * (dias / 30)) - ((ValorUMA * 15 / 12) * (dias / 30))
                primavacacionesexenta = ((ValorUMA * 15 / 12) * (dias / 30))

                If primavacacionesgravada > 0 Then
                    primag = primavacacionesgravada

                Else
                    primag = 0
                End If


                baseisr = (sueldobase - ((sueldo * (8.5070471 / 100)) / 2)) + (sueldo * (7.272727273 / 100)) + aguinaldog + primag - incapacidad
                baseSubsidiototal = subsidiomensual(baseisr)

            End If
            Return baseSubsidiototal
        Catch ex As Exception

        End Try



    End Function


    Function subsidiomensualCausado(monto As Double) As Double
        Dim excendente As Double
        Dim isr As Double
        Dim subsidio As Double



        Dim SQL As String

        Try


            'calculos

            'Calculamos isr

            '1.- buscamos datos para el calculo
            isr = 0
            SQL = "select * from isr where ((" & monto & ">=isr.limiteinf and " & monto & "<=isr.limitesup)"
            SQL &= " or (" & monto & ">=isr.limiteinf and isr.limitesup=0)) and fkiIdTipoPeriodo2=1"


            Dim rwISRCALCULO As DataRow() = nConsulta(SQL)
            If rwISRCALCULO Is Nothing = False Then
                excendente = monto - Double.Parse(rwISRCALCULO(0)("limiteinf").ToString)
                isr = (excendente * (Double.Parse(rwISRCALCULO(0)("porcentaje").ToString) / 100)) + Double.Parse(rwISRCALCULO(0)("cuotafija").ToString)

            End If
            subsidio = 0
            SQL = "select * from subsidio where ((" & monto & ">=subsidio.limiteinf and " & monto & "<=subsidio.limitesup)"
            SQL &= " or (" & monto & ">=subsidio.limiteinf and subsidio.limitesup=0)) and fkiIdTipoPeriodo2=1"


            Dim rwSubsidio As DataRow() = nConsulta(SQL)
            If rwSubsidio Is Nothing = False Then
                subsidio = Double.Parse(rwSubsidio(0)("credito").ToString)

            End If

            If isr >= subsidio Then
                subsidiomensualCausado = 0
            Else
                subsidiomensualCausado = subsidio
            End If


        Catch ex As Exception

        End Try



    End Function


    Function baseSubsidio(puesto As String, dias As Double, sdi As Double, incapacidad As Double) As Double
        Dim sueldo As Double
        Dim sueldobase As Double
        Dim baseisr As Double
        Dim isrcalculado As Double
        Dim aguinaldog As Double
        Dim primag As Double
        Dim sql As String
        Dim ValorUMA As Double
        Try

            sql = "select * from Salario "
            sql &= " where Anio=" & aniocostosocial
            sql &= " and iEstatus=1"
            Dim rwValorUMA As DataRow() = nConsulta(sql)
            If rwValorUMA Is Nothing = False Then
                ValorUMA = Double.Parse(rwValorUMA(0)("uma").ToString)
            Else
                ValorUMA = 0
                MessageBox.Show("No se encontro valor para UMA en el año: " & aniocostosocial)
            End If

            If puesto = "OFICIALES EN PRACTICAS: PILOTIN / ASPIRANTE" Then
                sueldo = sdi * dias
                sueldobase = sueldo
                baseisr = sueldobase - incapacidad
                baseSubsidio = subsidiomensualCausado(baseisr)
            Else
                sueldo = sdi * dias
                sueldobase = (sueldo * (26.19568006 / 100)) + ((sueldo * (8.5070471 / 100)) / 2) + ((sueldo * (8.5070471 / 100)) / 2) + (sueldo * (42.89215164 / 100)) + (sueldo * (9.677848468 / 100))

                'Aguinaldo gravado 

                If ((sueldobase / dias) * 15 / 12 * (dias / 30)) > ((ValorUMA * 30 / 12) * (dias / 30)) Then
                    'Aguinaldo gravado
                    aguinaldog = Math.Round(((sueldobase / dias) * 15 / 12 * (dias / 30)) - ((ValorUMA * 30 / 12) * (dias / 30)), 2)
                Else
                    'Aguinaldo gravado
                    aguinaldog = "0.00"
                End If

                'Prima de vacaciones

                'Calculos prima
                Dim primavacacionesgravada As Double
                Dim primavacacionesexenta As Double

                primavacacionesgravada = (sueldobase * 0.25 / 12 * (dias / 30)) - ((ValorUMA * 15 / 12) * (dias / 30))
                primavacacionesexenta = ((ValorUMA * 15 / 12) * (dias / 30))

                If primavacacionesgravada > 0 Then
                    primag = primavacacionesgravada

                Else
                    primag = 0
                End If

                baseisr = (sueldobase - ((sueldo * (8.5070471 / 100)) / 2)) + (sueldo * (7.272727273 / 100)) + aguinaldog + primag - incapacidad
                baseSubsidio = subsidiomensualCausado(baseisr)

            End If
            Return baseSubsidio
        Catch ex As Exception

        End Try



    End Function


    Private Function Incapacidades(tipo As String, valor As Double, sd As Double) As Double
        Dim incapacidad As Double
        incapacidad = 0.0
        Try
            If tipo = "Riesgo de trabajo" Then
                Incapacidades = 0
            ElseIf tipo = "Enfermedad general" Then
                Incapacidades = valor * sd
            ElseIf tipo = "Maternidad" Then
                Incapacidades = 0
            End If
            Return incapacidad
        Catch ex As Exception

        End Try
    End Function


    Private Sub cmdguardarfinal_Click(sender As Object, e As EventArgs) Handles cmdguardarfinal.Click
        Try
            Dim sql As String
            Dim sql2 As String
            sql = "select * from Nomina where fkiIdEmpresa=1 and fkiIdPeriodo=" & cboperiodo.SelectedValue
            sql &= " and iEstatusNomina=1 and iEstatus=1 and iEstatusEmpleado=" & cboserie.SelectedIndex
            sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex
            'Dim sueldobase, salariodiario, salariointegrado, sueldobruto, TiempoExtraFijoGravado, TiempoExtraFijoExento As Double
            'Dim TiempoExtraOcasional, DesSemObligatorio, VacacionesProporcionales, AguinaldoGravado, AguinaldoExento As Double
            'Dim PrimaVacGravada, PrimaVacExenta, TotalPercepciones, TotalPercepcionesISR As Double
            'Dim incapacidad, ISR, IMSS, Infonavit, InfonavitAnterior, InfonavitAjuste, PensionAlimenticia As Double
            'Dim Prestamo, Fonacot, NetoaPagar, Excedente, Total, ImssCS, RCVCS, InfonavitCS, ISNCS
            'sql = "EXEC getNominaXEmpresaXPeriodo " & gIdEmpresa & "," & cboperiodo.SelectedValue & ",1"

            Dim rwNominaGuardadaFinal As DataRow() = nConsulta(sql)

            If rwNominaGuardadaFinal Is Nothing = False Then
                MessageBox.Show("La nomina ya esta marcada como final, no  se pueden guardar cambios", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                MessageBox.Show("Se borraran los datos tanto de la nomina abordo como la de descanso", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)

                sql = "delete from Nomina"
                sql &= " where fkiIdEmpresa=1 and fkiIdPeriodo=" & cboperiodo.SelectedValue
                sql &= " and iEstatusNomina=0 and iEstatus=1 and iEstatusEmpleado=" & cboserie.SelectedIndex
                'sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex
                If nExecute(sql) = False Then
                    MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    'pnlProgreso.Visible = False
                    Exit Sub
                End If

                sql = "delete from DetalleDescInfonavit"
                sql &= " where fkiIdPeriodo=" & cboperiodo.SelectedValue
                sql &= " and iSerie=" & cboserie.SelectedIndex
                'sql &= " and iSerie=" & cboserie.SelectedIndex
                'sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex

                If nExecute(sql) = False Then
                    MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    'pnlProgreso.Visible = False
                    Exit Sub
                End If

                pnlProgreso.Visible = True

                Application.DoEvents()
                pnlCatalogo.Enabled = False
                pgbProgreso.Minimum = 0
                pgbProgreso.Value = 0
                pgbProgreso.Maximum = dtgDatos.Rows.Count


                For x As Integer = 0 To dtgDatos.Rows.Count - 1



                    sql = "EXEC [setNominaInsertar ] 0"
                    'periodo
                    sql &= "," & cboperiodo.SelectedValue
                    'idempleado
                    sql &= "," & dtgDatos.Rows(x).Cells(2).Value
                    'idempresa
                    sql &= ",1"
                    'Puesto
                    'buscamos el valor en la tabla
                    sql2 = "select * from puestos where cNombre='" & dtgDatos.Rows(x).Cells(11).FormattedValue & "'"

                    Dim rwPuesto As DataRow() = nConsulta(sql2)

                    sql &= "," & rwPuesto(0)("iIdPuesto")


                    'departamento
                    'buscamos el valor en la tabla
                    sql2 = "select * from departamentos where cNombre='" & dtgDatos.Rows(x).Cells(12).FormattedValue & "'"

                    Dim rwDepto As DataRow() = nConsulta(sql2)

                    sql &= "," & rwDepto(0)("iIdDepartamento")

                    'estatus empleado
                    sql &= "," & cboserie.SelectedIndex
                    'edad
                    sql &= "," & dtgDatos.Rows(x).Cells(10).Value
                    'puesto
                    sql &= ",'" & dtgDatos.Rows(x).Cells(11).FormattedValue & "'"
                    'buque
                    sql &= ",'" & dtgDatos.Rows(x).Cells(12).FormattedValue & "'"
                    'iTipo Infonavit
                    sql &= ",'" & dtgDatos.Rows(x).Cells(13).Value & "'"
                    'valor infonavit
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(14).Value = "", "0", dtgDatos.Rows(x).Cells(14).Value.ToString.Replace(",", ""))
                    'salario base
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(15).Value = "", "0", dtgDatos.Rows(x).Cells(15).Value.ToString.Replace(",", ""))
                    'salario diario
                    sql &= "," & dtgDatos.Rows(x).Cells(16).Value
                    'salario integrado
                    sql &= "," & dtgDatos.Rows(x).Cells(17).Value
                    'Dias trabajados
                    sql &= "," & dtgDatos.Rows(x).Cells(18).Value
                    'tipo incapacidad

                    sql &= ",'" & dtgDatos.Rows(x).Cells(19).Value & "'"
                    'numero dias incapacidad
                    sql &= "," & dtgDatos.Rows(x).Cells(20).Value
                    'sueldobruto
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(21).Value = "", "0", dtgDatos.Rows(x).Cells(21).Value.ToString.Replace(",", ""))
                    'tiempo extra fijo gravado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(22).Value = "", "0", dtgDatos.Rows(x).Cells(22).Value.ToString.Replace(",", ""))
                    'tiempo extra fijo exento
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(23).Value = "", "0", dtgDatos.Rows(x).Cells(23).Value.ToString.Replace(",", ""))
                    'Tiempo extra ocasional
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(24).Value = "", "0", dtgDatos.Rows(x).Cells(24).Value.ToString.Replace(",", ""))
                    'descanso semanal obligatorio
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(25).Value = "", "0", dtgDatos.Rows(x).Cells(25).Value.ToString.Replace(",", ""))
                    'vacaciones proporcionales
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(26).Value = "", "0", dtgDatos.Rows(x).Cells(26).Value.ToString.Replace(",", ""))
                    'aguinaldo gravado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(27).Value = "", "0", dtgDatos.Rows(x).Cells(27).Value.ToString.Replace(",", ""))
                    'aguinaldo exento
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(28).Value = "", "0", dtgDatos.Rows(x).Cells(28).Value.ToString.Replace(",", ""))
                    'prima vacacional gravado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(30).Value = "", "0", dtgDatos.Rows(x).Cells(30).Value.ToString.Replace(",", ""))
                    'prima vacacional exento
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(31).Value = "", "0", dtgDatos.Rows(x).Cells(31).Value.ToString.Replace(",", ""))

                    'totalpercepciones
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(33).Value = "", "0", dtgDatos.Rows(x).Cells(33).Value.ToString.Replace(",", ""))
                    'totalpercepcionesISR
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(34).Value = "", "0", dtgDatos.Rows(x).Cells(34).Value.ToString.Replace(",", ""))
                    'Incapacidad
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(35).Value = "", "0", dtgDatos.Rows(x).Cells(35).Value.ToString.Replace(",", ""))
                    'isr
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(36).Value = "", "0", dtgDatos.Rows(x).Cells(36).Value.ToString.Replace(",", ""))
                    'imss
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(37).Value = "", "0", dtgDatos.Rows(x).Cells(37).Value.ToString.Replace(",", ""))
                    'infonavit
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(38).Value = "", "0", dtgDatos.Rows(x).Cells(38).Value.ToString.Replace(",", ""))
                    'infonavit anterior
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(39).Value = "", "0", dtgDatos.Rows(x).Cells(39).Value.ToString.Replace(",", ""))
                    'ajuste infonavit
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(40).Value = "", "0", dtgDatos.Rows(x).Cells(40).Value.ToString.Replace(",", ""))
                    'Pension alimenticia
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(41).Value = "", "0", dtgDatos.Rows(x).Cells(41).Value.ToString.Replace(",", ""))
                    'Prestamo
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(42).Value = "", "0", dtgDatos.Rows(x).Cells(42).Value.ToString.Replace(",", ""))
                    'Fonacot
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(43).Value = "", "0", dtgDatos.Rows(x).Cells(43).Value.ToString.Replace(",", ""))
                    'Subsidio Generado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(44).Value = "", "0", dtgDatos.Rows(x).Cells(44).Value.ToString.Replace(",", ""))
                    'Subsidio Aplicado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(45).Value = "", "0", dtgDatos.Rows(x).Cells(45).Value.ToString.Replace(",", ""))
                    'Operadora
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(46).Value = "", "0", dtgDatos.Rows(x).Cells(46).Value.ToString.Replace(",", ""))
                    'Prestamo Personal Asimilado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(47).Value = "", "0", dtgDatos.Rows(x).Cells(47).Value.ToString.Replace(",", ""))
                    'Adeudo_Infonavit_Asimilado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(48).Value = "", "0", dtgDatos.Rows(x).Cells(48).Value.ToString.Replace(",", ""))
                    'Difencia infonavit Asimilado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(49).Value = "", "0", dtgDatos.Rows(x).Cells(49).Value.ToString.Replace(",", ""))
                    'Complemento Asimilado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(50).Value = "", "0", dtgDatos.Rows(x).Cells(50).Value.ToString.Replace(",", ""))
                    'Retenciones_Operadora
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(51).Value = "", "0", dtgDatos.Rows(x).Cells(51).Value.ToString.Replace(",", ""))
                    '% Comision
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(52).Value = "", "0", dtgDatos.Rows(x).Cells(52).Value.ToString.Replace(",", ""))
                    'Comision_Operadora
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(53).Value = "", "0", dtgDatos.Rows(x).Cells(53).Value.ToString.Replace(",", ""))
                    'Comision asimilados
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(54).Value = "", "0", dtgDatos.Rows(x).Cells(54).Value.ToString.Replace(",", ""))
                    'IMSS_CS
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(55).Value = "", "0", dtgDatos.Rows(x).Cells(55).Value.ToString.Replace(",", ""))
                    'RCV_CS
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(56).Value = "", "0", dtgDatos.Rows(x).Cells(56).Value.ToString.Replace(",", ""))
                    'Infonavit_CS
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(57).Value = "", "0", dtgDatos.Rows(x).Cells(57).Value.ToString.Replace(",", ""))
                    'ISN_CS
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(58).Value = "", "0", dtgDatos.Rows(x).Cells(58).Value.ToString.Replace(",", ""))
                    'Total Costo Social
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(59).Value = "", "0", dtgDatos.Rows(x).Cells(59).Value.ToString.Replace(",", ""))
                    'Subtotal
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(60).Value = "", "0", dtgDatos.Rows(x).Cells(60).Value.ToString.Replace(",", ""))
                    'IVA
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(61).Value = "", "0", dtgDatos.Rows(x).Cells(61).Value.ToString.Replace(",", ""))
                    'TOTAL DEPOSITO
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(62).Value = "", "0", dtgDatos.Rows(x).Cells(62).Value.ToString.Replace(",", ""))
                    'Estatus
                    sql &= ",1"
                    'Estatus Nomina
                    sql &= ",1"
                    'Tipo Nomina
                    sql &= "," & cboTipoNomina.SelectedIndex




                    If nExecute(sql) = False Then
                        MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        'pnlProgreso.Visible = False
                        Exit Sub
                    End If

                    'sql = "update empleadosC set fSueldoOrd=" & dtgDatos.Rows(x).Cells(6).Value & ", fCosto =" & dtgDatos.Rows(x).Cells(18).Value
                    'sql &= " where iIdEmpleadoC = " & dtgDatos.Rows(x).Cells(2).Value

                    'If nExecute(sql) = False Then
                    '    MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    '    'pnlProgreso.Visible = False
                    '    Exit Sub
                    'End If

                    pgbProgreso.Value += 1
                    Application.DoEvents()
                Next
                pnlProgreso.Visible = False
                pnlCatalogo.Enabled = True
                MessageBox.Show("Datos guardados correctamente", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

    Private Sub cboperiodo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboperiodo.SelectedIndexChanged
        Try
            dtgDatos.DataSource = ""
            dtgDatos.Columns.Clear()
        Catch ex As Exception

        End Try

    End Sub


    Private Sub dtgDatos_CellMouseDown(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dtgDatos.CellMouseDown
        Try
            If e.RowIndex > -1 Then
                dtgDatos.CurrentCell = dtgDatos.Rows(e.RowIndex).Cells(e.ColumnIndex)


            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub dtgDatos_CellMouseUp(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dtgDatos.CellMouseUp

    End Sub



    Private Sub dtgDatos_KeyPress(sender As Object, e As KeyPressEventArgs) Handles dtgDatos.KeyPress
        Try

            SoloNumero.NumeroDec(e, sender)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub cmdexcel_Click(sender As Object, e As EventArgs) Handles cmdexcel.Click
        Try

            Dim filaExcel As Integer = 0
            Dim dialogo As New SaveFileDialog()
            Dim periodo As String


            If dtgDatos.Rows.Count > 0 Then
                Dim ruta As String
                ruta = My.Application.Info.DirectoryPath() & "\Archivos\TMM.xlsx"
                'ruta = My.Application.Info.DirectoryPath() & "\Archivos\TMM.xlsx"

                Dim book As New ClosedXML.Excel.XLWorkbook(ruta)


                Dim libro As New ClosedXML.Excel.XLWorkbook

                book.Worksheet(1).CopyTo(libro, "NOMINA TOTAL")
                book.Worksheet(2).CopyTo(libro, "OPERADORA ABORDO")
                book.Worksheet(3).CopyTo(libro, "OPERADORA DESCANSO")
                book.Worksheet(4).CopyTo(libro, "DETALLE")
                'book.Worksheets(5).CopyTo(libro, "FACT")


                Dim hoja As IXLWorksheet = libro.Worksheets(0)
                Dim hoja2 As IXLWorksheet = libro.Worksheets(1)
                Dim hoja3 As IXLWorksheet = libro.Worksheets(2)
                Dim hoja4 As IXLWorksheet = libro.Worksheets(3)
                'Dim hoja5 As IXLWorksheet = libro.Worksheets(4)

                '<<<<<<<<<<<<<<<<<<<<<<Nomina Total>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
             

                    hoja.Cell(12, 1).Clear()

                    filaExcel = 13
                    Dim nombrebuque As String
                    Dim inicio As Integer = 0
                    Dim contadorexcelbuqueinicial As Integer = 0
                    Dim contadorexcelbuquefinal As Integer = 0
                    Dim total As Integer = dtgDatos.Rows.Count - 1
                    Dim filatmp As Integer = 13 - 4


                    Dim rwPeriodo0 As DataRow() = nConsulta("Select * from periodos where iIdPeriodo=" & cboperiodo.SelectedValue)
                    If rwPeriodo0 Is Nothing = False Then
                    periodo = MonthString(rwPeriodo0(0).Item("iMes")).ToUpper & " DE " & (rwPeriodo0(0).Item("iEjercicio"))

                    hoja.Cell(10, 2).Style.Font.SetBold(True)
                    hoja.Cell(10, 2).Style.NumberFormat.Format = "@"
                    hoja.Cell(10, 2).Value = periodo

                    End If

                    For x As Integer = 0 To dtgDatos.Rows.Count - 1

                        If inicio = x Then
                            contadorexcelbuqueinicial = filaExcel + x
                            nombrebuque = dtgDatos.Rows(x).Cells(12).Value
                        End If
                        If nombrebuque = dtgDatos.Rows(x).Cells(12).Value Then

                            hoja.Cell(filaExcel + x, 2).Value = dtgDatos.Rows(x).Cells(12).Value
                            hoja.Cell(filaExcel + x, 3).Value = dtgDatos.Rows(x).Cells(5).Value
                            hoja.Cell(filaExcel + x, 4).Value = dtgDatos.Rows(x).Cells(3).Value
                            hoja.Cell(filaExcel + x, 5).Value = dtgDatos.Rows(x).Cells(4).Value
                            hoja.Cell(filaExcel + x, 6).Value = dtgDatos.Rows(x).Cells(11).Value
                            hoja.Cell(filaExcel + x, 7).Value = dtgDatos.Rows(x).Cells(10).Value
                            hoja.Cell(filaExcel + x, 8).Value = dtgDatos.Rows(x).Cells(18).Value
                            hoja.Cell(filaExcel + x, 9).Value = dtgDatos.Rows(x).Cells(18).Value
                            hoja.Cell(filaExcel + x, 10).Value = "" 'dtgDatos.Rows(x).Cells().Value 'Descanso
                            hoja.Cell(filaExcel + x, 11).Value = "" 'dtgDatos.Rows(x).Cells().Value  'Abordo
                            hoja.Cell(filaExcel + x, 12).FormulaA1 = "=J" & filaExcel + x & "+ K" & filaExcel + x
                            hoja.Cell(filaExcel + x, 13).FormulaA1 = "='OPERADORA ABORDO'!AI" & filatmp + x & "+'OPERADORA DESCANSO'!AI" & filatmp + x
                            hoja.Cell(filaExcel + x, 14).FormulaA1 = "='OPERADORA ABORDO'!AJ" & filatmp + x & "+'OPERADORA DESCANSO'!AJ" & filatmp + x
                            hoja.Cell(filaExcel + x, 15).FormulaA1 = "L" & filaExcel + x & "-M" & filaExcel + x & "-N" & filaExcel + x
                            hoja.Cell(filaExcel + x, 16).FormulaA1 = "='OPERADORA ABORDO'!AI" & filatmp + x & "+'OPERADORA DESCANSO'!AI" & filatmp + x
                            hoja.Cell(filaExcel + x, 17).FormulaA1 = "O" & filaExcel + x & "-P" & filaExcel + x
                            hoja.Cell(filaExcel + x, 18).FormulaA1 = "='OPERADORA ABORDO'!AG" & filatmp + x & "+'OPERADORA ABORDO'!AH" & filatmp + x & "+'OPERADORA ABORDO'!AI" & filatmp + x & "+'OPERADORA ABORDO'!AJ" & filatmp + x & "+'OPERADORA DESCANSO'!AG" & filatmp + x & "+'OPERADORA DESCANSO'!AH9+'OPERADORA DESCANSO'!AI" & filatmp + x & "+'OPERADORA DESCANSO'!AJ" & filatmp + x
                            hoja.Cell(filaExcel + x, 19).Value = " "
                            hoja.Cell(filaExcel + x, 20).FormulaA1 = "(P" & filaExcel + x & "+R" & filaExcel + x & ")*2%"
                            hoja.Cell(filaExcel + x, 21).FormulaA1 = "=Q" & filaExcel + x & "*2%"
                            Dim csocial As Double = CDbl(dtgDatos.Rows(x).Cells(49).Value) + CDbl(dtgDatos.Rows(x).Cells(50).Value) + CDbl(dtgDatos.Rows(x).Cells(51).Value) + CDbl(dtgDatos.Rows(x).Cells(52).Value)
                        hoja.Cell(filaExcel + x, 22).Value = dtgDatos.Rows(x).Cells(59).Value 'COSTO SOCIAL
                            hoja.Cell(filaExcel + x, 24).FormulaA1 = "=P" & filaExcel + x & "+Q" & filaExcel + x & "+R" & filaExcel + x & "+T" & filaExcel + x & "+U" & filaExcel + x & "+V" & filaExcel + x
                            hoja.Cell(filaExcel + x, 25).FormulaA1 = "=X" & filaExcel + x & "*16%"
                            hoja.Cell(filaExcel + x, 26).FormulaA1 = "=X" & filaExcel + x & "+Y" & filaExcel + x


                        'Detalle

                     
                        Else
                            contadorexcelbuquefinal = filaExcel + x - 1
                        hoja.Cell(filaExcel + x, 12).FormulaA1 = "=SUM(L" & contadorexcelbuqueinicial & ":L" & contadorexcelbuquefinal & ")"
                        hoja.Cell(filaExcel + x, 13).FormulaA1 = "=SUM(M" & contadorexcelbuqueinicial & ":M" & contadorexcelbuquefinal & ")"
                        hoja.Cell(filaExcel + x, 14).FormulaA1 = "=SUM(N" & contadorexcelbuqueinicial & ":N" & contadorexcelbuquefinal & ")"
                        hoja.Cell(filaExcel + x, 15).FormulaA1 = "=SUM(O" & contadorexcelbuqueinicial & ":O" & contadorexcelbuquefinal & ")"
                        hoja.Cell(filaExcel + x, 16).FormulaA1 = "=SUM(P" & contadorexcelbuqueinicial & ":P" & contadorexcelbuquefinal & ")"
                        hoja.Cell(filaExcel + x, 17).FormulaA1 = "=SUM(Q" & contadorexcelbuqueinicial & ":Q" & contadorexcelbuquefinal & ")"
                        hoja.Cell(filaExcel + x, 18).FormulaA1 = "=SUM(R" & contadorexcelbuqueinicial & ":R" & contadorexcelbuquefinal & ")"
                        hoja.Cell(filaExcel + x, 20).FormulaA1 = "=SUM(T" & contadorexcelbuqueinicial & ":T" & contadorexcelbuquefinal & ")"
                        hoja.Cell(filaExcel + x, 21).FormulaA1 = "=SUM(U" & contadorexcelbuqueinicial & ":U" & contadorexcelbuquefinal & ")"
                        hoja.Cell(filaExcel + x, 22).FormulaA1 = "=SUM(V" & contadorexcelbuqueinicial & ":V" & contadorexcelbuquefinal & ")"
                        'hoja.Cell(filaExcel + x, 23).FormulaA1 = "=SUMA(W" & contadorexcelbuqueinicial & ":W" & contadorexcelbuquefinal & ")"
                        hoja.Cell(filaExcel + x, 24).FormulaA1 = "=SUM(X" & contadorexcelbuqueinicial & ":X" & contadorexcelbuquefinal & ")"
                        hoja.Cell(filaExcel + x, 25).FormulaA1 = "=SUM(Y" & contadorexcelbuqueinicial & ":Y" & contadorexcelbuquefinal & ")"
                        hoja.Cell(filaExcel + x, 26).FormulaA1 = "=SUM(Z" & contadorexcelbuqueinicial & ":Z" & contadorexcelbuquefinal & ")"

                            hoja.Range(filaExcel + x, 12, filaExcel + x, 26).Style.Fill.BackgroundColor = XLColor.PowderBlue
                            hoja.Range(filaExcel + x, 12, filaExcel + x, 26).Style.Font.SetBold(True)

                            nombrebuque = dtgDatos.Rows(x).Cells(12).Value
                            filaExcel = filaExcel + 2
                            contadorexcelbuqueinicial = filaExcel + x
                            contadorexcelbuquefinal = 0

                            hoja.Cell(filaExcel + x, 2).Value = dtgDatos.Rows(x).Cells(12).Value
                            hoja.Cell(filaExcel + x, 3).Value = dtgDatos.Rows(x).Cells(5).Value
                            hoja.Cell(filaExcel + x, 4).Value = dtgDatos.Rows(x).Cells(3).Value
                            hoja.Cell(filaExcel + x, 5).Value = dtgDatos.Rows(x).Cells(4).Value
                            hoja.Cell(filaExcel + x, 6).Value = dtgDatos.Rows(x).Cells(11).Value
                            hoja.Cell(filaExcel + x, 7).Value = dtgDatos.Rows(x).Cells(10).Value
                            hoja.Cell(filaExcel + x, 8).Value = dtgDatos.Rows(x).Cells(18).Value
                            hoja.Cell(filaExcel + x, 9).Value = dtgDatos.Rows(x).Cells(18).Value
                            hoja.Cell(filaExcel + x, 10).Value = "" 'dtgDatos.Rows(x).Cells().Value 'Descanso
                            hoja.Cell(filaExcel + x, 11).Value = "" 'dtgDatos.Rows(x).Cells().Value  'Abordo
                            hoja.Cell(filaExcel + x, 12).FormulaA1 = "=J" & filaExcel + x & "+ K" & filaExcel + x
                            hoja.Cell(filaExcel + x, 13).FormulaA1 = "='OPERADORA ABORDO'!AI" & filatmp + x & "+'OPERADORA DESCANSO'!AI" & filatmp + x
                            hoja.Cell(filaExcel + x, 14).FormulaA1 = "='OPERADORA ABORDO'!AJ" & filatmp + x & "+'OPERADORA DESCANSO'!AJ" & filatmp + x
                            hoja.Cell(filaExcel + x, 15).FormulaA1 = "L" & filaExcel + x & "-M" & filaExcel + x & "-N" & filaExcel + x
                            hoja.Cell(filaExcel + x, 16).FormulaA1 = "='OPERADORA ABORDO'!AI" & filatmp + x & "+'OPERADORA DESCANSO'!AI" & filatmp + x
                            hoja.Cell(filaExcel + x, 17).FormulaA1 = "O" & filaExcel + x & "-P" & filaExcel + x
                            hoja.Cell(filaExcel + x, 18).FormulaA1 = "='OPERADORA ABORDO'!AG" & filatmp + x & "+'OPERADORA ABORDO'!AH" & filatmp + x & "+'OPERADORA ABORDO'!AI" & filatmp + x & "+'OPERADORA ABORDO'!AJ" & filatmp + x & "+'OPERADORA DESCANSO'!AG" & filatmp + x & "+'OPERADORA DESCANSO'!AH9+'OPERADORA DESCANSO'!AI" & filatmp + x & "+'OPERADORA DESCANSO'!AJ" & filatmp + x
                            hoja.Cell(filaExcel + x, 19).Value = " "
                            hoja.Cell(filaExcel + x, 20).FormulaA1 = "(P" & filaExcel + x & "+R" & filaExcel + x & ")*2%"
                            hoja.Cell(filaExcel + x, 21).FormulaA1 = "=Q" & filaExcel + x & "*2%"
                            Dim csocial As Double = CDbl(dtgDatos.Rows(x).Cells(49).Value) + CDbl(dtgDatos.Rows(x).Cells(50).Value) + CDbl(dtgDatos.Rows(x).Cells(51).Value) + CDbl(dtgDatos.Rows(x).Cells(52).Value)
                            hoja.Cell(filaExcel + x, 22).Value = csocial 'COSTO SOCIAL
                            hoja.Cell(filaExcel + x, 23).Value = ""
                            hoja.Cell(filaExcel + x, 24).FormulaA1 = "=P" & filaExcel + x & "+Q" & filaExcel + x & "+R" & filaExcel + x & "+T" & filaExcel + x & "+U" & filaExcel + x & "+V" & filaExcel + x
                            hoja.Cell(filaExcel + x, 25).FormulaA1 = "=X" & filaExcel + x & "*16%"
                            hoja.Cell(filaExcel + x, 26).FormulaA1 = "=X" & filaExcel + x & "+Y" & filaExcel + x

                        End If
                    Next x
                    filaExcel = filaExcel + 2
                    contadorexcelbuquefinal = filaExcel + total - 1
                hoja.Cell(filaExcel + total, 12).FormulaA1 = "=SUM(L" & contadorexcelbuqueinicial & ":L" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 13).FormulaA1 = "=SUM(M" & contadorexcelbuqueinicial & ":M" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 14).FormulaA1 = "=SUM(N" & contadorexcelbuqueinicial & ":N" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 15).FormulaA1 = "=SUM(O" & contadorexcelbuqueinicial & ":O" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 16).FormulaA1 = "=SUM(P" & contadorexcelbuqueinicial & ":P" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 17).FormulaA1 = "=SUM(Q" & contadorexcelbuqueinicial & ":Q" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 18).FormulaA1 = "=SUM(R" & contadorexcelbuqueinicial & ":R" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 20).FormulaA1 = "=SUM(T" & contadorexcelbuqueinicial & ":T" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 21).FormulaA1 = "=SUM(U" & contadorexcelbuqueinicial & ":U" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 22).FormulaA1 = "=SUM(V" & contadorexcelbuqueinicial & ":V" & contadorexcelbuquefinal & ")"
                'hoja.Cell(filaExcel + total, 23).FormulaA1 = "=SUMA(W" & contadorexcelbuqueinicial & ":W" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 24).FormulaA1 = "=SUM(X" & contadorexcelbuqueinicial & ":X" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 25).FormulaA1 = "=SUM(Y" & contadorexcelbuqueinicial & ":Y" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 26).FormulaA1 = "=SUM(Z" & contadorexcelbuqueinicial & ":Z" & contadorexcelbuquefinal & ")"

                    hoja.Range(filaExcel + total, 12, filaExcel + total, 26).Style.Fill.BackgroundColor = XLColor.PowderBlue
                    hoja.Range(filaExcel + total, 12, filaExcel + total, 26).Style.Font.SetBold(True)


                    'Formulas



                    'Nomina Tottal

                    hoja.Cell(12, 27).Clear()
                    hoja.Cell(12, 28).Clear()
                    hoja.Cell(12, 29).Clear()
                    hoja.Cell(12, 30).Clear()
                    hoja.Cell(12, 31).Clear()
                    hoja.Cell(12, 32).Clear()
                    hoja.Cell(12, 33).Clear()



                    '<<<<<<<<<<<<<<<<<Operadora Abordo>>>>>>>>>>>>>>>>>>>>>>>>

                Dim rwPeriodo As DataRow() = nConsulta("Select (CONVERT(nvarchar(12),dFechaInicio,103) + ' al ' + CONVERT(nvarchar(12),dFechaFin,103)) as dFechaInicio from periodos where iIdPeriodo=" & cboperiodo.SelectedValue)
                If rwPeriodo Is Nothing = False Then
                    hoja2.Cell(4, 2).Value = "Periodo Mensual del " & rwPeriodo(0).Item("dFechaInicio")
                    hoja3.Cell(4, 2).Value = "Periodo Mensual del " & rwPeriodo(0).Item("dFechaInicio")

                End If


                    ''OPERADORA ABORDO
                    filaExcel = 9
                    For x As Integer = 0 To dtgDatos.Rows.Count - 1

                        ' dtgDatos.Rows(x).Cells(11).FormattedValue

                        hoja2.Cell(filaExcel, 1).Value = dtgDatos.Rows(x).Cells(3).Value
                        hoja2.Cell(filaExcel, 2).Value = dtgDatos.Rows(x).Cells(4).Value
                        hoja2.Cell(filaExcel, 3).Value = dtgDatos.Rows(x).Cells(5).Value
                        hoja2.Cell(filaExcel, 4).Value = dtgDatos.Rows(x).Cells(6).Value
                        hoja2.Cell(filaExcel, 5).Value = dtgDatos.Rows(x).Cells(7).Value
                        hoja2.Cell(filaExcel, 6).Value = dtgDatos.Rows(x).Cells(8).Value
                        hoja2.Cell(filaExcel, 7).Value = dtgDatos.Rows(x).Cells(9).Value
                        hoja2.Cell(filaExcel, 8).Value = dtgDatos.Rows(x).Cells(10).Value
                        hoja2.Cell(filaExcel, 9).Value = dtgDatos.Rows(x).Cells(11).FormattedValue
                        hoja2.Cell(filaExcel, 10).Value = dtgDatos.Rows(x).Cells(12).FormattedValue
                        hoja2.Cell(filaExcel, 11).Value = dtgDatos.Rows(x).Cells(13).Value
                        hoja2.Cell(filaExcel, 12).Value = dtgDatos.Rows(x).Cells(14).Value
                        hoja2.Cell(filaExcel, 13).Value = dtgDatos.Rows(x).Cells(15).Value
                        hoja2.Cell(filaExcel, 14).Value = dtgDatos.Rows(x).Cells(16).Value
                        hoja2.Cell(filaExcel, 15).Value = dtgDatos.Rows(x).Cells(18).Value
                        hoja2.Cell(filaExcel, 16).Value = dtgDatos.Rows(x).Cells(19).Value
                        hoja2.Cell(filaExcel, 17).Value = dtgDatos.Rows(x).Cells(20).Value
                        hoja2.Cell(filaExcel, 18).Value = dtgDatos.Rows(x).Cells(21).Value
                        hoja2.Cell(filaExcel, 19).Value = dtgDatos.Rows(x).Cells(22).Value
                        hoja2.Cell(filaExcel, 20).Value = dtgDatos.Rows(x).Cells(23).Value
                        hoja2.Cell(filaExcel, 21).Value = dtgDatos.Rows(x).Cells(24).Value
                        hoja2.Cell(filaExcel, 22).Value = dtgDatos.Rows(x).Cells(25).Value
                        hoja2.Cell(filaExcel, 23).Value = dtgDatos.Rows(x).Cells(26).Value
                        hoja2.Cell(filaExcel, 24).Value = dtgDatos.Rows(x).Cells(27).Value
                        hoja2.Cell(filaExcel, 25).Value = dtgDatos.Rows(x).Cells(28).Value
                        hoja2.Cell(filaExcel, 26).Value = dtgDatos.Rows(x).Cells(29).Value
                        hoja2.Cell(filaExcel, 27).Value = dtgDatos.Rows(x).Cells(30).Value
                        hoja2.Cell(filaExcel, 28).Value = dtgDatos.Rows(x).Cells(31).Value
                        hoja2.Cell(filaExcel, 29).Value = dtgDatos.Rows(x).Cells(32).Value
                        hoja2.Cell(filaExcel, 30).Value = dtgDatos.Rows(x).Cells(33).Value
                        hoja2.Cell(filaExcel, 31).Value = dtgDatos.Rows(x).Cells(34).Value
                        hoja2.Cell(filaExcel, 32).Value = dtgDatos.Rows(x).Cells(35).Value
                        hoja2.Cell(filaExcel, 33).Value = dtgDatos.Rows(x).Cells(36).Value
                        hoja2.Cell(filaExcel, 34).Value = dtgDatos.Rows(x).Cells(37).Value
                        hoja2.Cell(filaExcel, 35).Value = dtgDatos.Rows(x).Cells(38).Value
                        hoja2.Cell(filaExcel, 36).Value = dtgDatos.Rows(x).Cells(41).Value
                        hoja2.Cell(filaExcel, 37).Value = dtgDatos.Rows(x).Cells(45).Value
                        hoja2.Cell(filaExcel, 38).Value = dtgDatos.Rows(x).Cells(46).Value

                        filaExcel = filaExcel + 1


                    Next x

                    'Formulas

                    'Operadora Abordo

                    hoja2.Cell(filaExcel + 4, 18).FormulaA1 = "=SUM(R9:R" & filaExcel & ")"
                    hoja2.Cell(filaExcel + 4, 19).FormulaA1 = "=SUM(S9:S" & filaExcel & ")"
                    hoja2.Cell(filaExcel + 4, 20).FormulaA1 = "=SUM(T9:T" & filaExcel & ")"
                    hoja2.Cell(filaExcel + 4, 21).FormulaA1 = "=SUM(U9:U" & filaExcel & ")"
                    hoja2.Cell(filaExcel + 4, 22).FormulaA1 = "=SUM(V9:V" & filaExcel & ")"
                hoja2.Cell(filaExcel + 4, 23).FormulaA1 = "=SUM(W9:W" & filaExcel & ")"
                    hoja2.Cell(filaExcel + 4, 24).FormulaA1 = "=SUM(X9:X" & filaExcel & ")"
                    hoja2.Cell(filaExcel + 4, 25).FormulaA1 = "=SUM(Y9:Y" & filaExcel & ")"
                    hoja2.Cell(filaExcel + 4, 26).FormulaA1 = "=SUM(Z9:Z" & filaExcel & ")"
                    hoja2.Cell(filaExcel + 4, 27).FormulaA1 = "=SUM(AA9:AA" & filaExcel & ")"
                    hoja2.Cell(filaExcel + 4, 28).FormulaA1 = "=SUM(AB9:AB" & filaExcel & ")"
                    hoja2.Cell(filaExcel + 4, 29).FormulaA1 = "=SUM(AC9:AC" & filaExcel & ")"
                    hoja2.Cell(filaExcel + 4, 30).FormulaA1 = "=SUM(AD9:AD" & filaExcel & ")"
                    hoja2.Cell(filaExcel + 4, 31).FormulaA1 = "=SUM(AE9:AE" & filaExcel & ")"
                    hoja2.Cell(filaExcel + 4, 32).FormulaA1 = "=SUM(AF9:AF" & filaExcel & ")"
                    hoja2.Cell(filaExcel + 4, 33).FormulaA1 = "=SUM(AG9:AG" & filaExcel & ")"
                    hoja2.Cell(filaExcel + 4, 34).FormulaA1 = "=SUM(AH9:AH" & filaExcel & ")"
                    hoja2.Cell(filaExcel + 4, 35).FormulaA1 = "=SUM(AI9:AI" & filaExcel & ")"
                    hoja2.Cell(filaExcel + 4, 36).FormulaA1 = "=SUM(AJ9:AJ" & filaExcel & ")"
                    hoja2.Cell(filaExcel + 4, 37).FormulaA1 = "=SUM(AK9:AK" & filaExcel & ")"
                    hoja2.Cell(filaExcel + 4, 38).FormulaA1 = "=SUM(AL9:AL" & filaExcel & ")"


                'hoja2.Cell(8, 39).Clear()
                'hoja2.Cell(8, 40).Clear()
                'hoja2.Cell(8, 41).Clear()
                'hoja2.Cell(8, 42).Clear()
                'hoja2.Cell(8, 43).Clear()
                'hoja2.Cell(8, 44).Clear()
                'hoja2.Cell(8, 45).Clear()
                'hoja2.Cell(8, 46).Clear()
                'hoja2.Cell(8, 47).Clear()
                'hoja2.Cell(8, 48).Clear()
                'hoja2.Cell(8, 49).Clear()
                'hoja2.Cell(8, 50).Clear()
                'hoja2.Cell(8, 51).Clear()
                'hoja2.Cell(8, 52).Clear()

                limpiarCell(hoja2, 39) ', 1, dtgDatos.Rows.Count - 1)
                    '<<<<<<<<<<<<<<<Operadora Descanso>>>>>>>>>>>>>>>>>>

              

                    ''Operadora Descanso
                    filaExcel = 9
                    For x As Integer = 0 To dtgDatos.Rows.Count - 1

                        ' dtgDatos.Rows(x).Cells(11).FormattedValue

                        hoja3.Cell(filaExcel, 1).Value = dtgDatos.Rows(x).Cells(3).Value
                        hoja3.Cell(filaExcel, 2).Value = dtgDatos.Rows(x).Cells(4).Value
                        hoja3.Cell(filaExcel, 3).Value = dtgDatos.Rows(x).Cells(5).Value
                        hoja3.Cell(filaExcel, 4).Value = dtgDatos.Rows(x).Cells(6).Value
                        hoja3.Cell(filaExcel, 5).Value = dtgDatos.Rows(x).Cells(7).Value
                        hoja3.Cell(filaExcel, 6).Value = dtgDatos.Rows(x).Cells(8).Value
                        hoja3.Cell(filaExcel, 7).Value = dtgDatos.Rows(x).Cells(9).Value
                        hoja3.Cell(filaExcel, 8).Value = dtgDatos.Rows(x).Cells(10).Value
                        hoja3.Cell(filaExcel, 9).Value = dtgDatos.Rows(x).Cells(11).FormattedValue
                        hoja3.Cell(filaExcel, 10).Value = dtgDatos.Rows(x).Cells(12).FormattedValue
                    hoja3.Cell(filaExcel, 11).Value = dtgDatos.Rows(x).Cells(13).Value
                    If dtgDatos.Rows(x).Cells(11).FormattedValue = "OFICIALES EN PRACTICAS: PILOTIN / ASPIRANTE" Then

                        hoja3.Cell(filaExcel, 12).Value = 0
                        hoja3.Cell(filaExcel, 13).Value = 0
                        hoja3.Cell(filaExcel, 14).Value = 0
                        hoja3.Cell(filaExcel, 15).Value = 0
                        hoja3.Cell(filaExcel, 16).Value = 0
                        hoja3.Cell(filaExcel, 17).Value = 0
                        hoja3.Cell(filaExcel, 18).Value = 0
                        hoja3.Cell(filaExcel, 19).Value = 0
                        hoja3.Cell(filaExcel, 20).Value = 0
                        hoja3.Cell(filaExcel, 21).Value = 0
                        hoja3.Cell(filaExcel, 22).Value = 0
                        hoja3.Cell(filaExcel, 23).Value = 0
                        hoja3.Cell(filaExcel, 24).Value = 0
                        hoja3.Cell(filaExcel, 25).Value = 0
                        hoja3.Cell(filaExcel, 26).Value = 0
                        hoja3.Cell(filaExcel, 27).Value = 0
                        hoja3.Cell(filaExcel, 28).Value = 0
                        hoja3.Cell(filaExcel, 29).Value = 0
                        hoja3.Cell(filaExcel, 30).Value = 0
                        hoja3.Cell(filaExcel, 31).Value = 0
                        hoja3.Cell(filaExcel, 32).Value = 0
                        hoja3.Cell(filaExcel, 33).Value = 0
                        hoja3.Cell(filaExcel, 34).Value = 0
                        hoja3.Cell(filaExcel, 35).Value = 0
                        hoja3.Cell(filaExcel, 36).Value = 0
                        hoja3.Cell(filaExcel, 37).Value = 0
                        hoja3.Cell(filaExcel, 38).Value = 0


                    Else
                        hoja3.Cell(filaExcel, 12).Value = dtgDatos.Rows(x).Cells(14).Value
                        hoja3.Cell(filaExcel, 13).Value = dtgDatos.Rows(x).Cells(15).Value
                        hoja3.Cell(filaExcel, 14).Value = dtgDatos.Rows(x).Cells(16).Value
                        hoja3.Cell(filaExcel, 15).Value = dtgDatos.Rows(x).Cells(18).Value
                        hoja3.Cell(filaExcel, 16).Value = dtgDatos.Rows(x).Cells(19).Value
                        hoja3.Cell(filaExcel, 17).Value = dtgDatos.Rows(x).Cells(20).Value
                        hoja3.Cell(filaExcel, 18).Value = dtgDatos.Rows(x).Cells(21).Value
                        hoja3.Cell(filaExcel, 19).Value = dtgDatos.Rows(x).Cells(22).Value
                        hoja3.Cell(filaExcel, 20).Value = dtgDatos.Rows(x).Cells(23).Value
                        hoja3.Cell(filaExcel, 21).Value = dtgDatos.Rows(x).Cells(24).Value
                        hoja3.Cell(filaExcel, 22).Value = dtgDatos.Rows(x).Cells(25).Value
                        hoja3.Cell(filaExcel, 23).Value = dtgDatos.Rows(x).Cells(26).Value
                        hoja3.Cell(filaExcel, 24).Value = dtgDatos.Rows(x).Cells(27).Value
                        hoja3.Cell(filaExcel, 25).Value = dtgDatos.Rows(x).Cells(28).Value
                        hoja3.Cell(filaExcel, 26).Value = dtgDatos.Rows(x).Cells(29).Value
                        hoja3.Cell(filaExcel, 27).Value = dtgDatos.Rows(x).Cells(30).Value
                        hoja3.Cell(filaExcel, 28).Value = dtgDatos.Rows(x).Cells(31).Value
                        hoja3.Cell(filaExcel, 29).Value = dtgDatos.Rows(x).Cells(32).Value
                        hoja3.Cell(filaExcel, 30).Value = dtgDatos.Rows(x).Cells(33).Value
                        hoja3.Cell(filaExcel, 31).Value = dtgDatos.Rows(x).Cells(34).Value
                        hoja3.Cell(filaExcel, 32).Value = dtgDatos.Rows(x).Cells(35).Value
                        hoja3.Cell(filaExcel, 33).Value = dtgDatos.Rows(x).Cells(36).Value
                        hoja3.Cell(filaExcel, 34).Value = dtgDatos.Rows(x).Cells(37).Value
                        hoja3.Cell(filaExcel, 35).Value = dtgDatos.Rows(x).Cells(38).Value
                        hoja3.Cell(filaExcel, 36).Value = dtgDatos.Rows(x).Cells(41).Value
                        hoja3.Cell(filaExcel, 37).Value = dtgDatos.Rows(x).Cells(45).Value
                        hoja3.Cell(filaExcel, 38).Value = dtgDatos.Rows(x).Cells(46).Value
                    End If

                  
                    filaExcel = filaExcel + 1


                Next x

                    'Formulas
                    hoja3.Cell(filaExcel + 4, 18).FormulaA1 = "=SUM(R9:R" & filaExcel & ")"

                    'Operadora Descanso

                    hoja3.Cell(filaExcel + 4, 18).FormulaA1 = "=SUM(R9:R" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 19).FormulaA1 = "=SUM(S9:S" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 20).FormulaA1 = "=SUM(T9:T" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 21).FormulaA1 = "=SUM(U9:U" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 22).FormulaA1 = "=SUM(V9:V" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 23).FormulaA1 = "=SUM(W9:W" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 24).FormulaA1 = "=SUM(X9:X" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 25).FormulaA1 = "=SUM(Y9:Y" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 26).FormulaA1 = "=SUM(Z9:Z" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 27).FormulaA1 = "=SUM(AA9:AA" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 28).FormulaA1 = "=SUM(AB9:AB" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 29).FormulaA1 = "=SUM(AC9:AC" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 30).FormulaA1 = "=SUM(AD9:AD" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 31).FormulaA1 = "=SUM(AE9:AE" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 32).FormulaA1 = "=SUM(AF9:AF" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 33).FormulaA1 = "=SUM(AG9:AG" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 34).FormulaA1 = "=SUM(AH9:AH" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 35).FormulaA1 = "=SUM(AI9:AI" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 36).FormulaA1 = "=SUM(AJ9:AJ" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 37).FormulaA1 = "=SUM(AK9:AK" & filaExcel & ")"
                    hoja3.Cell(filaExcel + 4, 38).FormulaA1 = "=SUM(AL9:AL" & filaExcel & ")"


                '    hoja3.Cell(8, 39).Clear()
                '    hoja3.Cell(8, 40).Clear()
                '    hoja3.Cell(8, 41).Clear()
                '    hoja3.Cell(8, 42).Clear()
                '    hoja3.Cell(8, 43).Clear()
                '    hoja3.Cell(8, 44).Clear()
                'hoja3.Cell(8, 45).Clear()
                'hoja3.Cell(8, 46).Clear()
                'hoja3.Cell(8, 47).Clear()
                'hoja3.Cell(8, 48).Clear()
                'hoja3.Cell(8, 49).Clear()
                'hoja3.Cell(8, 50).Clear()
                'hoja3.Cell(8, 51).Clear()
                'hoja3.Cell(8, 52).Clear()

                limpiarCell(hoja3, 39) ', 1, dtgDatos.Rows.Count - 1)

               
                'hoja3.Range(8, 39, total, 45).Clear()



                '<<<<<<<<<<<<<<<Detalle>>>>>>>>>>>>>>>>>>
                filaExcel = 6
                filatmp = 9
                Dim cuenta, banco, clabe As String

                hoja4.Cell(4, 3).Style.Font.SetBold(True)
                hoja4.Cell(4, 3).Style.NumberFormat.Format = "@"
                hoja4.Cell(4, 3).Value = periodo

                For x As Integer = 0 To dtgDatos.Rows.Count - 1

                    hoja4.Cell(filaExcel, 6).Style.NumberFormat.Format = "@"
                    hoja4.Cell(filaExcel, 7).Style.NumberFormat.Format = "@"
                    hoja4.Range(filaExcel, 2, filaExcel, 9).Style.Font.SetBold(False)

                    Dim empleado As DataRow() = nConsulta("Select * from empleadosC where cCodigoEmpleado=" & dtgDatos.Rows(x).Cells(3).Value)
                    If empleado Is Nothing = False Then
                        cuenta = empleado(0).Item("NumCuenta")
                        clabe = empleado(0).Item("Clabe")
                        Dim bank As DataRow() = nConsulta("select * from bancos where iIdBanco =" & empleado(0).Item("fkiIdBanco"))
                        If bank Is Nothing = False Then
                            banco = bank(0).Item("cBANCO")
                        End If
                    End If



                    If inicio = x Then
                        contadorexcelbuqueinicial = filaExcel + x
                        nombrebuque = dtgDatos.Rows(x).Cells(12).Value
                    End If
                    If nombrebuque = dtgDatos.Rows(x).Cells(12).Value Then
                        hoja4.Cell(filaExcel, 2).Value = dtgDatos.Rows(x).Cells(12).Value
                        hoja4.Cell(filaExcel, 3).Value = dtgDatos.Rows(x).Cells(3).Value
                        hoja4.Cell(filaExcel, 4).Value = dtgDatos.Rows(x).Cells(4).Value
                        hoja4.Cell(filaExcel, 5).Value = banco
                        hoja4.Cell(filaExcel, 6).Value = clabe
                        hoja4.Cell(filaExcel, 7).Value = cuenta
                        hoja4.Cell(filaExcel, 8).FormulaA1 = "='OPERADORA ABORDO'!AI" & filatmp & "+'OPERADORA DESCANSO'!AI" & filatmp
                        hoja4.Cell(filaExcel, 9).FormulaA1 = "='NOMINA TOTAL'!Q" & filatmp + 4
                    Else
                        filatmp = filatmp + 2

                        nombrebuque = dtgDatos.Rows(x).Cells(12).Value
                        hoja4.Cell(filaExcel, 2).Value = dtgDatos.Rows(x).Cells(12).Value
                        hoja4.Cell(filaExcel, 3).Value = dtgDatos.Rows(x).Cells(3).Value
                        hoja4.Cell(filaExcel, 4).Value = dtgDatos.Rows(x).Cells(4).Value
                        hoja4.Cell(filaExcel, 5).Value = ""
                        hoja4.Cell(filaExcel, 6).Value = ""
                        hoja4.Cell(filaExcel, 7).Value = ""
                        hoja4.Cell(filaExcel, 8).FormulaA1 = "='OPERADORA ABORDO'!AI" & filatmp & "+'OPERADORA DESCANSO'!AI" & filatmp
                        'hoja4.Cell(filaExcel, 9).Value = "='NOMINA TOTAL'!Q"
                        hoja4.Cell(filaExcel, 9).FormulaA1 = "='NOMINA TOTAL'!Q" & filatmp + 4
                    End If



                    filaExcel = filaExcel + 1
                    filatmp = filatmp + 1


                Next x


                'Formulas
                hoja4.Range(filaExcel + 4, 8, filaExcel + 4, 11).Style.Font.SetBold(True)
                hoja4.Cell(filaExcel + 4, 8).FormulaA1 = "=SUM(H6:H" & filaExcel & ")"
                hoja4.Cell(filaExcel + 4, 9).FormulaA1 = "=SUM(I6:I" & filaExcel & ")"




                'Titulo
                Dim moment As Date = Date.Now()
                Dim month As Integer = moment.Month
                Dim year As Integer = moment.Year


                dialogo.FileName = "TMM " + MonthString(month).ToUpper + " " + year.ToString
                dialogo.Filter = "Archivos de Excel (*.xlsx)|*.xlsx"
                ''  dialogo.ShowDialog()

                If dialogo.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    ' OK button pressed
                    libro.SaveAs(dialogo.FileName)
                    libro = Nothing
                    MessageBox.Show("Archivo generado correctamente", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("No se guardo el archivo", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Try

    End Sub


    Public Sub limpiarCell(ByVal hoja As IXLWorksheet, ByVal celda As Integer) ', ByVal fila As Integer, ByVal filatotal As Integer)

        For x As Integer = celda To 200

            'For y As Integer = fila + 1 To filatotal + 20
            '    hoja.Cell(x, y).Clear()

            'Next y
            hoja.Cell(1, x).Clear()
            hoja.Cell(2, x).Clear()
            hoja.Cell(3, x).Clear()
            hoja.Cell(4, x).Clear()
            hoja.Cell(5, x).Clear()
            hoja.Cell(6, x).Clear()
            hoja.Cell(7, x).Clear()
            hoja.Cell(8, x).Clear()
            hoja.Cell(9, x).Clear()
            hoja.Cell(10, x).Clear()
            hoja.Cell(11, x).Clear()
            hoja.Cell(12, x).Clear()
            hoja.Cell(13, x).Clear()
            hoja.Cell(14, x).Clear()
        Next x
    End Sub

    Private Sub btnReporte_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReporte.Click
        Try

            Dim filaExcel As Integer = 0
            Dim dialogo As New SaveFileDialog()
            Dim periodo, fechadepago As String
            Dim mes As String
            Dim fechapagoletra() As String
            If dtgDatos.Rows.Count > 0 Then

                Dim rwPeriodo0 As DataRow() = nConsulta("Select * from periodos where iIdPeriodo=" & cboperiodo.SelectedValue)
                If rwPeriodo0 Is Nothing = False Then

                    periodo = MonthString(rwPeriodo0(0).Item("iMes")).ToUpper '& " DE " & (rwPeriodo0(0).Item("iEjercicio"))
                    mes = rwPeriodo0(0).Item("iMes")
                    fechapagoletra = rwPeriodo0(0).Item("dFechaFin").ToLongDateString().ToString.Split(" ")
                    fechadepago = rwPeriodo0(0).Item("dFechaFin")
                End If



                Dim ruta As String
                ruta = My.Application.Info.DirectoryPath() & "\Archivos\Reporte.xlsx"

                Dim book As New ClosedXML.Excel.XLWorkbook(ruta)


                Dim libro As New ClosedXML.Excel.XLWorkbook

                book.Worksheet(1).CopyTo(libro, periodo)
                book.Worksheet(2).CopyTo(libro, "DESGLOSE")
                book.Worksheet(3).CopyTo(libro, "RESUMEN")


                Dim hoja As IXLWorksheet = libro.Worksheets(0)
                Dim hoja2 As IXLWorksheet = libro.Worksheets(1)
                Dim hoja3 As IXLWorksheet = libro.Worksheets(2)




                '<<<<<<DESGLOCE>>>>>>>
                filaExcel = 2
                Dim nombrebuque As String
                Dim inicio As Integer = 0
                Dim contadorexcelbuqueinicial As Integer = 0
                Dim contadorexcelbuquefinal As Integer = 0
                Dim contadorexcelbuquefinalg As Integer = 0
                Dim total As Integer = dtgDatos.Rows.Count - 1
                Dim filatmp As Integer = 0

                For x As Integer = 0 To dtgDatos.Rows.Count - 1
                    If inicio = x Then
                        contadorexcelbuqueinicial = filaExcel + x
                        nombrebuque = dtgDatos.Rows(x).Cells(12).Value
                    End If
                    If nombrebuque = dtgDatos.Rows(x).Cells(12).Value Then

                        hoja2.Cell(filaExcel + x, 1).Value = fechadepago 'FECHA DE PAGO
                        hoja2.Cell(filaExcel + x, 2).Value = dtgDatos.Rows(x).Cells(3).Value
                        hoja2.Cell(filaExcel + x, 3).Value = dtgDatos.Rows(x).Cells(4).Value
                        hoja2.Cell(filaExcel + x, 4).Value = dtgDatos.Rows(x).Cells(6).Value
                        hoja2.Cell(filaExcel + x, 5).Value = dtgDatos.Rows(x).Cells(11).FormattedValue
                        hoja2.Cell(filaExcel + x, 6).Value = dtgDatos.Rows(x).Cells(18).Value
                        hoja2.Cell(filaExcel + x, 7).Value = dtgDatos.Rows(x).Cells(12).FormattedValue
                        hoja2.Cell(filaExcel + x, 8).Value = dtgDatos.Rows(x).Cells(15).Value
                        hoja2.Cell(filaExcel + x, 9).Value = CInt(dtgDatos.Rows(x).Cells(22).Value) + CInt(dtgDatos.Rows(x).Cells(23).Value) 'TE Gravado
                        hoja2.Cell(filaExcel + x, 10).Value = dtgDatos.Rows(x).Cells(24).Value
                        hoja2.Cell(filaExcel + x, 11).Value = dtgDatos.Rows(x).Cells(25).Value
                        hoja2.Cell(filaExcel + x, 12).Value = dtgDatos.Rows(x).Cells(26).Value
                        hoja2.Cell(filaExcel + x, 13).Value = dtgDatos.Rows(x).Cells(29).Value
                        hoja2.Cell(filaExcel + x, 14).Value = dtgDatos.Rows(x).Cells(32).Value
                        hoja2.Cell(filaExcel + x, 15).Value = dtgDatos.Rows(x).Cells(33).Value
                        hoja2.Cell(filaExcel + x, 16).Value = dtgDatos.Rows(x).Cells(56).Value
                        If dtgDatos.Rows(x).Cells(46).Value <> "" Then
                            hoja2.Cell(filaExcel + x, 17).Value = dtgDatos.Rows(x).Cells(46).Value * 2%
                        Else
                            hoja2.Cell(filaExcel + x, 17).Value = "0"
                        End If
                        If dtgDatos.Rows(x).Cells(56).Value <> "" Then
                            hoja2.Cell(filaExcel + x, 18).Value = dtgDatos.Rows(x).Cells(56).Value * 2%
                        Else
                            hoja2.Cell(filaExcel + x, 18).Value = "0"
                        End If
                        'hoja2.Cell(filaExcel + x, 17).Value = IIf(dtgDatos.Rows(x).Cells(46).Value <> "", dtgDatos.Rows(x).Cells(46).Value * 2%, "0") 'COMISION OPERADORA (Neto_pagar*2%)
                        'hoja2.Cell(filaExcel + x, 18).Value = IIf(dtgDatos.Rows(x).Cells(56).Value <> "", dtgDatos.Rows(x).Cells(56).Value * 2%, "0") 'COMISION COMPLEMENTE
                        hoja2.Cell(filaExcel + x, 19).Value = dtgDatos.Rows(x).Cells(45).Value 'Subsidio
                        hoja2.Cell(filaExcel + x, 20).Value = dtgDatos.Rows(x).Cells(49).Value
                        hoja2.Cell(filaExcel + x, 21).Value = dtgDatos.Rows(x).Cells(50).Value
                        hoja2.Cell(filaExcel + x, 22).Value = dtgDatos.Rows(x).Cells(51).Value
                        hoja2.Cell(filaExcel + x, 23).Value = dtgDatos.Rows(x).Cells(52).Value

                        hoja2.Cell(filaExcel + x, 24).FormulaA1 = "=SUMA(O" & filaExcel + x & ":W" & filaExcel + x & ")"
                        hoja2.Cell(filaExcel + x, 25).FormulaA1 = "=X" & filaExcel + x & "*16%"
                        hoja2.Cell(filaExcel + x, 26).FormulaA1 = "=X" & filaExcel & "+Y" & filaExcel + x




                    Else
                        filatmp = filatmp + 1

                        contadorexcelbuquefinal = filaExcel + x - 1
                        contadorexcelbuquefinalg = contadorexcelbuquefinal
                        hoja2.Cell(filaExcel + x, 7).Value = "SUMA " + nombrebuque
                        hoja2.Cell(filaExcel + x, 8).FormulaA1 = "=SUM(H" & contadorexcelbuqueinicial & ":H" & contadorexcelbuquefinal & ")"
                        hoja2.Cell(filaExcel + x, 9).FormulaA1 = "=SUM(I" & contadorexcelbuqueinicial & ":I" & contadorexcelbuquefinal & ")"
                        hoja2.Cell(filaExcel + x, 10).FormulaA1 = "=SUM(J" & contadorexcelbuqueinicial & ":J" & contadorexcelbuquefinal & ")"
                        hoja2.Cell(filaExcel + x, 11).FormulaA1 = "=SUM(K" & contadorexcelbuqueinicial & ":K" & contadorexcelbuquefinal & ")"
                        hoja2.Cell(filaExcel + x, 12).FormulaA1 = "=SUM(L" & contadorexcelbuqueinicial & ":L" & contadorexcelbuquefinal & ")"
                        hoja2.Cell(filaExcel + x, 13).FormulaA1 = "=SUM(M" & contadorexcelbuqueinicial & ":M" & contadorexcelbuquefinal & ")"
                        hoja2.Cell(filaExcel + x, 14).FormulaA1 = "=SUM(N" & contadorexcelbuqueinicial & ":N" & contadorexcelbuquefinal & ")"
                        hoja2.Cell(filaExcel + x, 15).FormulaA1 = "=SUM(O" & contadorexcelbuqueinicial & ":O" & contadorexcelbuquefinal & ")"
                        hoja2.Cell(filaExcel + x, 16).FormulaA1 = "=SUM(P" & contadorexcelbuqueinicial & ":P" & contadorexcelbuquefinal & ")"
                        hoja2.Cell(filaExcel + x, 17).FormulaA1 = "=SUM(Q" & contadorexcelbuqueinicial & ":Q" & contadorexcelbuquefinal & ")"
                        hoja2.Cell(filaExcel + x, 18).FormulaA1 = "=SUM(R" & contadorexcelbuqueinicial & ":R" & contadorexcelbuquefinal & ")"
                        hoja2.Cell(filaExcel + x, 19).FormulaA1 = "=SUM(S" & contadorexcelbuqueinicial & ":S" & contadorexcelbuquefinal & ")"
                        hoja2.Cell(filaExcel + x, 20).FormulaA1 = "=SUM(T" & contadorexcelbuqueinicial & ":T" & contadorexcelbuquefinal & ")"
                        hoja2.Cell(filaExcel + x, 21).FormulaA1 = "=SUM(U" & contadorexcelbuqueinicial & ":U" & contadorexcelbuquefinal & ")"
                        hoja2.Cell(filaExcel + x, 22).FormulaA1 = "=SUM(V" & contadorexcelbuqueinicial & ":V" & contadorexcelbuquefinal & ")"
                        hoja2.Cell(filaExcel + x, 23).FormulaA1 = "=SUM(W" & contadorexcelbuqueinicial & ":W" & contadorexcelbuquefinal & ")"
                        hoja2.Cell(filaExcel + x, 24).FormulaA1 = "=SUM(X" & contadorexcelbuqueinicial & ":X" & contadorexcelbuquefinal & ")"
                        hoja2.Cell(filaExcel + x, 25).FormulaA1 = "=SUM(Y" & contadorexcelbuqueinicial & ":Y" & contadorexcelbuquefinal & ")"
                        hoja2.Cell(filaExcel + x, 26).FormulaA1 = "=SUM(Z" & contadorexcelbuqueinicial & ":Z" & contadorexcelbuquefinal & ")"
                        hoja2.Range(filaExcel + x, 7, filaExcel + x, 26).Style.Font.SetBold(True)

                        '<<<<Mes>>>
                        hoja.Cell(5, 2).Style.NumberFormat.Format = "@"
                        hoja.Cell(5, 2).Value = fechapagoletra(1) & " " & fechapagoletra(2) & " " & fechapagoletra(3)
                        hoja.Cell(16, 2).Value = MonthString(mes - 1).ToUpper & " ADICIONALES"
                        hoja.Cell(24, 7).Value = MonthString(mes - 1).ToUpper
                        hoja.Cell(24, 18).Value = MonthString(mes - 1).ToUpper & " ADICIONALES"
                        hoja.Cell(24, 18).Style.Font.SetBold(True)
                        Select Case nombrebuque
                            Case "CEDROS", "ISLA CEDROS"
                                hoja.Cell(5, 4).FormulaA1 = "=DESGLOSE!H" & contadorexcelbuquefinal + 1
                                hoja.Cell(5, 5).FormulaA1 = "=DESGLOSE!I" & contadorexcelbuquefinal + 1
                                hoja.Cell(5, 6).FormulaA1 = "=DESGLOSE!J" & contadorexcelbuquefinal + 1
                                hoja.Cell(5, 7).FormulaA1 = "=DESGLOSE!K" & contadorexcelbuquefinal + 1
                                hoja.Cell(5, 8).FormulaA1 = "=DESGLOSE!L" & contadorexcelbuquefinal + 1
                                hoja.Cell(5, 9).FormulaA1 = "=DESGLOSE!M" & contadorexcelbuquefinal + 1
                                hoja.Cell(5, 10).FormulaA1 = "=DESGLOSE!N" & contadorexcelbuquefinal + 1
                                hoja.Cell(5, 11).FormulaA1 = "=DESGLOSE!O" & contadorexcelbuquefinal + 1
                                hoja.Cell(5, 12).FormulaA1 = "=DESGLOSE!P" & contadorexcelbuquefinal + 1
                                hoja.Cell(5, 13).FormulaA1 = "=DESGLOSE!Q" & contadorexcelbuquefinal + 1
                                hoja.Cell(5, 14).FormulaA1 = "=DESGLOSE!R" & contadorexcelbuquefinal + 1
                                hoja.Cell(5, 15).FormulaA1 = "=DESGLOSE!S" & contadorexcelbuquefinal + 1
                                hoja.Cell(5, 16).FormulaA1 = "=DESGLOSE!T" & contadorexcelbuquefinal + 1
                                hoja.Cell(5, 17).FormulaA1 = "=DESGLOSE!U" & contadorexcelbuquefinal + 1
                                hoja.Cell(5, 18).FormulaA1 = "=DESGLOSE!V" & contadorexcelbuquefinal + 1
                                hoja.Cell(5, 19).FormulaA1 = "=DESGLOSE!W" & contadorexcelbuquefinal + 1





                            Case "ISLA SAN JOSE"
                                hoja.Cell(6, 4).FormulaA1 = "=DESGLOSE!H" & contadorexcelbuquefinal + 1
                                hoja.Cell(6, 5).FormulaA1 = "=DESGLOSE!I" & contadorexcelbuquefinal + 1
                                hoja.Cell(6, 6).FormulaA1 = "=DESGLOSE!J" & contadorexcelbuquefinal + 1
                                hoja.Cell(6, 7).FormulaA1 = "=DESGLOSE!K" & contadorexcelbuquefinal + 1
                                hoja.Cell(6, 8).FormulaA1 = "=DESGLOSE!L" & contadorexcelbuquefinal + 1
                                hoja.Cell(6, 9).FormulaA1 = "=DESGLOSE!M" & contadorexcelbuquefinal + 1
                                hoja.Cell(6, 10).FormulaA1 = "=DESGLOSE!N" & contadorexcelbuquefinal + 1
                                hoja.Cell(6, 11).FormulaA1 = "=DESGLOSE!O" & contadorexcelbuquefinal + 1
                                hoja.Cell(6, 12).FormulaA1 = "=DESGLOSE!P" & contadorexcelbuquefinal + 1
                                hoja.Cell(6, 13).FormulaA1 = "=DESGLOSE!Q" & contadorexcelbuquefinal + 1
                                hoja.Cell(6, 14).FormulaA1 = "=DESGLOSE!R" & contadorexcelbuquefinal + 1
                                hoja.Cell(6, 15).FormulaA1 = "=DESGLOSE!S" & contadorexcelbuquefinal + 1
                                hoja.Cell(6, 16).FormulaA1 = "=DESGLOSE!T" & contadorexcelbuquefinal + 1
                                hoja.Cell(6, 17).FormulaA1 = "=DESGLOSE!U" & contadorexcelbuquefinal + 1
                                hoja.Cell(6, 18).FormulaA1 = "=DESGLOSE!V" & contadorexcelbuquefinal + 1
                                hoja.Cell(6, 19).FormulaA1 = "=DESGLOSE!W" & contadorexcelbuquefinal + 1


                            Case "ISLA GRANDE"
                                hoja.Cell(7, 4).FormulaA1 = "=DESGLOSE!H" & contadorexcelbuquefinal + 1
                                hoja.Cell(7, 5).FormulaA1 = "=DESGLOSE!I" & contadorexcelbuquefinal + 1
                                hoja.Cell(7, 6).FormulaA1 = "=DESGLOSE!J" & contadorexcelbuquefinal + 1
                                hoja.Cell(7, 7).FormulaA1 = "=DESGLOSE!K" & contadorexcelbuquefinal + 1
                                hoja.Cell(7, 8).FormulaA1 = "=DESGLOSE!L" & contadorexcelbuquefinal + 1
                                hoja.Cell(7, 9).FormulaA1 = "=DESGLOSE!M" & contadorexcelbuquefinal + 1
                                hoja.Cell(7, 10).FormulaA1 = "=DESGLOSE!N" & contadorexcelbuquefinal + 1
                                hoja.Cell(7, 11).FormulaA1 = "=DESGLOSE!O" & contadorexcelbuquefinal + 1
                                hoja.Cell(7, 12).FormulaA1 = "=DESGLOSE!P" & contadorexcelbuquefinal + 1
                                hoja.Cell(7, 13).FormulaA1 = "=DESGLOSE!Q" & contadorexcelbuquefinal + 1
                                hoja.Cell(7, 14).FormulaA1 = "=DESGLOSE!R" & contadorexcelbuquefinal + 1
                                hoja.Cell(7, 15).FormulaA1 = "=DESGLOSE!S" & contadorexcelbuquefinal + 1
                                hoja.Cell(7, 16).FormulaA1 = "=DESGLOSE!T" & contadorexcelbuquefinal + 1
                                hoja.Cell(7, 17).FormulaA1 = "=DESGLOSE!U" & contadorexcelbuquefinal + 1
                                hoja.Cell(7, 18).FormulaA1 = "=DESGLOSE!V" & contadorexcelbuquefinal + 1
                                hoja.Cell(7, 19).FormulaA1 = "=DESGLOSE!W" & contadorexcelbuquefinal + 1




                            Case "ISLA MIRAMAR"
                                hoja.Cell(8, 4).FormulaA1 = "=DESGLOSE!H" & contadorexcelbuquefinal + 1
                                hoja.Cell(8, 5).FormulaA1 = "=DESGLOSE!I" & contadorexcelbuquefinal + 1
                                hoja.Cell(8, 6).FormulaA1 = "=DESGLOSE!J" & contadorexcelbuquefinal + 1
                                hoja.Cell(8, 7).FormulaA1 = "=DESGLOSE!K" & contadorexcelbuquefinal + 1
                                hoja.Cell(8, 8).FormulaA1 = "=DESGLOSE!L" & contadorexcelbuquefinal + 1
                                hoja.Cell(8, 9).FormulaA1 = "=DESGLOSE!M" & contadorexcelbuquefinal + 1
                                hoja.Cell(8, 10).FormulaA1 = "=DESGLOSE!N" & contadorexcelbuquefinal + 1
                                hoja.Cell(8, 11).FormulaA1 = "=DESGLOSE!O" & contadorexcelbuquefinal + 1
                                hoja.Cell(8, 12).FormulaA1 = "=DESGLOSE!P" & contadorexcelbuquefinal + 1
                                hoja.Cell(8, 13).FormulaA1 = "=DESGLOSE!Q" & contadorexcelbuquefinal + 1
                                hoja.Cell(8, 14).FormulaA1 = "=DESGLOSE!R" & contadorexcelbuquefinal + 1
                                hoja.Cell(8, 15).FormulaA1 = "=DESGLOSE!S" & contadorexcelbuquefinal + 1
                                hoja.Cell(8, 16).FormulaA1 = "=DESGLOSE!T" & contadorexcelbuquefinal + 1
                                hoja.Cell(8, 17).FormulaA1 = "=DESGLOSE!U" & contadorexcelbuquefinal + 1
                                hoja.Cell(8, 18).FormulaA1 = "=DESGLOSE!V" & contadorexcelbuquefinal + 1
                                hoja.Cell(8, 19).FormulaA1 = "=DESGLOSE!W" & contadorexcelbuquefinal + 1


                            Case "ISLA MONSERRAT"
                                hoja.Cell(9, 4).FormulaA1 = "=DESGLOSE!H" & contadorexcelbuquefinal + 1
                                hoja.Cell(9, 5).FormulaA1 = "=DESGLOSE!I" & contadorexcelbuquefinal + 1
                                hoja.Cell(9, 6).FormulaA1 = "=DESGLOSE!J" & contadorexcelbuquefinal + 1
                                hoja.Cell(9, 7).FormulaA1 = "=DESGLOSE!K" & contadorexcelbuquefinal + 1
                                hoja.Cell(9, 8).FormulaA1 = "=DESGLOSE!L" & contadorexcelbuquefinal + 1
                                hoja.Cell(9, 9).FormulaA1 = "=DESGLOSE!M" & contadorexcelbuquefinal + 1
                                hoja.Cell(9, 10).FormulaA1 = "=DESGLOSE!N" & contadorexcelbuquefinal + 1
                                hoja.Cell(9, 11).FormulaA1 = "=DESGLOSE!O" & contadorexcelbuquefinal + 1
                                hoja.Cell(9, 12).FormulaA1 = "=DESGLOSE!P" & contadorexcelbuquefinal + 1
                                hoja.Cell(9, 13).FormulaA1 = "=DESGLOSE!Q" & contadorexcelbuquefinal + 1
                                hoja.Cell(9, 14).FormulaA1 = "=DESGLOSE!R" & contadorexcelbuquefinal + 1
                                hoja.Cell(9, 15).FormulaA1 = "=DESGLOSE!S" & contadorexcelbuquefinal + 1
                                hoja.Cell(9, 16).FormulaA1 = "=DESGLOSE!T" & contadorexcelbuquefinal + 1
                                hoja.Cell(9, 17).FormulaA1 = "=DESGLOSE!U" & contadorexcelbuquefinal + 1
                                hoja.Cell(9, 18).FormulaA1 = "=DESGLOSE!V" & contadorexcelbuquefinal + 1
                                hoja.Cell(9, 19).FormulaA1 = "=DESGLOSE!W" & contadorexcelbuquefinal + 1



                            Case "ISLA BLANCA"
                                hoja.Cell(10, 4).FormulaA1 = "=DESGLOSE!H" & contadorexcelbuquefinal + 1
                                hoja.Cell(10, 5).FormulaA1 = "=DESGLOSE!I" & contadorexcelbuquefinal + 1
                                hoja.Cell(10, 6).FormulaA1 = "=DESGLOSE!J" & contadorexcelbuquefinal + 1
                                hoja.Cell(10, 7).FormulaA1 = "=DESGLOSE!K" & contadorexcelbuquefinal + 1
                                hoja.Cell(10, 8).FormulaA1 = "=DESGLOSE!L" & contadorexcelbuquefinal + 1
                                hoja.Cell(10, 9).FormulaA1 = "=DESGLOSE!M" & contadorexcelbuquefinal + 1
                                hoja.Cell(10, 10).FormulaA1 = "=DESGLOSE!N" & contadorexcelbuquefinal + 1
                                hoja.Cell(10, 11).FormulaA1 = "=DESGLOSE!O" & contadorexcelbuquefinal + 1
                                hoja.Cell(10, 12).FormulaA1 = "=DESGLOSE!P" & contadorexcelbuquefinal + 1
                                hoja.Cell(10, 13).FormulaA1 = "=DESGLOSE!Q" & contadorexcelbuquefinal + 1
                                hoja.Cell(10, 14).FormulaA1 = "=DESGLOSE!R" & contadorexcelbuquefinal + 1
                                hoja.Cell(10, 15).FormulaA1 = "=DESGLOSE!S" & contadorexcelbuquefinal + 1
                                hoja.Cell(10, 16).FormulaA1 = "=DESGLOSE!T" & contadorexcelbuquefinal + 1
                                hoja.Cell(10, 17).FormulaA1 = "=DESGLOSE!U" & contadorexcelbuquefinal + 1
                                hoja.Cell(10, 18).FormulaA1 = "=DESGLOSE!V" & contadorexcelbuquefinal + 1
                                hoja.Cell(10, 19).FormulaA1 = "=DESGLOSE!W" & contadorexcelbuquefinal + 1



                            Case "ISLA CIARI"
                                hoja.Cell(11, 4).FormulaA1 = "=DESGLOSE!H" & contadorexcelbuquefinal + 1
                                hoja.Cell(11, 5).FormulaA1 = "=DESGLOSE!I" & contadorexcelbuquefinal + 1
                                hoja.Cell(11, 6).FormulaA1 = "=DESGLOSE!J" & contadorexcelbuquefinal + 1
                                hoja.Cell(11, 7).FormulaA1 = "=DESGLOSE!K" & contadorexcelbuquefinal + 1
                                hoja.Cell(11, 8).FormulaA1 = "=DESGLOSE!L" & contadorexcelbuquefinal + 1
                                hoja.Cell(11, 9).FormulaA1 = "=DESGLOSE!M" & contadorexcelbuquefinal + 1
                                hoja.Cell(11, 10).FormulaA1 = "=DESGLOSE!N" & contadorexcelbuquefinal + 1
                                hoja.Cell(11, 11).FormulaA1 = "=DESGLOSE!O" & contadorexcelbuquefinal + 1
                                hoja.Cell(11, 12).FormulaA1 = "=DESGLOSE!P" & contadorexcelbuquefinal + 1
                                hoja.Cell(11, 13).FormulaA1 = "=DESGLOSE!Q" & contadorexcelbuquefinal + 1
                                hoja.Cell(11, 14).FormulaA1 = "=DESGLOSE!R" & contadorexcelbuquefinal + 1
                                hoja.Cell(11, 15).FormulaA1 = "=DESGLOSE!S" & contadorexcelbuquefinal + 1
                                hoja.Cell(11, 16).FormulaA1 = "=DESGLOSE!T" & contadorexcelbuquefinal + 1
                                hoja.Cell(11, 17).FormulaA1 = "=DESGLOSE!U" & contadorexcelbuquefinal + 1
                                hoja.Cell(11, 18).FormulaA1 = "=DESGLOSE!V" & contadorexcelbuquefinal + 1
                                hoja.Cell(11, 19).FormulaA1 = "=DESGLOSE!W" & contadorexcelbuquefinal + 1



                            Case "ISLA JANITZIO"
                                hoja.Cell(12, 4).FormulaA1 = "=DESGLOSE!H" & contadorexcelbuquefinal + 1
                                hoja.Cell(12, 5).FormulaA1 = "=DESGLOSE!I" & contadorexcelbuquefinal + 1
                                hoja.Cell(12, 6).FormulaA1 = "=DESGLOSE!J" & contadorexcelbuquefinal + 1
                                hoja.Cell(12, 7).FormulaA1 = "=DESGLOSE!K" & contadorexcelbuquefinal + 1
                                hoja.Cell(12, 8).FormulaA1 = "=DESGLOSE!L" & contadorexcelbuquefinal + 1
                                hoja.Cell(12, 9).FormulaA1 = "=DESGLOSE!M" & contadorexcelbuquefinal + 1
                                hoja.Cell(12, 10).FormulaA1 = "=DESGLOSE!N" & contadorexcelbuquefinal + 1
                                hoja.Cell(12, 11).FormulaA1 = "=DESGLOSE!O" & contadorexcelbuquefinal + 1
                                hoja.Cell(12, 12).FormulaA1 = "=DESGLOSE!P" & contadorexcelbuquefinal + 1
                                hoja.Cell(12, 13).FormulaA1 = "=DESGLOSE!Q" & contadorexcelbuquefinal + 1
                                hoja.Cell(12, 14).FormulaA1 = "=DESGLOSE!R" & contadorexcelbuquefinal + 1
                                hoja.Cell(12, 15).FormulaA1 = "=DESGLOSE!S" & contadorexcelbuquefinal + 1
                                hoja.Cell(12, 16).FormulaA1 = "=DESGLOSE!T" & contadorexcelbuquefinal + 1
                                hoja.Cell(12, 17).FormulaA1 = "=DESGLOSE!U" & contadorexcelbuquefinal + 1
                                hoja.Cell(12, 18).FormulaA1 = "=DESGLOSE!V" & contadorexcelbuquefinal + 1
                                hoja.Cell(12, 19).FormulaA1 = "=DESGLOSE!W" & contadorexcelbuquefinal + 1


                            Case "ISLA SAN GABRIEL"
                                hoja.Cell(13, 4).FormulaA1 = "=DESGLOSE!H" & contadorexcelbuquefinal + 1
                                hoja.Cell(13, 5).FormulaA1 = "=DESGLOSE!I" & contadorexcelbuquefinal + 1
                                hoja.Cell(13, 6).FormulaA1 = "=DESGLOSE!J" & contadorexcelbuquefinal + 1
                                hoja.Cell(13, 7).FormulaA1 = "=DESGLOSE!K" & contadorexcelbuquefinal + 1
                                hoja.Cell(13, 8).FormulaA1 = "=DESGLOSE!L" & contadorexcelbuquefinal + 1
                                hoja.Cell(13, 9).FormulaA1 = "=DESGLOSE!M" & contadorexcelbuquefinal + 1
                                hoja.Cell(13, 10).FormulaA1 = "=DESGLOSE!N" & contadorexcelbuquefinal + 1
                                hoja.Cell(13, 11).FormulaA1 = "=DESGLOSE!O" & contadorexcelbuquefinal + 1
                                hoja.Cell(13, 12).FormulaA1 = "=DESGLOSE!P" & contadorexcelbuquefinal + 1
                                hoja.Cell(13, 13).FormulaA1 = "=DESGLOSE!Q" & contadorexcelbuquefinal + 1
                                hoja.Cell(13, 14).FormulaA1 = "=DESGLOSE!R" & contadorexcelbuquefinal + 1
                                hoja.Cell(13, 15).FormulaA1 = "=DESGLOSE!S" & contadorexcelbuquefinal + 1
                                hoja.Cell(13, 16).FormulaA1 = "=DESGLOSE!T" & contadorexcelbuquefinal + 1
                                hoja.Cell(13, 17).FormulaA1 = "=DESGLOSE!U" & contadorexcelbuquefinal + 1
                                hoja.Cell(13, 18).FormulaA1 = "=DESGLOSE!V" & contadorexcelbuquefinal + 1
                                hoja.Cell(13, 19).FormulaA1 = "=DESGLOSE!W" & contadorexcelbuquefinal + 1



                            Case "AMARRADOS"
                                hoja.Cell(14, 4).FormulaA1 = "=DESGLOSE!H" & contadorexcelbuquefinal + 1
                                hoja.Cell(14, 5).FormulaA1 = "=DESGLOSE!I" & contadorexcelbuquefinal + 1
                                hoja.Cell(14, 6).FormulaA1 = "=DESGLOSE!J" & contadorexcelbuquefinal + 1
                                hoja.Cell(14, 7).FormulaA1 = "=DESGLOSE!K" & contadorexcelbuquefinal + 1
                                hoja.Cell(14, 8).FormulaA1 = "=DESGLOSE!L" & contadorexcelbuquefinal + 1
                                hoja.Cell(14, 9).FormulaA1 = "=DESGLOSE!M" & contadorexcelbuquefinal + 1
                                hoja.Cell(14, 10).FormulaA1 = "=DESGLOSE!N" & contadorexcelbuquefinal + 1
                                hoja.Cell(14, 11).FormulaA1 = "=DESGLOSE!O" & contadorexcelbuquefinal + 1
                                hoja.Cell(14, 12).FormulaA1 = "=DESGLOSE!P" & contadorexcelbuquefinal + 1
                                hoja.Cell(14, 13).FormulaA1 = "=DESGLOSE!Q" & contadorexcelbuquefinal + 1
                                hoja.Cell(14, 14).FormulaA1 = "=DESGLOSE!R" & contadorexcelbuquefinal + 1
                                hoja.Cell(14, 15).FormulaA1 = "=DESGLOSE!S" & contadorexcelbuquefinal + 1
                                hoja.Cell(14, 16).FormulaA1 = "=DESGLOSE!T" & contadorexcelbuquefinal + 1
                                hoja.Cell(14, 17).FormulaA1 = "=DESGLOSE!U" & contadorexcelbuquefinal + 1
                                hoja.Cell(14, 18).FormulaA1 = "=DESGLOSE!V" & contadorexcelbuquefinal + 1
                                hoja.Cell(14, 19).FormulaA1 = "=DESGLOSE!W" & contadorexcelbuquefinal + 1


                        End Select


                        nombrebuque = dtgDatos.Rows(x).Cells(12).Value
                        filaExcel = filaExcel + 1
                        contadorexcelbuqueinicial = filaExcel + x
                        contadorexcelbuquefinal = 0


                        hoja2.Cell(filaExcel + x, 1).Value = fechadepago 'FECHA DE PAGO
                        hoja2.Cell(filaExcel + x, 2).Value = dtgDatos.Rows(x).Cells(3).Value
                        hoja2.Cell(filaExcel + x, 3).Value = dtgDatos.Rows(x).Cells(4).Value
                        hoja2.Cell(filaExcel + x, 4).Value = dtgDatos.Rows(x).Cells(6).Value
                        hoja2.Cell(filaExcel + x, 5).Value = dtgDatos.Rows(x).Cells(11).FormattedValue
                        hoja2.Cell(filaExcel + x, 6).Value = dtgDatos.Rows(x).Cells(18).Value
                        hoja2.Cell(filaExcel + x, 7).Value = dtgDatos.Rows(x).Cells(12).FormattedValue
                        hoja2.Cell(filaExcel + x, 8).Value = dtgDatos.Rows(x).Cells(15).Value
                        hoja2.Cell(filaExcel + x, 9).Value = CInt(dtgDatos.Rows(x).Cells(22).Value) + CInt(dtgDatos.Rows(x).Cells(23).Value) 'TE Gravado
                        hoja2.Cell(filaExcel + x, 10).Value = dtgDatos.Rows(x).Cells(24).Value
                        hoja2.Cell(filaExcel + x, 11).Value = dtgDatos.Rows(x).Cells(25).Value
                        hoja2.Cell(filaExcel + x, 12).Value = dtgDatos.Rows(x).Cells(26).Value
                        hoja2.Cell(filaExcel + x, 13).Value = dtgDatos.Rows(x).Cells(29).Value
                        hoja2.Cell(filaExcel + x, 14).Value = dtgDatos.Rows(x).Cells(32).Value
                        hoja2.Cell(filaExcel + x, 15).Value = dtgDatos.Rows(x).Cells(33).Value
                        hoja2.Cell(filaExcel + x, 16).Value = dtgDatos.Rows(x).Cells(56).Value
                        ' hoja2.Cell(filaExcel + x, 17).Value = dtgDatos.Rows(x).Cells(46).Value * 2% 'IIf(dtgDatos.Rows(x).Cells(46).Value <> "", dtgDatos.Rows(x).Cells(46).Value * 2%, 0) 'COMISION OPERADORA (Neto_pagar*2%)
                        If dtgDatos.Rows(x).Cells(46).Value <> "" Then
                            hoja2.Cell(filaExcel + x, 17).Value = dtgDatos.Rows(x).Cells(46).Value * 2%
                        Else
                            hoja2.Cell(filaExcel + x, 17).Value = "0"
                        End If
                        If dtgDatos.Rows(x).Cells(56).Value <> "" Then
                            hoja2.Cell(filaExcel + x, 18).Value = dtgDatos.Rows(x).Cells(56).Value * 2%
                        Else
                            hoja2.Cell(filaExcel + x, 18).Value = "0"
                        End If
                        'IIf(dtgDatos.Rows(x).Cells(56).Value <> "", dtgDatos.Rows(x).Cells(56).Value * 2%, 0) 'COMISION COMPLEMENTE
                        hoja2.Cell(filaExcel + x, 19).Value = dtgDatos.Rows(x).Cells(45).Value 'Subsidio
                        hoja2.Cell(filaExcel + x, 20).Value = dtgDatos.Rows(x).Cells(49).Value
                        hoja2.Cell(filaExcel + x, 21).Value = dtgDatos.Rows(x).Cells(50).Value
                        hoja2.Cell(filaExcel + x, 22).Value = dtgDatos.Rows(x).Cells(51).Value
                        hoja2.Cell(filaExcel + x, 23).Value = dtgDatos.Rows(x).Cells(52).Value
                        hoja2.Cell(filaExcel + x, 24).FormulaA1 = "=SUMA(O" & filaExcel + x & ":W" & filaExcel + x & ")"
                        hoja2.Cell(filaExcel + x, 25).FormulaA1 = "=X" & filaExcel + x & "*16%"
                        hoja2.Cell(filaExcel + x, 26).FormulaA1 = "=X" & filaExcel & "+Y" & filaExcel + x

                    End If


                Next x
                filaExcel = filaExcel + 1
                contadorexcelbuquefinal = filaExcel + total - 1
                hoja2.Cell(filaExcel + total, 7).Value = "SUMA " + nombrebuque
                hoja2.Cell(filaExcel + total, 8).FormulaA1 = "=SUM(H" & contadorexcelbuqueinicial & ":H" & contadorexcelbuquefinal & ")"
                hoja2.Cell(filaExcel + total, 9).FormulaA1 = "=SUM(I" & contadorexcelbuqueinicial & ":I" & contadorexcelbuquefinal & ")"
                hoja2.Cell(filaExcel + total, 10).FormulaA1 = "=SUM(J" & contadorexcelbuqueinicial & ":J" & contadorexcelbuquefinal & ")"
                hoja2.Cell(filaExcel + total, 11).FormulaA1 = "=SUM(K" & contadorexcelbuqueinicial & ":K" & contadorexcelbuquefinal & ")"
                hoja2.Cell(filaExcel + total, 12).FormulaA1 = "=SUM(L" & contadorexcelbuqueinicial & ":L" & contadorexcelbuquefinal & ")"
                hoja2.Cell(filaExcel + total, 13).FormulaA1 = "=SUM(M" & contadorexcelbuqueinicial & ":M" & contadorexcelbuquefinal & ")"
                hoja2.Cell(filaExcel + total, 14).FormulaA1 = "=SUM(N" & contadorexcelbuqueinicial & ":N" & contadorexcelbuquefinal & ")"
                hoja2.Cell(filaExcel + total, 15).FormulaA1 = "=SUM(O" & contadorexcelbuqueinicial & ":O" & contadorexcelbuquefinal & ")"
                hoja2.Cell(filaExcel + total, 16).FormulaA1 = "=SUM(P" & contadorexcelbuqueinicial & ":P" & contadorexcelbuquefinal & ")"
                hoja2.Cell(filaExcel + total, 17).FormulaA1 = "=SUM(Q" & contadorexcelbuqueinicial & ":Q" & contadorexcelbuquefinal & ")"
                hoja2.Cell(filaExcel + total, 18).FormulaA1 = "=SUM(R" & contadorexcelbuqueinicial & ":R" & contadorexcelbuquefinal & ")"
                hoja2.Cell(filaExcel + total, 19).FormulaA1 = "=SUM(S" & contadorexcelbuqueinicial & ":S" & contadorexcelbuquefinal & ")"
                hoja2.Cell(filaExcel + total, 20).FormulaA1 = "=SUM(T" & contadorexcelbuqueinicial & ":T" & contadorexcelbuquefinal & ")"
                hoja2.Cell(filaExcel + total, 21).FormulaA1 = "=SUM(U" & contadorexcelbuqueinicial & ":U" & contadorexcelbuquefinal & ")"
                hoja2.Cell(filaExcel + total, 22).FormulaA1 = "=SUM(V" & contadorexcelbuqueinicial & ":V" & contadorexcelbuquefinal & ")"
                hoja2.Cell(filaExcel + total, 23).FormulaA1 = "=SUM(W" & contadorexcelbuqueinicial & ":W" & contadorexcelbuquefinal & ")"
                hoja2.Cell(filaExcel + total, 24).FormulaA1 = "=SUM(X" & contadorexcelbuqueinicial & ":X" & contadorexcelbuquefinal & ")"
                hoja2.Cell(filaExcel + total, 25).FormulaA1 = "=SUM(Y" & contadorexcelbuqueinicial & ":Y" & contadorexcelbuquefinal & ")"
                hoja2.Cell(filaExcel + total, 26).FormulaA1 = "=SUM(Z" & contadorexcelbuqueinicial & ":Z" & contadorexcelbuquefinal & ")"
                hoja2.Range(filaExcel + total, 7, filaExcel + total, 26).Style.Font.SetBold(True)



                '<<<<<<<<<<RESUMEN>>>>>>>>>
                filaExcel = 2

                'For x As Integer = 0 To dtgDatos.Rows.Count - 1
                hoja3.Cell(1, 2).Style.NumberFormat.Format = "@"
                hoja3.Cell(1, 2).Value = fechapagoletra(1) & " " & fechapagoletra(2) & " " & fechapagoletra(3)
                hoja3.Cell(1, 17).Value = MonthString(mes - 1).ToUpper & " ADICIONALES"
                hoja3.Cell(1, 17).Style.Font.SetBold(True)

                hoja3.Cell(6, 4).FormulaA1 = "=" & periodo & "!D5"
                hoja3.Cell(7, 4).FormulaA1 = "=" & periodo & "!E5"
                hoja3.Cell(8, 4).FormulaA1 = "=" & periodo & "!F5"
                hoja3.Cell(9, 4).FormulaA1 = "=" & periodo & "!G5"
                'hoja3.Cell(10, 4).FormulaA1 = "=" & periodo & "!H5"
                hoja3.Cell(11, 4).FormulaA1 = "=" & periodo & "!I5"
                hoja3.Cell(12, 4).FormulaA1 = "=" & periodo & "!H5"
                hoja3.Cell(13, 4).FormulaA1 = "=" & periodo & "!J5"
                hoja3.Cell(16, 4).FormulaA1 = "=" & periodo & "!P5"
                hoja3.Cell(19, 4).FormulaA1 = "=" & periodo & "!O5"
                hoja3.Cell(20, 4).FormulaA1 = "=" & periodo & "!L5"
                hoja3.Cell(31, 4).FormulaA1 = "=" & periodo & "!P5"
                hoja3.Cell(32, 4).FormulaA1 = "=" & periodo & "!Q5"
                hoja3.Cell(33, 4).FormulaA1 = "=" & periodo & "!R5"
                hoja3.Cell(34, 4).FormulaA1 = "=" & periodo & "!S5"


                hoja3.Cell(6, 5).FormulaA1 = "=" & periodo & "!D6"
                hoja3.Cell(7, 5).FormulaA1 = "=" & periodo & "!E6"
                hoja3.Cell(8, 5).FormulaA1 = "=" & periodo & "!F6"
                hoja3.Cell(9, 5).FormulaA1 = "=" & periodo & "!G6"
                'hoja3.Cell(10, 5).FormulaA1 = "=" & periodo & "!H6"
                hoja3.Cell(11, 5).FormulaA1 = "=" & periodo & "!I6"
                hoja3.Cell(12, 5).FormulaA1 = "=" & periodo & "!H6"
                hoja3.Cell(13, 5).FormulaA1 = "=" & periodo & "!J6"
                hoja3.Cell(16, 5).FormulaA1 = "=" & periodo & "!P6"
                hoja3.Cell(19, 5).FormulaA1 = "=" & periodo & "!O6"
                hoja3.Cell(20, 5).FormulaA1 = "=" & periodo & "!L6"
                hoja3.Cell(31, 5).FormulaA1 = "=" & periodo & "!P6"
                hoja3.Cell(32, 5).FormulaA1 = "=" & periodo & "!Q6"
                hoja3.Cell(33, 5).FormulaA1 = "=" & periodo & "!R6"
                hoja3.Cell(34, 5).FormulaA1 = "=" & periodo & "!S6"



                hoja3.Cell(6, 6).FormulaA1 = "=" & periodo & "!D7"
                hoja3.Cell(7, 6).FormulaA1 = "=" & periodo & "!E7"
                hoja3.Cell(8, 6).FormulaA1 = "=" & periodo & "!F7"
                hoja3.Cell(9, 6).FormulaA1 = "=" & periodo & "!G7"
                'hoja3.Cell(10, 6).FormulaA1 = "=" & periodo & "!H7"
                hoja3.Cell(11, 6).FormulaA1 = "=" & periodo & "!I7"
                hoja3.Cell(12, 6).FormulaA1 = "=" & periodo & "!H7"
                hoja3.Cell(13, 6).FormulaA1 = "=" & periodo & "!J7"
                hoja3.Cell(16, 6).FormulaA1 = "=" & periodo & "!P7"
                hoja3.Cell(19, 6).FormulaA1 = "=" & periodo & "!O7"
                hoja3.Cell(20, 6).FormulaA1 = "=" & periodo & "!L7"
                hoja3.Cell(31, 6).FormulaA1 = "=" & periodo & "!P7"
                hoja3.Cell(32, 6).FormulaA1 = "=" & periodo & "!Q7"
                hoja3.Cell(33, 6).FormulaA1 = "=" & periodo & "!R7"
                hoja3.Cell(34, 6).FormulaA1 = "=" & periodo & "!S7"

                hoja3.Cell(6, 7).FormulaA1 = "=" & periodo & "!D8"
                hoja3.Cell(7, 7).FormulaA1 = "=" & periodo & "!E8"
                hoja3.Cell(8, 7).FormulaA1 = "=" & periodo & "!F8"
                hoja3.Cell(9, 7).FormulaA1 = "=" & periodo & "!G8"
                'hoja3.Cell(10, 7).FormulaA1 = "=" & periodo & "!H8"
                hoja3.Cell(11, 7).FormulaA1 = "=" & periodo & "!I8"
                hoja3.Cell(12, 7).FormulaA1 = "=" & periodo & "!H8"
                hoja3.Cell(13, 7).FormulaA1 = "=" & periodo & "!J8"
                hoja3.Cell(16, 7).FormulaA1 = "=" & periodo & "!P8"
                hoja3.Cell(19, 7).FormulaA1 = "=" & periodo & "!O8"
                hoja3.Cell(20, 7).FormulaA1 = "=" & periodo & "!L8"
                hoja3.Cell(31, 7).FormulaA1 = "=" & periodo & "!P8"
                hoja3.Cell(32, 7).FormulaA1 = "=" & periodo & "!Q8"
                hoja3.Cell(33, 7).FormulaA1 = "=" & periodo & "!R8"
                hoja3.Cell(34, 7).FormulaA1 = "=" & periodo & "!S8"

                hoja3.Cell(6, 8).FormulaA1 = "=" & periodo & "!D9"
                hoja3.Cell(7, 8).FormulaA1 = "=" & periodo & "!E9"
                hoja3.Cell(8, 8).FormulaA1 = "=" & periodo & "!F9"
                hoja3.Cell(9, 8).FormulaA1 = "=" & periodo & "!G9"
                'hoja3.Cell(10, 8).FormulaA1 = "=" & periodo & "!H9"
                hoja3.Cell(11, 8).FormulaA1 = "=" & periodo & "!I9"
                hoja3.Cell(12, 8).FormulaA1 = "=" & periodo & "!H9"
                hoja3.Cell(13, 8).FormulaA1 = "=" & periodo & "!J9"
                hoja3.Cell(16, 8).FormulaA1 = "=" & periodo & "!P9"
                hoja3.Cell(19, 8).FormulaA1 = "=" & periodo & "!O9"
                hoja3.Cell(20, 8).FormulaA1 = "=" & periodo & "!L9"
                hoja3.Cell(31, 8).FormulaA1 = "=" & periodo & "!P9"
                hoja3.Cell(32, 8).FormulaA1 = "=" & periodo & "!Q9"
                hoja3.Cell(33, 8).FormulaA1 = "=" & periodo & "!R9"
                hoja3.Cell(34, 8).FormulaA1 = "=" & periodo & "!S9"

                hoja3.Cell(6, 9).FormulaA1 = "=" & periodo & "!D10"
                hoja3.Cell(7, 9).FormulaA1 = "=" & periodo & "!E10"
                hoja3.Cell(8, 9).FormulaA1 = "=" & periodo & "!F10"
                hoja3.Cell(9, 9).FormulaA1 = "=" & periodo & "!G10"
                'hoja3.Cell(10, 9).FormulaA1 = "=" & periodo & "!H10"
                hoja3.Cell(11, 9).FormulaA1 = "=" & periodo & "!I10"
                hoja3.Cell(12, 9).FormulaA1 = "=" & periodo & "!H10"
                hoja3.Cell(13, 9).FormulaA1 = "=" & periodo & "!J10"
                hoja3.Cell(16, 9).FormulaA1 = "=" & periodo & "!P10"
                hoja3.Cell(19, 9).FormulaA1 = "=" & periodo & "!O10"
                hoja3.Cell(20, 9).FormulaA1 = "=" & periodo & "!L10"
                hoja3.Cell(31, 9).FormulaA1 = "=" & periodo & "!P10"
                hoja3.Cell(32, 9).FormulaA1 = "=" & periodo & "!Q10"
                hoja3.Cell(33, 9).FormulaA1 = "=" & periodo & "!R10"
                hoja3.Cell(34, 9).FormulaA1 = "=" & periodo & "!S10"


                hoja3.Cell(6, 10).FormulaA1 = "=" & periodo & "!D11"
                hoja3.Cell(7, 10).FormulaA1 = "=" & periodo & "!E11"
                hoja3.Cell(8, 10).FormulaA1 = "=" & periodo & "!F11"
                hoja3.Cell(9, 10).FormulaA1 = "=" & periodo & "!G11"
                'hoja3.Cell(10, 10).FormulaA1 = "=" & periodo & "!H11"
                hoja3.Cell(11, 10).FormulaA1 = "=" & periodo & "!I11"
                hoja3.Cell(12, 10).FormulaA1 = "=" & periodo & "!H11"
                hoja3.Cell(13, 10).FormulaA1 = "=" & periodo & "!J11"
                hoja3.Cell(16, 10).FormulaA1 = "=" & periodo & "!P11"
                hoja3.Cell(19, 10).FormulaA1 = "=" & periodo & "!O11"
                hoja3.Cell(20, 10).FormulaA1 = "=" & periodo & "!L11"

                hoja3.Cell(31, 10).FormulaA1 = "=" & periodo & "!P11"
                hoja3.Cell(32, 10).FormulaA1 = "=" & periodo & "!Q11"
                hoja3.Cell(33, 10).FormulaA1 = "=" & periodo & "!R11"
                hoja3.Cell(34, 10).FormulaA1 = "=" & periodo & "!S11"

                hoja3.Cell(6, 11).FormulaA1 = "=" & periodo & "!D12"
                hoja3.Cell(7, 11).FormulaA1 = "=" & periodo & "!E12"
                hoja3.Cell(8, 11).FormulaA1 = "=" & periodo & "!F12"
                hoja3.Cell(9, 11).FormulaA1 = "=" & periodo & "!G12"
                'hoja3.Cell(10, 11).FormulaA1 = "=" & periodo & "!H12"
                hoja3.Cell(11, 11).FormulaA1 = "=" & periodo & "!I12"
                hoja3.Cell(12, 11).FormulaA1 = "=" & periodo & "!H12"
                hoja3.Cell(13, 11).FormulaA1 = "=" & periodo & "!J12"
                hoja3.Cell(16, 11).FormulaA1 = "=" & periodo & "!P12"
                hoja3.Cell(19, 11).FormulaA1 = "=" & periodo & "!O12"
                hoja3.Cell(20, 11).FormulaA1 = "=" & periodo & "!L12"
                hoja3.Cell(31, 11).FormulaA1 = "=" & periodo & "!P12"
                hoja3.Cell(32, 11).FormulaA1 = "=" & periodo & "!Q12"
                hoja3.Cell(33, 11).FormulaA1 = "=" & periodo & "!R12"
                hoja3.Cell(34, 11).FormulaA1 = "=" & periodo & "!S12"

                hoja3.Cell(6, 12).FormulaA1 = "=" & periodo & "!D13"
                hoja3.Cell(7, 12).FormulaA1 = "=" & periodo & "!E13"
                hoja3.Cell(8, 12).FormulaA1 = "=" & periodo & "!F13"
                hoja3.Cell(9, 12).FormulaA1 = "=" & periodo & "!G13"
                'hoja3.Cell(10, 12).FormulaA1 = "=" & periodo & "!H13"
                hoja3.Cell(11, 12).FormulaA1 = "=" & periodo & "!I13"
                hoja3.Cell(12, 12).FormulaA1 = "=" & periodo & "!H13"
                hoja3.Cell(13, 12).FormulaA1 = "=" & periodo & "!J13"
                hoja3.Cell(16, 12).FormulaA1 = "=" & periodo & "!P13"
                hoja3.Cell(19, 12).FormulaA1 = "=" & periodo & "!O13"
                hoja3.Cell(20, 12).FormulaA1 = "=" & periodo & "!L13"
                hoja3.Cell(31, 12).FormulaA1 = "=" & periodo & "!P13"
                hoja3.Cell(32, 12).FormulaA1 = "=" & periodo & "!Q13"
                hoja3.Cell(33, 12).FormulaA1 = "=" & periodo & "!R13"
                hoja3.Cell(34, 12).FormulaA1 = "=" & periodo & "!S13"

                hoja3.Cell(6, 13).FormulaA1 = "=" & periodo & "!D14"
                hoja3.Cell(7, 13).FormulaA1 = "=" & periodo & "!E14"
                hoja3.Cell(8, 13).FormulaA1 = "=" & periodo & "!F14"
                hoja3.Cell(9, 13).FormulaA1 = "=" & periodo & "!G14"
                'hoja3.Cell(10, 13).FormulaA1 = "=" & periodo & "!H14"
                hoja3.Cell(11, 13).FormulaA1 = "=" & periodo & "!I14"
                hoja3.Cell(12, 13).FormulaA1 = "=" & periodo & "!H14"
                hoja3.Cell(13, 13).FormulaA1 = "=" & periodo & "!J14"
                hoja3.Cell(16, 13).FormulaA1 = "=" & periodo & "!P14"
                hoja3.Cell(19, 13).FormulaA1 = "=" & periodo & "!O14"
                hoja3.Cell(20, 13).FormulaA1 = "=" & periodo & "!L14"
                hoja3.Cell(31, 13).FormulaA1 = "=" & periodo & "!P14"
                hoja3.Cell(32, 13).FormulaA1 = "=" & periodo & "!Q14"
                hoja3.Cell(33, 13).FormulaA1 = "=" & periodo & "!R14"
                hoja3.Cell(34, 13).FormulaA1 = "=" & periodo & "!S14"

                'Titulo
                Dim moment As Date = Date.Now()
                Dim month As Integer = moment.Month
                Dim year As Integer = moment.Year


                dialogo.FileName = "Reporte " + periodo.ToUpper
                dialogo.Filter = "Archivos de Excel (*.xlsx)|*.xlsx"
                ''  dialogo.ShowDialog()

                If dialogo.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    ' OK button pressed
                    libro.SaveAs(dialogo.FileName)
                    libro = Nothing
                    MessageBox.Show("Archivo generado correctamente", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("No se guardo el archivo", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Try


    End Sub

    Private Sub tsbImportar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles tsbImportar.Click

    End Sub

    Private Sub cmdincidencias_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdincidencias.Click

    End Sub

    Private Sub cmdreiniciar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdreiniciar.Click
        Try
            Dim sql As String
            Dim resultado As Integer = MessageBox.Show("Se borraran los datos tanto de la nomina Abordo como Descanso,¿Desea reiniciar la nomina?", "Pregunta", MessageBoxButtons.YesNo)
            If resultado = DialogResult.Yes Then

                sql = "select * from Nomina where fkiIdEmpresa=1 and fkiIdPeriodo=" & cboperiodo.SelectedValue
                sql &= " and iEstatusNomina=1 and iEstatus=1 and iEstatusEmpleado=" & cboserie.SelectedIndex
                sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex

                Dim rwNominaGuardadaFinal As DataRow() = nConsulta(sql)



                If rwNominaGuardadaFinal Is Nothing = False Then
                    MessageBox.Show("La nomina ya esta marcada como final, no  se pueden guardar cambios.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    MessageBox.Show("Se borraran los datos tanto de la nomina abordo como la de descanso", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)

                    sql = "delete from Nomina"
                    sql &= " where fkiIdEmpresa=1 and fkiIdPeriodo=" & cboperiodo.SelectedValue
                    sql &= " and iEstatusNomina=0 and iEstatus=1 and iEstatusEmpleado=" & cboserie.SelectedIndex
                    'sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex

                    If nExecute(sql) = False Then
                        MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        'pnlProgreso.Visible = False
                        Exit Sub
                    End If

                    'borrar el detalle del infonavit


                    sql = "delete from DetalleDescInfonavit"
                    sql &= " where fkiIdPeriodo=" & cboperiodo.SelectedValue
                    sql &= " and iSerie=" & cboserie.SelectedIndex
                    'sql &= " and iSerie=" & cboserie.SelectedIndex
                    'sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex

                    If nExecute(sql) = False Then
                        MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        'pnlProgreso.Visible = False
                        Exit Sub
                    End If

                    MessageBox.Show("Nomina reiniciada correctamente, vuelva a cargar los datos", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    dtgDatos.DataSource = ""
                    dtgDatos.Columns.Clear()
                End If



            End If




        Catch ex As Exception

        End Try


    End Sub

    Private Sub tsbIEmpleados_Click(ByVal sender As Object, ByVal e As EventArgs) Handles tsbIEmpleados.Click
        Try
            Dim Forma As New frmEmpleados
            Forma.gIdEmpresa = gIdEmpresa
            Forma.gIdPeriodo = cboperiodo.SelectedValue
            Forma.gIdTipoPuesto = 1
            Forma.ShowDialog()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub dtgDatos_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles dtgDatos.CellClick
        Try
            If e.ColumnIndex = 0 Then
                dtgDatos.Rows(e.RowIndex).Cells(0).Value = Not dtgDatos.Rows(e.RowIndex).Cells(0).Value
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

    Private Sub dtgDatos_CellEnter(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles dtgDatos.CellEnter
        'MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    End Sub

    Private Sub TextboxNumeric_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        Try
            'Dim columna As Integer
            'Dim fila As Integer

            'columna = CInt(DirectCast(sender, System.Windows.Forms.DataGridView).CurrentCell.ColumnIndex)
            'Fila = CInt(DirectCast(sender, System.Windows.Forms.DataGridView).CurrentCell.RowIndex)


            Dim nonNumberEntered As Boolean

            nonNumberEntered = True

            If (Convert.ToInt32(e.KeyChar) >= 48 AndAlso Convert.ToInt32(e.KeyChar) <= 57) OrElse Convert.ToInt32(e.KeyChar) = 8 OrElse Convert.ToInt32(e.KeyChar) = 46 Then

                'If Convert.ToInt32(e.KeyChar) = 46 Then
                '    If InStr(dtgDatos.Rows(Fila).Cells(columna).Value, ".") = 0 Then
                '        nonNumberEntered = False
                '    Else
                '        nonNumberEntered = False
                '    End If
                'Else
                '    nonNumberEntered = False
                'End If
                nonNumberEntered = False
            End If

            If nonNumberEntered = True Then
                ' Stop the character from being entered into the control since it is non-numerical.
                e.Handled = True
            Else
                e.Handled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try



    End Sub

    Private Sub dtgDatos_CellEndEdit(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles dtgDatos.CellEndEdit
        Try
            If Not m_currentControl Is Nothing Then
                RemoveHandler m_currentControl.KeyPress, AddressOf TextboxNumeric_KeyPress
            End If
            If Not dgvCombo Is Nothing Then
                RemoveHandler dgvCombo.SelectedIndexChanged, AddressOf dvgCombo_SelectedIndexChanged
            End If
            If dgvCombo IsNot Nothing Then
                RemoveHandler dgvCombo.SelectedIndexChanged, New EventHandler(AddressOf dvgCombo_SelectedIndexChanged)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub dtgDatos_EditingControlShowing(ByVal sender As Object, ByVal e As DataGridViewEditingControlShowingEventArgs) Handles dtgDatos.EditingControlShowing
        Try
            Dim columna As Integer
            m_currentControl = Nothing
            columna = CInt(DirectCast(sender, System.Windows.Forms.DataGridView).CurrentCell.ColumnIndex)
            If columna = 15 Or columna = 18 Or columna = 39 Or columna = 40 Or columna = 41 Or columna = 42 Or columna = 43 Or columna = 10 Then
                AddHandler e.Control.KeyPress, AddressOf TextboxNumeric_KeyPress
                m_currentControl = e.Control
            End If


            dgvCombo = TryCast(e.Control, DataGridViewComboBoxEditingControl)

            If dgvCombo IsNot Nothing Then
                AddHandler dgvCombo.SelectedIndexChanged, New EventHandler(AddressOf dvgCombo_SelectedIndexChanged)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try




    End Sub

    Private Sub dtgDatos_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) Handles dtgDatos.ColumnHeaderMouseClick
        Try
            Dim newColumn As DataGridViewColumn = dtgDatos.Columns(e.ColumnIndex)

            Dim sql As String
            If e.ColumnIndex = 0 Then
                dtgDatos.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
            Else
                If e.ColumnIndex = 11 Then
                    'DirectCast(dtgDatos.Columns(11), DataGridViewComboBoxColumn).Sorted = True
                    Dim resultado As Integer = MessageBox.Show("Para realizar este ordenamiento es necesario guardar la nomina primeramente, ¿desea continuar?", "Pregunta", MessageBoxButtons.YesNo)
                    If resultado = DialogResult.Yes Then

                        cmdguardarnomina_Click(sender, e)
                        campoordenamiento = "nomina.Puesto,cNombreLargo"
                        llenargrid()
                    End If

                End If

                If e.ColumnIndex = 12 Then
                    Dim resultado As Integer = MessageBox.Show("Para realizar este ordenamiento es necesario guardar la nomina primeramente, ¿desea continuar?", "Pregunta", MessageBoxButtons.YesNo)
                    If resultado = DialogResult.Yes Then

                        cmdguardarnomina_Click(sender, e)
                        campoordenamiento = "Nomina.Buque,cNombreLargo"
                        llenargrid()
                    End If
                End If
                'dtgDatos.Columns(e.ColumnIndex).SortMode = DataGridViewColumnSortMode.Automatic
            End If

            For x As Integer = 0 To dtgDatos.Rows.Count - 1

                sql = "select * from empleadosC where iIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                Dim rwFila As DataRow() = nConsulta(sql)



                CType(Me.dtgDatos.Rows(x).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("cPuesto").ToString()

                CType(Me.dtgDatos.Rows(x).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("cFuncionesPuesto").ToString()
                dtgDatos.Rows(x).Cells(1).Value = x + 1
            Next


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

    Private Sub chkAll_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkAll.CheckedChanged
        For x As Integer = 0 To dtgDatos.Rows.Count - 1
            dtgDatos.Rows(x).Cells(0).Value = Not dtgDatos.Rows(x).Cells(0).Value
        Next
        chkAll.Text = IIf(chkAll.Checked, "Desmarcar todos", "Marcar todos")
    End Sub

    Private Sub cmdlayouts_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdlayouts.Click

    End Sub

    Function RemoverBasura(ByVal nombre As String) As String
        Dim COMPSTR As String = "áéíóúÁÉÍÓÚ.ñÑ"
        Dim REPLSTR As String = "aeiouAEIOU nN"
        Dim Posicion As Integer
        Dim cadena As String = ""
        Dim arreglo As Char() = nombre.ToCharArray()
        For x As Integer = 0 To arreglo.Length - 1
            Posicion = COMPSTR.IndexOf(arreglo(x))
            If Posicion <> -1 Then
                arreglo(x) = REPLSTR(Posicion)

            End If
            cadena = cadena & arreglo(x)
        Next
        Return cadena
    End Function

    Function TipoCuentaBanco(ByVal idempleado As String, ByVal idempresa As String) As String
        'Agregar el banco y el tipo de cuenta ya sea a terceros o interbancaria
        'Buscamos el banco y verificarmos el tipo de cuenta a tercero o interbancaria
        Dim Sql As String
        Dim cadenabanco As String
        cadenabanco = ""

        Sql = "select iIdempleadoC,NumCuenta,Clabe,cuenta2,clabe2,fkiIdBanco,fkiIdBanco2"
        Sql &= " from empleadosC"
        Sql &= " where fkiIdEmpresa=" & gIdEmpresa & " and iIdempleadoC=" & idempleado

        Dim rwDatosBanco As DataRow() = nConsulta(Sql)

        cadenabanco = "@"

        If rwDatosBanco Is Nothing = False Then
            If rwDatosBanco(0)("NumCuenta") = "" Then
                cadenabanco &= "I"
            Else
                cadenabanco &= "T"
            End If

            If rwDatosBanco(0)("fkiIdBanco") = "1" Then
                cadenabanco &= "-BANAMEX"
            ElseIf rwDatosBanco(0)("fkiIdBanco") = "4" Then
                cadenabanco &= "-BANCOMER"
            ElseIf rwDatosBanco(0)("fkiIdBanco") = "13" Then
                cadenabanco &= "-SCOTIABANK"
            ElseIf rwDatosBanco(0)("fkiIdBanco") = "18" Then
                cadenabanco &= "-BANORTE"
            Else
                cadenabanco &= "-OTRO"
            End If

            cadenabanco &= "/"

            If rwDatosBanco(0)("cuenta2") = "" Then
                cadenabanco &= "I"
            Else
                cadenabanco &= "T"
            End If

            If rwDatosBanco(0)("fkiIdBanco2") = "1" Then
                cadenabanco &= "-BANAMEX"
            ElseIf rwDatosBanco(0)("fkiIdBanco2") = "4" Then
                cadenabanco &= "-BANCOMER"
            ElseIf rwDatosBanco(0)("fkiIdBanco2") = "13" Then
                cadenabanco &= "-SCOTIABANK"
            ElseIf rwDatosBanco(0)("fkiIdBanco2") = "18" Then
                cadenabanco &= "-BANORTE"
            Else
                cadenabanco &= "-OTRO"
            End If


        End If

        Return cadenabanco
    End Function

    Function CalculoPrimaSindicato(ByVal idempleado As String, ByVal idempresa As String) As String
        'Agregar el banco y el tipo de cuenta ya sea a terceros o interbancaria
        'Buscamos el banco y verificarmos el tipo de cuenta a tercero o interbancaria
        Dim Sql As String
        Dim cadenabanco As String
        Dim dia As String
        Dim mes As String
        Dim anio As String
        Dim anios As Integer
        Dim sueldodiario As Double
        Dim dias As Integer

        Dim Prima As String


        cadenabanco = ""


        Sql = "select *"
        Sql &= " from empleadosC"
        Sql &= " where fkiIdEmpresa=" & gIdEmpresa & " and iIdempleadoC=" & idempleado

        Dim rwDatosBanco As DataRow() = nConsulta(Sql)

        cadenabanco = "@"
        Prima = "0"
        If rwDatosBanco Is Nothing = False Then

            If Double.Parse(rwDatosBanco(0)("fsueldoOrd")) > 0 Then
                dia = Date.Parse(rwDatosBanco(0)("dFechaAntiguedad").ToString).Day.ToString("00")
                mes = Date.Parse(rwDatosBanco(0)("dFechaAntiguedad").ToString).Month.ToString("00")
                anio = Date.Today.Year
                'verificar el periodo para saber si queda entre el rango de fecha

                sueldodiario = Double.Parse(rwDatosBanco(0)("fsueldoOrd")) / diasperiodo

                Sql = "select * from periodos where iIdPeriodo= " & cboperiodo.SelectedValue
                Dim rwPeriodo As DataRow() = nConsulta(Sql)

                If rwPeriodo Is Nothing = False Then
                    Dim FechaBuscar As Date = Date.Parse(dia & "/" & mes & "/" & anio)
                    Dim FechaInicial As Date = Date.Parse(rwPeriodo(0)("dFechaInicio"))
                    Dim FechaFinal As Date = Date.Parse(rwPeriodo(0)("dFechaFin"))
                    Dim FechaAntiguedad As Date = Date.Parse(rwDatosBanco(0)("dFechaAntiguedad"))

                    If FechaBuscar.CompareTo(FechaInicial) >= 0 And FechaBuscar.CompareTo(FechaFinal) <= 0 Then
                        'Estamos dentro del rango 
                        'Calculamos la prima

                        anios = DateDiff("yyyy", FechaAntiguedad, FechaBuscar)

                        dias = CalculoDiasVacaciones(anios)

                        'Calcular prima

                        Prima = Math.Round(sueldodiario * dias * 0.25, 2).ToString()




                    End If


                End If


            End If


        End If


        Return Prima


    End Function


    Function CalculoDiasVacaciones(ByVal anios As Integer) As Integer
        Dim dias As Integer

        If anios = 1 Then
            dias = 6
        End If

        If anios = 2 Then
            dias = 8
        End If

        If anios = 3 Then
            dias = 10
        End If

        If anios = 4 Then
            dias = 12
        End If

        If anios >= 5 And anios <= 9 Then
            dias = 14
        End If

        If anios >= 10 And anios <= 14 Then
            dias = 16
        End If

        If anios >= 15 And anios <= 19 Then
            dias = 18
        End If

        If anios >= 20 And anios <= 24 Then
            dias = 20
        End If

        If anios >= 25 And anios <= 29 Then
            dias = 22
        End If

        If anios >= 30 And anios <= 34 Then
            dias = 24
        End If

        Return dias
    End Function

    Private Sub dtgDatos_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dtgDatos.CellContentClick
        Try
            If e.RowIndex = -1 And e.ColumnIndex = 0 Then
                Return
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub tsbEmpleados_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbEmpleados.Click
        Dim frm As New frmImportarEmpleadosAlta
        frm.ShowDialog()
    End Sub

    Private Sub dtgDatos_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dtgDatos.DataError
        Try
            e.Cancel = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub EliminarDeLaListaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EliminarDeLaListaToolStripMenuItem.Click
        If dtgDatos.CurrentRow Is Nothing = False Then
            Dim resultado As Integer = MessageBox.Show("¿Desea eliminar a este trabajador de la lista?", "Pregunta", MessageBoxButtons.YesNo)
            If resultado = DialogResult.Yes Then

                dtgDatos.Rows.Remove(dtgDatos.CurrentRow)
            End If
        End If


    End Sub

    Private Sub cbodias_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub cboserie_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboserie.SelectedIndexChanged
        dtgDatos.Columns.Clear()
        dtgDatos.DataSource = ""


    End Sub

    Private Sub cmdrecibosA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdrecibosA.Click

    End Sub

<<<<<<< HEAD
    Private Sub cboTipoNomina_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTipoNomina.SelectedIndexChanged
        dtgDatos.Columns.Clear()
        dtgDatos.DataSource = ""
=======
    Private Sub cboTipoNomina_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboTipoNomina.SelectedIndexChanged
        'dtgDatos.Columns.Clear()
        'dtgDatos.DataSource = ""
>>>>>>> origin/master
    End Sub

    Private Sub AgregarTrabajadoresToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AgregarTrabajadoresToolStripMenuItem.Click
        Try
            Dim Forma As New frmAgregarEmpleado
            Dim ids As String()
            Dim sql As String
            Dim cadenaempleados As String
            If Forma.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim dsPeriodo As New DataSet
                dsPeriodo.Tables.Add("Tabla")
                dsPeriodo.Tables("Tabla").Columns.Add("Consecutivo")
                dsPeriodo.Tables("Tabla").Columns.Add("Id_empleado")
                dsPeriodo.Tables("Tabla").Columns.Add("CodigoEmpleado")
                dsPeriodo.Tables("Tabla").Columns.Add("Nombre")
                dsPeriodo.Tables("Tabla").Columns.Add("Status")
                dsPeriodo.Tables("Tabla").Columns.Add("RFC")
                dsPeriodo.Tables("Tabla").Columns.Add("CURP")
                dsPeriodo.Tables("Tabla").Columns.Add("Num_IMSS")
                dsPeriodo.Tables("Tabla").Columns.Add("Fecha_Nac")
                dsPeriodo.Tables("Tabla").Columns.Add("Edad")
                dsPeriodo.Tables("Tabla").Columns.Add("Puesto")
                dsPeriodo.Tables("Tabla").Columns.Add("Buque")
                dsPeriodo.Tables("Tabla").Columns.Add("Tipo_Infonavit")
                dsPeriodo.Tables("Tabla").Columns.Add("Valor_Infonavit")
                dsPeriodo.Tables("Tabla").Columns.Add("Sueldo_Base")
                dsPeriodo.Tables("Tabla").Columns.Add("Salario_Diario")
                dsPeriodo.Tables("Tabla").Columns.Add("Salario_Cotización")
                dsPeriodo.Tables("Tabla").Columns.Add("Dias_Trabajados")
                dsPeriodo.Tables("Tabla").Columns.Add("Tipo_Incapacidad")
                dsPeriodo.Tables("Tabla").Columns.Add("Número_días")
                dsPeriodo.Tables("Tabla").Columns.Add("Sueldo_Bruto")
                dsPeriodo.Tables("Tabla").Columns.Add("Tiempo_Extra_Fijo_Gravado")
                dsPeriodo.Tables("Tabla").Columns.Add("Tiempo_Extra_Fijo_Exento")
                dsPeriodo.Tables("Tabla").Columns.Add("Tiempo_Extra_Ocasional")
                dsPeriodo.Tables("Tabla").Columns.Add("Desc_Sem_Obligatorio")
                dsPeriodo.Tables("Tabla").Columns.Add("Vacaciones_proporcionales")
                dsPeriodo.Tables("Tabla").Columns.Add("Aguinaldo_gravado")
                dsPeriodo.Tables("Tabla").Columns.Add("Aguinaldo_exento")
                dsPeriodo.Tables("Tabla").Columns.Add("Total_Aguinaldo")
                dsPeriodo.Tables("Tabla").Columns.Add("Prima_vac_gravado")
                dsPeriodo.Tables("Tabla").Columns.Add("Prima_vac_exento")
                dsPeriodo.Tables("Tabla").Columns.Add("Total_Prima_vac")
                dsPeriodo.Tables("Tabla").Columns.Add("Total_percepciones")
                dsPeriodo.Tables("Tabla").Columns.Add("Total_percepciones_p/isr")
                dsPeriodo.Tables("Tabla").Columns.Add("Incapacidad")
                dsPeriodo.Tables("Tabla").Columns.Add("ISR")
                dsPeriodo.Tables("Tabla").Columns.Add("IMSS")
                dsPeriodo.Tables("Tabla").Columns.Add("Infonavit")
                dsPeriodo.Tables("Tabla").Columns.Add("Infonavit_bim_anterior")
                dsPeriodo.Tables("Tabla").Columns.Add("Ajuste_infonavit")
                dsPeriodo.Tables("Tabla").Columns.Add("Pension_Alimenticia")
                dsPeriodo.Tables("Tabla").Columns.Add("Prestamo")
                dsPeriodo.Tables("Tabla").Columns.Add("Fonacot")
                dsPeriodo.Tables("Tabla").Columns.Add("Subsidio_Generado")
                dsPeriodo.Tables("Tabla").Columns.Add("Subsidio_Aplicado")
                dsPeriodo.Tables("Tabla").Columns.Add("Operadora")
                dsPeriodo.Tables("Tabla").Columns.Add("Prestamo_Personal_A")
                dsPeriodo.Tables("Tabla").Columns.Add("Adeudo_Infonavit_A")
                dsPeriodo.Tables("Tabla").Columns.Add("Diferencia_Infonavit_A")
                dsPeriodo.Tables("Tabla").Columns.Add("Asimilados")
                dsPeriodo.Tables("Tabla").Columns.Add("Retenciones_Operadora")
                dsPeriodo.Tables("Tabla").Columns.Add("%_Comisión")
                dsPeriodo.Tables("Tabla").Columns.Add("Comisión_Operadora")
                dsPeriodo.Tables("Tabla").Columns.Add("Comisión_Asimilados")
                dsPeriodo.Tables("Tabla").Columns.Add("IMSS_CS")
                dsPeriodo.Tables("Tabla").Columns.Add("RCV_CS")
                dsPeriodo.Tables("Tabla").Columns.Add("Infonavit_CS")
                dsPeriodo.Tables("Tabla").Columns.Add("ISN_CS")
                dsPeriodo.Tables("Tabla").Columns.Add("Total_Costo_Social")
                dsPeriodo.Tables("Tabla").Columns.Add("Subtotal")
                dsPeriodo.Tables("Tabla").Columns.Add("IVA")
                dsPeriodo.Tables("Tabla").Columns.Add("TOTAL_DEPOSITO")


                ids = Forma.gidEmpleados.Split(",")
                If dtgDatos.Rows.Count > 0 Then
                    'Dim dt As DataTable = DirectCast(dtgDatos.DataSource, DataTable)
                    'dsPeriodo.Tables("Tabla") = dtgDatos.DataSource, DataTable
                    'Dim dt As DataTable = dsPeriodo.Tables("Tabla")


                    'For y As Integer = 0 To dt.Rows.Count - 1
                    '    dsPeriodo.Tables("Tabla").ImportRow(dt.Rows[y])
                    'Next

                    'Pasamos del datagrid al dataset ya creado
                    For y As Integer = 0 To dtgDatos.Rows.Count - 1

                        Dim fila As DataRow = dsPeriodo.Tables("Tabla").NewRow

                        fila.Item("Consecutivo") = (y + 1).ToString
                        fila.Item("Id_empleado") = dtgDatos.Rows(y).Cells(2).Value
                        fila.Item("CodigoEmpleado") = dtgDatos.Rows(y).Cells(3).Value
                        fila.Item("Nombre") = dtgDatos.Rows(y).Cells(4).Value
                        fila.Item("Status") = dtgDatos.Rows(y).Cells(5).Value
                        fila.Item("RFC") = dtgDatos.Rows(y).Cells(6).Value
                        fila.Item("CURP") = dtgDatos.Rows(y).Cells(7).Value
                        fila.Item("Num_IMSS") = dtgDatos.Rows(y).Cells(8).Value
                        fila.Item("Fecha_Nac") = dtgDatos.Rows(y).Cells(9).Value
                        'Dim tiempo As TimeSpan = Date.Now - Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString)

                        fila.Item("Edad") = dtgDatos.Rows(y).Cells(10).Value
                        fila.Item("Puesto") = dtgDatos.Rows(y).Cells(11).FormattedValue
                        fila.Item("Buque") = dtgDatos.Rows(y).Cells(12).FormattedValue

                        fila.Item("Tipo_Infonavit") = dtgDatos.Rows(y).Cells(13).Value
                        fila.Item("Valor_Infonavit") = IIf(dtgDatos.Rows(y).Cells(14).Value = "", "0", dtgDatos.Rows(y).Cells(14).Value.ToString.Replace(",", ""))
                        'salario base
                        fila.Item("Sueldo_Base") = dtgDatos.Rows(y).Cells(15).Value
                        fila.Item("Salario_Diario") = dtgDatos.Rows(y).Cells(16).Value
                        fila.Item("Salario_Cotización") = dtgDatos.Rows(y).Cells(17).Value
                        fila.Item("Dias_Trabajados") = dtgDatos.Rows(y).Cells(18).Value
                        fila.Item("Tipo_Incapacidad") = dtgDatos.Rows(y).Cells(19).Value
                        fila.Item("Número_días") = dtgDatos.Rows(y).Cells(20).Value
                        fila.Item("Sueldo_Bruto") = IIf(dtgDatos.Rows(y).Cells(21).Value = "", "0", dtgDatos.Rows(y).Cells(21).Value.ToString.Replace(",", ""))
                        fila.Item("Tiempo_Extra_Fijo_Gravado") = IIf(dtgDatos.Rows(y).Cells(22).Value = "", "0", dtgDatos.Rows(y).Cells(22).Value.ToString.Replace(",", ""))
                        fila.Item("Tiempo_Extra_Fijo_Exento") = IIf(dtgDatos.Rows(y).Cells(23).Value = "", "0", dtgDatos.Rows(y).Cells(23).Value.ToString.Replace(",", ""))
                        fila.Item("Tiempo_Extra_Ocasional") = IIf(dtgDatos.Rows(y).Cells(24).Value = "", "0", dtgDatos.Rows(y).Cells(24).Value.ToString.Replace(",", ""))
                        fila.Item("Desc_Sem_Obligatorio") = IIf(dtgDatos.Rows(y).Cells(25).Value = "", "0", dtgDatos.Rows(y).Cells(25).Value.ToString.Replace(",", ""))
                        fila.Item("Vacaciones_proporcionales") = IIf(dtgDatos.Rows(y).Cells(26).Value = "", "0", dtgDatos.Rows(y).Cells(26).Value.ToString.Replace(",", ""))
                        fila.Item("Aguinaldo_gravado") = IIf(dtgDatos.Rows(y).Cells(27).Value = "", "0", dtgDatos.Rows(y).Cells(27).Value.ToString.Replace(",", ""))
                        fila.Item("Aguinaldo_exento") = IIf(dtgDatos.Rows(y).Cells(28).Value = "", "0", dtgDatos.Rows(y).Cells(28).Value.ToString.Replace(",", ""))
                        fila.Item("Total_Aguinaldo") = IIf(dtgDatos.Rows(y).Cells(29).Value = "", "0", dtgDatos.Rows(y).Cells(29).Value.ToString.Replace(",", ""))
                        fila.Item("Prima_vac_gravado") = IIf(dtgDatos.Rows(y).Cells(30).Value = "", "0", dtgDatos.Rows(y).Cells(30).Value.ToString.Replace(",", ""))
                        fila.Item("Prima_vac_exento") = IIf(dtgDatos.Rows(y).Cells(31).Value = "", "0", dtgDatos.Rows(y).Cells(31).Value.ToString.Replace(",", ""))

                        fila.Item("Total_Prima_vac") = IIf(dtgDatos.Rows(y).Cells(32).Value = "", "0", dtgDatos.Rows(y).Cells(32).Value.ToString.Replace(",", ""))
                        fila.Item("Total_percepciones") = IIf(dtgDatos.Rows(y).Cells(33).Value = "", "0", dtgDatos.Rows(y).Cells(33).Value.ToString.Replace(",", ""))
                        fila.Item("Total_percepciones_p/isr") = IIf(dtgDatos.Rows(y).Cells(34).Value = "", "0", dtgDatos.Rows(y).Cells(34).Value.ToString.Replace(",", ""))
                        fila.Item("Incapacidad") = IIf(dtgDatos.Rows(y).Cells(35).Value = "", "0", dtgDatos.Rows(y).Cells(35).Value.ToString.Replace(",", ""))
                        fila.Item("ISR") = IIf(dtgDatos.Rows(y).Cells(36).Value = "", "0", dtgDatos.Rows(y).Cells(36).Value.ToString.Replace(",", ""))
                        fila.Item("IMSS") = IIf(dtgDatos.Rows(y).Cells(37).Value = "", "0", dtgDatos.Rows(y).Cells(37).Value.ToString.Replace(",", ""))
                        fila.Item("Infonavit") = IIf(dtgDatos.Rows(y).Cells(38).Value = "", "0", dtgDatos.Rows(y).Cells(38).Value.ToString.Replace(",", ""))
                        fila.Item("Infonavit_bim_anterior") = IIf(dtgDatos.Rows(y).Cells(39).Value = "", "0", dtgDatos.Rows(y).Cells(39).Value.ToString.Replace(",", ""))
                        fila.Item("Ajuste_infonavit") = IIf(dtgDatos.Rows(y).Cells(40).Value = "", "0", dtgDatos.Rows(y).Cells(40).Value.ToString.Replace(",", ""))
                        fila.Item("Pension_Alimenticia") = IIf(dtgDatos.Rows(y).Cells(41).Value = "", "0", dtgDatos.Rows(y).Cells(41).Value.ToString.Replace(",", ""))
                        fila.Item("Prestamo") = IIf(dtgDatos.Rows(y).Cells(42).Value = "", "0", dtgDatos.Rows(y).Cells(42).Value.ToString.Replace(",", ""))
                        fila.Item("Fonacot") = IIf(dtgDatos.Rows(y).Cells(43).Value = "", "0", dtgDatos.Rows(y).Cells(43).Value.ToString.Replace(",", ""))
                        fila.Item("Subsidio_Generado") = IIf(dtgDatos.Rows(y).Cells(44).Value = "", "0", dtgDatos.Rows(y).Cells(44).Value.ToString.Replace(",", ""))
                        fila.Item("Subsidio_Aplicado") = IIf(dtgDatos.Rows(y).Cells(45).Value = "", "0", dtgDatos.Rows(y).Cells(45).Value.ToString.Replace(",", ""))
                        fila.Item("Operadora") = IIf(dtgDatos.Rows(y).Cells(46).Value = "", "0", dtgDatos.Rows(y).Cells(46).Value.ToString.Replace(",", ""))
                        fila.Item("Prestamo_Personal_A") = IIf(dtgDatos.Rows(y).Cells(47).Value = "", "0", dtgDatos.Rows(y).Cells(47).Value.ToString.Replace(",", ""))
                        fila.Item("Adeudo_Infonavit_A") = IIf(dtgDatos.Rows(y).Cells(48).Value = "", "0", dtgDatos.Rows(y).Cells(48).Value.ToString.Replace(",", ""))
                        fila.Item("Diferencia_Infonavit_A") = IIf(dtgDatos.Rows(y).Cells(49).Value = "", "0", dtgDatos.Rows(y).Cells(49).Value.ToString.Replace(",", ""))
                        fila.Item("Asimilados") = IIf(dtgDatos.Rows(y).Cells(50).Value = "", "0", dtgDatos.Rows(y).Cells(50).Value.ToString.Replace(",", ""))
                        fila.Item("Retenciones_Operadora") = IIf(dtgDatos.Rows(y).Cells(51).Value = "", "0", dtgDatos.Rows(y).Cells(51).Value.ToString.Replace(",", ""))
                        fila.Item("%_Comisión") = IIf(dtgDatos.Rows(y).Cells(52).Value = "", "0", dtgDatos.Rows(y).Cells(52).Value.ToString.Replace(",", ""))
                        fila.Item("Comisión_Operadora") = IIf(dtgDatos.Rows(y).Cells(53).Value = "", "0", dtgDatos.Rows(y).Cells(53).Value.ToString.Replace(",", ""))
                        fila.Item("Comisión_Asimilados") = IIf(dtgDatos.Rows(y).Cells(54).Value = "", "0", dtgDatos.Rows(y).Cells(54).Value.ToString.Replace(",", ""))
                        fila.Item("IMSS_CS") = IIf(dtgDatos.Rows(y).Cells(55).Value = "", "0", dtgDatos.Rows(y).Cells(55).Value.ToString.Replace(",", ""))
                        fila.Item("RCV_CS") = IIf(dtgDatos.Rows(y).Cells(56).Value = "", "0", dtgDatos.Rows(y).Cells(56).Value.ToString.Replace(",", ""))
                        fila.Item("Infonavit_CS") = IIf(dtgDatos.Rows(y).Cells(57).Value = "", "0", dtgDatos.Rows(y).Cells(57).Value.ToString.Replace(",", ""))
                        fila.Item("ISN_CS") = IIf(dtgDatos.Rows(y).Cells(58).Value = "", "0", dtgDatos.Rows(y).Cells(58).Value.ToString.Replace(",", ""))
                        fila.Item("Total_Costo_Social") = IIf(dtgDatos.Rows(y).Cells(59).Value = "", "0", dtgDatos.Rows(y).Cells(59).Value.ToString.Replace(",", ""))
                        fila.Item("Subtotal") = IIf(dtgDatos.Rows(y).Cells(60).Value = "", "0", dtgDatos.Rows(y).Cells(60).Value.ToString.Replace(",", ""))
                        fila.Item("IVA") = IIf(dtgDatos.Rows(y).Cells(61).Value = "", "0", dtgDatos.Rows(y).Cells(61).Value.ToString.Replace(",", ""))
                        fila.Item("TOTAL_DEPOSITO") = IIf(dtgDatos.Rows(y).Cells(62).Value = "", "0", dtgDatos.Rows(y).Cells(62).Value.ToString.Replace(",", ""))

                        fila.Item("Neto_Pagar") = IIf(dtgDatos.Rows(y).Cells(46).Value = "", "0", dtgDatos.Rows(y).Cells(46).Value.ToString.Replace(",", ""))
                        fila.Item("Excendente") = IIf(dtgDatos.Rows(y).Cells(47).Value = "", "0", dtgDatos.Rows(y).Cells(47).Value.ToString.Replace(",", ""))
                        fila.Item("Total") = IIf(dtgDatos.Rows(y).Cells(48).Value = "", "0", dtgDatos.Rows(y).Cells(48).Value.ToString.Replace(",", ""))
                        fila.Item("IMSS_CS") = IIf(dtgDatos.Rows(y).Cells(49).Value = "", "0", dtgDatos.Rows(y).Cells(49).Value.ToString.Replace(",", ""))
                        fila.Item("RCV_CS") = IIf(dtgDatos.Rows(y).Cells(50).Value = "", "0", dtgDatos.Rows(y).Cells(50).Value.ToString.Replace(",", ""))
                        fila.Item("Infonavit_CS") = IIf(dtgDatos.Rows(y).Cells(51).Value = "", "0", dtgDatos.Rows(y).Cells(51).Value.ToString.Replace(",", ""))
                        fila.Item("ISN_CS") = IIf(dtgDatos.Rows(y).Cells(52).Value = "", "0", dtgDatos.Rows(y).Cells(52).Value.ToString.Replace(",", ""))
                        fila.Item("Prestamo_Personal") = IIf(dtgDatos.Rows(y).Cells(53).Value = "", "0", dtgDatos.Rows(y).Cells(53).Value.ToString.Replace(",", ""))
                        fila.Item("Adeudo_Infonavit") = IIf(dtgDatos.Rows(y).Cells(54).Value = "", "0", dtgDatos.Rows(y).Cells(54).Value.ToString.Replace(",", ""))
                        fila.Item("Diferencia_Infonavit") = IIf(dtgDatos.Rows(y).Cells(55).Value = "", "0", dtgDatos.Rows(y).Cells(55).Value.ToString.Replace(",", ""))
                        fila.Item("Complemento_Asimilados") = IIf(dtgDatos.Rows(y).Cells(56).Value = "", "0", dtgDatos.Rows(y).Cells(56).Value.ToString.Replace(",", ""))


                        dsPeriodo.Tables("Tabla").Rows.Add(fila)
                    Next

                    'Agregar a la tabla los datos que vienen de la busqueda de empleados
                    For x As Integer = 0 To ids.Length - 1

                        Dim fila As DataRow = dsPeriodo.Tables("Tabla").NewRow
                        'Dim fila As DataRow = dt.NewRow
                        'Dim fila As DataRow = dsPeriodo.Tables("Tabla").NewRow
                        sql = "select  * from empleadosC where " 'fkiIdClienteInter=-1"
                        sql &= " iIdEmpleadoC=" & ids(x)
                        sql &= " order by cFuncionesPuesto,cNombreLargo"
                        Dim rwEmpleado As DataRow() = nConsulta(sql)
                        If rwEmpleado Is Nothing = False Then
                            fila.Item("Consecutivo") = (dtgDatos.Rows.Count + x + 1).ToString
                            fila.Item("Id_empleado") = rwEmpleado(0)("iIdEmpleadoC").ToString
                            fila.Item("CodigoEmpleado") = rwEmpleado(0)("cCodigoEmpleado").ToString
                            fila.Item("Nombre") = rwEmpleado(0)("cNombreLargo").ToString.ToUpper()
                            fila.Item("Status") = IIf(rwEmpleado(0)("iOrigen").ToString = "1", "INTERINO", "PLANTA")
                            fila.Item("RFC") = rwEmpleado(0)("cRFC").ToString
                            fila.Item("CURP") = rwEmpleado(0)("cCURP").ToString
                            fila.Item("Num_IMSS") = rwEmpleado(0)("cIMSS").ToString

                            fila.Item("Fecha_Nac") = Date.Parse(rwEmpleado(0)("dFechaNac").ToString).ToShortDateString()
                            'Dim tiempo As TimeSpan = Date.Now - Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString)
                            fila.Item("Edad") = CalcularEdad(Date.Parse(rwEmpleado(0)("dFechaNac").ToString).Day, Date.Parse(rwEmpleado(0)("dFechaNac").ToString).Month, Date.Parse(rwEmpleado(0)("dFechaNac").ToString).Year)
                            fila.Item("Puesto") = rwEmpleado(0)("cPuesto").ToString
                            fila.Item("Buque") = "ECO III"

                            fila.Item("Tipo_Infonavit") = rwEmpleado(0)("cTipoFactor").ToString
                            fila.Item("Valor_Infonavit") = rwEmpleado(0)("fFactor").ToString
                            fila.Item("Sueldo_Base") = "0.00"
                            fila.Item("Salario_Diario") = rwEmpleado(0)("fSueldoBase").ToString
                            fila.Item("Salario_Cotización") = rwEmpleado(0)("fSueldoIntegrado").ToString
                            fila.Item("Dias_Trabajados") = "30"
                            fila.Item("Tipo_Incapacidad") = TipoIncapacidad(rwEmpleado(0)("iIdEmpleadoC").ToString, cboperiodo.SelectedValue)
                            fila.Item("Número_días") = NumDiasIncapacidad(rwEmpleado(0)("iIdEmpleadoC").ToString, cboperiodo.SelectedValue)
                            fila.Item("Sueldo_Bruto") = ""
                            fila.Item("Tiempo_Extra_Fijo_Gravado") = ""
                            fila.Item("Tiempo_Extra_Fijo_Exento") = ""
                            fila.Item("Tiempo_Extra_Ocasional") = ""
                            fila.Item("Desc_Sem_Obligatorio") = ""
                            fila.Item("Vacaciones_proporcionales") = ""
                            fila.Item("Aguinaldo_gravado") = ""
                            fila.Item("Aguinaldo_exento") = ""
                            fila.Item("Total_Aguinaldo") = ""
                            fila.Item("Prima_vac_gravado") = ""
                            fila.Item("Prima_vac_exento") = ""

                            fila.Item("Total_Prima_vac") = ""
                            fila.Item("Total_percepciones") = ""
                            fila.Item("Total_percepciones_p/isr") = ""
                            fila.Item("Incapacidad") = ""
                            fila.Item("ISR") = ""
                            fila.Item("IMSS") = ""
                            fila.Item("Infonavit") = ""
                            fila.Item("Infonavit_bim_anterior") = ""
                            fila.Item("Ajuste_infonavit") = ""
                            fila.Item("Pension_Alimenticia") = ""
                            fila.Item("Prestamo") = ""
                            fila.Item("Fonacot") = ""
                            fila.Item("Subsidio_Generado") = ""
                            fila.Item("Subsidio_Aplicado") = ""
                            fila.Item("Operadora") = ""
                            fila.Item("Prestamo_Personal_A") = ""
                            fila.Item("Adeudo_Infonavit_A") = ""
                            fila.Item("Diferencia_Infonavit_A") = ""
                            fila.Item("Asimilados") = ""
                            fila.Item("Retenciones_Operadora") = ""
                            fila.Item("%_Comisión") = ""
                            fila.Item("Comisión_Operadora") = ""
                            fila.Item("Comisión_Asimilados") = ""
                            fila.Item("IMSS_CS") = ""
                            fila.Item("RCV_CS") = ""
                            fila.Item("Infonavit_CS") = ""
                            fila.Item("ISN_CS") = ""
                            fila.Item("Total_Costo_Social") = ""
                            fila.Item("Subtotal") = ""
                            fila.Item("IVA") = ""
                            fila.Item("TOTAL_DEPOSITO") = ""


                            dsPeriodo.Tables("Tabla").Rows.Add(fila)
                            'dt.Rows.Add(fila)



                        End If

                    Next
                    'dtgDatos.DataSource = dt
                    dtgDatos.Columns.Clear()
                    Dim chk As New DataGridViewCheckBoxColumn()
                    dtgDatos.Columns.Add(chk)
                    chk.HeaderText = ""
                    chk.Name = "chk"
                    dtgDatos.DataSource = dsPeriodo.Tables("Tabla")

                    dtgDatos.Columns(0).Width = 30
                    dtgDatos.Columns(0).ReadOnly = True
                    dtgDatos.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                    'consecutivo
                    dtgDatos.Columns(1).Width = 60
                    dtgDatos.Columns(1).ReadOnly = True
                    dtgDatos.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    'idempleado
                    dtgDatos.Columns(2).Width = 100
                    dtgDatos.Columns(2).ReadOnly = True
                    dtgDatos.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    'codigo empleado
                    dtgDatos.Columns(3).Width = 100
                    dtgDatos.Columns(3).ReadOnly = True
                    dtgDatos.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    'Nombre
                    dtgDatos.Columns(4).Width = 250
                    dtgDatos.Columns(4).ReadOnly = True
                    'Estatus
                    dtgDatos.Columns(5).Width = 100
                    dtgDatos.Columns(5).ReadOnly = True
                    'RFC
                    dtgDatos.Columns(6).Width = 100
                    dtgDatos.Columns(6).ReadOnly = True
                    'dtgDatos.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                    'CURP
                    dtgDatos.Columns(7).Width = 150
                    dtgDatos.Columns(7).ReadOnly = True
                    'IMSS 

                    dtgDatos.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(8).ReadOnly = True
                    'Fecha_Nac
                    dtgDatos.Columns(9).Width = 150
                    dtgDatos.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(9).ReadOnly = True

                    'Edad
                    dtgDatos.Columns(10).ReadOnly = True
                    dtgDatos.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                    'Puesto
                    dtgDatos.Columns(11).ReadOnly = True
                    dtgDatos.Columns(11).Width = 200
                    dtgDatos.Columns.Remove("Puesto")

                    Dim combo As New DataGridViewComboBoxColumn

                    sql = "select * from puestos where iTipo=1 order by cNombre"

                    'Dim rwPuestos As DataRow() = nConsulta(sql)
                    'If rwPuestos Is Nothing = False Then
                    '    combo.Items.Add("uno")
                    '    combo.Items.Add("dos")
                    '    combo.Items.Add("tres")
                    'End If

                    nCargaCBO(combo, sql, "cNombre", "iIdPuesto")

                    combo.HeaderText = "Puesto"

                    combo.Width = 150
                    dtgDatos.Columns.Insert(11, combo)
                    'DirectCast(dtgDatos.Columns(11), DataGridViewComboBoxColumn).Sorted = True
                    'Dim combo2 As New DataGridViewComboBoxCell
                    'combo2 = CType(Me.dtgDatos.Rows(2).Cells(11), DataGridViewComboBoxCell)
                    'combo2.Value = combo.Items(11)



                    'dtgDatos.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                    'Buque
                    'dtgDatos.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(12).ReadOnly = True
                    dtgDatos.Columns(12).Width = 150
                    dtgDatos.Columns.Remove("Buque")

                    Dim combo2 As New DataGridViewComboBoxColumn

                    sql = "select * from departamentos where iEstatus=1 order by cNombre"

                    'Dim rwPuestos As DataRow() = nConsulta(sql)
                    'If rwPuestos Is Nothing = False Then
                    '    combo.Items.Add("uno")
                    '    combo.Items.Add("dos")
                    '    combo.Items.Add("tres")
                    'End If

                    nCargaCBO(combo2, sql, "cNombre", "iIdDepartamento")

                    combo2.HeaderText = "Buque"
                    combo2.Width = 150
                    dtgDatos.Columns.Insert(12, combo2)

                    'Tipo_Infonavit
                    dtgDatos.Columns(13).ReadOnly = True
                    dtgDatos.Columns(13).Width = 150
                    'dtgDatos.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight



                    'Valor_Infonavit
                    dtgDatos.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(14).ReadOnly = True
                    dtgDatos.Columns(14).Width = 150
                    'Sueldo_Base
                    dtgDatos.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    'dtgDatos.Columns(15).ReadOnly = True
                    dtgDatos.Columns(15).Width = 150
                    'Salario_Diario
                    dtgDatos.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(16).ReadOnly = True
                    dtgDatos.Columns(16).Width = 150
                    'Salario_Cotización
                    dtgDatos.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(17).ReadOnly = True
                    dtgDatos.Columns(17).Width = 150
                    'Dias_Trabajados
                    dtgDatos.Columns(18).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(18).Width = 150
                    'Tipo_Incapacidad
                    dtgDatos.Columns(19).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(19).ReadOnly = True
                    dtgDatos.Columns(19).Width = 150
                    'Número_días
                    dtgDatos.Columns(20).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(20).ReadOnly = True
                    dtgDatos.Columns(20).Width = 150
                    'Sueldo_Bruto
                    dtgDatos.Columns(21).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(21).ReadOnly = True
                    dtgDatos.Columns(21).Width = 150
                    'Tiempo_Extra_Fijo_Gravado
                    dtgDatos.Columns(22).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(22).ReadOnly = True
                    dtgDatos.Columns(22).Width = 150

                    'Tiempo_Extra_Fijo_Exento
                    dtgDatos.Columns(23).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(23).ReadOnly = True
                    dtgDatos.Columns(23).Width = 150

                    'Tiempo_Extra_Ocasional
                    dtgDatos.Columns(24).Width = 150
                    dtgDatos.Columns(24).ReadOnly = True
                    dtgDatos.Columns(24).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    'Desc_Sem_Obligatorio
                    dtgDatos.Columns(25).Width = 150
                    dtgDatos.Columns(25).ReadOnly = True
                    dtgDatos.Columns(25).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    'Vacaciones_proporcionales
                    dtgDatos.Columns(26).Width = 150
                    dtgDatos.Columns(26).ReadOnly = True
                    dtgDatos.Columns(26).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    'Aguinaldo_gravado
                    dtgDatos.Columns(27).Width = 150
                    dtgDatos.Columns(27).ReadOnly = True
                    dtgDatos.Columns(27).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    'Aguinaldo_exento
                    dtgDatos.Columns(28).Width = 150
                    dtgDatos.Columns(28).ReadOnly = True
                    dtgDatos.Columns(28).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    'Total_Aguinaldo
                    dtgDatos.Columns(29).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(29).Width = 150
                    dtgDatos.Columns(29).ReadOnly = True

                    'Prima_vac_gravado
                    dtgDatos.Columns(30).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(30).ReadOnly = True
                    dtgDatos.Columns(30).Width = 150
                    'Prima_vac_exento 
                    dtgDatos.Columns(31).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(31).ReadOnly = True
                    dtgDatos.Columns(31).Width = 150

                    'Total_Prima_vac
                    dtgDatos.Columns(32).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(32).ReadOnly = True
                    dtgDatos.Columns(32).Width = 150


                    'Total_percepciones
                    dtgDatos.Columns(33).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(33).ReadOnly = True
                    dtgDatos.Columns(33).Width = 150
                    'Total_percepciones_p/isr
                    dtgDatos.Columns(34).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(34).ReadOnly = True
                    dtgDatos.Columns(34).Width = 150

                    'Incapacidad
                    dtgDatos.Columns(35).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(35).ReadOnly = True
                    dtgDatos.Columns(35).Width = 150

                    'ISR
                    dtgDatos.Columns(36).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(36).ReadOnly = True
                    dtgDatos.Columns(36).Width = 150


                    'IMSS
                    dtgDatos.Columns(37).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(37).ReadOnly = True
                    dtgDatos.Columns(37).Width = 150

                    'Infonavit
                    dtgDatos.Columns(38).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    'dtgDatos.Columns(38).ReadOnly = True
                    dtgDatos.Columns(38).Width = 150
                    'Infonavit_bim_anterior
                    dtgDatos.Columns(39).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    'dtgDatos.Columns(39).ReadOnly = True
                    dtgDatos.Columns(39).Width = 150
                    'Ajuste_infonavit
                    dtgDatos.Columns(40).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    'dtgDatos.Columns(40).ReadOnly = True
                    dtgDatos.Columns(40).Width = 150
                    'Pension_Alimenticia
                    dtgDatos.Columns(41).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    'dtgDatos.Columns(40).ReadOnly = True
                    dtgDatos.Columns(41).Width = 150
                    'Prestamo
                    dtgDatos.Columns(42).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    'dtgDatos.Columns(42).ReadOnly = True
                    dtgDatos.Columns(42).Width = 150
                    'Fonacot
                    dtgDatos.Columns(43).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    'dtgDatos.Columns(43).ReadOnly = True
                    dtgDatos.Columns(43).Width = 150
                    'Subsidio_Generado
                    dtgDatos.Columns(44).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(44).ReadOnly = True
                    dtgDatos.Columns(44).Width = 150
                    'Subsidio_Aplicado
                    dtgDatos.Columns(45).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(45).ReadOnly = True
                    dtgDatos.Columns(45).Width = 150
                    'Operadora
                    dtgDatos.Columns(46).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(46).ReadOnly = True
                    dtgDatos.Columns(46).Width = 150

                    'Prestamo Personal Asimilado
                    dtgDatos.Columns(47).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    'dtgDatos.Columns(48).ReadOnly = True
                    dtgDatos.Columns(47).Width = 150

                    'Adeudo_Infonavit_Asimilado
                    dtgDatos.Columns(48).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    'dtgDatos.Columns(49).ReadOnly = True
                    dtgDatos.Columns(48).Width = 150

                    'Difencia infonavit Asimilado
                    dtgDatos.Columns(49).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    'dtgDatos.Columns(50).ReadOnly = True
                    dtgDatos.Columns(49).Width = 150

                    'Complemento Asimilado
                    dtgDatos.Columns(50).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(50).ReadOnly = True
                    dtgDatos.Columns(50).Width = 150

                    'Retenciones_Operadora
                    dtgDatos.Columns(51).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(51).ReadOnly = True
                    dtgDatos.Columns(51).Width = 150

                    '% Comision
                    dtgDatos.Columns(52).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(52).ReadOnly = True
                    dtgDatos.Columns(52).Width = 150

                    'Comision_Operadora
                    dtgDatos.Columns(53).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(53).ReadOnly = True
                    dtgDatos.Columns(53).Width = 150

                    'Comision asimilados
                    dtgDatos.Columns(54).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(54).ReadOnly = True
                    dtgDatos.Columns(54).Width = 150

                    'IMSS_CS
                    dtgDatos.Columns(55).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(55).ReadOnly = True
                    dtgDatos.Columns(55).Width = 150

                    'RCV_CS
                    dtgDatos.Columns(56).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(56).ReadOnly = True
                    dtgDatos.Columns(56).Width = 150

                    'Infonavit_CS
                    dtgDatos.Columns(57).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(57).ReadOnly = True
                    dtgDatos.Columns(57).Width = 150

                    'ISN_CS
                    dtgDatos.Columns(58).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(58).ReadOnly = True
                    dtgDatos.Columns(58).Width = 150

                    'Total Costo Social
                    dtgDatos.Columns(59).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(59).ReadOnly = True
                    dtgDatos.Columns(59).Width = 150

                    'Subtotal
                    dtgDatos.Columns(60).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(60).ReadOnly = True
                    dtgDatos.Columns(60).Width = 150

                    'IVA
                    dtgDatos.Columns(61).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(61).ReadOnly = True
                    dtgDatos.Columns(61).Width = 150

                    'TOTAL DEPOSITO
                    dtgDatos.Columns(62).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(62).ReadOnly = True
                    dtgDatos.Columns(62).Width = 150
                    'calcular()

                    'calcular()

                    'Cambiamos index del combo en el grid




                    For x As Integer = 0 To dtgDatos.Rows.Count - 1

                        sql = "select * from nomina where fkiIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                        sql &= " and fkiIdPeriodo=" & cboperiodo.SelectedValue
                        sql &= " and iEstatusEmpleado=" & cboserie.SelectedIndex
                        sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex
                        Dim rwFila As DataRow() = nConsulta(sql)

                        If rwFila Is Nothing = False Then
                            CType(Me.dtgDatos.Rows(x).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("Puesto").ToString()
                            CType(Me.dtgDatos.Rows(x).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("Buque").ToString()
                        Else
                            sql = "select * from empleadosC where iIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                            Dim rwEmpleado As DataRow() = nConsulta(sql)



                            CType(Me.dtgDatos.Rows(x).Cells(11), DataGridViewComboBoxCell).Value = rwEmpleado(0)("cPuesto").ToString()
                            CType(Me.dtgDatos.Rows(x).Cells(12), DataGridViewComboBoxCell).Value = rwEmpleado(0)("cFuncionesPuesto").ToString()
                        End If



                    Next

                    MessageBox.Show("Datos cargados", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                Else



                    cadenaempleados = ""

                    For x As Integer = 0 To ids.Length - 1
                        If x = 0 Then
                            cadenaempleados = " iIdEmpleadoC=" & ids(x)
                        Else
                            cadenaempleados &= "  or iIdEmpleadoC=" & ids(x)
                        End If
                    Next






                    sql = "select  * from empleadosC where " 'fkiIdClienteInter=-1"
                    sql &= cadenaempleados
                    sql &= " order by cFuncionesPuesto,cNombreLargo"

                    Dim rwDatosEmpleados As DataRow() = nConsulta(sql)
                    If rwDatosEmpleados Is Nothing = False Then
                        For x As Integer = 0 To rwDatosEmpleados.Length - 1


                            Dim fila As DataRow = dsPeriodo.Tables("Tabla").NewRow

                            fila.Item("Consecutivo") = (x + 1).ToString
                            fila.Item("Id_empleado") = rwDatosEmpleados(x)("iIdEmpleadoC").ToString
                            fila.Item("CodigoEmpleado") = rwDatosEmpleados(x)("cCodigoEmpleado").ToString
                            fila.Item("Nombre") = rwDatosEmpleados(x)("cNombreLargo").ToString.ToUpper()
                            fila.Item("Status") = IIf(rwDatosEmpleados(x)("iOrigen").ToString = "1", "INTERINO", "PLANTA")
                            fila.Item("RFC") = rwDatosEmpleados(x)("cRFC").ToString
                            fila.Item("CURP") = rwDatosEmpleados(x)("cCURP").ToString
                            fila.Item("Num_IMSS") = rwDatosEmpleados(x)("cIMSS").ToString

                            fila.Item("Fecha_Nac") = Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString).ToShortDateString()
                            'Dim tiempo As TimeSpan = Date.Now - Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString)
                            fila.Item("Edad") = CalcularEdad(Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString).Day, Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString).Month, Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString).Year)
                            fila.Item("Puesto") = rwDatosEmpleados(x)("cPuesto").ToString
                            fila.Item("Buque") = "ECO III"

                            fila.Item("Tipo_Infonavit") = rwDatosEmpleados(x)("cTipoFactor").ToString
                            fila.Item("Valor_Infonavit") = rwDatosEmpleados(x)("fFactor").ToString
                            fila.Item("Sueldo_Base") = "0.00"
                            fila.Item("Salario_Diario") = rwDatosEmpleados(x)("fSueldoBase").ToString
                            fila.Item("Salario_Cotización") = rwDatosEmpleados(x)("fSueldoIntegrado").ToString
                            fila.Item("Dias_Trabajados") = "30"
                            fila.Item("Tipo_Incapacidad") = TipoIncapacidad(rwDatosEmpleados(x)("iIdEmpleadoC").ToString, cboperiodo.SelectedValue)
                            fila.Item("Número_días") = NumDiasIncapacidad(rwDatosEmpleados(x)("iIdEmpleadoC").ToString, cboperiodo.SelectedValue)
                            fila.Item("Sueldo_Bruto") = ""
                            fila.Item("Tiempo_Extra_Fijo_Gravado") = ""
                            fila.Item("Tiempo_Extra_Fijo_Exento") = ""
                            fila.Item("Tiempo_Extra_Ocasional") = ""
                            fila.Item("Desc_Sem_Obligatorio") = ""
                            fila.Item("Vacaciones_proporcionales") = ""
                            fila.Item("Aguinaldo_gravado") = ""
                            fila.Item("Aguinaldo_exento") = ""
                            fila.Item("Total_Aguinaldo") = ""
                            fila.Item("Prima_vac_gravado") = ""
                            fila.Item("Prima_vac_exento") = ""
                            fila.Item("Total_Prima_vac") = ""
                            fila.Item("Total_percepciones") = ""
                            fila.Item("Total_percepciones_p/isr") = ""
                            fila.Item("Incapacidad") = ""
                            fila.Item("ISR") = ""
                            fila.Item("IMSS") = ""
                            fila.Item("Infonavit") = ""
                            fila.Item("Infonavit_bim_anterior") = ""
                            fila.Item("Ajuste_infonavit") = ""
                            fila.Item("Pension_Alimenticia") = ""
                            fila.Item("Prestamo") = ""
                            fila.Item("Fonacot") = ""
                            fila.Item("Subsidio_Generado") = ""
                            fila.Item("Subsidio_Aplicado") = ""
                            fila.Item("Operadora") = ""
                            fila.Item("Prestamo_Personal_A") = ""
                            fila.Item("Adeudo_Infonavit_A") = ""
                            fila.Item("Diferencia_Infonavit_A") = ""
                            fila.Item("Asimilados") = ""
                            fila.Item("Retenciones_Operadora") = ""
                            fila.Item("%_Comisión") = ""
                            fila.Item("Comisión_Operadora") = ""
                            fila.Item("Comisión_Asimilados") = ""
                            fila.Item("IMSS_CS") = ""
                            fila.Item("RCV_CS") = ""
                            fila.Item("Infonavit_CS") = ""
                            fila.Item("ISN_CS") = ""
                            fila.Item("Total_Costo_Social") = ""
                            fila.Item("Subtotal") = ""
                            fila.Item("IVA") = ""
                            fila.Item("TOTAL_DEPOSITO") = ""


                            dsPeriodo.Tables("Tabla").Rows.Add(fila)




                        Next

                        dtgDatos.Columns.Clear()
                        Dim chk As New DataGridViewCheckBoxColumn()
                        dtgDatos.Columns.Add(chk)
                        chk.HeaderText = ""
                        chk.Name = "chk"
                        dtgDatos.DataSource = dsPeriodo.Tables("Tabla")

                        dtgDatos.Columns(0).Width = 30
                        dtgDatos.Columns(0).ReadOnly = True
                        dtgDatos.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                        'consecutivo
                        dtgDatos.Columns(1).Width = 60
                        dtgDatos.Columns(1).ReadOnly = True
                        dtgDatos.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'idempleado
                        dtgDatos.Columns(2).Width = 100
                        dtgDatos.Columns(2).ReadOnly = True
                        dtgDatos.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'codigo empleado
                        dtgDatos.Columns(3).Width = 100
                        dtgDatos.Columns(3).ReadOnly = True
                        dtgDatos.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Nombre
                        dtgDatos.Columns(4).Width = 250
                        dtgDatos.Columns(4).ReadOnly = True
                        'Estatus
                        dtgDatos.Columns(5).Width = 100
                        dtgDatos.Columns(5).ReadOnly = True
                        'RFC
                        dtgDatos.Columns(6).Width = 100
                        dtgDatos.Columns(6).ReadOnly = True
                        'dtgDatos.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                        'CURP
                        dtgDatos.Columns(7).Width = 150
                        dtgDatos.Columns(7).ReadOnly = True
                        'IMSS 

                        dtgDatos.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(8).ReadOnly = True
                        'Fecha_Nac
                        dtgDatos.Columns(9).Width = 150
                        dtgDatos.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(9).ReadOnly = True

                        'Edad
                        dtgDatos.Columns(10).ReadOnly = True
                        dtgDatos.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                        'Puesto
                        dtgDatos.Columns(11).ReadOnly = True
                        dtgDatos.Columns(11).Width = 200
                        dtgDatos.Columns.Remove("Puesto")

                        Dim combo As New DataGridViewComboBoxColumn

                        sql = "select * from puestos where iTipo=1 order by cNombre"

                        'Dim rwPuestos As DataRow() = nConsulta(sql)
                        'If rwPuestos Is Nothing = False Then
                        '    combo.Items.Add("uno")
                        '    combo.Items.Add("dos")
                        '    combo.Items.Add("tres")
                        'End If

                        nCargaCBO(combo, sql, "cNombre", "iIdPuesto")

                        combo.HeaderText = "Puesto"

                        combo.Width = 150
                        dtgDatos.Columns.Insert(11, combo)
                        'DirectCast(dtgDatos.Columns(11), DataGridViewComboBoxColumn).Sorted = True
                        'Dim combo2 As New DataGridViewComboBoxCell
                        'combo2 = CType(Me.dtgDatos.Rows(2).Cells(11), DataGridViewComboBoxCell)
                        'combo2.Value = combo.Items(11)



                        'dtgDatos.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                        'Buque
                        'dtgDatos.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(12).ReadOnly = True
                        dtgDatos.Columns(12).Width = 150
                        dtgDatos.Columns.Remove("Buque")

                        Dim combo2 As New DataGridViewComboBoxColumn

                        sql = "select * from departamentos where iEstatus=1 order by cNombre"

                        'Dim rwPuestos As DataRow() = nConsulta(sql)
                        'If rwPuestos Is Nothing = False Then
                        '    combo.Items.Add("uno")
                        '    combo.Items.Add("dos")
                        '    combo.Items.Add("tres")
                        'End If

                        nCargaCBO(combo2, sql, "cNombre", "iIdDepartamento")

                        combo2.HeaderText = "Buque"
                        combo2.Width = 150
                        dtgDatos.Columns.Insert(12, combo2)

                        'Tipo_Infonavit
                        dtgDatos.Columns(13).ReadOnly = True
                        dtgDatos.Columns(13).Width = 150
                        'dtgDatos.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight



                        'Valor_Infonavit
                        dtgDatos.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(14).ReadOnly = True
                        dtgDatos.Columns(14).Width = 150
                        'Sueldo_Base
                        dtgDatos.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(15).ReadOnly = True
                        dtgDatos.Columns(15).Width = 150
                        'Salario_Diario
                        dtgDatos.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(16).ReadOnly = True
                        dtgDatos.Columns(16).Width = 150
                        'Salario_Cotización
                        dtgDatos.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(17).ReadOnly = True
                        dtgDatos.Columns(17).Width = 150
                        'Dias_Trabajados
                        dtgDatos.Columns(18).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(18).Width = 150
                        'Tipo_Incapacidad
                        dtgDatos.Columns(19).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(19).ReadOnly = True
                        dtgDatos.Columns(19).Width = 150
                        'Número_días
                        dtgDatos.Columns(20).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(20).ReadOnly = True
                        dtgDatos.Columns(20).Width = 150
                        'Sueldo_Bruto
                        dtgDatos.Columns(21).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(21).ReadOnly = True
                        dtgDatos.Columns(21).Width = 150
                        'Tiempo_Extra_Fijo_Gravado
                        dtgDatos.Columns(22).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(22).ReadOnly = True
                        dtgDatos.Columns(22).Width = 150

                        'Tiempo_Extra_Fijo_Exento
                        dtgDatos.Columns(23).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(23).ReadOnly = True
                        dtgDatos.Columns(23).Width = 150

                        'Tiempo_Extra_Ocasional
                        dtgDatos.Columns(24).Width = 150
                        dtgDatos.Columns(24).ReadOnly = True
                        dtgDatos.Columns(24).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Desc_Sem_Obligatorio
                        dtgDatos.Columns(25).Width = 150
                        dtgDatos.Columns(25).ReadOnly = True
                        dtgDatos.Columns(25).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Vacaciones_proporcionales
                        dtgDatos.Columns(26).Width = 150
                        dtgDatos.Columns(26).ReadOnly = True
                        dtgDatos.Columns(26).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Aguinaldo_gravado
                        dtgDatos.Columns(27).Width = 150
                        dtgDatos.Columns(27).ReadOnly = True
                        dtgDatos.Columns(27).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Aguinaldo_exento
                        dtgDatos.Columns(28).Width = 150
                        dtgDatos.Columns(28).ReadOnly = True
                        dtgDatos.Columns(28).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Total_Aguinaldo
                        dtgDatos.Columns(29).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(29).Width = 150
                        dtgDatos.Columns(29).ReadOnly = True

                        'Prima_vac_gravado
                        dtgDatos.Columns(30).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(30).ReadOnly = True
                        dtgDatos.Columns(30).Width = 150
                        'Prima_vac_exento 
                        dtgDatos.Columns(31).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(31).ReadOnly = True
                        dtgDatos.Columns(31).Width = 150

                        'Total_Prima_vac
                        dtgDatos.Columns(32).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(32).ReadOnly = True
                        dtgDatos.Columns(32).Width = 150


                        'Total_percepciones
                        dtgDatos.Columns(33).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(33).ReadOnly = True
                        dtgDatos.Columns(33).Width = 150
                        'Total_percepciones_p/isr
                        dtgDatos.Columns(34).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(34).ReadOnly = True
                        dtgDatos.Columns(34).Width = 150

                        'Incapacidad
                        dtgDatos.Columns(35).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(35).ReadOnly = True
                        dtgDatos.Columns(35).Width = 150

                        'ISR
                        dtgDatos.Columns(36).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(36).ReadOnly = True
                        dtgDatos.Columns(36).Width = 150


                        'IMSS
                        dtgDatos.Columns(37).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(37).ReadOnly = True
                        dtgDatos.Columns(37).Width = 150

                        'Infonavit
                        dtgDatos.Columns(38).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(38).ReadOnly = True
                        dtgDatos.Columns(38).Width = 150
                        'Infonavit_bim_anterior
                        dtgDatos.Columns(39).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(39).ReadOnly = True
                        dtgDatos.Columns(39).Width = 150
                        'Ajuste_infonavit
                        dtgDatos.Columns(40).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(40).ReadOnly = True
                        dtgDatos.Columns(40).Width = 150
                        'Pension_Alimenticia
                        dtgDatos.Columns(41).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(40).ReadOnly = True
                        dtgDatos.Columns(41).Width = 150
                        'Prestamo
                        dtgDatos.Columns(42).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(42).ReadOnly = True
                        dtgDatos.Columns(42).Width = 150
                        'Fonacot
                        dtgDatos.Columns(43).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(43).ReadOnly = True
                        dtgDatos.Columns(43).Width = 150
                        'Subsidio_Generado
                        dtgDatos.Columns(44).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(44).ReadOnly = True
                        dtgDatos.Columns(44).Width = 150
                        'Subsidio_Aplicado
                        dtgDatos.Columns(45).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(45).ReadOnly = True
                        dtgDatos.Columns(45).Width = 150
                        'Operadora
                        dtgDatos.Columns(46).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(46).ReadOnly = True
                        dtgDatos.Columns(46).Width = 150

                        'Prestamo Personal Asimilado
                        dtgDatos.Columns(47).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(48).ReadOnly = True
                        dtgDatos.Columns(47).Width = 150

                        'Adeudo_Infonavit_Asimilado
                        dtgDatos.Columns(48).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(49).ReadOnly = True
                        dtgDatos.Columns(48).Width = 150

                        'Difencia infonavit Asimilado
                        dtgDatos.Columns(49).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(50).ReadOnly = True
                        dtgDatos.Columns(49).Width = 150

                        'Complemento Asimilado
                        dtgDatos.Columns(50).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(50).ReadOnly = True
                        dtgDatos.Columns(50).Width = 150

                        'Retenciones_Operadora
                        dtgDatos.Columns(51).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(51).ReadOnly = True
                        dtgDatos.Columns(51).Width = 150

                        '% Comision
                        dtgDatos.Columns(52).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(52).ReadOnly = True
                        dtgDatos.Columns(52).Width = 150

                        'Comision_Operadora
                        dtgDatos.Columns(53).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(53).ReadOnly = True
                        dtgDatos.Columns(53).Width = 150

                        'Comision asimilados
                        dtgDatos.Columns(54).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(54).ReadOnly = True
                        dtgDatos.Columns(54).Width = 150

                        'IMSS_CS
                        dtgDatos.Columns(55).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(55).ReadOnly = True
                        dtgDatos.Columns(55).Width = 150

                        'RCV_CS
                        dtgDatos.Columns(56).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(56).ReadOnly = True
                        dtgDatos.Columns(56).Width = 150

                        'Infonavit_CS
                        dtgDatos.Columns(57).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(57).ReadOnly = True
                        dtgDatos.Columns(57).Width = 150

                        'ISN_CS
                        dtgDatos.Columns(58).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(58).ReadOnly = True
                        dtgDatos.Columns(58).Width = 150

                        'Total Costo Social
                        dtgDatos.Columns(59).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(59).ReadOnly = True
                        dtgDatos.Columns(59).Width = 150

                        'Subtotal
                        dtgDatos.Columns(60).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(60).ReadOnly = True
                        dtgDatos.Columns(60).Width = 150

                        'IVA
                        dtgDatos.Columns(61).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(61).ReadOnly = True
                        dtgDatos.Columns(61).Width = 150

                        'TOTAL DEPOSITO
                        dtgDatos.Columns(62).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(62).ReadOnly = True
                        dtgDatos.Columns(62).Width = 150

                        'calcular()

                        'Cambiamos index del combo en el grid

                        For x As Integer = 0 To dtgDatos.Rows.Count - 1

                            sql = "select * from empleadosC where iIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                            Dim rwFila As DataRow() = nConsulta(sql)



                            CType(Me.dtgDatos.Rows(x).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("cPuesto").ToString()
                            CType(Me.dtgDatos.Rows(x).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("cFuncionesPuesto").ToString()
                        Next




                        MessageBox.Show("Datos cargados", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("No hay datos en este período", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If




                    'No hay datos en este período


                End If




                'MessageBox.Show("Trabajadores asignados", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                'If cboempresa.SelectedIndex > -1 Then
                '    cargarlista()
                'End If
                'lsvLista.SelectedItems(0).Tag = ""
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub



    Function MonthString(ByRef month As Integer) As String

        Select Case month
            Case 1 : Return "Enero"
            Case 2 : Return "Febrero"
            Case 3 : Return "Marzo"
            Case 4 : Return "Abril"
            Case 5 : Return "Mayo"
            Case 6 : Return "Junio"
            Case 7 : Return "Julio"
            Case 8 : Return "Agosto"
            Case 9 : Return "Septiembre"
            Case 10 : Return "Octubre"
            Case 11 : Return "Noviembre"
            Case 12, 0 : Return "Diciembre"

        End Select

    End Function


End Class



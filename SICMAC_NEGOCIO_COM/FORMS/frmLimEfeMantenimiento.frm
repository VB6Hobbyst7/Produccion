VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLimEfeMantenimiento 
   Caption         =   "Límite de Efectivo : Mantenimiento"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9360
   Icon            =   "frmLimEfeMantenimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Mantenimiento"
      TabPicture(0)   =   "frmLimEfeMantenimiento.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FECobertura"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdCancelar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdLimpiar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdGuardar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboAnio2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboMes2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboAnio1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cboMes1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lstPoliza"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      Begin VB.ListBox lstPoliza 
         Height          =   840
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   3495
      End
      Begin VB.ComboBox cboMes1 
         Height          =   315
         ItemData        =   "frmLimEfeMantenimiento.frx":0326
         Left            =   720
         List            =   "frmLimEfeMantenimiento.frx":0328
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ComboBox cboAnio1 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2040
         Width           =   1215
      End
      Begin VB.ComboBox cboMes2 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ComboBox cboAnio2 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   5760
         Width           =   975
      End
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "&Limpiar"
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   5760
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   5760
         Width           =   975
      End
      Begin SICMACT.FlexEdit FECobertura 
         Height          =   3135
         Left            =   120
         TabIndex        =   4
         Top             =   2520
         Width           =   9135
         _extentx        =   16113
         _extenty        =   5530
         cols0           =   11
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   $"frmLimEfeMantenimiento.frx":032A
         encabezadosanchos=   "500-2200-1500-1500-1500-1500-0-0-0-0-0"
         font            =   "frmLimEfeMantenimiento.frx":03B3
         font            =   "frmLimEfeMantenimiento.frx":03DF
         font            =   "frmLimEfeMantenimiento.frx":040B
         font            =   "frmLimEfeMantenimiento.frx":0437
         font            =   "frmLimEfeMantenimiento.frx":0463
         fontfixed       =   "frmLimEfeMantenimiento.frx":048F
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-2-3-4-5-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "L-L-R-R-R-R-C-C-C-C-C"
         formatosedit    =   "0-0-4-4-4-4-0-0-0-0-0"
         textarray0      =   "Nro"
         lbbuscaduplicadotext=   -1
         colwidth0       =   495
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin VB.Label Label4 
         Caption         =   "Seleccionar Póliza"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Editar Registro"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   4560
         TabIndex        =   9
         Top             =   2040
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmLimEfeMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''************************************************************
''***Nombre:         frmLimEfeMantenimiento
''***Descripción:    Formulario que permite el mantenimiento de Límite de Efectivo
''***Creación:       ELRO el 20120725, según OYP-RFC077-2012
''************************************************************
'Option Explicit
'
'Dim fnIdCobAgePer As Integer
'Dim fsMesDesIni As String
'Dim fsAnioDesIni As String
'Dim fsMesHasIni As String
'Dim fsAnioHasIni As String
'
'Private Sub cargarAgencias()
'Dim oDCOMGeneral As New COMDConstSistema.DCOMGeneral
'Dim rsAgencias As New ADODB.Recordset
'Dim i As Integer
'
'Set rsAgencias = oDCOMGeneral.devolverAgenciasParaCoberturar
'
'If Not rsAgencias.BOF And Not rsAgencias.EOF Then
'i = 1
'Call LimpiaFlex(FECobertura)
'FECobertura.lbEditarFlex = True
'    Do While Not rsAgencias.EOF
'            FECobertura.AdicionaFila
'            FECobertura.TextMatrix(i, 1) = rsAgencias!cAgeDescripcion
'            FECobertura.TextMatrix(i, TipoCobertura.gCobBovEfe + 1) = Format(rsAgencias!BovedaEfe, "#,##0.00")
'            FECobertura.TextMatrix(i, TipoCobertura.gCobVenEfe + 1) = Format(rsAgencias!VentanillaEfe, "#,##0.00")
'            FECobertura.TextMatrix(i, TipoCobertura.gCobBovOro + 1) = Format(rsAgencias!BovedaOro, "#,##0.00")
'            FECobertura.TextMatrix(i, TipoCobertura.gCobVenOro + 1) = Format(rsAgencias!VentanillaOro, "#,##0.00")
'            FECobertura.TextMatrix(i, 6) = rsAgencias!cAgeCod
'            FECobertura.TextMatrix(i, 7) = 0#
'            FECobertura.TextMatrix(i, 8) = 0#
'            FECobertura.TextMatrix(i, 9) = 0#
'            FECobertura.TextMatrix(i, 9) = 0#
'            FECobertura.TextMatrix(i, 10) = 0#
'        i = i + 1
'        rsAgencias.MoveNext
'    Loop
'Else
'    MsgBox "No se registraron las Agencias.", vbInformation, "Aviso"
'End If
'
'End Sub
'
'Private Sub cargarCoberturasAgencias(ByVal pnIdCobAgePer As Integer)
'Dim oDCOMGeneral As New COMDConstSistema.DCOMGeneral
'Dim rsAgencias As New ADODB.Recordset
'Dim i As Integer
'
'Set rsAgencias = oDCOMGeneral.devolverCoberturasAgencias(pnIdCobAgePer)
'
'If Not rsAgencias.BOF And Not rsAgencias.EOF Then
'i = 1
'Call LimpiaFlex(FECobertura)
'FECobertura.lbEditarFlex = True
'    Do While Not rsAgencias.EOF
'            FECobertura.AdicionaFila
'            FECobertura.TextMatrix(i, 1) = rsAgencias!cAgeDescripcion
'            FECobertura.TextMatrix(i, TipoCobertura.gCobBovEfe + 1) = Format(rsAgencias!nBovedaEfe, "#,##0.00")
'            FECobertura.TextMatrix(i, TipoCobertura.gCobVenEfe + 1) = Format(rsAgencias!nVentanillaEfe, "#,##0.00")
'            FECobertura.TextMatrix(i, TipoCobertura.gCobBovOro + 1) = Format(rsAgencias!nBovedaOro, "#,##0.00")
'            FECobertura.TextMatrix(i, TipoCobertura.gCobVenOro + 1) = Format(rsAgencias!nVentanillaOro, "#,##0.00")
'            FECobertura.TextMatrix(i, 6) = rsAgencias!cAgeCod
'            FECobertura.TextMatrix(i, 7) = rsAgencias!IdBovedaEfectivo
'            FECobertura.TextMatrix(i, 8) = rsAgencias!IdVentanillaEfectivo
'            FECobertura.TextMatrix(i, 9) = rsAgencias!IdBovedaOro
'            FECobertura.TextMatrix(i, 10) = rsAgencias!IdVentanillaOro
'        i = i + 1
'        rsAgencias.MoveNext
'    Loop
'Else
'    MsgBox "No se registraron las coberturas en el Periodo seleccionado.", vbInformation, "Aviso"
'End If
'
'End Sub
'
'Private Sub cargarPeriodos()
'Dim oDCOMGeneral As New COMDConstSistema.DCOMGeneral
'Dim rsPolizas As New ADODB.Recordset
'Dim i As Integer
'
'Set rsPolizas = oDCOMGeneral.devolverPeriodos
'
'lstPoliza.Clear
'
'If Not rsPolizas.BOF And Not rsPolizas.EOF Then
'    Do While Not rsPolizas.EOF
'        lstPoliza.AddItem rsPolizas!cPeriodo
'        rsPolizas.MoveNext
'    Loop
'End If
'
'End Sub
'
'Private Sub cargarAnios()
'Dim i As Integer
'
'cboAnio1.Clear
'cboAnio2.Clear
'
'For i = 2012 To 2050
'    cboAnio1.AddItem i
'    cboAnio2.AddItem i
'Next i
'
'cboAnio1.ListIndex = -1
'cboAnio2.ListIndex = -1
'
'End Sub
'
'Private Sub cargarMeses()
'Dim i As Integer
'
'cboMes1.Clear
'cboMes2.Clear
'
'For i = 1 To 12
'    cboMes1.AddItem IIf(i = 1, "Enero", IIf(i = 2, "Febrero", IIf(i = 3, "Marzo", IIf(i = 4, "Abril", IIf(i = 5, "Mayo", IIf(i = 6, "Junio", IIf(i = 7, "Julio", IIf(i = 8, "Agosto", IIf(i = 9, "Setiembre", IIf(i = 10, "Octubre", IIf(i = 11, "Noviembre", IIf(i = 12, "Diciembre", "")))))))))))) & Space(50) & i
'    cboMes2.AddItem IIf(i = 1, "Enero", IIf(i = 2, "Febrero", IIf(i = 3, "Marzo", IIf(i = 4, "Abril", IIf(i = 5, "Mayo", IIf(i = 6, "Junio", IIf(i = 7, "Julio", IIf(i = 8, "Agosto", IIf(i = 9, "Setiembre", IIf(i = 10, "Octubre", IIf(i = 11, "Noviembre", IIf(i = 12, "Diciembre", "")))))))))))) & Space(50) & i
'Next i
'
'cboAnio1.ListIndex = -1
'cboAnio2.ListIndex = -1
'
'End Sub
'
'
'Private Function validarCampos() As Boolean
'Dim i, J As Integer
'
'J = FECobertura.Rows
'
' For i = 1 To J - 1
'    If FECobertura.TextMatrix(i, TipoCobertura.gCobBovEfe + 1) = "" Or _
'       FECobertura.TextMatrix(i, TipoCobertura.gCobVenEfe + 1) = "" Or _
'       FECobertura.TextMatrix(i, TipoCobertura.gCobBovOro + 1) = "" Or _
'       FECobertura.TextMatrix(i, TipoCobertura.gCobVenOro + 1) = "" Then
'       MsgBox "No ingreso el monto de una Cobertura en la fila de la " & FECobertura.TextMatrix(i, 1) & ", debe corregir antes de Guardar.", vbInformation, "Aviso"
'       validarCampos = False
'       FECobertura.SetFocus
'       Exit Function
'    End If
' Next i
'
'
' For i = 1 To J - 1
'    If CCur(FECobertura.TextMatrix(i, TipoCobertura.gCobBovEfe + 1)) = 0# Or _
'       CCur(FECobertura.TextMatrix(i, TipoCobertura.gCobVenEfe + 1)) = 0# Or _
'       CCur(FECobertura.TextMatrix(i, TipoCobertura.gCobBovOro + 1)) = 0# Or _
'       CCur(FECobertura.TextMatrix(i, TipoCobertura.gCobVenOro + 1)) = 0# Then
'       MsgBox "El monto de una Cobertura es cero en la fila de la " & FECobertura.TextMatrix(i, 1) & ", debe corregir antes de Guardar.", vbInformation, "Aviso"
'       validarCampos = False
'       FECobertura.SetFocus
'       Exit Function
'    End If
' Next i
'
'  For i = 1 To J - 1
'    If CCur(FECobertura.TextMatrix(i, TipoCobertura.gCobBovEfe + 1)) < 0# Or _
'       CCur(FECobertura.TextMatrix(i, TipoCobertura.gCobVenEfe + 1)) < 0# Or _
'       CCur(FECobertura.TextMatrix(i, TipoCobertura.gCobBovOro + 1)) < 0# Or _
'       CCur(FECobertura.TextMatrix(i, TipoCobertura.gCobVenOro + 1)) < 0# Then
'       MsgBox "El monto de una Cobertura tiene valor negativo en la fila de la " & FECobertura.TextMatrix(i, 1) & ", debe corregir antes de Guardar.", vbInformation, "Aviso"
'       validarCampos = False
'       FECobertura.SetFocus
'       Exit Function
'    End If
' Next i
'
'  If Trim(cboMes1) = "" Then
'    MsgBox "No seleccionó el Mes inicial, debe corregir antes de Guardar.", vbInformation, "Aviso"
'    validarCampos = False
'    cboMes1.SetFocus
'    Exit Function
'  End If
'
'  If Trim(cboAnio1) = "" Then
'    MsgBox "No seleccionó el Año inicial, debe corregir antes de Guardar.", vbInformation, "Aviso"
'    validarCampos = False
'    cboAnio1.SetFocus
'    Exit Function
'  End If
'
'  If Trim(cboMes2) = "" Then
'    MsgBox "No seleccionó el Mes final, debe corregir antes de Guardar.", vbInformation, "Aviso"
'    validarCampos = False
'    cboMes1.SetFocus
'    Exit Function
'  End If
'
' If Trim(cboAnio2) = "" Then
'    MsgBox "No seleccionó el Año final, debe corregir antes de Guardar.", vbInformation, "Aviso"
'    validarCampos = False
'    cboAnio2.SetFocus
'    Exit Function
' End If
'
'
' If CInt(Right(cboAnio2, 2)) < CInt(Right(cboAnio1, 2)) Then
'    MsgBox "El Año final es menor que Año inicial, debe corregir antes de Guardar.", vbInformation, "Aviso"
'    validarCampos = False
'    cboAnio2.SetFocus
'    Exit Function
' End If
'
' If CInt(Right(cboAnio2, 2)) = CInt(Right(cboAnio1, 2)) And CInt(Right(cboMes2, 2)) < CInt(Right(cboMes1, 2)) Then
'    MsgBox "El Mes final es menor que Mes inicial, debe corregir antes de Guardar.", vbInformation, "Aviso"
'    validarCampos = False
'    cboMes2.SetFocus
'    Exit Function
' End If
'
' validarCampos = True
'End Function
'
'Private Sub cmdCancelar_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdGuardar_Click()
'Dim oDCOMGeneral As New COMDConstSistema.DCOMGeneral
'Dim oNCOMContFunciones As New COMNContabilidad.NCOMContFunciones
'Dim oPista As New COMManejador.Pista
'Dim lbRango As Boolean
'
'
'If validarCampos = False Then Exit Sub
'
'
'If Trim(Right(cboMes1, 2)) <> fsMesDesIni Or _
'   Trim(Right(cboAnio1, 4)) <> fsAnioDesIni Or _
'   Trim(Right(cboMes2, 2)) <> fsMesHasIni Or _
'   Trim(Right(cboAnio2, 4)) <> fsAnioHasIni Then
'
'    lbRango = oDCOMGeneral.verificarRangoCoberturaAgencia(CInt(Right(cboMes1, 2)), _
'                                                          CInt(Right(cboAnio1, 4)), _
'                                                          CInt(Right(cboMes2, 2)), _
'                                                          CInt(Right(cboAnio2, 4)), _
'                                                          False, _
'                                                          fnIdCobAgePer)
'
'    If lbRango = False Then
'        MsgBox "Parametro de fecha seleccionado ya fue registrado.", vbInformation, "Aviso"
'        Exit Sub
'    End If
'End If
'
'If MsgBox("¿Esta seguro que desea Modificar?", vbYesNo, "Aviso") = vbYes Then
'
'    Dim lbConfirmar As Boolean
'    Dim i, J, lnIdCobAgePer, lnIdCoberturaAgencia1, lnIdCoberturaAgencia2, lnIdCoberturaAgencia3, lnIdCoberturaAgencia4 As Integer
'    Dim lsMovNro As String
'
'    J = FECobertura.Rows
'    lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'
'    lnIdCobAgePer = oDCOMGeneral.actualizarCoberturaAgenciaPeriodo(CInt(Right(cboMes1, 2)), _
'                                                                   CInt(Right(cboAnio1, 4)), _
'                                                                   CInt(Right(cboMes2, 2)), _
'                                                                   CInt(Right(cboAnio2, 4)), _
'                                                                   lsMovNro, fnIdCobAgePer)
'    If lnIdCobAgePer > 0 Then
'        For i = 1 To J - 1
'
'
'           If CInt(FECobertura.TextMatrix(i, 7)) > 0 Then
'            lnIdCoberturaAgencia1 = oDCOMGeneral.actualizarCoberturaAgencia(FECobertura.TextMatrix(i, 6), _
'                                                                            TipoCobertura.gCobBovEfe, _
'                                                                            lnIdCobAgePer, _
'                                                                            Moneda.gMonedaExtranjera, _
'                                                                            FECobertura.TextMatrix(i, TipoCobertura.gCobBovEfe + 1), _
'                                                                            lsMovNro, _
'                                                                            CInt(FECobertura.TextMatrix(i, 7)))
'            oPista.InsertarPista "Limites de Efectivo>>Mantenimiento", lsMovNro, gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gModificar, FECobertura.TextMatrix(i, 1) & " Bóveda para Efectivo: " & FECobertura.TextMatrix(i, TipoCobertura.gCobBovEfe + 1), lnIdCoberturaAgencia1, CodigosIdentificacionPistas.gCodigo
'           Else
'            lnIdCoberturaAgencia1 = oDCOMGeneral.registrarCoberturaAgencia(FECobertura.TextMatrix(i, 6), _
'                                                                          TipoCobertura.gCobBovEfe, _
'                                                                          lnIdCobAgePer, _
'                                                                          Moneda.gMonedaExtranjera, _
'                                                                          FECobertura.TextMatrix(i, TipoCobertura.gCobBovEfe + 1), _
'                                                                          lsMovNro)
'           oPista.InsertarPista "Limites de Efectivo>>Registro", lsMovNro, gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gInsertar, FECobertura.TextMatrix(i, 1) & " Bóveda para Efectivo: " & FECobertura.TextMatrix(i, TipoCobertura.gCobBovEfe + 1), lnIdCoberturaAgencia1, CodigosIdentificacionPistas.gCodigo
'
'           End If
'
'           If CInt(FECobertura.TextMatrix(i, 8)) > 0 Then
'            lnIdCoberturaAgencia2 = oDCOMGeneral.actualizarCoberturaAgencia(FECobertura.TextMatrix(i, 6), _
'                                                                           TipoCobertura.gCobVenEfe, _
'                                                                           lnIdCobAgePer, _
'                                                                           Moneda.gMonedaExtranjera, _
'                                                                           FECobertura.TextMatrix(i, TipoCobertura.gCobVenEfe + 1), _
'                                                                           lsMovNro, _
'                                                                           CInt(FECobertura.TextMatrix(i, 8)))
'            oPista.InsertarPista "Limites de Efectivo>>Mantenimiento", lsMovNro, gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gModificar, FECobertura.TextMatrix(i, 1) & " Ventanilla para Efectivo: " & FECobertura.TextMatrix(i, TipoCobertura.gCobVenEfe + 1), lnIdCoberturaAgencia2, CodigosIdentificacionPistas.gCodigo
'           Else
'            lnIdCoberturaAgencia2 = oDCOMGeneral.registrarCoberturaAgencia(FECobertura.TextMatrix(i, 6), _
'                                                                          TipoCobertura.gCobVenEfe, _
'                                                                          lnIdCobAgePer, _
'                                                                          Moneda.gMonedaExtranjera, _
'                                                                          FECobertura.TextMatrix(i, TipoCobertura.gCobVenEfe + 1), _
'                                                                          lsMovNro)
'            oPista.InsertarPista "Limites de Efectivo>>Registro", lsMovNro, gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gInsertar, FECobertura.TextMatrix(i, 1) & " Ventanilla para Efectivo: " & FECobertura.TextMatrix(i, TipoCobertura.gCobVenEfe + 1), lnIdCoberturaAgencia2, CodigosIdentificacionPistas.gCodigo
'
'           End If
'
'           If CInt(FECobertura.TextMatrix(i, 9)) > 0 Then
'            lnIdCoberturaAgencia3 = oDCOMGeneral.actualizarCoberturaAgencia(FECobertura.TextMatrix(i, 6), _
'                                                                           TipoCobertura.gCobBovOro, _
'                                                                           lnIdCobAgePer, _
'                                                                           Moneda.gMonedaExtranjera, _
'                                                                           FECobertura.TextMatrix(i, TipoCobertura.gCobBovOro + 1), _
'                                                                           lsMovNro, _
'                                                                           CInt(FECobertura.TextMatrix(i, 9)))
'            oPista.InsertarPista "Limites de Efectivo>>Mantenimiento", lsMovNro, gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gModificar, FECobertura.TextMatrix(i, 1) & " Bóveda para Oro: " & FECobertura.TextMatrix(i, TipoCobertura.gCobBovOro + 1), lnIdCoberturaAgencia3, CodigosIdentificacionPistas.gCodigo
'           Else
'            lnIdCoberturaAgencia3 = oDCOMGeneral.registrarCoberturaAgencia(FECobertura.TextMatrix(i, 6), _
'                                                                          TipoCobertura.gCobBovOro, _
'                                                                          lnIdCobAgePer, _
'                                                                          Moneda.gMonedaExtranjera, _
'                                                                          FECobertura.TextMatrix(i, TipoCobertura.gCobBovOro + 1), _
'                                                                          lsMovNro)
'           oPista.InsertarPista "Limites de Efectivo>>Registro", lsMovNro, gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gInsertar, FECobertura.TextMatrix(i, 1) & " Bóveda para Oro: " & FECobertura.TextMatrix(i, TipoCobertura.gCobBovOro + 1), lnIdCoberturaAgencia3, CodigosIdentificacionPistas.gCodigo
'
'           End If
'
'           If CInt(FECobertura.TextMatrix(i, 10)) > 0 Then
'            lnIdCoberturaAgencia4 = oDCOMGeneral.actualizarCoberturaAgencia(FECobertura.TextMatrix(i, 6), _
'                                                                           TipoCobertura.gCobVenOro, _
'                                                                           lnIdCobAgePer, _
'                                                                           Moneda.gMonedaExtranjera, _
'                                                                           FECobertura.TextMatrix(i, TipoCobertura.gCobVenOro + 1), _
'                                                                           lsMovNro, _
'                                                                           CInt(FECobertura.TextMatrix(i, 10)))
'            oPista.InsertarPista "Limites de Efectivo>>Mantenimiento", lsMovNro, gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gModificar, FECobertura.TextMatrix(i, 1) & " Ventanilla para Oro: " & FECobertura.TextMatrix(i, TipoCobertura.gCobVenOro + 1), lnIdCoberturaAgencia4, CodigosIdentificacionPistas.gCodigo
'           Else
'            lnIdCoberturaAgencia4 = oDCOMGeneral.registrarCoberturaAgencia(FECobertura.TextMatrix(i, 6), _
'                                                                          TipoCobertura.gCobVenOro, _
'                                                                          lnIdCobAgePer, _
'                                                                          Moneda.gMonedaExtranjera, _
'                                                                          FECobertura.TextMatrix(i, TipoCobertura.gCobVenOro + 1), _
'                                                                          lsMovNro)
'            oPista.InsertarPista "Limites de Efectivo>>Registro", lsMovNro, gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gInsertar, FECobertura.TextMatrix(i, 1) & " Ventanilla para Oro: " & FECobertura.TextMatrix(i, TipoCobertura.gCobVenOro + 1), lnIdCoberturaAgencia4, CodigosIdentificacionPistas.gCodigo
'
'           End If
'
'           If lnIdCoberturaAgencia1 = 0 Or lnIdCoberturaAgencia2 = 0 Or _
'              lnIdCoberturaAgencia3 = 0 Or lnIdCoberturaAgencia4 = 0 Then
'                MsgBox "No se pudo modificar todas las Coberturas del Seguro.", vbInformation, "Aviso"
'                cargarAgencias
'                cargarAnios
'                cargarMeses
'                Exit Sub
'           End If
'
'           lnIdCoberturaAgencia1 = 0
'           lnIdCoberturaAgencia2 = 0
'           lnIdCoberturaAgencia3 = 0
'           lnIdCoberturaAgencia4 = 0
'
'        Next i
'
'        MsgBox "Se modificaron correctamente los datos.", vbInformation, "Aviso"
'        fnIdCobAgePer = 0
'        fsMesDesIni = ""
'        fsAnioDesIni = ""
'        fsMesHasIni = ""
'        fsAnioHasIni = ""
'        cargarAgencias
'        cargarAnios
'        cargarMeses
'        cargarPeriodos
'    Else
'        MsgBox "No se pudo registrar las Coberturas del Seguro.", vbInformation, "Aviso"
'        fnIdCobAgePer = 0
'        fsMesDesIni = ""
'        fsAnioDesIni = ""
'        fsMesHasIni = ""
'        fsAnioHasIni = ""
'        cargarAgencias
'        cargarAnios
'        cargarMeses
'        cargarPeriodos
'        Exit Sub
'    End If
'
'End If
'
'Set oNCOMContFunciones = Nothing
'Set oPista = Nothing
'Set oDCOMGeneral = Nothing
'End Sub
'
'Private Sub cmdLimpiar_Click()
'    fnIdCobAgePer = 0
'    fsMesDesIni = ""
'    fsAnioDesIni = ""
'    fsMesHasIni = ""
'    fsAnioHasIni = ""
'    cargarPeriodos
'    cargarAgencias
'    cargarAnios
'    cargarMeses
'End Sub
'
'Private Sub Form_Load()
'    fnIdCobAgePer = 0
'    fsMesDesIni = ""
'    fsAnioDesIni = ""
'    fsMesHasIni = ""
'    fsAnioHasIni = ""
'    cargarPeriodos
'    cargarAgencias
'    cargarAnios
'    cargarMeses
'End Sub
'
'Private Sub lstPoliza_Click()
'Dim lsCodigos As String
'Dim lsIdCobAgePer As String
'Dim lsMesAnioDes As String
'Dim lsMesAnioHas As String
'Dim i, J, k, L As Integer
'
'
''Dim fsMesDesIni As String
''Dim fsAnioDesIni As String
''Dim fsMesHasIni As String
''Dim fsAnioHasIni As String
'
'
'lsCodigos = Trim(Right(lstPoliza.Text, 40))
'fnIdCobAgePer = CInt(Mid(lsCodigos, 1, InStr(1, lsCodigos, "-") - 1))
'lsMesAnioDes = Mid(lsCodigos, InStr(1, lsCodigos, "-") + 1, InStr(1, lsCodigos, ",") - 3)
'lsMesAnioHas = Right(lsCodigos, InStr(1, lsCodigos, ",") - 3)
'cargarCoberturasAgencias (fnIdCobAgePer)
'
'For i = 0 To cboMes1.ListCount - 1
'    cboMes1.ListIndex = i
'    If CInt(Trim(Right(cboMes1.Text, 3))) = CInt(Trim(Left(lsMesAnioDes, InStr(1, lsMesAnioDes, ".") - 1))) Then
'        fsMesDesIni = Trim(Left(lsMesAnioDes, InStr(1, lsMesAnioDes, ".") - 1))
'        cboMes1.ListIndex = i
'        Exit For
'    End If
'Next i
'For J = 0 To cboAnio1.ListCount - 1
'    cboAnio1.ListIndex = J
'    If CInt(Trim(Right(cboAnio1.Text, 4))) = CInt(Trim(Right(lsMesAnioDes, 4))) Then
'        fsAnioDesIni = Trim(Right(lsMesAnioDes, 4))
'        cboAnio1.ListIndex = J
'        Exit For
'    End If
'Next J
'For k = 0 To cboMes2.ListCount - 1
'    cboMes2.ListIndex = k
'    If CInt(Trim(Right(cboMes2.Text, 3))) = CInt(Trim(Left(lsMesAnioHas, InStr(1, lsMesAnioDes, ".") - 1))) Then
'        fsMesHasIni = Trim(Left(lsMesAnioHas, InStr(1, lsMesAnioDes, ".") - 1))
'        cboMes2.ListIndex = k
'        Exit For
'    End If
'Next k
'For L = 0 To cboAnio2.ListCount - 1
'    cboAnio2.ListIndex = L
'    If CInt(Trim(Right(cboAnio2.Text, 4))) = CInt(Trim(Right(lsMesAnioHas, 4))) Then
'        fsAnioHasIni = Trim(Right(lsMesAnioHas, 4))
'        cboAnio2.ListIndex = L
'        Exit For
'    End If
'Next L
'
'End Sub

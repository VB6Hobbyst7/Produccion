VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLimEfeRegistro 
   Caption         =   "Límite de Efectivo : Registro"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   Icon            =   "frmLimEfeRegistro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9375
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
      TabCaption(0)   =   "Registro"
      TabPicture(0)   =   "frmLimEfeRegistro.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FECobertura"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboMes1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboAnio1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboMes2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboAnio2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdGuardar"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdLimpiar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdCancelar"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   5760
         Width           =   975
      End
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "&Limpiar"
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   5760
         Width           =   975
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   5760
         Width           =   975
      End
      Begin VB.ComboBox cboAnio2 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cboMes2 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox cboAnio1 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cboMes1 
         Height          =   315
         ItemData        =   "frmLimEfeRegistro.frx":0326
         Left            =   720
         List            =   "frmLimEfeRegistro.frx":0328
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin SICMACT.FlexEdit FECobertura 
         Height          =   4335
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   9135
         _extentx        =   16113
         _extenty        =   7646
         cols0           =   7
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "Nro-Agencia-Bóveda $-Ventanilla $-Bóveda Oro-Ventanilla Oro-cAgeCod"
         encabezadosanchos=   "500-2200-1500-1500-1500-1500-0"
         font            =   "frmLimEfeRegistro.frx":032A
         font            =   "frmLimEfeRegistro.frx":0356
         font            =   "frmLimEfeRegistro.frx":0382
         font            =   "frmLimEfeRegistro.frx":03AE
         font            =   "frmLimEfeRegistro.frx":03DA
         fontfixed       =   "frmLimEfeRegistro.frx":0406
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-2-3-4-5-X"
         listacontroles  =   "0-0-0-0-0-0-0"
         encabezadosalineacion=   "L-L-R-R-R-R-C"
         formatosedit    =   "0-0-4-4-4-4-0"
         textarray0      =   "Nro"
         lbbuscaduplicadotext=   -1
         colwidth0       =   495
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   4560
         TabIndex        =   5
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Nuevo Registro"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmLimEfeRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''************************************************************
''***Nombre:         frmLimEfeRegistro
''***Descripción:    Formulario que permite el registro de Límite de Efectivo
''***Creación:       ELRO el 20120724, según OYP-RFC077-2012
''************************************************************
'Option Explicit
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
'        i = i + 1
'        rsAgencias.MoveNext
'    Loop
'Else
'    MsgBox "No se registraron las Agencias.", vbInformation, "Aviso"
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
'Dim i, j As Integer
'
'j = FECobertura.Rows
'
' For i = 1 To j - 1
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
' For i = 1 To j - 1
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
'  For i = 1 To j - 1
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
'Unload Me
'End Sub
'
'Private Sub cmdGuardar_Click()
'Dim oDCOMGeneral As New COMDConstSistema.DCOMGeneral
'Dim oNCOMContFunciones As New COMNContabilidad.NCOMContFunciones
'Dim oPista As New COMManejador.Pista
'Dim lbRango As Boolean
'
'If validarCampos = False Then Exit Sub
'
'lbRango = oDCOMGeneral.verificarRangoCoberturaAgencia(CInt(Right(cboMes1, 2)), _
'                                                      CInt(Right(cboAnio1, 4)), _
'                                                      CInt(Right(cboMes2, 2)), _
'                                                      CInt(Right(cboAnio2, 4)))
'
'If lbRango = False Then
'    MsgBox "Parametro de fecha seleccionado ya fue registrado.", vbInformation, "Aviso"
'    Exit Sub
'End If
'
'If MsgBox("¿Esta seguro que desea Guardar?", vbYesNo, "Aviso") = vbYes Then
'
'    Dim lbConfirmar As Boolean
'    Dim i, j, lnIdCobAgePer, lnIdCoberturaAgencia1, lnIdCoberturaAgencia2, lnIdCoberturaAgencia3, lnIdCoberturaAgencia4 As Integer
'    Dim lsMovNro As String
'
'    j = FECobertura.Rows
'    lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'
'    lnIdCobAgePer = oDCOMGeneral.registrarPeriodoCoberturaAgencia(CInt(Right(cboMes1, 2)), _
'                                                                  CInt(Right(cboAnio1, 4)), _
'                                                                  CInt(Right(cboMes2, 2)), _
'                                                                  CInt(Right(cboAnio2, 4)), _
'                                                                  lsMovNro)
'    If lnIdCobAgePer > 0 Then
'        For i = 1 To j - 1
'
'           lnIdCoberturaAgencia1 = oDCOMGeneral.registrarCoberturaAgencia(FECobertura.TextMatrix(i, 6), _
'                                                                          TipoCobertura.gCobBovEfe, _
'                                                                          lnIdCobAgePer, _
'                                                                          Moneda.gMonedaExtranjera, _
'                                                                          FECobertura.TextMatrix(i, TipoCobertura.gCobBovEfe + 1), _
'                                                                          lsMovNro)
'           oPista.InsertarPista "Limites de Efectivo>>Registro", lsMovNro, gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gInsertar, FECobertura.TextMatrix(i, 1) & " Bóveda para Efectivo: " & FECobertura.TextMatrix(i, TipoCobertura.gCobBovEfe + 1), lnIdCoberturaAgencia1, CodigosIdentificacionPistas.gCodigo
'
'           lnIdCoberturaAgencia2 = oDCOMGeneral.registrarCoberturaAgencia(FECobertura.TextMatrix(i, 6), _
'                                                                          TipoCobertura.gCobVenEfe, _
'                                                                          lnIdCobAgePer, _
'                                                                          Moneda.gMonedaExtranjera, _
'                                                                          FECobertura.TextMatrix(i, TipoCobertura.gCobVenEfe + 1), _
'                                                                          lsMovNro)
'           oPista.InsertarPista "Limites de Efectivo>>Registro", lsMovNro, gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gInsertar, FECobertura.TextMatrix(i, 1) & " Ventanilla para Efectivo: " & FECobertura.TextMatrix(i, TipoCobertura.gCobVenEfe + 1), lnIdCoberturaAgencia2, CodigosIdentificacionPistas.gCodigo
'
'           lnIdCoberturaAgencia3 = oDCOMGeneral.registrarCoberturaAgencia(FECobertura.TextMatrix(i, 6), _
'                                                                          TipoCobertura.gCobBovOro, _
'                                                                          lnIdCobAgePer, _
'                                                                          Moneda.gMonedaExtranjera, _
'                                                                          FECobertura.TextMatrix(i, TipoCobertura.gCobBovOro + 1), _
'                                                                          lsMovNro)
'           oPista.InsertarPista "Limites de Efectivo>>Registro", lsMovNro, gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gInsertar, FECobertura.TextMatrix(i, 1) & " Bóveda para Oro: " & FECobertura.TextMatrix(i, TipoCobertura.gCobBovOro + 1), lnIdCoberturaAgencia3, CodigosIdentificacionPistas.gCodigo
'
'           lnIdCoberturaAgencia4 = oDCOMGeneral.registrarCoberturaAgencia(FECobertura.TextMatrix(i, 6), _
'                                                                          TipoCobertura.gCobVenOro, _
'                                                                          lnIdCobAgePer, _
'                                                                          Moneda.gMonedaExtranjera, _
'                                                                          FECobertura.TextMatrix(i, TipoCobertura.gCobVenOro + 1), _
'                                                                          lsMovNro)
'           oPista.InsertarPista "Limites de Efectivo>>Registro", lsMovNro, gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gInsertar, FECobertura.TextMatrix(i, 1) & " Ventanilla para Oro: " & FECobertura.TextMatrix(i, TipoCobertura.gCobVenOro + 1), lnIdCoberturaAgencia4, CodigosIdentificacionPistas.gCodigo
'
'            If lnIdCoberturaAgencia1 = 0 Or lnIdCoberturaAgencia2 = 0 Or _
'               lnIdCoberturaAgencia3 = 0 Or lnIdCoberturaAgencia4 = 0 Then
'                oDCOMGeneral.eliminarCoberturasAgencias lnIdCobAgePer
'                oDCOMGeneral.eliminarCoberturaAgenciaPeriodo lnIdCobAgePer
'                MsgBox "No se pudo registrar las Coberturas del Seguro.", vbInformation, "Aviso"
'                cargarAgencias
'                cargarAnios
'                cargarMeses
'                Exit Sub
'            End If
'
'            lnIdCoberturaAgencia1 = 0
'            lnIdCoberturaAgencia2 = 0
'            lnIdCoberturaAgencia3 = 0
'            lnIdCoberturaAgencia4 = 0
'
'        Next i
'
'        MsgBox "Se Guardaron correctamente los datos.", vbInformation, "Aviso"
'        cargarAgencias
'        cargarAnios
'        cargarMeses
'    Else
'        MsgBox "No se pudo registrar las Coberturas del Seguro.", vbInformation, "Aviso"
'        cargarAgencias
'        cargarAnios
'        cargarMeses
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
'    cargarAgencias
'    cargarAnios
'    cargarMeses
'End Sub
'
'Private Sub Form_Load()
'    cargarAgencias
'    cargarAnios
'    cargarMeses
'End Sub
'
'

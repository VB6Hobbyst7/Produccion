VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCredGeneraReportePagoVarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Pagos de Servicios x Institucion"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10290
   Icon            =   "frmCredGeneraReportePagoVarios.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.OptionButton Option5 
         Caption         =   "Dia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         TabIndex        =   3
         Top             =   600
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Mes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4920
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox cboInstitucion 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   4695
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000000&
         Height          =   615
         Left            =   6360
         TabIndex        =   8
         Top             =   240
         Width           =   3495
         Begin VB.CommandButton cmdbuscar 
            Caption         =   "Buscar"
            Height          =   350
            Left            =   120
            TabIndex        =   5
            Top             =   160
            Width           =   975
         End
         Begin VB.CommandButton cmdimprimir 
            Caption         =   "Imprimir "
            Height          =   350
            Left            =   1200
            TabIndex        =   6
            Top             =   160
            Width           =   975
         End
         Begin VB.CommandButton cmdsalir 
            Caption         =   "Salir"
            Height          =   350
            Left            =   2280
            TabIndex        =   7
            Top             =   160
            Width           =   975
         End
      End
      Begin SICMACT.FlexEdit FEHojaEval 
         Height          =   4455
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   9975
         _extentx        =   17595
         _extenty        =   7858
         cols0           =   8
         highlight       =   1
         allowuserresizing=   3
         encabezadosnombres=   "#-Doc.Tpo-Doc.Num-Nombres-Monto-DPago-Recibo-Concepto"
         encabezadosanchos=   "400-800-1000-3000-1200-1200-1500-2000"
         font            =   "frmCredGeneraReportePagoVarios.frx":030A
         font            =   "frmCredGeneraReportePagoVarios.frx":0336
         font            =   "frmCredGeneraReportePagoVarios.frx":0362
         font            =   "frmCredGeneraReportePagoVarios.frx":038E
         font            =   "frmCredGeneraReportePagoVarios.frx":03BA
         fontfixed       =   "frmCredGeneraReportePagoVarios.frx":03E6
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-L-L-L-L-L-L-L"
         formatosedit    =   "0-0-0-0-0-0-0-0"
         textarray0      =   "#"
         lbpuntero       =   -1
         colwidth0       =   405
         rowheight0      =   300
      End
      Begin MSMask.MaskEdBox txtperiodo 
         Height          =   300
         Left            =   1440
         TabIndex        =   2
         Top             =   600
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Vigencia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblTipoEval 
         AutoSize        =   -1  'True
         Caption         =   "Institución:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label4 
         Caption         =   "Periodo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCredGeneraReportePagoVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pInst As String

Private Sub CmdBuscar_Click()
Dim oPers As COMDPersona.DCOMPersonas
Dim rs As ADODB.Recordset
Dim ind As Integer

    Set oPers = New COMDPersona.DCOMPersonas
    If Trim(Right(Me.cboInstitucion.Text, 13)) = "" Then
        MsgBox "Debe especificar una Institución para Iniciar búsqueda", vbInformation
        Exit Sub
    End If
    
     If Me.cboInstitucion.ListIndex = 0 Or Me.cboInstitucion.ListIndex = -1 Then
        MsgBox "Seleccione una Institución, Verifique !!", vbInformation, "Aviso"
        Exit Sub
     End If
     
     If Me.txtperiodo.Text = "" Then
        MsgBox "Fecha no válida, Verifique !!", vbInformation, "Aviso"
        Exit Sub
     End If
     
     If Not IsDate(txtperiodo) Then
        MsgBox "Fecha no válida, Verifique !!", vbInformation, "Aviso"
        Exit Sub
     End If

    If Option5.value Then 'dia
        ind = 1
    Else
        ind = 2
    End If
    
    Set rs = oPers.CargaDatosReportePagoServiciosxDoc(Trim(Right(Me.cboInstitucion.Text, 13)), Me.txtperiodo, ind)
        FEHojaEval.Clear
        FEHojaEval.Rows = 2
        FEHojaEval.FormaCabecera
        FEHojaEval.FormateaColumnas
        FEHojaEval.TextMatrix(1, 0) = "1"
    
    If Not (rs.BOF And rs.EOF) Then
            Dim i As Integer
            i = 0
            Call LimpiaFlex(FEHojaEval)
'            ConfigurarMSH
            Do Until rs.EOF
            With Me.FEHojaEval
                 i = i + 1
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 2, 0) = i 'rsX!id
                .TextMatrix(.Rows - 2, 1) = rs!nTipoDoc
                .TextMatrix(.Rows - 2, 2) = rs!cNumDoc
                .TextMatrix(.Rows - 2, 3) = rs!cNombre
                .TextMatrix(.Rows - 2, 4) = rs!Pago
                .TextMatrix(.Rows - 2, 5) = rs!dFechaPag
                .TextMatrix(.Rows - 2, 6) = rs!cCodServicio
                .TextMatrix(.Rows - 2, 7) = rs!cConcepto
            End With
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Else
        MsgBox "No existen Datos con la fecha Indicada, Verifique", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If
    pInst = Trim(Right(Me.cboInstitucion.Text, 13))
Set oPers = Nothing
End Sub

Sub ConfigurarMSH()
    FEHojaEval.Cols = 8
    FEHojaEval.Rows = 2

    With FEHojaEval
        .TextMatrix(0, 0) = "Id"
        .TextMatrix(0, 1) = "TipDoc"
        .TextMatrix(0, 2) = "NumDoc"
        .TextMatrix(0, 3) = "Nombre Cliente"
        .TextMatrix(0, 4) = "MontoPago"
        .TextMatrix(0, 5) = "Fecha Pago"
        .TextMatrix(0, 6) = "cCodServicio"
        .TextMatrix(0, 7) = "Concepto"

        .ColWidth(0) = 400
        .ColWidth(1) = 800
        .ColWidth(2) = 1000
        .ColWidth(3) = 3000
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 1800
        .ColWidth(7) = 2500

    End With
End Sub

Private Sub cmdImprimir_Click()
Dim sCadImp As String
    Dim oPrev As previo.clsprevio
    Dim oNCred As COMNCredito.NCOMCredito
    Dim pFecha As Date
    
    If Trim(Right(Me.cboInstitucion.Text, 13)) = "" Then
        MsgBox "Debe especificar una Institución para Iniciar búsqueda", vbInformation
        Exit Sub
    End If
    
     If Me.cboInstitucion.ListIndex = 0 Or Me.cboInstitucion.ListIndex = -1 Then
        MsgBox "Seleccione una Institución, Verifique !!", vbInformation, "Aviso"
        Exit Sub
     End If
     
    If Not IsDate(txtperiodo) Then
        MsgBox "La fecha Ingresada No es Correcta, Verfique", vbCritical, "Aviso"
        Exit Sub
    End If
    If Me.cboInstitucion.ListIndex = -1 Then
        MsgBox "Debe especificar una Institucion, Verfique", vbCritical, "Aviso"
        Exit Sub
    End If
    Set oPrev = New previo.clsprevio
    Set oNCred = New COMNCredito.NCOMCredito

    sCadImp = oNCred.ImprimeReportePagoserviciosVarios(gsCodUser, gdFecSis, CDate(txtperiodo), IIf(Me.Option5.value, 1, 3), gsNomCmac, pInst)
    
    previo.Show sCadImp, "Registro de Archivo Pago de Servicios Registrados", False
    Set oPrev = Nothing
    Set oNCred = Nothing
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
  CentraForm Me
    Me.Top = 0
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Call CargaInstitucion
End Sub

Private Sub CargaInstitucion()
Dim rsCred As ADODB.Recordset
Dim oCredD As COMDCredito.DCOMCredito
    
    Set oCredD = New COMDCredito.DCOMCredito
    Set rsCred = New ADODB.Recordset
    Set rsCred = oCredD.GetInstitucionesPrevioPago(gsCodAge)
    
    Call llenar_cbo(rsCred, Me.cboInstitucion)
    Set oGen = Nothing
    Set rsCred = Nothing
    Exit Sub
ERRORCargaInstitucion:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub

Sub llenar_cbo(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
pcboObjeto.Clear
Do While Not pRs.EOF
    pcboObjeto.AddItem Trim(pRs!cPersNombre) & Space(100) & Trim(str(pRs!cPersCod))
    pRs.MoveNext
Loop
pRs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
For i = 1 To nContFlex - 1
        FEHojaEval.TextMatrix(i, 1) = ""
        FEHojaEval.TextMatrix(i, 2) = ""
        FEHojaEval.TextMatrix(i, 3) = ""
        FEHojaEval.TextMatrix(i, 4) = ""
        FEHojaEval.TextMatrix(i, 5) = ""
        FEHojaEval.TextMatrix(i, 6) = ""
        FEHojaEval.TextMatrix(i, 7) = ""
        FEHojaEval.TextMatrix(i, 8) = ""
    Next i
End Sub

Private Sub TxtPeriodo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Option5.SetFocus
End Sub

'Dim pInst As String
'
'Private Sub cmdBuscar_Click()
'Dim oPers As COMDPersona.DCOMPersonas
'Dim rs As ADODB.Recordset
'Dim ind As Integer
'
'    Set oPers = New COMDPersona.DCOMPersonas
'    If Trim(Right(Me.cboInstitucion.Text, 13)) = "" Then
'        MsgBox "Debe especificar una Institución para Iniciar búsqueda", vbInformation
'        Exit Sub
'    End If
'
'     If Me.cboInstitucion.ListIndex = 0 Or Me.cboInstitucion.ListIndex = -1 Then
'        MsgBox "Seleccione una Institución, Verifique !!", vbInformation, "Aviso"
'        Exit Sub
'     End If
'
'     If Me.txtperiodo.Text = "" Then
'        MsgBox "Fecha no válida, Verifique !!", vbInformation, "Aviso"
'        Exit Sub
'     End If
'
'     If Not IsDate(txtperiodo) Then
'        MsgBox "Fecha no válida, Verifique !!", vbInformation, "Aviso"
'        Exit Sub
'     End If
'
'    If Option5.value Then 'dia
'        ind = 1
'    Else
'        ind = 2
'    End If
'
'    Set rs = oPers.CargaDatosReportePagoServiciosxDoc(Trim(Right(Me.cboInstitucion.Text, 13)), Me.txtperiodo, ind)
'        FEHojaEval.Clear
'        FEHojaEval.Rows = 2
'        FEHojaEval.FormaCabecera
'        FEHojaEval.FormateaColumnas
'        FEHojaEval.TextMatrix(1, 0) = "1"
'
'    If Not (rs.BOF And rs.EOF) Then
'            Dim i As Integer
'            i = 0
'            Call LimpiaFlex(FEHojaEval)
''            ConfigurarMSH
'            Do Until rs.EOF
'            With Me.FEHojaEval
'                 i = i + 1
'                .Rows = .Rows + 1
'                .TextMatrix(.Rows - 2, 0) = i 'rsX!id
'                .TextMatrix(.Rows - 2, 1) = rs!nTipoDoc
'                .TextMatrix(.Rows - 2, 2) = rs!cNumDoc
'                .TextMatrix(.Rows - 2, 3) = rs!cNombre
'                .TextMatrix(.Rows - 2, 4) = rs!Pago
'                .TextMatrix(.Rows - 2, 5) = rs!dFechaPag
'                .TextMatrix(.Rows - 2, 6) = rs!cCodServicio
'                .TextMatrix(.Rows - 2, 7) = rs!cConcepto
'            End With
'        rs.MoveNext
'    Loop
'    rs.Close
'    Set rs = Nothing
'    Else
'        MsgBox "No existen Datos con la fecha Indicada, Verifique", vbInformation, "Atención"
'        Screen.MousePointer = 0
'        Exit Sub
'    End If
'    pInst = Trim(Right(Me.cboInstitucion.Text, 13))
'Set oPers = Nothing
'End Sub
'
'Sub ConfigurarMSH()
'    FEHojaEval.Cols = 8
'    FEHojaEval.Rows = 2
'
'    With FEHojaEval
'        .TextMatrix(0, 0) = "Id"
'        .TextMatrix(0, 1) = "TipDoc"
'        .TextMatrix(0, 2) = "NumDoc"
'        .TextMatrix(0, 3) = "Nombre Cliente"
'        .TextMatrix(0, 4) = "MontoPago"
'        .TextMatrix(0, 5) = "Fecha Pago"
'        .TextMatrix(0, 6) = "cCodServicio"
'        .TextMatrix(0, 7) = "Concepto"
'
'        .ColWidth(0) = 400
'        .ColWidth(1) = 800
'        .ColWidth(2) = 1000
'        .ColWidth(3) = 3000
'        .ColWidth(4) = 1200
'        .ColWidth(5) = 1200
'        .ColWidth(6) = 1800
'        .ColWidth(7) = 2500
'
'    End With
'End Sub
'
'Private Sub cmdImprimir_Click()
'Dim sCadImp As String
'    Dim oPrev As previo.clsprevio
'    Dim oNCred As COMNCredito.NCOMCredito
'    Dim pFecha As Date
'
'    If Trim(Right(Me.cboInstitucion.Text, 13)) = "" Then
'        MsgBox "Debe especificar una Institución para Iniciar búsqueda", vbInformation
'        Exit Sub
'    End If
'
'     If Me.cboInstitucion.ListIndex = 0 Or Me.cboInstitucion.ListIndex = -1 Then
'        MsgBox "Seleccione una Institución, Verifique !!", vbInformation, "Aviso"
'        Exit Sub
'     End If
'
'    If Not IsDate(txtperiodo) Then
'        MsgBox "La fecha Ingresada No es Correcta, Verfique", vbCritical, "Aviso"
'        Exit Sub
'    End If
'    If Me.cboInstitucion.ListIndex = -1 Then
'        MsgBox "Debe especificar una Institucion, Verfique", vbCritical, "Aviso"
'        Exit Sub
'    End If
'    Set oPrev = New previo.clsprevio
'    Set oNCred = New COMNCredito.NCOMCredito
'
'    sCadImp = oNCred.ImprimeReportePagoserviciosVarios(gsCodUser, gdFecSis, CDate(txtperiodo), 3, gsNomCmac, pInst)
'
'    previo.Show sCadImp, "Registro de Archivo Pago de Servicios Registrados", False
'    Set oPrev = Nothing
'    Set oNCred = Nothing
'End Sub
'
'Private Sub cmdsalir_Click()
'Unload Me
'End Sub
'
'Private Sub Form_Load()
'  CentraForm Me
'    Me.Top = 0
'    Me.Left = (Screen.Width - Me.Width) / 2
'    Me.Icon = LoadPicture(App.path & gsRutaIcono)
'    Call CargaInstitucion
'End Sub
'
'Private Sub CargaInstitucion()
'Dim rsCred As ADODB.Recordset
'Dim oCredD As COMDCredito.DCOMCredito
'
'    Set oCredD = New COMDCredito.DCOMCredito
'    Set rsCred = New ADODB.Recordset
'    Set rsCred = oCredD.GetInstitucionesPrevioPago(gsCodAge)
'
'    Call llenar_cbo(rsCred, Me.cboInstitucion)
'    Set oGen = Nothing
'    Set rsCred = Nothing
'    Exit Sub
'ERRORCargaInstitucion:
'    MsgBox err.Description, vbCritical, "Aviso"
'End Sub
'
'Sub llenar_cbo(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
'pcboObjeto.Clear
'Do While Not pRs.EOF
'    pcboObjeto.AddItem Trim(pRs!cPersNombre) & Space(100) & Trim(str(pRs!cPersCod))
'    pRs.MoveNext
'Loop
'pRs.Close
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'For i = 1 To nContFlex - 1
'        FEHojaEval.TextMatrix(i, 1) = ""
'        FEHojaEval.TextMatrix(i, 2) = ""
'        FEHojaEval.TextMatrix(i, 3) = ""
'        FEHojaEval.TextMatrix(i, 4) = ""
'        FEHojaEval.TextMatrix(i, 5) = ""
'        FEHojaEval.TextMatrix(i, 6) = ""
'        FEHojaEval.TextMatrix(i, 7) = ""
'        FEHojaEval.TextMatrix(i, 8) = ""
'    Next i
'End Sub
'
'Private Sub TxtPeriodo_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then Option5.SetFocus
'End Sub

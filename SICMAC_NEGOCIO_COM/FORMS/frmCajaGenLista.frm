VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajaGenLista 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5460
   ClientLeft      =   2100
   ClientTop       =   2550
   ClientWidth     =   10605
   Icon            =   "frmCajaGenLista.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   6
      Top             =   705
      Width           =   1335
   End
   Begin VB.Frame FraFechas 
      Height          =   600
      Left            =   6795
      TabIndex        =   15
      Top             =   45
      Width           =   3660
      Begin MSMask.MaskEdBox txtDesde 
         Height          =   315
         Left            =   795
         TabIndex        =   16
         Top             =   188
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   8388608
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txthasta 
         Height          =   315
         Left            =   2490
         TabIndex        =   17
         Top             =   195
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   8388608
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   135
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1935
         TabIndex        =   18
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.CheckBox chkTodo 
      Caption         =   "Se&leccionar"
      Height          =   195
      Left            =   255
      TabIndex        =   0
      Top             =   30
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      Height          =   4005
      Left            =   120
      TabIndex        =   7
      Top             =   1065
      Width           =   10350
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   8700
         TabIndex        =   11
         Top             =   3405
         Width           =   1470
      End
      Begin VB.TextBox txtMovDesc 
         Height          =   510
         Left            =   105
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3420
         Width           =   6270
      End
      Begin VB.CommandButton cmdConfHabCG 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   7230
         TabIndex        =   8
         Top             =   3405
         Width           =   1470
      End
      Begin SICMACT.FlexEdit fgLista 
         Height          =   2940
         Left            =   120
         TabIndex        =   12
         Top             =   210
         Width           =   10080
         _ExtentX        =   17780
         _ExtentY        =   5186
         Cols0           =   15
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   $"frmCajaGenLista.frx":030A
         EncabezadosAnchos=   "350-1200-3500-1200-1500-1500-1200-0-0-0-0-0-0-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-R-L-L-R-L-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-2-0-0-2-0-0-0-0-0-0-0-0"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbOrdenaCol     =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.CommandButton cmdExtornaHabCG 
         Caption         =   "&Extornar"
         Height          =   375
         Left            =   7230
         TabIndex        =   10
         Top             =   3405
         Width           =   1470
      End
      Begin VB.Label lblDescripcion 
         Caption         =   "Descripción :"
         Height          =   240
         Left            =   135
         TabIndex        =   20
         Top             =   3180
         Width           =   1050
      End
   End
   Begin VB.Frame fraConfHab 
      Height          =   1005
      Left            =   120
      TabIndex        =   1
      Top             =   45
      Width           =   6645
      Begin SICMACT.TxtBuscar TxtBuscarOrig 
         Height          =   345
         Left            =   840
         TabIndex        =   2
         Top             =   180
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         ForeColor       =   16512
      End
      Begin SICMACT.TxtBuscar TxtBuscarDest 
         Height          =   345
         Left            =   840
         TabIndex        =   3
         Top             =   555
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         ForeColor       =   16512
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Destino :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   75
         TabIndex        =   14
         Top             =   615
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "Origen :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Left            =   75
         TabIndex        =   13
         Top             =   225
         Width           =   660
      End
      Begin VB.Label lblDestDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   1905
         TabIndex        =   5
         Top             =   570
         Width           =   4635
      End
      Begin VB.Label lblOrigDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   1905
         TabIndex        =   4
         Top             =   180
         Width           =   4620
      End
   End
End
Attribute VB_Name = "frmCajaGenLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnOpeBovedaCaja As OpeBovedaCajaGeneral
Dim lnOpeBovAge As OpeBovedaAgencia
Dim lsCtaDebe As String
Dim lsCtaHaber As String
Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral
Dim lbSalir As Boolean

Private Sub GridFormatoMoneda()
Dim i As Long
Dim nMoneda As Moneda
For i = 1 To fgLista.Rows - 1
    nMoneda = CLng(fgLista.TextMatrix(i, 9))
    If nMoneda = gMonedaExtranjera Then
        fgLista.row = i
        fgLista.BackColorRow &HC0FFC0
    End If
Next i
End Sub

Private Sub chkTodo_Click()
Me.fraConfHab.Enabled = chkTodo.value
If chkTodo.value = 1 Then
    If fraConfHab.Visible Then
        If TxtBuscarDest.Enabled Then
            TxtBuscarDest.SetFocus
        ElseIf TxtBuscarOrig.Enabled Then
            TxtBuscarOrig.SetFocus
       End If
    End If
Else
    If TxtBuscarDest.Enabled Then
        TxtBuscarDest = ""
        lblDestDesc = ""
    End If
    If TxtBuscarOrig.Enabled Then
        TxtBuscarOrig = ""
        lblOrigDesc = ""
    End If
   
End If
End Sub

Private Sub cmdConfHabCG_Click()
Dim lsMovNro As String
Dim lsMovNroHab As String
Dim lnImporteHab As Double
Dim ocon As COMNContabilidad.NCOMContFunciones  'NContFunciones
Dim ldFechaHab  As Date
Dim rsBill As ADODB.Recordset
Dim rsMon As ADODB.Recordset
Dim lsmensaje  As String
Dim lsObjOrig As String
Dim lsObjDest As String
Dim lnMovNroHab As Long, nFila As Long
Dim nMoneda As COMDConstantes.Moneda
Set rsBill = New ADODB.Recordset
Set rsMon = New ADODB.Recordset

Dim lsCadImp As String
Dim nFicSal As Integer

If fgLista.TextMatrix(1, 0) = "" Then Exit Sub
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Ingrese Descripción de Movimiento", vbInformation, "Aviso"
    txtMovDesc.SetFocus
    Exit Sub
End If
nFila = fgLista.row
lsMovNroHab = fgLista.TextMatrix(nFila, 8)
lnMovNroHab = fgLista.TextMatrix(nFila, 12)
lnImporteHab = CDbl(fgLista.TextMatrix(nFila, 3))
ldFechaHab = CDate(fgLista.TextMatrix(nFila, 1))
nMoneda = CLng(fgLista.TextMatrix(nFila, 9))

If lnImporteHab = 0 Then
    MsgBox "Monto de Habilitación no Válida", vbInformation, "Aviso"
    Exit Sub
End If
frmCajaGenEfectivo.Muestra lsMovNroHab, lnImporteHab, nMoneda
If Not frmCajaGenEfectivo.lbOk Then
    Set frmCajaGenEfectivo = Nothing
    Exit Sub
End If
Set frmCajaGenEfectivo = Nothing
Set ocon = New COMNContabilidad.NCOMContFunciones  'NContFunciones
Select Case gsOpeCod
    Case gOpeBoveAgeConfHabCG
        lsObjOrig = IIf(Mid(TxtBuscarOrig, 1, 3) = Mid(TxtBuscarDest, 1, 3), TxtBuscarOrig, TxtBuscarDest)
        lsObjDest = TxtBuscarDest
        lsmensaje = "Desea Confirmar la Habilitación de Caja General Y/O Agencias??"
    Case gOpeBoveCGConfHabAgeBove
        lsmensaje = "Desea Confirmar la Habilitación realizada por la Agencia??"
        lsObjOrig = TxtBuscarOrig
        lsObjDest = TxtBuscarDest
End Select
If MsgBox(lsmensaje, vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = ocon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
    If oCaja.GrabaConfHabEfectivo(lsMovNro, gsOpeCod, txtMovDesc, lnImporteHab, _
            lsObjOrig, lsObjDest, lnMovNroHab) = 0 Then
        
        Select Case gsOpeCod
            Case gOpeBoveAgeConfHabCG
                Dim oContImp As COMNContabilidad.NCOMContImprimir 'NContImprimir
                Dim lbOk As Boolean
                
                Set oContImp = New COMNContabilidad.NCOMContImprimir
                lbOk = True
                lsCadImp = oContImp.ImprimeBoletahabilitacion("", "HABILITACION DE CAJA GENERAL", _
                            TxtBuscarOrig.Text, lblOrigDesc.Caption, TxtBuscarDest.Text, lblDestDesc.Caption, nMoneda, gsOpeCod, _
                            lnImporteHab, gsNomAge, lsMovNro, sLpt, gsCodCMAC)
                
                Do While lbOk
                    nFicSal = FreeFile
                    Open sLpt For Output As nFicSal
                        Print #nFicSal, lsCadImp & Chr$(12)
                        Print #nFicSal, ""
                    Close #nFicSal
                    
                    If MsgBox("Desea Reimprimir Boleta de Operacion??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                        lbOk = False
                    End If
                Loop
                Set oContImp = Nothing
        End Select
        
        Set ocon = Nothing
        fgLista.EliminaFila fgLista.row
        
        If Not (fgLista.TextMatrix(1, 1) = "" And fgLista.Rows = 2) Then
            If MsgBox("Desea Realizar otra confirmación de Habilitación??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                txtMovDesc = ""
            Else
                Unload Me
            End If
        Else
            Unload Me
        End If
    End If
End If
Set oCaja = Nothing
End Sub
Private Sub cmdExtornaHabCG_Click()
Dim lsMovNro As String
Dim lsMovNroHab As String
Dim lnImporteHab As Currency
Dim ocon As COMNContabilidad.NCOMContFunciones  'NContFunciones
Dim ldFechaHab  As Date
Dim lnMovNroHab As Long
'EJVG20131209 ***
Dim lnMovNroConf As Long
Dim lnMovNroDep As Long
'END EJVG *******
Set ocon = New COMNContabilidad.NCOMContFunciones

If fgLista.TextMatrix(1, 0) = "" Then Exit Sub
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Ingrese Descripción de Movimiento", vbInformation, "Aviso"
    txtMovDesc.SetFocus
    Exit Sub
End If
lsMovNroHab = fgLista.TextMatrix(fgLista.row, 8)
If gsOpeCod = gOpeBoveAgeExtConfHabCG Then
    lnMovNroHab = fgLista.TextMatrix(fgLista.row, 10)
Else
    lnMovNroHab = fgLista.TextMatrix(fgLista.row, 12)
    'EJVG20131209 ***
    lnMovNroConf = CLng(fgLista.TextMatrix(fgLista.row, 13))
    lnMovNroDep = CLng(fgLista.TextMatrix(fgLista.row, 14))
    'END EJVG *******
End If
    
lnImporteHab = CCur(fgLista.TextMatrix(fgLista.row, 3))
ldFechaHab = CDate(fgLista.TextMatrix(fgLista.row, 1))
If lnImporteHab = 0 Then
    MsgBox "Monto de Habilitación no Válida", vbInformation, "Aviso"
    Exit Sub
End If
If MsgBox("Desea Extornar la habilitación seleccionada?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    If gdFecSis <> ldFechaHab Then
        If MsgBox("Se va a Realizar el Extorno de Movimientos de dias anteriores" & vbCrLf & " Desea Proseguir??", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If
    lsMovNro = ocon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
    'oCaja.ExtornaHabEfectivo gdFecSis, ldFechaHab, lsMovNro, lnMovNroHab, gsOpeCod, _
                    Trim(txtMovDesc.Text), lnImporteHab
    oCaja.ExtornaHabEfectivo gdFecSis, ldFechaHab, lsMovNro, lnMovNroHab, gsOpeCod, _
                    Trim(txtMovDesc.Text), lnImporteHab, lnMovNroConf, lnMovNroDep 'EJVG20131209
                    
    If gdFecSis <> ldFechaHab Then
        ImpAsientoContable lsMovNro, , , , True, True, txtMovDesc.Text, , lnImporteHab
    End If
    Set ocon = Nothing
    If MsgBox("Desea Realizar otro Extorno de Habilitacion", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        txtMovDesc = ""
        fgLista.EliminaFila fgLista.row
    Else
        Unload Me
    End If
End If
Set oCaja = Nothing

End Sub

Private Sub cmdProcesar_Click()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
If fraConfHab.Visible Then
    If fraConfHab.Enabled And TxtBuscarOrig = "" Then
        MsgBox "Ingrese Area de Origen a Buscar", vbInformation, "Aviso"
        If TxtBuscarOrig.Enabled Then TxtBuscarOrig.SetFocus
        Exit Sub
    End If
    If fraConfHab.Enabled And TxtBuscarDest = "" Then
        MsgBox "Ingrese Area de Destino a Buscar", vbInformation, "Aviso"
        If TxtBuscarDest.Enabled Then TxtBuscarDest.SetFocus
        Exit Sub
    End If
End If
If FraFechas.Visible Then
    If ValFecha(Me.txtDesde) = False Then Exit Sub
    If ValFecha(txthasta) = False Then Exit Sub
    
    If CDate(txtDesde) > CDate(txthasta) Then
        MsgBox "Fecha Inicial no puede ser mayor que la Final", vbInformation, "Aviso"
        Exit Sub
    End If
End If
Me.MousePointer = 11
fgLista.Clear
fgLista.FormaCabecera
fgLista.Rows = 2
Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
Select Case gsOpeCod
    Case gOpeBoveAgeExtHabAgeACG, gOpeBoveAgeExtHabEntreAge
        Set rs = oCaja.GetHabCajaGen(TxtBuscarOrig, CDate(txtDesde), CDate(txthasta), Mid(TxtBuscarDest, 1, 3), Mid(TxtBuscarDest, 4, 2), False)
    Case gOpeBoveAgeExtConfHabCG
        If TxtBuscarOrig = TxtBuscarDest Then
            Me.MousePointer = 0
            MsgBox "Agencia de Origen no puede ser igual que la de Destino", vbInformation, "Aviso"
            If TxtBuscarOrig.Enabled Then TxtBuscarOrig.SetFocus
            Exit Sub
        End If
        Set rs = oCaja.GetDatosConfHabilitacion(TxtBuscarOrig, CDate(txtDesde), CDate(txthasta), Mid(TxtBuscarDest, 1, 3), Mid(TxtBuscarDest, 4, 2))
    Case gOpeBoveAgeConfHabCG
        If TxtBuscarOrig = TxtBuscarDest Then
            Me.MousePointer = 0
            MsgBox "Agencia de Origen no puede ser igual que la de Destino", vbInformation, "Aviso"
            If Me.TxtBuscarOrig.Enabled Then TxtBuscarOrig.SetFocus
            Exit Sub
        End If
        'If Me.chkRemChq.value = 1 Then
        '    Set rs = oCaja.GetRetirosEfectivoCaja(gOpeCGOpeBancosRetEfecMN, CDate(txtDesde), CDate(txthasta))
        'Else
            Set rs = oCaja.GetHabCajaGen(TxtBuscarOrig, CDate(txtDesde), CDate(txthasta), Mid(TxtBuscarDest, 1, 3), Mid(TxtBuscarDest, 4, 2))
        'End If
End Select

If Not rs.EOF And Not rs.BOF Then
    'If chkRemChq.value = 0 Then
        Set fgLista.Recordset = rs
        GridFormatoMoneda
    'Else
    '    fgLista.Clear
    '    fgLista.Rows = 2
    '    fgLista.Cols = 2
    '    fgLista.EncabezadosAnchos = "350-600-1000-1000-2000-3300-1200-0-0"
    '    fgLista.EncabezadosNombres = "N°-Tipo-Número-Voucher-Banco-Cuenta-Importe-nMovNro-Concepto"
    '    fgLista.EncabezadosAlineacion = "C-C-L-L-L-L-R-L-L"
    '    fgLista.FormatosEdit = "0-0-0-0-0-0-2-0-0"
    '    Set fgLista.Recordset = rs
    '    GridFormatoMoneda
    'End If
    fgLista.SetFocus
Else
    Me.MousePointer = 0
    MsgBox "Datos no Encontrados ", vbInformation, "Aviso"
End If
rs.Close
Set rs = Nothing
Me.MousePointer = 0
Set oCaja = Nothing
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub fgLista_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMovDesc.SetFocus
End If
End Sub

Private Sub Form_Activate()
If lbSalir Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Dim oOpe As COMNCajaGeneral.NCOMCajaGeneral  'DOperacion
Dim rs As ADODB.Recordset
Dim sAreCodCaja As String, sAreaNomCaja As String

Set oOpe = New COMNCajaGeneral.NCOMCajaGeneral
Me.Icon = LoadPicture(App.path & gsRutaIcono)
lbSalir = False
txtDesde = gdFecSis
txthasta = gdFecSis
Me.Caption = gsOpeDesc
cmdExtornaHabCG.Visible = False
cmdConfHabCG.Visible = False

fraConfHab.Visible = False

Select Case gsOpeCod
    Case gOpeBoveAgeConfHabCG, gOpeBoveAgeExtConfHabCG
        FraFechas.Enabled = True
        Me.txtDesde.Enabled = True
        Me.txthasta.Enabled = True
        fraConfHab.Visible = True
        If gsOpeCod = gOpeBoveAgeExtConfHabCG Then
            cmdExtornaHabCG.Visible = True
        Else
            cmdConfHabCG.Visible = True
        End If
        'Set rs = New ADODB.Recordset
        Set rs = oOpe.GetOpeObj(gsOpeCod, "1")
        'MODIFICADO PW NO MUESTRA LA AGENCIA
        sAreCodCaja = rs("Codigo")
        sAreaNomCaja = rs("Descripcion")
        Set rs = oOpe.GetOpeObj(gsOpeCod, "0")
        TxtBuscarDest.Text = rs("Codigo") & gsCodAge
        TxtBuscarDest.Enabled = False
        lblDestDesc.Caption = gsNomAge
        rs.Close
        Set rs = Nothing
        chkTodo.Visible = False
        TxtBuscarOrig.Text = sAreCodCaja
        TxtBuscarOrig.Enabled = False
        lblOrigDesc.Caption = sAreaNomCaja
        If gOpeBoveAgeConfHabCG Then
            'chkRemChq.Visible = True
        End If
    Case gOpeBoveAgeExtHabAgeACG
        fraConfHab.Visible = True
        chkTodo.Visible = False
        Set rs = oOpe.GetOpeObj(gsOpeCod, "0")
        sAreCodCaja = rs("Codigo")
        sAreaNomCaja = rs("Descripcion")
        Set rs = oOpe.GetOpeObj(gsOpeCod, "1")
        TxtBuscarOrig.Text = rs("Codigo") & gsCodAge
        TxtBuscarOrig.Enabled = False
        lblOrigDesc.Caption = gsNomAge
        rs.Close
        Set rs = Nothing
        TxtBuscarDest.Text = sAreCodCaja
        TxtBuscarDest.Enabled = False
        lblDestDesc.Caption = sAreaNomCaja
        cmdExtornaHabCG.Visible = True
End Select

Set oOpe = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set oCaja = Nothing
End Sub
Private Sub txtBuscarDest_EmiteDatos()
lblDestDesc = TxtBuscarDest.psDescripcion
If cmdProcesar.Visible Then cmdProcesar.SetFocus
End Sub

Private Sub TxtBuscarOrig_EmiteDatos()
lblOrigDesc = TxtBuscarOrig.psDescripcion
If Me.TxtBuscarDest.Enabled Then
    If TxtBuscarDest.Visible Then TxtBuscarDest.SetFocus
ElseIf Me.cmdProcesar.Enabled Then
    If Me.cmdProcesar.Visible Then Me.cmdProcesar.SetFocus
    End If
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtDesde) = False Then Exit Sub
    If CDate(txtDesde) > CDate(txthasta) Then
        MsgBox "Fecha Inicial no puede ser mayor que la Final", vbInformation, "Aviso"
        Exit Sub
    End If
    txthasta.SetFocus
End If
End Sub


Private Sub txthasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txthasta) = False Then Exit Sub
    If CDate(txtDesde) > CDate(txthasta) Then
        MsgBox "Fecha Inicial no puede ser mayor que la Final", vbInformation, "Aviso"
        Exit Sub
    End If
    cmdProcesar.SetFocus
End If
End Sub
Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    KeyAscii = 0
    If cmdExtornaHabCG.Visible Then cmdExtornaHabCG.SetFocus
    If Me.cmdConfHabCG.Visible Then cmdConfHabCG.SetFocus
End If
End Sub

Public Sub ImpAsientoContable(ByVal psMovNro As String, Optional ByVal psDocVoucher As String = "", _
                                  Optional ByVal psDocTpo As String = "", Optional ByVal psDocumento As String = "", _
                                  Optional ByVal pbEfectivo As Boolean = False, _
                                  Optional ByVal pbIngreso As Boolean = False, _
                                  Optional ByVal psGlosa As String, Optional ByVal psPersCod As String, _
                                  Optional ByVal pnImporte As Currency, Optional ByVal pnTipoArendir As ArendirTipo, _
                                  Optional ByVal psNroRecViaticos As String = "", Optional pbHabEfectivo As Boolean = False, _
                                  Optional ByVal pnNumCopiasAsiento As Integer = 2)

Dim oContImp As COMNContabilidad.NCOMContImprimir 'NContImprimir
Dim oNContFunc As COMNContabilidad.NCOMContFunciones  'NContFunciones
Dim oPlant As COMDContabilidad.DCOMPlantilla

Set oContImp = New COMNContabilidad.NCOMContImprimir
Set oNContFunc = New COMNContabilidad.NCOMContFunciones
Set oPlant = New COMDContabilidad.DCOMPlantilla

Dim lsAsiento  As String
Dim lsTitulo As String
Dim lsVEOPSave As String
Dim lsRecibo As String
Dim lsOPSave As String
Dim lsHab As String
Dim lsPie As String
Dim lsOtraFirma As String
Dim i As Integer
Dim lsCopias As String
Dim lsCartas As String

Dim nFicSal As Integer
Dim oPrevio As previo.clsprevio

lsTitulo = ""
If psDocVoucher <> "" Then
    lsTitulo = " COMPROBANTE DE EGRESO N° " & psDocVoucher
End If
If pbEfectivo Then
    lsRecibo = oContImp.ImprimeReciboIngresoEgreso(psMovNro, gdFecSis, psGlosa, _
                                                   gsNomCmac, gsOpeCod, psPersCod, _
                                                   pnImporte, gnColPage, pnTipoArendir, psNroRecViaticos, pbIngreso)
                                                   
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
        Print #nFicSal, lsRecibo & Chr$(12)
        Print #nFicSal, ""
    Close #nFicSal
    
    If pbIngreso Then
        lsTitulo = "I N G R E S O   D E   E F E C T I V O"
    Else
        lsTitulo = "S A L I D A   D E   E F E C T I V O"
   End If
End If
lsPie = "179"

If pbHabEfectivo Then
    lsTitulo = "H A B I L I T A C I O N   D E   E F E C T I V O "
    lsHab = oContImp.ImprimeDocSalidaEfectivo(gnColPage, gdFecSis, gsOpeCod, psMovNro, gsNomCmac)
    
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
        Print #nFicSal, lsHab & Chr$(12)
        Print #nFicSal, ""
    Close #nFicSal
    
    lsPie = "158"
    lsOtraFirma = "RESPONSABLE TRASLADO"
End If
'lsAsiento = oContImp.ImprimeAsientoContable(psMovNro, gnLinPage, gnColPage, lsTitulo, , lsPie, lsOtraFirma)
lsAsiento = oContImp.ImprimeAsientoContable(psMovNro, gnLinPage, gnColPage, lsTitulo, , lsPie)
Select Case val(psDocTpo)
    Case TpoDocCheque  '  gnDocTpoCheque
        If psDocumento <> "" Then
            lsAsiento = psDocumento & lsAsiento
        End If
        For i = 1 To pnNumCopiasAsiento - 1
            lsCopias = lsCopias & Chr$(12) & lsAsiento
        Next
        lsAsiento = psDocumento & Chr$(12) + lsAsiento & lsCopias
    Case TpoDocCarta  ' gnDocTpoCarta
        If psDocumento <> "" Then
            frmCopiasImp.Show 1
            For i = 1 To frmCopiasImp.CopiasCartas - 1
                lsCartas = Chr$(12) + psDocumento
            Next i
            lsCartas = psDocumento + lsCartas
            pnNumCopiasAsiento = frmCopiasImp.CopiasAsientos
        End If
        For i = 1 To pnNumCopiasAsiento - 1
            lsCopias = lsCopias & Chr$(12) & lsAsiento
        Next
        lsAsiento = lsAsiento & lsCopias
        lsAsiento = IIf(lsCartas = "", "", lsCartas & Chr$(12)) + lsAsiento
        Set frmCopiasImp = Nothing
    Case TpoDocOrdenPago, TpoDocNotaCargo, TpoDocNotaAbono        'gnDocTpoOPago, TpoDocNotaCargo, TpoDocNotaAbono
        If psDocumento <> "" Then
            lsAsiento = psDocumento & lsAsiento
        End If
        For i = 1 To pnNumCopiasAsiento - 1
            lsCopias = lsCopias & Chr$(12) & lsAsiento
        Next
        lsAsiento = lsAsiento & lsCopias
    Case Else
        If pbHabEfectivo Then
            For i = 1 To pnNumCopiasAsiento - 1
                lsCopias = lsCopias & Chr$(12) & lsAsiento
            Next
            lsAsiento = lsAsiento & lsCopias
            If lsHab <> "" Then
                lsAsiento = lsAsiento & Chr$(12) & lsHab
            End If
        Else
            For i = 1 To pnNumCopiasAsiento - 1
                lsCopias = lsCopias & Chr$(12) & lsAsiento
            Next
            lsAsiento = lsAsiento & lsCopias
        End If
        If lsRecibo <> "" Then
            lsAsiento = lsAsiento & Chr$(12) & lsRecibo
        End If
End Select

If psDocTpo <> "" Then
    If psDocTpo = TpoDocOrdenPago And pbIngreso = False Then
        Set oPlant = New COMDContabilidad.DCOMPlantilla
         lsOPSave = oPlant.GetPlantillaDoc(IDPlantillaOP)
        Set oPlant = Nothing
        lsOPSave = lsOPSave & IIf(lsOPSave = "", "", Chr$(12)) & psDocumento
         
        Set oPlant = New COMDContabilidad.DCOMPlantilla
            oPlant.GrabaPlantilla "OPBatch", "Ordenes de Pago para impresiones en Batch", lsOPSave
        Set oPlant = Nothing
        
        Set oPlant = New COMDContabilidad.DCOMPlantilla
            lsVEOPSave = oPlant.GetPlantillaDoc(IDPlantillaVOP)
        Set oPlant = Nothing
        lsVEOPSave = lsVEOPSave & IIf(lsVEOPSave = "", "", Chr$(12)) & lsAsiento
        
        Set oPlant = New COMDContabilidad.DCOMPlantilla
            oPlant.GrabaPlantilla "OPVEBatch", "Voucher de egresos de Ordenes de Pago para impresiones en Batch", lsVEOPSave
        Set oPlant = Nothing
         
        If MsgBox(" ¿ Desea Imprimir Orden de Pago ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
            Set oPlant = New COMDContabilidad.DCOMPlantilla
                lsOPSave = oPlant.GetPlantillaDoc(IDPlantillaOP)
            Set oPlant = Nothing
    
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, lsOPSave & Chr$(12)
                Print #nFicSal, ""
            Close #nFicSal
            
            lsVEOPSave = oPlant.GetPlantillaDoc(IDPlantillaVOP)
            Set oPrevio = New previo.clsprevio
                oPrevio.Show lsVEOPSave, gsOpeDesc, False, gnLinPage, gImpresora
            Set oPrevio = Nothing
            
            Set oPlant = New COMDContabilidad.DCOMPlantilla
                oPlant.GrabaPlantilla IDPlantillaOP, "Ordenes de Pago para impresiones en Batch", ""
                oPlant.GrabaPlantilla IDPlantillaVOP, "Voucher de egresos de Ordenes de Pago para impresiones en Batch", ""
            Set oPlant = Nothing
            
       End If
    Else
     Set oPrevio = New previo.clsprevio
       oPrevio.Show lsAsiento, gsOpeDesc, False, gnLinPage, gImpresora
     Set oPrevio = Nothing
    End If
Else
   Set oPrevio = New previo.clsprevio
    oPrevio.Show lsAsiento, gsOpeDesc, False, gnLinPage, gImpresora
   Set oPrevio = Nothing
End If
Set oContImp = Nothing
Set oNContFunc = Nothing
End Sub


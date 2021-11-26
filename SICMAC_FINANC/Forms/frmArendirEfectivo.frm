VERSION 5.00
Begin VB.Form frmArendirEfectivo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A Rendir en Efectivo"
   ClientHeight    =   5790
   ClientLeft      =   2895
   ClientTop       =   1980
   ClientWidth     =   6255
   Icon            =   "frmArendirEfectivo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4785
      TabIndex        =   7
      Top             =   5340
      Width           =   1380
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3375
      TabIndex        =   6
      Top             =   5340
      Width           =   1380
   End
   Begin VB.Frame fraBilletajes 
      Caption         =   "Billetajes  :"
      Height          =   3915
      Left            =   75
      TabIndex        =   4
      Top             =   1365
      Width           =   6090
      Begin Sicmact.FlexEdit fgBilletajes 
         Height          =   3180
         Left            =   135
         TabIndex        =   5
         Top             =   225
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   5609
         Cols0           =   5
         FixedCols       =   2
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-Descripción-Monto-cEfectivoCod-nEfectivoValor"
         EncabezadosAnchos=   "350-3500-1200-0-0"
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
         ColumnasAEditar =   "X-X-2-X-X"
         ListaControles  =   "0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-R-C-C"
         FormatosEdit    =   "0-0-2-0-0"
         AvanceCeldas    =   1
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         CellBackColor   =   -2147483633
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   315
         Left            =   3690
         TabIndex        =   9
         Top             =   3450
         Width           =   1965
      End
      Begin VB.Label lbl3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL"
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
         Height          =   210
         Left            =   2580
         TabIndex        =   8
         Top             =   3495
         Width           =   615
      End
      Begin VB.Shape ShapeS 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Left            =   2220
         Top             =   3435
         Width           =   3465
      End
   End
   Begin VB.Frame fraDatosPrinc 
      Caption         =   "Datos Principales"
      Height          =   1275
      Left            =   75
      TabIndex        =   1
      Top             =   90
      Width           =   6090
      Begin VB.TextBox lblArendirMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   4125
         TabIndex        =   0
         Top             =   915
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   3345
         TabIndex        =   14
         Top             =   945
         Width           =   615
      End
      Begin VB.Label lblNroARendir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1005
         TabIndex        =   13
         Top             =   255
         Width           =   1200
      End
      Begin VB.Label lblPersCod 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1005
         TabIndex        =   12
         Top             =   585
         Width           =   1200
      End
      Begin VB.Label lblCaptionArendirNro 
         AutoSize        =   -1  'True
         Caption         =   "A Rendir N° :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   11
         Top             =   255
         Width           =   930
      End
      Begin VB.Label lblArendirArea 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2220
         TabIndex        =   10
         Top             =   255
         Width           =   3720
      End
      Begin VB.Label lblPersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2220
         TabIndex        =   3
         Top             =   585
         Width           =   3720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Persona :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   105
         TabIndex        =   2
         Top             =   600
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmArendirEfectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vbOk As Boolean
Dim rsAux As ADODB.Recordset
Dim lnMoneda As Moneda
Dim lnMonto As Currency
Dim lnArendirFase As ARendirFases
Dim lnDiferencia As Currency
Dim lsCaptionArendir As String
Dim lnParamDif As Currency
Dim lbModificaMonto As Boolean

Private Sub cmdAceptar_Click()
vbOk = True
lnDiferencia = 0
If CCur(lblArendirMonto) = 0 Then
    MsgBox "Monto a Pagar debe ser mayor a cero", vbInformation, "Aviso!"
    lblArendirMonto = Format(lnMonto, gsFormatoNumeroView)
    lblArendirMonto.SetFocus
    Exit Sub
End If
If Abs(lnMonto) < CCur(lblArendirMonto) Then
    MsgBox "Monto a Pagar no puede ser mayor a " & Format(lnMonto, gsFormatoNumeroView), vbInformation, "Aviso!"
    lblArendirMonto.SetFocus
    Exit Sub
End If
lnMonto = CCur(lblArendirMonto)

'If lnArendirFase <> ArendirRendicion Then
'    If CCur(lblTotal) <> CCur(lblArendirMonto) Then
'        MsgBox "Monto Ingresado no cubre el Monto de Arendir", vbInformation, "Aviso"
'        fgBilletajes.SetFocus
'        Exit Sub
'    End If
'Else
    If CCur(lblTotal) <> CCur(lblArendirMonto) Then
        If Abs(CCur(lblTotal) - CCur(lblArendirMonto)) <= lnParamDif Then
            If MsgBox("Monto Ingresado no cubre el Monto de Arendir" & Chr(13) & "Existe una diferencia la cual no ha sido cubierta" & Chr(13) & "Desea continuar???", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                fgBilletajes.SetFocus
                Exit Sub
            Else
                lnDiferencia = CCur(lblArendirMonto) - CCur(Abs(lblTotal))
            End If
        Else
            MsgBox "Monto Ingresado no cubre el Monto de Arendir", vbInformation, "Aviso"
            fgBilletajes.SetFocus
            Exit Sub
        End If
    End If
'End If
Set rsAux = fgBilletajes.GetRsNew

'Me.Hide
Unload Me
DoEvents
End Sub
Private Sub cmdCancelar_Click()
vbOk = False
Unload Me
End Sub
Public Property Get lbOk() As Variant
    lbOk = vbOk
End Property
Public Property Let lbOk(ByVal vNewValue As Variant)
    vbOk = vNewValue
End Property
Public Property Get rsEfectivo() As ADODB.Recordset
    Set rsEfectivo = rsAux
End Property
Public Property Let rsEfectivo(ByVal vNewValue As ADODB.Recordset)
    Set rsAux = vNewValue
End Property
Private Sub fgBilletajes_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim lnTotal As Currency
Dim lnValor As Currency
lnTotal = fgBilletajes.SumaRow(2)
lnValor = CCur(IIf(fgBilletajes.TextMatrix(pnRow, pnCol) = "", "0", fgBilletajes.TextMatrix(pnRow, pnCol)))
If Residuo(lnValor, fgBilletajes.TextMatrix(pnRow, 4)) Then
    If lnTotal > Abs(lnMonto) Then
        If MsgBox("Total no sebe Superar a lo solicitado. ¿ Desea Continuar ? ", vbQuestion + vbYesNo, "Confirmación") = vbNo Then
            Cancel = False
        Else
            lblTotal = Format(lnTotal, "#,####0.00")
        End If
    Else
        lblTotal = Format(lnTotal, "#,####0.00")
        Cancel = True
    End If
Else
    Cancel = False
End If
End Sub

Private Sub Form_Load()
Dim oGen As DGeneral
Set oGen = New DGeneral
lnParamDif = oGen.GetParametro(4000, 1001)
Set oGen = Nothing
CentraForm Me
Me.Caption = gsOpeDesc
CargaBilletajes
If lsCaptionArendir <> "" Then lblCaptionArendirNro = lsCaptionArendir
lblTotal = Format(fgBilletajes.SumaRow(2), "#,####0.00")
lblArendirMonto.Enabled = lbModificaMonto

End Sub
Public Sub Inicio(ByVal pTipoARendir As ArendirTipo, ByVal psNroArendir As String, _
            pnMoneda As Moneda, psAreaDesc As String, ByVal pnMontoARendir As Currency, _
            ByVal psPersCod As String, ByVal psPersNombre As String, Optional pnArendirFase As ARendirFases = ArendirSustentacion, _
            Optional psCaptionArendir As String, Optional pbModificaMonto As Boolean = False)

lbModificaMonto = pbModificaMonto
lsCaptionArendir = psCaptionArendir
lblPersCod = psPersCod
lblPersNombre = psPersNombre
lblArendirArea = psAreaDesc
lblNroARendir = psNroArendir
lnMoneda = pnMoneda
lnMonto = pnMontoARendir
lnArendirFase = pnArendirFase
lblArendirMonto = Format(Abs(pnMontoARendir), "##,###0.00")

If Val(pnMontoARendir) < 0 Then
    lblArendirMonto.ForeColor = &HFF&
Else
    lblArendirMonto.ForeColor = &HC00000
End If

CargaBilletajes
Me.Show 1
End Sub
Private Sub CargaBilletajes()
Dim Sql As String
Dim rs As ADODB.Recordset
Dim oContFunct As NContFunciones
Dim oEfec As Defectivo
Dim lnFila As Long

Set oContFunct = New NContFunciones
Set oEfec = New Defectivo

Set rs = New ADODB.Recordset
Set rs = oEfec.EmiteBilletajes(lnMoneda)
fgBilletajes.FontFixed.Bold = True
fgBilletajes.Clear
fgBilletajes.FormaCabecera
fgBilletajes.Rows = 2
Do While Not rs.EOF
    fgBilletajes.AdicionaFila
    lnFila = fgBilletajes.Row
    fgBilletajes.TextMatrix(lnFila, 1) = rs!Descripcion
    fgBilletajes.TextMatrix(lnFila, 2) = Format(rs!Monto, "0.00")
    fgBilletajes.TextMatrix(lnFila, 3) = rs!cEfectivoCod
    fgBilletajes.TextMatrix(lnFila, 4) = rs!nEfectivoValor
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
Set oContFunct = Nothing
Set oEfec = Nothing
'fgBilletajes.Row = 1
fgBilletajes.Col = 2
If fgBilletajes.Visible And fgBilletajes.Enabled Then
    fgBilletajes.SetFocus
End If
End Sub
Public Function Residuo(Dividendo As Currency, Divisor As Currency) As Boolean
Dim X As Currency
X = Round(Dividendo / Divisor, 0)
Residuo = True
X = X * Divisor
If X <> Dividendo Then
   Residuo = False
End If
End Function
Public Property Get vnDiferencia() As Currency
vnDiferencia = lnDiferencia
End Property
Public Property Let vnDiferencia(ByVal vNewValue As Currency)
lnDiferencia = vNewValue
End Property

Public Property Get vnMonto() As Currency
vnMonto = lnMonto
End Property
Public Property Let vnMonto(ByVal vNewValue As Currency)
lnMonto = vNewValue
End Property

Private Sub lblArendirMonto_GotFocus()
    fEnfoque lblArendirMonto
End Sub

Private Sub lblArendirMonto_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(lblArendirMonto, KeyAscii, 15, 2)
    If KeyAscii = 13 Then
        If lblArendirMonto = "" Then lblArendirMonto = "0.00"
        If lnMonto < CCur(lblArendirMonto) Then
            MsgBox "Monto a Pagar no puede ser mayor a " & Format(lnMonto, gsFormatoNumeroView), vbInformation, "Aviso!"
            lblArendirMonto.SetFocus
            Exit Sub
        End If
        fgBilletajes.SetFocus
    End If
End Sub

Private Sub lblArendirMonto_LostFocus()
lblArendirMonto = Format(lblArendirMonto, gsFormatoNumeroView)
End Sub

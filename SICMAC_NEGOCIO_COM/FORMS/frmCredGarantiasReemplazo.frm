VERSION 5.00
Begin VB.Form frmCredGarantiaReemplazo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reemplazo de Garantia"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14640
   Icon            =   "frmCredGarantiasReemplazo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   14640
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPorTipoCambioME 
      Enabled         =   0   'False
      Height          =   375
      Left            =   12840
      TabIndex        =   13
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtPorTipoCambioMN 
      Enabled         =   0   'False
      Height          =   375
      Left            =   9360
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtTipoCambio 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   12960
      TabIndex        =   6
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   11400
      TabIndex        =   5
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdReemplazar 
      Caption         =   "Reemplazar"
      Height          =   495
      Left            =   9840
      TabIndex        =   4
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Garantia Reemplazante"
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   14415
      Begin SICMACT.FlexEdit FLante 
         Height          =   1455
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   2566
         Cols0           =   9
         HighLight       =   2
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-Sel-Garantia-Código-Clasificación-Tipo Garantia-Moneda-Disponible-Gravado"
         EncabezadosAnchos=   "0-400-4000-1200-2000-2000-1200-1200-1200"
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
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-X-X-X-X-X-8"
         TextStyleFixed  =   4
         ListaControles  =   "0-4-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-C-C-L-R-R"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-2"
         AvanceCeldas    =   1
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Garantía Reemplazada"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   14415
      Begin SICMACT.FlexEdit FLGada 
         Height          =   1650
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   13980
         _ExtentX        =   24659
         _ExtentY        =   2910
         Cols0           =   9
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "-Sel-Garantia-Codigo-Clasificación-Tipo Garantia-Moneda-Gravado-Saldo Gravado"
         EncabezadosAnchos=   "0-400-4000-1200-2000-2000-1200-1200-1200"
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
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-4-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-L-R-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
         AvanceCeldas    =   1
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
      End
   End
   Begin SICMACT.ActXCodCta ActXCta 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      Texto           =   "Crédito"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "Por Tipo Cambio ME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10920
      TabIndex        =   12
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Por Tipo Cambio MN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   10
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de Cambio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmCredGarantiaReemplazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnContadorGada As Integer
Dim lnContadorAnte As Integer
Dim nSaldoTotalMN As Currency
Dim nSaldoTotalME As Currency
Dim nSaldoTotalNuevoMN As Currency
Dim nSaldoTotalNuevoME As Currency

Dim lnContadorActGada As Integer
Dim lnContadorActAnte As Integer

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    Call CargarDatos
 End If
End Sub

Private Sub CargarDatos()
Dim oGarantia As COMNCredito.NCOMGarantia
Dim OrsGRante As ADODB.Recordset
Dim OrsGRada As ADODB.Recordset
Dim row As Long
Dim i As Integer
'Set OrsGRante = New ADODB.Recordset
'Set OrsGRada = New ADODB.Recordset
lnContadorGada = 0
lnContadorAnte = 0
lnContadorActGada = 0
lnContadorActAnte = 0
Set oGarantia = New COMNCredito.NCOMGarantia
Call oGarantia.CargarDatosCambioGarantia(ActXCta.NroCuenta, gdFecSis, OrsGRada, OrsGRante)

FormateaFlex FLGada
FormateaFlex FLante

If Not (OrsGRada.BOF Or OrsGRada.EOF) Then
    Do While Not OrsGRada.EOF
            FLGada.AdicionaFila
            FLGada.TextMatrix(OrsGRada.Bookmark, 1) = "1"
            FLGada.TextMatrix(OrsGRada.Bookmark, 2) = OrsGRada!cGarantia
            FLGada.TextMatrix(OrsGRada.Bookmark, 3) = OrsGRada!cNumGarant
            FLGada.TextMatrix(OrsGRada.Bookmark, 4) = OrsGRada!Clasificacion
            FLGada.TextMatrix(OrsGRada.Bookmark, 5) = OrsGRada!cConsDescripcion
            FLGada.TextMatrix(OrsGRada.Bookmark, 6) = OrsGRada!nmoneda
            FLGada.TextMatrix(OrsGRada.Bookmark, 7) = OrsGRada!nGravado
            FLGada.TextMatrix(OrsGRada.Bookmark, 8) = OrsGRada!nGravament
            OrsGRada.MoveNext
            lnContadorGada = lnContadorGada + 1
    Loop
End If
FLGada.TopRow = 1
If Not (OrsGRante.BOF Or OrsGRante.EOF) Then
    Do While Not OrsGRante.EOF
            FLante.AdicionaFila
'            row = FLante.row
            FLante.TextMatrix(OrsGRante.Bookmark, 1) = "1"
            FLante.TextMatrix(OrsGRante.Bookmark, 2) = OrsGRante!cGarantia
            FLante.TextMatrix(OrsGRante.Bookmark, 3) = OrsGRante!cNumGarant
            FLante.TextMatrix(OrsGRante.Bookmark, 4) = OrsGRante!Clasificacion
            FLante.TextMatrix(OrsGRante.Bookmark, 5) = OrsGRante!cConsDescripcion
            FLante.TextMatrix(OrsGRante.Bookmark, 6) = OrsGRante!nmoneda
            FLante.TextMatrix(OrsGRante.Bookmark, 7) = OrsGRante!nPorGravar
            OrsGRante.MoveNext
            lnContadorAnte = lnContadorAnte + 1
    Loop
End If
FLante.TopRow = 1
    For i = 1 To lnContadorGada
        FLGada.TextMatrix(i, 1) = ""
    Next i
    For i = 1 To lnContadorAnte
        FLante.TextMatrix(i, 1) = ""
    Next i
    cmdReemplazar.Enabled = True
    ActXCta.Enabled = False
End Sub

Private Sub cmdCancelar_Click()
    Call InicializaControles
End Sub
Private Sub InicializaControles()
    FormateaFlex FLGada
    FormateaFlex FLante
    ActXCta.Cuenta = ""
    ActXCta.CMAC = gsCodCMAC
    ActXCta.Age = gsCodAge
    ActXCta.Prod = ""
    cmdReemplazar.Enabled = False
    ActXCta.Enabled = True
    txtPorTipoCambioMN.Text = ""
    txtPorTipoCambioME.Text = ""
End Sub

Private Sub cmdReemplazar_Click()

Dim oGarantia As COMDCredito.DCOMGarantia
Set oGarantia = New COMDCredito.DCOMGarantia
Dim blogicoGada As Boolean
Dim blogicoAnte As Boolean
Dim lsNumGaranAntigua As String
Dim lsNumGaranNueva As String
Dim lsMonGaranAntigua As String

Dim nSumaGada As Currency
Dim nSumaAnte As Currency


Dim lsMatrixGada() As String
Dim lsMatrixAnte() As String

Dim i As Integer

blogicoGada = False
blogicoAnte = False

If lnContadorGada = 0 Then
    MsgBox "El Crédito no tiene garantias asociadas", vbInformation
    Exit Sub
End If
If lnContadorAnte = 0 Then
    MsgBox "El Cliente no tiene nuevas garantias para asociar", vbInformation
    Exit Sub
End If

If lnContadorActGada > 1 Then
    MsgBox "no se puede seleccionar mas de una garantia", vbInformation
    Exit Sub
End If
If lnContadorActAnte > 1 Then
    MsgBox "no se puede seleccionar mas de una garantia", vbInformation
    Exit Sub
End If

nSumaGada = 0
For i = 1 To lnContadorGada
    If FLGada.TextMatrix(i, 1) = "." Then
        lsNumGaranAntigua = FLGada.TextMatrix(i, 3)
        lsMonGaranAntigua = IIf(FLGada.TextMatrix(i, 6) = "Soles", 1, 2)
        If Mid(ActXCta.NroCuenta, 9, 1) = "1" Then
            nSumaGada = nSumaGada + CCur(FLGada.TextMatrix(i, 7)) * IIf(FLGada.TextMatrix(i, 6) = "Soles", 1, CCur(txtTipoCambio.Text))
        Else
            nSumaGada = nSumaGada + CCur(FLGada.TextMatrix(i, 7)) / IIf(FLGada.TextMatrix(i, 6) = "Soles", CCur(txtTipoCambio.Text), 1)
        End If
        blogicoGada = True
    End If
Next i

If blogicoGada = False Then
    MsgBox "El Seleccionó ninguna garantia para Reemplazar", vbInformation
    Exit Sub
End If
For i = 1 To lnContadorAnte
    If FLante.TextMatrix(i, 1) = "." Then

        lsNumGaranNueva = FLante.TextMatrix(i, 3)
        If FLante.TextMatrix(i, 8) = "" Then
            MsgBox "Ingresar monto gravado de la garantia", vbInformation
            Exit Sub
        End If
        If Mid(ActXCta.NroCuenta, 9, 1) = "1" Then
            nSumaAnte = nSumaAnte + CCur(FLante.TextMatrix(i, 8)) * IIf(FLante.TextMatrix(i, 6) = "Soles", 1, CCur(txtTipoCambio.Text))
        Else
            nSumaAnte = nSumaAnte + CCur(FLante.TextMatrix(i, 8)) / IIf(FLante.TextMatrix(i, 6) = "Soles", CCur(txtTipoCambio.Text), 1)
        End If
        blogicoAnte = True
    End If
Next i

If blogicoAnte = False Then
    MsgBox "El Seleccionó ninguna garantia para usar como Reemplazante", vbInformation
    Exit Sub
End If

If lsMonGaranAntigua = "1" And (nSaldoTotalMN <> nSaldoTotalNuevoMN) Then
    MsgBox "El monto no cubre la Garantia", vbInformation
    Exit Sub
End If

If lsMonGaranAntigua = "2" And (nSaldoTotalME <> nSaldoTotalNuevoME) Then
    MsgBox "El monto no cubre la Garantia", vbInformation
    Exit Sub
End If

Call oGarantia.GuardarReemplazoDeGarantia(ActXCta.NroCuenta, lsNumGaranAntigua, lsNumGaranNueva, CCur(txtPorTipoCambioMN.Text), CCur(txtPorTipoCambioME.Text), lsMonGaranAntigua, CCur(txtTipoCambio.Text))
MsgBox "Proceso se realizó correctamente", vbInformation
Call CargarDatos
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub FLante_OnCellChange(pnRow As Long, pnCol As Long)
 Dim i As Integer
 nSaldoTotalNuevoMN = 0
 nSaldoTotalNuevoME = 0
 For i = 1 To lnContadorAnte
        If FLante.TextMatrix(i, 1) = "." Then
           If Trim(FLante.TextMatrix(i, 8)) = "" Then
                MsgBox "Ingresar monto de cobertura"
                Exit Sub
           End If
           If CDbl(FLante.TextMatrix(i, 8)) > CDbl(FLante.TextMatrix(i, 7)) Then
                MsgBox "El monto de la garantia no cubre el Gravament"
                FLante.TextMatrix(i, 1) = ""
                FLante.TextMatrix(i, 8) = ""
                Exit Sub
           End If
           nSaldoTotalNuevoMN = nSaldoTotalNuevoMN + IIf(FLante.TextMatrix(i, 6) = "Soles", 1, CCur(txtTipoCambio.Text)) * CCur(FLante.TextMatrix(i, 8))
           nSaldoTotalNuevoME = nSaldoTotalNuevoME + CCur(FLante.TextMatrix(i, 8)) / IIf(FLante.TextMatrix(i, 6) = "Soles", CCur(txtTipoCambio.Text), 1)

           nSaldoTotalNuevoME = Round(nSaldoTotalNuevoME, 2)
           nSaldoTotalNuevoMN = Round(nSaldoTotalNuevoMN, 2)
           If nSaldoTotalNuevoME = nSaldoTotalME Then
                nSaldoTotalNuevoMN = nSaldoTotalMN
           End If
        End If
 Next i
End Sub

Private Sub FLGada_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
Dim oGarantia As COMNCredito.NCOMGarantia
Set oGarantia = New COMNCredito.NCOMGarantia
Dim nCantidad As Integer


nSaldoTotalMN = 0
nSaldoTotalME = 0
lnContadorActGada = 0
Dim i As Integer
nCantidad = 0
    If FLGada.TextMatrix(pnRow, 1) = "." Then
        nCantidad = oGarantia.RecuperaCantidadCreditosxGarantia(FLGada.TextMatrix(pnRow, 3))
        If nCantidad > 1 Then
            MsgBox "Esta garantia es sabana, usar otra garantia", vbInformation
            FLGada.TextMatrix(pnRow, pnCol) = ""
            Exit Sub
        End If
    End If
        For i = 1 To lnContadorGada
            If FLGada.TextMatrix(i, 1) = "." Then
                If FLGada.TextMatrix(i, 6) = "Soles" Then
                    nSaldoTotalMN = nSaldoTotalMN + IIf(FLGada.TextMatrix(i, 6) = "Soles", 1, CCur(txtTipoCambio.Text)) * CCur(FLGada.TextMatrix(i, 7))
                    nSaldoTotalME = nSaldoTotalME + CCur(FLGada.TextMatrix(i, 7)) / IIf(FLGada.TextMatrix(i, 6) = "Soles", CCur(txtTipoCambio.Text), 1)
                Else
                    nSaldoTotalMN = nSaldoTotalMN + IIf(FLGada.TextMatrix(i, 6) = "Dolares", 1, CCur(txtTipoCambio.Text)) * CCur(FLGada.TextMatrix(i, 7))
                    nSaldoTotalME = nSaldoTotalME + CCur(FLGada.TextMatrix(i, 7)) / IIf(FLGada.TextMatrix(i, 6) = "Dolares", CCur(txtTipoCambio.Text), 1)
                End If
            End If
        Next i
        nSaldoTotalMN = Round(nSaldoTotalMN, 2)
        nSaldoTotalME = Round(nSaldoTotalME, 2)
        
        txtPorTipoCambioMN.Text = nSaldoTotalMN
        txtPorTipoCambioME.Text = nSaldoTotalME
        
        For i = 1 To lnContadorGada
        If FLGada.TextMatrix(i, 1) = "." Then
            lnContadorActGada = lnContadorActGada + 1
        End If
    Next i
End Sub

Private Sub Form_Load()
    Dim clsTC As COMDConstSistema.NCOMTipoCambio
    Dim nTC As Double
    Set clsTC = New COMDConstSistema.NCOMTipoCambio
    
    ActXCta.CMAC = gsCodCMAC
    ActXCta.Age = gsCodAge
    cmdReemplazar.Enabled = False
    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
    txtTipoCambio.Text = nTC
    CentraForm Me
End Sub

Private Sub FLAnte_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
Dim oGarantia As COMNCredito.NCOMGarantia
Set oGarantia = New COMNCredito.NCOMGarantia
Dim nCantidad As Integer
Dim i As Integer

nSaldoTotalNuevoMN = 0
nSaldoTotalNuevoME = 0
lnContadorActAnte = 0
nCantidad = 0
    If FLante.TextMatrix(pnRow, 1) = "1" Then
        nCantidad = oGarantia.RecuperaCantidadCreditosxGarantia(FLante.TextMatrix(pnRow, 3))
        If nCantidad = 1 Then
            MsgBox "Esta garantia ya esta usada por otro credito", vbInformation
            FLante.TextMatrix(pnRow, pnCol) = ""
            Exit Sub
        End If
    Else
        FLante.TextMatrix(pnRow, 8) = ""
    End If
    
    For i = 1 To lnContadorAnte
        If FLante.TextMatrix(i, 1) = "." Then
            lnContadorActAnte = lnContadorActAnte + 1
        End If
    Next i
    
End Sub

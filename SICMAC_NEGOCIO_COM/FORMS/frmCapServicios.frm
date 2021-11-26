VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCapServicios 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   Icon            =   "frmCapServicios.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMonto 
      Caption         =   "Monto a Pagar"
      Height          =   1545
      Left            =   45
      TabIndex        =   8
      Top             =   1260
      Width           =   4290
      Begin SICMACT.EditMoney txtMonto 
         Height          =   345
         Left            =   1260
         TabIndex        =   9
         Top             =   225
         Width           =   2055
         _extentx        =   3625
         _extenty        =   609
         font            =   "frmCapServicios.frx":030A
         forecolor       =   16711680
         text            =   "0"
         enabled         =   -1  'True
      End
      Begin VB.Label lblTotalPagar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   1260
         TabIndex        =   14
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total a Pagar :"
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   1125
         Width           =   1050
      End
      Begin VB.Line Line1 
         X1              =   1125
         X2              =   3555
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Label lblComision 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   1260
         TabIndex        =   12
         Top             =   615
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Comisión :"
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   300
         Width           =   540
      End
   End
   Begin MSComctlLib.ImageList imglst 
      Left            =   4575
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCapServicios.frx":0332
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCapServicios.frx":0784
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   350
      Left            =   2340
      TabIndex        =   3
      Top             =   2925
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   350
      Left            =   1260
      TabIndex        =   2
      Top             =   2925
      Width           =   975
   End
   Begin VB.Frame fraServicio 
      Caption         =   "Servicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1080
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4215
      Begin VB.PictureBox pctServ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   3480
         ScaleHeight     =   525
         ScaleWidth      =   525
         TabIndex        =   7
         Top             =   240
         Width           =   555
      End
      Begin MSMask.MaskEdBox txtCodigo 
         Height          =   345
         Left            =   1215
         TabIndex        =   1
         Top             =   628
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtNumRecibo 
         Height          =   350
         Left            =   1215
         TabIndex        =   0
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   705
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Recibo :"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   315
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmCapServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nServicio As COMDConstantes.CaptacOperacion

Dim nmoneda As COMDConstantes.Moneda

Private Function GetValorComision() As Double
Dim rsPar As Recordset
Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion

Set rsPar = oCap.GetTarifaParametro(nServicio, COMDConstantes.gMonedaNacional, COMDConstantes.gCostoComServPublico)
Set oCap = Nothing
If rsPar.EOF And rsPar.BOF Then
    GetValorComision = 0
Else
    GetValorComision = rsPar("nParValor")
End If
rsPar.Close
Set rsPar = Nothing
End Function

Private Sub LimpiaControles()
Select Case nServicio
    Case gServCobSedalib
        txtNumRecibo.Text = "___-________-__"
        txtCodigo.Text = "___________"
    Case gServCobHidrandina
        txtNumRecibo.Text = "___-__-__-____"
        txtCodigo.Text = "________"
    Case gServCobEdelnor
        txtNumRecibo.Text = "_-________"
        txtCodigo.Text = "________"
End Select
TxtMonto.Text = "0.00"
txtNumRecibo.SetFocus
End Sub

Public Sub Inicia(ByVal nServ As COMDConstantes.CaptacOperacion, Optional nMon As COMDConstantes.Moneda = COMDConstantes.gMonedaNacional)
nServicio = nServ
nmoneda = nMon
Select Case nServicio
    Case COMDConstantes.gServCobSedalib
        txtNumRecibo.Mask = "###-########-##"
        txtCodigo.Mask = "###########"
        lblCodigo.Caption = "Código :"
        Set pctServ.Picture = imglst.ListImages(1).Picture
        Me.Caption = "Captaciones - Servicios - SEDALIB"
    Case COMDConstantes.gServCobHidrandina
        txtNumRecibo.Mask = "###-##-##-####"
        txtCodigo.Mask = "########"
        lblCodigo.Caption = "N° Medidor :"
        Set pctServ.Picture = imglst.ListImages(2).Picture
        Me.Caption = "Captaciones - Servicios - HIDRANDINA"
    Case COMDConstantes.gServCobEdelnor
        txtNumRecibo.Mask = "C-########"
        txtCodigo.Mask = "########"
        lblCodigo.Caption = "N° Suministro :"
        Set pctServ.Picture = imglst.ListImages(2).Picture
        Me.Caption = "Captaciones - Servicios - EDELNOR"
End Select
lblComision.Caption = Format$(GetValorComision(), "#,##0.00")
Me.Show 1
End Sub

Private Sub cmdGrabar_Click()
Dim sNumRecibo As String
Dim sCodigo As String
Dim nMonto As Double, nMontoComision As Double
Dim lsBoleta As String
sNumRecibo = Replace(Trim(txtNumRecibo), "_", "", 1, , vbTextCompare)
sNumRecibo = Replace(Trim(txtNumRecibo), "-", "", 1, , vbTextCompare)
sCodigo = Replace(Trim(txtCodigo), "_", "", 1, , vbTextCompare)
sCodigo = Replace(Trim(txtCodigo), "-", "", 1, , vbTextCompare)
nMonto = TxtMonto.value
nMontoComision = CDbl(lblComision.Caption)
If sNumRecibo = "" Then
    MsgBox "Debe digitar un Número de Recibo Válido", vbInformation, "Aviso"
    txtNumRecibo.SetFocus
    Exit Sub
End If
If sCodigo = "" Then
    MsgBox "Debe digitar un " & lblCodigo & " Válido", vbInformation, "Aviso"
    txtCodigo.SetFocus
    Exit Sub
End If
If nMonto = 0 Then
    MsgBox "Monto NO Válido.", vbInformation, "Aviso"
    TxtMonto.SetFocus
    Exit Sub
End If
Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
If nServicio = COMDConstantes.gServCobSedalib Then
    If Not clsServ.ValSedalib(sNumRecibo, sCodigo, nMonto) Then
        MsgBox "Número de Recibo MAL DIGITADO, verifique y corrija por favor.", vbInformation, "Aviso"
        txtNumRecibo.SetFocus
        Exit Sub
    End If
End If
If clsServ.EsReciboCobrado(nServicio, sNumRecibo) Then
    MsgBox "Número de Recibo ya fue cobrado", vbInformation, "Aviso"
    txtNumRecibo.SetFocus
    Set clsServ = Nothing
    Exit Sub
End If
If MsgBox("¿Desea grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    Dim sMovNro As String
    Set clsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov = Nothing
    On Error GoTo ErrGraba
    clsServ.CapCobranzaServicios sMovNro, nServicio, sNumRecibo, sCodigo, nMonto, gsNomCmac, gsNomAge, sLpt, nMontoComision, , , , , gsCodCMAC, lsBoleta
    
    If Trim(lsBoleta) <> "" Then
        nFicSal = FreeFile
        Open sLpt For Output As nFicSal
            Print #nFicSal, lsBoleta
            Print #nFicSal, ""
        Close #nFicSal
    End If

    LimpiaControles
End If
Set clsServ = Nothing
Exit Sub
ErrGraba:
    MsgBox Err.Description, vbExclamation, "Error"
    
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub lblComision_Change()
Dim nMonto As Double, nComision As Double
nMonto = TxtMonto.value
nComision = CDbl(lblComision)
LblTotalPagar.Caption = Format$(nMonto + nComision, "#,##0.00")
End Sub

Private Sub txtcodigo_GotFocus()
With txtCodigo
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtMonto.SetFocus
End If
End Sub

Private Sub txtMonto_Change()
Dim nMonto As Double, nComision As Double
nMonto = TxtMonto.value
nComision = CDbl(lblComision)
LblTotalPagar.Caption = Format$(nMonto + nComision, "#,##0.00")
End Sub

Private Sub txtMonto_GotFocus()
TxtMonto.MarcaTexto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
End If
End Sub

Private Sub txtNumRecibo_GotFocus()
With txtNumRecibo
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtNumRecibo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCodigo.SetFocus
End If
End Sub

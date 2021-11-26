VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajaGenRepFlujos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3045
   ClientLeft      =   1485
   ClientTop       =   2535
   ClientWidth     =   8085
   Icon            =   "frmCajaGenRepFlujos.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   6780
      TabIndex        =   20
      Top             =   2475
      Width           =   1170
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   5625
      TabIndex        =   19
      Top             =   2475
      Width           =   1170
   End
   Begin VB.Frame fraFechas 
      Caption         =   "Fechas"
      Height          =   615
      Left            =   75
      TabIndex        =   14
      Top             =   2280
      Width           =   4215
      Begin MSMask.MaskEdBox txtdesde 
         Height          =   315
         Left            =   705
         TabIndex        =   15
         Top             =   210
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHasta 
         Height          =   315
         Left            =   2760
         TabIndex        =   16
         Top             =   210
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         Caption         =   "Hasta :"
         Height          =   195
         Left            =   2160
         TabIndex        =   18
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label9 
         Caption         =   "Desde :"
         Height          =   195
         Left            =   75
         TabIndex        =   17
         Top             =   270
         Width           =   555
      End
   End
   Begin VB.Frame FracuentaHasta 
      Caption         =   "Datos Generales"
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
      Height          =   1065
      Left            =   45
      TabIndex        =   7
      Top             =   1140
      Width           =   7965
      Begin SICMACT.TxtBuscar txtCtaIFHasta 
         Height          =   315
         Left            =   1155
         TabIndex        =   8
         Top             =   270
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         ForeColor       =   128
      End
      Begin VB.Label lblDescCtaHasta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   1140
         TabIndex        =   13
         Top             =   645
         Width           =   6645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Institución :"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4020
         TabIndex        =   12
         Top             =   322
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cuenta :"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   135
         TabIndex        =   11
         Top             =   690
         Width           =   945
      End
      Begin VB.Label lblDescIFHasta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   4890
         TabIndex        =   10
         Top             =   277
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta IF:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   135
         TabIndex        =   9
         Top             =   322
         Width           =   720
      End
   End
   Begin VB.Frame FraCuentaDesde 
      Caption         =   "Cuenta Desde"
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
      Height          =   1035
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7965
      Begin SICMACT.TxtBuscar txtCtaIFDesde 
         Height          =   315
         Left            =   1155
         TabIndex        =   1
         Top             =   255
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         ForeColor       =   -2147483635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta IF:"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   315
         Width           =   735
      End
      Begin VB.Label lblDescIFDesde 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   4890
         TabIndex        =   5
         Top             =   262
         Width           =   2895
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cuenta :"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   690
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Institución :"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   4035
         TabIndex        =   3
         Top             =   315
         Width           =   810
      End
      Begin VB.Label lblDescCtaDesde 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   1155
         TabIndex        =   2
         Top             =   637
         Width           =   6630
      End
   End
   Begin VB.CommandButton cmdSaldos 
      Caption         =   "&Saldos"
      Height          =   345
      Left            =   4455
      TabIndex        =   21
      Top             =   2475
      Width           =   1185
   End
End
Attribute VB_Name = "frmCajaGenRepFlujos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oOpe As DOperacion
Dim oCtaIf As COMNAuditoria.NCajaCtaIF
Dim oBarra As clsProgressBar
Dim oCont As COMNAuditoria.NContImprimir
Attribute oCont.VB_VarHelpID = -1

Private Sub CmdAceptar_Click()
Dim lsTexto As String
If Valida = False Then Exit Sub
    lsTexto = oCont.ImprimeFlujoCtasBancos(gnColPage, gnLinPage, gsOpeCod, Mid(txtCtaIFDesde, 4, 28), Mid(txtCtaIFHasta, 4, 28), txtdesde, txtHasta)
    If lsTexto = "" Then
        MsgBox "No Existen Movimientos para los datos Ingresados", vbInformation, "Aviso"
        Exit Sub
    End If
    EnviaPrevio lsTexto, "Me.Caption", gnLinPage, False
    'EnviaPrevioTes lsTexto, Me.Caption, gnLinPage, False
End Sub

Function Valida() As Boolean
Valida = True
If Len(Trim(txtCtaIFDesde)) = 0 Then
    MsgBox "Ingrese Cuenta de Institución Financiera Inicial", vbInformation, "Aviso"
    txtCtaIFDesde.SetFocus
    Valida = False
    Exit Function
End If
If Len(Trim(txtCtaIFHasta)) = 0 Then
    MsgBox "Ingrese Cuenta de Institución Financiera Final", vbInformation, "Aviso"
    txtCtaIFHasta.SetFocus
    Valida = False
    Exit Function
End If
If ValFecha(txtdesde) = False Then
    Valida = False
    Exit Function
End If
If ValFecha(txtHasta) = False Then
    Valida = False
    Exit Function
End If

End Function

Private Sub cmdSaldos_Click()
Dim lsTexto As String
On Error GoTo SaldosErr
lsTexto = oCont.ImprimeSaldosCtaIf(gnColPage, gnLinPage, gsOpeCod, txtHasta)
If lsTexto = "" Then
    MsgBox "No Existen Movimientos para los datos Ingresados", vbInformation, "Aviso"
    Exit Sub
End If
EnviaPrevio lsTexto, "Me.Caption", gnLinPage, False
'EnviaPrevioTes lsTexto, Me.Caption, gnLinPage, False
Exit Sub
SaldosErr:
    MsgBox Err.Description
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub Form_Load()
Set oOpe = New DOperacion
                'MsgBox "Antes de: Set oCtaIf = New COMNAuditoria.NCajaCtaIF"
Set oCtaIf = New COMNAuditoria.NCajaCtaIF
                'MsgBox "Antes de: Set oBarra = New clsProgressBar"
Set oBarra = New clsProgressBar
                'MsgBox "Antes de: Set oCont = New COMNAuditoria.NContImprimir"
Set oCont = New COMNAuditoria.NContImprimir
                'MsgBox "Antes de: CentraForm Me"
CentraForm Me
                'MsgBox "Antes de: txtCtaIFDesde.rs"
txtCtaIFDesde.rs = oOpe.GetRsOpeObj(gsOpeCod, "0")
                'MsgBox "Antes de: txtCtaIFHasta.rs"
txtCtaIFHasta.rs = oOpe.GetRsOpeObj(gsOpeCod, "0")
                'MsgBox "Antes de: Caption = gsOpeDesc"
Caption = gsOpeDesc
                'MsgBox "Antes de:txtdesde = gdFecSis"
txtdesde = gdFecSis
                'MsgBox "Antes de:txtHasta = gdFecSis"
txtHasta = gdFecSis

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oBarra = Nothing
Set oOpe = Nothing
Set oCtaIf = Nothing
Set oCont = Nothing
End Sub


Private Sub oCont_BarraClose()
oBarra.CloseForm Me
End Sub

Private Sub oCont_BarraProgress(value As Variant, psTitulo As String, psSubTitulo As String, psTituloBarra As String, ColorLetras As ColorConstants)
oBarra.Progress value, psTitulo, psSubTitulo, psTituloBarra, ColorLetras
End Sub

Private Sub oCont_BarraShow(pnMax As Variant)
oBarra.ShowForm Me
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.Max = pnMax
End Sub

Private Sub txtCtaIFDesde_EmiteDatos()
lblDescCtaDesde = oCtaIf.EmiteTipoCuentaIF(Mid(txtCtaIFDesde, 18, 10)) + " " + txtCtaIFDesde.psDescripcion
lblDescIFDesde = oCtaIf.NombreIF(Mid(txtCtaIFDesde, 4, 13))
txtCtaIFHasta.SetFocus
End Sub

Private Sub txtCtaIFHasta_EmiteDatos()
lblDescCtaHasta = oCtaIf.EmiteTipoCuentaIF(Mid(txtCtaIFHasta, 18, 10)) + " " + txtCtaIFHasta.psDescripcion
lblDescIFHasta = oCtaIf.NombreIF(Mid(txtCtaIFHasta, 4, 13))
txtdesde.SetFocus
End Sub

Private Sub txtDesde_GotFocus()
fEnfoque txtdesde
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtHasta.SetFocus
End If
End Sub

Private Sub txtHasta_GotFocus()
fEnfoque txtHasta
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdAceptar.SetFocus
End If
End Sub

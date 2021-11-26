VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajeroReportes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes de Cajero"
   ClientHeight    =   4545
   ClientLeft      =   2025
   ClientTop       =   2175
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCajeroReportes.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin Sicmact.TxtBuscar txtBuscarUser 
      Height          =   345
      Left            =   4815
      TabIndex        =   8
      Top             =   1275
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   609
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      sTitulo         =   ""
      ForeColor       =   8388608
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4395
      TabIndex        =   7
      Top             =   4035
      Width           =   1200
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   4395
      TabIndex        =   6
      Top             =   3660
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango Fechas"
      Height          =   1005
      Left            =   3945
      TabIndex        =   1
      Top             =   120
      Width           =   1920
      Begin MSMask.MaskEdBox txtdesde 
         Height          =   315
         Left            =   705
         TabIndex        =   2
         Top             =   225
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   -2147483635
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txthasta 
         Height          =   300
         Left            =   690
         TabIndex        =   3
         Top             =   585
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         _Version        =   393216
         ForeColor       =   -2147483635
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
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
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
         Left            =   90
         TabIndex        =   5
         Top             =   270
         Width           =   570
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   105
         TabIndex        =   4
         Top             =   615
         Width           =   525
      End
   End
   Begin MSComctlLib.TreeView TreeRepo 
      Height          =   4185
      Left            =   135
      TabIndex        =   0
      Top             =   150
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   7382
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   706
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3825
      Top             =   2805
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   17
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajeroReportes.frx":030A
            Key             =   "LibroClose"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajeroReportes.frx":06D0
            Key             =   "LibroOpen"
         EndProperty
      EndProperty
   End
   Begin Sicmact.TxtBuscar txtbuscarAge 
      Height          =   345
      Left            =   4815
      TabIndex        =   10
      Top             =   1650
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   609
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      sTitulo         =   ""
      ForeColor       =   8388608
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Agencia :"
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
      Left            =   3990
      TabIndex        =   11
      Top             =   1710
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User :"
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
      Left            =   3990
      TabIndex        =   9
      Top             =   1335
      Width           =   480
   End
End
Attribute VB_Name = "frmCajeroReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsKey  As String
Dim WithEvents oCajeroImp As NCajeroImp
Attribute oCajeroImp.VB_VarHelpID = -1
Dim oBarra As clsProgressBar
Private Sub ListadoReportes()
Dim nodX As Node
Set nodX = TreeRepo.Nodes.Add(, , "A", "Reportes Cajero", "LibroClose")
nodX.Expanded = True
nodX.ForeColor = vbBlue
Set nodX = TreeRepo.Nodes.Add("A", tvwChild, "A1", "Faltante y/o Sobrante", "LibroOpen")
Set nodX = TreeRepo.Nodes.Add("A", tvwChild, "A2", "Protocolo", "LibroOpen")
Set nodX = TreeRepo.Nodes.Add("A", tvwChild, "A3", "Compra / Venta ME", "LibroOpen")
Set nodX = TreeRepo.Nodes.Add("A", tvwChild, "A4", "Saldos Diarios en ME", "LibroOpen")
Set nodX = TreeRepo.Nodes.Add("A", tvwChild, "A5", "Operaciones de Notas de Abono y Cargo", "LibroOpen")
Set nodX = TreeRepo.Nodes.Add("A", tvwChild, "A6", "Operaciones Varias", "LibroOpen")
Set nodX = TreeRepo.Nodes.Add("A", tvwChild, "A7", "Extornos", "LibroOpen")
End Sub
Private Sub cmdImprimir_Click()
Dim lsTexto As String
Select Case lsKey
    Case "A1"
             
    Case "A3"
        
        lsTexto = oCajeroImp.GeneraRepoCompraVenta(gOpeCajeroMECompra, Me.txtDesde, txtHasta, txtBuscarUser, txtbuscarAge, _
                gsNomCmac, gsNomAge, "AHORRO ", gdFecSis, gnColPage, gnLinPage)
                
        lsTexto = lsTexto + oImpresora.gPrnSaltoPagina + oCajeroImp.GeneraRepoCompraVenta(gOpeCajeroMEVenta, Me.txtDesde, txtHasta, txtBuscarUser, txtbuscarAge, _
                gsNomCmac, gsNomAge, "AHORRO ", gdFecSis, gnColPage, gnLinPage)
        If Len(Trim(lsTexto)) <= 10 Then
            MsgBox "Datos no Encontrados", vbInformation, "Aviso"
        Else
            EnviaPrevio lsTexto, Me.Caption, gnLinPage, False
        End If
    Case Else
End Select
End Sub

Private Sub CmdSalir_Click()
Unload Me

End Sub

Private Sub Form_Load()
Dim oArea As DActualizaDatosArea
Dim oGen As DGeneral
Set oGen = New DGeneral
Set oArea = New DActualizaDatosArea
Set oBarra = New clsProgressBar
Set oCajeroImp = New NCajeroImp
txtbuscarAge.rs = oArea.GetAgencias
txtbuscarAge = gsCodAge
txtBuscarUser.psRaiz = "Usuarios"
txtBuscarUser.rs = oGen.GetUserAreaAgencia(gsCodArea, gsCodAge)

CentraForm Me
ListadoReportes
txtDesde = gdFecSis
txtHasta = gdFecSis
Set oGen = Nothing
Set oArea = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMdiMain.Enabled = True
frmMdiMain.SetFocus
End Sub

Private Sub oCajeroImp_CloseProgress()
oBarra.CloseForm Me
End Sub
Private Sub oCajeroImp_Progress(ByVal pnValor As Long, ByVal pnTotal As Long)
oBarra.Max = pnTotal
oBarra.Progress pnValor, "Reporte de Compra/Venta ME", "", ""
End Sub
Private Sub oCajeroImp_ShowProgress()
oBarra.ShowForm Me
End Sub
Private Sub TreeRepo_NodeClick(ByVal Node As MSComctlLib.Node)
lsKey = Node.Key

End Sub

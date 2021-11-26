VERSION 5.00
Begin VB.Form FrmLogOCompraPenalidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Penalidad"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7740
   Icon            =   "FrmLogOCompraPenalidad.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   7575
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   315
         Left            =   6300
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   315
         Left            =   5100
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   120
      TabIndex        =   5
      Top             =   60
      Width           =   7575
      Begin VB.Label LblCodProv 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1140
         TabIndex        =   15
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Dirección:"
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
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   960
         Width           =   885
      End
      Begin VB.Label lblRUC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1140
         TabIndex        =   10
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label lblDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1140
         TabIndex        =   9
         Top             =   900
         Width           =   5835
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "R.U.C:"
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
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   600
         Width           =   585
      End
      Begin VB.Label lblProveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2400
         TabIndex        =   7
         Top             =   180
         Width           =   4575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor:"
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
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Width           =   945
      End
   End
   Begin Sicmact.EditMoney edMonto 
      Height          =   315
      Left            =   1260
      TabIndex        =   3
      Top             =   1620
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12648447
      Text            =   "0"
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtConcepto 
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2340
      Width           =   7575
   End
   Begin VB.Label lblMonto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2640
      TabIndex        =   4
      Top             =   1620
      Width           =   5055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Concepto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2100
      Width           =   7545
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Importe:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   1620
      Width           =   705
   End
End
Attribute VB_Name = "FrmLogOCompraPenalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
    ReDim datos(1)
    datos(0).Concepto = ""
    datos(0).monto = 0
    datos(0).bit = 0
    datos(0).Nombre = Me.lblProveedor
    datos(0).ruc = Me.lblRuc
    datos(0).direccion = Me.lblDireccion
    datos(0).Concepto = Me.txtConcepto.Text
    datos(0).monto = Me.edMonto.value '& " " & Me.lblMonto
    datos(0).bit = 1
    CmdCancelar_Click
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub edMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtConcepto.SetFocus
    End If
End Sub

Private Sub edMonto_LostFocus()
    lblMonto = ConvNumLet(edMonto.value)
End Sub

Private Sub Form_Activate()
    Me.edMonto.SetFocus
End Sub

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.CmdAceptar.SetFocus
    End If
End Sub

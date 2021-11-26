VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmValCtaAhoAnt 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   Icon            =   "frmValCtaAhoAnt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1335
      Width           =   1035
   End
   Begin VB.Frame fraDato 
      Caption         =   "Dato"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1155
      Left            =   1860
      TabIndex        =   7
      Top             =   120
      Width           =   3195
      Begin MSMask.MaskEdBox txtCuenta 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###-###-###-#######A"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTarjeta 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-##-############ "
         PromptChar      =   "_"
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1380
      TabIndex        =   4
      Top             =   1335
      Width           =   1035
   End
   Begin VB.Frame fraTipoBusq 
      Caption         =   "Buscar por :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1155
      Left            =   60
      TabIndex        =   6
      Top             =   120
      Width           =   1755
      Begin VB.OptionButton optTipoBusq 
         Caption         =   "Tarjeta"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   720
         Width           =   1035
      End
      Begin VB.OptionButton optTipoBusq 
         Caption         =   "Código Antiguo"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   300
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmValCtaAhoAnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nProducto As Producto
Dim sCuenta As String

Public Function Inicia(ByVal nProd As Producto, Optional bTarjeta As Boolean = False) As String
If bTarjeta Then
    optTipoBusq(1).Visible = True
    txtTarjeta.Visible = True
    Me.Caption = "Relación Cuenta Antigua - Nuevo y Tarjeta Magnética"
Else
    optTipoBusq(1).Visible = False
    txtTarjeta.Visible = False
    Me.Caption = "Relación Cuenta Antigua - Nuevo"
End If
nProducto = nProd

Select Case nProducto
    Case gCapAhorros
        txtCuenta.Mask = "###-###-##-#########A"
        txtCuenta.Text = "108-___-__-__________"
    Case gCapCTS, gCapPlazoFijo
        txtCuenta.Mask = "###-###-##-#########A"
        txtCuenta.Text = "108-___-__-__________"
    Case gColConsuPrendario
        txtCuenta.Mask = "###-###-##-#########A"
        txtCuenta.Text = "108-___-__-__________"
    Case Else
        txtCuenta.Mask = "###-###-##-#########A"
        txtCuenta.Text = "108-___-__-__________"
End Select
sCuenta = ""
Me.Show 1
Inicia = sCuenta
End Function

Private Sub CmdAceptar_Click()
If optTipoBusq(0).value = True Then
    Dim clsGen As DGeneral
    Dim sCuentaAnt As String
    sCuentaAnt = Trim(Replace(txtCuenta.Text, "-", "", 1, , vbTextCompare))
    sCuentaAnt = Trim(Replace(sCuentaAnt, "_", "", 1, , vbTextCompare))
    If sCuentaAnt <> "" Then
        Set clsGen = New DGeneral
        sCuenta = clsGen.GetCuentaNueva(sCuentaAnt)
    Else
        MsgBox "Cuenta Incorrecta, por favor digite una cuenta de 12 dígitos", vbInformation, "Aviso"
        txtCuenta.SetFocus
        Exit Sub
    End If
End If
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Me.Icon = LoadPicture(App.path & gsRutaIcono)
CentraForm Me
End Sub

Private Sub optTipoBusq_Click(Index As Integer)
If Index = 0 Then
    txtTarjeta.Visible = False
    txtCuenta.Visible = True
    txtCuenta.SetFocus
ElseIf Index = 1 Then
    txtCuenta.Visible = False
    txtTarjeta.Visible = True
    txtTarjeta.SetFocus
End If
End Sub

Private Sub txtCuenta_GotFocus()
With txtCuenta
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtTarjeta_GotFocus()
With txtTarjeta
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

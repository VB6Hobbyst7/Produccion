VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAnexo15 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anexo 15 A"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10230
   Icon            =   "frmAnexo15.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   10230
   Begin VB.TextBox txtEstadoEncajeAcumuladoME 
      Height          =   375
      Left            =   7920
      TabIndex        =   27
      Text            =   "1434801.71"
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox txtEstadoEncajeAcumuladoMN 
      Height          =   375
      Left            =   5040
      TabIndex        =   26
      Text            =   "170098.48"
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10095
      Begin VB.TextBox txtPatrimonioEfec 
         Height          =   375
         Left            =   5040
         TabIndex        =   28
         Text            =   "49035908.60"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdImprimir 
         Cancel          =   -1  'True
         Caption         =   "Generar"
         Height          =   375
         Left            =   7080
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   8760
         TabIndex        =   15
         Top             =   360
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label19 
         Caption         =   "Patrimonio Efectivo"
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
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8520
      TabIndex        =   13
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Frame frValorEncaje 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   10095
      Begin VB.TextBox txtDepositoBCRAcumuladoME 
         BackColor       =   &H80000001&
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   7920
         TabIndex        =   25
         Text            =   "0"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtDepositoBCRAcumuladoMN 
         Height          =   375
         Left            =   5040
         TabIndex        =   24
         Text            =   "318334795.7"
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txtToseAcumuladoME 
         Height          =   375
         Left            =   7920
         TabIndex        =   21
         Text            =   "148119810.77"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtToseAcumuladoMN 
         Height          =   375
         Left            =   5040
         TabIndex        =   20
         Text            =   "2794941273.51"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtValorEncajeME 
         Height          =   375
         Left            =   7920
         TabIndex        =   8
         Text            =   "22818799.19"
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtValorEncaje 
         Height          =   375
         Left            =   5040
         TabIndex        =   7
         Text            =   "487818618.19"
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtToseME 
         Height          =   375
         Left            =   7920
         TabIndex        =   6
         Text            =   "246689721"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtToseMN 
         Height          =   375
         Left            =   5040
         TabIndex        =   5
         Text            =   "4298631122"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtFondoEncaje 
         Height          =   375
         Left            =   5040
         TabIndex        =   4
         Text            =   "2514905.20"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtFondoEncajeME 
         Height          =   375
         Left            =   7920
         TabIndex        =   3
         Text            =   "0"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txTasaS 
         Height          =   375
         Left            =   5040
         TabIndex        =   2
         Text            =   "4.28"
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox txtTasaD 
         Height          =   375
         Left            =   7920
         TabIndex        =   1
         Text            =   "0"
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   10080
         Y1              =   100
         Y2              =   100
      End
      Begin VB.Line Line5 
         X1              =   10080
         X2              =   10080
         Y1              =   4370
         Y2              =   100
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   10080
         Y1              =   4370
         Y2              =   4370
      End
      Begin VB.Line Line3 
         X1              =   4440
         X2              =   10080
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line2 
         X1              =   4440
         X2              =   4440
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line1 
         X1              =   7440
         X2              =   7440
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Label Label5 
         Caption         =   "Estado de encaje acumulado a la fecha"
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
         Left            =   240
         TabIndex        =   23
         Top             =   3360
         Width           =   3735
      End
      Begin VB.Label Label4 
         Caption         =   "Deposito BCR Acumulado a la fecha"
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
         Left            =   240
         TabIndex        =   22
         Top             =   2880
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   "Tose acumulado a la fecha"
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
         Left            =   240
         TabIndex        =   19
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Moneda Extranjera"
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
         Left            =   8040
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda Nacional"
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
         Left            =   5040
         TabIndex        =   17
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label13 
         Caption         =   "Encaje exigible (FEB. 2011)"
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
         Left            =   240
         TabIndex        =   12
         ToolTipText     =   "Encaje exigible"
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label Label16 
         Caption         =   "TOSE Base(FEB. 2011)"
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
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   4575
      End
      Begin VB.Label Label17 
         Caption         =   "Promedio de caja de mes anterior"
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
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   3855
      End
      Begin VB.Label Label20 
         Caption         =   "Tasa Promedio Overnight"
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
         Left            =   240
         TabIndex        =   9
         Top             =   3840
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmAnexo15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ldFecha As Date
Dim lsOpeCod As String
Dim lsMoneda As String
Dim lnTipoCambio As Currency

Private Sub cmdImprimir_Click()
    If CDbl(IIf(txtValorEncaje.Text = "", 0, txtValorEncaje.Text)) = 0 Then
        MsgBox "Ingresar valor encaje", vbCritical
        Exit Sub
    End If
    If CDbl(IIf(txtValorEncajeME = "", 0, txtValorEncajeME)) = 0 Then
        MsgBox "Ingresar valor encaje", vbCritical
        Exit Sub
    End If
    If CDbl(IIf(txtToseMN.Text = "", 0, txtToseMN.Text)) = 0 Then
        MsgBox "Ingresar TOSE", vbCritical
        Exit Sub
    End If
    If CDbl(IIf(txtToseME.Text = "", 0, txtToseME.Text)) = 0 Then
        MsgBox "Ingresar TOSE", vbCritical
        Exit Sub
    End If
    If CDbl(IIf(txtToseAcumuladoMN.Text = "", 0, txtToseAcumuladoMN.Text)) = 0 Then
        MsgBox "Ingresar TOSE acumulado", vbCritical
        Exit Sub
    End If
    If CDbl(IIf(txtToseAcumuladoME.Text = "", 0, txtToseAcumuladoME.Text)) = 0 Then
        MsgBox "Ingresar TOSE acumulado", vbCritical
        Exit Sub
    End If
    
    If CDbl(IIf(txtDepositoBCRAcumuladoMN.Text = "", 0, txtDepositoBCRAcumuladoMN.Text)) = 0 Then
        MsgBox "Ingresar deposito BCR", vbCritical
        Exit Sub
    End If
    'ALPA 20110609
'    If CInt(IIf(txtDepositoBCRAcumuladoME.Text = "", 0, txtDepositoBCRAcumuladoME.Text)) = 0 Then
'        MsgBox "Ingresar deposito BCR", vbCritical
'        Exit Sub
'    End If
    
    If CDbl(IIf(txtEstadoEncajeAcumuladoMN.Text = "", 0, txtEstadoEncajeAcumuladoMN.Text)) = 0 Then
        MsgBox "Ingresar encaje acumulado", vbCritical
        Exit Sub
    End If
    If CDbl(IIf(txtEstadoEncajeAcumuladoME.Text = "", 0, txtEstadoEncajeAcumuladoME.Text)) = 0 Then
        MsgBox "Ingresar encaje acumulado", vbCritical
        Exit Sub
    End If
    frmAnx15AReporteNew.ImprimeAnexo15A lsOpeCod, lsMoneda, ldFecha, lnTipoCambio, txtValorEncaje.Text, txtValorEncajeME.Text, txtToseMN.Text, txtToseME.Text, txtFondoEncaje.Text, txtFondoEncajeME.Text, txtPatrimonioEfec.Text, txTasaS.Text, txtTasaD.Text, txtToseAcumuladoMN.Text, txtToseAcumuladoME.Text, txtDepositoBCRAcumuladoMN.Text, txtDepositoBCRAcumuladoME.Text, txtEstadoEncajeAcumuladoMN.Text, txtEstadoEncajeAcumuladoME.Text
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub ImprimirAnexo15A(psOpeCod As String, psMoneda As String, pdFecha As Date, Optional ByVal pnTipoCambio As Currency = 1)
    ldFecha = pdFecha
    lsOpeCod = psOpeCod
    lsMoneda = psMoneda
    lnTipoCambio = pnTipoCambio
    Show
End Sub

Private Sub Form_Load()
txtFecha.Text = ldFecha
End Sub

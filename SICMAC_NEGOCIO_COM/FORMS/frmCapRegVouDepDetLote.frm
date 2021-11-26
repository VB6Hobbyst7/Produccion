VERSION 5.00
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCapRegVouDepDetLote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CTS Depósito en Lote"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3090
   Icon            =   "frmCapRegVouDepDetLote.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   3090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   885
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   885
   End
   Begin Spinner.uSpinner uspNroCli 
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Max             =   99999
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
   End
   Begin VB.Label Label1 
      Caption         =   "Número de Clientes:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   165
      Width           =   1575
   End
End
Attribute VB_Name = "frmCapRegVouDepDetLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************
'** Nombre : frmCapRegVouDepDetLote
'** Descripción : Registro de Número que podran aperturar cuentas segun TI-ERS044-2013
'** Creación : EJVG, 20130906 21:12:00 PM
'***************************************************************************************
Option Explicit
Dim fnNroCliIni As Long
Dim fbAceptar As Boolean

Private Sub Form_Unload(Cancel As Integer)
    fbAceptar = False
End Sub
Public Function Inicio(ByVal pnCantidad As Long) As Long
    fbAceptar = False
    Call SetNroClientes(pnCantidad)
    Show 1
    Inicio = GetNroClientes
End Function
Private Sub SetNroClientes(ByVal pnCantidad As Long)
    uspNroCli.valor = pnCantidad
    fnNroCliIni = pnCantidad
End Sub
Private Function GetNroClientes() As Long
    If fbAceptar Then
        GetNroClientes = uspNroCli.valor
    Else
        GetNroClientes = fnNroCliIni
    End If
End Function
Private Sub cmdAceptar_Click()
    fbAceptar = True
    Hide
End Sub
Private Sub cmdCancelar_Click()
    fbAceptar = False
    Hide
End Sub
Private Sub uspNroCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub

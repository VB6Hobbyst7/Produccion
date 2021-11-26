VERSION 5.00
Begin VB.Form frmColocEvalPorProducto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Productos por Evaluar"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmColocEvalPorProducto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkProducto 
      Caption         =   "Pignoraticio"
      Height          =   315
      Index           =   8
      Left            =   960
      TabIndex        =   9
      Top             =   4080
      Width           =   2490
   End
   Begin VB.CheckBox chkProducto 
      Caption         =   "Hipotecario"
      Height          =   315
      Index           =   7
      Left            =   960
      TabIndex        =   8
      Top             =   3600
      Width           =   2490
   End
   Begin VB.CheckBox chkProducto 
      Caption         =   "Consumo No Revolvente"
      Height          =   315
      Index           =   6
      Left            =   960
      TabIndex        =   7
      Top             =   3120
      Width           =   2490
   End
   Begin VB.CheckBox chkProducto 
      Caption         =   "Consumo Revolvente"
      Height          =   315
      Index           =   5
      Left            =   960
      TabIndex        =   6
      Top             =   2640
      Width           =   2490
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   465
      Left            =   1560
      TabIndex        =   5
      Top             =   4800
      Width           =   1590
   End
   Begin VB.CheckBox chkProducto 
      Caption         =   "Microempresas"
      Height          =   315
      Index           =   4
      Left            =   975
      TabIndex        =   4
      Top             =   2100
      Width           =   2490
   End
   Begin VB.CheckBox chkProducto 
      Caption         =   "Pequeñas Empresas"
      Height          =   315
      Index           =   3
      Left            =   975
      TabIndex        =   3
      Top             =   1650
      Width           =   2490
   End
   Begin VB.CheckBox chkProducto 
      Caption         =   "Medianas Empresas"
      Height          =   315
      Index           =   2
      Left            =   975
      TabIndex        =   2
      Top             =   1200
      Width           =   2490
   End
   Begin VB.CheckBox chkProducto 
      Caption         =   "Grandes Empresas"
      Height          =   315
      Index           =   1
      Left            =   975
      TabIndex        =   1
      Top             =   750
      Width           =   2490
   End
   Begin VB.CheckBox chkProducto 
      Caption         =   "Corporativos"
      Height          =   315
      Index           =   0
      Left            =   975
      TabIndex        =   0
      Top             =   300
      Width           =   2490
   End
End
Attribute VB_Name = "frmColocEvalPorProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MatIndices As Variant

Private Sub CmdAceptar_Click()
Dim i As Integer
'ReDim MatIndices(5)
ReDim MatIndices(9)

'For i = 0 To 4
For i = 0 To 8
    If chkProducto(i).value = 1 Then
        MatIndices(i) = 1
    Else
        MatIndices(i) = 0
    End If
Next i
Unload Me
End Sub

Private Sub Form_Load()
    Call CentraForm(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not IsArray(MatIndices) Then
   ReDim MatIndices(0)
End If
End Sub

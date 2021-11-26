VERSION 5.00
Begin VB.Form frmCredAlertaTempranaDet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredAlertaTempranaDet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDet 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.Frame frameLeyenda 
         Caption         =   "Leyenda:"
         ForeColor       =   &H8000000C&
         Height          =   1185
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Width           =   6855
         Begin VB.Label lblLeyenda 
            Caption         =   "lblLeyenda"
            Height          =   735
            Left            =   120
            TabIndex        =   11
            Top             =   285
            Width           =   6615
         End
      End
      Begin VB.TextBox txtDetalleValor 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   350
         Left            =   960
         TabIndex        =   7
         Top             =   1440
         Width           =   1515
      End
      Begin VB.TextBox txtDetalle 
         Enabled         =   0   'False
         Height          =   350
         Left            =   960
         TabIndex        =   6
         Top             =   915
         Width           =   6135
      End
      Begin VB.TextBox txtLimite 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   350
         Left            =   3645
         TabIndex        =   5
         Top             =   1440
         Width           =   1400
      End
      Begin VB.TextBox txtFormula 
         Enabled         =   0   'False
         Height          =   350
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   6135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   1080
         X2              =   7065
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         X1              =   1080
         X2              =   7080
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label lblUnidadMedida 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5120
         TabIndex        =   9
         Top             =   1515
         Width           =   90
      End
      Begin VB.Label lblValorCalculado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   375
         TabIndex        =   8
         Top             =   1485
         Width           =   585
      End
      Begin VB.Label lblValor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor límite :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2520
         TabIndex        =   3
         Top             =   1485
         Width           =   1110
      End
      Begin VB.Label lblDetalle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lblFormula 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fórmula :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   435
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmCredAlertaTempranaDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************
'** Nombre : frmCredAlertaTempranaDet
'** Descripción : Muestra las alertas tempranas por crédito
'** Referencia : ERS001-2017 - Metodología para monitoreo de Alertas Tempranas de Crédito - Detalle
'** Creación : LUCV, 20170131 13:51:01 PM
'**************************************************************************************************
Option Explicit
Dim fsTitulo As String
Dim fsFormula As String
Dim fsDetalle As String
Dim fnValor As Double
Dim fnValorLimite As Double
Dim fnEstado As Boolean
Dim fsUnidad As String
Dim fsLeyenda As String
Dim fbOk As Boolean

Private Sub Form_Load()
    fbOk = False
End Sub
Public Function Inicio(ByVal psTitulo As String, ByVal psFormula As String, ByVal psDetalleValor As String, ByVal pnValor As Double, ByVal psValorLimite As Double, ByVal pnEstado As Boolean, ByVal psUnidad As String, ByVal psLeyenda As String) As Boolean
    fsTitulo = psTitulo
    fsFormula = psFormula
    fsDetalle = psDetalleValor
    fnValor = pnValor
    fnValorLimite = psValorLimite
    fnEstado = pnEstado
    fsUnidad = psUnidad
    fsLeyenda = psLeyenda
    Call CargarDatos
    Show 1
    Inicio = fbOk
End Function

Private Sub CargarDatos()
    Caption = fsTitulo
    txtFormula.Text = fsFormula
    txtDetalle.Text = fsDetalle
    If fnEstado = 0 Then
        txtDetalleValor.BackColor = &HC0FFC0
    Else
       txtDetalleValor.BackColor = &HC0C0FF
    End If
    txtDetalleValor.Text = fnValor
    txtLimite.Text = fnValorLimite
    lblUnidadMedida.Caption = fsUnidad
    lblLeyenda.Caption = fsLeyenda
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
    If KeyCode = 113 And Shift = 0 Then
        KeyCode = 10
    End If
    If KeyCode = 27 And Shift = 0 Then
        Unload Me
        fbOk = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    fbOk = True
End Sub

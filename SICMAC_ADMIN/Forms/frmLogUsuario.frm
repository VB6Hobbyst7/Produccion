VERSION 5.00
Begin VB.Form frmLogUsuario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relación de Usuarios"
   ClientHeight    =   5070
   ClientLeft      =   1050
   ClientTop       =   1935
   ClientWidth     =   10065
   Icon            =   "frmLogUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   Begin Sicmact.FlexEdit fgeArea 
      Height          =   4185
      Left            =   135
      TabIndex        =   2
      Top             =   315
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   7382
      Cols0           =   4
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-cAreaCod-Nombre del Area-Estructura"
      EncabezadosAnchos=   "400-500-2400-800"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X"
      ListaControles  =   "0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L"
      FormatosEdit    =   "0-0-0-0"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      lbOrdenaCol     =   -1  'True
      Appearance      =   0
      RowHeight0      =   240
   End
   Begin Sicmact.FlexEdit fgeUsuario 
      Height          =   4185
      Left            =   4755
      TabIndex        =   1
      Top             =   300
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   7382
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-cPersCod-Usuario-Nombre del Usuario-cPersEstado"
      EncabezadosAnchos=   "450-0-700-2700-1000"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L-L"
      FormatosEdit    =   "0-0-0-0-0"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      lbOrdenaCol     =   -1  'True
      Appearance      =   0
      RowHeight0      =   240
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   7815
      TabIndex        =   0
      Top             =   4575
      Width           =   1305
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Usuarios por Area"
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
      Height          =   210
      Index           =   1
      Left            =   4875
      TabIndex        =   4
      Top             =   45
      Width           =   1650
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Areas"
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
      Height          =   210
      Index           =   0
      Left            =   255
      TabIndex        =   3
      Top             =   60
      Width           =   750
   End
End
Attribute VB_Name = "frmLogUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsDGnral As DLogGeneral

Private Sub cmdSalir_Click()
    Set clsDGnral = Nothing
    Unload Me
End Sub

Private Sub fgeArea_Click()
    Dim rs As ADODB.Recordset
    Set rs = clsDGnral.CargaUsuario(UsuTodosArea, fgeArea.TextMatrix(fgeArea.Row, 1))
    If rs.RecordCount > 0 Then
        Set fgeUsuario.Recordset = rs
    Else
        fgeUsuario.Clear
        fgeUsuario.FormaCabecera
        fgeUsuario.Rows = 2
    End If
    Set rs = Nothing
End Sub

Private Sub fgeUsuario_DblClick()
    If fgeUsuario.TextMatrix(fgeUsuario.Row, 2) <> "" Then
        gsCodUser = fgeUsuario.TextMatrix(fgeUsuario.Row, 2)
        MsgBox "Cargado el usuario : " & gsCodUser
        MDISicmact.Caption = gcTitModulo & Space(20) & gsDBName & Space(10) & gsServerName & "   " & gsCodUser & " " & gdFecSis
    End If
End Sub

Private Sub Form_Load()
    Set clsDGnral = New DLogGeneral
    Set fgeArea.Recordset = clsDGnral.CargaArea(AreaTotal)
    Call CentraForm(Me)
    Call fgeArea_Click
End Sub

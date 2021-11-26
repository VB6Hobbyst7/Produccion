VERSION 5.00
Begin VB.Form frmPigDeudaPendiente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Deuda Pendiente "
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5955
   Icon            =   "frmPigDeudaPendiente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   4485
      TabIndex        =   1
      Top             =   3195
      Width           =   1320
   End
   Begin SICMACT.FlexEdit feContratos 
      Height          =   3105
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   5477
      Cols0           =   4
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Item-Contrato-Estado-Deuda"
      EncabezadosAnchos=   "500-2400-1600-1200"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      EncabezadosAlineacion=   "C-C-C-R"
      FormatosEdit    =   "0-0-0-2"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmPigDeudaPendiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Inicia(ByVal psCodPers As String)
Dim oDatos As DPigContrato
Dim rs As Recordset

    Set oDatos = New DPigContrato
    Set rs = oDatos.dObtieneDeudaTotalDetallada(psCodPers, gdFecSis)
    
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            feContratos.AdicionaFila
            feContratos.TextMatrix(feContratos.Row, 1) = rs!cCtaCod
            feContratos.TextMatrix(feContratos.Row, 2) = rs!Estado
            feContratos.TextMatrix(feContratos.Row, 3) = rs!DeudaTotal
            rs.MoveNext
        Loop
    End If
    
    Set oDatos = Nothing
    Me.Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

VERSION 5.00
Begin VB.Form frmCapAdmNivelesDialog 
   Caption         =   "Detalle Niveles Aprobación"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5910
   Icon            =   "frmCapAdmNivelesDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4830
      TabIndex        =   2
      Top             =   5145
      Width           =   975
   End
   Begin VB.Frame frame1 
      Caption         =   "Detalle de Agencias y TEA"
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
      Height          =   4950
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   5685
      Begin SICMACT.FlexEdit grdAgencia 
         Height          =   4455
         Left            =   105
         TabIndex        =   1
         Top             =   345
         Width           =   5460
         _extentx        =   9631
         _extenty        =   7858
         cols0           =   6
         fixedcols       =   0
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "id-Agencia-TEA Adic.-grupo-Dias-SubPr"
         encabezadosanchos=   "0-2000-1000-0-800-800"
         font            =   "frmCapAdmNivelesDialog.frx":030A
         font            =   "frmCapAdmNivelesDialog.frx":0336
         font            =   "frmCapAdmNivelesDialog.frx":0362
         font            =   "frmCapAdmNivelesDialog.frx":038E
         fontfixed       =   "frmCapAdmNivelesDialog.frx":03BA
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0"
         encabezadosalineacion=   "L-L-C-C-C-C"
         formatosedit    =   "0-0-0-0-0-0"
         textarray0      =   "id"
         selectionmode   =   1
         lbbuscaduplicadotext=   -1  'True
         appearance      =   0
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCapAdmNivelesDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsLocal As ADODB.Recordset

Public Sub cargarRs(pRs As ADODB.Recordset)
    Set rsLocal = pRs
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cargarGrid
End Sub

Private Sub cargarGrid()
     Dim i As Integer
     grdAgencia.rsFlex = rsLocal
     For i = 1 To grdAgencia.Rows - 1
        grdAgencia.TextMatrix(i, 2) = Format$(grdAgencia.TextMatrix(i, 2), "#,##0.00")
        grdAgencia.TextMatrix(i, 5) = IIf(grdAgencia.TextMatrix(i, 5) = 1, "Si", "No")
     Next
End Sub

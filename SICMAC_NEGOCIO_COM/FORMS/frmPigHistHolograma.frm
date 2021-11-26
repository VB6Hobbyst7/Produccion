VERSION 5.00
Begin VB.Form frmPigHistHolograma 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "N° Pendientes"
   ClientHeight    =   2400
   ClientLeft      =   10770
   ClientTop       =   5340
   ClientWidth     =   2235
   Icon            =   "frmPigHistHolograma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   2235
   ShowInTaskbar   =   0   'False
   Begin SICMACT.FlexEdit FlexHolog 
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      _ExtentX        =   10186
      _ExtentY        =   3413
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Nro. Holog."
      EncabezadosAnchos=   "300-1200"
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
      ColumnasAEditar =   "X-X"
      ListaControles  =   "0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C"
      FormatosEdit    =   "5-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   6
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmPigHistHolograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* HISTORIAL DE HOLOGRAMAS
'Archivo:  frmPigHistHolograma.frm
'APRI   :  15/05/2019
'Resumen:  Nos permite visualizar los Nro. de holgramas pendientes.
Public Sub Inicio(ByVal pnCod As Long)
    Dim obj As New COMDColocPig.DCOMColPContrato
    Dim rs As ADODB.Recordset
    Dim i As Integer
  
    FlexHolog.Clear
    FormateaFlex FlexHolog
        Set rs = obj.ObtieneHologramasPendientes(pnCod)
        If Not (rs.EOF And rs.BOF) Then
            For i = 1 To rs.RecordCount
                FlexHolog.AdicionaFila
                FlexHolog.TextMatrix(i, 1) = rs!nHolograma
                rs.MoveNext
            Next i
        End If
        Me.Show 1

End Sub

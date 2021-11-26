VERSION 5.00
Begin VB.Form frmCredReprogSolicExonera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exoneraciones a solicitar"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   Icon            =   "frmCredReprogSolicExonera.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4560
      TabIndex        =   1
      Top             =   3120
      Width           =   1170
   End
   Begin SICMACT.FlexEdit feExonera 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5106
      Cols0           =   4
      HighLight       =   1
      EncabezadosNombres=   "-cExoneraCod-Exoneracion-Solicitar"
      EncabezadosAnchos=   "0-0-4500-800"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-3"
      ListaControles  =   "0-0-0-4"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "L-L-L-C"
      FormatosEdit    =   "0-1-0-0"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   3
      RowHeight0      =   300
   End
End
Attribute VB_Name = "frmCredReprogSolicExonera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmCredReprogSolicExonera
'** Descripción : Formulario para listar las exoneraciones a solicitar para reprogramar créditos
'**               según TI-ERS010-2016
'** Creación : JUEZ, 20160331 09:00:00 AM
'*****************************************************************************************************

Option Explicit

Dim rsExonera As ADODB.Recordset
Dim nTipoOpcion As Integer 'Agrego JOEP20171214 ACTA220-2017

Public Function ObtieneExoneracionesSolicitud(ByVal nTipoOp As Integer) As ADODB.Recordset
    nTipoOpcion = nTipoOp 'Agrego JOEP20171214 ACTA220-2017
    CargarExoneraciones
    Me.Show 1
    Set ObtieneExoneracionesSolicitud = rsExonera
End Function

Private Sub CargarExoneraciones()
Dim oDCred As COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
Dim lnFila As Integer

    Set oDCred = New COMDCredito.DCOMCredito
        Set rs = oDCred.RecuperaColocReprogExonera(nTipoOpcion)
    Set oDCred = Nothing
    
    If Not (rs.EOF And rs.BOF) Then
        lnFila = 0
        Do While Not rs.EOF
            feExonera.AdicionaFila
            lnFila = feExonera.row
            feExonera.TextMatrix(lnFila, 1) = rs!nExoneraCod
            feExonera.TextMatrix(lnFila, 2) = rs!cExoneraDesc
            rs.MoveNext
        Loop
        feExonera.TopRow = 1
        feExonera.row = 1
    End If
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsExonera = IIf(feExonera.rows - 1 > 0, feExonera.GetRsNew(), Nothing)
End Sub

VERSION 5.00
Begin VB.Form frmColPObservacionesRetasacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alertas de Retasación"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8055
   Icon            =   "frmColPObservacionesRetasacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SICMACT.FlexEdit FEObservacion 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   7880
      _ExtentX        =   13891
      _ExtentY        =   3836
      Cols0           =   6
      ScrollBars      =   2
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Prenda-Kilataje-P. Bruto-P. Neto-Observación"
      EncabezadosAnchos=   "400-2500-800-800-800-2500"
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
      ColumnasAEditar =   "X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483628
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-C-C-C-L"
      FormatosEdit    =   "0-0-0-0-0-0"
      TextArray0      =   "#"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblObservadas 
      Caption         =   "lblObservadas"
      Height          =   255
      Left            =   6840
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Observadas:"
      Height          =   255
      Left            =   5880
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblRetasadas 
      Caption         =   "lblRetasadas"
      Height          =   255
      Left            =   6840
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Retasadas:"
      Height          =   255
      Left            =   5880
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblNroRetasacion 
      Caption         =   "lblNroRetasacion"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblNroContrato 
      Caption         =   "lblNroContrato"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "N° Retasación:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "N° Contrato:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmColPObservacionesRetasacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************
'CREADO:         TORE 28/05/2018
'DESCRIPCIÓN:    Creado para la visualización de las retasaciones de
'                las joyas según ERS054-2017
'*********************************************************

Public Sub Observaciones(ByVal pscPersCod As String, ByVal psCtaCod As String, ByVal pnTpoProceso As Integer)
    Dim oDR As New COMDColocPig.DCOMColPContrato
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Set rs = oDR.ObtenerObservacionRetasacion(pscPersCod, psCtaCod, pnTpoProceso)
    If Not (rs.EOF And rs.BOF) Then
        lblNroContrato.Caption = rs!cPigCod
        lblNroRetasacion.Caption = rs!nNroRetasacion
        lblRetasadas.Caption = rs!Retasadas
        For i = 1 To rs.RecordCount
            FEObservacion.AdicionaFila
            FEObservacion.TextMatrix(i, 1) = rs!cDescrip
            FEObservacion.TextMatrix(i, 2) = rs!cKilataje & "K"
            FEObservacion.TextMatrix(i, 3) = Format(rs!nPesoBruto, gcFormView)
            FEObservacion.TextMatrix(i, 4) = Format(rs!nPesoNeto, gcFormView)
            FEObservacion.TextMatrix(i, 5) = rs!cObservaciones
            rs.MoveNext
        Next
        lblObservadas.Caption = CStr(i - 1)
        Show 1
    End If
    'Set rs = Nothing
    'Set oDR = Nothing
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

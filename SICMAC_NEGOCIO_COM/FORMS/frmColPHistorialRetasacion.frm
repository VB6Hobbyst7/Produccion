VERSION 5.00
Begin VB.Form frmColPHistorialRetasacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historial Retasacíon"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10320
   Icon            =   "frmColPHistorialRetasacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9000
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Retasacíon"
      Height          =   2775
      Left            =   50
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin SICMACT.FlexEdit FEDatos 
         Height          =   2415
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   4260
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Piezas-Material-PBruto-PNeto-Tasac-Descripcíon-Observacíon"
         EncabezadosAnchos=   "400-600-650-800-800-0-3500-3000"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-R-R-R-L-L"
         FormatosEdit    =   "0-0-0-2-2-2-0-0"
         TextArray0      =   "#"
         Enabled         =   0   'False
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmColPHistorialRetasacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmColPHistorialRetasacion
'** Descripción : Formulario que muestra la retasacion de un credito prendario
'** Creación    : RECO, 20140707 - ERS074-2014
'**********************************************************************************************
Option Explicit
Dim lsCtaCod As String

Private Sub Form_Load()
    With Screen
    Move (.Width - Width) / 2, (.Height - Height) / 2, Width, Height
    End With
End Sub

Public Sub CargarDatos(ByVal psCtaCod As String)
    Dim oColp As New COMNColoCPig.NCOMColPContrato
    Dim rs As New ADODB.Recordset
    Dim x As Integer
    
    Set rs = oColp.DevuelveHistorialRetasacion(psCtaCod)
    If Not (rs.BOF And rs.EOF) Then
        FEDatos.Clear
        FormateaFlex FEDatos
        For x = 1 To rs.RecordCount
            FEDatos.AdicionaFila
            FEDatos.TextMatrix(x, 1) = rs!nPiezas
            FEDatos.TextMatrix(x, 2) = rs!cKilataje
            FEDatos.TextMatrix(x, 3) = rs!nPesoBruto
            FEDatos.TextMatrix(x, 4) = rs!nPesoNeto
            FEDatos.TextMatrix(x, 5) = rs!nValTasac
            FEDatos.TextMatrix(x, 6) = rs!cDescrip
            FEDatos.TextMatrix(x, 7) = rs!cObservaciones
            rs.MoveNext
        Next
    End If
    Set rs = Nothing
End Sub

Public Sub Inicio(ByVal psCtaCod As String)
    lsCtaCod = psCtaCod
    Call CargarDatos(lsCtaCod)
    Me.Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub



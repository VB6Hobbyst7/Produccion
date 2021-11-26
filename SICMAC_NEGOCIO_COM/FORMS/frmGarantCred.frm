VERSION 5.00
Begin VB.Form frmGarantCred 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Créditos Asociados a la garantía"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   Icon            =   "frmGarantCred.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   7080
      TabIndex        =   13
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   " Créditos Asociados a la Garantía"
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Left            =   120
      TabIndex        =   9
      Top             =   1260
      Width           =   8295
      Begin SICMACT.FlexEdit feCreditos 
         Height          =   2655
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4683
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nº Crédito-Titular-Moneda-Monto-Monto Garantía"
         EncabezadosAnchos=   "400-1400-3000-700-1200-1200"
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
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos de la Garantía "
      ForeColor       =   &H00FF0000&
      Height          =   1140
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin VB.Label lblDisponible 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6360
         TabIndex        =   8
         Top             =   675
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Disponible:"
         Height          =   255
         Left            =   5520
         TabIndex        =   7
         Top             =   705
         Width           =   855
      End
      Begin VB.Label lblVRM 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3720
         TabIndex        =   6
         Top             =   675
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "VRM:"
         Height          =   255
         Left            =   3240
         TabIndex        =   5
         Top             =   705
         Width           =   495
      End
      Begin VB.Label lblValorComercial 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1440
         TabIndex        =   4
         Top             =   675
         Width           =   1695
      End
      Begin VB.Label lblDescripcion 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   6615
      End
      Begin VB.Label Label2 
         Caption         =   "Valor Comercial:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   710
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   390
         Width           =   975
      End
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   600
      TabIndex        =   12
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Total:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4380
      Width           =   495
   End
End
Attribute VB_Name = "frmGarantCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmGarantCred
'** Descripción : Formulario que muestra la lista de créditos relacionados a la garantía
'** Creación    : RECO, 20150421 - ERS010-2015
'**********************************************************************************************

Option Explicit
Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Public Sub Inicio(ByVal psNumGarant As String)
    If CargarDatos(psNumGarant) = True Then Me.Show 1
End Sub

Private Function CargarDatos(ByVal psNumGarant As String) As Boolean
    Dim obj As New COMNCredito.NCOMGarantia
    Dim RS As New ADODB.Recordset
    Dim nTotal As Double, i As Integer
    
    Set RS = obj.ObtieneCreditosPorGarantia(psNumGarant)
    If Not (RS.EOF And RS.BOF) Then
        lblDescripcion.Caption = RS!cDescripcion
        lblValorComercial.Caption = Format(RS!nTasacion, gsFormatoNumeroView)
        lblVRM.Caption = Format(RS!nRealizacion, gsFormatoNumeroView)
        lblDisponible.Caption = Format(RS!nDisponible, gsFormatoNumeroView)
        feCreditos.Clear
        FormateaFlex feCreditos
        For i = 1 To RS.RecordCount
            feCreditos.AdicionaFila
            feCreditos.TextMatrix(i, 1) = RS!cCtaCod
            feCreditos.TextMatrix(i, 2) = RS!cPersNombre
            feCreditos.TextMatrix(i, 3) = RS!cmoneda
            feCreditos.TextMatrix(i, 4) = Format(RS!nMontoCol, gsFormatoNumeroView)
            feCreditos.TextMatrix(i, 5) = Format(RS!nGravado, gsFormatoNumeroView)
            nTotal = nTotal + RS!nGravado
            RS.MoveNext
        Next
        lblTotal.Caption = Format(nTotal, gsFormatoNumeroView)
        CargarDatos = True
    Else
        CargarDatos = False
    End If
End Function

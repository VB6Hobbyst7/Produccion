VERSION 5.00
Begin VB.Form frmCapHistoricoRenovacionesDPF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historial de Renovaciones"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   Icon            =   "frmCapHistoricoRenovacionesDPF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin SICMACT.FlexEdit FEHistorico 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   3836
      Cols0           =   6
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "N° renov-Fecha-TEA-Monto-Plazo-Forma Retiro"
      EncabezadosAnchos=   "700-1200-1200-1200-1200-2500"
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
      TextArray0      =   "N° renov"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   705
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label lblTasaApertura 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblNombreApertura 
      Caption         =   "TEA Apertura :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmCapHistoricoRenovacionesDPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***************************************************************************************************
'***Nombre      : frmCapHistoricoRenovacionesDPF
'***Descripción : Formulario para DPF - Histórico de Renovaciones.
'***Creación    : ELRO el 20130111, según OYP-RFC123-2012
'***************************************************************************************************

Private fsCuenta As String

Public Sub iniciarHistoricoRenovacionesDPF(ByVal psCuenta As String)
fsCuenta = psCuenta
Show 1
End Sub

Private Sub obtenerHistoricoRenovacionesDPF(ByVal psCuenta As String)
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rsHist As ADODB.Recordset
    Dim i As Integer
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsHist = New ADODB.Recordset
    Set rsHist = clsMant.obtenerHistoricoRenovacionesDPF(psCuenta)
    
    Call LimpiaFlex(FEHistorico)
    i = 1
    FEHistorico.lbEditarFlex = True
        
    If Not (rsHist.EOF And rsHist.BOF) Then
        Do While Not rsHist.EOF
            FEHistorico.AdicionaFila
            FEHistorico.TextMatrix(i, 1) = rsHist!cFechaRenovacion
            FEHistorico.TextMatrix(i, 2) = Format$(ConvierteTNAaTEA(rsHist!nTEA), "#,##0.00")
            FEHistorico.TextMatrix(i, 3) = Format(rsHist!nSaldoDisponible, gcFormView)
            FEHistorico.TextMatrix(i, 4) = rsHist!nPlazo
            FEHistorico.TextMatrix(i, 5) = rsHist!cConsDescripcion
            i = i + 1
            rsHist.MoveNext
        Loop
    End If
    
    FEHistorico.lbEditarFlex = False
    Set rsHist = Nothing
    Set clsMant = Nothing
End Sub

Private Sub Form_Load()
    CentraForm Me
    'JAME20140509 ***
    Dim oCapta As New DCOMCaptaGenerales
    lblTasaApertura.Caption = Format(oCapta.DevuelveTasaAperturaPlazoFijo(fsCuenta), "#0.00")
    Set oCapta = Nothing
    'END JAME *******
    obtenerHistoricoRenovacionesDPF fsCuenta
End Sub


Private Sub cmdAceptar_Click()
Unload Me
End Sub

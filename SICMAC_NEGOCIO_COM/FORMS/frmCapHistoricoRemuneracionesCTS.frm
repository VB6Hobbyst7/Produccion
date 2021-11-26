VERSION 5.00
Begin VB.Form frmCapHistoricoRemuneracionesCTS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CTS - Histórico de Registro de n Remun. Brutas"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8430
   Icon            =   "frmCapHistoricoRemuneracionesCTS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Histórico de Registro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8235
      Begin SICMACT.FlexEdit FEHistorico 
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4471
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Nro-Fecha-Moneda-Total 4 Rem. Bruto-Agencia-Usuario"
         EncabezadosAnchos=   "500-1200-1200-1650-1650-1200"
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
         EncabezadosAlineacion=   "C-C-C-R-L-C"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "Nro"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCapHistoricoRemuneracionesCTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***************************************************************************************************
'***Nombre      : frmCapIndComDep
'***Descripción : Formulario para CTS - Histórico de Registro de 6 Remun. Brutas.
'***Creación    : ELRO el 20121015, según OYP-RFC101-2012
'***************************************************************************************************

Private fsCuenta As String

Public Sub iniciarHistoricoRemBruCTS(ByVal psCuenta As String)
'COMENTADO POR APRI20170707 (MEJORA) EL CODIGO SE PUSO EN EL LOAD
''JUEZ 20151114 **********************************************
'Dim clsDef As New COMNCaptaGenerales.NCOMCaptaDefinicion
'Dim nParRemBrutas As Integer
'nParRemBrutas = clsDef.GetCapParametroNew(gCapCTS, 0)!nUltRemunBrutas
'Me.Caption = "CTS - Histórico de Registro de " & CStr(nParRemBrutas) & " Remun. Brutas"
'FEHistorico.EncabezadosNombres = "Nro-Fecha-Moneda-Total " & CStr(nParRemBrutas) & " Rem. Bruto-Agencia-Usuario"
'Set clsDef = Nothing
''END JUEZ ***************************************************
fsCuenta = psCuenta
Show 1
End Sub


Private Sub obtieneHistoricoCTS(ByVal psCuenta As String)
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rsHist As ADODB.Recordset
    Dim i As Integer
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsHist = New ADODB.Recordset
    Set rsHist = clsMant.obtenerHistorialCaptacSueldosCTS(psCuenta)
    
    Call LimpiaFlex(FEHistorico)
    i = 1
    FEHistorico.lbEditarFlex = True
        
    If Not (rsHist.EOF And rsHist.BOF) Then
        Do While Not rsHist.EOF
            FEHistorico.AdicionaFila
            FEHistorico.TextMatrix(i, 1) = rsHist!cFecha
            FEHistorico.TextMatrix(i, 2) = rsHist!cMoneda
            FEHistorico.TextMatrix(i, 3) = Format(rsHist!nSueldoTotal, gcFormView)
            FEHistorico.TextMatrix(i, 4) = rsHist!cAgeDescripcion
            FEHistorico.TextMatrix(i, 5) = rsHist!cUser
            i = i + 1
            rsHist.MoveNext
        Loop
    End If
    
    FEHistorico.lbEditarFlex = False
End Sub

Private Sub Form_Load()
    CentraForm Me
    'JUEZ 20151114 **********************************************
    Dim clsDef As New COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim nParRemBrutas As Integer
    nParRemBrutas = clsDef.GetCapParametroNew(gCapCTS, 0)!nUltRemunBrutas
    Me.Caption = "CTS - Histórico de Registro de " & CStr(nParRemBrutas) & " Remun. Brutas"
    FEHistorico.EncabezadosNombres = "Nro-Fecha-Moneda-Total " & CStr(nParRemBrutas) & " Rem. Bruto-Agencia-Usuario"
    Set clsDef = Nothing
    'END JUEZ ***************************************************
    obtieneHistoricoCTS fsCuenta
End Sub

VERSION 5.00
Begin VB.Form frmCredMoraCuotas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuotas en Mora"
   ClientHeight    =   3420
   ClientLeft      =   2505
   ClientTop       =   3240
   ClientWidth     =   6420
   Icon            =   "frmCredMora.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   420
      Left            =   2430
      TabIndex        =   2
      Top             =   2895
      Width           =   1620
   End
   Begin VB.Frame Frame1 
      Height          =   2805
      Left            =   90
      TabIndex        =   0
      Top             =   15
      Width           =   6240
      Begin SICMACT.FlexEdit FEMora 
         Height          =   2445
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   4313
         Cols0           =   5
         FixedCols       =   0
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Cuota-Fecha Venc.-Monto-Mora-Cancelacion"
         EncabezadosAnchos=   "600-1200-1200-1200-1200"
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
            Weight          =   700
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
         EncabezadosAlineacion=   "C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "Cuota"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   600
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCredMoraCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub MostarMoraDetalle(ByVal psCtaCod As String, ByVal pdHoy As Date)
Dim i As Integer
Dim K As Integer
Dim nMonto As Double
Dim nMora As Double
Dim nCancel As Double
Dim nPos As Integer
Dim pMatCalend As Variant
Dim oCred As COMNCredito.NCOMCredito

Set oCred = New COMNCredito.NCOMCredito

pMatCalend = oCred.RecuperaMatrizCalendarioPendiente(psCtaCod)

Set oCred = Nothing


    nPos = 1
    For i = 0 To UBound(pMatCalend) - 1
        
        If CDate(pMatCalend(i, 0)) < pdHoy Then
            If nPos > 1 Then
                FEMora.AdicionaFila
            End If
            nMonto = 0
            nMora = 0
            nCancel = 0
            nMonto = CDbl(pMatCalend(i, 3)) + CDbl(pMatCalend(i, 4)) + CDbl(pMatCalend(i, 6))
            nMora = CDbl(pMatCalend(i, 6))
            For K = 0 To i
                nCancel = nCancel + CDbl(pMatCalend(K, 3)) + CDbl(pMatCalend(K, 4)) + CDbl(pMatCalend(K, 5)) + CDbl(pMatCalend(K, 6)) + CDbl(pMatCalend(K, 7)) + CDbl(pMatCalend(K, 8)) + CDbl(pMatCalend(K, 9)) + CDbl(pMatCalend(K, 11))
            Next K
            FEMora.TextMatrix(nPos, 0) = pMatCalend(i, 1)
            FEMora.TextMatrix(nPos, 1) = pMatCalend(i, 0)
            FEMora.TextMatrix(nPos, 2) = Format(nMonto, "#0.00")
            FEMora.TextMatrix(nPos, 3) = Format(nMora, "#0.00")
            FEMora.TextMatrix(nPos, 4) = Format(nCancel, "#0.00")
            nPos = nPos + 1
            
        End If
    Next i
    Me.Show 1
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
End Sub

VERSION 5.00
Begin VB.Form frmBancoPagadorExtornoAbonos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extorno de Abonos"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8925
   Icon            =   "frmBancoPagadorExtornoAbonos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "&Extornar"
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Operaciones de Abono"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin TarjAdm.FlexEdit flxProcesosExtorno 
         Height          =   2715
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   4789
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Archivo-Fecha Creación-Num. Reg-Usuario Procesa-Fecha Procesa"
         EncabezadosAnchos=   "300-2100-1800-1200-1500-1500-0"
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
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C"
         TextArray0      =   "N°"
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmBancoPagadorExtornoAbonos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExtornar_Click()
    Dim nIDProc As Integer
    Dim psPersCod As String
    
    If MsgBox("¿Esta seguro de extornar el proceso de abono seleccionado?", vbYesNo + vbInformation, "Aviso") = vbNo Then
        Exit Sub
    End If
    
        If flxProcesosExtorno.TextMatrix(flxProcesosExtorno.Row, 1) <> "" Then
            nIDProc = flxProcesosExtorno.TextMatrix(flxProcesosExtorno.Row, 6)
            Me.MousePointer = 1
            
            ''Realiza el extorno
            Call RealizaExtornoProcesoInicial(nIDProc, gsCodUser)
            MsgBox "Extorno Finalizado", vbInformation, "Aviso"
            Call CargarProcesosaExtornar
        End If
End Sub

Private Sub Form_Load()
    Call CargarProcesosaExtornar
End Sub

Public Sub CargarProcesosaExtornar()
  Dim nNumFila As Integer
    Dim rsProcExt As ADODB.Recordset
    Set rsProcExt = DevuelveProcesosaExtornar()
    
    Call LimpiaFlex(flxProcesosExtorno)
    If rsProcExt.RecordCount > 0 Then
        Me.flxProcesosExtorno.Clear
        Me.flxProcesosExtorno.Rows = 2
        Me.flxProcesosExtorno.FormaCabecera
        
        Do While Not rsProcExt.EOF
            
            If flxProcesosExtorno.TextMatrix(1, 1) <> "" Then
                flxProcesosExtorno.AdicionaFila
            End If
            
            nNumFila = flxProcesosExtorno.Rows - 1
            
            flxProcesosExtorno.TextMatrix(nNumFila, 0) = nNumFila
            flxProcesosExtorno.TextMatrix(nNumFila, 1) = rsProcExt!cArchivoNombre
            flxProcesosExtorno.TextMatrix(nNumFila, 2) = rsProcExt!dCabFecha_3
            flxProcesosExtorno.TextMatrix(nNumFila, 3) = rsProcExt!nCabNumRegistros_4
            flxProcesosExtorno.TextMatrix(nNumFila, 4) = rsProcExt!cUserProcesa
            flxProcesosExtorno.TextMatrix(nNumFila, 5) = rsProcExt!dFechaProcesa
            flxProcesosExtorno.TextMatrix(nNumFila, 6) = rsProcExt!nIDProc
            rsProcExt.MoveNext
        Loop
    Else
        MsgBox "No existen Operaciones para Extornar", vbInformation, "Aviso"
    End If
End Sub

VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCapIndComDep 
   Caption         =   "Definir Parámetros"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7740
   Icon            =   "frmCapIndComDep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5318
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Indicadores"
      TabPicture(0)   =   "frmCapIndComDep.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FEIndicador"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin SICMACT.FlexEdit FEIndicador 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4260
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Indicadores-Indice-nIndTpo-Id_IndConCli"
         EncabezadosAnchos=   "500-5500-1000-0-0"
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
         ColumnasAEditar =   "X-X-2-X-X"
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-C-C"
         FormatosEdit    =   "0-0-2-0-0"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCapIndComDep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***************************************************************************************************
'***Nombre      : frmCapIndComDep
'***Descripción : Formulario para Definir Parametros de los Indicadores de Comportamiento Depósitos.
'***Creación    : ELRO el 20120918, según OYP-RFC087-2012
'***************************************************************************************************

Private fsIndice As String

Private Sub devolverIndComDep()
Dim oNCOMCaptaReportes As New COMNCaptaGenerales.NCOMCaptaReportes
Dim rsIndComDep As New ADODB.Recordset
Dim i As Integer

Set rsIndComDep = oNCOMCaptaReportes.devolverIndComDep

If Not (rsIndComDep.BOF And rsIndComDep.EOF) Then
    LimpiaFlex FEIndicador
    FEIndicador.lbEditarFlex = True
    i = 1
    Do While Not rsIndComDep.EOF
        FEIndicador.AdicionaFila
        FEIndicador.TextMatrix(i, 1) = rsIndComDep!cConsDescripcion
        FEIndicador.TextMatrix(i, 2) = Format(rsIndComDep!nIndice, "##,##0.00") & "%"
        fsIndice = Format(rsIndComDep!nIndice, "##,##0.00") & "%"
        FEIndicador.TextMatrix(i, 3) = rsIndComDep!nIndTpo
        FEIndicador.TextMatrix(i, 4) = rsIndComDep!Id_IndComDep
        rsIndComDep.MoveNext
        i = i + 1
    Loop
Else
    MsgBox "No existen parametros.", vbInformation, "Aviso"
End If
End Sub

Private Sub modificarIndComDep()
    Dim oNCOMCaptaReportes As New COMNCaptaGenerales.NCOMCaptaReportes
    Dim lnId_IndComDep As Long
    Dim lnFila As Integer
    Dim lnIndice As Currency
    
    lnFila = FEIndicador.Row
    
    If InStr(Trim(FEIndicador.TextMatrix(lnFila, 2)), "%") > 0 Then
        If IsNumeric(Left(Trim(FEIndicador.TextMatrix(lnFila, 2)), Len(Trim(FEIndicador.TextMatrix(lnFila, 2))) - 1)) Then
            lnIndice = CCur(Left(Trim(FEIndicador.TextMatrix(lnFila, 2)), Len(Trim(FEIndicador.TextMatrix(lnFila, 2))) - 1))
        Else
            MsgBox "Ingresar un número", vbInformation, "Aviso"
            FEIndicador.TextMatrix(lnFila, 2) = fsIndice
            Exit Sub
        End If
    Else
        If IsNumeric(Trim(FEIndicador.TextMatrix(lnFila, 2))) Then
            lnIndice = CCur(Trim(FEIndicador.TextMatrix(lnFila, 2)))
        Else
            MsgBox "Ingresar un número", vbInformation, "Aviso"
            FEIndicador.TextMatrix(lnFila, 2) = fsIndice
            Exit Sub
        End If
    End If
    If lnIndice = 0 Then
        MsgBox "El Índice no debe ser cero.", vbInformation, "Aviso"
        FEIndicador.TextMatrix(lnFila, 2) = fsIndice
        Exit Sub
    ElseIf lnIndice < 0 Then
        MsgBox "El Índice no debe ser negativo.", vbInformation, "Aviso"
        FEIndicador.TextMatrix(lnFila, 2) = fsIndice
        Exit Sub
    End If
    
    lnId_IndComDep = oNCOMCaptaReportes.modificarIndComDep(CInt(FEIndicador.TextMatrix(lnFila, 4)), lnIndice)
    
    If lnId_IndComDep > 0 Then
        FEIndicador.TextMatrix(lnFila, 2) = Format(lnIndice, "##,##0.00") & "%"
    Else
        FEIndicador.TextMatrix(lnFila, 2) = fsIndice
    End If
End Sub

Private Sub FEIndicador_OnCellChange(pnRow As Long, pnCol As Long)
    modificarIndComDep
End Sub

Private Sub Form_Load()
    devolverIndComDep
End Sub



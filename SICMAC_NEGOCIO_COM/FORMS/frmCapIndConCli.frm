VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCapIndConCli 
   Caption         =   "Definir Parámetros"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7395
   Icon            =   "frmCapIndConCli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4260
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Indicadores"
      TabPicture(0)   =   "frmCapIndConCli.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FEIndicador"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin SICMACT.FlexEdit FEIndicador 
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2778
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Indicadores-SBS-Interno-nIndTpo-Id_IndConCli"
         EncabezadosAnchos=   "500-4000-1000-1000-0-0"
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
         ColumnasAEditar =   "X-X-2-3-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-R-C-C"
         FormatosEdit    =   "0-0-2-2-0-0"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label2 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   7080
         TabIndex        =   3
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "* N mayor a"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2040
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCapIndConCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***************************************************************************************************
'***Nombre      : frmCapIndConCli
'***Descripción : Formulario para Definir Parametros de los Indicadores de Concetración del Cliente.
'***Creación    : ELRO el 20120905, según OYP-RFC087-2012
'***************************************************************************************************

Private fsSBS, fsInterno As String

Private Sub devolverIndConCli()
Dim oNCOMCaptaReportes As New COMNCaptaGenerales.NCOMCaptaReportes
Dim rsIndConClin As New ADODB.Recordset
Dim i As Integer

Set rsIndConClin = oNCOMCaptaReportes.devolverIndConCli

If Not (rsIndConClin.BOF And rsIndConClin.EOF) Then
    LimpiaFlex FEIndicador
    FEIndicador.lbEditarFlex = True
    i = 1
    Do While Not rsIndConClin.EOF
        FEIndicador.AdicionaFila
        FEIndicador.TextMatrix(i, 1) = rsIndConClin!cConsDescripcion
        FEIndicador.TextMatrix(i, 2) = Format(rsIndConClin!nSBS, "##,##0.00") & "%"
        fsSBS = Format(rsIndConClin!nSBS, "##,##0.00") & "%"
        FEIndicador.TextMatrix(i, 3) = Format(rsIndConClin!nInterno, "##,##0.00") & "%"
        fsInterno = Format(rsIndConClin!nInterno, "##,##0.00") & "%"
        FEIndicador.TextMatrix(i, 4) = rsIndConClin!nIndTpo
        FEIndicador.TextMatrix(i, 5) = rsIndConClin!Id_IndConCli
        rsIndConClin.MoveNext
        i = i + 1
    Loop
Else
    MsgBox "No existen parametros.", vbInformation, "Aviso"
End If
End Sub

Private Sub modificarIndConCli(ByVal pnFila As Long, ByVal pnColumna As Long)
    Dim oNCOMCaptaReportes As New COMNCaptaGenerales.NCOMCaptaReportes
    Dim lnId_IndConCli As Long
    Dim lnFila As Integer
    Dim lnSBS, lnInterno As Currency
    
    lnFila = pnFila
    
    If pnColumna = 2 Then
    
        If InStr(Trim(FEIndicador.TextMatrix(lnFila, 2)), "%") > 0 Then
            If IsNumeric(Left(Trim(FEIndicador.TextMatrix(lnFila, 2)), Len(Trim(FEIndicador.TextMatrix(lnFila, 2))) - 1)) Then
                lnSBS = CCur(Left(Trim(FEIndicador.TextMatrix(lnFila, 2)), Len(Trim(FEIndicador.TextMatrix(lnFila, 2))) - 1))
                lnInterno = CCur(Left(Trim(FEIndicador.TextMatrix(lnFila, 3)), Len(Trim(FEIndicador.TextMatrix(lnFila, 3))) - 1))
            Else
                MsgBox "Ingresar un número", vbInformation, "Aviso"
                FEIndicador.TextMatrix(lnFila, 2) = fsSBS
                Exit Sub
            End If
        Else
            If IsNumeric(Trim(FEIndicador.TextMatrix(lnFila, 2))) Then
                lnSBS = CCur(Trim(FEIndicador.TextMatrix(lnFila, 2)))
                lnInterno = CCur(Left(Trim(FEIndicador.TextMatrix(lnFila, 3)), Len(Trim(FEIndicador.TextMatrix(lnFila, 3))) - 1))
            Else
                MsgBox "Ingresar un número", vbInformation, "Aviso"
                FEIndicador.TextMatrix(lnFila, 2) = fsSBS
                Exit Sub
            End If
        End If
       
          
        If lnSBS = 0 Then
            MsgBox "El Índice SBS no debe ser cero.", vbInformation, "Aviso"
            FEIndicador.TextMatrix(lnFila, 2) = fsSBS
            Exit Sub
        ElseIf lnSBS < 0 Then
            MsgBox "El Índice SBS no debe ser negativo.", vbInformation, "Aviso"
            FEIndicador.TextMatrix(lnFila, 2) = fsSBS
            Exit Sub
        End If
    
    ElseIf pnColumna = 3 Then
    
        If InStr(Trim(FEIndicador.TextMatrix(lnFila, 3)), "%") > 0 Then
            If IsNumeric(Left(Trim(FEIndicador.TextMatrix(lnFila, 3)), Len(Trim(FEIndicador.TextMatrix(lnFila, 3))) - 1)) Then
                lnSBS = CCur(Left(Trim(FEIndicador.TextMatrix(lnFila, 2)), Len(Trim(FEIndicador.TextMatrix(lnFila, 2))) - 1))
                lnInterno = CCur(Left(Trim(FEIndicador.TextMatrix(lnFila, 3)), Len(Trim(FEIndicador.TextMatrix(lnFila, 3))) - 1))
            Else
                MsgBox "Ingresar un número", vbInformation, "Aviso"
                FEIndicador.TextMatrix(lnFila, 3) = fsInterno
                Exit Sub
            End If
        Else
            If IsNumeric(Trim(FEIndicador.TextMatrix(lnFila, 3))) Then
                lnSBS = CCur(Left(Trim(FEIndicador.TextMatrix(lnFila, 2)), Len(Trim(FEIndicador.TextMatrix(lnFila, 2))) - 1))
                lnInterno = CCur(Trim(FEIndicador.TextMatrix(lnFila, 3)))
            Else
                MsgBox "Ingresar un número", vbInformation, "Aviso"
                FEIndicador.TextMatrix(lnFila, 3) = fsInterno
                Exit Sub
            End If
        End If
        
        If lnInterno = 0 Then
            MsgBox "El Índice Interno no debe ser cero.", vbInformation, "Aviso"
            FEIndicador.TextMatrix(lnFila, 3) = fsInterno
            Exit Sub
        ElseIf lnInterno < 0 Then
            MsgBox "El Índice Interno no debe ser negativo.", vbInformation, "Aviso"
            FEIndicador.TextMatrix(lnFila, 3) = fsInterno
            Exit Sub
        End If
    
    End If

    lnId_IndConCli = oNCOMCaptaReportes.modificarIndConCli(CInt(FEIndicador.TextMatrix(lnFila, 5)), lnSBS, lnInterno)
    
    If lnId_IndConCli > 0 Then
        FEIndicador.TextMatrix(lnFila, 2) = Format(lnSBS, "##,##0.00") & "%"
        FEIndicador.TextMatrix(lnFila, 3) = Format(lnInterno, "##,##0.00") & "%"
    Else
        FEIndicador.TextMatrix(lnFila, 2) = fsSBS
        FEIndicador.TextMatrix(lnFila, 3) = fsInterno
    End If
End Sub

Private Sub FEIndicador_OnCellChange(pnRow As Long, pnCol As Long)
    modificarIndConCli pnRow, pnCol
End Sub

Private Sub Form_Load()
    devolverIndConCli
End Sub


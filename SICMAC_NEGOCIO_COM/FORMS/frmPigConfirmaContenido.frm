VERSION 5.00
Begin VB.Form frmPigConfirmaContenido 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirmación de Contenido"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   Icon            =   "frmPigConfirmaContenido.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRegSobrante 
      Caption         =   "&Sobrantes/Faltantes"
      Height          =   375
      Left            =   6810
      TabIndex        =   18
      Top             =   8535
      Width           =   1710
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8640
      TabIndex        =   13
      Top             =   8535
      Width           =   1200
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "&Confirmar"
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      Top             =   8550
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Height          =   8040
      Left            =   45
      TabIndex        =   2
      Top             =   360
      Width           =   9900
      Begin VB.TextBox txtItemsS 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   8805
         TabIndex        =   17
         Top             =   7665
         Width           =   780
      End
      Begin VB.TextBox txtItemsF 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   8805
         TabIndex        =   16
         Top             =   7395
         Width           =   780
      End
      Begin VB.TextBox txtItemsC 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   8805
         TabIndex        =   15
         Top             =   7125
         Width           =   780
      End
      Begin VB.TextBox txtItemsV 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   8805
         TabIndex        =   14
         Top             =   6840
         Width           =   780
      End
      Begin VB.CommandButton CmdConfirmaTodos 
         Caption         =   "Seleccionar Todos"
         Height          =   330
         Left            =   420
         TabIndex        =   10
         Top             =   7245
         Width           =   1575
      End
      Begin SICMACT.FlexEdit feItems 
         Height          =   4170
         Left            =   75
         TabIndex        =   3
         Top             =   2625
         Width           =   9750
         _ExtentX        =   17198
         _ExtentY        =   7355
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-C-NumCta-Item-PesoNeto-NumPiezas-Descripcion"
         EncabezadosAnchos=   "0-400-2000-600-1200-1200-4250"
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
         ColumnasAEditar =   "X-1-X-X-X-X-X"
         ListaControles  =   "0-4-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit feGuias 
         Height          =   2040
         Left            =   105
         TabIndex        =   11
         Top             =   360
         Width           =   9705
         _ExtentX        =   17119
         _ExtentY        =   3598
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-C-Nro Guia-Origen-Motivo-Cant.-Clase"
         EncabezadosAnchos=   "0-400-1500-2600-3500-700-900"
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
         ColumnasAEditar =   "X-1-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-4-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-L-L-R-L"
         FormatosEdit    =   "0-0-0-1-1-3-1"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label7 
         Caption         =   "Items faltantes"
         Height          =   240
         Left            =   7245
         TabIndex        =   9
         Top             =   7410
         Width           =   1425
      End
      Begin VB.Label Label6 
         Caption         =   "Total en Valija"
         Height          =   240
         Left            =   7230
         TabIndex        =   8
         Top             =   6870
         Width           =   1440
      End
      Begin VB.Label Label5 
         Caption         =   "Items sobrantes"
         Height          =   240
         Left            =   7245
         TabIndex        =   7
         Top             =   7695
         Width           =   1425
      End
      Begin VB.Label Label4 
         Caption         =   "Items confirmados "
         Height          =   240
         Left            =   7245
         TabIndex        =   6
         Top             =   7140
         Width           =   1425
      End
      Begin VB.Label Label3 
         Caption         =   "Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   135
         TabIndex        =   5
         Top             =   2415
         Width           =   705
      End
      Begin VB.Label Label2 
         Caption         =   "Guias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   105
         TabIndex        =   4
         Top             =   150
         Width           =   675
      End
   End
   Begin VB.Label lblDestino 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DESTINO DE LA PIEZA"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3705
      TabIndex        =   1
      Top             =   90
      Width           =   3585
   End
   Begin VB.Label Label1 
      Caption         =   "Destino"
      Height          =   270
      Left            =   3030
      TabIndex        =   0
      Top             =   120
      Width           =   630
   End
End
Attribute VB_Name = "frmPigConfirmaContenido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConfirmar_Click()
Dim oPigGrabar As DPigActualizaBD
Dim i As Integer
Dim lsNumDoc As String

Set oPigGrabar = New DPigActualizaBD

    oPigGrabar.dBeginTrans
    For i = 1 To feItems.Rows - 1
    
        lsNumDoc = feItems.TextMatrix(i, 2)
        Call oPigGrabar.dUpdateColocPigGuiaDet(lsNumDoc, , , , , 1, False)
        
    Next i
    oPigGrabar.dCommitTrans

End Sub

Private Sub CmdConfirmaTodos_Click()
Dim i As Integer

    For i = 1 To feItems.Rows - 1
        feItems.TextMatrix(i, 1) = "1"
    Next i
    
    txtItemsC = feItems.Rows - 1
    txtItemsF = CInt(txtItemsV) - CInt(txtItemsC)
        
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub feGuias_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
Dim oPigContrato As DPigContrato
Dim lsNumDoc As String
Dim rs As Recordset
    
    If feGuias.TextMatrix(feGuias.Row, 1) = "." Then 'Guia seleccionada para Confirmacion
        
        Set oPigContrato = New DPigContrato
        lsNumDoc = feGuias.TextMatrix(pnRow, 2)
        txtItemsV = feGuias.TextMatrix(pnRow, 5)
        
        If feGuias.TextMatrix(feGuias.Row, 6) = "LOTE" Then
        
            Set rs = oPigContrato.dObtieneColocPigGuiaDetLotes(lsNumDoc)
    
            Do While Not rs.EOF
                
                feItems.AdicionaFila
                feItems.TextMatrix(feItems.Rows - 1, 2) = rs!cCtaCod
                feItems.TextMatrix(feItems.Rows - 1, 3) = rs!nItemPieza
                feItems.TextMatrix(feItems.Rows - 1, 4) = rs!PNeto
                feItems.TextMatrix(feItems.Rows - 1, 5) = rs!nPiezas
                feItems.TextMatrix(feItems.Rows - 1, 6) = ""
                
                rs.MoveNext
            Loop
                  
        ElseIf feGuias.TextMatrix(feGuias.Row, 6) = "PIEZAS" Then
            
            Set rs = oPigContrato.dObtieneColocPigGuiaDetPiezas(lsNumDoc)
            
            Do While Not rs.EOF
            
                feItems.AdicionaFila
                feItems.TextMatrix(feItems.Rows - 1, 2) = rs!cCtaCod
                feItems.TextMatrix(feItems.Rows - 1, 3) = rs!nItemPieza
                feItems.TextMatrix(feItems.Rows - 1, 4) = rs!PNeto
                feItems.TextMatrix(feItems.Rows - 1, 5) = 1
                feItems.TextMatrix(feItems.Rows - 1, 6) = rs!cDescripcion
            
                rs.MoveNext
            Loop
            
        End If
        Set oPigContrato = Nothing
    End If

End Sub

Private Sub feGuias_RowColChange()
Dim oPigContrato As DPigContrato
Dim lsNumDoc As String
Dim rs As Recordset
    
    If feGuias.TextMatrix(feGuias.Row, 1) = "." Then 'Guia seleccionada para Confirmacion
        
        Set oPigContrato = New DPigContrato
        lsNumDoc = feGuias.TextMatrix(feGuias.Row, 2)
        If feGuias.TextMatrix(feGuias.Row, 6) = "LOTE" Then
            Set rs = oPigContrato.dObtieneColocPigGuiaDetLotes(lsNumDoc)
            
    
            Do While Not rs.EOF
                
                feItems.AdicionaFila
                feItems.TextMatrix(feItems.Rows - 1, 2) = rs!cCtaCod
                feItems.TextMatrix(feItems.Rows - 1, 3) = rs!nItemPieza
                feItems.TextMatrix(feItems.Rows - 1, 4) = rs!PNeto
                feItems.TextMatrix(feItems.Rows - 1, 5) = rs!nPiezas
                feItems.TextMatrix(feItems.Rows - 1, 6) = ""
                
                rs.MoveNext
            Loop
                  
        ElseIf feGuias.TextMatrix(feGuias.Row, 6) = "PIEZAS" Then
            
            Set rs = oPigContrato.dObtieneColocPigGuiaDetPiezas(lsNumDoc)
            
            Do While Not rs.EOF
            
                feItems.AdicionaFila
                feItems.TextMatrix(feItems.Rows - 1, 2) = rs!cCtaCod
                feItems.TextMatrix(feItems.Rows - 1, 3) = rs!nItemPieza
                feItems.TextMatrix(feItems.Rows - 1, 4) = rs!PNeto
                feItems.TextMatrix(feItems.Rows - 1, 5) = 1
                feItems.TextMatrix(feItems.Rows - 1, 6) = rs!cDescripcion
            
                rs.MoveNext
            Loop
            
        End If
        Set oPigContrato = Nothing
    End If

End Sub

Private Sub feItems_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
      
    If feItems.TextMatrix(pnRow, 1) = "." Then
        txtItemsC = Val(txtItemsC) + 1
        txtItemsF = CInt(txtItemsV) - CInt(txtItemsC)
    Else
        If CInt(txtItemsC) > 0 Then txtItemsC = CInt(txtItemsC) - 1
    End If
   
End Sub

Private Sub Form_Load()
Dim oPigCont As DPigContrato
Dim rs As Recordset

lblDestino = gsNomAge

Set oPigCont = New DPigContrato
Set rs = oPigCont.dObtieneGuias(3, , gsCodAge)

Do While Not rs.EOF
    feGuias.AdicionaFila
    feGuias.TextMatrix(feGuias.Row, 2) = rs!cNumDoc
    feGuias.TextMatrix(feGuias.Row, 3) = rs!Origen
    feGuias.TextMatrix(feGuias.Row, 4) = rs!Motivo
    feGuias.TextMatrix(feGuias.Row, 5) = rs!nTotItem
    feGuias.TextMatrix(feGuias.Row, 6) = rs!TipoGuia
    rs.MoveNext
Loop

Set rs = Nothing
Set oPigCont = Nothing

End Sub

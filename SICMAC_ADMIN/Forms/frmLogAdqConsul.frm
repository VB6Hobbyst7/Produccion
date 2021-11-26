VERSION 5.00
Begin VB.Form frmLogAdqConsul 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plan Anual de Adquisiciones y Contrataciones : Consulta"
   ClientHeight    =   6960
   ClientLeft      =   555
   ClientTop       =   1035
   ClientWidth     =   10290
   Icon            =   "frmLogAdqConsul.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraContenedor 
      Caption         =   "Plan de Adquisición "
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
      Height          =   6285
      Left            =   150
      TabIndex        =   2
      Top             =   60
      Width           =   9960
      Begin Sicmact.FlexEdit fgeBS 
         Height          =   3990
         Left            =   165
         TabIndex        =   3
         Top             =   2160
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   7038
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Bien/Servicio-Unidad-Cantidad-Precio-SubTotal"
         EncabezadosAnchos=   "400-0-2200-700-900-1000-1100"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-R-R-R"
         FormatosEdit    =   "0-0-0-0-2-2-2"
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit fgeAdq 
         Height          =   1605
         Left            =   150
         TabIndex        =   4
         Top             =   240
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   2831
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Adquisición-Area-Periodo-Moneda-Estado-Imp"
         EncabezadosAnchos=   "400-750-0-1000-1200-1500-400"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-6"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-4"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-L-C-L-L-C"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         lbOrdenaCol     =   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit fgeBSMes 
         Height          =   3990
         Left            =   6945
         TabIndex        =   5
         Top             =   2160
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   7038
         Cols0           =   4
         HighLight       =   2
         AllowUserResizing=   1
         EncabezadosNombres=   "Mes-Código-Descripción-Cantidad"
         EncabezadosAnchos=   "400-0-1070-1000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0"
         EncabezadosAlineacion=   "R-L-L-R"
         FormatosEdit    =   "0-0-0-2"
         CantEntero      =   6
         CantDecimales   =   1
         AvanceCeldas    =   1
         TextArray0      =   "Mes"
         lbFlexDuplicados=   0   'False
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit fgeMes 
         Height          =   1545
         Left            =   210
         TabIndex        =   8
         Top             =   4575
         Visible         =   0   'False
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   2725
         Cols0           =   13
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Enero-Febrero-Marzo-Abril-Mayo-Junio-Julio-Agosto-Setiembre-Octubre-Noviembre-Diciembre"
         EncabezadosAnchos=   "400-550-550-550-550-550-550-550-550-550-550-550-550"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-L-L-L-L-L-L-L-L"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   165
         X2              =   9765
         Y1              =   1890
         Y2              =   1890
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000E&
         Index           =   1
         X1              =   165
         X2              =   9770
         Y1              =   1905
         Y2              =   1905
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Detalle"
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
         Index           =   3
         Left            =   195
         TabIndex        =   7
         Top             =   1950
         Width           =   900
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Mes"
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
         Index           =   0
         Left            =   7020
         TabIndex        =   6
         Top             =   1935
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdAdq 
      Caption         =   "&Imprimir"
      Height          =   405
      Left            =   5325
      TabIndex        =   1
      Top             =   6450
      Width           =   1260
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   405
      Left            =   8430
      TabIndex        =   0
      Top             =   6450
      Width           =   1260
   End
End
Attribute VB_Name = "frmLogAdqConsul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsDGnral As DLogGeneral

Private Sub cmdAdq_Click()
    Dim clsNImp As NLogImpre
    Dim clsPrevio As clsPrevio
    Dim rs As ADODB.Recordset
    Dim sImpre As String
    Dim nCont As Integer, nSum As Integer

    'Validación de Requerimientos
    Set rs = New ADODB.Recordset
    rs.Fields.Append "cAdqNro", adVarChar, 25, adFldMayBeNull
    rs.Open
    For nCont = 1 To fgeAdq.Rows - 1
        If (fgeAdq.TextMatrix(nCont, 6)) = "." Then
            rs.AddNew "cAdqNro", clsDGnral.GetnMovNro(fgeAdq.TextMatrix(nCont, 1))
            rs.Update
            rs.MoveNext
            nSum = nSum + 1
        End If
    Next
    If nSum = 0 Then
        Set rs = Nothing
        MsgBox "Por favor, determine que Planes se Imprimen", vbInformation, " Aviso "
        Exit Sub
    End If

'''    Set rs = New ADODB.Recordset
'''    rs.Fields.Append "cAdqNro", adVarChar, 25, adFldMayBeNull
'''    rs.Open
'''    rs.AddNew "cAdqNro", fgeAdq.TextMatrix(fgeAdq.Row, 1)
'''    rs.Update

    Set clsNImp = New NLogImpre
    sImpre = clsNImp.ImpReqAdqui(gsNomAge, gdFecSis, rs)
    Set clsNImp = Nothing
    Set rs = Nothing

    Set clsPrevio = New clsPrevio
    clsPrevio.Show sImpre, Me.Caption, True
    Set clsPrevio = Nothing
End Sub

Private Sub fgeAdq_OnRowChange(pnRow As Long, pnCol As Long)
    Dim clsDReq As DLogRequeri
    Dim rs As ADODB.Recordset
    Dim sAdqNro As String
    Dim nCont As Integer
    'Cargar información del Detalle
    Call Limpiar
    sAdqNro = Trim(fgeAdq.TextMatrix(fgeAdq.Row, 1))
    If Trim(sAdqNro) <> "" Then
        Set clsDReq = New DLogRequeri
        Set rs = New ADODB.Recordset
        'Set rs = clsDReq.CargaAdqDetalle(AdqDetUnRegistro, sAdqNro)
        Set rs = clsDReq.CargaReqDetalle(ReqDetUnRegistroConsul, clsDGnral.GetnMovNro(sAdqNro))
        If rs.RecordCount > 0 Then
            Set fgeBS.Recordset = rs
            fgeBS.AdicionaFila
            fgeBS.BackColorRow &HC0FFFF
            If Right(Trim(fgeAdq.TextMatrix(fgeAdq.Row, 4)), 1) = gMonedaNacional Then
                fgeBS.TextMatrix(fgeBS.Row, 2) = "T O T A L    S O L E S"
            Else
                fgeBS.TextMatrix(fgeBS.Row, 2) = "T O T A L    D O L A R E S"
            End If
            fgeBS.TextMatrix(fgeBS.Row, 6) = Format(fgeBS.SumaRow(6), "#,##0.00")
        End If
        Set rs = Nothing
        
        'Cargar información del DetMes
        Set rs = clsDReq.CargaReqDetMes(ReqDetMesUltTraNro, clsDGnral.GetnMovNro(sAdqNro))
        If rs.RecordCount > 0 Then
            Set fgeMes.Recordset = rs
            For nCont = 1 To fgeMes.Rows - 1
                fgeMes.TextMatrix(nCont, 0) = nCont
            Next
        End If
        Set rs = Nothing
        Set clsDReq = Nothing
    End If
    
End Sub

Private Sub fgeBS_OnRowChange(pnRow As Long, pnCol As Long)
    Dim nCont As Integer
    'Carga Meses del Item de acuerdo al Flex fgeMes
    If pnRow <= fgeMes.Rows - 1 Then
        For nCont = 1 To fgeBSMes.Rows - 1
            fgeBSMes.TextMatrix(nCont, 3) = fgeMes.TextMatrix(pnRow, nCont)
        Next
    Else
        For nCont = 1 To fgeBSMes.Rows - 1
            fgeBSMes.TextMatrix(nCont, 3) = ""
        Next
    End If
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Dim clsDReq As DLogRequeri
    'Dim clsDGnral As DLogGeneral
    Set rs = New ADODB.Recordset
    Set clsDReq = New DLogRequeri
    Set clsDGnral = New DLogGeneral
    Call CentraForm(Me)
    
    'Carga Meses
    fgeBSMes.rsFlex = clsDGnral.CargaConstante(gMeses)
    'Set clsDGnral = Nothing
    
    'Carga Plan de Obtención
    Set rs = clsDReq.CargaRequerimiento(gLogReqTipoConsolidado, ReqTodosFlexConsul, "")
    If rs.RecordCount > 0 Then
        fgeAdq.lbEditarFlex = True
        Set fgeAdq.Recordset = rs
        cmdAdq.Enabled = True
        Call fgeAdq_OnRowChange(fgeAdq.Row, fgeAdq.Col)
    Else
        cmdAdq.Enabled = False
    End If
    
    Set clsDReq = Nothing
End Sub


Private Sub cmdSalir_Click()
    Set clsDGnral = Nothing
    Unload Me
End Sub

Private Sub Limpiar()
    Dim nCont As Integer
    'Limpiar FLEX
    fgeBS.Clear
    fgeBS.FormaCabecera
    fgeBS.Rows = 2
    For nCont = 1 To fgeBSMes.Rows - 1
        fgeBSMes.TextMatrix(nCont, 3) = ""
    Next
    fgeMes.Clear
    fgeMes.FormaCabecera
    fgeMes.Rows = 2
End Sub



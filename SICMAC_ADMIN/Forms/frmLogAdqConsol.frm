VERSION 5.00
Begin VB.Form frmLogAdqConsol 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plan Anual de Adquisiciones y Contrataciones : Consolidación"
   ClientHeight    =   7440
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   11700
   Icon            =   "frmLogAdqConsol.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraContenedor 
      Caption         =   "Requerimientos a Consolidar "
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
      Height          =   6870
      Index           =   0
      Left            =   135
      TabIndex        =   8
      Top             =   360
      Width           =   11430
      Begin VB.CommandButton cmdBS 
         Caption         =   "&Consolidar"
         Height          =   390
         Index           =   1
         Left            =   9600
         TabIndex        =   15
         Top             =   4650
         Width           =   1305
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   390
         Left            =   9660
         TabIndex        =   14
         Top             =   6225
         Width           =   1305
      End
      Begin Sicmact.FlexEdit fgeReqDet 
         Height          =   2250
         Left            =   5745
         TabIndex        =   9
         Top             =   405
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   3969
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   1
         EncabezadosNombres=   "Item-Código-Descripción-Unidad-Cantidad-PrecioUni.-Sub Total"
         EncabezadosAnchos=   "400-0-1700-650-800-750-900"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-1-0-0-0-0-0"
         EncabezadosAlineacion=   "R-L-L-L-R-R-R"
         FormatosEdit    =   "0-0-0-0-2-2-2"
         CantEntero      =   6
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit fgeReq 
         Height          =   2430
         Left            =   150
         TabIndex        =   10
         Top             =   225
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   4286
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Requer.-Area-Periodo-Moneda-Tipo Req.-Opc"
         EncabezadosAnchos=   "400-700-1800-550-600-800-350"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         RowHeight0      =   285
      End
      Begin Sicmact.FlexEdit fgeBS 
         Height          =   3795
         Left            =   165
         TabIndex        =   12
         Top             =   3000
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   6694
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-cBSCod-Bien/Servicio-Unidad-Cantidad-PrecioProm-Sub Total"
         EncabezadosAnchos=   "400-0-1800-650-800-800-900"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         RowHeight0      =   285
      End
      Begin Sicmact.FlexEdit fgeBSMes 
         Height          =   3990
         Left            =   6060
         TabIndex        =   13
         Top             =   2805
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   7038
         Cols0           =   4
         HighLight       =   2
         AllowUserResizing=   1
         EncabezadosNombres=   "Mes-Código-Descripción-Cantidad"
         EncabezadosAnchos=   "400-0-1070-1000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Height          =   1425
         Left            =   165
         TabIndex        =   17
         Top             =   5235
         Visible         =   0   'False
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   2514
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
      Begin VB.Label lblEtiqueta 
         Caption         =   "Consolidado"
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
         Left            =   240
         TabIndex        =   16
         Top             =   2760
         Width           =   1530
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000E&
         Index           =   1
         X1              =   165
         X2              =   11275
         Y1              =   2715
         Y2              =   2715
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   165
         X2              =   11275
         Y1              =   2700
         Y2              =   2700
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Detalle de requerimiento seleccionado"
         ForeColor       =   &H8000000D&
         Height          =   210
         Index           =   2
         Left            =   5835
         TabIndex        =   11
         Top             =   180
         Width           =   2850
      End
   End
   Begin VB.ComboBox cboPeriodo 
      Height          =   315
      Left            =   9315
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   45
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.OptionButton optMoneda 
      Caption         =   "Dólares"
      Height          =   195
      Index           =   1
      Left            =   3345
      TabIndex        =   4
      Top             =   120
      Width           =   900
   End
   Begin VB.OptionButton optMoneda 
      Caption         =   "Soles"
      Height          =   195
      Index           =   0
      Left            =   2505
      TabIndex        =   3
      Top             =   120
      Value           =   -1  'True
      Width           =   750
   End
   Begin VB.TextBox txtTipCambio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6795
      TabIndex        =   0
      Top             =   75
      Width           =   945
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   0
      Top             =   -30
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Año :"
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
      Index           =   1
      Left            =   8715
      TabIndex        =   7
      Top             =   90
      Width           =   660
   End
   Begin VB.Label lblPeriodo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   9315
      TabIndex        =   6
      Top             =   60
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Consolidar a  :"
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
      Left            =   870
      TabIndex        =   2
      Top             =   120
      Width           =   1530
   End
   Begin VB.Label lblTipCambio 
      Caption         =   "Tipo Cambio :"
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
      Height          =   255
      Left            =   5550
      TabIndex        =   1
      Top             =   105
      Width           =   1260
   End
End
Attribute VB_Name = "frmLogAdqConsol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim psTpoReq As String
Dim clsDGnral As DLogGeneral

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

Private Sub fgeReq_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    'CHECK
    Dim clsDReq As DLogRequeri
    Dim rs As ADODB.Recordset
    Dim sReqNroAll As String
    Dim nCont As Integer
    If optMoneda(1).Value = True Then
        'DOLARES
        If Val(txtTipCambio.Text) = 0 Then
            'MsgBox "Falta ingresar el tipo de Cambio a utilizar", vbInformation, " Aviso "
            Exit Sub
        End If
    End If
    
    For nCont = 1 To fgeReq.Rows - 1
        If fgeReq.TextMatrix(nCont, 6) = "." Then
            sReqNroAll = sReqNroAll & clsDGnral.GetnMovNro(Trim(fgeReq.TextMatrix(nCont, 1))) & ","
        End If
    Next
    'Carga detalle requerimientos para aprobación CONSOLIDADO
    If Len(Trim(sReqNroAll)) > 0 Then
        sReqNroAll = Left(sReqNroAll, Len(sReqNroAll) - 1)
        Set clsDReq = New DLogRequeri
        Set rs = New ADODB.Recordset
        'Carga Consolidación de Requerimientos
        Set rs = clsDReq.CargaReqDetalle(ReqDetTodosConsol, sReqNroAll, , , IIf(optMoneda(0).Value = True, True, False), Val(txtTipCambio.Text))
        If rs.RecordCount > 0 Then
            Set fgeBS.Recordset = rs
            fgeBS.AdicionaFila
            fgeBS.BackColorRow &HC0FFFF
            fgeBS.TextMatrix(fgeBS.Row, 2) = "T O T A L "
            fgeBS.TextMatrix(fgeBS.Row, 6) = Format(fgeBS.SumaRow(6), "#,##0.00")
        End If
        Set rs = Nothing
        
        'Cargar información del Detalle
        'Set rs = clsDReq.CargaReqDetalle(ReqDetUnRegistroTramiteUlt, sReqNroAll)
        'If rs.RecordCount > 0 Then Set fgeBS.Recordset = rs
        'Set rs = Nothing
        
        'Cargar información del DetMes
        Set rs = clsDReq.CargaReqDetMes(ReqDetMesTodosConsol, sReqNroAll)
        If rs.RecordCount > 0 Then
            Set fgeMes.Recordset = rs
            For nCont = 1 To fgeMes.Rows - 1
                fgeMes.TextMatrix(nCont, 0) = nCont
            Next
        End If
        Set rs = Nothing
        
        Set clsDReq = Nothing
        
        'Actualiza fgeBSDetMes
        Call fgeBS_OnRowChange(fgeBS.Row, fgeBS.Col)
        fgeBS.Row = 1
    Else
        'limpiar
        fgeBS.Clear
        fgeBS.FormaCabecera
        fgeBS.Rows = 2
        fgeMes.Clear
        fgeMes.FormaCabecera
        fgeMes.Rows = 2
        For nCont = 1 To fgeBSMes.Rows - 1
            fgeBSMes.TextMatrix(nCont, 3) = ""
        Next
    End If
End Sub
Private Sub fgeReq_OnRowChange(pnRow As Long, pnCol As Long)
    Dim clsDReq As DLogRequeri
    Dim rs As ADODB.Recordset
    Dim sReqNro As String
    'Cargar información del Detalle
    If optMoneda(1).Value = True Then
        'DOLARES
        If Val(txtTipCambio.Text) = 0 Then
            MsgBox "Falta ingresar el tipo de Cambio a utilizar", vbInformation, " Aviso "
            Exit Sub
        End If
    End If
    sReqNro = Trim(fgeReq.TextMatrix(fgeReq.Row, 1))
    If Trim(sReqNro) <> "" Then
        Set clsDReq = New DLogRequeri
        Set rs = New ADODB.Recordset
        'Set rs = clsDReq.CargaObtDetalle(ObtDetParaAdquisi, sReqNro)
        Set rs = clsDReq.CargaReqDetalle(ReqDetUnRegistroConsol, clsDGnral.GetnMovNro(sReqNro), , , IIf(optMoneda(0).Value = True, True, False), Val(txtTipCambio.Text))
        Set clsDReq = Nothing
        If rs.RecordCount > 0 Then
            Set fgeReqDet.Recordset = rs
            fgeReqDet.AdicionaFila
            fgeReqDet.BackColorRow &HC0FFFF
            fgeReqDet.TextMatrix(fgeReqDet.Row, 2) = "T O T A L "
            fgeReqDet.TextMatrix(fgeReqDet.Row, 6) = Format(fgeReqDet.SumaRow(6), "#,##0.00")
        End If
        Set rs = Nothing
    End If
End Sub

Private Sub Form_Load()
    'Dim clsDGnral As DLogGeneral
    Dim rs As ADODB.Recordset
    Set clsDGnral = New DLogGeneral
    
    Call CentraForm(Me)
    Usuario.Inicio gsCodUser
    If Len(Usuario.AreaCod) = 0 Then
        cmdBS(1).Enabled = False
        MsgBox "Usuario no determinado", vbInformation, "Aviso"
        Exit Sub
    End If
    
    'Para cargar el PERIODO
    'Set clsDGnral = New DLogGeneral
    Set rs = New ADODB.Recordset
    Set rs = clsDGnral.CargaPeriodo
    Call CargaCombo(rs, cboPeriodo)
    cboPeriodo.Visible = True
    'Carga Meses
    fgeBSMes.rsFlex = clsDGnral.CargaConstante(gMeses)
    'Set clsDGnral = Nothing
    
    Call CargaRequeri
End Sub

Private Sub cmdBS_Click(Index As Integer)
    Dim nReqNro As Integer, nAdqNro As Integer
    Dim sAdqNro As String, sReqNro As String, sMoneda As String, sBSCod As String, sActualiza As String
    Dim nPrecio As Currency, nCantidad As Currency
    Dim nCont As Integer, nSum As Integer, nResult As Integer
    Dim nBs As Integer, nBSMes As Integer
    Dim clsDMov As DLogMov
    'Dim clsDGnral As DLogGeneral
    
    Select Case Index
        Case 1:
            'CONSOLIDAR
            'Validación de Requerimientos
            For nCont = 1 To fgeReq.Rows - 1
                If (fgeReq.TextMatrix(nCont, 6)) = "." Then
                    nSum = nSum + 1
                    Exit For
                End If
            Next
            If nSum = 0 Then
                MsgBox "Por favor, determine que Requerimientos se consolidan", vbInformation, " Aviso "
                Exit Sub
            End If
            'Periodo
            If Trim(cboPeriodo.Text) = "" Then
                MsgBox "Determine el periodo de la Consolidación", vbInformation, " Aviso"
                Exit Sub
            End If
            If MsgBox("¿ Estás seguro de consolidar estos Requerimientos ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                'Set clsDGnral = New DLogGeneral
                sAdqNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                'Set clsDGnral = Nothing
                
                sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                Set clsDMov = New DLogMov
                
                'Grabación de MOV
                clsDMov.InsertaMov sAdqNro, Trim(Str(gLogOpeReqRegistro)), "", gLogReqEstadoInicio
                nAdqNro = clsDMov.GetnMovNro(sAdqNro)
                clsDMov.InsertaMovRef nAdqNro, nAdqNro
                
                'Inserta Requerimiento Consolidado
                clsDMov.InsertaRequeri nAdqNro, cboPeriodo.Text, gLogReqTipoConsolidado, _
                     "", ""
                
                clsDMov.InsertaReqTramite nAdqNro, nAdqNro, Usuario.AreaCod, "", _
                     "", gLogReqEstadoInicio, IIf(optMoneda(0).Value = True, gMonedaNacional, gMonedaExtranjera), sActualiza
                
                'Inserta ReqTramite (Consolidado) y ReqCon
                For nCont = 1 To fgeReq.Rows - 1
                    If (fgeReq.TextMatrix(nCont, 6)) = "." Then
                        sReqNro = Trim(fgeReq.TextMatrix(nCont, 1))
                        nReqNro = clsDMov.GetnMovNro(sReqNro)
                        sMoneda = Right(Trim(fgeReq.TextMatrix(nCont, 4)), 1)
                            
                            'sMoneda = "1"
                        'Inserta ReqCon - Requerimientos que se Consolidan
                        clsDMov.InsertaReqCon nReqNro, nAdqNro
                        
                        'Cambiar el estado a los requerimientos consolidados
                        clsDMov.InsertaReqTramite nReqNro, nAdqNro, Usuario.AreaCod, "", _
                            "", gLogReqEstadoConsolida, sMoneda, sActualiza
                        
                        'sReqNroAll = sReqNroAll & "','" & sReqNro
                    End If
                Next
                'sReqNroAll = Mid(sReqNroAll, 4)
                'Inserta el Detalle de los Bienes/Servicios
                nBs = 0: nBSMes = 0
                For nBs = 1 To fgeBS.Rows - 2
                    sBSCod = fgeBS.TextMatrix(nBs, 1)
                    nPrecio = CCur(IIf(fgeBS.TextMatrix(nBs, 5) = "", 0, fgeBS.TextMatrix(nBs, 5)))
                    clsDMov.InsertaReqDetalle nAdqNro, nAdqNro, sBSCod, _
                         nPrecio, sActualiza
                    'Inserta los Meses de los Bienes/Servicios
                    For nBSMes = 1 To fgeMes.Cols - 1
                        nCantidad = CCur(IIf(fgeMes.TextMatrix(nBs, nBSMes) = "", 0, fgeMes.TextMatrix(nBs, nBSMes)))
                        If nCantidad > 0 Then
                            clsDMov.InsertaReqDetMes nAdqNro, nAdqNro, sBSCod, _
                                 Val(fgeBSMes.TextMatrix(nBSMes, 1)), nCantidad
                        End If
                    Next
                Next
                
                'Ejecuta todos los querys en una transacción
                'nResult = clsDMov.EjecutaBatch
                Set clsDMov = Nothing
                
                If nResult = 0 Then
                    cmdBS(1).Enabled = False
                    Call CargaRequeri
                Else
                    MsgBox "Error al grabar la información", vbInformation, " Aviso "
                End If
            End If
        Case Else
            MsgBox "Comando no reconocido", vbInformation, " Aviso"
    End Select
End Sub

Private Sub cmdSalir_Click()
    Set clsDGnral = Nothing
    Unload Me
End Sub

Private Sub Limpiar()
    Dim nCont As Integer
    'Limpiar FLEX
    cboPeriodo.ListIndex = -1
    fgeReq.Clear
    fgeReq.FormaCabecera
    fgeReq.Rows = 2
    fgeReqDet.Clear
    fgeReqDet.FormaCabecera
    fgeReqDet.Rows = 2
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

Private Sub CargaRequeri()
    Dim clsDReq As DLogRequeri
    Dim rs As ADODB.Recordset
    'Actualiza FLEX
    Call Limpiar
    
    'Carga Requerimientos a Consolidar
    Set clsDReq = New DLogRequeri
    Set rs = New ADODB.Recordset
    Set rs = clsDReq.CargaRequerimiento(0, ReqTodosFlexConsol, "")
    Set clsDReq = Nothing
    If rs.RecordCount > 0 Then
        Set fgeReq.Recordset = rs
        fgeReq.lbEditarFlex = True
        cmdBS(1).Enabled = True
        Call fgeReq_OnRowChange(fgeReq.Row, fgeReq.Col)
    Else
        cmdBS(1).Enabled = False
    End If
    Set rs = Nothing
End Sub

Private Sub optMoneda_Click(Index As Integer)
    Call fgeReq_OnRowChange(fgeReq.Row, fgeReq.Col)
    Call fgeReq_OnCellCheck(fgeReq.Row, fgeReq.Col)
End Sub

Private Sub txtTipCambio_GotFocus()
    Dim nCont As Integer
    fgeReqDet.Clear
    fgeReqDet.FormaCabecera
    fgeReqDet.Rows = 2
    fgeBS.Clear
    fgeBS.FormaCabecera
    fgeBS.Rows = 2
    fgeMes.Clear
    fgeMes.FormaCabecera
    fgeMes.Rows = 2
    For nCont = 1 To fgeBSMes.Rows - 1
        fgeBSMes.TextMatrix(nCont, 3) = ""
    Next
End Sub
Private Sub txtTipCambio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtTipCambio, KeyAscii, 8, 4)
    If KeyAscii = 13 Then
        Call optMoneda_Click(IIf(optMoneda(0).Value = True, 0, 1))
    End If
End Sub
Private Sub txtTipCambio_LostFocus()
    Call optMoneda_Click(IIf(optMoneda(0).Value = True, 0, 1))
End Sub

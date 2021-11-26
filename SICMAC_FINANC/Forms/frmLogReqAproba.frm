VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmLogReqAproba 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6750
   ClientLeft      =   630
   ClientTop       =   1260
   ClientWidth     =   11070
   Icon            =   "frmLogReqAproba.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   Begin Sicmact.FlexEdit fgeReqDet 
      Height          =   2550
      Left            =   165
      TabIndex        =   5
      Top             =   3510
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   4498
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   1
      EncabezadosNombres=   "Item-Código-Descripción-Unidad-Cantidad-PrecioUni.-Sub Total"
      EncabezadosAnchos=   "400-0-1700-650-800-750-900"
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
      RowHeight0      =   285
   End
   Begin Sicmact.FlexEdit fgeMes 
      Height          =   1290
      Left            =   165
      TabIndex        =   16
      Top             =   4785
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   2275
      Cols0           =   13
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-Enero-Febrero-Marzo-Abril-Mayo-Junio-Julio-Agosto-Setiembre-Octubre-Noviembre-Diciembre"
      EncabezadosAnchos=   "350-400-400-400-400-400-400-400-400-400-400-400-400"
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
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L-L-L-L-L-L-L-L-L-L"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      Appearance      =   0
      ColWidth0       =   345
      RowHeight0      =   285
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Aprobar"
      Height          =   390
      Index           =   2
      Left            =   6675
      TabIndex        =   15
      Top             =   6165
      Width           =   1305
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   9315
      TabIndex        =   10
      Top             =   6165
      Width           =   1305
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Rechazar"
      Height          =   390
      Index           =   1
      Left            =   4620
      TabIndex        =   9
      Top             =   6165
      Width           =   1305
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Cancelar"
      Height          =   390
      Index           =   0
      Left            =   2355
      TabIndex        =   8
      Top             =   6150
      Width           =   1305
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   210
      Top             =   6135
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin Spinner.uSpinner spinPeriodo 
      Height          =   300
      Left            =   9105
      TabIndex        =   2
      Top             =   120
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   529
      Max             =   9999
      Min             =   2000
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
   End
   Begin Sicmact.FlexEdit fgeReq 
      Height          =   2550
      Left            =   165
      TabIndex        =   4
      Top             =   645
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   4498
      Cols0           =   6
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-cAreaCod-Area-Requerimiento-Necesidad-Requerimiento"
      EncabezadosAnchos=   "400-0-2700-2000-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L-L-L"
      FormatosEdit    =   "0-0-0-0-0-0"
      TextArray0      =   "Item"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   285
   End
   Begin RichTextLib.RichTextBox rtfDescri 
      Height          =   2580
      Index           =   0
      Left            =   5760
      TabIndex        =   11
      Top             =   615
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4551
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MaxLength       =   8000
      TextRTF         =   $"frmLogReqAproba.frx":030A
   End
   Begin RichTextLib.RichTextBox rtfDescri 
      Height          =   2580
      Index           =   1
      Left            =   5760
      TabIndex        =   12
      Top             =   3480
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4551
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MaxLength       =   8000
      TextRTF         =   $"frmLogReqAproba.frx":038C
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Necesidad"
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
      Index           =   5
      Left            =   6000
      TabIndex        =   14
      Top             =   405
      Width           =   1080
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Requerimiento"
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
      Index           =   2
      Left            =   6000
      TabIndex        =   13
      Top             =   3270
      Width           =   1425
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Requerimientos :"
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
      Left            =   315
      TabIndex        =   7
      Top             =   420
      Width           =   1515
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Detalle :"
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
      Index           =   4
      Left            =   345
      TabIndex        =   6
      Top             =   3255
      Width           =   810
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
      Left            =   8535
      TabIndex        =   3
      Top             =   180
      Width           =   660
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Area :"
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
      Left            =   390
      TabIndex        =   1
      Top             =   105
      Width           =   750
   End
   Begin VB.Label lblAreaDes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1155
      TabIndex        =   0
      Top             =   75
      Width           =   4110
   End
End
Attribute VB_Name = "frmLogReqAproba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim psTpoReq As String
Dim clsDGnral As DLogGeneral
Dim clsDReq As DLogRequeri

Public Sub Inicio(ByVal psTipoReq As String)
psTpoReq = psTipoReq

If psTpoReq = "1" Then
    Me.Caption = "Aprobación de Proyección de Requerimiento"
Else
    Me.Caption = "Aprobación de Requerimiento Extemporaneo"
End If
Me.Show 1
End Sub

Private Sub cmdReq_Click(Index As Integer)
    Dim sReqNro As String, sReqTraNro As String, sBSCod As String, sCtaCont As String, sActualiza As String
    Dim nPrecio As Currency, nCantidad As Currency
    Dim nCont As Integer, nResult As Integer
    Dim nBs As Integer, nBSMes As Integer
    Dim clsDMov As DLogMov
    Dim clsDGnral As DLogGeneral
    
    Select Case Index
        Case 0:
            'Cancelar
            If MsgBox("¿ Estás seguro de cancelar toda la operación ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Call cmdSalir_Click
            End If
        Case 1:
            'Rechazar
            If MsgBox("¿ Estás seguro de Rechazar el requerimiento ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Set clsDGnral = New DLogGeneral
                sReqNro = Trim(fgeReq.TextMatrix(fgeReq.Row, 3))
                sReqTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                Set clsDGnral = Nothing
                
                Set clsDMov = New DLogMov
                'Inserta Mov
                clsDMov.InsertaMov sReqTraNro, Trim(Str(gLogOpeReqTramite)), "", Trim(Str(gLogReqEstadoRechazado))
                clsDMov.InsertaMovRef sReqTraNro, sReqNro
                
                'Inserta tramite
                clsDMov.InsertaReqTramite sReqNro, sReqTraNro, Usuario.AreaCod, "", _
                    "", gLogReqEstadoRechazado, gLogReqFlujoSin, sActualiza
                
                'Inserta LogReqDetalle
                nBs = 0: nBSMes = 0
                For nBs = 1 To fgeReqDet.Rows - 2
                    sBSCod = fgeReqDet.TextMatrix(nBs, 1)
                    nPrecio = CCur(IIf(fgeReqDet.TextMatrix(nBs, 5) = "", 0, fgeReqDet.TextMatrix(nBs, 5)))
                    clsDMov.InsertaReqDetalle sReqNro, sReqTraNro, sBSCod, _
                        "", 0, nPrecio, "", sActualiza
                    For nBSMes = 1 To fgeMes.Cols - 1
                        nCantidad = CCur(IIf(fgeMes.TextMatrix(nBs, nBSMes) = "", 0, fgeMes.TextMatrix(nBs, nBSMes)))
                        If nCantidad > 0 Then
                            clsDMov.InsertaReqDetMes sReqNro, sReqTraNro, sBSCod, _
                                 nBSMes, nCantidad
                        End If
                    Next
                Next
                'Ejecuta todos los querys en una transacción
                'nResult = clsDMov.EjecutaBatch
                Set clsDMov = Nothing
                
                If nResult = 0 Then
                    cmdReq(0).Enabled = False
                    cmdReq(1).Enabled = False
                    cmdReq(2).Enabled = False
                    Call spinPeriodo_Change
                Else
                    MsgBox "Error al rechazar el requerimiento", vbInformation, " Aviso "
                End If
            End If
        Case 2:
            'Aprobar
            If MsgBox("¿ Estás seguro de Aprobar este Requerimiento ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Set clsDGnral = New DLogGeneral
                sReqNro = Trim(fgeReq.TextMatrix(fgeReq.Row, 3))
                sReqTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                Set clsDGnral = Nothing
                
                sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                Set clsDMov = New DLogMov
                
                'Grabación de MOV - MOVREF
                clsDMov.InsertaMov sReqTraNro, Trim(Str(gLogOpeReqTramite)), "", Trim(Str(gLogReqEstadoAceptado))
                clsDMov.InsertaMovRef sReqTraNro, sReqNro
                
                'Inserta LogReqTramite
                clsDMov.InsertaReqTramite sReqNro, sReqTraNro, Usuario.AreaCod, "", _
                    "", gLogReqEstadoAceptado, gLogReqFlujoSin, sActualiza

                'Inserta LogReqDetalle
                nBs = 0: nBSMes = 0
                For nBs = 1 To fgeReqDet.Rows - 2
                    sBSCod = fgeReqDet.TextMatrix(nBs, 1)
                    nPrecio = CCur(IIf(fgeReqDet.TextMatrix(nBs, 5) = "", 0, fgeReqDet.TextMatrix(nBs, 5)))
                    clsDMov.InsertaReqDetalle sReqNro, sReqTraNro, sBSCod, _
                        "", 0, nPrecio, "", sActualiza
                    For nBSMes = 1 To fgeMes.Cols - 1
                        nCantidad = CCur(IIf(fgeMes.TextMatrix(nBs, nBSMes) = "", 0, fgeMes.TextMatrix(nBs, nBSMes)))
                        If nCantidad > 0 Then
                            clsDMov.InsertaReqDetMes sReqNro, sReqTraNro, sBSCod, _
                                 nBSMes, nCantidad
                        End If
                    Next
                Next
                
                'Ejecuta todos los querys en una transacción
                'nResult = clsDMov.EjecutaBatch
                Set clsDMov = Nothing
                
                If nResult = 0 Then
                    cmdReq(0).Enabled = False
                    cmdReq(1).Enabled = False
                    cmdReq(2).Enabled = False
                    Call spinPeriodo_Change
                Else
                    MsgBox "Error al grabar la información", vbInformation, " Aviso "
                End If
            End If
        Case Else
            MsgBox "Opción no reconocida", vbInformation, "Aviso"
    End Select
End Sub

Private Sub cmdSalir_Click()
    Set clsDGnral = Nothing
    Set clsDReq = Nothing
    Unload Me
End Sub


Private Sub fgeReq_OnRowChange(pnRow As Long, pnCol As Long)
    Dim rs As ADODB.Recordset
    Dim sReqNro As String
    'Cargar información del Detalle
    If Trim(fgeReq.TextMatrix(fgeReq.Row, 1)) <> "" Then
        sReqNro = fgeReq.TextMatrix(fgeReq.Row, 3)
        If Trim(sReqNro) <> "" Then
            'Carga Textos
            rtfDescri(0).Text = fgeReq.TextMatrix(fgeReq.Row, 4)
            rtfDescri(1).Text = fgeReq.TextMatrix(fgeReq.Row, 5)
            'Carga Flex de Detalle
            Set rs = New ADODB.Recordset
            Set rs = clsDReq.CargaReqDetalle(ReqDetUnRegistroConsol, sReqNro)
            If rs.RecordCount > 0 Then
                Set fgeReqDet.Recordset = rs
                fgeReqDet.AdicionaFila
                fgeReqDet.BackColorRow &HC0FFFF
                fgeReqDet.TextMatrix(fgeReqDet.Row, 2) = "T O T A L "
                fgeReqDet.TextMatrix(fgeReqDet.Row, 6) = Format(fgeReqDet.SumaRow(6), "#,##0.00")
            End If
            Set rs = Nothing
            'Cargar información del DetMes
            Set rs = clsDReq.CargaReqDetMes(sReqNro, "")
            If rs.RecordCount > 0 Then Set fgeMes.Recordset = rs
            Set rs = Nothing
        End If
    End If
End Sub

Private Sub Form_Load()
    Set clsDGnral = New DLogGeneral
    Set clsDReq = New DLogRequeri
    Call CentraForm(Me)
    'Carga información de la relación usuario-area
    Usuario.Inicio gsCodUser
    If Len(Usuario.AreaCod) = 0 Then
        MsgBox "Usuario no determinado", vbInformation, "Aviso"
        Exit Sub
    End If
    lblAreaDes.Caption = Usuario.AreaNom
    'Inicia Periodo
    spinPeriodo.Valor = Year(gdFecSis) + 1
End Sub

Private Sub Limpiar()
    'Limpiar FLEX
    fgeReq.Clear
    fgeReq.FormaCabecera
    fgeReq.Rows = 2
    fgeReqDet.Clear
    fgeReqDet.FormaCabecera
    fgeReqDet.Rows = 2
    fgeMes.Clear
    fgeMes.FormaCabecera
    fgeMes.Rows = 2
    rtfDescri(0).Text = ""
    rtfDescri(1).Text = ""
End Sub

Private Sub spinPeriodo_Change()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'Actualiza FLEX
    Limpiar
    'Carga requerimientos para aprobación
    Set rs = clsDReq.CargaRequerimiento(psTpoReq, ReqTodosFlexApro, "", "", spinPeriodo.Valor)
    If rs.RecordCount > 0 Then
        Set fgeReq.Recordset = rs
        cmdReq(0).Enabled = True
        cmdReq(1).Enabled = True
        cmdReq(2).Enabled = True
        Call fgeReq_OnRowChange(fgeReq.Row, fgeReq.Col)
    Else
        cmdReq(0).Enabled = False
        cmdReq(1).Enabled = False
        cmdReq(2).Enabled = False
    End If
End Sub


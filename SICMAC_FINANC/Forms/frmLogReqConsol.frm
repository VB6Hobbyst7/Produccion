VERSION 5.00
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmLogReqConsol 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consolidación de ..."
   ClientHeight    =   6600
   ClientLeft      =   150
   ClientTop       =   1140
   ClientWidth     =   11685
   Icon            =   "frmLogReqConsol.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Cancelar"
      Height          =   390
      Index           =   0
      Left            =   3660
      TabIndex        =   5
      Top             =   6090
      Width           =   1305
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Grabar"
      Height          =   390
      Index           =   1
      Left            =   5925
      TabIndex        =   4
      Top             =   6090
      Width           =   1305
   End
   Begin Sicmact.FlexEdit fgeReq 
      Height          =   2580
      Left            =   135
      TabIndex        =   3
      Top             =   690
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   4551
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-cAreaCod-Area-Requerimiento-Ok"
      EncabezadosAnchos=   "400-0-2200-2100-400"
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
      ColumnasAEditar =   "X-X-X-X-4"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-4"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L-C"
      FormatosEdit    =   "0-0-0-0-0"
      TextArray0      =   "Item"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   285
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   8865
      TabIndex        =   0
      Top             =   6090
      Width           =   1305
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   45
      Top             =   6150
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin Sicmact.FlexEdit fgeBS 
      Height          =   5265
      Left            =   5760
      TabIndex        =   6
      Top             =   690
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   9287
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-cBSCod-Bien/Servicio-Unidad-Cantidad-PrecioProm-Sub Total"
      EncabezadosAnchos=   "400-0-1800-650-800-800-900"
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
      ColumnasAEditar =   "X-X-X-X-X-5-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-L-L-L-R-R-R"
      FormatosEdit    =   "0-0-0-0-2-2-2"
      AvanceCeldas    =   1
      TextArray0      =   "Item"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   285
   End
   Begin Spinner.uSpinner spinPeriodo 
      Height          =   300
      Left            =   9915
      TabIndex        =   7
      Top             =   105
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
   Begin Sicmact.FlexEdit fgeReqDet 
      Height          =   2415
      Left            =   135
      TabIndex        =   12
      Top             =   3540
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   4260
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
      Left            =   285
      TabIndex        =   11
      Top             =   3315
      Width           =   810
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Detalle de Requerimientos Consolidado :"
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
      Left            =   5805
      TabIndex        =   10
      Top             =   450
      Width           =   3525
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
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   450
      Width           =   1515
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
      Left            =   9345
      TabIndex        =   8
      Top             =   165
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
      Left            =   405
      TabIndex        =   2
      Top             =   105
      Width           =   750
   End
   Begin VB.Label lblAreaDes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1170
      TabIndex        =   1
      Top             =   75
      Width           =   4320
   End
End
Attribute VB_Name = "frmLogReqConsol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim psTpoReq As String
Dim clsDReq As DLogRequeri

Public Sub Inicio(ByVal psTipoReq As String)
psTpoReq = psTipoReq
If psTpoReq = "1" Then
    Me.Caption = "Consolidación de Proyección de Requerimientos"
Else
    Me.Caption = "Consolidación de Requerimiento Extemporaneo"
End If
Me.Show 1
End Sub

Private Sub fgeReq_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    'CHECK
    Dim rs As ADODB.Recordset
    Dim sReqNro As String
    Dim nCont As Integer
    For nCont = 1 To fgeReq.Rows - 1
        If fgeReq.TextMatrix(nCont, 4) = "." Then
            sReqNro = sReqNro & Trim(fgeReq.TextMatrix(nCont, 3)) & "','"
        End If
    Next
    'Carga detalle requerimientos para aprobación CONSOLIDADO
    If Len(Trim(sReqNro)) > 0 Then
        sReqNro = Left(sReqNro, Len(sReqNro) - 3)
        Set rs = New ADODB.Recordset
        Set rs = clsDReq.CargaReqDetalle(ReqDetTodosFlex, sReqNro)
        If rs.RecordCount > 0 Then
            Set fgeBS.Recordset = rs
            fgeBS.AdicionaFila
            fgeBS.BackColorRow &HC0FFFF
            fgeBS.TextMatrix(fgeBS.Row, 2) = "T O T A L "
            fgeBS.TextMatrix(fgeBS.Row, 6) = Format(fgeBS.SumaRow(6), "#,##0.00")
        End If
        Set rs = Nothing
    Else
        fgeBS.Clear
        fgeBS.FormaCabecera
        fgeBS.Rows = 2
    End If
End Sub

Private Sub fgeReq_OnRowChange(pnRow As Long, pnCol As Long)
    Dim rs As ADODB.Recordset
    Dim sReqNro As String
    'Cargar información del Detalle
    If Trim(fgeReq.TextMatrix(fgeReq.Row, 1)) <> "" Then
        sReqNro = fgeReq.TextMatrix(fgeReq.Row, 3)
        If Trim(sReqNro) <> "" Then
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
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Set clsDReq = New DLogRequeri
    Set rs = New ADODB.Recordset
    Call CentraForm(Me)
    'Carga información de la relación usuario-area
    Usuario.Inicio gsCodUser
    If Len(Usuario.AreaCod) = 0 Then
        MsgBox "Usuario no determinado", vbInformation, "Aviso"
        Exit Sub
    End If
    lblAreaDes.Caption = Usuario.AreaNom
    
    'Inicia Periodo
    spinPeriodo.Valor = IIf(psTpoReq = "1", Year(gdFecSis) + 1, Year(gdFecSis))
End Sub

Private Sub cmdReq_Click(Index As Integer)
    Dim sObtNro As String, sReqNro As String, sBSCod As String, sCtaCont As String, sActualiza As String
    Dim nPrecio As Currency, nCantidad As Currency
    Dim nCont As Integer, nSum As Integer, nResult As Integer
    Dim clsDMov As DLogMov
    Dim clsDGnral As DLogGeneral
    
    Select Case Index
        Case 0:
            'Cancelar
            If MsgBox("¿ Estás seguro de cancelar toda la operación ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Call cmdSalir_Click
            End If
        Case 1:
            'Validación de Requerimientos
            For nCont = 1 To fgeReq.Rows - 1
                If (fgeReq.TextMatrix(nCont, 4)) = "." Then
                    nSum = nSum + 1
                End If
            Next
            If nSum = 0 Then
                MsgBox "Por favor, determine que requerimiento(s) se consolidan", vbInformation, " Aviso "
                Exit Sub
            End If
            'If psTpoReq = "2" Then
            '    For nCont = 1 To fgeBS.Rows - 1
            '        If CCur(IIf(Trim(fgeBS.TextMatrix(nCont, 5)) = "", 0, fgeBS.TextMatrix(nCont, 5))) = 0 Then
            '            MsgBox "Por favor, ingrese los precios de los requerimientos", vbInformation, " Aviso "
            '            Exit Sub
            '        End If
            '    Next
            'End If
        
            If MsgBox("¿ Estás seguro de consolidar en un solo Plan ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Set clsDGnral = New DLogGeneral
                sObtNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                Set clsDGnral = Nothing
                
                sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                Set clsDMov = New DLogMov
                
                'Grabación de MOV
                clsDMov.InsertaMov sObtNro, Trim(Str(gLogOpeObtRegistro)), "", Trim(Str(gLogObtEstadoInicio))
                clsDMov.InsertaMovRef sObtNro, sObtNro
                
                'Inserta LogObtencion
                clsDMov.InsertaObtencion psTpoReq, sObtNro, spinPeriodo.Valor, sActualiza
                
                'Inserta LogReqObt
                For nCont = 1 To fgeReq.Rows - 1
                    If (fgeReq.TextMatrix(nCont, 4)) = "." Then
                        sReqNro = Trim(fgeReq.TextMatrix(nCont, 3))
                        clsDMov.InsertaReqObt sReqNro, sObtNro
                        'If psTpoReq = "1" Then
                            clsDMov.InsertaReqTramite sReqNro, sObtNro, Usuario.AreaCod, "", _
                                "", gLogReqEstadoConsolida, gLogReqFlujoSin, sActualiza
                        'Else
                        '    clsDMov.InsertaReqTramite sReqNro, sObtNro, Usuario.AreaCod, "", _
                        '        "", gLogReqEstadoPrecio, gLogReqFlujoSin, sActualiza
                        'End If
                    End If
                Next
                'Inserta LogObtDetalle
                For nCont = 1 To fgeBS.Rows - 2
                    sBSCod = fgeBS.TextMatrix(nCont, 1)
                    nCantidad = CCur(IIf(Trim(fgeBS.TextMatrix(nCont, 4)) = "", 0, fgeBS.TextMatrix(nCont, 4)))
                    nPrecio = CCur(IIf(Trim(fgeBS.TextMatrix(nCont, 5)) = "", 0, fgeBS.TextMatrix(nCont, 5)))
                    sCtaCont = ""
                    clsDMov.InsertaObtDetalle sObtNro, sBSCod, _
                        nCantidad, nPrecio, sCtaCont, sActualiza
                Next
                
                'Ejecuta todos los querys en una transacción
                'nResult = clsDMov.EjecutaBatch
                Set clsDMov = Nothing
                
                If nResult = 0 Then
                    cmdReq(0).Enabled = False
                    cmdReq(1).Enabled = False
                    fgeReq.lbEditarFlex = False
                    Call spinPeriodo_Change
                Else
                    MsgBox "Error al grabar la información", vbInformation, " Aviso "
                End If
            End If
        Case Else
            MsgBox "Comando no reconocido", vbInformation, " Aviso"
    End Select
End Sub

Private Sub cmdSalir_Click()
    Set clsDReq = Nothing
    Unload Me
End Sub


Private Sub Limpiar()
    'Limpiar FLEX
    fgeReq.Clear
    fgeReq.FormaCabecera
    fgeReq.Rows = 2
    fgeBS.Clear
    fgeBS.FormaCabecera
    fgeBS.Rows = 2
    fgeReqDet.Clear
    fgeReqDet.FormaCabecera
    fgeReqDet.Rows = 2
End Sub

Private Sub spinPeriodo_Change()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'Actualiza FLEX
    Limpiar
    'Carga requerimientos para aprobación
    Set rs = clsDReq.CargaRequerimiento(psTpoReq, ReqTodosFlexConsol, "", "", spinPeriodo.Valor)
    If rs.RecordCount > 0 Then
        Set fgeReq.Recordset = rs
        Call fgeReq_OnRowChange(fgeReq.Row, fgeReq.Col)
        cmdReq(0).Enabled = True
        cmdReq(1).Enabled = True
    Else
        cmdReq(0).Enabled = False
        cmdReq(1).Enabled = False
    End If
End Sub

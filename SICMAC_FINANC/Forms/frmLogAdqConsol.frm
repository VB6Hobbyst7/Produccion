VERSION 5.00
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmLogAdqConsol 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plan Anual de Adquisición"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   1905
   ClientWidth     =   11580
   Icon            =   "frmLogAdqConsol.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optPlan 
      Caption         =   "Extemporaneo"
      Height          =   330
      Index           =   1
      Left            =   7500
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   30
      Width           =   1305
   End
   Begin VB.OptionButton optPlan 
      Caption         =   "Normal"
      Height          =   330
      Index           =   0
      Left            =   6180
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   30
      Value           =   -1  'True
      Width           =   1305
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   8805
      TabIndex        =   2
      Top             =   6060
      Width           =   1305
   End
   Begin VB.CommandButton cmdBS 
      Caption         =   "&Grabar"
      Height          =   390
      Index           =   1
      Left            =   5865
      TabIndex        =   1
      Top             =   6060
      Width           =   1305
   End
   Begin VB.CommandButton cmdBS 
      Caption         =   "&Cancelar"
      Height          =   390
      Index           =   0
      Left            =   3600
      TabIndex        =   0
      Top             =   6060
      Width           =   1305
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   0
      Top             =   6105
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin Sicmact.FlexEdit fgeBS 
      Height          =   5265
      Left            =   5760
      TabIndex        =   3
      Top             =   660
      Width           =   5685
      _ExtentX        =   10028
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
   Begin Spinner.uSpinner spinPeriodo 
      Height          =   300
      Left            =   9855
      TabIndex        =   4
      Top             =   75
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
   Begin Sicmact.FlexEdit fgeObtDet 
      Height          =   2415
      Left            =   135
      TabIndex        =   5
      Top             =   3510
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
   Begin Sicmact.FlexEdit fgeObt 
      Height          =   2505
      Left            =   135
      TabIndex        =   14
      Top             =   660
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   4419
      Cols0           =   3
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-Obtención-Ok"
      EncabezadosAnchos=   "400-3200-400"
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
      ColumnasAEditar =   "X-X-2"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-4"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "L-L-C"
      FormatosEdit    =   "0-0-0"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   285
   End
   Begin VB.Label lblAreaDes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1110
      TabIndex        =   11
      Top             =   45
      Width           =   4320
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
      Left            =   345
      TabIndex        =   10
      Top             =   75
      Width           =   750
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
      Left            =   9285
      TabIndex        =   9
      Top             =   135
      Width           =   660
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Planes de obtención :"
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
      TabIndex        =   8
      Top             =   420
      Width           =   2010
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Detalle de planes consolidados :"
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
      Left            =   5865
      TabIndex        =   7
      Top             =   435
      Width           =   3000
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
      Left            =   225
      TabIndex        =   6
      Top             =   3285
      Width           =   810
   End
End
Attribute VB_Name = "frmLogAdqConsol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim psTpoReq As String
Dim clsDReq As DLogRequeri

Private Sub fgeObt_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    'CHECK
    Dim rs As ADODB.Recordset
    Dim sObtNro As String
    Dim nCont As Integer
    For nCont = 1 To fgeObt.Rows - 1
        If fgeObt.TextMatrix(nCont, 2) = "." Then
            sObtNro = sObtNro & Trim(fgeObt.TextMatrix(nCont, 1)) & "','"
        End If
    Next
    'Carga detalle requerimientos para aprobación CONSOLIDADO
    If Len(Trim(sObtNro)) > 0 Then
        sObtNro = Left(sObtNro, Len(sObtNro) - 3)
        Set rs = New ADODB.Recordset
        'Carga Consolidación de Plan Obtención
        Set rs = clsDReq.CargaObtDetalle(ObtDetParaAdquisiConsol, sObtNro)
        If rs.RecordCount > 0 Then
            Set fgeBS.Recordset = rs
            fgeBS.AdicionaFila
            fgeBS.BackColorRow &HC0FFFF
            fgeBS.TextMatrix(fgeBS.Row, 2) = "T O T A L "
            fgeBS.TextMatrix(fgeBS.Row, 6) = Format(fgeBS.SumaRow(6), "#,##0.00")
        End If
        Set rs = Nothing
    Else
        'limpiar
        fgeBS.Clear
        fgeBS.FormaCabecera
        fgeBS.Rows = 2
    End If
End Sub
Private Sub fgeObt_OnRowChange(pnRow As Long, pnCol As Long)
    Dim rs As ADODB.Recordset
    Dim sObtNro As String
    'Cargar información del Detalle
    sObtNro = Trim(fgeObt.TextMatrix(fgeObt.Row, 1))
    If Trim(sObtNro) <> "" Then
        Set rs = New ADODB.Recordset
        Set rs = clsDReq.CargaObtDetalle(ObtDetParaAdquisi, sObtNro)
        If rs.RecordCount > 0 Then
            Set fgeObtDet.Recordset = rs
            fgeObtDet.AdicionaFila
            fgeObtDet.BackColorRow &HC0FFFF
            fgeObtDet.TextMatrix(fgeObtDet.Row, 2) = "T O T A L "
            fgeObtDet.TextMatrix(fgeObtDet.Row, 6) = Format(fgeObtDet.SumaRow(6), "#,##0.00")
        End If
        Set rs = Nothing
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
    
    Call optPlan_Click(0)
    'Inicia Periodo
    spinPeriodo.Valor = IIf(psTpoReq = "1", Year(gdFecSis) + 1, Year(gdFecSis))
End Sub

Private Sub cmdBS_Click(Index As Integer)
    Dim sAdqNro As String, sObtNro As String, sBSCod As String, sActualiza As String
    Dim nPrecio As Currency, nCantidad As Currency
    Dim nCont As Integer, nSum As Integer, nResult As Integer
    Dim clsDMov As DLogMov
    Dim clsDGnral As DLogGeneral
    
    Select Case Index
        Case 0:
            'CANCELAR
            If MsgBox("¿ Estás seguro de cancelar toda la operación ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Call cmdSalir_Click
            End If
        Case 1:
            'GRABAR
            'Validación de Requerimientos
            For nCont = 1 To fgeObt.Rows - 1
                If (fgeObt.TextMatrix(nCont, 2)) = "." Then
                    nSum = nSum + 1
                End If
            Next
            If nSum = 0 Then
                MsgBox "Por favor, determine que Planes se consolidan", vbInformation, " Aviso "
                Exit Sub
            End If
            If MsgBox("¿ Estás seguro de consolidar en una Adquisición ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Set clsDGnral = New DLogGeneral
                sAdqNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                Set clsDGnral = Nothing
                
                sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                Set clsDMov = New DLogMov
                
                'Grabación de MOV
                clsDMov.InsertaMov sAdqNro, Trim(Str(gLogOpeAdqRegistro)), "", Trim(Str(gLogAdqEstadoInicio))
                clsDMov.InsertaMovRef sAdqNro, sAdqNro
                
                'Inserta LogAdquisicion
                clsDMov.InsertaAdquisicion IIf(optPlan(0).Value = True, "1", "2"), sAdqNro, _
                    spinPeriodo.Valor, sActualiza
                
                'Inserta ObtAdq y Actualiza LogObtencion
                For nCont = 1 To fgeObt.Rows - 1
                    If (fgeObt.TextMatrix(nCont, 2)) = "." Then
                        sObtNro = fgeObt.TextMatrix(nCont, 1)
                        clsDMov.ActualizaObtencion sObtNro, gLogObtEstadoAdquisi, sActualiza
                        clsDMov.InsertaObtAdq sObtNro, sAdqNro
                    End If
                Next
                
                'Inserta LogAdqDetalle
                For nCont = 1 To fgeBS.Rows - 2
                    sBSCod = fgeBS.TextMatrix(nCont, 1)
                    nCantidad = CCur(IIf(Trim(fgeBS.TextMatrix(nCont, 4)) = "", 0, fgeBS.TextMatrix(nCont, 4)))
                    nPrecio = CCur(IIf(Trim(fgeBS.TextMatrix(nCont, 5)) = "", 0, fgeBS.TextMatrix(nCont, 5)))
                    clsDMov.InsertaAdqDetalle sAdqNro, sBSCod, _
                        nCantidad, nPrecio, sActualiza
                Next
                
                'Ejecuta todos los querys en una transacción
                'nResult = clsDMov.EjecutaBatch
                Set clsDMov = Nothing
                
                If nResult = 0 Then
                    cmdBS(0).Enabled = False
                    cmdBS(1).Enabled = False
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
    fgeObt.Clear
    fgeObt.FormaCabecera
    fgeObt.Rows = 2
    fgeObtDet.Clear
    fgeObtDet.FormaCabecera
    fgeObtDet.Rows = 2
    fgeBS.Clear
    fgeBS.FormaCabecera
    fgeBS.Rows = 2
End Sub

Private Sub optPlan_Click(Index As Integer)
    psTpoReq = IIf(optPlan(0).Value = True, "1", "2")
    spinPeriodo_Change
End Sub

Private Sub spinPeriodo_Change()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'Actualiza FLEX
    Limpiar
    'Carga Plan de Obtención
    Set rs = clsDReq.CargaObtencion(psTpoReq, spinPeriodo.Valor, gLogObtEstadoAceptado)
    If rs.RecordCount > 0 Then
        Set fgeObt.Recordset = rs
        fgeObt.lbEditarFlex = True
        cmdBS(0).Enabled = True
        cmdBS(1).Enabled = True
        Call fgeObt_OnRowChange(fgeObt.Row, fgeObt.Col)
    Else
        cmdBS(0).Enabled = False
        cmdBS(1).Enabled = False
    End If

End Sub


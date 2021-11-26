VERSION 5.00
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmLogReqAprobaExtem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aprobación de Plan Extemporaneo"
   ClientHeight    =   6255
   ClientLeft      =   300
   ClientTop       =   1455
   ClientWidth     =   11130
   Icon            =   "frmLogReqAprobaExtem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBS 
      Caption         =   "&Aprobar"
      Height          =   390
      Index           =   2
      Left            =   6795
      TabIndex        =   13
      Top             =   5670
      Width           =   1305
   End
   Begin VB.CommandButton cmdBS 
      Caption         =   "&Cancelar"
      Height          =   390
      Index           =   0
      Left            =   2460
      TabIndex        =   2
      Top             =   5670
      Width           =   1305
   End
   Begin VB.CommandButton cmdBS 
      Caption         =   "&Rechazar"
      Height          =   390
      Index           =   1
      Left            =   4710
      TabIndex        =   1
      Top             =   5670
      Width           =   1305
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   9150
      TabIndex        =   0
      Top             =   5670
      Width           =   1305
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   90
      Top             =   5685
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin Sicmact.FlexEdit fgeBS 
      Height          =   4815
      Left            =   4185
      TabIndex        =   3
      Top             =   705
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   8493
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-cBSCod-Bien/Servicio-Unidad-Cantidad-Precio-CtaContable"
      EncabezadosAnchos=   "400-0-2000-700-850-950-1500"
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
      EncabezadosAlineacion=   "L-L-L-L-R-R-L"
      FormatosEdit    =   "0-0-0-0-2-2-0"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   2
      lbFormatoCol    =   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   285
   End
   Begin Spinner.uSpinner spinPeriodo 
      Height          =   300
      Left            =   8055
      TabIndex        =   4
      Top             =   90
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
   Begin Sicmact.FlexEdit fgeObt 
      Height          =   2475
      Left            =   165
      TabIndex        =   5
      Top             =   705
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4366
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-Obtención"
      EncabezadosAnchos=   "400-3200"
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
      ColumnasAEditar =   "X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "L-L"
      FormatosEdit    =   "0-0"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   285
   End
   Begin Sicmact.FlexEdit fgeReq 
      Height          =   1980
      Left            =   165
      TabIndex        =   6
      Top             =   3540
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3493
      Cols0           =   3
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-Area-Requerimiento"
      EncabezadosAnchos=   "400-1500-2000"
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
      ColumnasAEditar =   "X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L"
      FormatosEdit    =   "0-0-0"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   285
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
      Left            =   7485
      TabIndex        =   12
      Top             =   150
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
      Left            =   285
      TabIndex        =   11
      Top             =   135
      Width           =   750
   End
   Begin VB.Label lblAreaDes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1050
      TabIndex        =   10
      Top             =   105
      Width           =   4110
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Plan de Obtención :"
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
      Left            =   270
      TabIndex        =   9
      Top             =   465
      Width           =   1830
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Detalle de Plan de Obtención :"
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
      Left            =   4335
      TabIndex        =   8
      Top             =   465
      Width           =   2865
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Requerimientos del Plan :"
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
      Left            =   255
      TabIndex        =   7
      Top             =   3285
      Width           =   2865
   End
End
Attribute VB_Name = "frmLogReqAprobaExtem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim psTpoReq As String
Dim clsDReq As DLogRequeri

Private Sub cmdBS_Click(Index As Integer)
    Dim sObtNro As String, sObtTraNro As String
    Dim sReqNro As String, sActualiza As String
    Dim nCont As Integer, nResult As Integer
    Dim clsDMov As DLogMov
    Dim clsDGnral As DLogGeneral
    Dim rs As ADODB.Recordset
    
    Select Case Index
        Case 0:
            'Cancelar
            If MsgBox("¿ Estás seguro de cancelar toda la operación ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Call cmdSalir_Click
            End If
        Case 1:
            'Rechazar
            If MsgBox("¿ Estás seguro de Rechazar el Plan Extemporaneo ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Set clsDGnral = New DLogGeneral
                sObtNro = Trim(fgeObt.TextMatrix(fgeReq.Row, 1))
                sObtTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                Set clsDGnral = Nothing
                
                Set clsDMov = New DLogMov
                'Inserta Mov
                clsDMov.InsertaMov sObtTraNro, Trim(Str(gLogOpeObtTramite)), "", Trim(Str(gLogObtEstadoRechazado))
                
                clsDMov.ActualizaObtencion sObtNro, gLogObtEstadoRechazado, sActualiza
                
                'Inserta en LogReqTramite
                For nCont = 1 To fgeReq.Rows - 1
                    sReqNro = fgeReq.TextMatrix(nCont, 2)
                    clsDMov.InsertaReqTramite sReqNro, sObtTraNro, Usuario.AreaCod, "", _
                        "", gLogReqEstadoRechazado, gLogReqFlujoSin, sActualiza
                Next
                
                'Ejecuta todos los querys en una transacción
                'nResult = clsDMov.EjecutaBatch
                Set clsDMov = Nothing
                
                If nResult = 0 Then
                    cmdBS(0).Enabled = False
                    cmdBS(1).Enabled = False
                    cmdBS(2).Enabled = False
                    Call spinPeriodo_Change
                Else
                    MsgBox "Error al rechazar el Plan Extemporaneo", vbInformation, " Aviso "
                End If
            End If
        Case 2:
            'Aprobar
            If MsgBox("¿ Estás seguro de Aprobar el Plan Extemporaneo ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Set clsDGnral = New DLogGeneral
                sObtNro = Trim(fgeObt.TextMatrix(fgeObt.Row, 1))
                sObtTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                Set clsDGnral = Nothing
                
                sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                Set clsDMov = New DLogMov
                
                'Grabación de MOV
                clsDMov.InsertaMov sObtTraNro, Trim(Str(gLogOpeObtTramite)), "", Trim(Str(gLogObtEstadoAceptado))
                
                clsDMov.ActualizaObtencion sObtNro, gLogObtEstadoAceptado, sActualiza
                
                'Inserta en LogReqTramite
                For nCont = 1 To fgeReq.Rows - 1
                    sReqNro = fgeReq.TextMatrix(nCont, 2)
                    clsDMov.InsertaReqTramite sReqNro, sObtTraNro, Usuario.AreaCod, "", _
                        "", gLogReqEstadoAceptado, gLogReqFlujoSin, sActualiza
                Next
                
                'Ejecuta todos los querys en una transacción
                'nResult = clsDMov.EjecutaBatch
                Set clsDMov = Nothing
                
                If nResult = 0 Then
                    cmdBS(0).Enabled = False
                    cmdBS(1).Enabled = False
                    cmdBS(2).Enabled = False
                    Call spinPeriodo_Change
                Else
                    MsgBox "Error al Aprobar el Plan Extemporaneo", vbInformation, " Aviso "
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

Private Sub fgeObt_OnRowChange(pnRow As Long, pnCol As Long)
    'Refrescar los detalles para ingresar ctas.cntables.
    Dim sObtNro As String
    Dim rs As ADODB.Recordset
    
    sObtNro = Trim(fgeObt.TextMatrix(fgeObt.Row, 1))
    If Trim(sObtNro) <> "" Then
        Set rs = New ADODB.Recordset
        Set rs = clsDReq.CargaObtDetalle(ObtDetParaAprobar, sObtNro)
        If rs.RecordCount > 0 Then
            Set fgeBS.Recordset = rs
            fgeBS.lbEditarFlex = True
            
            'Carga Flex de Requerimientos de Plan Obtención
            Set rs = clsDReq.CargaRequerimiento(psTpoReq, ReqTodosObten, "", sObtNro)
            If rs.RecordCount > 0 Then
                Set fgeReq.Recordset = rs
            End If
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
    
    'OJO. Siempre es Extemporaneo
    psTpoReq = "2"
    
    'Inicia Periodo
    
    spinPeriodo.Valor = Year(gdFecSis)
End Sub

Private Sub Limpiar()
    fgeObt.Clear
    fgeObt.FormaCabecera
    fgeObt.Rows = 2
    fgeReq.Clear
    fgeReq.FormaCabecera
    fgeReq.Rows = 2
    fgeBS.Clear
    fgeBS.FormaCabecera
    fgeBS.Rows = 2
    fgeBS.lbEditarFlex = False
End Sub

Private Sub spinPeriodo_Change()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'Actualiza FLEX
    Limpiar
        'Carga Plan de Obtención
    Set rs = clsDReq.CargaObtencion(psTpoReq, spinPeriodo.Valor, gLogObtEstadoCuenta)
    If rs.RecordCount > 0 Then
        Set fgeObt.Recordset = rs
        cmdBS(0).Enabled = True
        cmdBS(1).Enabled = True
        cmdBS(2).Enabled = True
        Call fgeObt_OnRowChange(fgeObt.Row, fgeObt.Col)
    Else
        cmdBS(0).Enabled = False
        cmdBS(1).Enabled = False
        cmdBS(2).Enabled = False
    End If
End Sub



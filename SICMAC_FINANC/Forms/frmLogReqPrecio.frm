VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmLogReqPrecio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Precios de Bienes/Servicios"
   ClientHeight    =   6060
   ClientLeft      =   510
   ClientTop       =   1665
   ClientWidth     =   10950
   Icon            =   "frmLogReqPrecio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReqPre 
      Caption         =   "&Grabar"
      Height          =   390
      Index           =   1
      Left            =   5625
      TabIndex        =   18
      Top             =   5595
      Width           =   1305
   End
   Begin VB.CommandButton cmdReqPre 
      Caption         =   "&Cancelar"
      Height          =   390
      Index           =   0
      Left            =   3585
      TabIndex        =   17
      Top             =   5595
      Width           =   1305
   End
   Begin TabDlg.SSTab sstReq 
      Height          =   4710
      Left            =   90
      TabIndex        =   7
      Top             =   795
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   8308
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "S&ustentación"
      TabPicture(0)   =   "frmLogReqPrecio.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "rtfDescri(0)"
      Tab(0).Control(1)=   "rtfDescri(1)"
      Tab(0).Control(2)=   "lblEtiqueta(3)"
      Tab(0).Control(3)=   "lblEtiqueta(4)"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Detalle"
      TabPicture(1)   =   "frmLogReqPrecio.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblEtiqueta(6)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblEtiqueta(5)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fgeMes"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fgeBSMes"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fgeBS"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin Sicmact.FlexEdit fgeBS 
         Height          =   3990
         Left            =   90
         TabIndex        =   16
         Top             =   540
         Width           =   7770
         _ExtentX        =   13705
         _ExtentY        =   7038
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   1
         EncabezadosNombres=   "Item-Código-Descripción-Unidad-Moneda-Pre.Uni.Ref.-Precio-Cuenta Contable"
         EncabezadosAnchos=   "450-1000-2500-700-650-900-1200-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-6-7"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-1"
         EncabezadosAlineacion=   "R-L-L-L-L-R-R-L"
         FormatosEdit    =   "0-0-0-0-0-2-2-0"
         CantEntero      =   6
         CantDecimales   =   1
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   2
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
      End
      Begin Sicmact.FlexEdit fgeBSMes 
         Height          =   3990
         Left            =   7905
         TabIndex        =   8
         Top             =   540
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   7038
         Cols0           =   4
         HighLight       =   2
         AllowUserResizing=   1
         EncabezadosNombres=   "Mes-Código-Descripción-Cantidad"
         EncabezadosAnchos=   "400-0-1070-900"
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
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
      End
      Begin Sicmact.FlexEdit fgeMes 
         Height          =   1425
         Left            =   90
         TabIndex        =   9
         Top             =   2730
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
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-L-L-L-L-L-L-L-L"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
      End
      Begin RichTextLib.RichTextBox rtfDescri 
         Height          =   4035
         Index           =   0
         Left            =   -74880
         TabIndex        =   10
         Top             =   555
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   7117
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         MaxLength       =   8000
         TextRTF         =   $"frmLogReqPrecio.frx":0342
      End
      Begin RichTextLib.RichTextBox rtfDescri 
         Height          =   4035
         Index           =   1
         Left            =   -69600
         TabIndex        =   11
         Top             =   555
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   7117
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         MaxLength       =   8000
         TextRTF         =   $"frmLogReqPrecio.frx":03C4
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Bienes/Servicios"
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
         Left            =   210
         TabIndex        =   15
         Top             =   360
         Width           =   1500
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
         Index           =   6
         Left            =   7995
         TabIndex        =   14
         Top             =   345
         Width           =   675
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
         Index           =   3
         Left            =   -74760
         TabIndex        =   13
         Top             =   345
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
         Index           =   4
         Left            =   -69465
         TabIndex        =   12
         Top             =   360
         Width           =   1425
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   8760
      TabIndex        =   0
      Top             =   5610
      Width           =   1305
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   120
      Top             =   5580
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label lblReqNro 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1170
      TabIndex        =   6
      Top             =   435
      Width           =   2625
   End
   Begin VB.Label lblPeriodo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   8100
      TabIndex        =   5
      Top             =   105
      Width           =   840
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
      Left            =   7500
      TabIndex        =   4
      Top             =   150
      Width           =   660
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Número :"
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
      Left            =   345
      TabIndex        =   3
      Top             =   450
      Width           =   825
   End
   Begin VB.Label lblAreaDes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1170
      TabIndex        =   2
      Top             =   105
      Width           =   4110
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
      TabIndex        =   1
      Top             =   135
      Width           =   750
   End
End
Attribute VB_Name = "frmLogReqPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim psTpoReq As String
Dim psFrmTpo As String
Dim p_sReqNro As String
Dim clsDGnral As DLogGeneral
Dim clsDReq As DLogRequeri

Public Sub Inicio(ByVal psTipoReq As String, ByVal psFormTpo As String, ByVal psReqNro As String)
psTpoReq = psTipoReq
psFrmTpo = psFormTpo
p_sReqNro = psReqNro

If psTpoReq = "1" Then
    If psFrmTpo = "1" Then
        Me.Caption = "Registro de Precios de Proyección de Requerimiento"
    ElseIf psFrmTpo = "2" Then
        Me.Caption = "Registro de Cuentas Contables de Proyección de Requerimiento"
    End If
Else
    If psFrmTpo = "1" Then
        Me.Caption = "Registro de Precios de Requerimiento Extemporaneo"
    ElseIf psFrmTpo = "2" Then
        Me.Caption = "Registro de Cuentas Contables de Requerimiento Extemporaneo"
    End If
End If
If psFrmTpo = "2" Then
    fgeBS.EncabezadosAnchos = "450-1000-2500-700-0-0-1000-1800"
    fgeBS.ColumnasAEditar = "X-X-X-X-X-X-X-7"
    fgeBS.lbFlexDuplicados = True
End If
Me.Show 1
End Sub

Private Sub cmdReqPre_Click(Index As Integer)
    Dim clsDMov As DLogMov
    Dim sReqNro As String, sReqTraNro As String, sBSCod As String, sActualiza As String
    Dim nRefPrecio As Currency, nPrecio As Currency, nCant As Currency
    Dim nBs As Integer, nBSMes As Integer, nResult As Integer
    Dim sCtaCont As String
    Dim bGraba_Parte As Boolean
    
    bGraba_Parte = False
    
    Select Case Index
        Case 0:
            'CANCELAR
            If MsgBox("¿ Estás seguro de cancelar toda la operación ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Call cmdSalir_Click
            End If
        Case 1:
            'GRABAR
            If psFrmTpo = "1" Then
                'Validación para Precio
                For nBs = 1 To fgeBS.Rows - 1
                    If Not (Val(fgeBS.TextMatrix(nBs, 6)) > 0) Then
                        MsgBox "Falta ingresar precio en el item " & nBs & " (" & fgeBS.TextMatrix(nBs, 2) & ")", vbInformation, " Aviso "
                        Exit Sub
                    End If
                Next
            Else
                'Validación para Cta.Cont.
                For nBs = 1 To fgeBS.Rows - 1
                    If Not (Val(fgeBS.TextMatrix(nBs, 7)) > 0) Then
                        MsgBox "Falta ingresar cuenta contable en el item " & nBs & " (" & fgeBS.TextMatrix(nBs, 2) & ")", vbInformation, " Aviso "
                        Exit Sub
                    End If
                Next
            End If
            
            If MsgBox("¿ Estás seguro de Grabar estos datos ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                
                'If psTpoReq = "1" Then
                    sReqNro = lblReqNro.Caption
                    sReqTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    
                    'If psFrmTpo = "1" Then
                        'Grabación de Precio
                        clsDMov.InsertaMov sReqTraNro, Trim(Str(gLogOpeReqTramite)), "", Trim(Str(gLogReqEstadoPrecio))
                        clsDMov.InsertaMovRef sReqTraNro, sReqNro
                        
                        clsDMov.InsertaReqTramite sReqNro, sReqTraNro, Usuario.AreaCod, "", _
                            "", Trim(Str(gLogReqEstadoPrecio)), gLogReqFlujoSin, sActualiza
                    'Else
                        'Grabación de Cta.Cont.
                        'clsDMov.InsertaMov sReqTraNro, Trim(Str(gLogOpeReqTramite)), "", Trim(Str(gLogReqEstadoCuenta))
                        
                        'clsDMov.InsertaReqTramite sReqNro, sReqTraNro, Usuario.AreaCod, "", _
                            "", Trim(Str(gLogReqEstadoCuenta)), gLogReqFlujoSin, sActualiza
                    'End If
                    
                    'Si no ha modificado detalle, lo agrega tal como está
                    nBs = 0: nBSMes = 0
                    For nBs = 1 To fgeBS.Rows - 1
                        sBSCod = fgeBS.TextMatrix(nBs, 1)
                        nRefPrecio = CCur(IIf(fgeBS.TextMatrix(nBs, 5) = "", 0, fgeBS.TextMatrix(nBs, 5)))
                        nPrecio = CCur(IIf(fgeBS.TextMatrix(nBs, 6) = "", 0, fgeBS.TextMatrix(nBs, 6)))
                        sCtaCont = Trim(fgeBS.TextMatrix(nBs, 7))
                        clsDMov.InsertaReqDetalle sReqNro, sReqTraNro, sBSCod, _
                            Trim(Right(fgeBS.TextMatrix(nBs, 4), 2)), nRefPrecio, nPrecio, sCtaCont, sActualiza
                        For nBSMes = 1 To fgeMes.Cols - 1
                            nCant = CCur(IIf(fgeMes.TextMatrix(nBs, nBSMes) = "", 0, fgeMes.TextMatrix(nBs, nBSMes)))
                            If nCant > 0 Then
                                clsDMov.InsertaReqDetMes sReqNro, sReqTraNro, sBSCod, _
                                     Val(fgeBSMes.TextMatrix(nBSMes, 1)), nCant
                            End If
                        Next
                    Next
                
                'End If
                'Ejecuta todos los querys en una transacción
                'nResult = clsDMov.EjecutaBatch
                Set clsDMov = Nothing
                
                If nResult = 0 Then
                    cmdReqPre(0).Enabled = False
                    cmdReqPre(1).Enabled = False
                    fgeBS.lbEditarFlex = False
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
    Set clsDReq = Nothing
    Unload Me
End Sub


Private Sub fgeBS_OnRowChange(pnRow As Long, pnCol As Long)
    Dim nCont As Integer
    'Carga Meses del Item de acuerdo al Flex fgeMes
    For nCont = 1 To fgeBSMes.Rows - 1
        fgeBSMes.TextMatrix(nCont, 3) = fgeMes.TextMatrix(pnRow, nCont)
    Next
End Sub

Private Sub Form_Load()
    Dim clsDCC As DCtaCont
    Set clsDGnral = New DLogGeneral
    Set clsDReq = New DLogRequeri
    Call CentraForm(Me)
    'Carga información de la relación usuario-area
    Usuario.Inicio gsCodUser
    If Len(Usuario.AreaCod) = 0 Then
        MsgBox "Usuario no determinado", vbInformation, "Aviso"
        Exit Sub
    End If
    'Si es psFrmTpo = "2" cargar RS de CtasCntbles.
    If psFrmTpo = "2" Then
        Set clsDCC = New DCtaCont
        fgeBS.rsTextBuscar = clsDCC.CargaCtaCont
        Set clsDCC = Nothing
    End If
    'Carga Meses
    fgeBSMes.rsFlex = clsDGnral.CargaConstante(gMeses)
    'Carga Registro
    lblReqNro.Caption = p_sReqNro
    lblAreaDes.Caption = Usuario.AreaNom
    Call CargaDatos(lblReqNro.Caption)
End Sub

Private Sub CargaDatos(ByVal psReqNro As String)
    Dim sBSCod As String
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    Set rs = clsDReq.CargaRequerimiento(psTpoReq, ReqUnRegistro, "", psReqNro)
    If rs.RecordCount = 1 Then
        With rs
            lblAreaDes.Caption = !cAreaDescripcion
            lblPeriodo.Caption = !cLogReqPeriodo
            rtfDescri(0).Text = !cLogReqNecesidad
            rtfDescri(1).Text = !cLogReqRequerimiento
        End With
    Else
        Set rs = Nothing
        MsgBox "Problemas al cargar información del Requerimiento", vbInformation, " Aviso"
        Exit Sub
    End If
    Set rs = Nothing
    
    'Cargar información del Detalle
    Set rs = clsDReq.CargaReqDetalle(ReqDetUnRegistroPrecio, psReqNro)
    If rs.RecordCount > 0 Then Set fgeBS.Recordset = rs
    Set rs = Nothing
    
    'Cargar información del DetMes
    Set rs = clsDReq.CargaReqDetMes(psReqNro, "")
    If rs.RecordCount > 0 Then Set fgeMes.Recordset = rs
    Set rs = Nothing
    
    'Actualiza fgeBSDetMes
    Call fgeBS_OnRowChange(fgeBS.Row, fgeBS.Col)
End Sub

VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLogReqPrecio 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6135
   ClientLeft      =   510
   ClientTop       =   1665
   ClientWidth     =   10950
   Icon            =   "frmLogReqPrecio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReqPre 
      Caption         =   "&Rechazar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   3
      Left            =   6615
      TabIndex        =   22
      Top             =   5625
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmdReqPre 
      Caption         =   "&Aprobar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   2
      Left            =   4830
      TabIndex        =   21
      Top             =   5625
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmdReqPre 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   1
      Left            =   5910
      TabIndex        =   15
      Top             =   5625
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmdReqPre 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   0
      Left            =   3735
      TabIndex        =   14
      Top             =   5625
      Width           =   1305
   End
   Begin TabDlg.SSTab sstReq 
      Height          =   4920
      Left            =   90
      TabIndex        =   4
      Top             =   585
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   8678
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "S&ustentación"
      TabPicture(0)   =   "frmLogReqPrecio.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "rtfDescri(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "rtfDescri(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblEtiqueta(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblEtiqueta(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Detalle"
      TabPicture(1)   =   "frmLogReqPrecio.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblEtiqueta(6)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblEtiqueta(5)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblMoneda"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblEtiqueta(8)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblMonedaFinal"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblTipCambio"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "fgeMes"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "fgeBSMes"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "fgeBS"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cboMoneda"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtTipCambio"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      Begin VB.TextBox txtTipCambio 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   5085
         TabIndex        =   23
         Top             =   375
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.ComboBox cboMoneda 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5085
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   375
         Visible         =   0   'False
         Width           =   1335
      End
      Begin Sicmact.FlexEdit fgeBS 
         Height          =   3780
         Left            =   90
         TabIndex        =   13
         Top             =   960
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   6668
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   1
         EncabezadosNombres=   "Item-Código-Descripción-Unidad-Pre.Uni.Ref.-Pre.Unitario-Cta.Contable-Cta.Cont.Descripción"
         EncabezadosAnchos=   "450-1000-2300-700-900-1000-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-5-6-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-1-0"
         EncabezadosAlineacion=   "R-L-L-L-R-R-L-L"
         FormatosEdit    =   "0-0-0-0-2-2-0-0"
         CantEntero      =   6
         CantDecimales   =   1
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   2
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   450
         RowHeight0      =   285
      End
      Begin Sicmact.FlexEdit fgeBSMes 
         Height          =   4095
         Left            =   7905
         TabIndex        =   5
         Top             =   645
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   7223
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
         RowHeight0      =   240
      End
      Begin Sicmact.FlexEdit fgeMes 
         Height          =   1425
         Left            =   225
         TabIndex        =   6
         Top             =   3255
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
         RowHeight0      =   225
      End
      Begin RichTextLib.RichTextBox rtfDescri 
         Height          =   4185
         Index           =   0
         Left            =   -74880
         TabIndex        =   7
         Top             =   555
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   7382
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         MaxLength       =   8000
         TextRTF         =   $"frmLogReqPrecio.frx":0342
      End
      Begin RichTextLib.RichTextBox rtfDescri 
         Height          =   4185
         Index           =   1
         Left            =   -69600
         TabIndex        =   8
         Top             =   555
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   7382
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         MaxLength       =   8000
         TextRTF         =   $"frmLogReqPrecio.frx":03B0
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
         Left            =   3720
         TabIndex        =   24
         Top             =   435
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lblMonedaFinal 
         Caption         =   "Moneda Final :"
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
         Left            =   3630
         TabIndex        =   19
         Top             =   435
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Moneda :"
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
         Index           =   8
         Left            =   255
         TabIndex        =   18
         Top             =   405
         Width           =   885
      End
      Begin VB.Label lblMoneda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1410
         TabIndex        =   17
         Top             =   390
         Width           =   1110
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
         Left            =   240
         TabIndex        =   12
         Top             =   720
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
         Left            =   8100
         TabIndex        =   11
         Top             =   420
         Width           =   480
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
         TabIndex        =   10
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
         TabIndex        =   9
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
   Begin Sicmact.TxtBuscar txtBuscar 
      Height          =   300
      Left            =   1275
      TabIndex        =   16
      Top             =   135
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   529
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      TipoBusqueda    =   2
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   0
      Top             =   5610
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label lblPeriodo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   8100
      TabIndex        =   3
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
      TabIndex        =   2
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
      Left            =   405
      TabIndex        =   1
      Top             =   165
      Width           =   825
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
Dim psReqNro As String
Dim clsDGnral As DLogGeneral
Dim clsDReq As DLogRequeri

Public Sub Inicio(ByVal psTipoReq As String, ByVal psFormTpo As String, Optional ByVal psRequeriNro As String = "")
psTpoReq = psTipoReq
psFrmTpo = psFormTpo
psReqNro = psRequeriNro
Me.Show 1
End Sub

Private Sub cmdReqPre_Click(Index As Integer)
    Dim clsDMov As DLogMov
    Dim nReqNro As Integer, nReqTraNro As Integer, nReqTraNroAnt As Integer
    Dim sReqNro As String, sReqTraNro As String, sReqTraNroAnt As String, sBSCod As String, sActualiza As String
    Dim nMoneda As Integer
    Dim nRefPrecio As Currency, nPrecio As Currency, nCant As Currency, nTipCambio As Currency
    Dim nBs As Integer, nBSMes As Integer, nResult As Integer
    Dim sPlaPreNro As String, sPlaRubro As String, sPlaCtaCont As String
    
    Dim bGraba_Parte As Boolean
    
    bGraba_Parte = False
    
    Select Case Index
        Case 0:
            'CANCELAR
            If MsgBox("¿ Estás seguro de cancelar toda la operación ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                If psFrmTpo = "1" Then
                    cmdReqPre(0).Enabled = False
                    cmdReqPre(1).Enabled = False
                    cboMoneda.Enabled = False
                ElseIf psFrmTpo = "2" Then
                    cmdReqPre(0).Enabled = False
                    cmdReqPre(1).Enabled = False
                ElseIf psFrmTpo = "3" Then
                    cmdReqPre(0).Enabled = False
                    cmdReqPre(2).Enabled = False
                    cmdReqPre(3).Enabled = False
                End If
                txtBuscar.Text = ""
                Call Limpiar
            End If
        Case 1:
            'GRABAR (PRECIO - CTA.CONTABLE)
            If psFrmTpo = "1" Then
                'Validación para Precio
                For nBs = 1 To fgeBS.Rows - 1
                    If Not (Val(fgeBS.TextMatrix(nBs, 5)) > 0) Then
                        MsgBox "Falta ingresar precio en el item " & nBs & " (" & fgeBS.TextMatrix(nBs, 2) & ")", vbInformation, " Aviso "
                        Exit Sub
                    End If
                Next
                'Moneda Final
                If Trim(cboMoneda.Text) = "" Then
                    MsgBox "Falta determinar la moneda del requerimiento", vbInformation, " Aviso "
                    Exit Sub
                End If
                nMoneda = Val(Right(Trim(cboMoneda.Text), 1))
                If Not (nMoneda = gMonedaNacional Or nMoneda = gMonedaExtranjera) Then
                    MsgBox "Moneda no reconocida", vbInformation, " Aviso "
                    Exit Sub
                End If
            Else
                'Validación para Cta.Cont.
                For nBs = 1 To fgeBS.Rows - 1
                    If Not (Val(fgeBS.TextMatrix(nBs, 6)) > 0) Then
                        MsgBox "Falta ingresar cuenta contable en el item " & nBs & " (" & fgeBS.TextMatrix(nBs, 2) & ")", vbInformation, " Aviso "
                        Exit Sub
                    End If
                Next
                nMoneda = Val(Right(Trim(lblMoneda.Caption), 1))
                If Not (nMoneda = gMonedaNacional Or nMoneda = gMonedaExtranjera) Then
                    MsgBox "Moneda no reconocida", vbInformation, " Aviso "
                    Exit Sub
                End If
                'Tipo Cambio
                If Right(lblMoneda.Caption, 1) = gMonedaExtranjera And psFrmTpo = "2" Then
                    nTipCambio = Val(txtTipCambio.Text)
                    If nTipCambio = 0 Then
                        MsgBox "Falta ingresar el Tipo de Cambio ", vbInformation, " Aviso "
                        Exit Sub
                    End If
                Else
                    nTipCambio = 0
                End If
            End If
            
            If MsgBox("¿ Estás seguro de Grabar estos datos ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                
                sReqNro = txtBuscar.Text
                sReqTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                If psFrmTpo = "2" Then
                    nReqTraNroAnt = clsDReq.CargaUltReqDetNro(clsDGnral.GetnMovNro(sReqNro))
                    'nReqTraNroAnt = clsDMov.GetnMovNro(sReqTraNroAnt)
                End If
                
                sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                Set clsDMov = New DLogMov
                
                If psFrmTpo = "1" Then
                    'Precio (Inserta)
                    clsDMov.InsertaMov sReqTraNro, Trim(Str(gLogOpeReqTramite)), "", gLogReqEstadoPrecio
                    nReqTraNro = clsDMov.GetnMovNro(sReqTraNro)
                    nReqNro = clsDMov.GetnMovNro(sReqNro)
                    clsDMov.InsertaMovRef nReqTraNro, nReqNro
                    
                    clsDMov.InsertaReqTramite nReqNro, nReqTraNro, Usuario.AreaCod, "", _
                        "", gLogReqEstadoPrecio, nMoneda, sActualiza
                
                    'Si no ha modificado detalle, lo agrega tal como está
                    nBs = 0: nBSMes = 0
                    For nBs = 1 To fgeBS.Rows - 1
                        sBSCod = fgeBS.TextMatrix(nBs, 1)
                        nRefPrecio = CCur(IIf(fgeBS.TextMatrix(nBs, 5) = "", 0, fgeBS.TextMatrix(nBs, 5)))
                        'Inserta ReqDetalle
                        clsDMov.InsertaReqDetalle nReqNro, nReqTraNro, sBSCod, _
                             nRefPrecio, sActualiza
                        For nBSMes = 1 To fgeMes.Cols - 1
                            nCant = CCur(IIf(fgeMes.TextMatrix(nBs, nBSMes) = "", 0, fgeMes.TextMatrix(nBs, nBSMes)))
                            If nCant > 0 Then
                                'Inserta ReqDetMes
                                clsDMov.InsertaReqDetMes nReqNro, nReqTraNro, sBSCod, _
                                     Val(fgeBSMes.TextMatrix(nBSMes, 1)), nCant
                            End If
                        Next
                    Next
                
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        cmdReqPre(0).Enabled = False
                        cmdReqPre(1).Enabled = False
                        fgeBS.lbEditarFlex = False
                        cboMoneda.Enabled = False
                        Call CargaTxtBuscar
                    Else
                        MsgBox "Error al grabar la información", vbInformation, " Aviso "
                    End If
                
                Else
                    'Cta.Cont. (Actualiza)
                    clsDMov.InsertaMov sReqTraNro, Trim(Str(gLogOpeReqTramite)), "", Trim(Str(gLogReqEstadoCuenta))
                    nReqTraNro = clsDMov.GetnMovNro(sReqTraNro)
                    nReqNro = clsDMov.GetnMovNro(sReqNro)
                    clsDMov.InsertaMovRef nReqTraNro, nReqNro
                    
                    If Right(lblMoneda.Caption, 1) = gMonedaExtranjera And psFrmTpo = "2" Then
                        'Inserta Tipo Cambio Actualiza la moneda
                        clsDMov.ActualizaRequeriTipCambio nReqNro, nTipCambio
                    End If
                    
                    clsDMov.InsertaReqTramite nReqNro, nReqTraNro, Usuario.AreaCod, "", _
                        "", gLogReqEstadoCuenta, nMoneda, sActualiza
                    
                    'Actualiza las ctas. contables
                    nBs = 0: nBSMes = 0
                    For nBs = 1 To fgeBS.Rows - 1
                        sBSCod = fgeBS.TextMatrix(nBs, 1)
                        sPlaCtaCont = Trim(fgeBS.TextMatrix(nBs, 6))
                        sPlaPreNro = Left(Trim(fgeBS.TextMatrix(nBs, 7)), InStr(1, fgeBS.TextMatrix(nBs, 7), "-") - 1)
                        sPlaRubro = Mid(Trim(fgeBS.TextMatrix(nBs, 7)), InStr(1, fgeBS.TextMatrix(nBs, 7), "-") + 1)
                        'Inserta ReqDetalle
                        clsDMov.ActualizaReqDetalle nReqNro, nReqTraNroAnt, sBSCod, _
                             sPlaPreNro, sPlaRubro, sPlaCtaCont, sActualiza, Me.lblPeriodo.Caption
                    Next
                
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        cmdReqPre(0).Enabled = False
                        cmdReqPre(1).Enabled = False
                        fgeBS.lbEditarFlex = False
                        txtTipCambio.Enabled = False
                        Call CargaTxtBuscar
                    Else
                        MsgBox "Error al grabar la información", vbInformation, " Aviso "
                    End If
                End If
            End If
        Case 2:
            nMoneda = Val(Right(Trim(lblMoneda.Caption), 1))
            If Not (nMoneda = gMonedaNacional Or nMoneda = gMonedaExtranjera) Then
                MsgBox "Moneda no reconocida", vbInformation, " Aviso "
                Exit Sub
            End If
            'APROBAR
            If MsgBox("¿ Estás seguro de Aprobar este requerimiento ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                sReqNro = txtBuscar.Text
                sReqTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                
                sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                Set clsDMov = New DLogMov
                
                'Inserta Mov - MovRef
                clsDMov.InsertaMov sReqTraNro, Trim(Str(gLogOpeReqTramite)), "", Trim(Str(gLogReqEstadoAceptado))
                nReqTraNro = clsDMov.GetnMovNro(sReqTraNro)
                nReqNro = clsDMov.GetnMovNro(sReqNro)
                clsDMov.InsertaMovRef nReqTraNro, nReqNro
                
                clsDMov.InsertaReqTramite nReqNro, nReqTraNro, Usuario.AreaCod, "", _
                    "", gLogReqEstadoAceptado, nMoneda, sActualiza
                
                'Ejecuta todos los querys en una transacción
                'nResult = clsDMov.EjecutaBatch
                Set clsDMov = Nothing
                
                If nResult = 0 Then
                    cmdReqPre(0).Enabled = False
                    cmdReqPre(2).Enabled = False
                    cmdReqPre(3).Enabled = False
                    Call CargaTxtBuscar
                Else
                    MsgBox "Error al grabar la información", vbInformation, " Aviso "
                End If
            End If
        Case 3:
            nMoneda = Val(Right(Trim(lblMoneda.Caption), 1))
            If Not (nMoneda = gMonedaNacional Or nMoneda = gMonedaExtranjera) Then
                MsgBox "Moneda no reconocida", vbInformation, " Aviso "
                Exit Sub
            End If
            'RECHAZAR
            If MsgBox("¿ Estás seguro de Rechazar este requerimiento ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                sReqNro = txtBuscar.Text
                sReqTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                
                sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                Set clsDMov = New DLogMov
                
                'Inserta Mov - MovRef
                clsDMov.InsertaMov sReqTraNro, Trim(Str(gLogOpeReqTramite)), "", gLogReqEstadoRechazado
                nReqTraNro = clsDMov.GetnMovNro(sReqTraNro)
                nReqNro = clsDMov.GetnMovNro(sReqNro)
                clsDMov.InsertaMovRef nReqTraNro, nReqNro
                
                clsDMov.InsertaReqTramite nReqNro, nReqTraNro, Usuario.AreaCod, "", _
                    "", gLogReqEstadoRechazado, nMoneda, sActualiza
            
                'Ejecuta todos los querys en una transacción
                'nResult = clsDMov.EjecutaBatch
                Set clsDMov = Nothing
                
                If nResult = 0 Then
                    cmdReqPre(0).Enabled = False
                    cmdReqPre(2).Enabled = False
                    cmdReqPre(3).Enabled = False
                    Call CargaTxtBuscar
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

Private Sub fgeBS_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)
    Dim nReqNro As Integer, nReqTraNro As Integer
    Dim sReqNro As String, sReqTraNro As String, sBSCod As String, sCtaCont As String
    Dim clsDReq As DLogRequeri
    Dim nTipCambio As Currency
    
    If Right(lblMoneda.Caption, 1) = gMonedaExtranjera Then
        nTipCambio = CCur(IIf(txtTipCambio.Text = "", 0, txtTipCambio.Text))
        If nTipCambio <= 0 Then
            MsgBox "Por favor ingrese primero el tipo de cambio", vbInformation, " Aviso"
            Exit Sub
        End If
    Else
        nTipCambio = 1
    End If
    sReqNro = txtBuscar.Text
    nReqNro = clsDGnral.GetnMovNro(sReqNro)
    Set clsDReq = New DLogRequeri
    nReqTraNro = clsDReq.CargaUltReqDetNro(nReqNro)
    Set clsDReq = Nothing
    sBSCod = fgeBS.TextMatrix(fgeBS.Row, 1)
    sCtaCont = Trim(fgeBS.TextMatrix(fgeBS.Row, 6) & Space(40) & fgeBS.TextMatrix(fgeBS.Row, 7))
    
    sCtaCont = frmLogReqPresupu.Inicio("1", lblPeriodo.Caption, nReqNro, _
         nReqTraNro, sBSCod, sCtaCont, nTipCambio)
         
    If sCtaCont <> "" Then
        'fgeBS.TextMatrix(fgeBS.Row, 6) = sCtaCont
        psCodigo = Trim(Left(sCtaCont, 40))
        psDescripcion = Trim(Mid(sCtaCont, 40))
    End If
End Sub

Private Sub fgeBS_OnRowChange(pnRow As Long, pnCol As Long)
    Dim nCont As Integer
    'Carga Meses del Item de acuerdo al Flex fgeMes
    For nCont = 1 To fgeBSMes.Rows - 1
        fgeBSMes.TextMatrix(nCont, 3) = fgeMes.TextMatrix(pnRow, nCont)
    Next
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Set clsDGnral = New DLogGeneral
    Set clsDReq = New DLogRequeri
    Set rs = New ADODB.Recordset
    Call CentraForm(Me)
     
    Usuario.Inicio gsCodUser
    If Len(Usuario.AreaCod) = 0 Then
        txtBuscar.Enabled = False
        cmdReqPre(0).Enabled = False
        cmdReqPre(1).Enabled = False
        cmdReqPre(2).Enabled = False
        cmdReqPre(3).Enabled = False
        sstReq.Enabled = False
        MsgBox "Usuario no determinado", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If psTpoReq = "1" Then
        If psFrmTpo = "1" Then
            Me.Caption = "Requerimiento Regular : Precios Referenciales"
        ElseIf psFrmTpo = "2" Then
            Me.Caption = "Requerimiento Regular : Cuentas Contables"
        ElseIf psFrmTpo = "3" Then
            Me.Caption = "Requerimiento Regular : Aprobación o Rechazo"
        End If
    Else
        If psFrmTpo = "1" Then
            Me.Caption = "Requerimiento Extemporaneo : Precios Referenciales"
        ElseIf psFrmTpo = "2" Then
            Me.Caption = "Requerimiento Extemporaneo : Cuentas Contables"
        ElseIf psFrmTpo = "3" Then
            Me.Caption = "Requerimiento Extemporaneo : Aprobación o Rechazo"
        End If
    End If
    'Si es psFrmTpo = "2" cargar RS de CtasCntbles.
    If psFrmTpo = "1" Then
        'PRECIOS
        lblMonedaFinal.Visible = True
        cboMoneda.Visible = True
        cmdReqPre(1).Visible = True
        'Monedas
        Set rs = clsDGnral.CargaConstante(gMoneda, False)
        Call CargaCombo(rs, cboMoneda)
    ElseIf psFrmTpo = "2" Then
        'CTAS.CONTABLES
        fgeBS.TipoBusqueda = BuscaLibre
        cmdReqPre(1).Visible = True
        'Cambiar FLEX
        fgeBS.EncabezadosAnchos = "450-1000-2300-700-900-0-1500"
    ElseIf psFrmTpo = "3" Then
        'APROBACION
        cmdReqPre(0).Left = cmdReqPre(0).Left - 700
        cmdReqPre(2).Visible = True
        cmdReqPre(3).Visible = True
        fgeBS.EncabezadosAnchos = "450-1000-2300-700-900-0-0"
    Else
        MsgBox "Opción no reconocida", vbInformation, " Aviso "
        Exit Sub
    End If
    'Carga Meses
    fgeBSMes.rsFlex = clsDGnral.CargaConstante(gMeses)
    
    'Carga los requerimientos pendientes del area
    Call CargaTxtBuscar
End Sub

Private Sub txtBuscar_EmiteDatos()
    Dim sReqNro As String, sBSCod As String
    Dim rs As ADODB.Recordset
    Dim nReqTraNro  As Integer, nCont As Integer
    
    If txtBuscar.OK = False Then
        Exit Sub
    End If
    
    sReqNro = txtBuscar.Text
    If sReqNro <> "" Then
        Set rs = New ADODB.Recordset
        Set rs = clsDReq.CargaRequerimiento(psTpoReq, ReqUnRegistroTramite, "", clsDGnral.GetnMovNro(sReqNro))
        If rs.RecordCount = 1 Then
            With rs
                'lblAreaDes.Caption = !cAreaDescripcion
                lblPeriodo.Caption = !nLogReqPeriodo
                lblMoneda.Caption = !cLogReqMoneda
                rtfDescri(0).Text = !cLogReqNecesidad
                rtfDescri(1).Text = !cLogReqRequerimiento
                If Val(Right(lblMoneda.Caption, 1)) = gMonedaExtranjera And psFrmTpo = "2" Then
                    lblTipCambio.Visible = True
                    txtTipCambio.Visible = True
                    txtTipCambio.Enabled = True
                Else
                    lblTipCambio.Visible = False
                    txtTipCambio.Visible = False
                    txtTipCambio.Enabled = False
                End If
            End With
        Else
            cmdReqPre(0).Enabled = False
            cmdReqPre(1).Enabled = False
            Set rs = Nothing
            MsgBox "Problemas al cargar información del Requerimiento", vbInformation, " Aviso"
            Exit Sub
        End If
        Set rs = Nothing
        fgeBS.lbEditarFlex = True
        If psFrmTpo = "1" Then
            cboMoneda.Enabled = True
            cboMoneda.ListIndex = -1
            cmdReqPre(0).Enabled = True
            cmdReqPre(1).Enabled = True
        ElseIf psFrmTpo = "2" Then
            cmdReqPre(0).Enabled = True
            cmdReqPre(1).Enabled = True
        ElseIf psFrmTpo = "3" Then
            cmdReqPre(0).Enabled = True
            cmdReqPre(2).Enabled = True
            cmdReqPre(3).Enabled = True
        End If
        'Cargar información del Detalle
        Set rs = clsDReq.CargaReqDetalle(ReqDetUnRegistroTramiteUlt, clsDGnral.GetnMovNro(sReqNro))
        If rs.RecordCount > 0 Then Set fgeBS.Recordset = rs
        Set rs = Nothing
        
        'Cargar información del DetMes
        Set rs = clsDReq.CargaReqDetMes(ReqDetMesUltTraNro, clsDGnral.GetnMovNro(sReqNro))
        If rs.RecordCount > 0 Then
            Set fgeMes.Recordset = rs
            For nCont = 1 To fgeMes.Rows - 1
                fgeMes.TextMatrix(nCont, 0) = nCont
            Next
        End If
        Set rs = Nothing
        
        'Actualiza fgeBSDetMes
        Call fgeBS_OnRowChange(fgeBS.Row, fgeBS.Col)
    End If

End Sub

Private Sub CargaTxtBuscar()
    Dim rsReqTree As ADODB.Recordset
    Set rsReqTree = New ADODB.Recordset
    'Carga los requerimientos para precios
    If psFrmTpo = "1" Then
        'PRECIOS
        Set rsReqTree = clsDReq.CargaRequerimiento(psTpoReq, ReqTodosTraPrecio, "")
    ElseIf psFrmTpo = "2" Then
        'CTAS.CONTABLES
        Set rsReqTree = clsDReq.CargaRequerimiento(psTpoReq, ReqTodosTraCuenta, "")
    ElseIf psFrmTpo = "3" Then
        'APROBACION O RECHAZO
        Set rsReqTree = clsDReq.CargaRequerimiento(psTpoReq, ReqTodosTraAproba, "")
    End If
    
    If rsReqTree.RecordCount > 0 Then
        txtBuscar.EditFlex = True
        txtBuscar.rs = rsReqTree
        txtBuscar.Enabled = True
    Else
        txtBuscar.Enabled = False
    End If
    Set rsReqTree = Nothing
End Sub

Private Sub Limpiar()
    Dim nCont As Integer
    'Carga los requerimientos pendientes del area
    Call CargaTxtBuscar
    'Otros
    cboMoneda.ListIndex = -1
    lblMoneda.Caption = ""
    lblPeriodo.Caption = ""
    rtfDescri(0).Text = ""
    rtfDescri(1).Text = ""
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
End Sub

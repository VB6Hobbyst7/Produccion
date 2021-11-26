VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmLogSelEntBase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de la Entrega de Bases"
   ClientHeight    =   5700
   ClientLeft      =   1860
   ClientTop       =   2025
   ClientWidth     =   9255
   Icon            =   "frmLogSelEntBase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab sstSeleccion 
      Height          =   4305
      Left            =   105
      TabIndex        =   7
      Top             =   750
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   7594
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Detalle"
      TabPicture(0)   =   "frmLogSelEntBase.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblEtiqueta(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblEtiqueta(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblEtiqueta(8)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblAdqNro"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblEtiqueta(7)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblResNro"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblResFec"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblCostoBase"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line1(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line1(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fgeBS"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Postores"
      TabPicture(1)   =   "frmLogSelEntBase.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "rtfObserva"
      Tab(1).Control(1)=   "cmdPostor(1)"
      Tab(1).Control(2)=   "cmdPostor(0)"
      Tab(1).Control(3)=   "fgePostor"
      Tab(1).Control(4)=   "lblObserva"
      Tab(1).Control(5)=   "lblEtiqueta(1)"
      Tab(1).ControlCount=   6
      Begin RichTextLib.RichTextBox rtfObserva 
         Height          =   2970
         Left            =   -70335
         TabIndex        =   21
         Top             =   750
         Visible         =   0   'False
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   5239
         _Version        =   393217
         Enabled         =   -1  'True
         MaxLength       =   4000
         TextRTF         =   $"frmLogSelEntBase.frx":0342
      End
      Begin VB.CommandButton cmdPostor 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   -70185
         TabIndex        =   9
         Top             =   3810
         Width           =   1155
      End
      Begin VB.CommandButton cmdPostor 
         Caption         =   "&Agregar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   -71745
         TabIndex        =   8
         Top             =   3810
         Width           =   1155
      End
      Begin Sicmact.FlexEdit fgePostor 
         Height          =   2955
         Left            =   -74670
         TabIndex        =   10
         Top             =   765
         Width           =   8460
         _ExtentX        =   14923
         _ExtentY        =   5212
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Código-Nombre-Observacion"
         EncabezadosAnchos=   "400-1700-4500-0"
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
         ColumnasAEditar =   "X-1-X-X"
         ListaControles  =   "0-1-0-0"
         EncabezadosAlineacion=   "C-L-L-L"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbFormatoCol    =   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   285
         TipoBusPersona  =   1
      End
      Begin Sicmact.FlexEdit fgeBS 
         Height          =   3330
         Left            =   2445
         TabIndex        =   14
         Top             =   765
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   5874
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-cBSCod-Bien/Servicio-Unidad-Cantidad-PrecioProm-Sub Total"
         EncabezadosAnchos=   "400-0-2000-650-900-900-1000"
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
      Begin VB.Label lblObserva 
         Caption         =   "Observación"
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
         Left            =   -70125
         TabIndex        =   22
         Top             =   495
         Width           =   1140
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   2325
         X2              =   2325
         Y1              =   540
         Y2              =   4065
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2310
         X2              =   2310
         Y1              =   525
         Y2              =   4065
      End
      Begin VB.Label lblCostoBase 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   270
         TabIndex        =   20
         Top             =   2235
         Width           =   1395
      End
      Begin VB.Label lblResFec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   285
         TabIndex        =   19
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label lblResNro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   270
         TabIndex        =   18
         Top             =   735
         Width           =   1875
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Costo bases"
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
         Index           =   7
         Left            =   300
         TabIndex        =   17
         Top             =   1995
         Width           =   1170
      End
      Begin VB.Label lblAdqNro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3945
         TabIndex        =   16
         Top             =   450
         Width           =   2760
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Adquisición :"
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
         Left            =   2670
         TabIndex        =   15
         Top             =   495
         Width           =   1230
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Resolución"
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
         Left            =   300
         TabIndex        =   13
         Top             =   495
         Width           =   1110
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Fecha"
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
         Left            =   300
         TabIndex        =   12
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Postores "
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
         Left            =   -74535
         TabIndex        =   11
         Top             =   495
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   7560
      TabIndex        =   3
      Top             =   5190
      Width           =   1305
   End
   Begin VB.CommandButton cmdAdq 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   1
      Left            =   4935
      TabIndex        =   2
      Top             =   5190
      Width           =   1290
   End
   Begin VB.CommandButton cmdAdq 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   0
      Left            =   3015
      TabIndex        =   1
      Top             =   5190
      Width           =   1290
   End
   Begin Sicmact.TxtBuscar txtSelNro 
      Height          =   285
      Left            =   1155
      TabIndex        =   0
      Top             =   375
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TipoBusqueda    =   2
      sTitulo         =   ""
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   -30
      Top             =   5235
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label lblAreaDes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1170
      TabIndex        =   6
      Top             =   60
      Width           =   3705
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
      Left            =   330
      TabIndex        =   5
      Top             =   105
      Width           =   555
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
      Index           =   5
      Left            =   300
      TabIndex        =   4
      Top             =   420
      Width           =   870
   End
End
Attribute VB_Name = "frmLogSelEntBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim psFrmTpo As String
Dim pnRowAnt As Integer

Public Sub Inicio(ByVal psFormTpo As String)
psFrmTpo = psFormTpo
Me.Show 1
End Sub

Private Sub cmdAdq_Click(Index As Integer)
    Dim clsDMov As DLogMov
    Dim clsDGnral As DLogGeneral
    Dim sSelNro As String, sSelTraNro As String, sPersCod As String
    Dim sObserva As String, sActualiza As String
    Dim nCont As Integer, nResult As Integer
    Select Case Index
        Case 0:
            'CANCELAR
            If MsgBox("¿ Estás seguro de cancelar toda la operación ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Call Limpiar
                txtSelNro.Text = ""
                fgePostor.lbEditarFlex = False
                cmdPostor(0).Enabled = False
                cmdPostor(1).Enabled = False
                cmdAdq(0).Enabled = False
                cmdAdq(1).Enabled = False
                
                Call CargaTxtSelNro
            End If
        Case 1:
            'GRABAR
            If psFrmTpo = "1" Then
                'REGISTRO DE POSTORES
                For nCont = 1 To fgePostor.Rows - 1
                    If fgePostor.TextMatrix(nCont, 1) = "" Then
                        MsgBox "Falta determinar el postor(es)", vbInformation, " Aviso"
                        Exit Sub
                    End If
                Next
                If MsgBox("¿ Estás seguro de Grabar estos Postores ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                    sSelNro = txtSelNro.Text
                    
                    Set clsDGnral = New DLogGeneral
                    sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDGnral = Nothing
                    
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    
                    'Grabación de MOV - MOVREF
                    clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", Trim(Str(gLogSelEstadoRegBase))
                    clsDMov.InsertaMovRef sSelTraNro, sSelNro
                    
                    'Actualiza LogSeleccion
                    clsDMov.ActualizaSeleccion sSelNro, gdFecSis, "", "", "", _
                        sActualiza, gLogSelEstadoRegBase
    
                    clsDMov.EliminaSelPostor sSelNro
                    
                    For nCont = 1 To fgePostor.Rows - 1
                        sPersCod = fgePostor.TextMatrix(nCont, 1)
                        clsDMov.InsertaSelPostor sSelNro, sPersCod, sActualiza
                    Next
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        fgePostor.lbEditarFlex = False
                        cmdPostor(0).Enabled = False
                        cmdPostor(1).Enabled = False
                        cmdAdq(0).Enabled = False
                        cmdAdq(1).Enabled = False
                        Call CargaTxtSelNro
                    Else
                        MsgBox "Error al grabar la información", vbInformation, " Aviso "
                    End If
                End If
            ElseIf psFrmTpo = "2" Then
                'OBSERVACION DE POSTORES
                Call fgePostor_OnRowChange(fgePostor.Row, fgePostor.Col)
                If MsgBox("¿ Estás seguro de Grabar estas Observaciones ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                    sSelNro = txtSelNro.Text
                    
                    Set clsDGnral = New DLogGeneral
                    sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDGnral = Nothing
                    
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    
                    'Grabación de MOV - MOVREF
                    clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", Trim(Str(gLogSelEstadoRegBase))
                    clsDMov.InsertaMovRef sSelTraNro, sSelNro
                    
                    For nCont = 1 To fgePostor.Rows - 1
                        sPersCod = fgePostor.TextMatrix(nCont, 1)
                        sObserva = fgePostor.TextMatrix(nCont, 3)
                        clsDMov.ActualizaSelPostor sSelNro, sPersCod, sObserva, sActualiza
                    Next
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        rtfObserva.Locked = True
                        cmdAdq(0).Enabled = False
                        cmdAdq(1).Enabled = False
                        Call CargaTxtSelNro
                    Else
                        MsgBox "Error al grabar la información", vbInformation, " Aviso "
                    End If
                End If
            End If
        Case Else
            MsgBox "Opción no reconocida", vbInformation, "Aviso"
    End Select
End Sub

Private Sub cmdPostor_Click(Index As Integer)
    Dim nBSRow As Integer
    Dim nResult As Integer
    'Botones de comandos del detalle de bienes/servicios
    If Index = 0 Then

        'Agregar en Flex
        fgePostor.AdicionaFila
        fgePostor.SetFocus
    ElseIf Index = 1 Then
        'Eliminar en Flex
        nBSRow = fgePostor.Row
        If MsgBox("¿ Estás seguro de eliminar " & fgePostor.TextMatrix(nBSRow, 2) & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            fgePostor.EliminaFila nBSRow
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub fgePostor_OnRowChange(pnRow As Long, pnCol As Long)
If fgePostor.TextMatrix(1, 1) <> "" Then
If psFrmTpo = "2" Then
    If pnRowAnt > 0 Then fgePostor.TextMatrix(pnRowAnt, 3) = rtfObserva.Text
    rtfObserva.Text = fgePostor.TextMatrix(pnRow, 3)
    pnRowAnt = pnRow
End If
End If
End Sub

Private Sub Form_Load()
    Call CentraForm(Me)
    'Carga información de la relación usuario-area
    Usuario.Inicio gsCodUser
    If Len(Usuario.AreaCod) = 0 Then
        MsgBox "Usuario no determinado", vbInformation, "Aviso"
        Exit Sub
    End If
    lblAreaDes.Caption = Usuario.AreaNom
    
    If psFrmTpo = "1" Then
        'REGISTRO DE POSTORES
        Me.Caption = "Registro de la Entrega de Bases"
    ElseIf psFrmTpo = "2" Then
        'OBSERVACIONES DE POSTORES
        Me.Caption = "Registro de Observaciones de Bases"
        fgePostor.EncabezadosAnchos = "400-0-3500-0"
        fgePostor.Width = fgePostor.Width - 4200
        lblObserva.Visible = True
        rtfObserva.Visible = True
        cmdPostor(0).Visible = False
        cmdPostor(1).Visible = False
    Else
        MsgBox "Tipo Formulario no reconocido", vbInformation, " Aviso"
        Exit Sub
    End If
    Call CargaTxtSelNro
End Sub

Private Sub txtSelNro_EmiteDatos()
    Dim clsDAdq As DLogAdquisi
    Dim rs As ADODB.Recordset
    Dim sSelNro As String
    
    If txtSelNro.Ok = False Then Exit Sub
    
    Set clsDAdq = New DLogAdquisi
    Set rs = New ADODB.Recordset
    
    Call Limpiar
    cmdAdq(0).Enabled = True
    cmdAdq(1).Enabled = True
    If psFrmTpo = "1" Then
        fgePostor.lbEditarFlex = True
        cmdPostor(0).Enabled = True
        cmdPostor(1).Enabled = True
    ElseIf psFrmTpo = "2" Then
        rtfObserva.Locked = False
    End If
    sSelNro = txtSelNro.Text
    Set rs = clsDAdq.CargaSeleccion(SelUnRegistro, sSelNro)
    If rs.RecordCount > 0 Then
        With rs
            lblResFec.Caption = Format(!dLogSelRes, "dd/mm/yyyy")
            lblResNro.Caption = !cLogSelResNro
            lblCostoBase.Caption = Format(!nLogSelCostoBase, "#0.0")
            lblAdqNro.Caption = !cLogAdqNro
            
            'Muestra detalle de Bienes/Servicios
            Set rs = clsDAdq.CargaAdqDetalle(AdqDetUnRegistro, !cLogAdqNro)
            If rs.RecordCount > 0 Then
                Set fgeBS.Recordset = rs
                fgeBS.AdicionaFila
                fgeBS.BackColorRow &HC0FFFF
                fgeBS.TextMatrix(fgeBS.Row, 2) = "T O T A L  R E F E R E N C I A L"
                fgeBS.TextMatrix(fgeBS.Row, 6) = Format(fgeBS.SumaRow(6), "#,##0.00")
            End If
            
            'Muestra postores anteriores
            Set rs = clsDAdq.CargaSelPostor(sSelNro)
            If rs.RecordCount > 0 Then
                Set fgePostor.Recordset = rs
                pnRowAnt = 0
                Call fgePostor_OnRowChange(fgePostor.Row, fgePostor.Col)
            End If
        End With
    End If
    
    Set rs = Nothing
    Set clsDAdq = Nothing
    
End Sub

Private Sub Limpiar()
    lblResNro.Caption = ""
    lblResFec.Caption = ""
    lblCostoBase.Caption = ""
    lblAdqNro.Caption = ""
    fgeBS.Clear
    fgeBS.FormaCabecera
    fgeBS.Rows = 2
    fgePostor.Clear
    fgePostor.FormaCabecera
    fgePostor.Rows = 2
    If psFrmTpo = "2" Then
        rtfObserva.Text = ""
    End If
End Sub

Private Sub CargaTxtSelNro()
    Dim clsDAdq As DLogAdquisi
    Dim rs As ADODB.Recordset
    Set clsDAdq = New DLogAdquisi
    Set rs = New ADODB.Recordset
    If psFrmTpo = "1" Then
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, "", gLogSelEstadoRegBase, gLogSelEstadoCotizacion)
    ElseIf psFrmTpo = "2" Then
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, "", gLogSelEstadoRegBase)
    End If
    If rs.RecordCount > 0 Then
        txtSelNro.rs = rs
    Else
        txtSelNro.Enabled = False
    End If
    Set rs = Nothing
    Set clsDAdq = Nothing
End Sub

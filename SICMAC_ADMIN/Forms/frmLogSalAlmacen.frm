VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogSalAlmacen 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12330
   Icon            =   "frmLogSalAlmacen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   12330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAsiento 
      Caption         =   "A&siento Cnt"
      Height          =   390
      Left            =   10200
      TabIndex        =   26
      Top             =   5340
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CheckBox chkConStock 
      Appearance      =   0  'Flat
      Caption         =   "Solo Bienes con stock mayor a LO PEDIDO"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4155
      TabIndex        =   25
      Top             =   5415
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   390
      Left            =   30
      TabIndex        =   10
      Top             =   5325
      Width           =   1020
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   390
      Left            =   2205
      TabIndex        =   9
      Top             =   5340
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   1110
      TabIndex        =   8
      Top             =   5340
      Width           =   1020
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   11280
      TabIndex        =   7
      Top             =   5340
      Width           =   1020
   End
   Begin MSMask.MaskEdBox mskFecha 
      Height          =   285
      Left            =   10995
      TabIndex        =   11
      Top             =   450
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin Sicmact.TxtBuscar txtGuiaRemInt 
      Height          =   300
      Left            =   1245
      TabIndex        =   12
      Top             =   795
      Width           =   1395
      _extentx        =   2461
      _extenty        =   529
      appearance      =   0
      appearance      =   0
      font            =   "frmLogSalAlmacen.frx":030A
      appearance      =   0
      stitulo         =   ""
   End
   Begin Sicmact.TxtBuscar txtTransportista 
      Height          =   285
      Left            =   1245
      TabIndex        =   19
      Top             =   1395
      Width           =   2475
      _extentx        =   4366
      _extenty        =   503
      appearance      =   0
      appearance      =   0
      font            =   "frmLogSalAlmacen.frx":0336
      appearance      =   0
      tipobusqueda    =   3
      stitulo         =   ""
      enabledtext     =   0
   End
   Begin VB.Frame framCont 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   15
      TabIndex        =   0
      Top             =   1680
      Width           =   12285
      Begin VB.Frame fraComentario 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Comentario"
         ForeColor       =   &H00800000&
         Height          =   750
         Left            =   2295
         TabIndex        =   5
         Top             =   2775
         Width           =   9930
         Begin VB.TextBox txtComentario 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   90
            MaxLength       =   300
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   210
            Width           =   9765
         End
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   345
         Left            =   60
         TabIndex        =   3
         Top             =   2835
         Width           =   1020
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   345
         Left            =   1155
         TabIndex        =   2
         Top             =   2835
         Width           =   1020
      End
      Begin Sicmact.FlexEdit FlexSerie 
         Height          =   2610
         Left            =   7800
         TabIndex        =   1
         Top             =   180
         Width           =   4425
         _extentx        =   7805
         _extenty        =   4604
         cols0           =   7
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Serie-Alm.-idx-IGV-Valor-Serie Real"
         encabezadosanchos=   "300-1700-0-0-700-1200-1500"
         font            =   "frmLogSalAlmacen.frx":0362
         font            =   "frmLogSalAlmacen.frx":038E
         font            =   "frmLogSalAlmacen.frx":03BA
         font            =   "frmLogSalAlmacen.frx":03E6
         font            =   "frmLogSalAlmacen.frx":0412
         fontfixed       =   "frmLogSalAlmacen.frx":043E
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         tipobusqueda    =   2
         columnasaeditar =   "X-1-X-X-X-X-6"
         textstylefixed  =   3
         listacontroles  =   "0-1-0-0-0-0-0"
         encabezadosalineacion=   "C-L-L-C-R-R-L"
         formatosedit    =   "0-0-0-0-2-2-0"
         avanceceldas    =   1
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbbuscaduplicadotext=   -1
         appearance      =   0
         colwidth0       =   300
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin Sicmact.FlexEdit FlexDetalle 
         Height          =   2610
         Left            =   45
         TabIndex        =   4
         Top             =   165
         Width           =   7710
         _extentx        =   13600
         _extenty        =   4604
         cols0           =   9
         highlight       =   1
         allowuserresizing=   1
         rowsizingmode   =   1
         encabezadosnombres=   "#-Codigo-Descripción-Cant. Sol.-Cant. Anten.-Stock-Cta-Id-PP"
         encabezadosanchos=   "300-1200-3000-1000-1000-800-0-0-0"
         font            =   "frmLogSalAlmacen.frx":046C
         font            =   "frmLogSalAlmacen.frx":0498
         font            =   "frmLogSalAlmacen.frx":04C4
         font            =   "frmLogSalAlmacen.frx":04F0
         font            =   "frmLogSalAlmacen.frx":051C
         fontfixed       =   "frmLogSalAlmacen.frx":0548
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         columnasaeditar =   "X-1-X-3-X-X-X-X-X"
         textstylefixed  =   3
         listacontroles  =   "0-1-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-L-L-R-R-R-C-C-C"
         formatosedit    =   "0-0-0-2-2-0-0-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbflexduplicados=   0
         lbbuscaduplicadotext=   -1
         appearance      =   0
         colwidth0       =   300
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
   End
   Begin Sicmact.TxtBuscar txtAlmacen 
      Height          =   300
      Left            =   7065
      TabIndex        =   23
      Top             =   780
      Visible         =   0   'False
      Width           =   1140
      _extentx        =   2011
      _extenty        =   529
      appearance      =   0
      appearance      =   0
      font            =   "frmLogSalAlmacen.frx":0576
      appearance      =   0
   End
   Begin Sicmact.TxtBuscar txtArea 
      Height          =   285
      Left            =   1245
      TabIndex        =   17
      Top             =   1080
      Width           =   2475
      _extentx        =   4366
      _extenty        =   503
      appearance      =   0
      appearance      =   0
      font            =   "frmLogSalAlmacen.frx":05A2
      appearance      =   0
      stitulo         =   ""
   End
   Begin Sicmact.TxtBuscar txtAlmacenDestino 
      Height          =   300
      Left            =   9840
      TabIndex        =   30
      Top             =   1080
      Visible         =   0   'False
      Width           =   1140
      _extentx        =   2011
      _extenty        =   529
      appearance      =   0
      appearance      =   0
      font            =   "frmLogSalAlmacen.frx":05CE
      appearance      =   0
   End
   Begin VB.Label lblAlmDest 
      Caption         =   "Almacen de Destino"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   8280
      TabIndex        =   29
      Top             =   1110
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblAlmacenDestino 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8235
      TabIndex        =   28
      Top             =   1410
      Visible         =   0   'False
      Width           =   4035
   End
   Begin VB.Label lblAlmacenG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   8235
      TabIndex        =   27
      Top             =   795
      Visible         =   0   'False
      Width           =   4035
   End
   Begin VB.Label lblTransportistaG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3750
      TabIndex        =   20
      Top             =   1410
      Width           =   4440
   End
   Begin VB.Label lblTransportista 
      Caption         =   "Responsable :"
      Height          =   240
      Left            =   105
      TabIndex        =   22
      Top             =   1440
      Width           =   1125
   End
   Begin VB.Label lblAreaG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3750
      TabIndex        =   18
      Top             =   1110
      Width           =   4440
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Guia de Remisión : 2001-0000001"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   45
      TabIndex        =   16
      Top             =   60
      Width           =   9495
   End
   Begin VB.Label lblFecha 
      Caption         =   "Fecha :"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   10995
      TabIndex        =   15
      Top             =   255
      Width           =   660
   End
   Begin VB.Label lblGuaRemIntG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2640
      TabIndex        =   14
      Top             =   795
      Width           =   3540
   End
   Begin VB.Label lblGuaRemInt 
      Caption         =   "Requerimiento :"
      Height          =   210
      Left            =   90
      TabIndex        =   13
      Top             =   825
      Width           =   1230
   End
   Begin VB.Label lblAlmacen 
      Caption         =   "Almacen"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   6300
      TabIndex        =   24
      Top             =   825
      Width           =   690
   End
   Begin VB.Label lblArea 
      Caption         =   "Area :"
      Height          =   195
      Left            =   105
      TabIndex        =   21
      Top             =   1125
      Width           =   1125
   End
End
Attribute VB_Name = "frmLogSalAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsOpeCod As String
Dim lnMovNroG As Long
Dim lbIngreso As Boolean
Dim lbConfirma As Boolean
Dim lbAlmacen As Boolean
Dim lnRowFlex As Long
Dim lnOpeDoc As Long
Dim lbExtorno As Boolean
Dim lbRechazo As Boolean
Dim lbReporte As Boolean
Dim lbModifica  As Boolean
Dim lbGrabar As Boolean


Dim lbRepRequerimiento As Boolean

'ARLO 20170126******************
Dim objPista As COMManejador.Pista
Dim lsPalabra, lsAccion As String
'*******************************

Private Sub cmdAgregar_Click()
    If lsOpeCod = gnAlmaSalXAtencion Then
        If Me.txtArea.Text = "" Then
            MsgBox "Debe elegir un area.", vbInformation, "Aviso"
            txtArea.SetFocus
            Exit Sub
        ElseIf Me.txtAlmacen.Text = "" Then
            MsgBox "Debe elegir un almacen.", vbInformation, "Aviso"
            txtAlmacen.SetFocus
            Exit Sub
        End If
    ElseIf Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4) Then
        If Me.txtAlmacen.Text = "" Then
            MsgBox "Debe elegir un almacen.", vbInformation, "Aviso"
            txtAlmacen.SetFocus
            Exit Sub
        End If
        
        If Me.txtTransportista.Text = "" And lsOpeCod <> gnAlmaSalXAjuste Then
            MsgBox "Debe elegir un responsable.", vbInformation, "Aviso"
            txtTransportista.SetFocus
            Exit Sub
        End If
'
'        If lsOpeCod <> gnAlmaSalXAtencion Then 'prueba
'            Me.FlexObjeto.AdicionaFila
'            Me.FlexObjeto.TextMatrix(FlexObjeto.Rows - 1, 1) = "2"
'            Me.FlexObjeto.TextMatrix(FlexObjeto.Rows - 1, 2) = FlexObjeto.Rows - 1
'        End If 'fin prueba
    End If
    If Me.FlexDetalle.TextMatrix(1, 1) <> "" Then
        If FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 1) <> "" Or FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 3) <> "" Then
            Me.FlexDetalle.AdicionaFila , , True
        End If
    Else
        Me.FlexDetalle.AdicionaFila
        FlexDetalle_RowColChange
    End If
    Me.FlexDetalle.SetFocus
End Sub

Private Sub cmdAsiento_Click()
    Dim oAsiento As NContImprimir
    Set oAsiento = New NContImprimir
    Dim lsOpeCodLocal As String
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    Dim lsMovNro As String
    Dim oMov As DMov
    Set oMov = New DMov
    
    If lnMovNroG = 0 Then
        MsgBox "Debe ingresar una Güia de Remisión.", vbInformation, "Aviso"
        Me.txtGuiaRemInt.SetFocus
        Exit Sub
    End If
    
    lsMovNro = oMov.GetcMovNro(lnMovNroG)
    lsOpeCodLocal = GetOpeMov(lnMovNroG)
    
    If lsOpeCodLocal = gnAlmaSalXAtencion Or lsOpeCodLocal = gnAlmaSalXAjuste Then
        oPrevio.Show oAsiento.ImprimeAsientoContable(lsMovNro, 60, 80), Caption, True
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdEliminar_Click()
    Dim i As Integer
    Dim lnEncontrar As Integer
    Dim lnContador As Integer
    
    If MsgBox("Desea Eliminar la fila, si ha incluido numeros de serie para este producto se perderan.", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub

    For i = 1 To CInt(Me.FlexSerie.Rows - 1)
        If FlexSerie.TextMatrix(i, 3) = Me.FlexDetalle.TextMatrix(FlexDetalle.row, 0) Then
            lnContador = lnContador + 1
        End If
    Next i
    
    i = 0
    While lnEncontrar < lnContador
        i = i + 1
        If FlexSerie.TextMatrix(i, 3) = Me.FlexDetalle.TextMatrix(FlexDetalle.row, 0) Then
            Me.FlexSerie.EliminaFila i
            lnEncontrar = lnEncontrar + 1
            i = i - 1
        End If
    Wend
    
    Me.FlexDetalle.EliminaFila Me.FlexDetalle.row
End Sub

Private Sub cmdGrabar_Click()
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    
    If Not Valida Then Exit Sub
    
    Dim i As Integer
    Dim oAlmacen As DMov
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lnItem As Integer
    Dim lsBSCod As String
    Dim lsDocNI As String
    Dim ldFechaOC As Date
    Set oAlmacen = New DMov
    Dim oOpe As DOperaciones
    Set oOpe = New DOperaciones
    Dim lsCtaCnt As String
    Dim oAsiento As NContImprimir
    Dim lsOpeCodLocal As String
    Dim lsCtaCntTemp As String
    Dim lnMontoActivo As Currency
    Dim lsAG As String
    Dim lsTipoMovAlm As String
    Dim lnStock As Double
    Dim lnCostoTotal As Double
    Dim lnPreProm As Double
    Dim lnPreuni As Double
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim lsSQL As String
    
    Dim rsDistrib As ADODB.Recordset
    Dim lnItemDistri As Integer
    Dim lnMontoDistri As Currency
    Dim lnSumaDistri As Currency, lnDistri As Currency
    Dim lnPrimerItem As Integer
    
    '*** PEAC 20111102
    Dim rsAsiento As ADODB.Recordset
    Set rsAsiento = New ADODB.Recordset
    Dim lnMontoDebe As Double
    Dim lnMontoHaber As Double
    Dim lnItenDebe As Double
    Dim lnItenHaber As Double
    Dim lnImporteDebe As Double
    Dim lnImporteHaber As Double
    '*** PEAC
    Dim lbUnicaCtaCont As Boolean 'EJVG20140320
    'PASI20151126 ERS0782015
    Dim MatOrdP As TMatOrdPago
    Dim bExisteOrdP As Boolean
    Dim nItemOrdP As Integer
    Dim nCantSol As Integer
    Dim fsMatOrdPag() As String
    'end PASI***
    'GITU
    If Left(gsOpeCod, 4) = "5911" Then 'Ingresos
        lsTipoMovAlm = "I"
    ElseIf Left(gsOpeCod, 4) = "5912" Then
        lsTipoMovAlm = "S"
    End If
    
    If MsgBox("Desea Grabar los cambios Realizados ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    lbGrabar = True
    If lbModifica Then
        'GRABACION DE MANTENIMIENTO DE SALIDAS
        
        lsDocNI = Right(Me.lblTit, 13)
        
        'SOLO ALMACEN
        oAlmacen.BeginTrans
          'Inserta Mov
          lsMovNro = oAlmacen.GeneraMovNro(CDate(Me.mskFecha), Right(gsCodAge, 2), gsCodUser)
          
          lsOpeCodLocal = GetOpeMov(lnMovNroG)
          lsOpeCod = lsOpeCodLocal
          
          If lsOpeCod = gnAlmaSalXAtencion Or lsOpeCod = gnAlmaSalXAjuste Then   'Confirmacion de Solicitud de Bienes
              oAlmacen.InsertaMov lsMovNro, lsOpeCod, Me.txtComentario.Text
          ElseIf lsOpeCod = gnAlmaReqAreaExt Then                                ' Rechazo
              oAlmacen.InsertaMov lsMovNro, lsOpeCod, Me.txtComentario.Text, gMovEstContabRechazado, gMovFlagDeExtorno
          ElseIf Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4) Then                          ' Salidas
              oAlmacen.InsertaMov lsMovNro, lsOpeCod, Me.txtComentario.Text, gMovEstLogSaleBienAlmacen
          End If
          
          lnMovNro = oAlmacen.GetnMovNro(lsMovNro)
          
          If Me.txtTransportista.Text <> "" Then oAlmacen.InsertaMovGasto lnMovNro, Me.txtTransportista.Text, ""
          
          If lnMovNroG <> 0 Then
             oAlmacen.ActualizaMov lnMovNroG, , gMovEstContabNoContable, gMovFlagEliminado    'Modificado Siempre y cuando no se confirma
             oAlmacen.InsertaMovRef lnMovNro, lnMovNroG
             oAlmacen.InsertaMovRefAnt lnMovNro, lnMovNroG
             oAlmacen.EliminaMovBSSerieparaActualizar lnMovNroG
          End If
          
          'Inserta Documentos
          oAlmacen.InsertaMovDoc lnMovNro, lnOpeDoc, lsDocNI, Format(CDate(mskFecha.Text), gsFormatoFecha)
          
          For i = 1 To Me.FlexDetalle.Rows - 1
            lsBSCod = Me.FlexDetalle.TextMatrix(i, 1)
            If lsOpeCod = gnAlmaSalXAtencion Or lsOpeCod = gnAlmaSalXAjuste Then
                If CCur(Me.FlexDetalle.TextMatrix(i, 3)) > 0 Then
                    
                    If lsOpeCod = gnAlmaSalXAtencion Then
                        oAlmacen.InsertaMovBS lnMovNro, i, Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(i, 1)
                    ElseIf Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4) Then
                        oAlmacen.InsertaMovBS lnMovNro, i, Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(i, 1)
                    Else
                        MsgBox "Ver..."
                    End If
                    oAlmacen.InsertaMovCant lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 3)
                    
                    'If Left(Me.FlexDetalle.TextMatrix(I, 6), 2) <> "18" Then
                    'control por bienes
                    If InStr(1, Me.FlexDetalle.TextMatrix(i, 2), "[S]") <> 0 Then
                        lnMontoActivo = GetValorTotalSalida(i)
                        If lnMontoActivo <> 0 Then
                            oAlmacen.InsertaMovCta lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 6), lnMontoActivo * -1
                        Else
                            oAlmacen.InsertaMovCta lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 6), CCur(Me.FlexDetalle.TextMatrix(i, 3)) * CCur(Me.FlexDetalle.TextMatrix(i, 7) * -1)
                        End If
                    Else
                        oAlmacen.InsertaMovCta lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 6), CCur(Me.FlexDetalle.TextMatrix(i, 3)) * CCur(Me.FlexDetalle.TextMatrix(i, 7) * -1)
                    End If
                        'oALmacen.InsertaMovCta lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 6), CCur(Me.FlexDetalle.TextMatrix(i, 3)) * CCur(Me.FlexDetalle.TextMatrix(i, 7) * -1)
                    'End If
                End If
            Else
                If Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4) Then   'Otras Salidas
                    
                    oAlmacen.InsertaMovCant lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 3)
                    oAlmacen.InsertaMovBS lnMovNro, i, Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(i, 1)
                    oAlmacen.InsertaMovCta lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 6), IIf(Me.FlexDetalle.TextMatrix(i, 3) = "", 0, Me.FlexDetalle.TextMatrix(i, 7))
                Else ' Solcicitud
                    oAlmacen.InsertaMovCant lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 3)
                    oAlmacen.InsertaMovBS lnMovNro, i, Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(i, 1)
                    oAlmacen.InsertaMovCta lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 6), IIf(Me.FlexDetalle.TextMatrix(i, 5) = "", 0, Me.FlexDetalle.TextMatrix(i, 5))
                End If
            End If
            
            'Inserta AreaAgencia
            If Me.txtArea.Text <> "" Then
                oAlmacen.InsertaMovObj lnMovNro, i, 1, "13"
                oAlmacen.InsertaMovObjAgenciaArea lnMovNro, i, 1, IIf(Len(Me.txtArea.Text) = 5, Right(Me.txtArea.Text, 2), ""), Left(Me.txtArea.Text, 3)
            End If
          Next i
          
          For i = 1 To Me.FlexDetalle.Rows - 1
            lsBSCod = Me.FlexDetalle.TextMatrix(i, 1)
            If lsOpeCod = gnAlmaSalXAtencion Or lsOpeCod = gnAlmaSalXAjuste Then
                If CCur(Me.FlexDetalle.TextMatrix(i, 3)) > 0 Then
                    If Left(Me.FlexDetalle.TextMatrix(i, 6), 2) <> "18" Then  'POR DEFINIR
                        lsCtaCnt = oAlmacen.GetOpeCtaCta(lsOpeCod, "", Me.FlexDetalle.TextMatrix(i, 6), Format(Me.txtAlmacen.Text, "00"))
                        lsCtaCntTemp = oAlmacen.GetCtaOpeBS(lsOpeCod, Me.FlexDetalle.TextMatrix(i, 1))
                        If lsCtaCntTemp <> "" Then
                            lsCtaCnt = lsCtaCntTemp
                        End If
                        'If Val(Me.txtAlmacenDestino.Text) <= "9" Then
                        '    lsAG = "0" + Me.txtAlmacenDestino.Text
                        ' Else
                        '    lsAG = Me.txtAlmacenDestino.Text
                        'End If
                        
                        lsCtaCnt = Replace(lsCtaCnt, "AG", GetCtaAreaAge(Left(Me.txtArea.Text, 3), Mid(Me.txtArea.Text, 4, 2)))
                        'lsCtaCnt = Replace(lsCtaCnt, "AG", lsAG)
                        If lsCtaCnt = "" And Left(Me.FlexDetalle.TextMatrix(i, 6), 2) <> "18" Then
                            oAlmacen.RollbackTrans
                            MsgBox "La cuenta contable : " & Me.FlexDetalle.TextMatrix(i, 6) & " no tienen una cuenta referenciada. Ver OPECTACTA....", vbInformation
                            Exit Sub
                        End If
                        'Montos por Serie
                        If InStr(1, Me.FlexDetalle.TextMatrix(i, 2), "[S]") <> 0 Then
                            lnMontoActivo = GetValorTotalSalida(i)
                            If lnMontoActivo <> 0 Then
                                oAlmacen.InsertaMovCta lnMovNro, Me.FlexDetalle.Rows - 1 + i, lsCtaCnt, lnMontoActivo
                            Else
                                oAlmacen.InsertaMovCta lnMovNro, Me.FlexDetalle.Rows - 1 + i, lsCtaCnt, CCur(Me.FlexDetalle.TextMatrix(i, 3)) * CCur(Me.FlexDetalle.TextMatrix(i, 7))
                            End If
                        Else
                            oAlmacen.InsertaMovCta lnMovNro, Me.FlexDetalle.Rows - 1 + i, lsCtaCnt, CCur(Me.FlexDetalle.TextMatrix(i, 3)) * CCur(Me.FlexDetalle.TextMatrix(i, 7))
                        End If
                        'oALmacen.InsertaMovCta lnMovNro, Me.FlexDetalle.Rows - 1 + i, lsCtaCnt, CCur(Me.FlexDetalle.TextMatrix(i, 3)) * CCur(Me.FlexDetalle.TextMatrix(i, 7))
                    End If
                End If
            End If
          Next i
          
         If FlexSerie.TextMatrix(1, 1) <> "" Then
            For i = 1 To Me.FlexSerie.Rows - 1
              If InStr(1, Me.FlexDetalle.TextMatrix(CInt(Me.FlexSerie.TextMatrix(i, 3)), 2), "[S]") <> 0 Then
                If Me.FlexSerie.TextMatrix(i, 5) <> 0 Then
                    oAlmacen.InsertaMovSalidaBSSerie lnMovNro, CInt(Me.FlexSerie.TextMatrix(i, 3)), Me.FlexDetalle.TextMatrix(CInt(Me.FlexSerie.TextMatrix(i, 3)), 1), Me.FlexSerie.TextMatrix(i, 1), Me.FlexSerie.TextMatrix(i, 4), Me.FlexSerie.TextMatrix(i, 5)
                Else
                    oAlmacen.InsertaMovSalidaBSSerie lnMovNro, CInt(Me.FlexSerie.TextMatrix(i, 3)), Me.FlexDetalle.TextMatrix(CInt(Me.FlexSerie.TextMatrix(i, 3)), 1), Me.FlexSerie.TextMatrix(i, 1), Me.FlexSerie.TextMatrix(i, 4), Me.FlexDetalle.TextMatrix(CInt(Me.FlexSerie.TextMatrix(i, 3)), 7)
                End If
                'oALmacen.InsertaMovSalidaBSSerie lnMovNro, CInt(Me.FlexSerie.TextMatrix(i, 3)), Me.FlexDetalle.TextMatrix(CInt(Me.FlexSerie.TextMatrix(i, 3)), 1), Me.FlexSerie.TextMatrix(i, 1)
                ''oAlmacen.InsertaMovBS 'ACTIVO FIJO
              End If
            Next i
          End If
        oAlmacen.CommitTrans
        
        Set oAsiento = New NContImprimir
        If lsOpeCod = gnAlmaSalXAtencion Or lsOpeCod = gnAlmaSalXAjuste Then oPrevio.Show oAsiento.ImprimeAsientoContable(lsMovNro, 60, 80), Caption, True
        
        'ARLO 20170126 ***
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", Me.Caption & " El Requemiento N° : " & Trim(txtGuiaRemInt.Text)
        Set objPista = Nothing
        '***
        cmdImprimir_Click
        Unload Me
        
        
        Exit Sub
        
    End If 'FIN DE MODIFICACION DE MANTENIMIENTO DE SALIDAS
    
    
    If lbRechazo Then 'RECHAZO DE ITEMS
        oAlmacen.BeginTrans
        'Inserta Mov
        lsMovNro = oAlmacen.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        oAlmacen.InsertaMov lsMovNro, lsOpeCod, Me.txtComentario.Text, gMovEstContabRechazado, gMovFlagVigente
        lnMovNro = oAlmacen.GetnMovNro(lsMovNro)
        
        If Me.txtTransportista.Text <> "" Then oAlmacen.InsertaMovGasto lnMovNro, Me.txtTransportista.Text, ""
        
        If lnMovNroG <> 0 Then
           oAlmacen.ActualizaMov lnMovNroG, , gMovEstContabRechazado, gMovFlagExtornado   'Modificado Siempre y cuando no se confirma
           oAlmacen.InsertaMovRef lnMovNro, lnMovNroG
           oAlmacen.EliminaMovBSSerieparaActualizar lnMovNroG
           oAlmacen.ModificaFlagDocumento Me.txtGuiaRemInt.Text, lnOpeDoc, 1
        End If
        oAlmacen.CommitTrans
        
        'ARLO 20170126 ***
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Se ha Rechazado el sobrante de pedido del Requemiento N° : " & Trim(txtGuiaRemInt.Text)
        Set objPista = Nothing
        '***
        
        Unload Me
        Exit Sub
    End If
    
    If lbExtorno Or lsOpeCod = gnAlmaReqAreaExt Then
        oAlmacen.BeginTrans
        'Inserta Mov
        lsMovNro = oAlmacen.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        oAlmacen.InsertaMov lsMovNro, lsOpeCod, Me.txtComentario.Text, gMovEstContabRechazado, gMovFlagDeExtorno
        lnMovNro = oAlmacen.GetnMovNro(lsMovNro)
        
        If Me.txtTransportista.Text <> "" Then oAlmacen.InsertaMovGasto lnMovNro, Me.txtTransportista.Text, ""
        'Extorno Gitu
        If lsOpeCod <> gnAlmaReqAreaExt Then
            Set rs = oOpe.GetMaestroAlmacen(Me.txtAlmacen.Text)
      
            For i = 1 To Me.FlexDetalle.Rows - 1
                lsBSCod = Me.FlexDetalle.TextMatrix(i, 1)
                'If CCur(Me.FlexDetalle.TextMatrix(i, 4)) > 0 Then
                    rs.Filter = "cBSCod = '" & Me.FlexDetalle.TextMatrix(i, 1) & "'"
                    If rs.RecordCount <> 0 Then
                        lnStock = rs!nStock + Me.FlexDetalle.TextMatrix(i, 3)
                        lnPreProm = rs!nPreprom
                        lnPreuni = rs!nPreUni
                        lnCostoTotal = rs!nCostoTotal + (Me.FlexDetalle.TextMatrix(i, 3) * lnPreProm)
                        
                        oAlmacen.ActualizaMasterAlmacen lsBSCod, lnStock, lnPreuni, lnPreProm, lnCostoTotal, txtAlmacen.Text, gdFecSis, lsTipoMovAlm, lsMovNro
                        oAlmacen.EliminaKardexAlmacen lsBSCod, lnMovNroG, Me.txtAlmacen.Text
                    End If
                'End If
            Next i
        End If
        'Fin Gitu
        If lnMovNroG <> 0 Then
           oAlmacen.ActualizaMov lnMovNroG, , gMovEstContabNoContable, gMovFlagModificado  'Modificado Siempre y cuando no se confirma
           oAlmacen.InsertaMovRef lnMovNro, lnMovNroG
           oAlmacen.EliminaMovBSSerieparaActualizar lnMovNroG
        End If
        oAlmacen.CommitTrans
        
        'ARLO 20170126 ***
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Se ha Extornado el Requemiento N° : " & Trim(txtGuiaRemInt.Text)
        Set objPista = Nothing
        '***
        
        Unload Me
        Exit Sub
    End If
    
    
    If lsOpeCod = gnAlmaReqAreaReg Or Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4) Then
       lsDocNI = oAlmacen.GeneraDocNro(lnOpeDoc, gMonedaExtranjera, Year(gdFecSis))
    Else
       lsDocNI = Right(Me.lblTit, 13)
    End If
    
    'PASI201511226 ERS0782015
    If gsOpeCod = "591201" Then
    bExisteOrdP = False
    For i = 1 To Me.FlexDetalle.Rows - 1
        If Me.FlexDetalle.TextMatrix(i, 1) = "11102002005" Or Me.FlexDetalle.TextMatrix(i, 1) = "11102002006" Then
            bExisteOrdP = True
            nItemOrdP = i
            nCantSol = Me.FlexDetalle.TextMatrix(i, 4)
        End If
    Next
    If bExisteOrdP Then
        MsgBox "Se requiere asignar el rango de Ordenes de Pago.", vbInformation, "Aviso"
        fsMatOrdPag = frmAsignaOrdPago.Inicio(nItemOrdP, lblTransportistaG.Caption, lblAreaG.Caption, nCantSol)
        If Len(fsMatOrdPag(1, 1)) = 0 Then Exit Sub
        If Not Len(fsMatOrdPag(1, 1)) = 0 Then
            MatOrdP.nItem = fsMatOrdPag(1, 1)
            MatOrdP.nCantidad = fsMatOrdPag(2, 1)
            MatOrdP.nRangoIni = fsMatOrdPag(3, 1)
            MatOrdP.nRangoFin = fsMatOrdPag(4, 1)
            MatOrdP.cGlosa = fsMatOrdPag(5, 1)
        End If
    End If
    End If
    'end PASI*****
    oAlmacen.BeginTrans
      'Inserta Mov
      lsMovNro = oAlmacen.GeneraMovNro(CDate(mskFecha.Text), Right(gsCodAge, 2), gsCodUser)
      
      If lsOpeCod = gnAlmaSalXAtencion Or lsOpeCod = gnAlmaSalXAjuste Or lsOpeCod = gnAlmaSalXTransferenciaOrigen Then    'Confirmacion de Solicitud de Bienes
        oAlmacen.InsertaMov lsMovNro, lsOpeCod, Me.txtComentario.Text
      ElseIf lsOpeCod = gnAlmaReqAreaExt Then                                ' Rechazo
        oAlmacen.InsertaMov lsMovNro, lsOpeCod, Me.txtComentario.Text, gMovEstContabRechazado, gMovFlagDeExtorno
      ElseIf Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4) Then                          ' Salidas
        oAlmacen.InsertaMov lsMovNro, lsOpeCod, Me.txtComentario.Text, gMovEstLogSaleBienAlmacen
      Else                                                              ' Mantenimiento e ingreso
        oAlmacen.InsertaMov lsMovNro, lsOpeCod, Me.txtComentario.Text, gMovEstContabNoContable
      End If
      
      lnMovNro = oAlmacen.GetnMovNro(lsMovNro)
      
      If Me.txtTransportista.Text <> "" Then oAlmacen.InsertaMovGasto lnMovNro, Me.txtTransportista.Text, ""
      
      If lnMovNroG <> 0 Then
         oAlmacen.ActualizaMov lnMovNroG, , 5  'Modificado Siempre y cuando no se confirma
         oAlmacen.InsertaMovRef lnMovNro, lnMovNroG
         oAlmacen.EliminaMovBSSerieparaActualizar lnMovNroG
      End If
      
      'Inserta Documentos
      oAlmacen.InsertaMovDoc lnMovNro, lnOpeDoc, lsDocNI, Format(CDate(mskFecha.Text), gsFormatoFecha)
      
      'Gitu
      Set rs = oOpe.GetMaestroAlmacen(Me.txtAlmacen.Text)
      
      For i = 1 To Me.FlexDetalle.Rows - 1
        lsBSCod = Me.FlexDetalle.TextMatrix(i, 1)
        If lsOpeCod = gnAlmaSalXAtencion Or lsOpeCod = gnAlmaSalXAjuste Or lsOpeCod = gnAlmaSalXTransferenciaOrigen Then
            If CCur(Me.FlexDetalle.TextMatrix(i, 4)) > 0 Then
                rs.Filter = "cBSCod = '" & Me.FlexDetalle.TextMatrix(i, 1) & "'"
                If rs.RecordCount <> 0 Then
                    lnStock = rs!nStock - Me.FlexDetalle.TextMatrix(i, 4)
                    lnPreProm = rs!nPreprom
                    lnPreuni = rs!nPreUni
                    lnCostoTotal = rs!nCostoTotal - (Me.FlexDetalle.TextMatrix(i, 4) * lnPreProm)
                        
                    oAlmacen.ActualizaMasterAlmacen lsBSCod, lnStock, lnPreuni, lnPreProm, lnCostoTotal, txtAlmacen.Text, gdFecSis, lsTipoMovAlm, lsMovNro
                    oAlmacen.InsertaKardexAlmacen lnMovNro, i, lsBSCod, lsDocNI, Me.FlexDetalle.TextMatrix(i, 4), lnPreuni, lnStock, lnPreProm, lnCostoTotal, txtAlmacen.Text, lsTipoMovAlm, lsMovNro
                End If
            End If
        End If
      Next i
      'Fin Gitu
      
      For i = 1 To Me.FlexDetalle.Rows - 1
        lsBSCod = Me.FlexDetalle.TextMatrix(i, 1)
        If lsOpeCod = gnAlmaSalXAtencion Or lsOpeCod = gnAlmaSalXAjuste Or lsOpeCod = gnAlmaSalXTransferenciaOrigen Then
            If CCur(Me.FlexDetalle.TextMatrix(i, 4)) > 0 Then
                
                If lsOpeCod = gnAlmaSalXAtencion Then
                    oAlmacen.InsertaMovBS lnMovNro, i, Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(i, 1), Me.txtAlmacen.Text
                ElseIf Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4) Then
                    oAlmacen.InsertaMovBS lnMovNro, i, Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(i, 1), Me.txtAlmacen.Text
                Else
                    MsgBox "Ver..."
                End If
                oAlmacen.InsertaMovCant lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 4)
                'oAlmacen.InsertaMovCant lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 4)
                If Left(Me.FlexDetalle.TextMatrix(i, 6), 2) <> "18" Then
                    If InStr(1, Me.FlexDetalle.TextMatrix(i, 2), "[S]") <> 0 Then
                        lnMontoActivo = GetValorTotalSalida(i)
                        If lnMontoActivo <> 0 Then
                            oAlmacen.InsertaMovCta lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 6), lnMontoActivo * -1
                        Else
                            oAlmacen.InsertaMovCta lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 6), CCur(Me.FlexDetalle.TextMatrix(i, 4)) * CCur(Me.FlexDetalle.TextMatrix(i, 7) * -1)
                        End If
                    Else
                        oAlmacen.InsertaMovCta lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 6), CCur(Me.FlexDetalle.TextMatrix(i, 4)) * CCur(Me.FlexDetalle.TextMatrix(i, 7) * -1)
                    End If
                End If
            End If
        Else
            If Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4) Then   'Otras Salidas
                
                oAlmacen.InsertaMovBS lnMovNro, i, Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(i, 1), Me.txtAlmacen.Text
                oAlmacen.InsertaMovCant lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 4)
                oAlmacen.InsertaMovCta lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 6), CCur(IIf(Me.FlexDetalle.TextMatrix(i, 4) = "", 0, Me.FlexDetalle.TextMatrix(i, 4)) * CCur(Me.FlexDetalle.TextMatrix(i, 7)) * -1)
            Else ' Solcicitud
                
                oAlmacen.InsertaMovBS lnMovNro, i, Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(i, 1), Me.txtAlmacen.Text
                oAlmacen.InsertaMovCant lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 3)
                oAlmacen.InsertaMovCta lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 6), IIf(Me.FlexDetalle.TextMatrix(i, 5) = "", 0, Me.FlexDetalle.TextMatrix(i, 5))
                'oALmacen.InsertaMovAgencia lnMovNro, I, Format(txtAlmacen.Text, "00")
            End If
        End If
        
        'Inserta AreaAgencia
        If Me.txtArea.Text <> "" Then
            oAlmacen.InsertaMovObj lnMovNro, i, 1, "13"
            oAlmacen.InsertaMovObjAgenciaArea lnMovNro, i, 1, IIf(Len(Me.txtArea.Text) = 5, Right(Me.txtArea.Text, 2), ""), Left(Me.txtArea.Text, 3)
        End If
      Next i
      
      
      lnItemDistri = Me.FlexDetalle.Rows - 1 '*** PEAC 20101013
      
      For i = 1 To Me.FlexDetalle.Rows - 1
        lsBSCod = Me.FlexDetalle.TextMatrix(i, 1)
        If lsOpeCod = gnAlmaSalXAtencion Or lsOpeCod = gnAlmaSalXAjuste Or lsOpeCod = gnAlmaSalXTransferenciaOrigen Then
            If CCur(Me.FlexDetalle.TextMatrix(i, 4)) > 0 Then
                lsCtaCntTemp = ""
                lsCtaCnt = oAlmacen.GetOpeCtaCta(lsOpeCod, "", Me.FlexDetalle.TextMatrix(i, 6), Format(Me.txtAlmacen.Text, "00"))
                lsCtaCntTemp = oAlmacen.GetCtaOpeBS(lsOpeCod, Me.FlexDetalle.TextMatrix(i, 1))
                If lsCtaCntTemp <> "" Then
                    lsCtaCnt = lsCtaCntTemp
                End If
                If lsOpeCod = gnAlmaSalXTransferenciaOrigen Then
                    If Val(Me.txtAlmacenDestino.Text) <= "9" Then
                        lsAG = "0" + Me.txtAlmacenDestino.Text
                     Else
                        lsAG = Me.txtAlmacenDestino.Text
                    End If
                    lsCtaCnt = Replace(lsCtaCnt, "AG", lsAG)
                 Else
                    lsCtaCnt = Replace(lsCtaCnt, "AG", oAlmacen.GetCtaAreaAge(Left(Me.txtArea.Text, 3), Mid(Me.txtArea.Text, 4, 2)))
                End If
                                
                If lsCtaCnt = "" And Left(Me.FlexDetalle.TextMatrix(i, 6), 2) <> "18" Then
                    oAlmacen.RollbackTrans
                    MsgBox "La cuenta contable : " & Me.FlexDetalle.TextMatrix(i, 6) & " no tienen una cuenta referenciada. Ver OPECTACTA....", vbInformation
                    Exit Sub
                End If
                If InStr(1, Me.FlexDetalle.TextMatrix(i, 2), "[S]") <> 0 Then
                    lnMontoActivo = GetValorTotalSalida(i)
                    If lnMontoActivo <> 0 Then
                        oAlmacen.InsertaMovCta lnMovNro, Me.FlexDetalle.Rows - 1 + i, lsCtaCnt, lnMontoActivo
                    Else
                        oAlmacen.InsertaMovCta lnMovNro, Me.FlexDetalle.Rows - 1 + i, lsCtaCnt, CCur(Me.FlexDetalle.TextMatrix(i, 4)) * CCur(Me.FlexDetalle.TextMatrix(i, 7))
                    End If
                Else
                lbUnicaCtaCont = oAlmacen.EsUnicaCtaCont(FlexDetalle.TextMatrix(i, 1)) 'EJVG20140320
                    '*********************** PEAC 20101013 --- DISTRIBUCION
                    If Len(Me.txtArea.Text) > 3 And Trim(Str(CInt(Right(Me.txtArea.Text, 2)))) <> Me.txtAlmacen.Text Then
                        'lsCtaCnt = Left(Me.FlexDetalle.TextMatrix(I, 6), Len(Me.FlexDetalle.TextMatrix(I, 6)) - 2) + Right(Me.txtArea.Text, 2)
                        lsCtaCnt = IIf(lbUnicaCtaCont, FlexDetalle.TextMatrix(i, 6), Left(FlexDetalle.TextMatrix(i, 6), Len(FlexDetalle.TextMatrix(i, 6)) - 2) + Right(txtArea.Text, 2)) 'EJVG20140320
                        oAlmacen.InsertaMovCta lnMovNro, Me.FlexDetalle.Rows - 1 + i, lsCtaCnt, CCur(Me.FlexDetalle.TextMatrix(i, 4)) * CCur(Me.FlexDetalle.TextMatrix(i, 7))
                        oAlmacen.InsertaMovBS lnMovNro, Me.FlexDetalle.Rows - 1 + i, 0, FlexDetalle.TextMatrix(i, 1), Val(Right(Me.txtArea.Text, 2)) 'EJVG20140320
                        oAlmacen.InsertaMovCant lnMovNro, Me.FlexDetalle.Rows - 1 + i, Me.FlexDetalle.TextMatrix(i, 4)
                    Else
                    
                    If lsOpeCod = "591201" And Me.txtAlmacen.Text = "1" And Left(Me.txtArea.Text, 3) <> "026" Then

                        Set rsDistrib = New ADODB.Recordset
                        lsSQL = " exec stp_sel_AgenciaPorcentajeGastos "
                        Set rsDistrib = oAlmacen.CargaRecordSet(lsSQL)
                        
                        lnMontoDistri = Round(CCur(Me.FlexDetalle.TextMatrix(i, 4)) * CCur(Me.FlexDetalle.TextMatrix(i, 7)), 2)
                        lnSumaDistri = 0
                        'lnPrimerItem = 0
                        Do While Not rsDistrib.EOF
                            lnItemDistri = lnItemDistri + 1
                            
'                            If lnPrimerItem = 0 Then
'                                lnPrimerItem = lnItemDistri
'                            End If
                            
                            lnDistri = Round(lnMontoDistri * rsDistrib!nAgePorcentaje / 100, 2)
                            
                            oAlmacen.InsertaMovCta lnMovNro, lnItemDistri, Left(lsCtaCnt, Len(lsCtaCnt) - 2) + rsDistrib!cAgeCod, lnDistri
                            lnSumaDistri = lnSumaDistri + lnDistri
                            rsDistrib.MoveNext
                        Loop
                        
                        If lnSumaDistri > lnMontoDistri Then
                            Call oAlmacen.ActualizaMovCta(lnMovNro, lnItemDistri, , lnDistri - (lnSumaDistri - lnMontoDistri))
                        ElseIf lnMontoDistri > lnSumaDistri Then
                            Call oAlmacen.ActualizaMovCta(lnMovNro, lnItemDistri, , lnDistri + (lnMontoDistri - lnSumaDistri))
                        End If
                        
                        RSClose rsDistrib

                        Set rsDistrib = Nothing
                    Else
                        oAlmacen.InsertaMovCta lnMovNro, Me.FlexDetalle.Rows - 1 + i, lsCtaCnt, CCur(Me.FlexDetalle.TextMatrix(i, 4)) * CCur(Me.FlexDetalle.TextMatrix(i, 7))
                    End If
                    
                    End If
                    'oALmacen.InsertaMovCta lnMovNro, Me.FlexDetalle.Rows - 1 + i, lsCtaCnt, CCur(Me.FlexDetalle.TextMatrix(i, 4)) * CCur(Me.FlexDetalle.TextMatrix(i, 7))
                    
                    '******************************* FIN PEAC
                End If
            End If
        End If
      Next i
      
    'PASI20151127 ERS0782015
    If Not MatOrdP.nCantidad = 0 Then
        oAlmacen.InsertaMovBSOrdPago lnMovNro, MatOrdP.nItem, MatOrdP.nCantidad, MatOrdP.nRangoIni, MatOrdP.nRangoFin, MatOrdP.cGlosa
    End If
    'end PASI
      
    '*** PEAC 20111102 - verifica descuadres en el asiento
    Set rsAsiento = oAlmacen.CargaRecordSet("select nMovItem,cCtaContCod,case when nMovImporte > 0 then nMovImporte else 0 end ndebe,case when nMovImporte < 0 then nMovImporte*-1 else 0 end nhaber from MovCta where nMovNro=" & lnMovNro)
    
    lnMontoDebe = 0: lnMontoHaber = 0
    lnItenDebe = 0: lnItenHaber = 0: lnImporteDebe = 0: lnImporteHaber = 0
    Do While Not rsAsiento.EOF
        If rsAsiento!nDebe > 0 And lnItenDebe = 0 And Left(rsAsiento!cCtaContCod, 2) = "45" Then
            lnItenDebe = rsAsiento!nMovItem
            lnImporteDebe = rsAsiento!nDebe
            
        ElseIf rsAsiento!nHaber > 0 And lnItenHaber = 0 And Left(rsAsiento!cCtaContCod, 2) = "45" Then
            lnItenHaber = rsAsiento!nMovItem
            lnImporteHaber = rsAsiento!nHaber
        End If
        
        lnMontoDebe = lnMontoDebe + rsAsiento!nDebe
        lnMontoHaber = lnMontoHaber + rsAsiento!nHaber
        rsAsiento.MoveNext
    Loop
    
    If lnMontoDebe > lnMontoHaber Then
        oAlmacen.ActualizaMovCta lnMovNro, lnItenDebe, , lnImporteDebe - (lnMontoDebe - lnMontoHaber)
    ElseIf lnMontoHaber > lnMontoDebe Then
        oAlmacen.ActualizaMovCta lnMovNro, lnItenDebe, , lnImporteDebe + (lnMontoHaber - lnMontoDebe)
    End If
    '*** FIN PEAC
      
     If FlexSerie.TextMatrix(1, 1) <> "" Then
        For i = 1 To Me.FlexSerie.Rows - 1
          If InStr(1, Me.FlexDetalle.TextMatrix(CInt(Me.FlexSerie.TextMatrix(i, 3)), 2), "[S]") <> 0 Then
            If Me.FlexSerie.TextMatrix(i, 5) <> 0 Then
                oAlmacen.InsertaMovSalidaBSSerie lnMovNro, CInt(Me.FlexSerie.TextMatrix(i, 3)), Me.FlexDetalle.TextMatrix(CInt(Me.FlexSerie.TextMatrix(i, 3)), 1), Me.FlexSerie.TextMatrix(i, 1), Me.FlexSerie.TextMatrix(i, 4), Me.FlexSerie.TextMatrix(i, 5)
            Else
                oAlmacen.InsertaMovSalidaBSSerie lnMovNro, CInt(Me.FlexSerie.TextMatrix(i, 3)), Me.FlexDetalle.TextMatrix(CInt(Me.FlexSerie.TextMatrix(i, 3)), 1), Me.FlexSerie.TextMatrix(i, 1), Me.FlexSerie.TextMatrix(i, 4), Me.FlexDetalle.TextMatrix(CInt(Me.FlexSerie.TextMatrix(i, 3)), 7)
            End If
            'INSERTA ACTVO FIJO
            'oAlmacen.InsertaMovBSActivoFijo Year(gdFecSis), lnMovNro, Me.FlexDetalle.TextMatrix(CInt(Me.FlexSerie.TextMatrix(I, 0)), 1), Me.FlexSerie.TextMatrix(I, 1), CCur(Me.FlexDetalle.TextMatrix(CInt(Me.FlexSerie.TextMatrix(I, 0)), 7)), 0, CDate(Me.mskFecha.Text), Left(Me.txtArea.Text, 3), Mid(Me.txtArea.Text, 4, 2)
          End If
        Next i
      End If
                   
    oAlmacen.CommitTrans
    lbGrabar = True
    
    Set oAsiento = New NContImprimir
    
    cmdImprimir_Click
    
    If lsOpeCod = gnAlmaSalXAtencion Or lsOpeCod = gnAlmaSalXAjuste Or lsOpeCod = gnAlmaSalXTransferenciaOrigen Then oPrevio.Show oAsiento.ImprimeAsientoContable(lsMovNro, 60, 80), Caption, True
    
        'ARLO 20170126 ***
        If (gsOpeCod = 591001) Then
        lsPalabra = "Registrado"
        lsAccion = "1"
        ElseIf (gsOpeCod = 591002) Then
        lsPalabra = "Actualizado"
        lsAccion = "2"
        ElseIf (gsOpeCod = 591201) Then
        lsPalabra = "Dado de Salida"
        lsAccion = "3"
        End If
        
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, lsAccion, "Se ha " & lsPalabra & " El Requemiento N° : " & lsDocNI
        Set objPista = Nothing
        '***
        
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
    
    If Me.txtGuiaRemInt.Visible Then
        If Me.txtGuiaRemInt.Text = "" Then
            MsgBox "Debe elegir una nota de ingreso.", vbInformation, "Aviso"
            Me.txtGuiaRemInt.SetFocus
            Exit Sub
        End If
    Else
        If Me.lblTransportistaG.Caption = "" Then
            Exit Sub
        End If
    End If
    
    Dim oPrevio As clsPrevio
    Set oPrevio = New clsPrevio
    
    Dim lsCadena As String
    Dim lsCadenaSerie As String
    Dim lsDocNom As String * 20
    Dim lsDocFec As String * 12
    Dim lsDocNum As String * 20
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lsItem As String * 5
    Dim lsCodigo As String * 15
    Dim lsNombre As String * 45
    Dim lsUnidad As String * 10
    Dim lsCantidad As String * 10
    Dim lsPrecio As String * 15
    Dim lsTotal As String * 15
    Dim i As Long
    Dim J As Long
    Dim nSuma As Double
    Dim oALmacenRep As DLogAlmacen
    Set oALmacenRep = New DLogAlmacen
    lsCadena = ""
     
    lsCadena = lsCadena & CabeceraPagina1(lblTit.Caption & " - " & Me.lblAlmacenG.Caption, lnPagina, lnItem, gsNomAge, gsEmpresa, CDate(mskFecha.Text), Mid(lsOpeCod, 3, 1)) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & JustificaTextoCadena(Me.txtComentario.Text, 110, 5) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & Me.lblTransportista.Caption & PstaNombre(Me.lblTransportistaG.Caption) & "       AREA  : " & lblAreaG.Caption & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & "MOTIVO DE INGRESO : " & Me.Caption & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & "      DOCUMENTOS : -----------------------------------------" & oImpresora.gPrnSaltoLinea
   
    
    If lsOpeCod = gnAlmaSalXTransferenciaOrigen Then
        lsCadena = lsCadena & Encabezado1("ITEM;5;CODIGO;9; ;10;DESCRIPCION;15; ;30;PRECIO;10; ;5;CANT.ATEND;17; ;5;DIFERENCIA;10; ;3;", lnItem)
    Else
        lsCadena = lsCadena & Encabezado1("ITEM;5;CODIGO;9; ;10;DESCRIPCION;15; ;30;CANT.SOL;10; ;5;CANT.ATEND;17; ;5;DIFERENCIA;10; ;3;", lnItem)
    End If
    
    lsTotal = ""
    For i = 1 To Me.FlexDetalle.Rows - 1
        lsItem = Format(i, "0000")
        lsCodigo = Me.FlexDetalle.TextMatrix(i, 1)
        lsNombre = Me.FlexDetalle.TextMatrix(i, 2)
        
        If lbModifica Then
            RSet lsCantidad = Format(Me.FlexDetalle.TextMatrix(i, 4), "#,##0.00")
            RSet lsPrecio = Format(Me.FlexDetalle.TextMatrix(i, 4), "#,##0.00")
        Else
            If lsOpeCod = gnAlmaSalXTransferenciaOrigen Then
                RSet lsCantidad = Format(Me.FlexDetalle.TextMatrix(i, 7), "#,##0.00")
                RSet lsPrecio = Format(Me.FlexDetalle.TextMatrix(i, 4), "#,##0.00")
            Else
                RSet lsCantidad = Format(Me.FlexDetalle.TextMatrix(i, 3), "#,##0.00")
                RSet lsPrecio = Format(Me.FlexDetalle.TextMatrix(i, 4), "#,##0.00")
            End If
        End If
        If IsNumeric(Me.FlexDetalle.TextMatrix(i, 4)) And IsNumeric(Me.FlexDetalle.TextMatrix(i, 3)) Then
            RSet lsTotal = Format(CCur(Me.FlexDetalle.TextMatrix(i, 3)) - CCur(Me.FlexDetalle.TextMatrix(i, 4)), "#,##0.00")
        End If
        'Obtener la suma de Los Items
        If lsOpeCod = 591201 Then
            nSuma = nSuma + oALmacenRep.GetPrePromedio(1, lsCodigo, 0) * IIf(Val(lsPrecio) = 0, Val(lsPrecio), Val(lsPrecio))
            Else
            nSuma = nSuma + oALmacenRep.GetPrePromedio(1, lsCodigo, 0) * IIf(Val(lsCantidad) = 0, Val(lsPrecio), Val(lsCantidad))
        End If
        lsCadena = lsCadena & Space(4) & "  " & lsItem & lsCodigo & "  " & lsNombre & lsCantidad & "  " & lsPrecio & "  " & lsTotal & oImpresora.gPrnSaltoLinea
        
        lnItem = lnItem + 1
        If InStr(1, Me.FlexDetalle.TextMatrix(i, 2), "[S]") <> 0 Then
            lsItem = ""
            lsCodigo = ""
            lsCadenaSerie = ""
            For J = 1 To Me.FlexSerie.Rows - 1
                If Me.FlexSerie.TextMatrix(J, 3) = Me.FlexDetalle.TextMatrix(i, 0) Then
                    If lsCadenaSerie = "" Then
                        lsCadenaSerie = Me.FlexSerie.TextMatrix(J, 1)
                    Else
                        lsCadenaSerie = lsCadenaSerie & " / " & Me.FlexSerie.TextMatrix(J, 1)
                        lnItem = lnItem + 1
                    End If
                    If J Mod 3 = 0 Then
                        lsCadena = lsCadena & Space(4) & lsItem & lsCodigo & "    " & lsCadenaSerie & oImpresora.gPrnSaltoLinea
                        lsCadenaSerie = ""
                    End If
                End If
            Next J
            lsCadena = lsCadena & Space(5) & lsItem & lsCodigo & "    " & lsCadenaSerie & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
        End If
        
        If lnItem > 44 Then
             lnItem = 0
             lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
             lsCadena = lsCadena & CabeceraPagina1(lblTit.Caption, lnPagina, lnItem, gsNomAge, gsEmpresa, CDate(mskFecha.Text), Mid(lsOpeCod, 3, 1)) & oImpresora.gPrnSaltoLinea
             lsCadena = lsCadena & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
             
             'lsCadena = lsCadena & Me.fraProveedor.Caption & " : " & PstaNombre(Me.lblProveedorNombre.Caption) & oImpresora.gPrnSaltoLinea
             lsCadena = lsCadena & Space(5) & "MOTIVO DE INGRESO : " & Me.Caption & oImpresora.gPrnSaltoLinea
             lsCadena = lsCadena & Space(5) & "      DOCUMENTOS : -----------------------------------------" & oImpresora.gPrnSaltoLinea
            
             lsCadena = lsCadena & Encabezado1("ITEM;5;CODIGO;9; ;10;DESCRIPCION;15; ;30;CANT.SOL;10; ;5;CANT.ATEND;17; ;5;DIFERENCIA;10; ;3;", lnItem)
        End If
    Next i
    
    'lsCadena = lsCadena & Space(5) & String(120, "=") & oImpresora.gPrnSaltoLinea & Space(2) & String(90, " ") + "VALOR REQUERIMIENTO" + Str(Format(nSuma, "########.#0")) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & String(120, "=") & oImpresora.gPrnSaltoLinea & Space(2) & String(90, " ") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    
    
    lsCadena = lsCadena & Space(5) & "----------------------------            ----------------------------       ---------------------------" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & "    FIRMA ALMACENERO                             Logistica                             Vo Bo          " & oImpresora.gPrnSaltoLinea
    
    oPrevio.Show lsCadena, Caption, True, 66, gImpresora
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub FlexDetalle_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 1 Then
        Me.FlexDetalle.TextMatrix(pnRow, 6) = ""
    End If
End Sub

Private Sub FlexDetalle_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    Dim lnStock As Double
    
    If lsOpeCod = gnAlmaReqAreaReg Or lsOpeCod = gnAlmaReqAreaMant Then
        If lsOpeCod = gnAlmaSalXDevolEmbargo Then
            lnStock = oAlmacen.GetStock("-1", FlexDetalle.TextMatrix(pnRow, 1), 2)
        Else
            lnStock = oAlmacen.GetStock(Me.txtAlmacen.Text, FlexDetalle.TextMatrix(pnRow, 1), 0)
        End If
        If CDbl(FlexDetalle.TextMatrix(pnRow, 3)) > lnStock Or CDbl(FlexDetalle.TextMatrix(pnRow, 3)) <= 0 Then   'EJVG20111111
            MsgBox "No hay existencia disponible en los almacenes para satisfacer su pedido.", vbInformation, "Aviso"
            'AVISA QUE NO HAY STOC PERO DEJA PASAR
            'Cancel = False
            If Me.chkConStock.value = 1 Then
                Cancel = False
            End If
        End If
    Else
    
        If FlexDetalle.TextMatrix(pnRow, 4) = "" Then
            FlexDetalle.TextMatrix(pnRow, 4) = FlexDetalle.TextMatrix(pnRow, 3)
        End If
        If lsOpeCod = gnAlmaSalXAtencion Then
            If CDbl(Int(FlexDetalle.TextMatrix(pnRow, 3))) < CDbl(FlexDetalle.TextMatrix(pnRow, 4)) Or CDbl(Int(FlexDetalle.TextMatrix(pnRow, 4))) > CDbl(FlexDetalle.TextMatrix(pnRow, 5)) Then
                Cancel = False
            End If
        ElseIf Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4) Then
            If CDbl(Int(FlexDetalle.TextMatrix(pnRow, 4))) > CDbl(FlexDetalle.TextMatrix(pnRow, 5)) Then
                Cancel = False
            End If
        End If
    End If
End Sub

Private Sub FlexDetalle_RowColChange()
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    Dim lsCtaCnt As String
    Dim i As Integer
    Dim lnContador As Integer
    Dim lnEncontrar As Integer
    
    
    If (Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4) And Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1) <> "") Or lbModifica Then
        If lsOpeCod = gnAlmaSalXDevolEmbargo Then
            Me.FlexDetalle.TextMatrix(FlexDetalle.row, 5) = oAlmacen.GetStock(Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1), 2)
            'Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 5) = oAlmacen.GetStockEmbargo(Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 1))
            Me.FlexDetalle.TextMatrix(FlexDetalle.row, 7) = oAlmacen.GetPrePromedio(Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1), 2)
        Else
            Me.FlexDetalle.TextMatrix(FlexDetalle.row, 5) = oAlmacen.GetStock(Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1), 0)
            Me.FlexDetalle.TextMatrix(FlexDetalle.row, 7) = oAlmacen.GetPrePromedio(Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1), 0)
        End If
        
    End If
    
    If Me.FlexDetalle.col = 1 Then
        If lsOpeCod = "571009" Or lsOpeCod = "572009" Then
            'Me.FlexDetalle.rsTextBuscar = oAlmacen.GetBienesAlmacen(, "11',12','13")
        Else
            'Me.FlexDetalle.rsTextBuscar = oAlmacen.GetBienesAlmacen(, "11','12','13")
        End If
    End If
        
    If lbModifica Then
        lsCtaCnt = GetCtaCntBS(Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1), GetOpeMov(lnMovNroG), Format(Me.txtAlmacen.Text, "00"))
    Else
        If FlexDetalle.TextMatrix(FlexDetalle.row, 6) = "" Then
            lsCtaCnt = GetCtaCntBS(Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1), lsOpeCod, Format(Me.txtAlmacen.Text, "00"))
        End If
    End If
    
    If FlexDetalle.TextMatrix(FlexDetalle.row, 6) = "" Then
        FlexDetalle.TextMatrix(FlexDetalle.row, 6) = lsCtaCnt
    End If
    
    If (InStr(1, Me.FlexDetalle.TextMatrix(FlexDetalle.row, 2), "[S]") <> 0 And lsOpeCod <> "571001" And lsOpeCod <> "571001") Or (Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1) = "" And lsOpeCod <> "571001" And lsOpeCod <> "571001") Then
        
        lnContador = 0
         
        If Not IsNumeric(Me.FlexDetalle.TextMatrix(FlexDetalle.row, 4)) Then Exit Sub
        
        For i = 1 To CInt(Me.FlexSerie.Rows - 1)
            If FlexSerie.TextMatrix(i, 3) = Me.FlexDetalle.TextMatrix(FlexDetalle.row, 0) Then
                lnContador = lnContador + 1
                FlexSerie.RowHeight(i) = 285
            Else
                FlexSerie.RowHeight(i) = 0
            End If
        Next i
        
        If lnContador <> CInt(Me.FlexDetalle.TextMatrix(FlexDetalle.row, 4)) Then
            i = 0
            While lnEncontrar < lnContador
                i = i + 1
                If FlexSerie.TextMatrix(i, 3) = Me.FlexDetalle.TextMatrix(FlexDetalle.row, 0) Then
                    Me.FlexSerie.EliminaFila i
                    lnEncontrar = lnEncontrar + 1
                    i = i - 1
                End If
            Wend
            For i = 1 To CInt(Me.FlexDetalle.TextMatrix(FlexDetalle.row, 4))
                If Me.FlexSerie.TextMatrix(1, 3) = "" Then
                    Me.FlexSerie.AdicionaFila
                Else
                    Me.FlexSerie.AdicionaFila , , True
                End If
                FlexSerie.TextMatrix(FlexSerie.Rows - 1, 0) = Me.FlexDetalle.TextMatrix(FlexDetalle.row, 0)
                FlexSerie.TextMatrix(FlexSerie.Rows - 1, 2) = Me.FlexDetalle.TextMatrix(FlexDetalle.row, 0)
                FlexSerie.TextMatrix(FlexSerie.Rows - 1, 3) = Me.FlexDetalle.TextMatrix(FlexDetalle.row, 0)
            Next i
            FlexSerie_RowColChange
        End If
    Else
        For i = 1 To CInt(Me.FlexSerie.Rows - 1)
            FlexSerie.RowHeight(i) = 0
        Next i
    End If
    
    Set oAlmacen = Nothing
End Sub


Private Sub FlexSerie_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    Dim lnIGV As Currency
    Dim lnValor As Currency
    
        
    If psDataCod <> "" Then
        oAlmacen.GetPrecioActivo Me.FlexDetalle.TextMatrix(FlexSerie.TextMatrix(pnRow, 3), 1), FlexSerie.TextMatrix(pnRow, 1), lnIGV, lnValor
        FlexSerie.TextMatrix(pnRow, 4) = lnIGV
        FlexSerie.TextMatrix(pnRow, 5) = lnValor
    End If
End Sub

Private Sub FlexSerie_RowColChange()
    Dim i As Integer
    Dim lsCadena As String
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    
    lsCadena = ""
    If lnRowFlex <> Me.FlexSerie.row Then
        
        lsCadena = "''"
        For i = 1 To Me.FlexSerie.Rows - 1
            If Me.FlexSerie.TextMatrix(i, 1) <> "" And FlexDetalle.TextMatrix(FlexDetalle.row, 0) = Me.FlexSerie.TextMatrix(i, 3) Then
                If lsCadena = "''" Then
                    lsCadena = "'" & Me.FlexSerie.TextMatrix(i, 1) & "'"
                Else
                    lsCadena = lsCadena & ",'" & Me.FlexSerie.TextMatrix(i, 1) & "'"
                End If
            End If
        Next i
        
        If lsOpeCod = gnAlmaSalXAtencion Then
            Me.FlexSerie.rsTextBuscar = oAlmacen.GetSerieBien(FlexDetalle.TextMatrix(FlexDetalle.row, 1), Right(lsOpeCod, 2), lsCadena)
        ElseIf Mid(lsOpeCod, 4, 1) = "3" Then
            Me.FlexSerie.rsTextBuscar = oAlmacen.GetSerieBien(FlexDetalle.TextMatrix(FlexDetalle.row, 1), Me.txtAlmacen.Text, lsCadena)
        End If
        
        
        lnRowFlex = Me.FlexSerie.row
        
    End If
End Sub

Private Sub Form_Load()
    Dim oDoc As DOperaciones
    Set oDoc = New DOperaciones
    Dim oDoc1 As DOperacion
    Set oDoc1 = New DOperacion
    Dim oArea As DActualizaDatosArea
    Set oArea = New DActualizaDatosArea
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oPersona As UPersona
    Set oPersona = New UPersona
    Dim oMov As DMov
    Set oMov = New DMov
    
    lbGrabar = False
    
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    
    lnMovNroG = 0
    
    If gsCodPersUser <> "" And lsOpeCod = gnAlmaReqAreaReg Then
        oPersona.ObtieneClientexCodigo gsCodPersUser
        Me.txtTransportista.Text = gsCodPersUser
        Me.lblTransportistaG.Caption = oPersona.sPersNombre
    End If
    
    Me.FlexDetalle.rsTextBuscar = oAlmacen.GetBienesAlmacen(, "" & gnLogBSTpoBienConsumo & "','" & gnLogBSTpoBienFijo & "','" & gnLogBSTpoBienNoDepreciable & "")
    
    Me.txtAlmacen.rs = oDoc.GetAlmacenes
    Me.txtAlmacenDestino.rs = oDoc.GetAlmacenes
    
'    Me.txtAlmacen.Text = "1"
    Me.txtAlmacen.Text = gsCodAge
    'EJVG20140319 ***
    If lsOpeCod = gnAlmaReqAreaReg Then
        'txtAlmacen.Enabled = False
    ElseIf lsOpeCod = gnAlmaSalXAtencion Then
        txtAlmacen.Enabled = False
        txtAlmacen_EmiteDatos
    ElseIf Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4) Then 'Otras Salidas
        txtAlmacen.Enabled = False
    End If
    'END EJVG *******
    Me.lblAlmacenG.Caption = txtAlmacen.psDescripcion
    
    Me.txtAlmacenDestino.Text = gsCodAge
    Me.lblAlmacenDestino.Caption = txtAlmacenDestino.psDescripcion
    
    Me.mskFecha = Format(gdFecSis, gsFormatoFechaView)
    
    If lsOpeCod = gnAlmaSalXTransferenciaOrigen Then
        Me.txtArea.rs = oArea.GetAgenciasAreas("023%")
    Else
        Me.txtArea.rs = oArea.GetAgenciasAreas
    End If
    
    'Habilitar CAmbio de Persona para el Ara de Logistica
    If gsCodArea = gsLogistica Then
        Me.txtTransportista.Enabled = True
    End If
    
    Set rs = oDoc1.CargaOpeDoc(lsOpeCod, "2")
    
    If rs.EOF And rs.BOF Then
        MsgBox "La operación no tiene ningun documento agregado.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    lnOpeDoc = rs.Fields(1)
     
    If lbExtorno Or lbModifica Then
        Me.lblTit.Caption = "Guia de Remisión : "
        FlexDetalle.lbEditarFlex = False
        
        Me.txtGuiaRemInt.rs = oDoc.GetGuiaSalida("13", lsOpeCod, "70")
        
        If lbModifica Then
            Me.cmdAsiento.Visible = True
            FlexDetalle.ColumnasAEditar = "X-1-X-3-X-X-X-X"
            FlexDetalle.lbEditarFlex = True
        End If
        
    Else
        'If (Not lbIngreso And Not lbRepRequerimiento) And ((Left(lsOpeCod, 4) = Left(gnAlmaReqAreaReg, 4) Or lsOpeCod = gnAlmaSalXAtencion)) Then Me.txtGuiaRemInt.rs = odoc.GetGuiaInterna("13", lsOpeCod, "70")
        'If Not lbIngreso And Not lbRepRequerimiento Then Me.txtGuiaRemInt.rs = oDoc.GetGuiaInterna("13", lsOpeCod, "70")
    End If
     
    If Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4) Then
        Me.lblTit.Caption = "Guia de Remisión : " & oMov.GeneraDocNro(lnOpeDoc, gMonedaExtranjera, Year(gdFecSis))
        FlexDetalle.ColumnasAEditar = "X-1-X-X-4-X-X-X"
    Else
        FlexDetalle.ColumnasAEditar = "X-1-X-3-X-X-X-X"
        If lbIngreso Then
            Me.lblTit.Caption = "Requerimiento : " & oMov.GeneraDocNro(lnOpeDoc, gMonedaExtranjera, Year(gdFecSis))
        Else
           'Me.lblTit.Caption = "Requerimiento : "
        End If
    End If
    
    Me.lblGuaRemInt.Visible = Not lbIngreso
    Me.lblGuaRemIntG.Visible = Not lbIngreso
    Me.txtGuiaRemInt.Visible = Not lbIngreso
    Me.cmdImprimir.Visible = True
    Me.lblAlmacen.Visible = lbAlmacen
    Me.txtAlmacen.Visible = lbAlmacen
    Me.lblAlmacen.Visible = lbAlmacen
    Me.lblAlmacenG.Visible = lbAlmacen

               
    If lsOpeCod = gnAlmaSalXAtencion Or lsOpeCod = gnAlmaReqAreaExt Then
        If lbConfirma Then
            Me.cmdAgregar.Visible = False
            Me.cmdEliminar.Visible = False
            
            If lsOpeCod = gnAlmaSalXAtencion Or lsOpeCod = gnAlmaReqAreaExt Then
                Me.cmdGrabar.Caption = "&Confirma"
                Me.txtGuiaRemInt.Visible = True
                Me.lblGuaRemInt.Visible = True
                Me.lblGuaRemIntG.Visible = True
            Else
                Me.cmdGrabar.Caption = "&Rechaza"
                Me.txtGuiaRemInt.Visible = False
                Me.lblGuaRemInt.Visible = False
                Me.lblGuaRemIntG.Visible = False
            End If
            Me.txtTransportista.Enabled = False
            Me.txtArea.Enabled = False
        End If
    ElseIf Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4) Then
        If lsOpeCod = gnAlmaSalXAjuste Then 'Ajuste
           ' Me.txtArea.Visible = False
           ' Me.lblArea.Visible = False
           ' Me.lblAreaG.Visible = False
            'Me.lblTransportista.Visible = False
            'Me.lblTransportistaG.Visible = False
            'Me.txtTransportista.Visible = False
        ElseIf lsOpeCod = gnAlmaSalXProvGarantRepa Or lsOpeCod = gnAlmaSalXProvDevCompras Or lsOpeCod = gnAlmaSalXProvDemosOtros Then
            Me.txtArea.Visible = False
            Me.lblArea.Visible = False
            Me.lblAreaG.Visible = False
            Me.lblTransportista.Caption = "Proveedor"
        ElseIf lsOpeCod = gnAlmaSalXTransferenciaOrigen Then
            Me.txtArea.Visible = True
            Me.lblArea.Visible = True
            Me.lblAreaG.Visible = True
            Me.txtTransportista.Enabled = True
            Me.txtAlmacenDestino.Visible = True
            Me.lblAlmacenDestino.Visible = True
            Me.lblAlmDest.Visible = True
            
        ElseIf lsOpeCod = gnAlmaSalXDevolEmbargo Then
            Me.txtArea.Visible = False
            Me.lblArea.Visible = False
            Me.lblAreaG.Visible = False
        ElseIf lsOpeCod = gnAlmaSalXAreasDevGaranRepa Then
            Me.txtArea.Visible = False
            Me.lblArea.Visible = False
            Me.lblAreaG.Visible = False
        ElseIf lsOpeCod = gnAlmaSalXOtrosMotivos Then
            Me.txtArea.Visible = False
            Me.lblArea.Visible = False
            Me.lblAreaG.Visible = False
        ElseIf lsOpeCod = gnAlmaSalXAjuste Then
            Me.txtArea.Visible = False
            Me.lblArea.Visible = False
            Me.lblAreaG.Visible = False
        End If
        
        Me.txtGuiaRemInt.Visible = False
        Me.lblGuaRemInt.Visible = False
        Me.lblGuaRemIntG.Visible = False
        
    End If
    
    If lbExtorno Or lbModifica Then
        Me.txtAlmacen.Enabled = False
        Me.txtTransportista.Enabled = False
         Me.cmdGrabar.Caption = "&Extornar"
    ElseIf lbModifica Then
        
    End If
    
    If lsOpeCod = "591002" Or lsOpeCod = "591003" Or lsOpeCod = "591004" Then
        Me.txtAlmacen.Enabled = False
        If (Not lbIngreso And Not lbRepRequerimiento) And ((Left(lsOpeCod, 4) = Left(gnAlmaReqAreaReg, 4) Or lsOpeCod = gnAlmaSalXAtencion)) Then
            Me.txtGuiaRemInt.rs = oDoc.GetGuiaInterna("13", lsOpeCod, "70", -1)
        End If
    End If
    
    If lsOpeCod = gnAlmaExtornoXSalida Then 'EJVG20111114
        cmdAgregar.Visible = False
        cmdEliminar.Visible = False
    End If
'    lnMovNroG = 0
    lnRowFlex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not lbGrabar Then
       If MsgBox("Desea Salir sin grabar? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            Cancel = 1
       End If
    End If
End Sub

Private Sub mskFecha_GotFocus()
    mskFecha.SelStart = 0
    mskFecha.SelLength = 50
End Sub

Private Sub mskFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cmdGrabar.Enabled Then
        Me.cmdGrabar.SetFocus
    End If
End Sub

Private Sub mskFecha_LostFocus()
    If Not IsDate(Me.mskFecha.Text) Then
        MsgBox "Debe ingresar una fecha correcta", vbInformation, "Aviso"
        mskFecha_GotFocus
        mskFecha.SetFocus
    End If
End Sub

Private Sub txtAlmacen_EmiteDatos()
    Dim lnI As Integer
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    
    Me.FlexDetalle.Rows = 2
    Me.FlexDetalle.Clear
    Me.FlexDetalle.FormaCabecera
    
    If Me.txtAlmacen.Text = "" Then Exit Sub
    Me.lblAlmacenG.Caption = Me.txtAlmacen.psDescripcion
    
    If Left(lsOpeCod, 4) = "5910" Or Left(lsOpeCod, 4) = "5913" Or Left(lsOpeCod, 4) = "5914" Then
        Exit Sub
    End If
    
    txtGuiaRemInt.Text = ""
    
    Dim oDoc As DOperaciones
    Set oDoc = New DOperaciones
    
    If (Not lbIngreso And Not lbRepRequerimiento) And ((Left(lsOpeCod, 4) = Left(gnAlmaReqAreaReg, 4) Or lsOpeCod = gnAlmaSalXAtencion)) Then
        Me.txtGuiaRemInt.rs = oDoc.GetGuiaInterna("13", lsOpeCod, "70", Me.txtAlmacen.Text)
    End If
    
'    For lnI = 1 To Me.FlexDetalle.Rows - 1
'        If Me.FlexDetalle.TextMatrix(lnI, 1) <> "" Then
'            If lsOpeCod = gnAlmaSalXDevolEmbargo Then
'                Me.FlexDetalle.TextMatrix(lnI, 5) = oALmacen.GetStock(Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(lnI, 1), 2)
'                Me.FlexDetalle.TextMatrix(lnI, 7) = oALmacen.GetPrePromedio(Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(lnI, 1), 2)
'            Else
'                Me.FlexDetalle.TextMatrix(lnI, 5) = oALmacen.GetStock(Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(lnI, 1), 0)
'                Me.FlexDetalle.TextMatrix(lnI, 7) = oALmacen.GetPrePromedio(Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(lnI, 1), 0)
'            End If
'        End If
'    Next lnI
End Sub

Private Sub txtAlmacenDestino_EmiteDatos()
    Dim lnI As Integer
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    
    Me.FlexDetalle.Rows = 2
    Me.FlexDetalle.Clear
    Me.FlexDetalle.FormaCabecera
    
    If Me.txtAlmacenDestino.Text = "" Then Exit Sub
    Me.lblAlmacenDestino.Caption = Me.txtAlmacenDestino.psDescripcion
    
    If Left(lsOpeCod, 4) = "5910" Or Left(lsOpeCod, 4) = "5913" Or Left(lsOpeCod, 4) = "5914" Then
        Exit Sub
    End If
    
    'txtGuiaRemInt.Text = ""
    
    'Dim oDoc As DOperaciones
    'Set oDoc = New DOperaciones
    
    'If (Not lbIngreso And Not lbRepRequerimiento) And ((Left(lsOpeCod, 4) = Left(gnAlmaReqAreaReg, 4) Or lsOpeCod = gnAlmaSalXAtencion)) Then
    '    Me.txtGuiaRemInt.rs = oDoc.GetGuiaInterna("13", lsOpeCod, "70", Me.txtAlmacen.Text)
    'End If
End Sub
Private Sub TxtArea_EmiteDatos()
    Me.lblAreaG.Caption = txtArea.psDescripcion
End Sub
Private Sub txtComentario_GotFocus()
    txtComentario.SelStart = 0
    txtComentario.SelLength = 300
End Sub
Private Sub txtComentario_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
End Sub

Private Sub txtGuiaRemInt_EmiteDatos()
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   Dim oOpe As DOperaciones
   Set oOpe = New DOperaciones
   Dim lnItemAnt As String
   Dim i As Integer
   Dim lnMovNroAnt As Long
   Dim lbAtendioPar As Boolean
       Dim oAlmacen As DLogAlmacen
   Set oAlmacen = New DLogAlmacen
   
   lbAtendioPar = False
   
    If Not lbReporte Then
        If Me.txtAlmacen.Text = "" And lsOpeCod <> gnAlmaReqAreaMant And lsOpeCod <> gnAlmaReqAreaExt Then
             MsgBox "Debe ingresar un almacen.", vbInformation, "Aviso"
             Me.txtGuiaRemInt.Text = ""
             Me.lblGuaRemIntG.Caption = ""
             txtAlmacen.SetFocus
             Exit Sub
        End If
        
        Me.lblGuaRemIntG.Caption = Left(txtGuiaRemInt.psDescripcion, 50)
        
         If Me.lblGuaRemIntG.Caption <> "" Then
             
             lnMovNroG = Mid(txtGuiaRemInt.psDescripcion, InStr(1, txtGuiaRemInt.psDescripcion, "[") + 1, InStr(1, txtGuiaRemInt.psDescripcion, "]") - InStr(1, txtGuiaRemInt.psDescripcion, "[") - 1)
             
             'Mid(lsOpeCod, 4, 1) = "2" And Right(lsOpeCod, 2) <> "99"
             If Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4) Then
                 'Me.lblTit.Caption = "Guia de Remisión : " & txtGuiaRemInt.Text
             Else
                 If lbExtorno Or lbModifica Then
                     Me.lblTit.Caption = "Guia de Remisión : " & txtGuiaRemInt.Text
                 Else
                     Me.lblTit.Caption = "Requerimiento : " & txtGuiaRemInt.Text
                 End If
             End If
             
             If lbExtorno Or lbModifica Then
                 Set rs = oOpe.GetGuiaSalidaDet(Me.txtGuiaRemInt.Text, IIf(lsOpeCod = gnAlmaReqAreaMant Or lsOpeCod = gnAlmaReqAreaExt Or Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4), True, False), Me.txtAlmacen.Text)
                 If rs.EOF And rs.BOF Then
                    MsgBox "No Existen Registros.", vbInformation, "Aviso"
                    Exit Sub
                 End If
                 Me.mskFecha.Text = Format(rs!dDocFecha, gsFormatoFechaView)
             Else
                 If lsOpeCod = gnAlmaSalXDevolEmbargo Then
                    Set rs = oOpe.GetDetGuiaRemision(Me.txtGuiaRemInt.Text, 2, IIf(lsOpeCod = gnAlmaReqAreaRechPar Or lsOpeCod = gnAlmaReqAreaMant Or lsOpeCod = gnAlmaReqAreaExt Or Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4), True, False), Me.txtAlmacen.Text)
                 Else
                    If lbRepRequerimiento Then
                        Set rs = oOpe.GetDetGuiaRemision(Me.txtGuiaRemInt.Text, 0, IIf(lsOpeCod = gnAlmaReqAreaRechPar Or lsOpeCod = gnAlmaReqAreaMant Or lsOpeCod = gnAlmaReqAreaExt Or Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4), True, False), Me.txtAlmacen.Text, "", False)
                    Else
                        Set rs = oOpe.GetDetGuiaRemision(Me.txtGuiaRemInt.Text, 0, IIf(lsOpeCod = gnAlmaReqAreaRechPar Or lsOpeCod = gnAlmaReqAreaMant Or lsOpeCod = gnAlmaReqAreaExt Or Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4), True, False), Me.txtAlmacen.Text)
                    End If
                 End If
                                 
                 If rs.EOF And rs.BOF Then
                    MsgBox "No existen datos.", vbInformation, "Aviso"
                    Exit Sub
                 End If
                 
                 If lsOpeCod = gnAlmaSalXAtencion Then
                    Me.mskFecha.Text = Format(gdFecSis, gsFormatoFechaView)
                 Else
                    Me.mskFecha.Text = Format(rs!dDocFecha, gsFormatoFechaView)
                End If
             End If
             
             If Not IsNull(rs!cPersCod) Then
                 Me.txtTransportista.Text = rs!cPersCod
                 Me.lblTransportistaG.Caption = rs!cPersNombre
             Else
                 Me.txtTransportista.Text = ""
                 Me.lblTransportistaG.Caption = ""
             End If
             
             If lsOpeCod = gnAlmaSalXAtencion Then
                Me.txtComentario.Text = "SALIDA POR ATENCION REQUERIMIENTO " & Me.txtGuiaRemInt.Text
             Else
                 Me.txtComentario.Text = rs!cMovDesc
             End If
             Me.txtArea.Text = rs!cAreaCod & rs!cAgeCod
             Me.lblAreaG.Caption = rs!cAreaDescripcion + "  " + GetDescAgencia(rs!cAgeCod)
             
             If Left(lsOpeCod, 4) = "5910" Or Left(lsOpeCod, 4) = "5913" Or Left(lsOpeCod, 4) = "5914" Then
                Me.txtAlmacen.Text = rs!nMovBsOrden
                txtAlmacen_EmiteDatos
             End If
             
             FlexDetalle.Clear
             FlexDetalle.FormaCabecera
             FlexDetalle.Rows = 2
             
             lnItemAnt = 0
             
             FlexSerie.Clear
             FlexSerie.Rows = 2
             FlexSerie.FormaCabecera
             
             lnMovNroAnt = rs!nMovNro
             
             While Not rs.EOF
                 If lnItemAnt <> rs!nMovItem Then
                     Me.FlexDetalle.AdicionaFila
                     Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 1) = rs!cBSCod
                     Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 2) = rs!cBSDescripcion
                     Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 3) = rs!nMovCant - rs!Atendidos
                     
                     If lsOpeCod = gnAlmaReqAreaMant And rs!Atendidos <> 0 Then
                         lbAtendioPar = True
                     End If
                     
                     Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 4) = 0
                     Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 6) = rs!cCtaContCod
                     
                     If lsOpeCod = gnAlmaSalXAtencion Then
                         Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 4) = 0
                         Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 5) = rs!Stock
                         Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 7) = rs!PrePromedio
                     Else
                         Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 4) = ""
                         Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 5) = ""
                     End If
                     If lbModifica Then
                         If lsOpeCod = gnAlmaSalXDevolEmbargo Then
                            Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 5) = oAlmacen.GetStock(Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1), 2)
                            Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 7) = oAlmacen.GetPrePromedio(Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1), 2)
                         Else
                            Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 5) = oAlmacen.GetStock(Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1), 0)
                            Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 7) = oAlmacen.GetPrePromedio(Me.txtAlmacen.Text, Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1), 0)
                         End If
                     End If
                 End If
                 
            If InStr(1, rs!cBSDescripcion, "[S]") <> 0 And lsOpeCod <> gnAlmaSalXAtencion And Left(lsOpeCod, 4) <> Left(gnAlmaReqAreaReg, 4) Then
                Me.FlexSerie.AdicionaFila
                Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 0) = rs!nMovItem
                Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 1) = IIf(IsNull(rs!cSerie), "", rs!cSerie)
                Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 2) = rs!nMovItem
                Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 3) = rs!nMovItem
                Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 4) = rs!nIGV
                Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 5) = rs!nValor
            End If
             
                 lnItemAnt = rs!nMovItem
                 rs.MoveNext
             Wend
         End If
         
         If lsOpeCod = gnAlmaReqAreaMant And lsOpeCod = gnAlmaReqAreaExt Then
             If lbAtendioPar Then
                 Me.cmdAgregar.Enabled = False
                 Me.cmdEliminar.Enabled = False
                 Me.cmdGrabar.Enabled = False
             Else
                 Me.cmdAgregar.Enabled = True
                 Me.cmdEliminar.Enabled = True
                 Me.cmdGrabar.Enabled = True
             End If
         End If
    Else 'REPORTE
        
        If Me.txtAlmacen.Text = "" And lsOpeCod <> gnAlmaReqAreaMant And lsOpeCod <> gnAlmaReqAreaExt Then
             MsgBox "Debe ingresar un almacen.", vbInformation, "Aviso"
             Me.txtGuiaRemInt.Text = ""
             Me.lblGuaRemIntG.Caption = ""
             txtAlmacen.SetFocus
             Exit Sub
        End If
        
        Me.lblGuaRemIntG.Caption = Left(txtGuiaRemInt.psDescripcion, 50)
        
         If Me.lblGuaRemIntG.Caption <> "" Then
             
             lnMovNroG = Mid(txtGuiaRemInt.psDescripcion, InStr(1, txtGuiaRemInt.psDescripcion, "[") + 1, InStr(1, txtGuiaRemInt.psDescripcion, "]") - InStr(1, txtGuiaRemInt.psDescripcion, "[") - 1)
             
             'Mid(lsOpeCod, 4, 1) = "2" And Right(lsOpeCod, 2) <> "99"
             If Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4) Then
                 'Me.lblTit.Caption = "Guia de Remisión : " & txtGuiaRemInt.Text
             Else
                 If lbExtorno Or lbModifica Then
                     Me.lblTit.Caption = "Guia de Remisión : " & txtGuiaRemInt.Text
                 ElseIf lbReporte Then
                    Me.lblTit.Caption = "Guia de Remisión : " & txtGuiaRemInt.Text
                 Else
                     Me.lblTit.Caption = "Requerimiento : " & txtGuiaRemInt.Text
                 End If
             End If
             
             If lbExtorno Or lbModifica Then
                 Set rs = oOpe.GetGuiaSalidaDet(Me.txtGuiaRemInt.Text, IIf(lsOpeCod = gnAlmaReqAreaMant Or lsOpeCod = gnAlmaReqAreaExt Or Left(lsOpeCod, 4) = Left(gnAlmaSalXAtencion, 4), True, False), lsOpeCod)
             Else
                 Set rs = oOpe.GetDetNotaIngresoReporte(Me.txtGuiaRemInt.Text, "10", "71")
             End If
            
             Me.mskFecha.Text = Format(rs!dDocFecha, gsFormatoFechaView)

             
             If rs.EOF And rs.BOF Then Exit Sub
             
             If Not IsNull(rs!cPersCod) Then
                 Me.txtTransportista.Text = rs!cPersCod & ""
                 Me.lblTransportistaG.Caption = rs!cPersNombre & ""
             Else
                 Me.txtTransportista.Text = ""
                 Me.lblTransportistaG.Caption = ""
             End If
             Me.txtComentario.Text = rs!cMovDesc
             Me.txtArea.Text = rs!cAreaCod & rs!cAgeCod
             Me.lblAreaG.Caption = rs!cAreaDescripcion
             FlexDetalle.Clear
             FlexDetalle.FormaCabecera
             FlexDetalle.Rows = 2
             
             lnItemAnt = 0
             
             FlexSerie.Clear
             FlexSerie.Rows = 2
             FlexSerie.FormaCabecera
             
             lnMovNroAnt = rs!nMovNro
             
             While Not rs.EOF
                 If lnItemAnt <> rs!nMovItem Then
                     Me.FlexDetalle.AdicionaFila
                     Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 1) = rs!cBSCod
                     Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 2) = rs!cBSDescripcion
                     Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 4) = rs!nMovCant
                     
                     'Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 4) = 0
                     Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 6) = rs!cCtaContCod
                     
                     If lsOpeCod = gnAlmaSalXAtencion Then
                         'Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 4) = 0
                         'Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 5) = rs!Stock
                         'Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 7) = rs!PrePromedio
                     Else
                         'Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 4) = ""
                         'Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 5) = ""
                     End If
                 End If
                 
                 If InStr(1, rs!cBSDescripcion, "[S]") <> 0 Then
                     Me.FlexSerie.AdicionaFila
                     Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 0) = rs!nMovItem
                     Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 1) = IIf(IsNull(rs!cSerie), "", rs!cSerie)
                     Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 2) = rs!nMovItem
                     Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 3) = rs!nMovItem
                     Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 4) = rs!nIGV
                     Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 5) = rs!nValor
                 End If
                 
                 lnItemAnt = rs!nMovItem
                 rs.MoveNext
             Wend
         End If
    
    End If
End Sub

Private Sub txtGuiaRemInt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtGuiaRemInt.Text = Format(gdFecSis, "yyyy") & "-" & Format(txtGuiaRemInt.Text, "00000000")
    End If
End Sub

Private Sub txtTransportista_EmiteDatos()
    Me.lblTransportistaG.Caption = txtTransportista.psDescripcion
End Sub

Public Sub Ini(psOpeCod As String, psCaption As String, Optional pbIngreso As Boolean = True, Optional pbConfirma As Boolean = False, Optional pbAlmacen As Boolean = False, Optional pbExtorno As Boolean = False, Optional pbRechazo As Boolean = False, Optional pbModifica As Boolean = False)
    lsOpeCod = psOpeCod
    lbRepRequerimiento = False
    lbIngreso = pbIngreso
    lbConfirma = pbConfirma
    lbAlmacen = pbAlmacen
    lbExtorno = pbExtorno
    lbRechazo = pbRechazo
    lbReporte = False
    lbModifica = pbModifica
    Caption = psCaption
    Me.Show 1
End Sub

Private Function Valida() As Boolean
    Dim i As Integer
    Dim J As Integer
    Dim lnContador As Integer
    Dim oSal As DLogAlmacen
    Set oSal = New DLogAlmacen
        
    If oSal.CierreMesLogistica(CDate(Me.mskFecha.Text), txtAlmacen.Text) Then
        MsgBox "No se puede modificar el documento, la fecha que se desea ingresar es una fecha anterior a un cierre de almacen, lo que modificaria los reportes de cierre.", vbInformation, "Aviso"
        mskFecha.SetFocus
        Valida = False
        Exit Function
    End If
        
        
    If Me.txtArea.Text = "" And txtArea.Visible Then
        MsgBox "Debe ingresar el Area.", vbInformation, "Aviso"
        Valida = False
        'Me.txtArea.SetFocus
        Exit Function
    End If
    
    If lbExtorno Then
        Valida = True
        Exit Function
    End If
        
    If lbModifica Then
        For i = 1 To Me.FlexDetalle.Rows - 1
            If Me.FlexDetalle.TextMatrix(i, 1) = "" Then
                MsgBox "Debe ingresar un bien valido para el registro " & i & " .", vbInformation, "Aviso"
                FlexDetalle.col = 1
                FlexDetalle.row = i
                Me.FlexDetalle.SetFocus
                Valida = False
                Exit Function
            ElseIf Not IsNumeric(Me.FlexDetalle.TextMatrix(i, 3)) Then
                MsgBox "Debe ingresar una cantidad valido para el registro " & i & " .", vbInformation, "Aviso"
                FlexDetalle.col = 1
                FlexDetalle.row = i
                Me.FlexDetalle.SetFocus
                Valida = False
                Exit Function
            ElseIf InStr(1, Me.FlexDetalle.TextMatrix(i, 2), "[S]") <> 0 And lsOpeCod <> gnAlmaReqAreaReg And lsOpeCod <> gnAlmaReqAreaMant And lsOpeCod <> gnAlmaReqAreaExt Then
                    lnContador = 0

                    For J = 1 To CInt(Me.FlexSerie.Rows - 1)
                        If FlexSerie.TextMatrix(J, 3) = Me.FlexDetalle.TextMatrix(i, 0) And FlexSerie.TextMatrix(J, 1) <> "" Then
                            lnContador = lnContador + 1
                        End If
                    Next J
                If IsNumeric(Me.FlexDetalle.TextMatrix(i, 4)) Then
                    If lnContador <> CInt(Me.FlexDetalle.TextMatrix(i, 4)) Then
                        MsgBox "Debe ingresar una numeros serie valida para el registro " & i & " .", vbInformation, "Aviso"
                        FlexDetalle.col = 1
                        FlexDetalle.row = i
                        FlexSerie.row = lnContador + 1
                        Me.FlexDetalle.SetFocus
                        Valida = False
                        Exit Function
                    End If
                End If
            End If
        Next i
        Valida = True
        Exit Function
    End If
    
    If Not lbIngreso And lsOpeCod = gnAlmaSalXAtencion Then
        If Me.txtGuiaRemInt.Text = "" Then
            MsgBox "Debe ingresar un requerimiento.", vbInformation, "Aviso"
            Me.txtGuiaRemInt.SetFocus
            Valida = False
            Exit Function
        ElseIf Me.txtAlmacen.Text = "" Then
            MsgBox "Debe elegir un almacen.", vbInformation, "Aviso"
            txtAlmacen.SetFocus
            Exit Function
        End If
    End If
    
    If Me.FlexDetalle.TextMatrix(1, 1) = "" Then
        MsgBox "Debe ingresar por lo menos un producto.", vbInformation, "Aviso"
        Me.cmdAgregar.SetFocus
        Valida = False
        Exit Function
    End If
    
    For i = 1 To Me.FlexDetalle.Rows - 1
        If Me.FlexDetalle.TextMatrix(i, 1) = "" Then
            MsgBox "Debe ingresar un bien valido para el registro " & i & " .", vbInformation, "Aviso"
            FlexDetalle.col = 1
            FlexDetalle.row = i
            Me.FlexDetalle.SetFocus
            Valida = False
            Exit Function
        ElseIf Me.FlexDetalle.TextMatrix(i, 3) = "" And Left(lsOpeCod, 4) <> "5912" Then
            MsgBox "Debe ingresar una Cantidad valido para el registro " & i & " .", vbInformation, "Aviso"
            FlexDetalle.col = 3
            FlexDetalle.row = i
            Me.FlexDetalle.SetFocus
            Valida = False
            Exit Function
        ElseIf InStr(1, Me.FlexDetalle.TextMatrix(i, 2), "[S]") <> 0 And lsOpeCod <> gnAlmaReqAreaReg And lsOpeCod <> gnAlmaReqAreaMant And lsOpeCod <> gnAlmaReqAreaExt And lsOpeCod <> gnAlmaReqAreaRechPar Then
            lnContador = 0
            
            For J = 1 To CInt(Me.FlexSerie.Rows - 1)
                If FlexSerie.TextMatrix(J, 3) = Me.FlexDetalle.TextMatrix(i, 0) And FlexSerie.TextMatrix(J, 1) <> "" Then
                    lnContador = lnContador + 1
                End If
            Next J
        
            If lnContador <> CInt(Me.FlexDetalle.TextMatrix(i, 4)) Then
                MsgBox "Debe ingresar un serie valida para el registro " & i & " .", vbInformation, "Aviso"
                FlexDetalle.col = 1
                FlexDetalle.row = i
                FlexSerie.row = lnContador + 1
                Me.FlexDetalle.SetFocus
                Valida = False
                
                Exit Function
            ElseIf Not IsNumeric(Me.FlexDetalle.TextMatrix(i, 3)) And Mid(lsOpeCod, 4, 1) <> "3" Then
                MsgBox "Debe ingresar un valor valido para el registro " & i & " .", vbInformation, "Aviso"
                FlexDetalle.col = 3
                FlexDetalle.row = i
                Me.FlexDetalle.SetFocus
                Valida = False
                Exit Function
            ElseIf Not IsNumeric(Me.FlexDetalle.TextMatrix(i, 4)) And Mid(lsOpeCod, 4, 1) <> "1" Then
                MsgBox "Debe ingresar un valor valido para el registro " & i & " .", vbInformation, "Aviso"
                FlexDetalle.col = 4
                FlexDetalle.row = i
                Me.FlexDetalle.SetFocus
                Valida = False
                Exit Function
            End If
            
            For J = 1 To CInt(Me.FlexSerie.Rows - 1)
                If FlexSerie.TextMatrix(J, 3) = Me.FlexDetalle.TextMatrix(i, 0) Then
                    If (Not lbIngreso And Not VerfBSSerieMov(Me.FlexDetalle.TextMatrix(i, 1), FlexSerie.TextMatrix(J, 1), lnMovNroG)) And (VerfBSSerie(Me.FlexDetalle.TextMatrix(i, 1), FlexSerie.TextMatrix(J, 1), "0") And lbIngreso) Then
                        MsgBox "El bien ya fue ingresado o no ha sido descargado de almacen, para el registro " & i & " .", vbInformation, "Aviso"
                        FlexSerie.col = 1
                        FlexSerie.row = J
                        FlexDetalle.row = CInt(FlexSerie.TextMatrix(FlexSerie.row, 3))
                        FlexDetalle_RowColChange
                        Me.FlexSerie.SetFocus
                        Valida = False
                        Exit Function
                    End If
                End If
            Next J
        
        ElseIf Not IsNumeric(Me.FlexDetalle.TextMatrix(i, 3)) And Mid(lsOpeCod, 4, 1) = "" Then
            MsgBox "Debe ingresar un valor valido para el registro " & i & " .", vbInformation, "Aviso"
            FlexDetalle.col = 3
            FlexDetalle.row = i
            Me.FlexDetalle.SetFocus
            Valida = False
            Exit Function
        ElseIf Not IsNumeric(Me.FlexDetalle.TextMatrix(i, 4)) And Mid(lsOpeCod, 4, 1) <> "1" And Mid(lsOpeCod, 4, 1) <> "0" Then
            MsgBox "Debe ingresar un valor valido para el registro " & i & " .", vbInformation, "Aviso"
            FlexDetalle.col = 4
            FlexDetalle.row = i
            Me.FlexDetalle.SetFocus
            Valida = False
            Exit Function
        ElseIf Not IsNumeric(Me.FlexDetalle.TextMatrix(i, 6)) And lsOpeCod <> gnAlmaReqAreaMant And lsOpeCod <> gnAlmaReqAreaExt And lsOpeCod <> gnAlmaReqAreaRechPar Then
            MsgBox "No se ha definido Cta Contable para el registro (Defina una cuenta contable para este producto) " & i & " .", vbInformation, "Aviso"
            FlexDetalle.col = 4
            FlexDetalle.row = i
            Me.FlexDetalle.SetFocus
            Valida = False
            Exit Function
        End If
        'EJVG20140320 *** Valida la cantidad de atención
        If lsOpeCod = gnAlmaSalXAtencion Then
            If Val(FlexDetalle.TextMatrix(i, 4)) <= 0 Then
                MsgBox "Debe ingresar la cantidad con que se está atendiendo el requerimiento", vbInformation, "Aviso"
                FlexDetalle.col = 4
                FlexDetalle.row = i
                FlexDetalle.SetFocus
                Valida = False
                Exit Function
            End If
        End If
        'END EJVG *******
    Next i
    
    If txtComentario.Text = "" Then
        Valida = False
        MsgBox "Debe ingresar un comentario valido.", vbInformation, "Aviso"
        txtComentario.SetFocus
        Exit Function
    End If
    
    Valida = True
End Function

Public Sub InicioRep(psDoc As String)
    Dim oDoc As DOperaciones
    Set oDoc = New DOperaciones
    lbReporte = True
    lbExtorno = False
    lsOpeCod = gnAlmaIngXComprasConfirma
    Me.txtGuiaRemInt.rs = oDoc.GetNotaIngresoReporte(MovEstado.gMovEstContabMovContable, gnAlmarReporteMovNotIng, "71", psDoc)
    Me.cmdAgregar.Visible = False
    Me.cmdEliminar.Visible = False
    Me.FlexDetalle.lbEditarFlex = False
    Me.FlexSerie.lbEditarFlex = False
    Me.cmdGrabar.Visible = False
    Me.cmdCancelar.Visible = False
    Me.txtTransportista.Enabled = False
    Me.txtAlmacen.Enabled = False
    Me.cmdImprimir.Visible = True
    Me.cmdImprimir.Enabled = True
    Me.cmdAsiento.Visible = True
    lbGrabar = True
    Me.Show 1
End Sub

Public Sub InicioRepReq(psDoc As String)
    Dim oDoc As DOperaciones
    Set oDoc = New DOperaciones
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    lbReporte = False
    lbExtorno = False
    lbModifica = False
    lbRepRequerimiento = True
    
    lsOpeCod = gnAlmaIngXComprasConfirma
    
    Caption = "Requerimiento de Areas"
    
    Set rs = oDoc.GetNotaIngresoReporte("" & MovEstado.gMovEstContabNoContable & "','5", gnAlmarReporteMovNotIng, "70", psDoc)
    
    If rs.EOF And rs.EOF Then
        MsgBox "El Requerimiento ha sido grabado en blanco o sin Datos.", vbInformation, "Aviso"
        lbGrabar = True
        Unload Me
        Exit Sub
    End If
    
    Me.txtGuiaRemInt.rs = rs
    
    Me.cmdAgregar.Visible = False
    Me.cmdEliminar.Visible = False
    Me.FlexDetalle.lbEditarFlex = False
    Me.FlexSerie.lbEditarFlex = False
    Me.cmdGrabar.Visible = False
    Me.cmdCancelar.Visible = False
    Me.txtTransportista.Enabled = False
    Me.txtAlmacen.Enabled = False
    Me.cmdImprimir.Visible = True
    Me.cmdImprimir.Enabled = True
    Me.cmdAsiento.Visible = False
    lbGrabar = True
    Me.Show 1
End Sub

Private Function GetValorTotalSalida(pnRow As Integer) As Currency
    Dim lnI As Long
    Dim lnAcum As Currency
    
    lnAcum = 0
    For lnI = 1 To Me.FlexSerie.Rows - 1
        If Me.FlexSerie.TextMatrix(lnI, 3) = pnRow And IsNumeric(FlexSerie.TextMatrix(lnI, 5)) Then
            lnAcum = lnAcum + CCur(FlexSerie.TextMatrix(lnI, 5))
        End If
    Next lnI

    GetValorTotalSalida = lnAcum
End Function

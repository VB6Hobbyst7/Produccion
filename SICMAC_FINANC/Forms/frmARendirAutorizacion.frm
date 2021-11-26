VERSION 5.00
Begin VB.Form frmARendirAutorizacion 
   Caption         =   "Operaciones:"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12165
   Icon            =   "frmARendirAutorizacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   12165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "E&liminar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   13
      Top             =   4200
      Width           =   1380
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10800
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.CheckBox chkTodo 
      Caption         =   "&Todos"
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
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Value           =   1  'Checked
      Width           =   900
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Frame Frame2 
      Caption         =   "Glosa"
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
      Height          =   915
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   5265
      Begin VB.TextBox txtMovDesc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   195
         Width           =   5010
      End
   End
   Begin VB.CommandButton cmdAprobar 
      Caption         =   "&Aprobar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   6
      ToolTipText     =   "Ingresar Saldo a Caja General"
      Top             =   4200
      Width           =   1380
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   7
      Top             =   4200
      Width           =   1380
   End
   Begin VB.Frame FraSeleccion 
      Enabled         =   0   'False
      Height          =   945
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   10155
      Begin Sicmact.TxtBuscar txtBuscarAgenciaArea 
         Height          =   330
         Left            =   1425
         TabIndex        =   0
         Top             =   180
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         lbUltimaInstancia=   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Area/Agencia :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   9
         Top             =   225
         Width           =   1185
      End
      Begin VB.Label lblAgenciaArea 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2520
         TabIndex        =   1
         Top             =   195
         Width           =   7455
      End
      Begin VB.Label lblAgeDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1410
         TabIndex        =   2
         Top             =   525
         Width           =   8565
      End
   End
   Begin Sicmact.FlexEdit FEARendir 
      Height          =   2820
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   4974
      Cols0           =   15
      HighLight       =   2
      AllowUserResizing=   1
      EncabezadosNombres=   "#--Número-Area-Agencia-Persona Solicitante-Fecha-Importe-Concepto-cCodArea-cPersCod-nMovNro-cDocTpo-cCodAge-Cuenta"
      EncabezadosAnchos=   "0-550-1500-1800-1800-2400-900-1100-0-0-0-0-0-0-1500"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-1-X-X-X-X-X-X-X-X-X-X-X-X-14"
      TextStyleFixed  =   4
      ListaControles  =   "0-4-0-0-0-0-0-0-0-0-0-0-0-0-1"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-L-L-L-C-R-C-C-C-C-C-C-L"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   6
      lbFormatoCol    =   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      RowHeight0      =   300
   End
End
Attribute VB_Name = "frmARendirAutorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'************************************************************************************************
'***Nombre      : frmARendirAutorizacion
'***Descripción : Formulario para Aprobar las Solicitudes de los Váticos y Otros Gastos A Rendir.
'***Creación    : ELRO el 20120423
'************************************************************************************************

Dim lbMN As Boolean
Dim lsMoney As String
Dim lsSimbolo As String
Dim lSalir As Boolean
Dim lnTipoArendir As ArendirTipo
Dim lnArendirFase As ARendirFases

Dim ValorCelda1, ValorCelda2, ValorCelda3 As String 'NAGL 20180205
Dim ValorCelda4, ValorCelda5, ValorCelda6 As String 'NAGL 20180205

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************


Public Sub iniciarAprobacionViaticos()
Me.Caption = gsOpeDesc
Me.Show 1
End Sub
Public Sub iniciarAprobacionARendir()
Me.Caption = gsOpeDesc
Me.Show 1
End Sub

Private Sub limpiarCampos()
txtMovDesc = ""
End Sub

Private Sub cargarPendientes()
Dim oNArendir As NARendir
Set oNArendir = New NARendir
Dim rsPendientes As ADODB.Recordset
Set rsPendientes = New ADODB.Recordset
Dim lsArea As String
Dim lsDescArea As String
Dim X As Integer 'NAGL ERS075-2017 20171120


If gsOpeCod = CStr(gCGArendirViatAprobMN) Or gsOpeCod = CStr(gCGArendirViatAprobME) Then
    Set rsPendientes = oNArendir.devolverARendirViaticosPendientes(IIf(chkTodo, 1, 0), Mid(gsOpeCod, 3, 1), IIf(chkTodo = False, Left(txtBuscarAgenciaArea, 3), ""), IIf(chkTodo = False, IIf(Len(Trim(txtBuscarAgenciaArea)) = 5, Right(txtBuscarAgenciaArea, 2), "01"), ""), gArendirTipoViaticos)

ElseIf gsOpeCod = CStr(gCGArendirCtaAprobMN) Or gsOpeCod = CStr(gCGArendirCtaAprobME) Then
    Set rsPendientes = oNArendir.devolverARendirCtaPendientes(IIf(chkTodo, 1, 0), Mid(gsOpeCod, 3, 1), IIf(chkTodo = False, Left(txtBuscarAgenciaArea, 3), ""), IIf(chkTodo = False, IIf(Len(Trim(txtBuscarAgenciaArea)) = 5, Right(txtBuscarAgenciaArea, 2), "01"), ""), gArendirTipoCajaGeneral)
End If


lSalir = False

FEARendir.Clear
'FEARendir.FormaCabecera
'FEARendir.Rows = 2
'If rsPendientes.EOF And rsPendientes.BOF Then
'   MsgBox "No existe Solicitudes Pendientes", vbInformation, "Aviso"
'   rsPendientes.Close
'   Set rsPendientes = Nothing
'   Exit Sub
'End If
'If Not rsPendientes.EOF And Not rsPendientes.BOF Then
'    Set FEARendir.Recordset = rsPendientes
'    FEARendir.FormatoPersNom 4
'    If FEARendir.Visible Then
'        FEARendir.SetFocus
'    End If
'End If
'FEARendir.FormateaColumnas
'rsPendientes.Close 'Comentado by NAGL 20171120
 '************BEGIN**************************************
FormateaFlex FEARendir
If Not (rsPendientes.EOF And rsPendientes.BOF) Then
    For X = 1 To rsPendientes.RecordCount
        FEARendir.AdicionaFila , , True
        FEARendir.TextMatrix(X, 2) = rsPendientes!Documento
        FEARendir.TextMatrix(X, 3) = rsPendientes!AREA
        FEARendir.TextMatrix(X, 4) = rsPendientes!Agencia
        FEARendir.TextMatrix(X, 5) = PstaNombre(rsPendientes!Empleado)
        FEARendir.TextMatrix(X, 6) = Format(rsPendientes!Fecha, "DD/MM/YYYY")
        FEARendir.TextMatrix(X, 7) = Format(rsPendientes!Monto, "#,##0.00")
        FEARendir.TextMatrix(X, 8) = rsPendientes!Descripcion
        FEARendir.TextMatrix(X, 9) = rsPendientes!cAreaCod
        FEARendir.TextMatrix(X, 10) = rsPendientes!cPersCod
        FEARendir.TextMatrix(X, 11) = rsPendientes!nMovNroSolicitud
        FEARendir.TextMatrix(X, 13) = rsPendientes!cAgeCod
        rsPendientes.MoveNext
    Next
Else
 MsgBox "No existen Solicitudes Pendientes..!!", vbOKOnly + vbInformation, "Atención"
End If
'***********Agregado by NAGL ERS075-2017 20171120********
Set rsPendientes = Nothing
Set oNArendir = Nothing
End Sub

Private Sub actualizarArendir()
If FEARendir.TextMatrix(1, 2) <> "" Then 'NAGL Cambió de (1,1) a (1,2) y (FEARendir.Row, 7) a (FEARendir.Row, 8) 20171120
    txtMovDesc = FEARendir.TextMatrix(FEARendir.Row, 8)
Else
    txtMovDesc = ""
End If
End Sub


Private Sub chkTodo_Click()
If chkTodo.value = 0 Then
    FraSeleccion.Enabled = True
    txtBuscarAgenciaArea.SetFocus
    LimpiaFlex FEARendir
    txtMovDesc = ""
Else
    FraSeleccion.Enabled = False
    txtBuscarAgenciaArea.Text = ""
    lblAgenciaArea = ""
    lblAgeDesc = ""
    LimpiaFlex FEARendir
    txtMovDesc = ""
  End If
End Sub

Private Sub chkTodo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdBuscar.SetFocus
End If
End Sub

Private Sub cmdAprobar_Click()
Dim oNArendir As NARendir
Set oNArendir = New NARendir
Dim oDOperacion As DOperacion
Set oDOperacion = New DOperacion
Dim oNContFunciones As NContFunciones
Set oNContFunciones = New NContFunciones
'Dim lsCtaConITFD As String
'Dim lsCtaConITFH As String
'Dim lsGlosa As String 'Comentado by NAGL 20171122
Dim lsMovNro As String
Dim lsPersCod As String
Dim lnImporte As Currency
'Dim lsTpoIf As String
'Dim lsPersCodIf As String
'Dim lsCtaBanco As String 'Comentado by NAGL 20171122
Dim lsDocNro As String
Dim lsFecha As String
Dim lnMovAnt As Long
Dim lsCuentaAho As String
'Dim rsBilletaje As ADODB.Recordset
'Set rsBilletaje = New ADODB.Recordset
'Dim lsCtaContDebe As String
'Dim lsCtaContHaber As String 'Comentado by NAGL 20171122
Dim lsNroDocViatico As String

'*****NAGL ERS075-2017***
Dim I, nFilasTotal As Integer, nCantAp As Integer, Iopc As Integer
Dim lsCtaAhorro As String, smsAprob As String
Dim nItem As Integer
Dim cMovNroCadena As String
I = 1
'****END NAGL 20171122***

'If FEARendir.TextMatrix(FEARendir.Row, 9) = gsCodPersUser Then
'    MsgBox "Usted no puede aprobar su propia Solicitud", vbInformation, "Aviso"
'    cmdAprobar.SetFocus
'    Exit Sub
'End If 'Comentado by NAGL 20171122

If gsOpeCod = gCGArendirViatAprobMN Or gsOpeCod = gCGArendirViatAprobME Then
    lnArendirFase = ArendirAtencion
    lnTipoArendir = gArendirTipoViaticos
ElseIf gsOpeCod = gCGArendirCtaAprobMN Or gsOpeCod = gCGArendirCtaAprobME Then
    lnArendirFase = ArendirAtencion
    lnTipoArendir = gArendirTipoCajaGeneral
End If

'**********************BEGIN NAGL ERS075-2017 20171121*****************************
nFilasTotal = FEARendir.Rows - 1
If FEARendir.TextMatrix(1, 2) = "" Then
    MsgBox "No se existe ninguna solicitud para aprobar..!!", vbInformation, "!Aviso¡"
    Exit Sub
End If

If ValidaRegistros("A") Then
    For I = 1 To nFilasTotal
        If FEARendir.TextMatrix(I, 1) = "." And FEARendir.TextMatrix(I, 14) <> "" Then
           nCantAp = nCantAp + 1
           Iopc = I
        End If
    Next I
    If nCantAp = 1 Then
       smsAprob = "Desea aprobar la Solicitud " & Trim(FEARendir.TextMatrix(Iopc, 2)) & "..?"
    Else
       smsAprob = "Seguro de Aprobar las Solicitudes Seleccionadas..?"
    End If
    
    If MsgBox(smsAprob, vbYesNo + vbQuestion, "Aviso") = vbYes Then
        cmdEliminar.Enabled = False '***Agregado por ELRO el 20120929, según TIC1209270003
        cmdEditar.Enabled = False '***Agregado por ELRO el 20120929, según TIC1209270003
        cmdAprobar.Caption = "Procesando" '***Agregado by NAGL ERS075-2017 20171127
        cmdAprobar.Enabled = False
        cmdSalir.Enabled = False '***Agregado por ELRO el 20120929, según TIC1209270003
        
        For I = 1 To nFilasTotal
            If FEARendir.TextMatrix(I, 1) = "." And FEARendir.TextMatrix(I, 14) <> "" Then
                lsPersCod = FEARendir.TextMatrix(I, 10)
                lnImporte = CCur(FEARendir.TextMatrix(I, 7))
                lnMovAnt = FEARendir.TextMatrix(I, 11)
                lsCtaAhorro = FEARendir.TextMatrix(I, 14)
                lsDocNro = Mid(FEARendir.TextMatrix(I, 2), 5, Len(FEARendir.TextMatrix(I, 2)))
                lsFecha = gdFecSis
                
                If lnTipoArendir = gArendirTipoViaticos Then
                    lsNroDocViatico = FEARendir.TextMatrix(I, 2)
                Else
                    lsNroDocViatico = ""
                End If
                
                lsMovNro = oNContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'cMovnro
                
                If oNArendir.GrabaAtencionArendirNew(lnTipoArendir, lsFecha, lsMovNro, gsOpeCod, lsDocNro, txtMovDesc, lsPersCod, lnImporte, lnMovAnt, lsCtaAhorro, gbBitCentral) = 0 Then
                    cMovNroCadena = cMovNroCadena & lsMovNro & ","
                Else
                    MsgBox "No se aprobó la Solicitud.", vbInformation, "!Aviso¡"
                    cmdAprobar.Caption = "Aceptar"
                    Exit Sub
                End If
                 nItem = nItem + 1
            End If
        Next I
        
    'Se Imprime los Asientos Correspondientes
    If gsOpeCod = "401125" Or gsOpeCod = "402125" Or gsOpeCod = "401225" Or gsOpeCod = "402225" Then
            'MsgBox "Se ha aprobado la solicitud !!", vbInformation, "Aviso"
        If nItem > 1 Then
                ImprimeAsientoContable cMovNroCadena, "17", , , , , , , , , , , , , , , , , , , , True
            Else
                cMovNroCadena = Replace(cMovNroCadena, ",", "")
                ImprimeAsientoContable cMovNroCadena, "17", , , , , , , , , , , , , , , , , , , , False
            End If
    End If
'********************END***********************************************************
        'ARLO20170208
        Set objPista = New COMManejador.Pista
        'gsOpeCod = LogPistaCierreDiarioCont
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Aprobo A Rendir N° " & FEARendir.TextMatrix(FEARendir.Row, 2)
        Set objPista = Nothing
        '*******
        I = 1
        '*****NAGL ERS0752017 20171129****
        Do While I <= nFilasTotal
            If oNArendir.ExistProcesoPost(FEARendir.TextMatrix(I, 11)) Then
                FEARendir.EliminaFila I
                If (I - 1) > 0 Or nFilasTotal > 1 Then
                    I = I - 1
                    nFilasTotal = nFilasTotal - 1
                End If
            End If
            I = I + 1
        Loop
        '****NAGL ERS0752017 20171129*****
            
        'FEARendir.EliminaFila FEARendir.Row
        cmdEliminar.Enabled = True '***Agregado por ELRO el 20120929, según TIC1209270003
        cmdEditar.Enabled = True '***Agregado por ELRO el 20120929, según TIC1209270003
        cmdAprobar.Caption = "Aprobar" '***Agregado by NAGL ERS075-2017 20171127
        cmdAprobar.Enabled = True
        cmdSalir.Enabled = True '***Agregado por ELRO el 20120929, según TIC1209270003
        
        Call limpiarCampos
  End If
End If

'*****************Comentado by NAGL 20171127************
'lsPersCod = FEARendir.TextMatrix(FEARendir.Row, 9)
'lnImporte = CCur(FEARendir.TextMatrix(FEARendir.Row, 6))
'lnMovAnt = FEARendir.TextMatrix(FEARendir.Row, 10)
'lsTpoIf = ""
'lsPersCodIf = ""
'lsCtaBanco = ""
'Set rsBilletaje = Nothing
'lsFecha = CStr(gdFecha)
'If lnTipoArendir = gArendirTipoViaticos Then
'    lsNroDocViatico = FEARendir.TextMatrix(FEARendir.Row, 1)
'Else
'    lsNroDocViatico = ""
'End If
'lsCtaContDebe = oDOperacion.EmiteOpeCta(gsOpeCod, "D", , , ObjEntidadesFinancieras)
'lsCtaContHaber = oDOperacion.EmiteOpeCta(gsOpeCod, "H", , , ObjEntidadesFinancieras)
'If lsCtaContDebe = "" Or lsCtaContHaber = "" Then
'    MsgBox "Cuentas Contables no determinadas correctamente." & oImpresora.gPrnSaltoLinea & "consulte con sistemas", vbInformation, "Aviso"
'    Exit Sub
'End If
'lsCtaConITFD = oDOperacion.EmiteOpeCta(gsOpeCod, "D", 2)
'lsCtaConITFH = oDOperacion.EmiteOpeCta(gsOpeCod, "H", 2)
'If lsCtaConITFD = "" Or lsCtaConITFH = "" Then
'    MsgBox "Cuentas Contables ITF no determinadas correctamente." & oImpresora.gPrnSaltoLinea & "consulte con sistemas", vbInformation, "Aviso"
'    Exit Sub
'End If
'If MsgBox("Desea aprobar la Solicitud " & Trim(FEARendir.TextMatrix(FEARendir.Row, 1)), vbYesNo + vbQuestion, "Aviso") = vbYes Then
'    If oNArendir.GrabaAtencionArendir(lnTipoArendir, lsMovNro, _
'                                      gsOpeCod, txtMovDesc, _
'                                      lsCtaContDebe, lsCtaContHaber, _
'                                      lsPersCod, lnImporte, _
'                                      "", "", _
'                                      "", rsBilletaje, _
'                                      "", "", _
'                                      "", "", _
'                                      "", lnMovAnt, _
'                                      "", gbBitCentral, , _
'                                      lsCtaConITFD, lsCtaConITFH, _
'                                      gnImpITF) = 0 Then
''VAPI RS179-2014, Si es una operacion de aprobacion de A rendir o A rendir viaticos se registra como un movimiento no contable (13) YA NO IMPRIMIMOS LOS ASIENTOS
'    If gsOpeCod = "401125" Or gsOpeCod = "402125" Or gsOpeCod = "401225" Or gsOpeCod = "402225" Then
'            MsgBox "Se ha aprobado la solicitud", vbInformation, "Aviso"
'    Else
'            ImprimeAsientoContable lsMovNro, , , , , , _
'                                   txtMovDesc, lsPersCod, lnImporte, _
'                                   lnTipoArendir, lsNroDocViatico, , , , _
'                                   "17"
'    End If
''FIN VAPI
'        If FEARendir.TextMatrix(1, 0) = "" Then
'            Call limpiarCampos
'            Unload Me
'            Exit Sub
'        End If
'
'        Call limpiarCampos
'    End If
'End IF ***********************************************************
End Sub

Private Sub cmdEliminar_Click()
Dim oNArendir As NARendir
Set oNArendir = New NARendir
Dim oNContFunciones As NContFunciones
Set oNContFunciones = New NContFunciones
Dim lsMovNro As String
Dim lsOpeCodEliminar As String
Dim I As Integer, nRowsTotal As Integer, nCant As Integer, Iopc As Integer 'NAGL ERS0752017 20171129
Dim smsElim As String, smsElimFin As String 'NAGL ERS0752017 20171129
I = 1 'NAGL 20171129
On Error GoTo ErrcmdEliminar

    If gsOpeCod = gCGArendirViatAprobMN Or gsOpeCod = gCGArendirViatAprobME Then
        If gsOpeCod = gCGArendirViatAprobMN Then
            lsOpeCodEliminar = gCGArendirViatRechMN
        Else
            lsOpeCodEliminar = gCGArendirViatRechME
        End If
               
    ElseIf gsOpeCod = gCGArendirCtaAprobMN Or gsOpeCod = gCGArendirCtaAprobME Then
        If gsOpeCod = gCGArendirCtaAprobMN Then
            lsOpeCodEliminar = gCGArendirCtaRechMN
        Else
            lsOpeCodEliminar = gCGArendirCtaRechME
        End If
    End If
    
    nRowsTotal = FEARendir.Rows - 1 'NAGL ERS0752017 20171129
    
    If FEARendir.TextMatrix(1, 2) = "" Then 'NAGL Camio de (1,1) a (1,2)
        MsgBox "No existen Solicitudes para Eliminar...!", vbInformation, "Aviso"
        Exit Sub
    End If
    
'************Comentado by NAGL 20171129***********
'    If MsgBox(" ¿ Seguro de Eliminar A Rendir N° :" & FEARendir.TextMatrix(FEARendir.Row, 1) & vbCrLf & vbCrLf & "Solicitado por :" & FEARendir.TextMatrix(FEARendir.Row, 4) & vbCrLf & "De :" & FEARendir.TextMatrix(FEARendir.Row, 2) & " - " & FEARendir.TextMatrix(FEARendir.Row, 3) & vbCrLf, vbQuestion + vbYesNo, "Confirmación") = vbYes Then
'        lsMovNro = oNContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'        If oNArendir.GrabaRechazoSolARendir(lsMovNro, lsOpeCodEliminar, FEARendir.TextMatrix(FEARendir.Row, 10), "Solicitud A Rendir N° " & FEARendir.TextMatrix(FEARendir.Row, 1) & " eliminado") = 0 Then
'*************************************************

If ValidaRegistros("E") Then 'NAGL ERS0752017 20171129
    '*******NAGL ERS0752017 20171129******'
    For I = 1 To nRowsTotal
        If FEARendir.TextMatrix(I, 1) = "." Then
           nCant = nCant + 1
           Iopc = I
        End If
    Next I
    If nCant = 1 Then
       smsElim = " ¿ Seguro de Eliminar A Rendir N°: " & FEARendir.TextMatrix(Iopc, 2) & vbCrLf & vbCrLf & "Solicitado por: " & FEARendir.TextMatrix(Iopc, 5) & vbCrLf & "De: " & FEARendir.TextMatrix(Iopc, 3) & " - " & FEARendir.TextMatrix(Iopc, 4) 'NAGL Agregó una columna mas a c/columna, tambien cambio de FEARendir.Row a i
       smsElimFin = "La Solicitud N°: " & FEARendir.TextMatrix(Iopc, 2) & " fue eliminada satisfactoriamente..!"
    Else
       smsElim = " Seguro de Eliminar las Solicitudes Seleccionadas..?"
       smsElimFin = "Las Solicitudes fueron eliminadas satisfactoriamente ..!"
    End If
    '*******NAGL ERS0752017 20171129******'
    
    If MsgBox(smsElim & vbCrLf, vbQuestion + vbYesNo, "Confirmación") = vbYes Then
        For I = 1 To nRowsTotal
            If FEARendir.TextMatrix(I, 1) = "." Then
                lsMovNro = oNContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                If oNArendir.GrabaRechazoSolARendir(lsMovNro, lsOpeCodEliminar, FEARendir.TextMatrix(I, 11), "Solicitud A Rendir N° " & FEARendir.TextMatrix(I, 2) & " eliminado") = 0 Then 'NAGL Agregó una columna mas a c/columna
                    'ARLO20170208
                    Set objPista = New COMManejador.Pista
                    'gsOpeCod = LogPistaCierreDiarioCont
                    objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", Me.Caption & " Elimino A Rendir N° " & FEARendir.TextMatrix(FEARendir.Row, 2) 'NAGL Cambio (FEARendir.Row, 1) de 1 a 2
                    Set objPista = Nothing
                    '*******
                Else
                    MsgBox "No se eliminó la Solicitud.", vbInformation, "!Aviso¡"
                    Exit Sub
                End If
            End If
         Next I
        '*****NAGL ERS0752017 20171129****
        I = 1
        Do While I <= nRowsTotal
            If oNArendir.ExistProcesoPost(FEARendir.TextMatrix(I, 11)) Then
                FEARendir.EliminaFila I
                If (I - 1) > 0 Or nRowsTotal > 1 Then
                    I = I - 1
                    nRowsTotal = nRowsTotal - 1
                End If
            End If
            I = I + 1
        Loop
        MsgBox smsElimFin, vbInformation, "!Aviso¡"
        FEARendir.SetFocus
        'Me.txtMovDesc = ""
        '****NAGL ERS0752017 20171129*****
    End If
End If
    Set oNArendir = Nothing
    Set oNContFunciones = Nothing
    Exit Sub
ErrcmdEliminar:
    MsgBox "Error N° [" & Err.Number & "]" & TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Private Sub FEARendir_Click()
Call actualizarArendir
End Sub

Private Sub FEARendir_GotFocus()
Call actualizarArendir
End Sub

Private Sub FEARendir_OnRowChange(pnRow As Long, pnCol As Long)
Call actualizarArendir
End Sub
Private Sub FEARendir_RowColChange()
Call actualizarArendir
End Sub

Private Sub cmdBuscar_Click()
Call cargarPendientes
End Sub

Private Sub cmdEditar_Click()
If FEARendir.TextMatrix(1, 2) = "" Then 'NAGL Cambio de (1,1) a (1,2)
    MsgBox "No existen Solicitudes para editar...!", vbInformation, "Aviso"
    Exit Sub
End If

If gsOpeCod = CStr(gCGArendirViatAprobMN) Or gsOpeCod = CStr(gCGArendirViatAprobME) Then
    Call frmViaticosSol.iniciarEditar(Right(Trim(FEARendir.TextMatrix(FEARendir.Row, 2)), 8), IIf(Mid(gsOpeCod, 3, 1) = "1", CStr(gCGArendirViatSolcEditMN), CStr(gCGArendirViatSolcEditME))) 'NAGL Cambió de 1 a 2
    If Mid(gsOpeCod, 3, 1) = gMonedaNacional Then
        gsOpeCod = gCGArendirViatAprobMN
    Else
        gsOpeCod = gCGArendirViatAprobME
    End If
ElseIf gsOpeCod = CStr(gCGArendirCtaAprobMN) Or gsOpeCod = CStr(gCGArendirCtaAprobME) Then
    Call frmARendirSolicitud.iniciarEdicion(gArendirTipoCajaGeneral, , IIf(Mid(gsOpeCod, 3, 1) = "1", CStr(gCGArendirCtaSolcEditMN), CStr(gCGArendirCtaSolcEditME)), FEARendir.TextMatrix(FEARendir.Row, 11)) 'NAGL Cambió de 10 a 11
    If Mid(gsOpeCod, 3, 1) = gMonedaNacional Then
        gsOpeCod = gCGArendirCtaAprobMN
    Else
        gsOpeCod = gCGArendirCtaAprobME
    End If
End If
Call cargarPendientes
    'ARLO20170208
    Set objPista = New COMManejador.Pista
    'gsOpeCod = LogPistaCierreDiarioCont
    objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "2", Me.Caption & " Mofidifico A Rendir N° " & FEARendir.TextMatrix(FEARendir.Row, 1)
    Set objPista = Nothing
    '*******
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim oDActualizaDatosArea As DActualizaDatosArea
Set oDActualizaDatosArea = New DActualizaDatosArea

CentraForm Me

Me.Caption = Me.Caption & gsOpeDesc

If Mid(gsOpeCod, 3, 1) = gMonedaNacional Then
   lbMN = True
   lsMoney = gcMN
   lsSimbolo = gcMN
Else
   lbMN = False
   lsMoney = gcME
   lsSimbolo = gcME
   If gnTipCambio = 0 Then
      If Not GetTipCambio(gdFecSis) Then
         lSalir = True
         Exit Sub
      End If
   End If
End If


cmdEditar.Visible = True

txtBuscarAgenciaArea.rs = oDActualizaDatosArea.GetAgenciasAreas
Set oDActualizaDatosArea = Nothing
End Sub

Private Sub txtBuscarAgenciaArea_EmiteDatos()
Dim oDActualizaDatosArea As DActualizaDatosArea
Set oDActualizaDatosArea = New DActualizaDatosArea
lblAgenciaArea = oDActualizaDatosArea.GetNombreAreas(Mid(txtBuscarAgenciaArea, 1, 3))
lblAgeDesc = oDActualizaDatosArea.GetNombreAgencia(IIf(Trim(Mid(txtBuscarAgenciaArea, 4, 2)) = "", "01", Mid(txtBuscarAgenciaArea, 4, 2)))
If txtBuscarAgenciaArea <> "" Then
   cmdBuscar.SetFocus
Else
   txtBuscarAgenciaArea.SetFocus
End If
End Sub


Private Function ValidaRegistros(psTipo As String) As Boolean
Dim I, nRowsTotal As Integer, Cant As Integer
nRowsTotal = FEARendir.Rows - 1
Cant = 0

If (psTipo = "A") Then 'Valida para la Aprobación
    For I = 1 To nRowsTotal
        If (FEARendir.TextMatrix(I, 1) = "." And FEARendir.TextMatrix(I, 14) = "") Then
             Cant = Cant + 1
        End If
    Next I
    If Cant = nRowsTotal Or Cant > 1 Then
        MsgBox "Existen Solicitudes que no se encuentran vinculadas a alguna Cuenta de Ahorro, Imposible de Continuar.." & "!!", vbInformation, "Aviso"
        FEARendir.SetFocus
        Exit Function
    End If
    
    Cant = 0
    For I = 1 To nRowsTotal
        If (FEARendir.TextMatrix(I, 1) = "") Then
             Cant = Cant + 1
        End If
    Next I
    
    If Cant = nRowsTotal Then
        MsgBox "No se encuentra ningún check activado, Imposible de Continuar..!!", vbInformation, "Aviso"
        FEARendir.SetFocus
        Exit Function
    End If
    
    For I = 1 To nRowsTotal
        If (FEARendir.TextMatrix(I, 1) = "." And FEARendir.TextMatrix(I, 14) = "") Then
            MsgBox "Se debe seleccionar la cuenta de Ahorro, antes de activar el check, en la Solicitud " & FEARendir.TextMatrix(I, 2) & "!!", vbInformation, "Aviso"
            FEARendir.SetFocus
            Exit Function
        ElseIf FEARendir.TextMatrix(FEARendir.Row, 10) = gsCodPersUser Then
            MsgBox "Usted no puede aprobar su propia Solicitud", vbInformation, "Aviso"
            cmdAprobar.SetFocus
            Exit Function
        End If
    Next I
Else 'Valida para la Eliminación
    For I = 1 To nRowsTotal
        If FEARendir.TextMatrix(I, 1) <> "." Then
           Cant = Cant + 1
        End If
    Next I
    If Cant = nRowsTotal Then
        MsgBox "No se encuentra ningún check activado, Imposible de Eliminar..!!", vbInformation, "Aviso"
        Exit Function
    End If
End If
ValidaRegistros = True
End Function 'NAGL ERS075-2017 20171116


Private Sub FEARendir_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
Dim SelCtaAH As New frmSeleccionCtaAhorroERViaticos
Dim lsPersCod As String, lsMoneda As String
lsPersCod = FEARendir.TextMatrix(pnRow, 10)
lsMoneda = Mid(gsOpeCod, 3, 1)
Call frmSeleccionCtaAhorroERViaticos.Inicio(lsPersCod, lsMoneda)
FEARendir.TextMatrix(FEARendir.Row, 14) = frmSeleccionCtaAhorroERViaticos.lsCtaAhorro
End Sub 'NAGL ERS075-2017 20171116

'**********NAGL 20180205***********************************************'
Private Sub FEARendir_EnterCell()
  If FEARendir.col = 2 Then
        ValorCelda1 = FEARendir.TextMatrix(FEARendir.Row, FEARendir.col)
  ElseIf FEARendir.col = 3 Then
        ValorCelda2 = FEARendir.TextMatrix(FEARendir.Row, FEARendir.col)
  ElseIf FEARendir.col = 4 Then
        ValorCelda3 = FEARendir.TextMatrix(FEARendir.Row, FEARendir.col)
  ElseIf FEARendir.col = 5 Then
        ValorCelda4 = FEARendir.TextMatrix(FEARendir.Row, FEARendir.col)
  ElseIf FEARendir.col = 6 Then
        ValorCelda5 = FEARendir.TextMatrix(FEARendir.Row, FEARendir.col)
  ElseIf FEARendir.col = 7 Then
        ValorCelda6 = FEARendir.TextMatrix(FEARendir.Row, FEARendir.col)
  End If
End Sub

Private Sub FEARendir_OnCellChange(pnRow As Long, pnCol As Long)
  If (pnCol = 2) Then
        FEARendir.TextMatrix(pnRow, pnCol) = ValorCelda1
  ElseIf (pnCol = 3) Then
        FEARendir.TextMatrix(pnRow, pnCol) = ValorCelda2
  ElseIf (pnCol = 4) Then
        FEARendir.TextMatrix(pnRow, pnCol) = ValorCelda3
  ElseIf (pnCol = 5) Then
        FEARendir.TextMatrix(pnRow, pnCol) = ValorCelda4
  ElseIf (pnCol = 6) Then
        FEARendir.TextMatrix(pnRow, pnCol) = ValorCelda5
  ElseIf (pnCol = 7) Then
        FEARendir.TextMatrix(pnRow, pnCol) = ValorCelda6
  End If
End Sub
'***END NAGL 20180205***************************************************'


VERSION 5.00
Begin VB.Form frmCredAdmiGastos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administracion de Gastos"
   ClientHeight    =   6180
   ClientLeft      =   1665
   ClientTop       =   2070
   ClientWidth     =   8700
   Icon            =   "frmCredAdmiGastos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   Begin SICMACT.FlexEdit FEGastos 
      Height          =   3675
      Left            =   120
      TabIndex        =   26
      Top             =   1800
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   6482
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      EncabezadosNombres=   "Nº-OK-Cuota-Gasto-Monto-Pagado-nPrdConcepto"
      EncabezadosAnchos=   "400-400-600-3600-1400-1400-0"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-1-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-4-0-0-0-0-0"
      BackColorControl=   65535
      BackColorControl=   65535
      BackColorControl=   65535
      EncabezadosAlineacion=   "C-C-C-L-R-R-C"
      FormatosEdit    =   "0-0-0-0-2-2-0"
      AvanceCeldas    =   1
      TextArray0      =   "Nº"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483635
   End
   Begin VB.CheckBox chkTodos 
      Caption         =   "Todos"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox CboAplicado 
      Enabled         =   0   'False
      Height          =   315
      Left            =   7080
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton CmdNuevaBusq 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5535
      TabIndex        =   22
      Top             =   5640
      Width           =   1395
   End
   Begin VB.CheckBox ChkTodas 
      Caption         =   "Todas"
      Enabled         =   0   'False
      Height          =   210
      Left            =   1470
      TabIndex        =   15
      Top             =   1020
      Width           =   750
   End
   Begin VB.ComboBox CboCuotas 
      Enabled         =   0   'False
      Height          =   315
      Left            =   690
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   975
      Width           =   690
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6960
      TabIndex        =   10
      Top             =   5640
      Width           =   1530
   End
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1605
      TabIndex        =   9
      Top             =   5640
      Width           =   1530
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   60
      TabIndex        =   8
      Top             =   5640
      Width           =   1530
   End
   Begin VB.Frame FraDatos 
      Height          =   1140
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   8415
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   6075
         TabIndex        =   6
         Top             =   630
         Width           =   1215
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   345
         Left            =   6075
         TabIndex        =   5
         Top             =   255
         Width           =   1215
      End
      Begin VB.TextBox TxtMontoGasto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   825
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   630
         Width           =   1095
      End
      Begin VB.ComboBox CboGasto 
         Height          =   315
         Left            =   825
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   210
         Width           =   5145
      End
      Begin VB.Label Label2 
         Caption         =   "Monto :"
         Height          =   225
         Left            =   195
         TabIndex        =   3
         Top             =   660
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Gasto :"
         Height          =   270
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   555
      End
   End
   Begin SICMACT.ActXCodCta ActxCuenta 
      Height          =   420
      Left            =   180
      TabIndex        =   7
      Top             =   105
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   741
      Texto           =   "Credito :"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.Label Label9 
      Caption         =   "Aplicado :"
      Height          =   240
      Left            =   6345
      TabIndex        =   23
      Top             =   660
      Width           =   720
   End
   Begin VB.Label LblMoneda 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7080
      TabIndex        =   21
      Top             =   1005
      Width           =   1440
   End
   Begin VB.Label Label7 
      Caption         =   "Moneda :"
      Height          =   225
      Left            =   6345
      TabIndex        =   20
      Top             =   1013
      Width           =   690
   End
   Begin VB.Label LblSaldo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4900
      TabIndex        =   19
      Top             =   990
      Width           =   1350
   End
   Begin VB.Label Label6 
      Caption         =   "Saldo :"
      Height          =   225
      Left            =   4440
      TabIndex        =   18
      Top             =   1020
      Width           =   510
   End
   Begin VB.Label LblPrestamo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3010
      TabIndex        =   17
      Top             =   990
      Width           =   1350
   End
   Begin VB.Label Label5 
      Caption         =   "Prestamo :"
      Height          =   225
      Left            =   2330
      TabIndex        =   16
      Top             =   1020
      Width           =   870
   End
   Begin VB.Label Label4 
      Caption         =   "Cuota        :"
      Height          =   225
      Left            =   120
      TabIndex        =   14
      Top             =   1020
      Width           =   870
   End
   Begin VB.Label Label3 
      Caption         =   "Titular        :"
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   645
      Width           =   510
   End
   Begin VB.Label lblTitular 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C16A0B&
      Height          =   285
      Left            =   690
      TabIndex        =   11
      Top             =   630
      Width           =   5550
   End
End
Attribute VB_Name = "frmCredAdmiGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lAltoUp As Integer
Dim lAltoDown As Integer
Dim RGastos As ADODB.Recordset
Dim R As ADODB.Recordset
Dim nCuotaPend As Integer
Dim nDesembolsoPend As Integer
Dim objPista As COMManejador.Pista
'APRI20181121 ERS071-2018
Dim i As Integer
Dim dVigencia As Date
Dim nMontoCol As Double
Dim nTasaInteres As Double
'END APRI


Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim RT As ADODB.Recordset

    CboAplicado.Clear
    Set oCons = New COMDConstantes.DCOMConstantes
    Set RT = oCons.RecuperaConstantes(gColocCalendApl)
    Set oCons = Nothing
    Do While Not RT.EOF
        If RT!nConsValor = gColocCalendAplCuota Or RT!nConsValor = gColocCalendAplDesembolso Then
            CboAplicado.AddItem Trim(RT!cConsDescripcion) & Space(100) & Trim(RT!nConsValor)
        End If
        RT.MoveNext
    Loop
    RT.Close
    
End Sub

Private Sub CargaGastos(ByVal pnApli As ColocCalendApl)
Dim oGasto As COMDCredito.DCOMGasto
Dim RG As ADODB.Recordset

    Set oGasto = New COMDCredito.DCOMGasto
    If pnApli = gColocCalendAplCuota Then
        Set RG = oGasto.RecuperaGastosAplicablesCuotas(CInt(Mid(ActxCuenta.NroCuenta, 9, 1)), "MA", , , , True, True)
        'WIOR 20150709 AGREGO EL VALOR 'TRUE' AL PARAMETRO pbPolizaInc
        'APRI20180821 ERS061-2018 AGREGO EL VALOR 'TRUE' AL PARAMETRO pbSegMYPE
    Else
        Set RG = oGasto.RecuperaGastosAplicablesDesembolso(CInt(Mid(ActxCuenta.NroCuenta, 9, 1)), , "'MA'")
    End If
    Set oGasto = Nothing
    CboGasto.Clear
    Do While Not RG.EOF
        CboGasto.AddItem Trim(RG!CDescripcion) & Space(100) & Trim(RG!nPrdConceptoCod)
        RG.MoveNext
    Loop
End Sub

Private Sub HabilitaActualizacion(ByVal pbActualizar As Boolean)
    If pbActualizar Then
        'DGGastos.Height = lAltoUp
        FEGastos.Height = lAltoUp 'APRI20181121 ERS071-2018
    Else
        'DGGastos.Height = lAltoDown
        FEGastos.Height = lAltoDown 'APRI20181121 ERS071-2018
    End If
    'DGGastos.Enabled = Not pbActualizar
    FEGastos.Enabled = Not pbActualizar 'APRI20181121 ERS071-2018
    cmdAgregar.Enabled = Not pbActualizar
    cmdEliminar.Enabled = Not pbActualizar
    CboCuotas.Enabled = Not pbActualizar
    ChkTodas.Enabled = Not pbActualizar
    CmdSalir.Enabled = Not pbActualizar
    CmdNuevaBusq.Enabled = Not pbActualizar
    FraDatos.Enabled = pbActualizar
    FraDatos.Visible = pbActualizar
    chkTodos.value = 0 'APRI20181121 ERS071-2018
    chkTodos.Enabled = Not pbActualizar 'APRI20181121 ERS071-2018
    ChkTodas.value = 0 'APRI20181121 ERS071-2018
    If CboGasto.ListCount > 0 Then
        CboGasto.ListIndex = 0
    End If
    TxtMontoGasto.Text = "0.00"
    CboAplicado.Enabled = Not pbActualizar
    
End Sub

Private Sub CargaDatos()
Dim oCred As COMDCredito.DCOMCredito
Dim oCal As COMDCredito.DCOMCalendario
Dim i As Integer

    Set oCred = New COMDCredito.DCOMCredito
    Set R = oCred.RecuperaDatosCreditoVigente(ActxCuenta.NroCuenta, , True)
    Set oCred = Nothing
    CboCuotas.Clear
    ChkTodas.value = 0
    chkTodos.value = 0 'APRI20181121 ERS071-2018
    If R.RecordCount > 0 Then
        nDesembolsoPend = IIf(IsNull(R!nNroProxDesemb), 1, R!nNroProxDesemb)
        nCuotaPend = R!nNroProxCuota
        lblTitular.Caption = Space(1) & PstaNombre(R!cPersNombre)
        LblPrestamo.Caption = Format(R!nMontoCol, "###," & String(15, "#") & "#0.00") & " " 'WIOR 20150709 MODIFICO EL FORMATO
        LblSaldo.Caption = Format(R!nSaldo, "###," & String(15, "#") & "#0.00") & " " 'WIOR 20150709 MODIFICO EL FORMATO
        lblMoneda.Caption = Trim(R!cMoneda)
        cmdAceptar.Enabled = True
        CmdSalir.Enabled = True
        ChkTodas.Enabled = True
        CboCuotas.Enabled = True
        cmdAgregar.Enabled = True
        cmdEliminar.Enabled = True
        ActxCuenta.Enabled = False
        CboAplicado.Enabled = True
        'APRI20181121 ERS071-2018
        dVigencia = R!dFecDesemb
        nMontoCol = R!nMontoCol
        nTasaInteres = R!nTasaInteres
        'END APRI
    Else
        CboCuotas.Enabled = False
        ChkTodas.Enabled = False
        lblTitular.Caption = ""
        LblPrestamo.Caption = ""
        LblSaldo.Caption = ""
        lblMoneda.Caption = ""
        cmdAceptar.Enabled = False
        CmdSalir.Enabled = False
        ActxCuenta.Enabled = True
        CboAplicado.Enabled = False
        chkTodos.Enabled = False 'APRI20181121 ERS071-2018
        'Set DGGastos.DataSource = Nothing
        'DGGastos.Refresh
        MsgBox "No se ha Podido Encontrar el Credito ", vbInformation, "Aviso"
        Exit Sub
    End If
    CboAplicado.ListIndex = -1
    CboAplicado.ListIndex = 0
    
End Sub

Private Sub ActxCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CargaDatos
    End If
End Sub


Private Sub CboAplicado_Click()
Dim oCal As COMDCredito.DCOMCalendario
Dim i As Integer
Dim nCuotas As Integer
    
    If CboAplicado.ListIndex = -1 Then
        Exit Sub
    End If

    Set oCal = New COMDCredito.DCOMCalendario
    nCuotas = oCal.RecuperaNroCuotas(ActxCuenta.NroCuenta, CInt(Trim(Right(CboAplicado, 10))), R!nNroCalen)
    Set oCal = Nothing
    CboCuotas.Clear
    'Set DGGastos.DataSource = Nothing
    'DGGastos.Refresh
    'APRI20181025 ERS071-2018
    chkTodos.value = 0
    ChkTodas.value = 0
    FEGastos.Clear
    FEGastos.FormaCabecera
    FEGastos.rows = 2
    chkTodos.Enabled = IIf(CInt(Trim(Right(CboAplicado.Text, 10))) = 0, False, True)
    'END APRI
    If CInt(Trim(Right(CboAplicado, 10))) = gColocCalendAplCuota Then
        For i = IIf(nCuotaPend = 0, 1, nCuotaPend) To nCuotas + nCuotaPend - 1
            CboCuotas.AddItem Trim(str(i))
        Next i
    Else
        For i = nDesembolsoPend To nCuotas
            CboCuotas.AddItem Trim(str(i))
        Next i
    End If
    If CboCuotas.ListCount > 0 Then
        CboCuotas.ListIndex = 0
    End If
    
    Call CargaGastos(CInt(Trim(Right(CboAplicado, 10))))
End Sub

Private Sub CboCuotas_Click()
Dim oCal As COMDCredito.DCOMCalendario

    If Trim(CboCuotas.Text) = "" Then
        Exit Sub
    End If
    Set oCal = New COMDCredito.DCOMCalendario
    Set RGastos = oCal.RecuperaGastosCuotaDesemb(ActxCuenta.NroCuenta, R!nNroCalen, CInt(Trim(Right(CboAplicado.Text, 10))), CInt(CboCuotas.Text))
    Set oCal = Nothing
    'Set DGGastos.DataSource = RGastos
    'APRI20181121 ERS071-2018
    chkTodos.Enabled = IIf(CInt(Trim(Right(CboAplicado.Text, 10))) = 0, False, True)
    FEGastos.Clear
    FEGastos.FormaCabecera
    FEGastos.rows = 2
    If Not RGastos.EOF And Not RGastos.BOF Then
        Set FEGastos.Recordset = RGastos
    End If
    'END APRI
    
End Sub

Private Sub CboCuotas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CboCuotas_Click
    End If
End Sub


Private Sub CboGasto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtMontoGasto.SetFocus
    End If
End Sub

Private Sub ChkTodas_Click()
Dim oCal As COMDCredito.DCOMCalendario
    If ChkTodas.value = 1 Then
        Set oCal = New COMDCredito.DCOMCalendario
        Set RGastos = oCal.RecuperaGastosCuotaDesemb(ActxCuenta.NroCuenta, R!nNroCalen, gColocCalendAplCuota)
        Set oCal = Nothing
        'Set DGGastos.DataSource = RGastos
        'APRI20181121 ERS071-2018
        FEGastos.Clear
        FEGastos.FormaCabecera
        FEGastos.rows = 2
        If Not RGastos.EOF And Not RGastos.BOF Then
            Set FEGastos.Recordset = RGastos
        End If
        chkTodos.Enabled = True
        'END APRI
    Else
        CboCuotas_Click
    End If
End Sub

Private Sub chkTodos_Click()
     If chkTodos.value Then
        For i = 1 To FEGastos.rows - 1
            If FEGastos.TextMatrix(i, 5) = 0 Then
                FEGastos.TextMatrix(i, 1) = "1"
            End If
        Next i
    Else
        For i = 1 To FEGastos.rows - 1
            If FEGastos.TextMatrix(i, 1) = "." Then
                FEGastos.TextMatrix(i, 1) = " "
            End If
        Next i
    End If
End Sub

Private Sub CmdAceptar_Click()
Dim oBase As COMDCredito.DCOMCredActBD
Dim oGasto As COMDCredito.DCOMGasto
    Set oGasto = New COMDCredito.DCOMGasto
    If Trim(CboGasto.Text) = "" Then
        MsgBox "Escoja un Gasto", vbInformation, "Aviso"
        CboGasto.SetFocus
        Exit Sub
    End If
    If CDbl(TxtMontoGasto.Text) = 0 Then
        MsgBox "Monto debe ser Mayor a Cero", vbInformation, "Aviso"
        TxtMontoGasto.SetFocus
        Exit Sub
    End If
    
    '* CUSCO ****
    If ChkTodas.value = 0 Then
        If oGasto.ExisteGasto(ActxCuenta.NroCuenta, R!nNroCalen, CInt(Trim(Right(CboAplicado.Text, 10))), CInt(CboCuotas.Text), CLng(Trim(Right(CboGasto.Text, 10)))) Then
            MsgBox "Gasto ya Existe para esta Cuota ", vbInformation, "Aviso"
            CboGasto.SetFocus
            Exit Sub
        End If
        Set oGasto = Nothing
        Set oBase = New COMDCredito.DCOMCredActBD
        Call oBase.dInsertColocCalendDet(ActxCuenta.NroCuenta, R!nNroCalen, CInt(Trim(Right(CboAplicado.Text, 10))), CInt(CboCuotas.Text), CLng(Trim(Right(CboGasto.Text, 15))), CDbl(TxtMontoGasto.Text), 0#, "", False)
        'PEAC Y MAVM 26122009
        objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Inserta Gasto: " & CLng(Trim(Right(CboGasto.Text, 10))) & ", Cuota: " & CInt(CboCuotas.Text), ActxCuenta.NroCuenta, gCodigoCuenta

        Set oBase = Nothing
     Else   'Aplicar a todas las Cuotas
        Dim MatCuotas() As Integer
        Dim i As Integer
        Dim sMensaje As String
        
        ReDim MatCuotas(CboCuotas.ListCount)
        For i = 0 To CboCuotas.ListCount - 1
            MatCuotas(i) = CInt(CboCuotas.List(i))
        Next
        Set oBase = New COMDCredito.DCOMCredActBD
        sMensaje = oBase.InsercionGastosxCuotaLote(ActxCuenta.NroCuenta, R!nNroCalen, CInt(Trim(Right(CboAplicado.Text, 10))), MatCuotas, CLng(Trim(Right(CboGasto.Text, 15))), CDbl(TxtMontoGasto.Text))
        
        ''*** PEAC 20090126
        objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Inserta Gasto: " & CLng(Trim(Right(CboGasto.Text, 15))) & ", Cuota: Todas", ActxCuenta.NroCuenta, gCodigoCuenta
                
        Set oBase = Nothing
        If sMensaje <> "" Then
            MsgBox sMensaje, vbInformation, "Mensaje"
        End If
     End If
     '**********
     ActualizarTCEA ActxCuenta.NroCuenta 'APRI20181121 ERS071-2018
     HabilitaActualizacion False
     'CUSCO
     Call ChkTodas_Click
End Sub

Private Sub cmdAgregar_Click()
    If Trim(CboCuotas.Text) = "" Then
        MsgBox "Debe Seleccionar Una Cuota o Un Desembolso", vbInformation, "Aviso"
        Exit Sub
    End If
    HabilitaActualizacion True
End Sub

Private Sub cmdCancelar_Click()
    HabilitaActualizacion False
End Sub

Private Sub CmdEliminar_Click()
Dim oBase As COMDCredito.DCOMCredActBD
Dim nRegistros As Integer 'APRI20181121 ERS071-2018
    
    'APRI20181121 ERS071-2018
    nRegistros = 0
    For i = 1 To FEGastos.rows - 1
        If FEGastos.TextMatrix(i, 1) = "." Then
            nRegistros = nRegistros + 1
        End If
    Next i
    'END APRI
    'If RGastos.RecordCount > 0 Then
    If nRegistros > 0 Then 'APRI20181121 ERS071 - 2018
        'If RGastos!nMontoPagado > 0 Then
        '    MsgBox "El Gasto No se Puede Eliminar Porque posee una Amortizacion", vbInformation, "Aviso"
        '    Exit Sub
        'End If
        
        If MsgBox("Se va a Eliminar el Gasto, Desea Continuar?", vbInformation + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
            Exit Sub
        End If
        Set oBase = New COMDCredito.DCOMCredActBD
        'Call oBase.dDeleteColocCalendDet(ActxCuenta.NroCuenta, R!nNroCalen, CInt(Trim(Right(CboAplicado.Text, 10))), RGastos!nCuota, RGastos!nPrdConceptoCod)
        
        'objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Eliminacion del Gasto: " & RGastos!nPrdConceptoCod & ", Cuota: " & RGastos!nCuota, ActxCuenta.NroCuenta, gCodigoCuenta
        
         'APRI20181121 ERS071-2018
         For i = 1 To FEGastos.rows - 1
            If FEGastos.TextMatrix(i, 1) = "." Then
                Call oBase.dDeleteColocCalendDet(ActxCuenta.NroCuenta, R!nNroCalen, CInt(Trim(Right(CboAplicado.Text, 10))), FEGastos.TextMatrix(i, 2), FEGastos.TextMatrix(i, 6))
        
                objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Eliminacion del Gasto: " & FEGastos.TextMatrix(i, 6) & ", Cuota: " & FEGastos.TextMatrix(i, 2), ActxCuenta.NroCuenta, gCodigoCuenta
            End If
        Next i
         ActualizarTCEA ActxCuenta.NroCuenta
         'END APRI
        Set oBase = Nothing
       
        'CUSCO
        Call ChkTodas_Click
    Else
        MsgBox "No se ha seleccionado ningún registro", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdNuevaBusq_Click()
    CboCuotas.Enabled = False
    ChkTodas.Enabled = False
    lblTitular.Caption = ""
    LblPrestamo.Caption = ""
    LblSaldo.Caption = ""
    lblMoneda.Caption = ""
    cmdAceptar.Enabled = False
    CmdSalir.Enabled = False
    ActxCuenta.Enabled = True
    ActxCuenta.NroCuenta = ""
    ActxCuenta.CMAC = gsCodCMAC
    ActxCuenta.Age = gsCodAge
    'Set DGGastos.DataSource = Nothing
    'DGGastos.Refresh
    FormateaFlex FEGastos 'APRI20181121 ERS071-2018
    chkTodos.Enabled = False 'APRI20181121 ERS071-2018
    CboAplicado.Enabled = False
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    ActxCuenta.SetFocusProd
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And ActxCuenta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.inicia(gColPYMEEmp, False)
        If sCuenta <> "" Then
            ActxCuenta.NroCuenta = sCuenta
            ActxCuenta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    lAltoUp = 2595
    lAltoDown = 3720
    CentraForm Me
    ActxCuenta.CMAC = gsCodCMAC
    ActxCuenta.Age = gsCodAge
    Call CargaControles
    
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCredAdministrarGastos
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub

Private Sub TxtMontoGasto_GotFocus()
    fEnfoque TxtMontoGasto
End Sub

Private Sub TxtMontoGasto_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtMontoGasto, KeyAscii)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub

Private Sub TxtMontoGasto_LostFocus()
    If Trim(TxtMontoGasto.Text) = "" Then
        TxtMontoGasto.Text = "0.00"
    End If
    TxtMontoGasto.Text = Format(TxtMontoGasto.Text, "#0.00")
End Sub

Public Sub Inicio(ByVal psCtaCod As String)
    Me.ActxCuenta.NroCuenta = psCtaCod
    CargaDatos
    CmdNuevaBusq.Visible = False
    Me.Show 1
    
End Sub
'APRI20181121 ERS071-2018
Private Sub ActualizarTCEA(ByVal pcCtaCod As String)
    Dim oCredito As COMDCredito.DCOMCredito
    Dim oCalend As COMDCredito.DCOMCalendario
    Dim oBase As COMDCredito.DCOMCredActBD
    Dim oCred As COMNCredito.NCOMCredito
    
    Dim NewTCEA As Double
    Dim rsTipPeriodo As ADODB.Recordset
    Dim R As ADODB.Recordset
    Dim nTipoPeriodo As Integer
    Dim MatCalend As Variant
    
    Set oCredito = New COMDCredito.DCOMCredito
    Set oCalend = New COMDCredito.DCOMCalendario
    Set oBase = New COMDCredito.DCOMCredActBD
    Set oCred = New COMNCredito.NCOMCredito
          
    Set rsTipPeriodo = oCredito.IdentificarTipoPeriodo(pcCtaCod)
    If Not (rsTipPeriodo.EOF And rsTipPeriodo.BOF) Then
        nTipoPeriodo = rsTipPeriodo!nTpPeriodo
    End If
    
    Set R = oCalend.ObtenerCalendarioCredito(pcCtaCod)
  
    ReDim MatCalend(R.RecordCount, 3)
    
    If R.RecordCount > 0 Then
        Do While Not R.EOF
            MatCalend(R.Bookmark - 1, 0) = Format(R!dVenc, "dd/mm/yyyy")
            MatCalend(R.Bookmark - 1, 1) = Trim(str(R!nCuota))
            MatCalend(R.Bookmark - 1, 2) = Trim(str(R!nMontoCuota))
            R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
    
      
    NewTCEA = oCred.GeneraTasaCostoEfectivoAnual(dVigencia, nMontoCol, MatCalend, nTasaInteres, pcCtaCod, nTipoPeriodo)  'Para calcular la TCEA
    
    If NewTCEA > 0 Then
            Call oBase.ActualizarTCEA(pcCtaCod, NewTCEA)
    End If
    
    oBase.ExonerarSeguroTrama pcCtaCod
    
    Set rsTipPeriodo = Nothing
    Set oCalend = Nothing
    Set oCredito = Nothing
    Set oCred = Nothing
    Set oBase = Nothing
    
End Sub

Private Sub FEGastos_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    If FEGastos.TextMatrix(pnRow, 5) = 0 Then
        If FEGastos.TextMatrix(pnRow, 1) = "." Then
            FEGastos.TextMatrix(pnRow, 1) = "1"
        Else
            FEGastos.TextMatrix(pnRow, 1) = " "
        End If
    Else
        MsgBox "No se puede seleccionar porque el gasto posee una Amortizacion", vbInformation, "Aviso"
        FEGastos.TextMatrix(pnRow, 1) = " "
    End If
End Sub

'END APRI

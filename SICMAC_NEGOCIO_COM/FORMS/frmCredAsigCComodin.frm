VERSION 5.00
Begin VB.Form frmCredAsigCComodin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignar Cuota Comodin"
   ClientHeight    =   4155
   ClientLeft      =   1860
   ClientTop       =   2820
   ClientWidth     =   9000
   Icon            =   "frmCredAsigCComodin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   5490
      TabIndex        =   7
      Top             =   3630
      Width           =   1110
   End
   Begin VB.CommandButton CmdDeshacer 
      Caption         =   "&Deshacer"
      Enabled         =   0   'False
      Height          =   390
      Left            =   1230
      TabIndex        =   6
      Top             =   3645
      Width           =   1110
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   7770
      TabIndex        =   4
      Top             =   3630
      Width           =   1110
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Grabar"
      Enabled         =   0   'False
      Height          =   390
      Left            =   6615
      TabIndex        =   3
      Top             =   3630
      Width           =   1110
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Enabled         =   0   'False
      Height          =   390
      Left            =   90
      TabIndex        =   2
      Top             =   3645
      Width           =   1110
   End
   Begin SICMACT.ActXCodCta ActxCta 
      Height          =   525
      Left            =   165
      TabIndex        =   0
      Top             =   135
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   926
      Texto           =   "Credito :"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin SICMACT.FlexEdit FECalBPag 
      Height          =   2880
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   5080
      Cols0           =   8
      FixedCols       =   0
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Fecha Venc.-Cuota-Cuotas-Capital-Interes-Int. Gracia-Gastos-Saldo Capital"
      EncabezadosAnchos=   "1000-600-1200-1000-1000-1000-1000-1200"
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
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-C-R-R-R-R-R-R"
      FormatosEdit    =   "0-0-2-3-2-2-2-2"
      TextArray0      =   "Fecha Venc."
      SelectionMode   =   1
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   1005
      RowHeight0      =   300
   End
   Begin VB.Label LblTitular 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3855
      TabIndex        =   5
      Top             =   195
      Width           =   5010
   End
End
Attribute VB_Name = "frmCredAsigCComodin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MatCal As Variant
Dim MatCalTmp As Variant
Dim nPlazo As Integer
Dim nTipoCuota As Integer

Private Sub CargaMatrizFlex()
Dim i As Integer
    LimpiaFlex FECalBPag
    For i = 0 To UBound(MatCal) - 1
        Call FECalBPag.AdicionaFila(, , True)
        FECalBPag.TextMatrix(i + 1, 0) = MatCal(i, 0)  'Fecha
        FECalBPag.TextMatrix(i + 1, 1) = MatCal(i, 1)  'Nro Cuota
        FECalBPag.TextMatrix(i + 1, 2) = Format(CDbl(MatCal(i, 3)) + CDbl(MatCal(i, 4)) + CDbl(MatCal(i, 5)) + CDbl(MatCal(i, 9)), "#0.00") '
        FECalBPag.TextMatrix(i + 1, 3) = MatCal(i, 3)  'Capital
        FECalBPag.TextMatrix(i + 1, 4) = MatCal(i, 4)  'Interes
        FECalBPag.TextMatrix(i + 1, 5) = MatCal(i, 5)  'Interes Gracia
        FECalBPag.TextMatrix(i + 1, 6) = MatCal(i, 9)  'Gastos
        FECalBPag.TextMatrix(i + 1, 7) = MatCal(i, 10) 'Saldo
    Next i
    
End Sub

Private Sub HabilitaActualizar(ByVal pbAct As Boolean)
    ActxCta.Enabled = Not pbAct
    CmdAplicar.Enabled = pbAct
    CmdGrabar.Enabled = pbAct
    FECalBPag.Enabled = pbAct
    CmdDeshacer.Enabled = Not pbAct
End Sub

Private Sub LimpiaPantalla()
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    LblTitular.Caption = ""
    Call LimpiaFlex(FECalBPag)
    Call HabilitaActualizar(False)
End Sub

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
'Dim odCalend As COMNCredito.NCOMCredito
'Dim oDCred As COMDCredito.DCOMCredito
Dim oCred As COMNCredito.NCOMCredito
Dim sTitular As String
Dim sMensaje As String
'Dim R As ADODB.Recordset

    If KeyAscii = 13 Then
        'Set oDCred = New COMDCredito.DCOMCredito
        'Set R = oDCred.RecuperaDatosCreditoVigente(ActxCta.NroCuenta)
        'Set oDCred = Nothing
        Set oCred = New COMNCredito.NCOMCredito
        Call oCred.CargarDatosCComodin(ActxCta.NroCuenta, sTitular, nTipoCuota, nPlazo, MatCal, sMensaje)
        Set oCred = Nothing
        
        LimpiaFlex FECalBPag
        
        If sMensaje <> "" Then
            MsgBox sMensaje, vbInformation, "Aviso"
            HabilitaActualizar False
            Call LimpiaPantalla
            Exit Sub
        End If
        
        LblTitular.Caption = PstaNombre(sTitular)
        Call CargaMatrizFlex
        HabilitaActualizar True
        
        'If R.RecordCount > 0 Then
        '    If R!bCuotaCom = 1 Then
        '        LblTitular.Caption = PstaNombre(R!cPersNombre)
        '        R.Close
        '        Set oDCred = New COMDCredito.DCOMCredito
        '        Set R = oDCred.RecuperaColocacEstado(ActxCta.NroCuenta, gColocEstAprob)
        '        Set oDCred = Nothing
        '        nTipoCuota = IIf(IsNull(R!nPeriodoFechaFija), 0, R!nPeriodoFechaFija)
        '        nPlazo = IIf(IsNull(R!nPlazo), 0, R!nPlazo)
        '        R.Close
        '        Set odCalend = New COMNCredito.NCOMCredito
        '        MatCal = odCalend.RecuperaMatrizCalendarioPendiente(ActxCta.NroCuenta)
        '        Set odCalend = Nothing
        '        Call CargaMatrizFlex
        '        HabilitaActualizar True
        '    Else
        '        If R!bCuotaCom = 2 Then
        '            MsgBox "Cuota Comodin Ya fue Asignada para este Credito", vbInformation, "Aviso"
        '        Else
        '            MsgBox "Credito No es de Tipo Cuota Comodin", vbInformation, "Aviso"
        '        End If
        '        HabilitaActualizar False
        '        Call LimpiaPantalla
        '        R.Close
        '        Exit Sub
        '    End If
        'Else
        '    HabilitaActualizar False
        '    MsgBox "No se encuentra el Credito, Verifique que este Vigente", vbInformation, "Aviso"
        '    Call LimpiaPantalla
        '    R.Close
        '    Exit Sub
        'End If
    End If
End Sub

Private Sub CmdAplicar_Click()
Dim i As Integer
Dim oDCred As COMNCredito.NCOMCredito
        
        If CDate(FECalBPag.TextMatrix(FECalBPag.Row, 0)) < gdFecSis Then
            MsgBox "La cuota ya esta Vencida, Seleccione Otra Cuota", vbInformation, "Aviso"
            Exit Sub
        End If
        
        MatCalTmp = MatCal
        Set oDCred = New COMNCredito.NCOMCredito
        MatCal = oDCred.ReprogramarCreditoSoloFechasMemoria(FECalBPag.Row, nTipoCuota, nPlazo, MatCal)
        Set oDCred = Nothing
        Call CargaMatrizFlex
        CmdAplicar.Enabled = False
        CmdDeshacer.Enabled = True
End Sub

Private Sub cmdCancelar_Click()
    LimpiaPantalla
    CmdDeshacer.Enabled = False
End Sub

Private Sub CmdDeshacer_Click()
    MatCal = MatCalTmp
    Call CargaMatrizFlex
    CmdAplicar.Enabled = True
    CmdDeshacer.Enabled = False
End Sub

Private Sub cmdGrabar_Click()
Dim oCred As COMDCredito.DCOMCredActBD
    If MsgBox("Se va a Actualizar el Calendario de Pagos, Desea Continuar ???", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    Set oCred = New COMDCredito.DCOMCredActBD
    'Call oCred.ActualizaFechaCalendarioMatriz(ActxCta.NroCuenta, MatCal)
    'Call oCred.dUpdateColocacCred(ActxCta.NroCuenta, , , , , , , , , , , , , , , , , , False, 2)
    Call oCred.GrabarDatosCComodin(ActxCta.NroCuenta, MatCal)
    Set oCred = Nothing
    Call LimpiaPantalla
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And ActxCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColPYMEEmp, False)
        If sCuenta <> "" Then
            ActxCta.NroCuenta = sCuenta
            ActxCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    CentraForm Me
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
End Sub

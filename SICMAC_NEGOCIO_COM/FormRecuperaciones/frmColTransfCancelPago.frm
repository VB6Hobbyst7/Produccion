VERSION 5.00
Begin VB.Form frmColTransfCancelPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recuperaciones - Distribucón de Créditos Transferidos"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7590
   Icon            =   "frmColTransfCancelPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Crédito"
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
      Height          =   1815
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   7335
      Begin VB.CheckBox chkCancelCredito 
         Caption         =   "Cancelar Crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   39
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   360
         Left            =   6090
         TabIndex        =   31
         Top             =   360
         Width           =   1095
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   465
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   820
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblMetLiquid 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6240
         TabIndex        =   38
         Top             =   990
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Met.Liquid."
         Height          =   195
         Index           =   4
         Left            =   5400
         TabIndex        =   37
         Top             =   1005
         Width           =   780
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H8000000E&
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
         Height          =   285
         Left            =   930
         TabIndex        =   36
         Top             =   960
         Width           =   4305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   990
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Ing. Recup."
         Height          =   195
         Index           =   6
         Left            =   5400
         TabIndex        =   34
         Top             =   1395
         Width           =   840
      End
      Begin VB.Label lblIngRecup 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6240
         TabIndex        =   33
         Top             =   1350
         Width           =   930
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Distribución"
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
      Height          =   3735
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   7335
      Begin VB.TextBox txtGlosa 
         Height          =   615
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2880
         Width           =   6855
      End
      Begin VB.CheckBox chkAut 
         Caption         =   "Con autorización de Gerencia"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   2610
         Visible         =   0   'False
         Width           =   2415
      End
      Begin SICMACT.EditMoney AXMontos 
         Height          =   285
         Index           =   0
         Left            =   3720
         TabIndex        =   6
         Top             =   480
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney AXMontos 
         Height          =   285
         Index           =   1
         Left            =   3720
         TabIndex        =   7
         Top             =   780
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney AXMontos 
         Height          =   285
         Index           =   2
         Left            =   3720
         TabIndex        =   8
         Top             =   1080
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney AXMontos 
         Height          =   285
         Index           =   3
         Left            =   3720
         TabIndex        =   9
         Top             =   1380
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Monto a Canc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   3720
         TabIndex        =   29
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Deuda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   2520
         TabIndex        =   28
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Gastos:"
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   27
         Top             =   1410
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Int. Compensatorio:"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   26
         Top             =   825
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Int. Moratorio:"
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   25
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label lblCapitalD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
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
         Height          =   265
         Left            =   2160
         TabIndex        =   24
         Top             =   480
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Capital:"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   23
         Top             =   525
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descuento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   5400
         TabIndex        =   22
         Top             =   240
         Width           =   930
      End
      Begin VB.Label lblIntCompD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
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
         Height          =   265
         Left            =   2160
         TabIndex        =   21
         Top             =   780
         Width           =   1260
      End
      Begin VB.Label lblIntMoratD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
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
         Height          =   265
         Left            =   2160
         TabIndex        =   20
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label lblGastosD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
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
         Height          =   265
         Left            =   2160
         TabIndex        =   19
         Top             =   1380
         Width           =   1260
      End
      Begin VB.Label lblTotalD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
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
         Height          =   285
         Left            =   2160
         TabIndex        =   18
         Top             =   1980
         Width           =   1260
      End
      Begin VB.Label lblCapitalDS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5280
         TabIndex        =   17
         Top             =   480
         Width           =   1260
      End
      Begin VB.Label lblIntCompDS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5280
         TabIndex        =   16
         Top             =   780
         Width           =   1260
      End
      Begin VB.Label lblIntMoratDS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5280
         TabIndex        =   15
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label lblGastosDS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5280
         TabIndex        =   14
         Top             =   1380
         Width           =   1260
      End
      Begin VB.Label lblTotalP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
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
         Height          =   285
         Left            =   3720
         TabIndex        =   13
         Top             =   1980
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   1440
         TabIndex        =   12
         Top             =   2040
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Glosa"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   11
         Top             =   2625
         Width           =   405
      End
      Begin VB.Label lblTotalDS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Height          =   285
         Left            =   5280
         TabIndex        =   10
         Top             =   1980
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
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
      Height          =   390
      Left            =   4080
      TabIndex        =   2
      Top             =   5880
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
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
      Height          =   390
      Left            =   6420
      TabIndex        =   1
      Top             =   5880
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancelar 
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
      Height          =   390
      Left            =   5250
      TabIndex        =   0
      Top             =   5880
      Width           =   990
   End
End
Attribute VB_Name = "frmColTransfCancelPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fnSaldoCap As Currency, fnSaldoIntComp As Currency, fnSaldoIntMorat As Currency, fnSaldoGasto As Currency
Dim fnNewSaldoCap As Currency
Dim fnNroCalen As Integer
Dim fmMatGastos As Variant

Dim fnFechaIngRecup As Date

Dim fnPorcComision As Double
Dim fsFecUltPago As String
Dim fsCondicion As String, fsDemanda As String
Dim fsCancSKMayorCero As String

Dim objPista As COMManejador.Pista
Private Sub Form_Load()
    LimpiarCampos
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
    Call HabilitaControles(False, False, True, False)
End Sub
Private Sub cmdBuscar_Click()
    Dim loPers As comdpersona.UCOMPersona
    Dim lsPersCod As String, lsPersNombre As String
    Dim lsEstados As String
    Dim loPersCredito  As COMDColocRec.DCOMColRecCredito
    Dim lrCreditos As New ADODB.Recordset
    Dim loCuentas As comdpersona.UCOMProdPersona
    
    On Error GoTo ControlError
    
    Set loPers = New comdpersona.UCOMPersona
        Set loPers = frmBuscaPersona.Inicio
        If Not loPers Is Nothing Then
            lsPersCod = loPers.sPersCod
            lsPersNombre = loPers.sPersNombre
        Else
            Exit Sub
        End If
    Set loPers = Nothing
    
    ' Selecciona Estados
    lsEstados = gColocEstTransferido
    
    If Trim(lsPersCod) <> "" Then
        Set loPersCredito = New COMDColocRec.DCOMColRecCredito
            Set lrCreditos = loPersCredito.dObtieneCreditosDePersona(lsPersCod, lsEstados)
        Set loPersCredito = Nothing
    End If
    
    Set loCuentas = New comdpersona.UCOMProdPersona
        Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrCreditos)
        If loCuentas.sCtaCod <> "" Then
            AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
            AXCodCta.SetFocusCuenta
        End If
    Set loCuentas = Nothing
Exit Sub
ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub
Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BuscaCredito (AXCodCta.NroCuenta)
End Sub
Private Sub BuscaCredito(ByVal psCtaCod As String)
    Dim lbOk As Boolean
    Dim lrValida As ADODB.Recordset
    Dim loValCredito As COMNColocRec.NColRecValida
    Dim lrDatCredito As ADODB.Recordset
    Dim lrDatGastos As New ADODB.Recordset
    Dim loValCred As COMDColocRec.DCOMColRecCredito
    Dim loCredRec As COMDColocRec.DCOMColRecCredito
    Dim lrCIMG As ADODB.Recordset
    Dim lnDiasUltTrans As Integer
    'On Error GoTo ControlError
    
    Dim lsmensaje As String

    'valida Contrato
    Set loValCredito = New COMNColocRec.NColRecValida
    Set lrValida = loValCredito.CredTransfValidaCambioMetodoLiquid(psCtaCod, lsmensaje)
    If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        Exit Sub
    End If
    'Carga Datos
    Set loValCred = New COMDColocRec.DCOMColRecCredito
        Set lrDatCredito = loValCred.CredTransfObtieneDatosCancelacion(psCtaCod, lsmensaje)
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
        End If
    Set loValCred = Nothing
    
    If lrDatCredito Is Nothing Then   ' Hubo un Error
        MsgBox "No se Encontro el Credito o No se ha realizado registro de expediente", vbInformation, "Aviso"
        LimpiarCampos
        Set lrDatCredito = Nothing
        Exit Sub
    End If
    
    Set loCredRec = New COMDColocRec.DCOMColRecCredito
        Set lrDatGastos = loCredRec.dObtieneListaGastosxCredito(psCtaCod, lsmensaje, True)
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
        End If
    Set loCredRec = Nothing
    
    Dim nMontoSaldo As Double
    '*** Carga Gastos en Matriz
    Dim i As Integer
    ReDim fmMatGastos(0)
    ReDim fmMatGastos(lrDatGastos.RecordCount, 11)
    Do While Not lrDatGastos.EOF
        If lrDatGastos!nColocRecGastoEstado = gColRecGastoEstPendiente Then
            fmMatGastos(i, 1) = lrDatGastos!nNroGastoCta
            fmMatGastos(i, 2) = lrDatGastos!nMonto
            fmMatGastos(i, 3) = lrDatGastos!nMontoPagado
            fmMatGastos(i, 4) = lrDatGastos!nColocRecGastoEstado
            fmMatGastos(i, 5) = "N" ' Estado del Gasto
            fmMatGastos(i, 6) = 0 '(fmMatGastos(i, 2) - fmMatGastos(i, 3)) 'avmm 0  ' Monto a Cubrir del Gasto
            fmMatGastos(i, 7) = lrDatGastos!nPrdConceptoCod
            nMontoSaldo = nMontoSaldo + (fmMatGastos(i, 2) - fmMatGastos(i, 3))
            i = i + 1
        End If
        lrDatGastos.MoveNext
    Loop
    fnSaldoGasto = nMontoSaldo
        
        fnSaldoCap = lrValida!nSaldo
        fnSaldoIntComp = lrValida!nSaldoIntComp
        fnSaldoIntMorat = lrValida!nSaldoIntMor
        fnSaldoGasto = nMontoSaldo
        fsFecUltPago = CDate(fgFechaHoraGrab(lrValida!cUltimaActualizacion))
        lnDiasUltTrans = CDate(Format(gdFecSis, "dd/mm/yyyy")) - CDate(Format(fsFecUltPago, "dd/mm/yyyy"))
        
        'Muestra Datos
        Me.lblCliente.Caption = PstaNombre(Trim(lrDatCredito!cPersNombre))
        Me.lblMetLiquid.Caption = lrDatCredito!cMetLiquid
        fnFechaIngRecup = Format(lrDatCredito!dFecTransf, "dd/MM/yyyy")
        Me.lblIngRecup = fnFechaIngRecup
        
    Set lrDatCredito = Nothing
    
    Call HabilitaControles(True, True, True, True)
    
    'Obtiene los montos grabados de la misma fecha
    Set loCredRec = New COMDColocRec.DCOMColRecCredito
    Set lrCIMG = loCredRec.CredTransfObtieneDistribucionCIMGCobranza(psCtaCod, gdFecSis)
    Set loCredRec = Nothing
    If Not lrCIMG.EOF Then
        AXMontos(0).Text = Format(lrCIMG!nCapital, "#0.00")
        AXMontos(1).Text = Format(lrCIMG!nIntComp, "#0.00")
        AXMontos(2).Text = Format(lrCIMG!nMora, "#0.00")
        AXMontos(3).Text = Format(lrCIMG!nGasto, "#0.00")

        lblTotalP.Caption = Format(lrCIMG!nCapital + lrCIMG!nIntComp + lrCIMG!nMora + lrCIMG!nGasto, "#0.00")
        AXMontos(0).SetFocus
    Else
        lblCapitalDS.Caption = Format(0, "#0.00")
        lblIntCompDS.Caption = Format(0, "#0.00")
        lblIntMoratDS.Caption = Format(0, "#0.00")
        lblGastosDS.Caption = Format(0, "#0.00")
        
        lblTotalDS.Caption = Format(0, "#0.00")
        'Me.txtMetLiq.Enabled = True
        'txtMetLiq.SetFocus
    End If
    Set lrCIMG = Nothing
        
    Dim loCalcula As COMNColocRec.NCOMColRecCalculos
    Set loCalcula = New COMNColocRec.NCOMColRecCalculos
         
    Set loCalcula = Nothing
    fnSaldoIntComp = lrValida!nSaldoIntComp '+ fnIntCompGenerado
    fnSaldoIntMorat = lrValida!nSaldoIntMor '+ fnIntMoraGenerado
    
     Me.lblCapitalD.Caption = Format(fnSaldoCap, "#,##0.00")
     Me.lblIntCompD.Caption = Format(fnSaldoIntComp, "#,##0.00")
     Me.lblIntMoratD.Caption = Format(fnSaldoIntMorat, "#,##0.00")
     Me.lblGastosD.Caption = Format(fnSaldoGasto, "#,##0.00")
     Me.lblTotalD.Caption = Format(fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat + fnSaldoGasto, "#,##0.00")
    
        lblCapitalDS.Caption = CDbl(lblCapitalD.Caption) - CDbl(AXMontos(0).Text)
        lblIntCompDS.Caption = CDbl(lblIntCompD.Caption) - CDbl(AXMontos(1).Text)
        lblIntMoratDS.Caption = CDbl(lblIntMoratD.Caption) - CDbl(AXMontos(2).Text)
        lblGastosDS.Caption = CDbl(lblGastosD.Caption) - CDbl(AXMontos(3).Text)
        
        lblTotalDS.Caption = CDbl(lblCapitalDS.Caption) + CDbl(lblIntCompDS.Caption) + CDbl(lblIntMoratDS.Caption) + CDbl(lblGastosDS.Caption)
    
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
        
    AXCodCta.Enabled = False
    CmdBuscar.Enabled = False
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & err.Number & " " & err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub
Private Sub HabilitaControles(ByVal pbCmdGrabar As Boolean, ByVal pbCmdCancelar As Boolean, _
            ByVal pbCmdSalir As Boolean, ByVal pbTxt As Boolean)
    cmdGrabar.Enabled = pbCmdGrabar
    cmdCancelar.Enabled = pbCmdCancelar
    CmdSalir.Enabled = pbCmdSalir
    Me.AXMontos(0).Enabled = pbTxt
    Me.AXMontos(1).Enabled = pbTxt
    Me.AXMontos(2).Enabled = pbTxt
    Me.AXMontos(3).Enabled = pbTxt
    'txtMetLiq.Enabled = pbTxt
End Sub
Private Sub cmdGrabar_Click()
    Dim lsmensaje As String
    Dim PorcPagMin As String
    
    If Me.chkAut.value = 1 Then
        If Me.txtGlosa.Text <> "" Then
            RegistrarPagoCancelacion
        Else
            MsgBox "No registró el detalle(Glosa) de la autorización de la Gerencia", vbInformation, "Aviso"
            Me.txtGlosa.SetFocus
        End If
    Else
        'FRHU 20150612 Observacion
        If txtGlosa.Text = "" Then
            MsgBox "Por favor, ingrese una glosa.", vbInformation, "Aviso"
            Me.txtGlosa.SetFocus
            Exit Sub
        End If
        'FIN FRHU 20150612
        RegistrarPagoCancelacion
    End If
End Sub
Public Sub RegistrarPagoCancelacion()
    Dim loContFunct As COMNContabilidad.NCOMContFunciones
    Dim loGrabar As COMNColocRec.NCOMColRecCredito
    Dim lsMovNro As String
    Dim lsFechaHoraGrab As String
    Dim msgPregunta As String, msgListo As String, msgPista As String
    If chkCancelCredito.value = 1 Then
        'msgPregunta = " Desea Grabar la cancelación del crédito?"
        msgPregunta = " ¿Desea grabar la cancelación del crédito?" 'FRHU 20150612
        msgListo = "El crédito " & AXCodCta.NroCuenta & " está listo para ser cancelado en Operaciones"
        'msgPista = "Grabar Cancelación Pago Credito Transferido FOCMACM"
        msgPista = "Grabar cancelación pago crédito transferido FOCMAC" 'FRHU 20150612
        gsOpeCod = gCredCancelacionPagoTransfFocmacm
    Else
        'msgPregunta = "Desea Grabar la Distribucion de Pago del crédito?"
        msgPregunta = "¿Desea grabar la distribución de pago del crédito?" 'FRHU 20150612
        msgListo = "El crédito " & AXCodCta.NroCuenta & " está listo para que se pueda realizar el Pago en Operaciones"
        'msgPista = "Grabar Distribución Pago Credito Transferido FOCMACM"
        msgPista = "Grabar distribución pago crédito Transferido FOCMAC" 'FRHU 20150612
        gsOpeCod = gCredDistribucionPagoTransfFocmacm
    End If
        
    'If Len(txtMetLiq.Text) = 4 Then
        If MsgBox(msgPregunta, vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
            Set loContFunct = New COMNContabilidad.NCOMContFunciones
                lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            Set loContFunct = Nothing
            
            lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
            Set loGrabar = New COMNColocRec.NCOMColRecCredito
                Call loGrabar.nRegistraPagoCancelacionCredTransferidos(AXCodCta.NroCuenta, gdFecSis, CCur(AXMontos(0).Text), CCur(AXMontos(1).Text), CCur(AXMontos(2).Text), CCur(AXMontos(3).Text), _
                                                                       "CGIM", txtGlosa.Text, fnNroCalen, lsMovNro, IIf(chkCancelCredito.value = 1, 1, 0))
                'Por el momento el txtMetLiq.Text se reemplazdo siempre por "CGIM", porque primero se debe pagar capital siempre
                Call loGrabar.nRegistraAutorizacionPagoTransferencia(AXCodCta.NroCuenta, lsMovNro, gdFecSis, lblMetLiquid.Caption, 1, "CGIM", txtGlosa.Text)
                
                MsgBox msgListo, vbInformation, "Aviso"
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, msgPista, AXCodCta.NroCuenta, gCodigoCuenta
                Set objPista = Nothing
            Set loGrabar = Nothing
            Call HabilitaControles(False, False, True, False)
            LimpiarCampos
            AXCodCta.Enabled = True
            AXCodCta.SetFocus
            CmdBuscar.Enabled = True
        Else
            MsgBox " Grabación cancelada ", vbInformation, " Aviso "
        End If
    'Else
        'MsgBox "Debe Ingresar correctamente el Metodo de Liquidacion para cancelar", vbExclamation, "Alerta"
        'Me.txtMetLiq.SetFocus
    'End If
Exit Sub
ControlError:
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub
Private Sub LimpiarCampos()
    Me.AXCodCta.NroCuenta = ""
    Me.AXCodCta.CMAC = "109"
    Me.AXCodCta.Age = gsCodAge
    Me.lblCliente.Caption = ""
    Me.lblMetLiquid.Caption = ""
    Me.lblIngRecup.Caption = ""
    'Me.txtMetLiq.Text = ""
    Me.lblCapitalD.Caption = ""
    Me.lblIntCompD.Caption = ""
    Me.lblIntMoratD.Caption = ""
    Me.lblGastosD.Caption = ""
    Me.lblTotalD.Caption = ""
    Me.AXMontos(0).Text = ""
    Me.AXMontos(1).Text = ""
    Me.AXMontos(2).Text = ""
    Me.AXMontos(3).Text = ""
    Me.lblTotalP.Caption = ""
    Me.lblCapitalDS.Caption = ""
    Me.lblIntCompDS.Caption = ""
    Me.lblIntMoratDS.Caption = ""
    Me.lblGastosDS.Caption = ""
    Me.lblTotalDS.Caption = ""
    Me.chkAut.value = 0
    Me.txtGlosa.Text = ""
    If gsCodCargo = "002017" Then 'Solo Jefe Recuperaciones
        Me.chkAut.Visible = True
    Else
        Me.chkAut.Visible = False
    End If
End Sub
Private Sub cmdCancelar_Click()
    LimpiarCampos
    Call HabilitaControles(False, False, True, False)
    AXCodCta.Enabled = True
    CmdBuscar.Enabled = True
End Sub
Private Sub AXMontos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If ValidarMontos Then
            If Index = 3 Then cmdGrabar.SetFocus Else AXMontos(Index + 1).SetFocus
        Else
            MsgBox "Los montos a pagar no deben superar a los montos de la deuda", vbInformation, "Alerta"
        End If
    End If
End Sub
Private Function ValidarMontos() As Boolean
    Dim i As Integer, nTotalP As Double, nTotalCIM As Double, nTotalDS As Double
    ValidarMontos = True
    If CDbl(AXMontos(0).Text) > CDbl(lblCapitalD.Caption) Then
        AXMontos(0).SetFocus
        ValidarMontos = False
        Exit Function
    End If
    If CDbl(AXMontos(1).Text) > CDbl(lblIntCompD.Caption) Then
        AXMontos(1).SetFocus
        ValidarMontos = False
        Exit Function
    End If
    If CDbl(AXMontos(2).Text) > CDbl(lblIntMoratD.Caption) Then
        AXMontos(2).SetFocus
        ValidarMontos = False
        Exit Function
    End If
    If CDbl(AXMontos(3).Text) > CDbl(lblGastosD.Caption) Then
        AXMontos(3).SetFocus
        ValidarMontos = False
        Exit Function
    End If
    For i = 0 To 3
        If i = 0 Then
            lblCapitalDS.Caption = Format(CDbl(lblCapitalD.Caption) - CDbl(AXMontos(i).Text), "#0.00")
            nTotalDS = nTotalDS + CDbl(lblCapitalDS.Caption)
        ElseIf i = 1 Then
            lblIntCompDS.Caption = Format(CDbl(lblIntCompD.Caption) - CDbl(AXMontos(i).Text), "#0.00")
            nTotalDS = nTotalDS + CDbl(lblIntCompDS.Caption)
        ElseIf i = 2 Then
            lblIntMoratDS.Caption = Format(CDbl(lblIntMoratD.Caption) - CDbl(AXMontos(i).Text), "#0.00")
            nTotalDS = nTotalDS + CDbl(lblIntMoratDS.Caption)
        Else
            lblGastosDS.Caption = Format(CDbl(lblGastosD.Caption) - CDbl(AXMontos(i).Text), "#0.00")
            nTotalDS = nTotalDS + CDbl(lblGastosDS.Caption)
        End If

        nTotalP = nTotalP + CDbl(AXMontos(i).Text)
        If i <> 3 Then nTotalCIM = nTotalCIM + CDbl(AXMontos(i).Text)
    Next
    'nTotalP = nTotalP + nTotalCIM
    nTotalP = nTotalP 'FRHU 20150612
    lblTotalP.Caption = Format(nTotalP, "#0.00")
    lblTotalDS.Caption = Format(nTotalDS, "#0.00")
End Function
'Private Sub txtMetLiq_KeyPress(KeyAscii As Integer)
'    Dim vCadMet As String
'    Dim x As Byte
'    KeyAscii = fgIntfMayusculas(KeyAscii)
'    vCadMet = "GCIM"
'    For x = 1 To Len(txtMetLiq)
'        vCadMet = Replace(vCadMet, Mid(txtMetLiq, x, 1), "", , , vbTextCompare)
'    Next
'    If InStr(1, vCadMet, Chr(KeyAscii), vbTextCompare) > 0 Or KeyAscii = 8 Or KeyAscii = 13 Then
'        If KeyAscii = 13 Then
'            cmdGrabar.Enabled = True
'            cmdGrabar.SetFocus
'        End If
'    Else
'        KeyAscii = 0
'    End If
'End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub

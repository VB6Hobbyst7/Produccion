VERSION 5.00
Begin VB.Form FrmExtornoAprobacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extorno de Aprobacion"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7125
   Icon            =   "FrmExtornoAprobacion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUsuario 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1230
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2460
      Width           =   4515
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos del Credito"
      Height          =   2535
      Left            =   -120
      TabIndex        =   6
      Top             =   1320
      Width           =   7215
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   2010
         Width           =   1485
      End
      Begin VB.TextBox txtHora 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1590
         Width           =   1515
      End
      Begin VB.TextBox txtEstado 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   1545
      End
      Begin VB.TextBox txtTitular 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   330
         Width           =   4515
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Monto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   540
         TabIndex        =   14
         Top             =   2040
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   645
         TabIndex        =   13
         Top             =   1620
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   390
         TabIndex        =   12
         Top             =   1170
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   495
         TabIndex        =   9
         Top             =   780
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   540
         TabIndex        =   7
         Top             =   390
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la cuenta"
      Height          =   1275
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   7095
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5580
         TabIndex        =   19
         Top             =   360
         Width           =   1425
      End
      Begin VB.ComboBox cboResolucion 
         Height          =   315
         ItemData        =   "FrmExtornoAprobacion.frx":030A
         Left            =   1110
         List            =   "FrmExtornoAprobacion.frx":0314
         TabIndex        =   4
         Top             =   720
         Width           =   1485
      End
      Begin VB.CommandButton CmdExtorno 
         Caption         =   "Extorno"
         Height          =   375
         Left            =   3870
         TabIndex        =   3
         Top             =   780
         Width           =   1425
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   3870
         TabIndex        =   2
         Top             =   330
         Width           =   1425
      End
      Begin SICMACT.ActXCodCta ActXCodCta1 
         Height          =   375
         Left            =   90
         TabIndex        =   1
         Top             =   300
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Cuenta"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Resolver"
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
         Left            =   210
         TabIndex        =   5
         Top             =   750
         Width           =   765
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   10
      Top             =   2160
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Titular:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   540
      TabIndex        =   8
      Top             =   1740
      Width           =   600
   End
End
Attribute VB_Name = "FrmExtornoAprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bIsBN As Boolean 'RIRO20170822
Private sPersCod As String 'RIRO20170822
Private nPrdEstado As Integer 'RIRO20170822

Private Sub ActXCodCta1_Change()
    If Len(ActXCodCta1.NroCuenta) > 0 Then
        CmdExtorno.Enabled = True
    End If
End Sub

Private Sub ActXCodCta1_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub


Private Sub ActXCodCta1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdBuscar_Click
    End If
End Sub

Private Sub cmdBuscar_Click()
    If Len(ActXCodCta1.NroCuenta) = 18 Then
        If ValidaCuenta = True Then
            Call CargarDatos
         Else
            If cboResolucion.ListIndex = 0 Then
                MsgBox "El credito no esta aprobado o el crédito ya fue desembolsado", vbInformation, "AVISO"
            Else
                MsgBox "El credito no esta rechazado o el crédito ya fue desembolsado", vbInformation, "AVISO"
            End If
            txtTitular.Text = ""
            txtUsuario = ""
            txtEstado = ""
            txtHora = ""
            txtMonto = "0.00"
         End If
    Else
        MsgBox "El codigo de la cuenta no esta completa", vbInformation, "AVISO"
        Exit Sub
    End If
End Sub
Private Sub cmdCancelar_Click()
    txtTitular.Text = ""
     txtUsuario = ""
     txtEstado = ""
     txtHora = ""
     txtMonto = "0.00"
     txtMonto.ForeColor = vbBlack
     ActXCodCta1.Cuenta = ""
     ActXCodCta1.Prod = ""
     'ActXCodCta1.Texto = ""
     ActXCodCta1.SetFocus
     bIsBN = False 'RIRO20170814
     sPersCod = "" 'RIRO20170814
     nPrdEstado = 0  'RIRO20170814
End Sub

Private Sub cmdExtorno_Click()
    Dim oExterno As COMDCredito.DCOMCredExtorno
    If MsgBox("Esta seguro que desea extornarlo?", vbQuestion + vbYesNo, "AVISO") = vbYes Then
        If txtUsuario <> "" And txtEstado <> "" And txtTitular <> "" And txtHora <> "" And txtMonto <> "" Then
            Set oExterno = New COMDCredito.DCOMCredExtorno
            
            'RIRO20170815 *****************
            Dim lbResultadoVisto As Boolean
            If CLng(Format(GetHoraServer, "hhmmss")) > 115900 And bIsBN And nPrdEstado = gColocEstAprob Then
                Dim sMensaje As String
                Dim loVistoElectronico As frmVistoElectronico
                Set loVistoElectronico = New frmVistoElectronico
                                                
                If MsgBox("El crédito que intenta extornar ya fue enviado al Banco de la Nación" & vbNewLine & _
                          "se recomienda extornar la aprobación el día de mañana antes del medio día." & vbNewLine & _
                          "¿Desea continuar de todas maneras?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                    
                    lbResultadoVisto = loVistoElectronico.Inicio(21, gCredExtAprobacion, sPersCod)
                    If Not lbResultadoVisto Then
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
            
            End If
            'END RIRO *********************
            'INICIO EAAS20180925 SEGUN TIC1808130006
            VerificarFechaSistemaAntesDelExtorno Me, True
            'FIN EAAS20180925 SEGUN TIC1808130006
            If oExterno.Extorno(ActXCodCta1.NroCuenta, gdFecSis, gsCodAge, gsCodUser) = True Then 'LUCV20180601, Según ERS022-2018
                If lbResultadoVisto Then
                    loVistoElectronico.RegistraVistoElectronico (0) 'RIRO20170815
                End If
                MsgBox "Se extorno correctamente", vbInformation, "AVISO"
                cmdCancelar_Click
            Else
                MsgBox "Error en el  extorno ", vbInformation, "AVISO"
            End If
        Else
            MsgBox "Existe datos incompletos", vbInformation, "AVISO"
        End If
    End If
End Sub

Private Sub Form_Activate()
        Me.ActXCodCta1.SetFocusProd
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim nProducto As Producto
    If KeyCode = vbKeyF12 And ActXCodCta1.Enabled = True Then 'F12
        Dim bRetSinTarjeta As Boolean
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(nProducto, bRetSinTarjeta)
        If sCuenta <> "" Then
            ActXCodCta1.NroCuenta = sCuenta
            ActXCodCta1.SetFocusCuenta
        End If
    End If
    
End Sub

Private Sub Form_Load()
    Me.ActXCodCta1.CMAC = gsCodCMAC '"108"
    Me.ActXCodCta1.Age = gsCodAge
    cboResolucion.ListIndex = 0
    bIsBN = False 'RIRO20170814
    sPersCod = "" 'RIRO20170814
    nPrdEstado = 0  'RIRO20170814
End Sub

Sub CargarDatos()
    Dim oExtorno As COMDCredito.DCOMCredExtorno
    Dim rs As ADODB.Recordset
    Dim nResolver As Integer
    On Error GoTo ErrHandler
        If cboResolucion.ListIndex = 0 Then
            nResolver = 1
        Else
            nResolver = 2
        End If
        
        Set oExtorno = New COMDCredito.DCOMCredExtorno
        Set rs = oExtorno.ObtenerDatosExtorno(ActXCodCta1.NroCuenta, nResolver)
        Set oExtorno = Nothing
        If Not rs.EOF And Not rs.BOF Then
            'By Capi 31102008 validacion para el extorno
             If InStr(IIf(IsNull(rs!Estado), "", rs!Estado), "REFINANCIADO") <> 0 And DateDiff("d", rs!dFecha, gdFecSis) <> 0 Then
                MsgBox "EXTORNO SOLO PROCEDE EL DIA DE LA TRANSACCION...PROCESO CANCELADO", vbInformation, "AVISO"
                CmdExtorno.Enabled = False
                CmdBuscar.SetFocus
             Else
             'End by
                txtTitular = IIf(IsNull(rs!cTitular), "", rs!cTitular)
                txtEstado = IIf(IsNull(rs!Estado), "", rs!Estado)
                txtUsuario = IIf(IsNull(rs!Usuario), "", rs!Usuario)
                txtHora = IIf(IsNull(rs!cHora), "", rs!cHora)
                bIsBN = IIf(IsNull(rs!BN), "", rs!BN) ' RIRO20170815
                sPersCod = IIf(IsNull(rs!cPersCod), "", rs!cPersCod) ' RIRO20170815
                nPrdEstado = IIf(IsNull(rs!nPrdEstado), 0, rs!nPrdEstado) ' RIRO20170815
                
                If rs!nMoneda = 2 Then
                    txtMonto.ForeColor = vbGreen
                Else
                    txtMonto.ForeColor = vbBlue
                End If
                txtMonto = Format(IIf(IsNull(rs!nMonto), 0, rs!nMonto), "#0.00")
            End If
        Else
            MsgBox "No se ha podido cargar datos" & vbCrLf & _
                   "Consulte por favor con el area de sistemas", vbInformation, "AVISO"
        End If
        
    Exit Sub
ErrHandler:
    If Not rs Is Nothing Then Set rs = Nothing
    If Not oExtorno Is Nothing Then Set oExtorno = Nothing
    MsgBox "Se ha producico un error", vbInformation, "AVISO"
End Sub

Function ValidaCuenta() As Boolean
    Dim oExtorno As COMDCredito.DCOMCredExtorno
    Dim nValida As Integer
    On Error GoTo ErrHandler
        Set oExtorno = New COMDCredito.DCOMCredExtorno
        nValida = oExtorno.validaExtorno(ActXCodCta1.NroCuenta, IIf(cboResolucion.ListIndex = 0, 1, 2))
        Set oExtorno = Nothing
        If nValida = 1 Then
            ValidaCuenta = True
        Else
            ValidaCuenta = False
        End If
    Exit Function
ErrHandler:
    If Not oExtorno Is Nothing Then Set oExtorno = Nothing
    ValidaCuenta = False
End Function


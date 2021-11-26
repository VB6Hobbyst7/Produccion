VERSION 5.00
Begin VB.Form frmCapConsultaSaldos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4200
   ClientLeft      =   1035
   ClientTop       =   1755
   ClientWidth     =   7185
   Icon            =   "frmCapConsultaSaldos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Top             =   3750
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   3750
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   3750
      Width           =   1095
   End
   Begin VB.Frame fraSaldos 
      Caption         =   "Saldos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1455
      Left            =   60
      TabIndex        =   12
      Top             =   2160
      Width           =   7035
      Begin VB.Label lblAvisoCTS 
         Caption         =   "* Saldo total de retiro de todas las cuentas CTS del cliente con el mismo empleador"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   27
         Top             =   840
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label lblMon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   4800
         TabIndex        =   24
         Top             =   308
         Width           =   255
      End
      Begin VB.Label lblMon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1620
         TabIndex        =   23
         Top             =   1028
         Width           =   255
      End
      Begin VB.Label lblMon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1620
         TabIndex        =   22
         Top             =   668
         Width           =   255
      End
      Begin VB.Label lblMon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1620
         TabIndex        =   21
         Top             =   308
         Width           =   255
      End
      Begin VB.Label lblEtqSaldoRetiro 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Retiro :"
         Height          =   195
         Left            =   3780
         TabIndex        =   20
         Top             =   315
         Width           =   960
      End
      Begin VB.Label lblSaldoRetiro 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   5160
         TabIndex        =   19
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label lblInteres 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1920
         TabIndex        =   18
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label lblSaldoContable 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1920
         TabIndex        =   17
         Top             =   600
         Width           =   1635
      End
      Begin VB.Label lblSaldoDisponible 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Interés a la Fecha:"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1028
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Contable :"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   668
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Disponible :"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   308
         Width           =   1275
      End
   End
   Begin VB.Frame fraCuenta 
      Caption         =   "Datos Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2085
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   7035
      Begin VB.Frame Frame2 
         Height          =   1305
         Left            =   4380
         TabIndex        =   6
         Top             =   660
         Width           =   2535
         Begin VB.Label lblVencimiento 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   870
            TabIndex        =   26
            Top             =   570
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label LblVence 
            AutoSize        =   -1  'True
            Caption         =   "Vence:"
            Height          =   195
            Left            =   90
            TabIndex        =   25
            Top             =   630
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   195
            Left            =   90
            TabIndex        =   10
            Top             =   270
            Width           =   690
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   90
            TabIndex        =   9
            Top             =   975
            Width           =   525
         End
         Begin VB.Label lblApertura 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   870
            TabIndex        =   8
            Top             =   210
            Width           =   1575
         End
         Begin VB.Label lblEstado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   870
            TabIndex        =   7
            Top             =   915
            Width           =   1575
         End
      End
      Begin SICMACT.FlexEdit grdCliente 
         Height          =   1125
         Left            =   180
         TabIndex        =   1
         Top             =   795
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   1984
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Cliente-RE-cPersCod"
         EncabezadosAnchos=   "250-3000-500-0"
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
         ColumnasAEditar =   "X-X-X-X"
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-C"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   435
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   767
         Texto           =   "Cuenta N°"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3900
         TabIndex        =   11
         Top             =   240
         Width           =   2955
      End
   End
End
Attribute VB_Name = "frmCapConsultaSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nProducto As COMDConstantes.Producto
Dim nTipoCuenta As COMDConstantes.ProductoCuentaTipo
Dim nMoneda As COMDConstantes.Moneda
Dim nOperacion As COMDConstantes.CaptacOperacion
'************************** MADM 20100927 ************************
Dim nPersoneria As PersPersoneria
Dim nEstado As COMDConstantes.CaptacEstado
Dim sMensaje As String, sBoleta As String, sBoletaITF As String
'************************** MADM *********************************
Dim sNumTarj As String
Dim sCuenta As String
Dim oAho As COMDCaptaGenerales.DCOMCaptaGenerales
Dim nValCons As Integer

Public Sub inicia(ByVal nProd As Producto)
nProducto = nProd
txtCuenta.NroCuenta = ""
Select Case nProducto
    Case gCapAhorros
        lblEtqSaldoRetiro.Visible = True
        lblSaldoRetiro.Visible = True
        lblEtqSaldoRetiro.Caption = "Bloq. Parcial:"
        lblMon(3).Visible = True
        Me.Caption = "Captaciones - Ahorros - Consulta de Saldos"
        nOperacion = gAhoConsSaldo
        lblAvisoCTS.Visible = False 'JUEZ 20130727
    Case gCapPlazoFijo
        lblEtqSaldoRetiro.Visible = False
        lblSaldoRetiro.Visible = False
        LblVence.Visible = True
        lblVencimiento.Visible = True
        lblMon(3).Visible = False
        Me.Caption = "Captaciones - Plazo Fijo - Consulta de Saldos"
        nOperacion = gPFConsSaldo
        lblAvisoCTS.Visible = False 'JUEZ 20130727
    Case gCapCTS
        lblEtqSaldoRetiro.Visible = True
        lblSaldoRetiro.Visible = True
        lblMon(3).Visible = True
        Me.Caption = "Captaciones - CTS - Consulta de Saldos"
        nOperacion = gCTSConsSaldo
        lblAvisoCTS.Visible = True 'JUEZ 20130727
End Select
txtCuenta.Prod = Trim(nProducto)
txtCuenta.EnabledProd = False
txtCuenta.CMAC = gsCodCMAC
txtCuenta.EnabledCMAC = False

    Set oAho = New COMDCaptaGenerales.DCOMCaptaGenerales

    nValCons = oAho.GetConsultaSaldoSinTarjeta(gsCodCargo)

    'ADD By GITU para el uso de las operaciones con tarjeta
    If (gnCodOpeTarj = 1 And nValCons = 0) Then
        sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(nOperacion), sNumTarj, nProducto)
        If sCuenta <> "123456789" Then
            If val(Mid(sCuenta, 6, 3)) <> nProducto And sCuenta <> "" Then
                MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
                Exit Sub
            End If
            If sCuenta <> "" Then
                txtCuenta.NroCuenta = sCuenta
                'txtCuenta.SetFocusCuenta
                ObtieneDatosCuenta sCuenta
            End If
            If sCuenta <> "" Then
                Me.Show 1
            End If
        Else
            Me.Show 1
        End If
    Else
        Me.Show 1
    End If
    'End GITU
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales  'NCapMantenimiento
Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento   'NCapMovimientos
Dim rsCta As ADODB.Recordset
Dim rsRel  As New ADODB.Recordset
Dim nEstado As COMDConstantes.CaptacEstado ' Comentado por MADM 20100927
Dim nRow As Long
Dim sMoneda As String, sPersona As String, sMon As String
Dim nSaldoDisp As Double, nSaldoCnt As Double, nTasa As Double, nIntAcum As Double
Dim nBloqueoParcial As Double
Dim dUltMov As Date, dUltimoEstado As Date, dVenc As String
Dim i As Integer


'MADM 20101112
'If nProducto = gCapAhorros Then 'modif.por BRGO 20110211 'Comentado x JUEZ 20140220
'    If Not CargoAutomatico(sCuenta, 1) Then Exit Sub
'End If
'END MADM

Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsCta = New ADODB.Recordset
Set rsCta = clsMant.GetDatosCuenta(sCuenta)
If Not (rsCta.EOF And rsCta.BOF) Then
    'JUEZ 20140220 **********************************
    If nProducto = gCapAhorros Then
        If MsgBox("Se va a realizar el cargo automático por la consulta de Saldos, Desea Continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
        If Not CargoAutomatico(sCuenta, 1) Then Exit Sub
    End If
    'END JUEZ ***************************************
    lblEstado = rsCta("cEstado")
    lblEstado.ToolTipText = rsCta("cEstado")
    nEstado = rsCta("nPrdEstado")
    dUltimoEstado = rsCta("dPrdEstado")
    lblApertura = Format$(rsCta("dApertura"), "dd mmm yyyy hh:mm")
    dUltMov = rsCta("dUltCierre")
    ' MADM 20100927
    nPersoneria = rsCta("nPersoneria")
    nEstado = rsCta("nPrdEstado")
    ' END MADM 20100927
    
    If Mid(sCuenta, 6, 3) = Producto.gCapPlazoFijo Then
        dVenc = Format(DateAdd("d", rsCta("nPlazo"), rsCta("dRenovacion")), "dd mmm yyyy")
    End If
    
    'Add By Gitu 09102009
    If Mid(sCuenta, 6, 3) = Producto.gCapAhorros Then
        If (rsCta("nTpoPrograma") = 3 Or rsCta("nTpoPrograma") = 2) Then
            LblVence.Visible = True
            lblVencimiento.Visible = True
            dVenc = Format(DateAdd("d", rsCta("nPlazo"), rsCta("dRenoPanderito")), "dd mmm yyyy")
        End If
    End If
    'End Gitu
    If nEstado = gCapEstAnulada Or nEstado = gCapEstCancelada Then
        nSaldoDisp = 0
        nSaldoCnt = 0
        nIntAcum = 0
    Else
        nSaldoDisp = rsCta("nSaldoDisp")
        nSaldoCnt = rsCta("nSaldo")
        nIntAcum = rsCta("nIntAcum")
    End If
    nTasa = rsCta("nTasaInteres")
    
    nMoneda = CLng(Mid(sCuenta, 9, 1))
    If nMoneda = gMonedaNacional Then
        sMoneda = "MONEDA NACIONAL"
        lblSaldoDisponible.BackColor = &HC0FFFF
        lblSaldoContable.BackColor = &HC0FFFF
        lblInteres.BackColor = &HC0FFFF
        lblSaldoRetiro.BackColor = &HC0FFFF
        For i = 0 To 3
            lblMon(i) = "S/."
        Next i
    Else
        sMoneda = "MONEDA EXTRANJERA"
        lblSaldoDisponible.BackColor = &HC0FFC0
        lblSaldoContable.BackColor = &HC0FFC0
        lblInteres.BackColor = &HC0FFC0
        lblSaldoRetiro.BackColor = &HC0FFC0
        For i = 0 To 3
            lblMon(i) = "US$"
        Next i
    End If
      
    lblSaldoDisponible = Format$(nSaldoDisp, "#,##0.00")
    lblSaldoContable = Format$(nSaldoCnt, "#,##0.00")
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    If nEstado = gCapEstCancelada Or nEstado = gCapEstAnulada Then
        lblInteres = Format$(nIntAcum, "#,##0.00")
    Else
        lblInteres = Format$(clsCap.GetCapInteresAlaFecha(nProducto, nSaldoDisp, nTasa, dUltMov, nIntAcum, gdFecSis), "#,##0.00")
    End If
    Set clsCap = Nothing
    Select Case nProducto
        Case gCapAhorros
            If rsCta("bOrdPag") Then
                lblMensaje = "AHORROS CON ORDEN DE PAGO" & Chr$(13) & sMoneda
            Else
                lblMensaje = "AHORROS SIN ORDEN DE PAGO" & Chr$(13) & sMoneda
            End If
            lblVencimiento.Caption = dVenc ' Add By Gitu 09102009
            lblSaldoRetiro = Format$(rsCta("nBloqueoParcial"), "#,##0.00")
        Case gCapPlazoFijo
            lblMensaje = "DEPOSITO A PLAZO FIJO" & Chr$(13) & sMoneda
            lblVencimiento.Caption = dVenc
            
        Case gCapCTS
            lblMensaje = "DEPOSITO CTS" & Chr$(13) & sMoneda
            lblSaldoRetiro = Format$(rsCta("nSaldRetiro"), "#,##0.00")
    End Select
    Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
    sPersona = ""
    Do While Not rsRel.EOF
        If sPersona <> rsRel("cPersCod") Then
            grdCliente.AdicionaFila
            nRow = grdCliente.Rows - 1
            grdCliente.TextMatrix(nRow, 1) = UCase(PstaNombre(rsRel("Nombre")))
            grdCliente.TextMatrix(nRow, 2) = Left(UCase(rsRel("Relacion")), 2)
            grdCliente.TextMatrix(nRow, 3) = Trim(rsRel("cPersCod")) 'FRHU ERS077-2015 20151204
            sPersona = rsRel("cPersCod")
        End If
        rsRel.MoveNext
    Loop
    rsRel.Close
    Set rsRel = Nothing
    txtCuenta.Enabled = False
    cmdImprimir.Enabled = True
    cmdCancelar.Enabled = True
    Dim clsMov As COMNContabilidad.NCOMContFunciones 'NContFunciones
    Dim sMovNro As String
    Set clsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov = Nothing
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    clsCap.CapConsultaSaldosMovimiento sCuenta, sMovNro, nOperacion, nSaldoDisp, nSaldoCnt, sNumTarj
    Set clsCap = Nothing
Else
    MsgBox "Cuenta NO Existe", vbInformation, "Aviso"
    txtCuenta.SetFocusCuenta
End If
End Sub

Private Sub LimpiaControles()
grdCliente.Clear
grdCliente.Rows = 2
grdCliente.FormaCabecera
txtCuenta.NroCuenta = ""
txtCuenta.Prod = Trim(nProducto)
txtCuenta.EnabledProd = False
txtCuenta.CMAC = gsCodCMAC
txtCuenta.EnabledCMAC = False
cmdCancelar.Enabled = False
lblApertura = ""
lblEstado = ""
cmdImprimir.Enabled = True
cmdCancelar.Enabled = True
txtCuenta.Enabled = True
txtCuenta.SetFocus
lblMensaje = ""
lblInteres = ""
lblSaldoContable = ""
lblSaldoDisponible = ""
lblInteres.BackColor = vbWhite
lblSaldoContable.BackColor = vbWhite
lblSaldoDisponible.BackColor = vbWhite
lblVencimiento.Caption = ""
'JUEZ 20140220 ***************************
If Trim(nProducto) = gCapAhorros Or Trim(nProducto) = gCapCTS Then
    lblSaldoRetiro = ""
    lblSaldoRetiro.BackColor = vbWhite
End If
'END JUEZ ********************************
Dim i As Integer
For i = 0 To 3
    lblMon(i) = "S/."
Next i
End Sub

Private Sub cmdCancelar_Click()
LimpiaControles
End Sub

Private Sub cmdImprimir_Click()
    Dim sCad As String, sMoneda As String, sCuenta As String, sCadAux As String
    Dim nCarLin As Integer, nLinPag As Integer, nCntPag As Integer
    Dim i As Integer, J As Integer
    Dim sAgencia  As String * 15
    Dim sSaldoDisponible As String * 13, sSaldoContable As String * 13, sInteresMes As String * 13
    Dim sSaldoRetiro As String * 13
    Dim sCliente As String * 65
    Dim sRelacionCta As String * 2
    Dim sTitRp1 As String, sTitRp2 As String, sNumPag As String
    Dim rs As New ADODB.Recordset
    Dim lsCadImp As String
    Dim clsPrev As previo.clsprevio
    Dim cCapImp As COMNCaptaGenerales.NCOMCaptaImpresion
    'WIOR 20130122 *******************
    Dim objPista As COMManejador.Pista
    Dim sComentario As String
    'WIOR FIN ************************
    'FRHU ERS077-2015 20151204
    Dim item As Integer
    For item = 1 To grdCliente.Rows - 1
        Call VerSiClienteActualizoAutorizoSusDatos(grdCliente.TextMatrix(item, 3), nOperacion)
    Next item
    'FRHU ERS077-2015 20151204
    'WIOR 20121009**********************************************************
    If nOperacion = gAhoConsSaldo Or nOperacion = gPFConsSaldo Or nOperacion = gCTSConsSaldo Then
        Dim oDPersona As COMDPersona.DCOMPersona
        Dim rsPersonaCred As ADODB.Recordset
        Dim rsPersona As ADODB.Recordset
        Dim Cont As Integer
        Set oDPersona = New COMDPersona.DCOMPersona
        
        
        Set rsPersonaCred = oDPersona.ObtenerPersCuentaRelac(Trim(txtCuenta.NroCuenta), gCapRelPersTitular)
        
        If rsPersonaCred.RecordCount > 0 Then
            If Not (rsPersonaCred.EOF And rsPersonaCred.BOF) Then
                For Cont = 0 To rsPersonaCred.RecordCount - 1
                    Set rsPersona = oDPersona.ObtenerUltimaVisita(Trim(rsPersonaCred!cPersCod))
                    If rsPersona.RecordCount > 0 Then
                        If Not (rsPersona.EOF And rsPersona.BOF) Then
                            If Trim(rsPersona!sUsual) = "3" Then
                            MsgBox PstaNombre(Trim(rsPersonaCred!cPersNombre), True) & "." & Chr(10) & "CLIENTE OBSERVADO: " & Trim(rsPersona!cVisObserva), vbInformation, "Aviso"
                                Call frmPersona.Inicio(Trim(rsPersonaCred!cPersCod), PersonaActualiza)
                            End If
                        End If
                    End If
                    Set rsPersona = Nothing
                    rsPersonaCred.MoveNext
                Next Cont
            End If
        End If
    End If
    'WIOR FIN ***************************************************************

    sCad = ""
    sCadAux = ""
    nCarLin = 85
    nLinPag = 65
    sCuenta = txtCuenta.NroCuenta
    sMoneda = IIf(nMoneda = gMonedaNacional, "SOLES", "DOLARES")
    sTitRp1 = "CONSULTA DE SALDOS"
    sNumPag = ""
    '******OJO PUSE SPACE 20
    nCntPag = 1
    sNumPag = FillNum(Trim(nCntPag), 4, " ")
    
    RSet sSaldoDisponible = lblSaldoDisponible
    RSet sSaldoContable = lblSaldoContable
    RSet sInteresMes = lblInteres
    RSet sSaldoRetiro = lblSaldoRetiro

    With rs
        .Fields.Append "cCliente", adVarChar, 150
        .Fields.Append "cRelaCta", adVarChar, 50
        .Open
        'Llenar Recordset
        For J = 1 To grdCliente.Rows - 1
            .AddNew
            .Fields("cCliente") = grdCliente.TextMatrix(J, 1)
            .Fields("cRelaCta") = grdCliente.TextMatrix(J, 2)
        Next J
    End With

    Set cCapImp = New COMNCaptaGenerales.NCOMCaptaImpresion
    lsCadImp = cCapImp.fgImprimirConsultaSaldos(rs, nCarLin, sTitRp1, sTitRp2, sMoneda, sNumPag, gsNomAge, gdFecSis, sCuenta, lblApertura, _
                             sSaldoDisponible, sSaldoContable, sSaldoRetiro, lblMon(0), _
                             sInteresMes, lblEstado, gsCodUser, nProducto, gCapCTS, gbImpTMU)
    Set cCapImp = Nothing
    'WIOR 20130122 ************************************
    sComentario = "Consulta de Saldos "
    Select Case nOperacion
        Case gAhoConsSaldo: sComentario = sComentario & "- AHORROS"
        Case gPFConsSaldo: sComentario = sComentario & "- PLAZO FIJO"
        Case gCTSConsSaldo: sComentario = sComentario & "- CTS"
    End Select
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista nOperacion, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gConsultar, sComentario, Trim(txtCuenta.NroCuenta), gCodigoCuenta
    Set objPista = Nothing
    'WIOR FIN *****************************************
    If gbImpTMU = False Then
        Set clsPrev = New previo.clsprevio
        clsPrev.Show lsCadImp, "Extracto de Cuenta", True, , gImpresora
        Set clsPrev = Nothing
    Else
        Dim lbOk As Boolean
        lbOk = True
        Do While lbOk
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, lsCadImp
                Print #nFicSal, ""
            Close #nFicSal
            If MsgBox("Desea Reimprimir Boleta de Consulta??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                lbOk = False
            End If
        Loop
    End If
    
    'MADM 20100928
     If sMensaje <> "" Then MsgBox sMensaje, vbInformation, "Aviso"
     If sBoleta <> "" Then ImprimeBoleta sBoleta
     If sBoletaITF <> "" Then ImprimeBoleta sBoletaITF, "Boleta ITF"
    'END MADM
    
End Sub

Private Sub ImprimeBoleta(ByVal sBoleta As String, Optional ByVal sMensaje As String = "Boleta Operación")
Dim nFicSal As Integer
Do
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
    Print #nFicSal, sBoleta '& oImpresora.gPrnSaltoLinea '& oImpresora.gPrnSaltoLinea
    Print #nFicSal, ""
    Close #nFicSal
Loop Until MsgBox("¿Desea Re-Imprimir " & sMensaje & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo
End Sub


Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
         
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.inicia(nProducto, False)
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
    
    '**DAOR 20081125, Para tarjetas ***********************
    If KeyCode = vbKeyF10 And txtCuenta.Enabled Then
        sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(nOperacion), sNumTarj, nProducto)
        If val(Mid(sCuenta, 6, 3)) <> nProducto And sCuenta <> "" Then
            MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
    '*******************************************************
    
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
'MADM 20101112
If KeyAscii = 13 Then
    Dim sCta As String
    sCta = txtCuenta.NroCuenta
    'JUEZ 20140220 ***************************************************************
    'If nProducto = gCapAhorros Then ' BRGO 20110211 Si es Cta.Ahorros realiza consulta
    '    If MsgBox("Se va a realizar el cargo automático por la consulta de Saldos, Desea Continuar?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    '       ObtieneDatosCuenta sCta
    '    Else
    '        Exit Sub
    '    End If
    'Else
    '    ObtieneDatosCuenta sCta
    'End If
    ObtieneDatosCuenta sCta
    frmSegSepelioAfiliacion.Inicio sCta
    'END JUEZ ********************************************************************
End If
End Sub

Private Function CargoAutomatico(psCuenta As String, pnCntPag As Integer) As Boolean
    Dim sMensajeCola As String
    Dim bExito As Boolean
    Dim nMonto As Double
    Dim bAplicar As Boolean
    bAplicar = True
    
    sBoleta = "" 'JUEZ 20150423
    bExito = False
    If bAplicar Then
        If nProducto = gCapAhorros And nPersoneria <> gPersonaJurCFLCMAC Then 'Ahorros y que no sean CMACs
            If nEstado <> gCapEstAnulada And nEstado <> gCapEstCancelada Then
                nMoneda = CLng(Mid(psCuenta, 9, 1))
                If nMoneda = gMonedaNacional Then
                        nMonto = GetMontoDescuento(2106, pnCntPag)
                Else
                        nMonto = GetMontoDescuento(2112, pnCntPag)
                End If
                If nMonto > 0 Then
                    Dim oCap As COMNCaptaGenerales.NCOMCaptaMovimiento  'NCapMovimientos
                    Dim sMovNro As String
                    Dim oMov As COMNContabilidad.NCOMContFunciones  'NContFunciones
                    Dim nFlag As Double, nitf As Currency
                    
                    nitf = fgITFCalculaImpuesto(nMonto)
                    Set oMov = New COMNContabilidad.NCOMContFunciones
                    sMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                    Set oMov = Nothing
                    
                    Set oCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
                    'oCap.IniciaImpresora gImpresora
                    nFlag = oCap.CapCargoCuentaAho(psCuenta, nMonto, gAhoConsSaldo, sMovNro, "Descuento Emisión Extracto", , , , , , , , , gsNomAge, sLpt, , , , , gsCodCMAC, , gsCodAge, , , nitf, , , , , sMensaje, sBoleta, sBoletaITF, , , gbImpTMU)
                    Set oCap = Nothing
                    
                    If nFlag = 0 Then
                        MsgBox "Cuenta No Posee saldo suficiente para el descuento.", vbInformation, "Aviso"
                        bExito = False
                    Else
                        bExito = True
                    End If
                Else
                    bExito = False
                End If
            End If
        End If
        CargoAutomatico = bExito
    Else
        CargoAutomatico = bExito
    End If
End Function

Private Function GetMontoDescuento(pnTipoDescuento As CaptacParametro, Optional pnCntPag As Integer = 0) As Double
Dim oParam As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim rsPar As New ADODB.Recordset

Set oParam = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Set rsPar = oParam.GetTarifaParametro(nOperacion, nMoneda, pnTipoDescuento)
Set oParam = Nothing


If rsPar.EOF And rsPar.BOF Then
    GetMontoDescuento = 0
Else
    Select Case pnTipoDescuento
        Case gDctoExtMNxPag, gDctoExtMExPag
            GetMontoDescuento = rsPar("nParValor") * pnCntPag
        Case gDctoExtMN, gDctoExtME
            GetMontoDescuento = rsPar("nParValor")
        Case Else
            GetMontoDescuento = rsPar("nParValor")
    End Select
End If
rsPar.Close
Set rsPar = Nothing
End Function

'Dim nPersoneria As PersPersoneria
'nPersoneria = rsCta("nPersoneria")


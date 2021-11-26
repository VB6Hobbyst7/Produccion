VERSION 5.00
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCapOrdPagSolicitud 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   Icon            =   "frmCapOrdPagSolicitud.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   105
      TabIndex        =   6
      Top             =   4560
      Width           =   990
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7890
      TabIndex        =   5
      Top             =   4560
      Width           =   1035
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   6750
      TabIndex        =   4
      Top             =   4560
      Width           =   1035
   End
   Begin VB.Frame fraSolicitud 
      Caption         =   "Solicitud"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2115
      Left            =   5340
      TabIndex        =   11
      Top             =   2355
      Width           =   4020
      Begin Spinner.uSpinner TxtNumTal 
         Height          =   345
         Left            =   1665
         TabIndex        =   27
         Top             =   630
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
      End
      Begin VB.ComboBox cboNumOP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         ItemData        =   "frmCapOrdPagSolicitud.frx":030A
         Left            =   1680
         List            =   "frmCapOrdPagSolicitud.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   210
         Width           =   1305
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "ITF "
         Height          =   195
         Left            =   3015
         TabIndex        =   29
         Top             =   750
         Width           =   285
      End
      Begin VB.Label lblITF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   375
         Left            =   2805
         TabIndex        =   28
         Top             =   1020
         Width           =   690
      End
      Begin VB.Line Line1 
         X1              =   135
         X2              =   3690
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Label lblMon 
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
         Left            =   3600
         TabIndex        =   26
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "# Talonarios :"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   668
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "# Ordenes de Pago :"
         Height          =   195
         Left            =   135
         TabIndex        =   24
         Top             =   270
         Width           =   1485
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Monto Descuento :"
         Height          =   195
         Left            =   135
         TabIndex        =   23
         Top             =   1110
         Width           =   1365
      End
      Begin VB.Label lblDescuento 
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
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1665
         TabIndex        =   22
         Top             =   1020
         Width           =   1035
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Del :"
         Height          =   195
         Left            =   300
         TabIndex        =   21
         Top             =   1695
         Width           =   330
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Al :"
         Height          =   195
         Left            =   1935
         TabIndex        =   20
         Top             =   1695
         Width           =   225
      End
      Begin VB.Label lblInicio 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   690
         TabIndex        =   19
         Top             =   1605
         Width           =   1155
      End
      Begin VB.Label lblFin 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   2235
         TabIndex        =   18
         Top             =   1605
         Width           =   1155
      End
   End
   Begin VB.Frame fraHistoria 
      Caption         =   "Historia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2115
      Left            =   60
      TabIndex        =   10
      Top             =   2355
      Width           =   5235
      Begin SICMACT.FlexEdit grdHistoria 
         Height          =   1800
         Left            =   90
         TabIndex        =   2
         Top             =   240
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   3175
         Cols0           =   5
         EncabezadosNombres=   "#-Inicio-Fin-Estado-Fecha"
         EncabezadosAnchos=   "350-1000-1000-1200-1200"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "#"
         ColWidth0       =   345
         RowHeight0      =   300
      End
   End
   Begin VB.Frame fraCuenta 
      Caption         =   "Datos Cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2295
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   9300
      Begin SICMACT.FlexEdit grdCliente 
         Height          =   1455
         Left            =   60
         TabIndex        =   1
         Top             =   735
         Width           =   5130
         _ExtentX        =   9049
         _ExtentY        =   2566
         Cols0           =   3
         EncabezadosNombres=   "#-Nombre-RE"
         EncabezadosAnchos=   "350-3500-500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0"
         EncabezadosAlineacion=   "C-C-C"
         FormatosEdit    =   "0-0-0"
         TextArray0      =   "#"
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin VB.Frame fraDetCuenta 
         Height          =   1515
         Left            =   5280
         TabIndex        =   8
         Top             =   660
         Width           =   3495
         Begin VB.Label lblFirmas 
            Alignment       =   2  'Center
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
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1140
            TabIndex        =   17
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lblTipoCuenta 
            Alignment       =   2  'Center
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
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1140
            TabIndex        =   16
            Top             =   660
            Width           =   2175
         End
         Begin VB.Label lblApertura 
            Alignment       =   2  'Center
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
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1140
            TabIndex        =   15
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "# Firmas :"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   1170
            Width           =   690
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cuenta :"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   750
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   330
            Width           =   690
         End
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   435
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   767
         Texto           =   "Cuenta N°"
      End
      Begin VB.Label lblDatosCuenta 
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
         Height          =   435
         Left            =   3945
         TabIndex        =   9
         Top             =   240
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmCapOrdPagSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim nMontoDescuento As Double
Dim nSaldoMinimo As Double
Dim nMaxNumOP As Long
Dim lbITFCtaExonerada As Boolean
'By capi 21012009
Dim objPista As COMManejador.Pista
'End by
Dim nRedondeoITF As Double ' BRGO 20110908




Private Sub CargaNumTalonario(ByVal nmoneda As Moneda)
    Dim rsOP As New ADODB.Recordset
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsOP = clsMant.GetTarifaOrdenPago(nmoneda)
    
    'madm 20090104 --------------------------------------------------
'    Do While Not rsOP.EOF
'        cboNumOP.AddItem rsOP("nNumOP") & Space(100) & rsOP("nCosto")
'        rsOP.MoveNext
'    Loop
    
    If Not rsOP.EOF Then
        cboNumOP.AddItem rsOP("nNumOP") & Space(100) & rsOP("nCosto")
    End If
    cboNumOP.ListIndex = 0
    Set clsMant = Nothing
End Sub

Private Sub AgregaHistoria(ByVal rsHist As Recordset)
Dim nFila As Long
Do While Not rsHist.EOF
    If grdHistoria.TextMatrix(1, 0) <> "" Then grdHistoria.Rows = grdHistoria.Rows + 1
    nFila = grdHistoria.Rows - 1
    grdHistoria.TextMatrix(nFila, 0) = nFila
    grdHistoria.TextMatrix(nFila, 1) = rsHist("nInicio")
    grdHistoria.TextMatrix(nFila, 2) = rsHist("nFin")
    grdHistoria.TextMatrix(nFila, 3) = Trim(rsHist("cEstado"))
    grdHistoria.TextMatrix(nFila, 4) = Format$(rsHist("dFecha"), "dd/mm/yyyy")
    rsHist.MoveNext
Loop
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim rsCta As New ADODB.Recordset, rsRel As New ADODB.Recordset
Dim oCuenta As COMNCaptaGenerales.NCOMCaptaGenerales
Dim nFila As Long
Dim sPersona As String
Dim oPar As COMNCaptaGenerales.NCOMCaptaDefinicion


Set oCuenta = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsCta = oCuenta.GetDatosCuenta(sCuenta)
If Not (rsCta.EOF And rsCta.BOF) Then
    Dim nEstado As COMDConstantes.CaptacEstado
    Dim bOrdPag As Boolean
    nEstado = rsCta("nPrdEstado")
    If nEstado <> gCapEstActiva Then
        Select Case nEstado
            Case gCapEstBloqRetiro, gCapEstBloqTotal
                MsgBox "Cuenta Bloqueada.", vbInformation, "Aviso"
            Case gCapEstAnulada, gCapEstCancelada
                MsgBox "Cuenta Cancelada o Anulada.", vbInformation, "Aviso"
        End Select
        txtCuenta.Cuenta = ""
        txtCuenta.SetFocusCuenta
        Exit Sub
    End If
    bOrdPag = rsCta("bOrdPag")
    If Not bOrdPag Then
        MsgBox "Cuenta NO fue aperturada con Ordenes de Pago", vbInformation, "Aviso"
        txtCuenta.Cuenta = ""
        txtCuenta.SetFocusCuenta
        Exit Sub
    End If
    lblDatosCuenta = "CUENTA CON ORDEN DE PAGO" & Chr$(13)
    Set oPar = New COMNCaptaGenerales.NCOMCaptaDefinicion
    If CLng(Mid(sCuenta, 9, 1)) = gMonedaNacional Then
        lblDescuento.BackColor = &H80000005
        lblDatosCuenta = lblDatosCuenta & "MONEDA NACIONAL"
        nSaldoMinimo = oPar.GetCapParametro(gSaldMinAhoMN)
        lblMon = "S/."
    Else
        lblDescuento.BackColor = &HC0FFC0
        lblDatosCuenta = lblDatosCuenta & "MONEDA EXTRANJERA"
        nSaldoMinimo = oPar.GetCapParametro(gSaldMinAhoME)
        lblMon = "US$"
    End If
    
    lbITFCtaExonerada = fgITFVerificaExoneracion(sCuenta)
    fgITFParamAsume Mid(sCuenta, 4, 2), Mid(sCuenta, 6, 3)
    
    Me.lblITF.BackColor = lblDescuento.BackColor
    
    Set oPar = Nothing
    lblApertura = Format$(rsCta("dApertura"), "dd-mmm-yyyy")
    lblTipoCuenta = Trim(rsCta("cTipoCuenta"))
    lblFirmas.Caption = rsCta("nFirmas")
    Set rsRel = oCuenta.GetPersonaCuenta(sCuenta)
    sPersona = ""
    Do While Not rsRel.EOF
        If sPersona <> rsRel("cPersCod") Then
            grdCliente.AdicionaFila
            nFila = grdCliente.Rows - 1
            grdCliente.TextMatrix(nFila, 1) = UCase(PstaNombre(rsRel("Nombre")))
            grdCliente.TextMatrix(nFila, 2) = Left(UCase(rsRel("Relacion")), 2)
            sPersona = rsRel("cPersCod")
        End If
        rsRel.MoveNext
    Loop
    rsRel.Close
    Set rsRel = Nothing
    
    Dim rsHist As ADODB.Recordset
    Set rsHist = oCuenta.GetHistOrdPagEmision(sCuenta)
    AgregaHistoria rsHist
    nMaxNumOP = oCuenta.GetMaxOrdPagEmitida(sCuenta)
    If nMaxNumOP = 0 Then
        lblInicio = Mid(sCuenta, 9, 1) & Format$(1, "0000000")
    Else
        lblInicio = Trim(nMaxNumOP + 1)
    End If
    CargaNumTalonario CLng(Mid(sCuenta, 9, 1))
    fraCuenta.Enabled = False
    fraHistoria.Enabled = True
    fraSolicitud.Enabled = True
    cmdGrabar.Enabled = True
    cmdCancelar.Enabled = True
    cboNumOP.SetFocus
Else
    MsgBox "Cuenta Cancelada, Anulada, o Sin Orden de Pago.", vbInformation, "Aviso"
    cmdCancelar_Click
End If
Set oCuenta = Nothing
End Sub

Private Sub cboNumOP_Click()
Dim nFin As Long
Dim nNumTal As Integer, nNumOP As Integer
Dim nCostoTalon As Double
nNumTal = TxtNumTal.valor
nNumOP = CInt(Trim(Left(cboNumOP, 10)))
nCostoTalon = CDbl(Trim(Right(cboNumOP, 10)))
'Add By Gitu 04-05-2010
If nMaxNumOP > 0 Then
    nMontoDescuento = nCostoTalon * nNumTal
Else
    nMontoDescuento = nCostoTalon * (nNumTal - 1)
End If
'End Gitu
lblDescuento = Format$(nMontoDescuento, "#,##0.00")
nFin = CLng(lblInicio) - 1
nFin = nFin + nNumTal * nNumOP
lblFin = Trim(nFin)
End Sub

Private Sub cmdCancelar_Click()
cmdGrabar.Enabled = False
fraCuenta.Enabled = True
fraHistoria.Enabled = False
fraSolicitud.Enabled = False
txtCuenta.Age = Right(gsCodAge, 2)
txtCuenta.Prod = gCapAhorros
txtCuenta.EnabledAge = False
txtCuenta.EnabledProd = False
txtCuenta.CMAC = gsCodCMAC
txtCuenta.EnabledCMAC = False
txtCuenta.Cuenta = ""
txtCuenta.SetFocusCuenta
cmdCancelar.Enabled = False
lblInicio = ""
lblFin = ""
lblDescuento.BackColor = &H80000005
lblDatosCuenta = ""
lblApertura = ""
lblTipoCuenta = ""
lblFirmas = ""
nMontoDescuento = 0
lblDescuento = "0.00"
cboNumOP.Clear
TxtNumTal.valor = 1
lblMon = "S/."
txtCuenta.EnabledCta = True
grdCliente.Rows = 2
grdCliente.Clear
grdCliente.FormaCabecera
grdHistoria.Rows = 2
grdHistoria.Clear
grdHistoria.FormaCabecera
nRedondeoITF = 0

End Sub

Private Sub CmdGrabar_Click()
Dim loMov As COMDMov.DCOMMov
Dim psMensajeValida As String, lnSaldo As Double

If lblInicio = "" Or lblFin = "" Then
    MsgBox "Error en la generación de los números a emitir", vbExclamation, "Error"
    Unload Me
End If
'If Len(lblInicio) <> 8 Or Len(lblFin) <> 8 Then
'    MsgBox "Error en la generación de los números a emitir", vbExclamation, "Error"
'    Unload Me
'End If
'JUEZ 20160428 *****************************
If CLng(TxtNumTal.valor) <= 0 Then
    MsgBox "El número de talonarios no puede ser cero", vbInformation, "Aviso"
    TxtNumTal.SetFocus
    Exit Sub
End If
'END JUEZ **********************************
If MsgBox("Desea grabar la solicitud de emisión de la Orden de Pago", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim sCuenta As String, sMovNro As String
    Dim oCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
    Dim oMov As COMNContabilidad.NCOMContFunciones 'NContFunciones
    
    Set oMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oMov = Nothing
    
    sCuenta = txtCuenta.NroCuenta
    'JUEZ 20160428 *************************************************************
    If nMontoDescuento > 0 Then
        Set oCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
        lnSaldo = oCapMov.CapCargoCuentaAho(sCuenta, nMontoDescuento, gAhoDctoEmiOP, sMovNro, "Nro Talonarios: " & Format(TxtNumTal.valor, "00") & " De " & Left(cboNumOP.Text, 5) & "  OP ", , "", "", , False, , , , gsNomAge, sLpt, , , , , , , , , gbITFAplica, CCur(Me.lblITF.Caption), gbITFAsumidoAho, gITFCobroCargo, , , psMensajeValida)
        If psMensajeValida <> "" Then
            MsgBox psMensajeValida, vbInformation, "Aviso"
            cmdCancelar_Click
            Exit Sub
        End If
    End If
    
    Dim oCapMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
    Dim nInicio As Long, nFin As Long, nNumTal As Long
    
    Set oCapMov = Nothing
    nInicio = CLng(lblInicio)
    nFin = CLng(lblFin)
    nNumTal = TxtNumTal.valor
    
    Set oCapMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        oCapMant.AgregaSolicitudOrdPagTal sCuenta, sMovNro, nInicio, nFin, nNumTal
    Set oCapMant = Nothing
    
    '*** BRGO 20110906 ***************************
    If gITF.gbITFAplica Then
       Set loMov = New COMDMov.DCOMMov
       Call loMov.InsertaMovRedondeoITF(sMovNro, 1, CCur(Me.lblITF) + nRedondeoITF, CCur(Me.lblITF))
       Set loMov = Nothing
    End If
    '*** BRGO

    'By Capi 21012009
    objPista.InsertarPista gsOpeCod, sMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, , sCuenta, gCodigoCuenta
    'End by

    cmdCancelar_Click
    'END JUEZ ******************************************************************
End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.inicia(gCapAhorros, False)
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Me.Caption = "Orden Pago - Solicitud Emisión"
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
fraHistoria.Enabled = False
fraSolicitud.Enabled = False
txtCuenta.Age = gsCodAge
txtCuenta.Prod = gCapAhorros
txtCuenta.EnabledAge = False
txtCuenta.EnabledProd = False
txtCuenta.CMAC = gsCodCMAC
txtCuenta.EnabledCMAC = False
txtCuenta.EnabledCta = True
lblInicio = ""
lblFin = ""
nMontoDescuento = 0
lblDescuento = "0.00"
lblDatosCuenta = ""
lblApertura = ""
lblTipoCuenta = ""
lblFirmas = ""
cboNumOP.Clear
TxtNumTal.valor = 1
lblMon = "S/."
'By Capi 20012009
Set objPista = New COMManejador.Pista
gsOpeCod = gAhoSolicitudOP
'End By


End Sub

Private Sub lblDescuento_Change()
    If gbITFAplica And Not lbITFCtaExonerada Then
        If Not gbITFAsumidoAho Then
            If Not IsNumeric(lblDescuento.Caption) Then lblDescuento.Caption = "0.00"
            Me.lblITF.Caption = Format(fgITFCalculaImpuesto(lblDescuento.Caption), "#,##0.00")
            '*** BRGO 20110908 ************************************************
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
            If nRedondeoITF > 0 Then
               Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
            End If
            '*** END BRGO
        End If
    End If
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim sCuenta As String
    sCuenta = txtCuenta.NroCuenta
    ObtieneDatosCuenta sCuenta
End If
End Sub

Private Sub txtNumTal_Change()
Dim nFin As Long
Dim nNumTal As Integer, nNumOP As Integer
Dim nCostoTalon As Double
If cboNumOP.Text <> "" Then
    nNumTal = TxtNumTal.valor
    nNumOP = CInt(Trim(Left(cboNumOP, 10)))
    nCostoTalon = CDbl(Trim(Right(cboNumOP, 10)))
    'Add By Gitu 04-05-2010
    If nMaxNumOP > 0 Then
        nMontoDescuento = nCostoTalon * nNumTal
    Else
        'nMontoDescuento = nCostoTalon * (nNumTal - 1)
        nMontoDescuento = nCostoTalon * (IIf(nNumTal = 0, 1, nNumTal) - 1)
    End If
    'End Gitu
    lblDescuento = Format$(nMontoDescuento, "#,##0.00")
    nFin = CLng(lblInicio) - 1
    nFin = nFin + nNumTal * nNumOP
    lblFin = Trim(nFin)
End If
End Sub

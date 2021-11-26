VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajeroOpeDevCredPers 
   Caption         =   "Devolución de Creditos por Convenio"
   ClientHeight    =   5175
   ClientLeft      =   2355
   ClientTop       =   2370
   ClientWidth     =   7635
   Icon            =   "frmCajeroOpeDevCredPers.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   7635
   Begin VB.Frame Frame3 
      Height          =   1260
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   7410
      Begin VB.CommandButton btn_buscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   28
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Soles"
         Height          =   240
         Index           =   3
         Left            =   180
         TabIndex        =   23
         Top             =   150
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Dolares"
         Height          =   240
         Index           =   2
         Left            =   1020
         TabIndex        =   22
         Top             =   165
         Width           =   1050
      End
      Begin MSMask.MaskEdBox txNotaIni 
         Height          =   300
         Left            =   1080
         TabIndex        =   26
         Top             =   480
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txNotaFin 
         Height          =   300
         Left            =   3360
         TabIndex        =   27
         Top             =   480
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         Height          =   195
         Left            =   75
         TabIndex        =   25
         Top             =   510
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   2595
         TabIndex        =   24
         Top             =   525
         Width           =   465
      End
   End
   Begin VB.TextBox txtGlosa 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   210
      MaxLength       =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   3600
      Width           =   4395
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   6210
      TabIndex        =   12
      Top             =   4740
      Width           =   1365
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   345
      Left            =   4860
      TabIndex        =   11
      Top             =   4740
      Width           =   1365
   End
   Begin VB.Frame Frame2 
      Height          =   1995
      Left            =   180
      TabIndex        =   8
      Top             =   1320
      Width           =   7410
      Begin SICMACT.FlexEdit fgDev 
         Height          =   1635
         Left            =   915
         TabIndex        =   10
         Top             =   195
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   2884
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Item-Registro-N° Cheque-Monto-cPersCodIns-cMovNro-cAge"
         EncabezadosAnchos=   "450-1400-1500-1500-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-R-C-C-C"
         FormatosEdit    =   "0-0-0-2-2-2-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   450
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1260
      Left            =   165
      TabIndex        =   0
      Top             =   45
      Width           =   7410
      Begin VB.OptionButton optMon 
         Caption         =   "Dolares"
         Height          =   240
         Index           =   1
         Left            =   1020
         TabIndex        =   5
         Top             =   165
         Width           =   1050
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Soles"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   150
         Value           =   -1  'True
         Width           =   780
      End
      Begin SICMACT.TxtBuscar txtCodPers 
         Height          =   315
         Left            =   900
         TabIndex        =   1
         Top             =   495
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
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
         TipoBusqueda    =   3
         sTitulo         =   ""
         TipoBusPers     =   1
      End
      Begin VB.Label lblCodInst 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   900
         TabIndex        =   13
         Top             =   825
         Width           =   1365
      End
      Begin VB.Label lblNomInst 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   7
         Top             =   825
         Width           =   4815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Institución:"
         Height          =   195
         Left            =   75
         TabIndex        =   6
         Top             =   885
         Width           =   765
      End
      Begin VB.Label lblnomcli 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2745
         TabIndex        =   3
         Top             =   495
         Width           =   4350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Persona :"
         Height          =   195
         Left            =   75
         TabIndex        =   2
         Top             =   510
         Width           =   675
      End
   End
   Begin VB.Label lblGlosa 
      Caption         =   "Glosa"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   210
      TabIndex        =   20
      Top             =   3390
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Monto :"
      Height          =   195
      Left            =   4815
      TabIndex        =   18
      Top             =   3540
      Width           =   540
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H80000001&
      Height          =   300
      Left            =   5565
      TabIndex        =   17
      Top             =   4230
      Width           =   1755
   End
   Begin VB.Label lblITF 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H80000001&
      Height          =   300
      Left            =   5565
      TabIndex        =   16
      Top             =   3870
      Width           =   1755
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Total :"
      Height          =   195
      Left            =   4815
      TabIndex        =   15
      Top             =   4305
      Width           =   450
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "ITF :"
      Height          =   195
      Left            =   4815
      TabIndex        =   14
      Top             =   3975
      Width           =   330
   End
   Begin VB.Label lblMonto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   345
      Left            =   5565
      TabIndex        =   9
      Top             =   3450
      Width           =   1755
   End
End
Attribute VB_Name = "frmCajeroOpeDevCredPers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnOpeCod As Long
Dim lsOperacion As String
Dim bOpeAfecta As Boolean
Dim lsCaption As String
Dim objPista As COMManejador.Pista 'MAVM 05102010
Dim lnCastDev As Integer 'MADM 20110930
'**MADM 20101006 *************************************
Dim loVistoElectronico As frmVistoElectronico
Dim lbVistoVal As Boolean
'*****************************************************
Dim nRedondeoITF As Double

Public Sub Inicia(ByVal pnOpeCod As Long, ByVal PsOperacion As String)
lnOpeCod = pnOpeCod
lsOperacion = PsOperacion
bOpeAfecta = VerifOpeVariasAfectaITF(Str(lnOpeCod))
lsCaption = Mid(PsOperacion, 3, Len(PsOperacion) - 2)
'MADM 20110630
If lnOpeCod <> 300508 Then
'    lblTotal.ForeColor = &HFFFFFF
'    lblITF.ForeColor = &HFFFFFF
'    lblMonto.ForeColor = &HFFFFFF
    Me.Frame1.Visible = True
    Me.Frame3.Visible = False
    Me.txNotaFin.Visible = False
    Me.txNotaIni.Visible = False
    Me.btn_buscar.Visible = False
    lnCastDev = 0
Else
    lnCastDev = 1
    lblTotal.ForeColor = &H80000012
    lblITF.ForeColor = &H80000012
    lblMonto.ForeColor = &H80000012
    Me.Frame3.Visible = True
    Me.Frame1.Visible = False
    Me.txNotaFin.Visible = True
    Me.txNotaIni.Visible = True
    Me.btn_buscar.Visible = True
    Me.txNotaFin.Text = gdFecSis
    Me.txNotaIni.Text = gdFecSis
    Me.Caption = lsCaption
End If
'END MADM
Me.Show 1
End Sub

Private Sub btn_buscar_Click()
If Me.txNotaFin.Text <> "" And Me.txNotaIni <> "" Then
    If IsDate(txNotaIni) And IsDate(txNotaFin) Then
        CargaDatosFecha CDate(Me.txNotaIni.Text), CDate(Me.txNotaFin.Text), IIf(Me.optMon(3).value = True, gMonedaNacional, IIf(Me.optMon(2).value = True, gMonedaExtranjera, gMonedaNacional))
        Me.txtGlosa.SetFocus
    Else
        MsgBox "Los Valores indicados en los textos no son Correctos"
    End If
Else
    MsgBox "Complete los Datos para Generar el Reporte"
End If

End Sub

Private Sub cmdCancelar_Click()
Frame1.Enabled = True
Frame2.Enabled = True
lblTotal = "0.00"
lblCodInst = ""
lblNomInst = ""
fgDev.Clear
fgDev.Rows = 2
fgDev.FormaCabecera
Me.txtCodPers = ""
Me.LblNomCli = ""
Me.txtGlosa = ""
Me.lblMonto = "0.00"
Me.lblITF = "0.00"
nRedondeoITF = 0
End Sub

Private Sub cmdGrabar_Click()
'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
'fin Comprobacion si es RFIII

Dim CodOpe As String
Dim lnMonto As Currency
Dim Moneda As String
Dim lsMov As String
Dim lsMovITF As String
Dim lsDocumento As String
Dim i As Long
Dim clsCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
Dim clsCont As COMNContabilidad.NCOMContFunciones
Set clsCont = New COMNContabilidad.NCOMContFunciones
Dim oFunFecha As New COMDConstSistema.DCOMGeneral
Dim oITF As New COMDConstSistema.FCOMITF

Dim oCred As COMDCredito.DCOMCredito
Set oCred = New COMDCredito.DCOMCredito

lnMonto = CCur(lblMonto.Caption)
lsMov = oFunFecha.FechaHora(gdFecSis)
Set oFunFecha = Nothing
lsDocumento = ""

Dim lnMovNro As Long
Dim lnMovNroITF As Long
Dim lbBan As Boolean
Dim lnMovNroBR As Long 'ALPA20131001
Dim loMov As COMDMov.DCOMMov
Set loMov = New COMDMov.DCOMMov
lnMovNroBR = 0 'ALPA20131002
'MADM 20101012--
Dim lbVistoVal As Boolean
Set loVistoElectronico = New frmVistoElectronico 'madm 20101007
lbVistoVal = False
'---------------
'LblItf.Caption = 0#

If lnOpeCod <> 300508 Then
    If Len(LblNomCli.Caption) = 0 Then
        MsgBox "Ingrese un Nombre", vbInformation, "Aviso"
        txtCodPers.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtGlosa.Text)) = 0 Then
        MsgBox "Ingrese la glosa o comentario correspondiente", vbInformation, "Aviso"
        txtGlosa.SetFocus
        Exit Sub
    End If
End If

If lnOpeCod = 300503 Then
    lsMov = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
End If

If MsgBox("Desea Grabar la Información", vbQuestion + vbYesNo, "Aviso") = vbYes Then

If lnOpeCod <> 300508 Then
  lbVistoVal = loVistoElectronico.inicio(3, lnOpeCod) 'madm 20101007
  If Not (lbVistoVal) Then Exit Sub
End If
    
If lnOpeCod <> 300508 Then
    lnMovNro = clsCapMov.OtrasOperaciones(lsMov, lnOpeCod, lnMonto, lsDocumento, txtGlosa.Text, IIf(Me.optMon(0).value = True, gMonedaNacional, gMonedaExtranjera), txtCodPers.Text, , , , , , , lnMovNroBR)
    'ALPA20131001*********************************
    If lnMovNroBR = 0 Then
        MsgBox "La operación no se realizó, favor intente nuevamente", vbInformation, "Aviso"
        Exit Sub
    End If
    '*********************************************
End If

If lnOpeCod <> 300508 Then
    For i = 1 To fgDev.Rows - 1
        'MADM 20110930 lnCastDev
        'oCred.ActualizaColocacConvenioRegDevOpe Trim(fgDev.TextMatrix(i, 4)), Trim(txtCodPers.Text), Trim(fgDev.TextMatrix(i, 1)), lnMovNro, lnCastDev, Trim(fgDev.TextMatrix(i, 5))
        oCred.ActualizaColocacConvenioRegDevOpe Trim(fgDev.TextMatrix(i, 4)), Trim(txtCodPers.Text), Trim(fgDev.TextMatrix(i, 1)), lnMovNro, 0, 0
    Next
Else
    For i = 1 To fgDev.Rows - 1
        'MADM 20110930 lnCastDev - lnMonto
        lsMov = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        lnMovNro = clsCapMov.OtrasOperaciones(lsMov, lnOpeCod, Trim(fgDev.TextMatrix(i, 3)), lsDocumento, txtGlosa.Text, IIf(Me.optMon(0).value = True, gMonedaNacional, gMonedaExtranjera), txtCodPers.Text, Trim(fgDev.TextMatrix(i, 5)), , , , , , lnMovNroBR)
        'ALPA20131001*********************************
        If lnMovNroBR = 0 Then
            MsgBox "La operación no se realizó, favor intente nuevamente", vbInformation, "Aviso"
            Exit Sub
        End If
        '*********************************************
        oCred.ActualizaColocacConvenioRegDevOpe Trim(fgDev.TextMatrix(i, 4)), Trim(txtCodPers.Text), Trim(fgDev.TextMatrix(i, 1)), lnMovNro, lnCastDev, Trim(fgDev.TextMatrix(i, 5))
    Next
End If
'MARG ERS052-2017------------------------------------
loVistoElectronico.RegistraVistoElectronico lnMovNro, , gsCodUser, lnMovNro
'END MARG ----------------------------------------------
    'MADM 20111002
    If lnOpeCod <> 300508 Then
        oITF.gbITFAplica = True
    Else
        oITF.gbITFAplica = False
    End If
    
    If oITF.gbITFAplica And CCur(Me.lblITF.Caption) <> 0 Then
        lsMovITF = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser, lsMov)
        lnMovNroITF = clsCapMov.OtrasOperaciones(lsMovITF, COMDConstantes.gITFCobroEfectivo, Abs(Me.lblITF.Caption), lsDocumento, Me.txtGlosa.Text, IIf(Me.optMon(0).value = True, COMDConstantes.gMonedaNacional, COMDConstantes.gMonedaExtranjera), txtCodPers.Text, , , , , , , lnMovNroBR)
        'ALPA20131001*********************************
        If lnMovNroBR = 0 Then
            MsgBox "La operación no se realizó, favor intente nuevamente", vbInformation, "Aviso"
            Exit Sub
        End If
        '*********************************************
        Call loMov.InsertaMovRedondeoITF(lsMovITF, 1, CCur(lblITF.Caption) + nRedondeoITF, CCur(lblITF.Caption)) 'BRGO 20110914
        Set loMov = Nothing
    End If
    
    'MAVM 05102010 ***
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista gsOpeCod, lsMov, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Devolución Créditos Personales", txtCodPers.Text, gCodigoPersona
    Set objPista = Nothing
    '***
    
    Dim oBol As COMNCaptaGenerales.NCOMCaptaImpresion
    Dim oBolITF As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim lsBoleta As String
    Dim lsBoletaITF As String
    Dim nFicSal As Integer
    
    Set oBol = New COMNCaptaGenerales.NCOMCaptaImpresion
        lsBoleta = oBol.ImprimeBoleta("OTRAS OPERACIONES", Left(lsCaption, 15), "", Str(lblMonto), Me.LblNomCli.Caption, "________" & IIf(optMon(0).value = True, gMonedaNacional, gMonedaExtranjera), lsDocumento, 0, "0", IIf(Len(lsDocumento) = 0, "", "Nro Documento"), 0, 0, False, False, , , , False, , "Nro Ope. : " & Str(lnMovNro), , gdFecSis, gsNomAge, gsCodUser, sLpt, , False, (lblITF.Caption * -1))
    Set oBol = Nothing
    
    Do
         If Trim(lsBoleta) <> "" Then
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                Print #nFicSal, ""
            Close #nFicSal
          End If
                   
    Loop While MsgBox("Desea Re Imprimir ?", vbQuestion + vbYesNo, "Aviso") = vbYes
  'INICIO JHCU ENCUESTA 16-10-2019
    Encuestas gsCodUser, gsCodAge, "ERS0292019", gsOpeCod
    'FIN
    cmdCancelar_Click
End If
Set oITF = Nothing
Set clsCapMov = Nothing
Set clsCont = Nothing
Set oCred = Nothing
End Sub

Private Sub Form_Load()
CentraForm Me
gsOpeCod = gOtrOpeDevolucionCredPersonal 'MAVM 05102010
Me.txtCodPers.psDescripcion = gsOpeCod 'MADM 20101012
End Sub


Private Sub optMon_Click(Index As Integer)
If lnOpeCod <> 300508 Then
    If optMon(0).value Then
        lblTotal.BackColor = vbWhite
    Else
        lblTotal.BackColor = &H80FF80
    End If
Else
    If optMon(3).value Then
        lblTotal.BackColor = vbWhite
    Else
        lblTotal.BackColor = &H80FF80
    End If
End If
End Sub

Private Sub txNotaFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        btn_buscar.SetFocus
    End If
End Sub

Private Sub txNotaIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txNotaFin.SetFocus
    End If
End Sub

Private Sub txtCodPers_EmiteDatos()
LblNomCli = Me.txtCodPers.psDescripcion
CargaDatos Trim(txtCodPers), IIf(Me.optMon(0).value = True, gMonedaNacional, gMonedaExtranjera)
cmdGrabar.SetFocus
Me.txtCodPers.psDescripcion = gsOpeCod 'MADM 20101012
End Sub
Sub CargaDatos(ByVal psPersCod As String, ByVal pnMoneda As COMDConstantes.Moneda)
Dim oCred As COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
Dim lnTotal As Currency
Set oCred = New COMDCredito.DCOMCredito
Dim oITF As New COMDConstSistema.FCOMITF

If psPersCod = "" Then
    Exit Sub
End If
lblTotal = "0.00"
lblITF = "0.00"
lblMonto = "0.00"
lblCodInst = ""
lblNomInst = ""
fgDev.Clear
fgDev.Rows = 2
fgDev.FormaCabecera

 oITF.fgITFParametros
oITF.gbITFAplica = True
'madm 20101012 - parametro gscodage
'''If Mid(psPersCod, 4, 2) <> gsCodAge Then
'''    MsgBox "Cliente no tiene registrado devoluciones por pagar en esta Agencia", vbInformation, "Aviso"
'''    cmdCancelar_Click
'''    Exit Sub
'''End If

Set rs = oCred.GetPersDevConv(psPersCod, pnMoneda, gsCodAge)
'end madm

If Not rs.EOF And Not rs.BOF Then
      
    lblCodInst = rs!cPersCodIns
    lblNomInst = Trim(rs!cPersNombre)
    Do While Not rs.EOF
        fgDev.AdicionaFila
        fgDev.TextMatrix(fgDev.row, 1) = rs!dRegistro
        fgDev.TextMatrix(fgDev.row, 2) = rs!cNroDoc
        fgDev.TextMatrix(fgDev.row, 3) = rs!nMonto
                   
        If lnOpeCod = 300503 Then
           'madm 20101012
            If fgDev.TextMatrix(fgDev.row, 3) > 0 Then
                fgDev.TextMatrix(fgDev.row, 3) = 0#
            End If
            'end madm
        End If
        
        fgDev.TextMatrix(fgDev.row, 4) = Trim(rs!cPersCodIns)
        fgDev.TextMatrix(fgDev.row, 5) = ""
        
        If lnOpeCod = 300508 Then
            
        End If
        
        lnTotal = lnTotal + rs!nMonto

        rs.MoveNext
    Loop
    lblMonto.Caption = Format(lnTotal, "#0.00")
    'If oITF.gbITFAplica And bOpeAfecta Then
        lblITF.Caption = Format(oITF.fgITFCalculaImpuesto(CDbl(lblMonto.Caption)), "#,##0.00")
    'End If
    lblTotal.Caption = Format(CCur(CCur(lblMonto.Caption) - Me.lblITF.Caption), "#,##0.00")
    If lnOpeCod = 300503 Then
        Frame1.Enabled = False
    End If
Else
    MsgBox "Cliente no tiene registrado devoluciones por pagar", vbInformation, "Aviso"
    cmdCancelar_Click
End If
rs.Close
Set rs = Nothing
Set oCred = Nothing
Set oITF = Nothing
End Sub

Sub CargaDatosFecha(ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal pnMoneda As COMDConstantes.Moneda)
Dim oCred As COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
Dim lnTotal As Currency
Set oCred = New COMDCredito.DCOMCredito
Dim oITF As New COMDConstSistema.FCOMITF

lblTotal = "0.00"
lblITF = "0.00"
lblMonto = "0.00"
lblCodInst = ""
lblNomInst = ""
fgDev.Clear
fgDev.Rows = 2
fgDev.FormaCabecera

oITF.fgITFParametros
oITF.gbITFAplica = False
'madm 20101012 - parametro gscodage
'''If Mid(psPersCod, 4, 2) <> gsCodAge Then
'''    MsgBox "Cliente no tiene registrado devoluciones por pagar en esta Agencia", vbInformation, "Aviso"
'''    cmdCancelar_Click
'''    Exit Sub
'''End If

Set rs = oCred.GetPersDevConvFecha(pdFecIni, pdFecFin, pnMoneda, gsCodAge)
'end madm

If Not rs.EOF And Not rs.BOF Then
      
    lblCodInst = rs!cPersCodIns
    lblNomInst = Trim(rs!cPersNombre)
    Do While Not rs.EOF
        fgDev.AdicionaFila
        fgDev.TextMatrix(fgDev.row, 1) = rs!dRegistro
        fgDev.TextMatrix(fgDev.row, 2) = rs!cNroDoc
        fgDev.TextMatrix(fgDev.row, 3) = rs!nMonto
                   
        If lnOpeCod = 300503 Then
           'madm 20101012
            If fgDev.TextMatrix(fgDev.row, 3) > 0 Then
                fgDev.TextMatrix(fgDev.row, 3) = 0#
            End If
            'end madm
        End If
        
        fgDev.TextMatrix(fgDev.row, 4) = lblCodInst
        fgDev.TextMatrix(fgDev.row, 5) = Trim(rs!nMovNroReg)
        lnTotal = lnTotal + rs!nMonto

        rs.MoveNext
    Loop
    lblMonto.Caption = Format(lnTotal, "#0.00")
    If oITF.gbITFAplica Then
        lblITF.Caption = Format(oITF.fgITFCalculaImpuesto(CDbl(lblMonto.Caption)), "#,##0.00")
        '*** BRGO 20110908 ************************************************
        nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
        If nRedondeoITF > 0 Then
            Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
        End If
        '*** END BRGO
    End If
    lblTotal.Caption = Format(CCur(CCur(lblMonto.Caption) - Me.lblITF.Caption), "#,##0.00")
Else
    MsgBox "No se encuentran devoluciones por pagar segun los criterios de búsqueda", vbInformation, "Aviso"
    cmdCancelar_Click
End If
rs.Close
Set rs = Nothing
Set oCred = Nothing
Set oITF = Nothing
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
End Sub

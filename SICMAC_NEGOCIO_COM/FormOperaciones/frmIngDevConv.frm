VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIngDevConv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de clientes para devolucion por Descto.Planilla"
   ClientHeight    =   5400
   ClientLeft      =   2940
   ClientTop       =   2955
   ClientWidth     =   7200
   Icon            =   "frmIngDevConv.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7200
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5850
      TabIndex        =   11
      Top             =   4935
      Width           =   1200
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   4665
      TabIndex        =   10
      Top             =   4935
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos a Ingresar"
      Height          =   3615
      Left            =   225
      TabIndex        =   13
      Top             =   1290
      Width           =   6840
      Begin VB.ComboBox cbocredito 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox txtDisponible 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3960
         MaxLength       =   20
         TabIndex        =   22
         Top             =   1440
         Width           =   2010
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   375
         Left            =   6240
         TabIndex        =   20
         ToolTipText     =   "Buscar Cheque"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtglosa 
         Height          =   1050
         Left            =   165
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   2505
         Width           =   3330
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   5085
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   3090
         Width           =   1485
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "Dolares"
         Height          =   285
         Index           =   1
         Left            =   5220
         TabIndex        =   4
         Top             =   180
         Width           =   1080
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "Soles"
         Height          =   285
         Index           =   0
         Left            =   4395
         TabIndex        =   3
         Top             =   195
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtNroChq 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3990
         MaxLength       =   20
         TabIndex        =   7
         Top             =   990
         Width           =   2010
      End
      Begin MSComCtl2.DTPicker txtRegistro 
         Height          =   330
         Left            =   1380
         TabIndex        =   6
         Top             =   990
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         Format          =   77463553
         CurrentDate     =   38461
      End
      Begin SICMACT.TxtBuscar txtBuscaPers 
         Height          =   345
         Left            =   720
         TabIndex        =   2
         Top             =   210
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
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
      End
      Begin VB.Label lblcredestado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   24
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Credito:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Disponible Ch.:"
         Height          =   195
         Left            =   2760
         TabIndex        =   21
         Top             =   1440
         Width           =   1065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Glosa:"
         Height          =   195
         Left            =   135
         TabIndex        =   19
         Top             =   2280
         Width           =   450
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4020
         TabIndex        =   18
         Top             =   3135
         Width           =   900
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Height          =   465
         Left            =   3840
         TabIndex        =   17
         Top             =   3030
         Width           =   2790
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nro Cheque :"
         Height          =   195
         Left            =   2940
         TabIndex        =   16
         Top             =   1035
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Registro:"
         Height          =   195
         Left            =   165
         TabIndex        =   15
         Top             =   1035
         Width           =   1125
      End
      Begin VB.Label lblNomPers 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   720
         TabIndex        =   5
         Top             =   555
         Width           =   5955
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   270
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Institucion:"
      Height          =   1065
      Left            =   195
      TabIndex        =   12
      Top             =   210
      Width           =   6825
      Begin SICMACT.TxtBuscar txtCodIns 
         Height          =   360
         Left            =   225
         TabIndex        =   0
         Top             =   210
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   635
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
      End
      Begin VB.Label lblNomInst 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   210
         TabIndex        =   1
         Top             =   585
         Width           =   6315
      End
   End
End
Attribute VB_Name = "frmIngDevConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsCodIns As String
Dim lsNomIns As String
Dim lnOpeCod As COMDConstantes.CaptacOperacion
Dim lsDescOpe As String
Dim lnSaldoCheque As Double
Dim lnIncluyeRegDevConvenio As Integer
Dim oDocRec As UDocRec 'EJVG20140227

Public Sub Inicio(ByVal pnOpeCod As Long, Optional ByVal psCodInst As String = "", Optional ByVal psNomInst As String = "", Optional psDescOpe As String = "")

lsCodIns = psCodInst
lsNomIns = psNomInst
txtCodIns = lsCodIns
lblNomInst = lsNomIns
lnOpeCod = pnOpeCod
lsDescOpe = psDescOpe
If psCodInst = "" Then
    Me.txtCodIns.Enabled = True
Else
    Me.txtCodIns.Enabled = False
End If
Me.Show 1
End Sub

Private Sub CmdGrabar_Click()
Dim CodOpe As String
Dim lnMonto As Currency
Dim Moneda As String
Dim lsMov As String
Dim lsMovITF As String
Dim lsDocumento As String
Dim i As Long
Dim clsCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim lnMovNro As Long
Dim lnMovNroITF As Long
Dim lnMovNroBR As Long 'ALPA20131004
Dim lbBan As Boolean
Dim clsCont As COMNContabilidad.NCOMContFunciones
Dim oFunFecha As New COMDConstSistema.DCOMGeneral
'Dim lnOpeCod As Integer
Dim oCred As COMDCredito.DCOMCredito
Dim lsCaption As String
If Not ValidaSeleccionCheque Then Exit Sub 'EJVG20140228

'*** PEAC 20090829 - se agregó este if si es en efectivo o cheque
If val(Me.txtDisponible.Text) > 0 Then
    '*** PEAC 20090323
    If CCur(txtMonto.Text) > CCur(Me.txtDisponible.Text) Then
        MsgBox "El monto ingresado es superior al monto disponible de este cheque.", vbInformation, "Aviso"
        Exit Sub
    End If
Else
    MsgBox "El Monto disponible de este cheque no permite la operación.", vbInformation, "Aviso"
    Exit Sub
End If
'****************


If txtCodIns.Text = "" Then
    MsgBox "Institución no válida", vbInformation, "Aviso"
    Exit Sub
End If
If Len(lblNomPers.Caption) = 0 Then
    MsgBox "Ingrese la persona para el registro", vbInformation, "Aviso"
    txtBuscaPers.SetFocus
    Exit Sub
ElseIf Len(Trim(txtglosa.Text)) = 0 Then
    MsgBox "Ingrese la glosa o comentario correspondiente", vbInformation, "Aviso"
    txtglosa.SetFocus
    Exit Sub
End If
If val(Me.txtMonto) = 0 Then
    MsgBox "Monto No válido", vbInformation, "aviso"
    Me.txtMonto.SetFocus
    Exit Sub
End If
 
Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
Set clsCont = New COMNContabilidad.NCOMContFunciones

Set oCred = New COMDCredito.DCOMCredito
 
lnMonto = CCur(txtMonto.Text)
lsMov = oFunFecha.FechaHora(gdFecSis)
Set oFunFecha = Nothing
lsDocumento = ""
lsDocumento = Me.txtNroChq.Text
 
'MADM 20110224
If cbocredito.ListIndex = -1 Then
   MsgBox "Debe seleccionar un Credito para realizar la devolución", vbInformation, "Aviso"
   Exit Sub
End If
     'MADM 20101108
    'If oCred.PertenecePersDevConv(txtBuscaPers.Text, lsCodIns) = False Then
    '   MsgBox "Persona no tiene Créditos Vigentes", vbInformation, "Aviso"
    '   Exit Sub
    'End If
    'END MADM
'END MADM

lsMov = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

If MsgBox("Desea Grabar la Información", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    lnMovNro = clsCapMov.OtrasOperaciones(lsMov, lnOpeCod, lnMonto, lsDocumento, txtglosa.Text, IIf(Me.optMoneda(0).value = True, gMonedaNacional, gMonedaExtranjera), Me.txtBuscaPers.Text, , , , , , , lnMovNroBR)
    'ALPA20131001*****************************
    If lnMovNroBR = 0 Then
        MsgBox "La operación no se realizó, favor intente nuevamente", vbInformation, "Aviso"
        Exit Sub
    End If
    
    'oCred.InsertaColocacConvenioRegDevolucion Me.txtCodIns.Text, Me.txtBuscaPers.Text, Me.txtRegistro.value, gsCodAge, IIf(Me.optMoneda(0).value = True, gMonedaNacional, gMonedaExtranjera), TxtMonto.Text, Trim(txtNroChq.Text), lnMovNro, Left(cbocredito.Text, 18)
    oCred.InsertaColocacConvenioRegDevolucion Me.txtCodIns.Text, Me.txtBuscaPers.Text, Me.txtRegistro.value, gsCodAge, IIf(Me.optMoneda(0).value = True, gMonedaNacional, gMonedaExtranjera), txtMonto.Text, oDocRec.fsNroDoc, lnMovNro, Left(cbocredito.Text, 18), oDocRec.fnTpoDoc, oDocRec.fsPersCod, oDocRec.fsIFTpo, oDocRec.fsIFCta

    '*****************************************
    'If gbITFAplica And CCur(Me.lblITF.Caption) > 0 Then
    '    lsMovITF = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser, lsMov)
    '    lnMovNroITF = clsCapMov.OtrasOperaciones(lsMovITF, gITFCobroEfectivo, Me.lblITF.Caption, lsDocumento, Me.txtglosa.Text, IIf(Me.optMon(0).value = True, gMonedaNacional, gMonedaExtranjera), txtCodPers.Text)
    'End If
    Dim oBol As COMNCaptaGenerales.NCOMCaptaImpresion
    Dim lsBoleta As String
    Dim nFicSal As String
    Set oBol = New COMNCaptaGenerales.NCOMCaptaImpresion
        lsBoleta = oBol.ImprimeBoleta(lsDescOpe, Left(lsCaption, 15), "", str(lnMonto), Me.lblNomPers.Caption, "________" & IIf(optMoneda(0).value = True, gMonedaNacional, gMonedaExtranjera), lsDocumento, 0, "0", IIf(Len(lsDocumento) = 0, "", "Nro Documento"), 0, 0, False, False, , , , False, , "Nro Ope. : " & str(lnMovNro), , gdFecSis, gsNomAge, gsCodUser, sLpt, , False)
    Set oBol = Nothing
    Do
        If Trim(lsBoleta) <> "" Then
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, lsBoleta
                Print #nFicSal, ""
            Close #nFicSal
        End If
        'If gbITFAplica And CCur(Me.lblITF.Caption) > 0 Then
        '    fgITFImprimeBoleta LblNomCli.Caption, CCur(Me.lblITF.Caption), Me.Caption, lnMovNroITF, , , , , , False
        'End If
        
    Loop While MsgBox("Desea Re Imprimir ?", vbQuestion + vbYesNo, "Aviso") = vbYes
    
    LimpiaControles
End If
Set oCred = Nothing
Set clsCapMov = Nothing
Set clsCont = Nothing
End Sub
Sub LimpiaControles()
If txtCodIns.Enabled Then
    Me.txtCodIns = ""
    Me.lblNomInst = ""
End If
Me.txtBuscaPers = ""
Me.lblNomPers = ""
Me.txtglosa = ""
Me.txtMonto = "0.00"
Me.txtNroChq = ""
Me.txtDisponible.Text = "0.00"
Me.txtRegistro = gdFecSis
Me.cmdGrabar.Enabled = True
Me.cbocredito.Clear
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()
    lnIncluyeRegDevConvenio = 1 '' 1= si incluye 0= no incluye
'MADM 20110608 - 20110225 - FILTRAR CHEQUES
    If lsCodIns <> "" Then
        'EJVG20140227 ***
        'MatDatos = frmBuscaCheque.BuscaCheque(1, , 0, lnIncluyeRegDevConvenio, lsCodIns)
        Dim oForm As New frmChequeBusquedaOtros
        Set oDocRec = oForm.IniciarBusquedaCredConvenio(IIf(Me.optMoneda(0).value = True, gMonedaNacional, gMonedaExtranjera), lsCodIns)
        Set oForm = Nothing
        txtNroChq.Text = oDocRec.fsNroDoc
        txtDisponible.Text = Format(oDocRec.fnMonto, gsFormatoNumeroView)
        'END EJVG *******
    Else
        'MatDatos = frmBuscaCheque.BuscaCheque(1, , 0, lnIncluyeRegDevConvenio)
    End If
'    MatDatos = frmBuscaCheque.BuscaCheque(1, , 0, lnIncluyeRegDevConvenio)
    'If MatDatos(0) <> "" Then
    '    txtNroChq.Text = MatDatos(4)
    '    txtDisponible.Text = MatDatos(0) ''- lnSaldoCheque
    'End If
End Sub

Private Sub Form_Load()
Dim oPersonas As COMDPersona.DCOMPersonas
Set oPersonas = New COMDPersona.DCOMPersonas
txtCodIns.rs = oPersonas.RecuperaPersonasTipo_Arbol(COMDConstantes.gPersTipoConvenio)
Set oPersonas = Nothing

CentraForm Me
Me.txtRegistro = gdFecSis


End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oDocRec = Nothing
End Sub

Private Sub optMoneda_Click(Index As Integer)
    Set oDocRec = New UDocRec
    txtNroChq.Text = oDocRec.fsNroDoc
    txtDisponible.Text = Format(oDocRec.fnMonto, gsFormatoNumeroView)
End Sub

Private Sub txtBuscaPers_EmiteDatos()
Dim oCred As COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
Set oCred = New COMDCredito.DCOMCredito

lblNomPers = Me.txtBuscaPers.psDescripcion

'MADM 20110224
Set rs = New ADODB.Recordset
Set rs = oCred.GetCreditoPersInstitucion(txtBuscaPers.Text, lsCodIns)

If Not (rs.EOF Or rs.BOF) Then
    Set oCred = Nothing
    If lblNomPers <> "" Then
        Me.txtRegistro.SetFocus
        Call CargaCboCred(rs, cbocredito)
        Me.lblcredestado = Trim(Right(Me.cbocredito.Text, 20))
        Me.cmdGrabar.Enabled = True
    End If
Else
    MsgBox "No podrá pagar la Devolución, debido a que la persona no tiene Créditos con la Institución ", vbInformation, "Aviso"
    Me.cmdGrabar.Enabled = False
    Exit Sub
End If
'END MADM

End Sub
'MADM 20110224
Private Sub CargaCboCred(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
pcboObjeto.Clear
Do While Not pRs.EOF
    pcboObjeto.AddItem Trim(pRs!Cuenta) & Space(100) & Trim(pRs!Estado)
    pRs.MoveNext
Loop
pRs.Close
End Sub
Private Sub cbocredito_Click()
If Me.cbocredito.ListIndex <> -1 Then
    Me.lblcredestado = Trim(Right(Me.cbocredito.Text, 20))
End If
End Sub
'END MADM

Private Sub txtBuscaPers_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'Me.txtRegistro.SetFocus
    If txtRegistro.Visible And txtRegistro.Enabled Then txtRegistro.SetFocus 'EJVG20140228
End If
End Sub

Private Sub txtCodIns_EmiteDatos()
Me.lblNomInst = Me.txtCodIns.psDescripcion
'Me.txtBuscaPers.SetFocus
If txtBuscaPers.Visible And txtBuscaPers.Enabled Then txtBuscaPers.SetFocus 'EJVG20140228
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
End Sub

Private Sub txtMonto_GotFocus()
fEnfoque txtMonto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMonto, KeyAscii, 12, 3)
If KeyAscii = 13 Then
    'Me.cmdGrabar.SetFocus
    If cmdGrabar.Visible And cmdGrabar.Enabled Then Me.cmdGrabar.SetFocus 'EJVG20140228
End If
End Sub

Private Sub txtMonto_LostFocus()
Me.txtMonto = Format(txtMonto, "#0.000")
End Sub

Private Sub txtNroChq_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'Me.txtglosa.SetFocus
    If txtglosa.Visible And txtglosa.Enabled Then txtglosa.SetFocus 'EJVG20140228
End If
End Sub
Private Sub txtRegistro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'Me.txtNroChq.SetFocus
    If txtNroChq.Visible And txtNroChq.Enabled Then txtNroChq.SetFocus 'EJVG20140228
End If
End Sub

'EJVG20140228 ***
Private Function ValidaSeleccionCheque() As Boolean
    ValidaSeleccionCheque = True
    If oDocRec Is Nothing Then
        ValidaSeleccionCheque = False
        Exit Function
    End If
    If Len(Trim(oDocRec.fsNroDoc)) = 0 Then
        ValidaSeleccionCheque = False
        Exit Function
    End If
End Function
'END EJVG *******

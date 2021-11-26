VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A7C47A80-96CC-11CF-8B85-0020AFE89883}#4.0#0"; "SigBox.OCX"
Begin VB.Form frmCapTarjetaRegistro 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   Icon            =   "frmCapTarjetaRegistro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1365
      TabIndex        =   18
      Top             =   4935
      Width           =   1170
   End
   Begin VB.Frame fraTarjAnt 
      Caption         =   "Tarjetas Relacionadas"
      Height          =   1590
      Left            =   105
      TabIndex        =   16
      Top             =   3255
      Width           =   4005
      Begin SICMACT.FlexEdit grdTarjetasAnt 
         Height          =   1170
         Left            =   105
         TabIndex        =   17
         Top             =   280
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   2064
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Cuenta-Tarjeta"
         EncabezadosAnchos=   "250-1800-1600"
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
         ColumnasAEditar =   "X-X-X"
         ListaControles  =   "0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C"
         FormatosEdit    =   "0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         RowHeight0      =   240
      End
   End
   Begin VB.PictureBox LogoPict 
      Height          =   315
      Left            =   210
      Picture         =   "frmCapTarjetaRegistro.frx":030A
      ScaleHeight     =   255
      ScaleWidth      =   330
      TabIndex        =   15
      Top             =   4305
      Visible         =   0   'False
      Width           =   390
   End
   Begin SigBoxLib.SigBox boxFirma 
      Height          =   615
      Left            =   2850
      TabIndex        =   14
      Top             =   4125
      Visible         =   0   'False
      Width           =   975
      _Version        =   262144
      _ExtentX        =   1720
      _ExtentY        =   1085
      _StockProps     =   233
      Appearance      =   1
      TitleText       =   ""
      PromptText      =   ""
      ConnectToPad    =   0
      Picture         =   "frmCapTarjetaRegistro.frx":2104
      DebugFileName   =   "SigBox1.TXT"
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   105
      TabIndex        =   6
      Top             =   4935
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelaCard 
      Caption         =   "Cancelar &Registro"
      Height          =   375
      Left            =   6090
      TabIndex        =   5
      Top             =   4935
      Width           =   1590
   End
   Begin VB.CommandButton cmdRegCard 
      Caption         =   "Registrar &Tarjeta"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   4935
      Width           =   1590
   End
   Begin VB.Frame fraCuentas 
      Height          =   3060
      Left            =   105
      TabIndex        =   13
      Top             =   105
      Width           =   7785
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   350
         Left            =   3780
         TabIndex        =   1
         Top             =   240
         Width           =   500
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   435
         Left            =   105
         TabIndex        =   0
         Top             =   210
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   767
         Texto           =   "Cuenta N°"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin SICMACT.FlexEdit grdCuentas 
         Height          =   645
         Left            =   105
         TabIndex        =   2
         Top             =   735
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   1138
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Cuenta-Apertura-Estado-Tipo Cuenta-Firm-Tipo Tasa"
         EncabezadosAnchos=   "250-1900-1000-1200-1200-400-1200"
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
         BackColor       =   16777215
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         RowHeight0      =   240
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCliente 
         Height          =   1485
         Left            =   195
         TabIndex        =   3
         Top             =   1560
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   2619
         _Version        =   393216
         FocusRect       =   0
         HighLight       =   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame fraRegCard 
      Caption         =   "Registro de Tarjeta"
      Height          =   1560
      Left            =   4200
      TabIndex        =   7
      Top             =   3255
      Width           =   3705
      Begin VB.TextBox lblpassw1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1365
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   735
         Width           =   810
      End
      Begin VB.TextBox lblpassw2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2415
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   735
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Password :"
         Height          =   165
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número :"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   293
         Width           =   645
      End
      Begin VB.Label lblTrack1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   990
         TabIndex        =   10
         Top             =   225
         Width           =   2565
      End
   End
End
Attribute VB_Name = "frmCapTarjetaRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nNumCuentas As Integer
Dim lsCodTar As String
Dim lsPassw As String
Dim lsmensaje As String
Dim lbIni As Boolean
Dim gbGrabaFirma As Boolean
Dim lbCargaHijo As Boolean
Dim VPersona As String




Private Sub SetupGridCliente()
Dim i As Integer
For i = 1 To grdCliente.Rows - 1
    grdCliente.MergeCol(i) = True
Next i
grdCliente.MergeCells = flexMergeFree
grdCliente.Cols = 13
grdCliente.ColWidth(0) = 100
grdCliente.ColWidth(1) = 3500
grdCliente.ColWidth(2) = 400
grdCliente.ColWidth(3) = 3500
grdCliente.ColWidth(4) = 1500
grdCliente.ColWidth(5) = 1000
grdCliente.ColWidth(6) = 600
grdCliente.ColWidth(7) = 1500
grdCliente.ColWidth(8) = 0
grdCliente.ColWidth(9) = 0
grdCliente.ColWidth(10) = 0
grdCliente.ColWidth(11) = 0
grdCliente.ColWidth(12) = 0

grdCliente.TextMatrix(0, 1) = "Nombre"
grdCliente.TextMatrix(0, 2) = "RE"
grdCliente.TextMatrix(0, 3) = "Direccion"
grdCliente.TextMatrix(0, 4) = "Zona"
grdCliente.TextMatrix(0, 5) = "Fono"
grdCliente.TextMatrix(0, 6) = "ID"
grdCliente.TextMatrix(0, 7) = "ID N°"
End Sub

Private Sub ObtieneDatosTarjeta()
Dim nItem As Long
Dim sPersona As String
Dim rsTarj As New ADODB.Recordset
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
nItem = grdCliente.Row
sPersona = grdCliente.TextMatrix(nItem, 8)
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsTarj = clsMant.GetPersonaTarj(sPersona)
Set clsMant = Nothing
If Not (rsTarj.EOF And rsTarj.BOF) Then
    Set grdTarjetasAnt.Recordset = rsTarj
Else
    grdTarjetasAnt.Clear
    grdTarjetasAnt.Rows = 2
    grdTarjetasAnt.FormaCabecera
End If
rsTarj.Close
Set rsTarj = Nothing

End Sub

Private Sub UnSetupPad()
boxFirma.Clear
boxFirma.MemoryClear
boxFirma.MagCardEnabled = False
If boxFirma.ConnectedToPad = True Then
    boxFirma.LoadLogoPicture LogoPict
End If
boxFirma.ConnectToPad = Never
End Sub

Public Sub Inicia(Optional sCuenta As String = "")

If sCuenta <> "" Then
    ObtieneDatosCuenta sCuenta
Else
    txtCuenta.CMAC = gsCodCMAC
    txtCuenta.Age = Right(gsCodAge, 2)
    txtCuenta.EnabledCMAC = False
    txtCuenta.EnabledAge = False
End If
Me.Show 1
End Sub

Private Sub ObtieneDatosPersona(ByVal sCuenta As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsCta As ADODB.Recordset

Set rsCta = New ADODB.Recordset
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = clsMant.GetPersonaCuenta(sCuenta)
Set clsMant = Nothing
If Not (rsCta.EOF And rsCta.BOF) Then
    Set grdCliente.DataSource = rsCta
    cmdRegCard.Enabled = True
    SetupGridCliente
Else
    MsgBox "Cuenta no posee relacion con Persona", vbExclamation, "Aviso"
    txtCuenta.SetFocusCuenta
    grdCuentas.Clear
    grdCuentas.FormaCabecera
End If
rsCta.Close
Set rsCta = Nothing

End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim rsCta As ADODB.Recordset, rsRel As ADODB.Recordset
Dim nEstado As COMDConstantes.CaptacEstado
Dim sSql As String

Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)
Set clsMant = Nothing
If Not (rsCta.EOF And rsCta.BOF) Then
    nEstado = rsCta("nPrdEstado")
    If nEstado <> gCapEstAnulada And nEstado <> gCapEstCancelada Then
        Dim nItem As Long
        grdCuentas.Clear
        grdCuentas.FormaCabecera
        grdCuentas.Rows = 2
        grdCuentas.AdicionaFila
        nItem = grdCuentas.Row
        grdCuentas.TextMatrix(nItem, 1) = rsCta("cCtaCod")
        grdCuentas.TextMatrix(nItem, 2) = Format$(rsCta("dApertura"), "dd-mm-yyyy")
        grdCuentas.TextMatrix(nItem, 3) = rsCta("cEstado")
        grdCuentas.TextMatrix(nItem, 4) = rsCta("cTipoCuenta")
        grdCuentas.TextMatrix(nItem, 5) = rsCta("nFirmas")
        grdCuentas.TextMatrix(nItem, 6) = rsCta("cTipoTasa")
        txtCuenta.Enabled = False
        cmdBuscar.Enabled = False
        ObtieneDatosPersona sCuenta
    Else
        MsgBox "Cuenta Anulada o Cancelada", vbInformation, "Aviso"
        txtCuenta.SetFocusCuenta
    End If
Else
    MsgBox "Cuenta no existe", vbInformation, "Aviso"
    txtCuenta.SetFocusCuenta
End If
rsCta.Close
Set rsCta = Nothing
End Sub

Private Function GetCardNumber(ByVal psTrack As String) As String
    GetCardNumber = Mid(psTrack, 3, 16)
End Function

Private Function ConectaPad() As Boolean
On Error GoTo ErrPad
boxFirma.ConnectToPad = Always
ConectaPad = boxFirma.ConnectedToPad
If Not ConectaPad Then
    MsgBox "Error de Conexión Con Pad Electrónico." & Chr(13) & Chr(13) & "Verifique la Conexión o Consulte con su Administrador", vbInformation, "Aviso "
End If
Exit Function
ErrPad:
    MsgBox Err.Description, vbCritical, "Error"
End Function

Private Sub BoxFirma_MagCard(ByVal timedOut As Boolean)
Dim lsPassw2 As String
Dim lsFecha As String
Dim lsAnio As String
Dim lsMes As String
Dim lsDia As String
Dim i As Integer
If Len(boxFirma.MagCardTrack1) > 0 Then
    lsCodTar = GetCardNumber(boxFirma.MagCardTrack1)
    lblTrack1 = Format(lsCodTar, "@@@@ @@@@ @@@@ @@@@")
    lsFecha = Mid(Me.boxFirma.MagCardTrack1, 60, 6)
    lsAnio = Left(lsFecha, 4)
    lsMes = Right(lsFecha, 2)
    If lsMes = "02" Then
        lsDia = "29"
    Else
        lsDia = "30"
    End If
    DoEvents
    lsPassw = Trim(boxFirma.GetNumber("Ingrese Password", 4, 0, 60))
    Do While Len(lsPassw) < 4
        If lsPassw = "" Then
            MsgBox "Operación Cancelada por el Usuario", vbInformation, "Aviso"
            InHabilitar
            UnSetupPad
            Exit Sub
        Else
            MsgBox "Password no Válido. Debe poseer Cuatro digitos", vbInformation, "Aviso"
            lsPassw = ""
            lsPassw = Trim(boxFirma.GetNumber("Ingrese Password", 4, 0, 60))
        End If
    Loop
    lblpassw1 = lsPassw
    DoEvents
    lsPassw2 = boxFirma.GetNumber("Confirme su Pasword", 4, 0, 60)
    i = 1
    Do While i <= 3
        If lsPassw <> lsPassw2 Then
            If i = 3 Then
                MsgBox "Numero de Reintentos Agotados. Vuelva a Realizar la Operación", vbInformation, "Aviso"
                InHabilitar
                UnSetupPad
                Exit Sub
            Else
                i = i + 1
                MsgBox "Confirmación de Password Incorrecta. Por Favor Reintente", vbInformation, "Aviso"
                lsPassw2 = boxFirma.GetNumber("Confirme su Pasword", 4, 0, 60)
                lblpassw2 = lsPassw2
            End If
        Else
            lblpassw2 = lsPassw2
            If MsgBox("Password Confirmado Correctamente" + Chr(13) + Chr(13) + "Desea Grabar el Registro de Tarjeta??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                AgregaTarjeta (lsCodTar)
                InHabilitar
                UnSetupPad
                Exit Do
            Else
                InHabilitar
                UnSetupPad
                Exit Do
            End If
        End If
    Loop
Else
    MsgBox "Error de Lectura de Tarjeta", vbInformation, "Aviso"
    UnSetupPad
    InHabilitar
End If
End Sub

Private Sub cmdBuscar_Click()
Dim clsPers As COMDPersona.UCOMPersona
Set clsPers = New COMDPersona.UCOMPersona
Set clsPers = frmBuscaPersona.Inicio
If Not clsPers Is Nothing Then
    Dim sPers As String
    Dim rsPers As ADODB.Recordset
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
    Dim sCta As String
    Dim sRelac As String * 15
    Dim sEstado As String
    Dim clsCuenta As UCapCuenta
    sPers = clsPers.sPersCod
    VPersona = sPers
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set rsPers = clsCap.GetCuentasPersona(sPers, , True)
    Set clsCap = Nothing
    If Not (rsPers.EOF And rsPers.EOF) Then
        Do While Not rsPers.EOF
            sCta = rsPers("cCtaCod")
            sRelac = rsPers("cRelacion")
            sEstado = Trim(rsPers("cEstado"))
            frmCapMantenimientoCtas.lstCuentas.AddItem sCta & Space(2) & sRelac & Space(2) & sEstado
            rsPers.MoveNext
        Loop
        Set clsCuenta = New UCapCuenta
        Set clsCuenta = frmCapMantenimientoCtas.Inicia
        If clsCuenta Is Nothing Then
        Else
            If clsCuenta.sCtaCod <> "" Then
                txtCuenta.Age = Mid(clsCuenta.sCtaCod, 4, 2)
                txtCuenta.Prod = Mid(clsCuenta.sCtaCod, 6, 3)
                txtCuenta.Cuenta = Mid(clsCuenta.sCtaCod, 9, 10)
            End If
        End If
        Set clsCuenta = Nothing
    Else
        MsgBox "Persona no posee ninguna cuenta de captaciones.", vbInformation, "Aviso"
    End If
    rsPers.Close
    Set rsPers = Nothing
End If
txtCuenta.SetFocusCuenta
End Sub

Private Sub cmdCancelaCard_Click()
InHabilitar
UnSetupPad
End Sub

Private Sub InHabilitar()
boxFirma.MagCardEnabled = False
cmdRegCard.Enabled = False
cmdCancelar.Enabled = False
End Sub

Private Sub Habilitar()
boxFirma.MagCardEnabled = True
cmdRegCard.Enabled = True
lblpassw1.Text = ""
lblpassw2.Text = ""
lblTrack1.Caption = ""
cmdCancelar.Enabled = True
End Sub

Private Sub AgregaTarjeta(ByVal psCodTarj As String)
Dim clsMov As COMNContabilidad.NCOMContFunciones 'NContFunciones
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim nItem As Long
Dim sMovNro As String, sPersona As String, sTarjeta As String
Dim sCuenta As String, sClave As String
Dim CLSSERV As COMNCaptaServicios.NCOMCaptaServicios 'NCapServicios
Set CLSSERV = New COMNCaptaServicios.NCOMCaptaServicios
Dim lsCadImp As String
Dim loPrevio As previo.clsPrevio
Dim lsmensaje As String

sClave = Encripta(Trim(lblpassw1.Text), True)

'sClave = Encripta(sClave, False)   prueba desemcripta

Set clsMov = New COMNContabilidad.NCOMContFunciones
sMovNro = clsMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
Set clsMov = Nothing

nItem = grdCliente.Row
If VPersona = "" Then
    sPersona = Trim(grdCliente.TextMatrix(nItem, 8))
Else
    sPersona = VPersona
End If

'sTarjeta = Trim(Replace(lblTrack1, "-", ""))
sTarjeta = Trim(psCodTarj)
sCuenta = txtCuenta.NroCuenta

Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
If clsMant.AgregaTarjeta(sTarjeta, sClave, sMovNro, sPersona, sCuenta, DateAdd("y", 1, gdFecSis), lsmensaje) Then
    lsCadImp = CLSSERV.ImprimeBolTarjeta("REGISTRO TARJETA", _
                                    Trim(grdCliente.TextMatrix(1, 1)), sTarjeta, _
                                    "TARJEA MAGNETICA", gdFecSis, gsNomAge, _
                                    gsCodUser, sLpt)
    Do
       Set loPrevio = New previo.clsPrevio
            loPrevio.PrintSpool sLpt, lsCadImp, False
            loPrevio.PrintSpool sLpt, Chr(10) & Chr(10) & Chr(10) & Chr(10) & lsCadImp, False
       Set loPrevio = Nothing
        
    Loop Until MsgBox("DESEA REIMPRIMIR BOLETA?", vbYesNo, "AVISO") = vbNo

    cmdCancelar_Click
Else
    If lsmensaje <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
    End If
    InHabilitar
    cmdCancelar.Enabled = True
    lblTrack1.Caption = ""
    lblpassw1.Text = ""
    lblpassw2.Text = ""
    cmdRegCard.Enabled = True
    cmdCancelaCard.Enabled = True
    UnSetupPad
    VPersona = ""
End If
Set clsMant = Nothing
Set CLSSERV = Nothing
End Sub

Private Sub cmdCancelar_Click()
grdCliente.Clear
grdCliente.Rows = 2
grdCuentas.Clear
grdCuentas.FormaCabecera
grdCuentas.Rows = 2
grdTarjetasAnt.Clear
grdTarjetasAnt.FormaCabecera
grdTarjetasAnt.Rows = 2
lblTrack1 = ""
lblpassw1 = ""
lblpassw2 = ""
txtCuenta.NroCuenta = ""
txtCuenta.CMAC = gsCodCMAC
txtCuenta.EnabledCMAC = False
txtCuenta.Age = gsCodAge
txtCuenta.EnabledAge = False
txtCuenta.Enabled = True
cmdBuscar.Enabled = True
InHabilitar
UnSetupPad
cmdCancelaCard.Enabled = False
End Sub

Private Sub cmdRegCard_Click()

Dim nCOM As COMDConstantes.TipoPuertoSerial
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim sMaquina As String
Dim lnNumTar As String
Dim lsCaption As String
'opciones de validacion

sMaquina = GetComputerName
Set clsGen = New COMDConstSistema.DCOMGeneral
    nCOM = clsGen.GetPuertoPeriferico(COMDConstantes.gPerifPENWARE, sMaquina)
    clsGen.GetPerifericosPC (sMaquina)
Set clsGen = Nothing

If nCOM = -1 Then
    Set clsGen = New COMDConstSistema.DCOMGeneral
        nCOM = clsGen.GetPuertoPeriferico(COMDConstantes.gPerifPINPAD, sMaquina)
    Set clsGen = Nothing

    'GnTipoPinPad = ObtieneTipoPinPad()
    
'*********  CONEXION CON VERIFONE 5000  *****
Dim lsNumTarTemp As String
Dim lsNumTar As String
    
    If Not GmyPSerial Is Nothing Then
        GmyPSerial.Disconnect
        Set GmyPSerial = Nothing
    End If
    Set GmyPSerial = CreateObject("HComPinpad.Pinpad")
    If GmyPSerial.ConnectionTest = 0 Then
        Call GmyPSerial.Connect(CInt(nCOM), 9600)
        If GmyPSerial.ConnectionTest = 1 Then
             If GmyPSerial.ReadCardIniConf("PASE SU TARJETA") = 1 Then
                    lsNumTar = GetNumTarjeta_Vrf5000
                    lsNumTarTemp = lsNumTar
                    If Len(lsNumTar) <> 16 Then
                        MsgBox "Error en la Lectura de Tarjeta.", vbInformation, "Aviso"
                        GmyPSerial.Disconnect
                        Set GmyPSerial = Nothing
                        Exit Sub
                    End If
                    MsgBox "PROCEDA A INGRESAR SU CLAVE...", vbInformation, "AVISO"
                    Me.Caption = "Ingrese la Clave de la Tarjeta."
                    Call GrabaTarjetaPINPADV_5000(nCOM, lsNumTar)
             End If
        End If
    End If
    If Not GmyPSerial Is Nothing Then
        GmyPSerial.Disconnect
        Set GmyPSerial = Nothing
    End If
Else
    If ConectaPad Then
        Habilitar
        MuestraPantalla
    End If
End If

End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub MuestraPantalla()
boxFirma.Clear
boxFirma.TitleText = "CMAC CUSCO"
boxFirma.PromptText = "Pase su Tarjeta"
boxFirma.PutBitmap 2, 2, App.path & "\Logo5.bmp"
End Sub


Private Function GrabaTarjetaPINPAD(ByVal pnCom As COMDConstantes.TipoPuertoSerial)
Dim sNumTar As String
Dim sClaveTar As String
Dim lnErr As Long
Dim lnNumOp As Integer
Dim sTitulo As String

sTitulo = Me.Caption

'ppoa Modificacion
'Call IniciaPinPad(Trim(IIf(pnCom = gPuertoSerialCOM1, gPuertoSerialCOM1, gPuertoSerialCOM2)))
If Not WriteToLcd("Pase su Tarjeta") Then
    FinalizaPinPad
    MsgBox "No se Realizó Envío", vbInformation, "Aviso"
    Exit Function
End If

'ppoa Modificacion
'*************
    If GmyPSerial.ReadCardIni = 1 Then
    End If
'*********

sNumTar = GetNumTarjeta


lsCodTar = sNumTar
lblTrack1 = sNumTar
'sNumTar = Replace(sNumTar, "-", "", 1, , vbTextCompare)

If Len(sNumTar) <> 16 Then
    MsgBox "Error en la Lectura de Tarjeta.", vbInformation, "Aviso"
    FinalizaPinPad
    Exit Function
End If
MsgBox "PROCEDA A INGRESAR SU CLAVE...", vbInformation, "AVISO"
Me.Caption = "Ingrese la Clave de la Tarjeta."

'If Not WriteToLcd("                                       ") Then
'    FinalizaPinPad
'    MsgBox "No se Realizó Envío", vbInformation, "Aviso"
'    Exit Function
'End If
'ppoa Modificacion
sClaveTar = GetClaveTarjeta
                        
                        
If sClaveTar = "" Then
    MsgBox "Debe Ingresar una Clave Valida.", vbInformation, "Aviso"
    lblTrack1 = ""
    Exit Function
End If

lblpassw1 = sClaveTar
sClaveTar = ""
lnNumOp = 0

While lnNumOp < 3 And lblpassw1 <> sClaveTar
    sClaveTar = GetClaveTarjeta
    lnNumOp = lnNumOp + 1
    If lblpassw1 <> sClaveTar And lnNumOp < 3 Then
        MsgBox "La clave es errada. Re-Ingrese su Clave.", vbInformation, "Aviso"
    End If
Wend

lblpassw2 = sClaveTar
If lblpassw1 = lblpassw2 Then
    If MsgBox("Desea Registrar la Tarjeta ? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        AgregaTarjeta lsCodTar
    End If
Else
    MsgBox "La clave no ha sido reconocida, el proceso sera Cancelado.", vbInformation, "Aviso"
End If

FinalizaPinPad
cmdCancelaCard.Enabled = False
Me.Caption = sTitulo

End Function

Private Function GrabaTarjetaPINPADV_5000(ByVal pnCom As COMDConstantes.TipoPuertoSerial, ByVal psCodTarj As String)
Dim sNumTar As String
Dim sClaveTar As String
Dim lnErr As Long
Dim lnNumOp As Integer
Dim sTitulo As String

'MsgBox "PROCEDA A INGRESAR SU CLAVE...", vbInformation, "AVISO"
Me.Caption = "Ingrese la Clave de la Tarjeta."
sClaveTar = GetClaveTarjeta_Vrf5000("INGRESE CLAVE ")
If sClaveTar = "" Then
    MsgBox "Debe Ingresar una Clave Valida.", vbInformation, "Aviso"
    lblTrack1 = ""
    Exit Function
End If

lblpassw1 = sClaveTar
sClaveTar = ""
lnNumOp = 0

While lnNumOp < 3 And lblpassw1 <> sClaveTar
    sClaveTar = GetClaveTarjeta("Conf.Clave [" & Trim(Str(lnNumOp + 1)) & "]")
    lnNumOp = lnNumOp + 1
    If lblpassw1 <> sClaveTar And lnNumOp < 3 Then
        MsgBox "La clave es errada. Re-Ingrese su Clave.", vbInformation, "Aviso"
    End If
Wend

lblpassw2 = sClaveTar
If lblpassw1 = lblpassw2 Then
    If MsgBox("Desea Registrar la Tarjeta ? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        AgregaTarjeta (psCodTarj)
    End If
Else
    MsgBox "La clave no ha sido reconocida, el proceso sera Cancelado.", vbInformation, "Aviso"
End If
cmdCancelaCard.Enabled = False
Me.Caption = sTitulo
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gCapAhorros, False)
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
VPersona = ""
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Me.Caption = "Captaciones - Registro de Tarjeta"
txtCuenta.CMAC = gsCodCMAC
txtCuenta.EnabledCMAC = False
txtCuenta.Age = gsCodAge
txtCuenta.EnabledAge = False
cmdRegCard.Enabled = False
cmdCancelaCard.Enabled = True
SetupGridCliente
End Sub

Private Sub grdCliente_Click()
ObtieneDatosTarjeta
End Sub

Private Sub grdCliente_RowColChange()
ObtieneDatosTarjeta
End Sub


Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim sCta As String
    sCta = txtCuenta.NroCuenta
    ObtieneDatosCuenta sCta
End If
End Sub

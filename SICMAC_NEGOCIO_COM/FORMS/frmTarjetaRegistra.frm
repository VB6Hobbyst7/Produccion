VERSION 5.00
Begin VB.Form frmTarjetaRegistra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar Tarjeta"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7860
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraRegCard 
      Caption         =   "Registro de Tarjeta"
      Height          =   720
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   7545
      Begin VB.Label lblTrack1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   2430
         TabIndex        =   12
         Top             =   255
         Width           =   3165
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número :"
         Height          =   195
         Left            =   1080
         TabIndex        =   11
         Top             =   300
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdRegCard 
      Caption         =   "Registrar &Tarjeta"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelaCard 
      Caption         =   "Cancelar &Registro"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7455
      Begin VB.Label LblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   7215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   7215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Direccion"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   1695
   End
   Begin SICMACT.FlexEdit grdCuentas 
      Height          =   1845
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   2831
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
End
Attribute VB_Name = "frmTarjetaRegistra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cPersCod As String
Private Sub cmdBuscar_Click()
Dim ObjTarj As COMNCaptaServicios.NCOMCaptaTarjeta
Set ObjTarj = New COMNCaptaServicios.NCOMCaptaTarjeta
Dim clsPers As COMDPersona.UCOMPersona
Set clsPers = New COMDPersona.UCOMPersona
Set clsPers = frmBuscaPersona.Inicio
If Not clsPers Is Nothing Then
    Me.LblNombre.Caption = Trim(clsPers.sPersNombre)
    Me.lblDireccion.Caption = Trim(clsPers.sPersDireccDomicilio)
    cPersCod = clsPers.sPersCod
    ObtieneDatosCuenta (cPersCod)
    cmdRegCard.Enabled = True
    If ObjTarj.VerificaPersonaTarjetaActiva(cPersCod) Then
        MsgBox "La persona ya posse una tarjeta activa", vbInformation, "AVISO"
        cmdCancelaCard_Click
    End If
Else
    cPersCod = ""
End If
Set ObjTarj = Nothing
End Sub

Private Sub cmdCancelaCard_Click()
cPersCod = ""
Me.lblDireccion = ""
Me.LblNombre = ""
Me.lblTrack1 = ""
grdCuentas.Clear
grdCuentas.FormaCabecera
grdCuentas.Rows = 2
cmdRegCard.Enabled = False
End Sub


Private Function Get_Banda_Tarj() As Boolean
Dim cTarjeta As String
Dim ObjTarj As COMNCaptaServicios.NCOMCaptaTarjeta
Set ObjTarj = New COMNCaptaServicios.NCOMCaptaTarjeta
'VerificaTarjetaActivaReg
    Dim I_RESULT As Integer
    Dim S_RESULT As String
    Dim L_RESULT As Long
    CargaValoresPinPad_ACS
    Get_Banda_Tarj = False
    I_RESULT = OpenReader()
    S_RESULT = CStr(I_RESULT)
    'Call ClearDisplay
    'Call DisplayString(I_LINE, I_OFFSET, I_TYPE, "Pase la Tarjeta")
    'LightOn (50)
    'I_RESULT = ReadMagneticCard(I_MIN, I_TRACK)
    I_RESULT = ReadMagneticCard(60, 1)
    S_RESULT = CStr(I_RESULT)
    'LightOn (0)
    If I_RESULT > 0 Then
        L_RESULT = ReadTrack(1, 0, 110)
        S_RESULT = sANSIPtrZToVBString(L_RESULT)
        lblTrack1 = Mid(S_RESULT, 3, 16)
        cTarjeta = Mid(S_RESULT, 3, 16)
        If Not IsNumeric(Mid(S_RESULT, 3, 16)) Then
            MsgBox "Tarjeta Invalida", vbInformation, "AVISO"
            CloseReader
            Get_Banda_Tarj = False
            Exit Function
        End If
        If Not ObjTarj.VerificaTarjetaActivaReg(cTarjeta) Then
            MsgBox "Tarjeta Invalida, para Registro", vbInformation, "AVISO"
            CloseReader
            Set ObjTarj = Nothing
            Get_Banda_Tarj = False
            Exit Function
        End If
        S_PAN = cTarjeta
        S_PVKI = gsPVKi
        S_IPSERVER = gsServPindVerify
        S_PVV = ObjTarj.Get_Tarj_PVV(S_PAN)
        MsgBox "Ingrese la Clave.."
        I_RESULT = CheckPIN(I_TIMEOUT, S_IPSERVER, S_PAN, S_PVKI, S_PVV)
        
        If I_RESULT = 1 Then
            AgregaTarjeta
            Get_Banda_Tarj = True
        ElseIf I_RESULT = -1 Then
            MsgBox "Pin Incorrecto", vbInformation, "AVISO"
            Get_Banda_Tarj = False
        Else
            MsgBox "Error de Conexion", vbInformation, "AVISO"
            Get_Banda_Tarj = False
        End If
    End If
    CloseReader
    Set ObjTarj = Nothing
    
End Function


Private Sub cmdRegCard_Click()
If Trim(LblNombre.Caption) = "" Then
    MsgBox "Debe de buscar primero al cliente para hacer la afiliacion", vbInformation, "AVISO"
    Me.CmdBuscar.SetFocus
    Exit Sub
End If
'If Trim(lblTrack1.Caption) = "" Then
'    MsgBox "Pase la Tarjeta por el Pind Pad", vbInformation, "AVISO"
'    Exit Sub
'End If


If Get_Banda_Tarj() Then
Else
    MsgBox "Error al Registrar la Tarjeta, vuelva a intentarlo", vbInformation
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub CambioClave(ByVal psTarjeta As String)
Dim cTarjeta As String
Dim sMovNro As String
Dim ObjTarj As COMNCaptaServicios.NCOMCaptaTarjeta
Set ObjTarj = New COMNCaptaServicios.NCOMCaptaTarjeta
Dim clsMov As COMNContabilidad.NCOMContFunciones
Set clsMov = New COMNContabilidad.NCOMContFunciones


'VerificaTarjetaActivaReg
    Dim I_RESULT As Integer
    Dim S_RESULT As String
    Dim L_RESULT As Long
    MsgBox "Se procedera al cambio de la Clave", vbInformation, "AVISO"
    CargaValoresPinPad_ACS
    I_RESULT = OpenReader()
    S_RESULT = CStr(I_RESULT)

    'I_RESULT = ReadMagneticCard(I_MIN, I_TRACK)
    'I_RESULT = ReadMagneticCard(60, 1)
    S_RESULT = CStr(I_RESULT)
    If I_RESULT > 0 Then
        'L_RESULT = ReadTrack(1, 0, 110)
        'S_RESULT = sANSIPtrZToVBString(L_RESULT)
        cTarjeta = psTarjeta
        'If Not IsNumeric(Mid(S_RESULT, 3, 16)) Then
        '    MsgBox "Tarjeta Invalida", vbInformation, "AVISO"
        '    CloseReader
        '    Exit Sub
        'End If
        'If Not ObjTarj.VerificaTarjetaActiva(cTarjeta) Then
        '    MsgBox "Esta tarjeta no se encuentra activa", vbInformation, "AVISO"
        '    CloseReader
        '    Set ObjTarj = Nothing
        '    Exit Sub
        'End If
        'txtTarjeta.Text = Format(Mid(S_RESULT, 3, 16), "####-####-####-####")
        S_PAN = cTarjeta
        S_PVKI = gsPVKi
        S_IPSERVER = gsServPindVerify
        S_PVV = ObjTarj.Get_Tarj_PVV(S_PAN)
        I_RESULT = CheckPIN(I_TIMEOUT, S_IPSERVER, S_PAN, S_PVKI, S_PVV)
        'L_RESULT = ""
        If I_RESULT = 1 Then
            If IsNumeric(Mid(S_RESULT, 3, 16)) Then
                I_CHNG_PASS = 1
                LightOn (30)
                MsgBox "Ingrese su nueva Clave", vbInformation, "AVISO"
                L_RESULT = ChangePIN(I_TIMEOUT, S_IPSERVER, S_PAN, S_PVKI, I_CHNG_PASS, "NUEVA CLAVE")
                S_RESULT = CStr(I_RESULT)
                S_RESULT = sANSIPtrZToVBString(L_RESULT)
                LightOn (0)
                If IsNumeric(Trim(S_RESULT)) Then
                    Call ObjTarj.Put_Tarj_PVV(S_PAN, S_RESULT)
                    sMovNro = clsMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
                    Call ObjTarj.RegistraCambioPind(S_PAN, Format(FechaHora(gdFecSis), "mm/dd/YYYY HH:MM:SS"), sMovNro)
                Else
                    MsgBox "Error al Cambiar la Clave", vbInformation
                End If
            End If
        ElseIf I_RESULT = -1 Then
            MsgBox "Clave incorrecta", vbInformation, "AVISO"
        ElseIf I_RESULT = -2 Then
            MsgBox "No existe conexion con el servidor PINDVERIFY", vbInformation, "AVISO"
        End If
    End If
    CloseReader
    Set ObjTarj = Nothing
End Sub
Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub
Private Sub AgregaTarjeta()
Dim clsMov As COMNContabilidad.NCOMContFunciones 'NContFunciones
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim nItem As Long
Dim sMovNro As String, sPersona As String, sTarjeta As String
Dim sCuenta As String, sClave As String
Dim CLSSERV As COMNCaptaServicios.NCOMCaptaServicios 'NCapServicios
Set CLSSERV = New COMNCaptaServicios.NCOMCaptaServicios
Dim lscadimp As String
Dim loPrevio As previo.clsPrevio
Dim lsmensaje As String

Dim ObjTarj As COMNCaptaServicios.NCOMCaptaTarjeta
Set ObjTarj = New COMNCaptaServicios.NCOMCaptaTarjeta



Set clsMov = New COMNContabilidad.NCOMContFunciones
sMovNro = clsMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
Set clsMov = Nothing

sPersona = cPersCod
sTarjeta = Trim(Replace(lblTrack1, "-", ""))
Call ObjTarj.RegistraTarjetaPersona(sTarjeta, sMovNro, sPersona)
lscadimp = CLSSERV.ImprimeBolTarjeta("REGISTRO TARJETA", _
                                Trim(Me.LblNombre.Caption), sTarjeta, _
                                "TARJEA MAGNETICA", gdFecSis, gsNomAge, _
                                gsCodUser, sLpt)
Do
   Set loPrevio = New previo.clsPrevio
        loPrevio.PrintSpool sLpt, lscadimp, False
        loPrevio.PrintSpool sLpt, Chr(10) & Chr(10) & Chr(10) & Chr(10) & lscadimp, False
   Set loPrevio = Nothing
    
Loop Until MsgBox("DESEA REIMPRIMIR BOLETA?", vbYesNo, "AVISO") = vbNo

If lsmensaje <> "" Then
    MsgBox lsmensaje, vbInformation, "Aviso"
End If


lblTrack1.Caption = ""
cmdRegCard.Enabled = False
cmdCancelaCard.Enabled = True
Set CLSSERV = Nothing
End Sub


Private Sub ObtieneDatosCuenta(ByVal cPersCod As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim rsCta As ADODB.Recordset, rsRel As ADODB.Recordset
Dim nEstado As COMDConstantes.CaptacEstado
Dim sSQL As String

Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetCuentasPersona(cPersCod, gCapAhorros, True)
Set clsMant = Nothing
If Not (rsCta.EOF And rsCta.BOF) Then
    'nEstado = rsCta("nPrdEstado")
    'If nEstado <> gCapEstAnulada And nEstado <> gCapEstCancelada Then
        
        Dim nItem As Long
        grdCuentas.Clear
        grdCuentas.FormaCabecera
        grdCuentas.Rows = 2
        While Not rsCta.EOF
            grdCuentas.AdicionaFila
            nItem = grdCuentas.Row
            grdCuentas.TextMatrix(nItem, 1) = rsCta("cCtaCod")
            grdCuentas.TextMatrix(nItem, 2) = Format$(rsCta("dApertura"), "dd-mm-yyyy")
            grdCuentas.TextMatrix(nItem, 3) = rsCta("cEstado")
            grdCuentas.TextMatrix(nItem, 4) = rsCta("cProducto")
            grdCuentas.TextMatrix(nItem, 5) = rsCta("nFirmas")
            grdCuentas.TextMatrix(nItem, 6) = rsCta("nTasaInteres")
            rsCta.MoveNext
        Wend
    '    txtCuenta.Enabled = False
    '    CmdBuscar.Enabled = False
    '    ObtieneDatosPersona sCuenta
    'Else
    '    MsgBox "Cuenta Anulada o Cancelada", vbInformation, "Aviso"
    '    txtCuenta.SetFocusCuenta
    'End If
Else
    MsgBox "No existen cuentas", vbInformation, "Aviso"
    'txtCuenta.SetFocusCuenta
End If
rsCta.Close
Set rsCta = Nothing
End Sub


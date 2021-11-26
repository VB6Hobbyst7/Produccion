VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTarjetaCambioClave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de Clave "
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7260
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraHistoria 
      Caption         =   "Historia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1905
      Left            =   120
      TabIndex        =   9
      Top             =   2100
      Width           =   3585
      Begin SICMACT.FlexEdit grdTarjetaEstado 
         Height          =   1590
         Left            =   120
         TabIndex        =   10
         Top             =   210
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   2805
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Fecha-Estado"
         EncabezadosAnchos=   "250-1000-2000"
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
         ColWidth0       =   255
         RowHeight0      =   300
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   2880
      Width           =   1680
   End
   Begin VB.Frame fraTarjeta 
      Height          =   750
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3405
      Begin MSMask.MaskEdBox txtTarjeta 
         Height          =   375
         Left            =   945
         TabIndex        =   5
         Top             =   210
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "####-####-####-####"
         Mask            =   "####-####-####-####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   105
         TabIndex        =   6
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Frame fraPersona 
      Caption         =   "Persona"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1170
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   7155
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCliente 
         Height          =   855
         Left            =   105
         TabIndex        =   3
         Top             =   210
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   1508
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdRegCard 
      Caption         =   "Cambiar Clave"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblMensaje 
      Caption         =   "Presiones <F11> para activar la lectura de la tarjeta...."
      Height          =   495
      Left            =   3675
      TabIndex        =   8
      Top             =   128
      Width           =   3375
   End
End
Attribute VB_Name = "FrmTarjetaCambioClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancelar_Click()
LimpiaPantalla
End Sub

Private Sub cmdRegCard_Click()
MsgBox "Pase su Tarjeta por la Banda Magnetica", vbInformation, "AVISO"
If Trim(Replace(txtTarjeta.Text, "-", "")) = "" Then
    Exit Sub
End If
Get_Banda_Tarj
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim cTarjeta As String
Dim ObjTarj As COMNCaptaServicios.NCOMCaptaTarjeta
Dim I_RESULT As Integer
Dim S_RESULT As String
Dim L_RESULT As Long

If KeyCode = vbKeyF11 Then 'F11
    MsgBox "Pase la Tarjeta", vbInformation, "AVISO"
    Set ObjTarj = New COMNCaptaServicios.NCOMCaptaTarjeta
    'VerificaTarjetaActivaReg
    CargaValoresPinPad_ACS
    I_RESULT = OpenReader()
    S_RESULT = CStr(I_RESULT)

    'I_RESULT = ReadMagneticCard(I_MIN, I_TRACK)
    I_RESULT = ReadMagneticCard(30, 1)
    S_RESULT = CStr(I_RESULT)
    If I_RESULT > 0 Then
        L_RESULT = ReadTrack(1, 0, 110)
        S_RESULT = sANSIPtrZToVBString(L_RESULT)
        Me.txtTarjeta = Format(Mid(S_RESULT, 3, 16), "####-####-####-####")
        cTarjeta = Mid(S_RESULT, 3, 16)
        If Not IsNumeric(Mid(S_RESULT, 3, 16)) Then
            MsgBox "Tarjeta Invalida", vbInformation, "AVISO"
            CloseReader
            Set ObjTarj = Nothing
            Exit Sub
        End If
        If Not ObjTarj.VerificaTarjetaActiva(cTarjeta) Then
            MsgBox "Esta tarjeta no se encuentra activa", vbInformation, "AVISO"
            CloseReader
            Set ObjTarj = Nothing
            Exit Sub
        End If
    End If
    Set ObjTarj = Nothing
    SetupGridCliente
    CloseReader
    Call txtTarjeta_KeyPress(13)
    'CloseReader
    'Get_Banda_Tarj

End If

End Sub

Public Sub LimpiaPantalla()
grdCliente.Clear
grdCliente.Rows = 2
SetupGridCliente
'grdClienteTarj.Clear
'grdClienteTarj.Rows = 2
'grdClienteTarj.FormaCabecera
grdTarjetaEstado.Clear
grdTarjetaEstado.Rows = 2
grdTarjetaEstado.FormaCabecera
txtTarjeta.Text = "____-____-____-____"
fraTarjeta.Enabled = True
'fraEstado.Enabled = False
cmdCancelar.Enabled = False
'lblpassw1 = ""
'lblpassw2 = ""
'lblTrack1 = ""
'txtGlosa = ""
'cmdGrabar.Enabled = False
End Sub
Public Sub SetupGridCliente()
Dim I As Integer
For I = 1 To grdCliente.Rows - 1
    grdCliente.MergeCol(I) = True
Next I
grdCliente.MergeCells = flexMergeFree
grdCliente.BandExpandable(0) = True
grdCliente.Cols = 9
grdCliente.ColWidth(0) = 100
grdCliente.ColWidth(1) = 3500
grdCliente.ColWidth(2) = 3500
grdCliente.ColWidth(3) = 1500
grdCliente.ColWidth(4) = 1000
grdCliente.ColWidth(5) = 600
grdCliente.ColWidth(6) = 1500
grdCliente.ColWidth(7) = 0
grdCliente.ColWidth(8) = 0
grdCliente.TextMatrix(0, 1) = "Nombre"
grdCliente.TextMatrix(0, 2) = "Dirección"
grdCliente.TextMatrix(0, 3) = "Zona"
grdCliente.TextMatrix(0, 4) = "Fono"
grdCliente.TextMatrix(0, 5) = "ID"
grdCliente.TextMatrix(0, 6) = "ID N°"
End Sub
Private Sub Get_Banda_Tarj()
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
    CargaValoresPinPad_ACS

    I_RESULT = OpenReader()
    S_RESULT = CStr(I_RESULT)

    'I_RESULT = ReadMagneticCard(I_MIN, I_TRACK)
    I_RESULT = ReadMagneticCard(60, 1)
    S_RESULT = CStr(I_RESULT)
    If I_RESULT > 0 Then
        L_RESULT = ReadTrack(1, 0, 110)
        S_RESULT = sANSIPtrZToVBString(L_RESULT)
        cTarjeta = Mid(S_RESULT, 3, 16)
        If Not IsNumeric(Mid(S_RESULT, 3, 16)) Then
            MsgBox "Tarjeta Invalida", vbInformation, "AVISO"
            CloseReader
            Exit Sub
        End If
        If Not ObjTarj.VerificaTarjetaActiva(cTarjeta) Then
            MsgBox "Esta tarjeta no se encuentra activa", vbInformation, "AVISO"
            CloseReader
            Set ObjTarj = Nothing
            Exit Sub
        End If
        txtTarjeta.Text = Format(Mid(S_RESULT, 3, 16), "####-####-####-####")
        S_PAN = cTarjeta
        S_PVKI = gsPVKi
        S_IPSERVER = gsServPindVerify
        S_PVV = ObjTarj.Get_Tarj_PVV(S_PAN)
        MsgBox "Ingrese su clave", vbInformation, "AVISO"
        I_RESULT = CheckPIN(I_TIMEOUT, S_IPSERVER, S_PAN, S_PVKI, S_PVV)
        'L_RESULT = ""
        If I_RESULT = 1 Then
            If IsNumeric(Mid(S_RESULT, 3, 16)) Then
                I_CHNG_PASS = 1
                LightOn (30)
                MsgBox "Ingrese la nueva Clave", vbInformation
                L_RESULT = ChangePIN(I_TIMEOUT, S_IPSERVER, S_PAN, S_PVKI, I_CHNG_PASS, "NUEVA CLAVE")
                S_RESULT = CStr(I_RESULT)
                S_RESULT = sANSIPtrZToVBString(L_RESULT)
                LightOn (0)
                If IsNumeric(Trim(S_RESULT)) Then
                    Call ObjTarj.Put_Tarj_PVV(S_PAN, S_RESULT)
                    sMovNro = clsMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
                    Call ObjTarj.RegistraCambioPind(S_PAN, Format(FechaHora(gdFecSis & " " & Time), "mm/dd/YYYY HH:MM:SS"), sMovNro)
                Else
                    MsgBox "Error al Cambiar la Clave", vbInformation
                End If
            End If
        ElseIf I_RESULT = -1 Then
            MsgBox "Clave incorrecta", vbInformation, "AVISO"
        ElseIf I_RESULT = -2 Then
            MsgBox "No existe conexion con el servidor PINDVERIFY", vbInformation, "AVISO"
        Else
            MsgBox "Error de espera"
        End If
    End If
    CloseReader
    Set ObjTarj = Nothing
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
SetupGridCliente
End Sub

Private Sub txtTarjeta_KeyPress(KeyAscii As Integer)
Dim rsTarj As New ADODB.Recordset
Dim ObjTarj As COMNCaptaServicios.NCOMCaptaTarjeta
Dim sPersona As String
Dim cTarjeta As String
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
If KeyAscii = 13 Then
    
    cTarjeta = Trim(Replace(Me.txtTarjeta.Text, "-", ""))
    If Trim(cTarjeta) = "" Then
        MsgBox "Nro de Tarjeta Incorrecta", vbInformation, "AVISO"
        LimpiaPantalla
        Exit Sub
    End If
    Set ObjTarj = New COMNCaptaServicios.NCOMCaptaTarjeta
    If Not ObjTarj.VerificaTarjetaActiva(cTarjeta) Then
        MsgBox "Tarjeta Invalida", vbInformation, "AVISO"
        Set ObjTarj = Nothing
        LimpiaPantalla
        Exit Sub
    End If
    Set rsTarj = New ADODB.Recordset
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    sPersona = ObjTarj.Get_Tarj_Cod_Titular(cTarjeta)
    If Trim(sPersona) <> "" Then
        Set rsTarj = clsMant.GetDatosPersona(sPersona)
        Set grdCliente.Recordset = rsTarj
        Set rsTarj = ObjTarj.Get_Tarj_HistorialEst(cTarjeta)
        Set grdTarjetaEstado.Recordset = rsTarj
    End If
    Set clsMant = Nothing
    SetupGridCliente
    Set rsTarj = Nothing
    Set ObjTarj = Nothing
    Me.cmdRegCard.Enabled = True
    Me.cmdCancelar.Enabled = True
End If
End Sub



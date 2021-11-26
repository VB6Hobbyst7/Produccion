VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DB786848-D4E8-474E-A2C2-DCBC1D43DA22}#2.0#0"; "OCXTarjeta.ocx"
Begin VB.Form frmSegTarjetaAfiliacionExt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extorno de Afiliación de Seguro de Tarjeta de Débito"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   Icon            =   "frmSegTarjetaAfiliacionExt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "Extornar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   2040
      TabIndex        =   17
      Top             =   3600
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   3240
      TabIndex        =   16
      Top             =   3600
      Width           =   1170
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   4560
      TabIndex        =   0
      Top             =   3600
      Width           =   1050
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos de Afiliación"
      TabPicture(0)   =   "frmSegTarjetaAfiliacionExt.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblNumTarjeta"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdLecTarj"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frmMotExtorno"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame frmMotExtorno 
         Caption         =   "Motivos del Extorno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2940
         Left            =   1560
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   2845
         Begin VB.ComboBox cmbMotivos 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmSegTarjetaAfiliacionExt.frx":0326
            Left            =   240
            List            =   "frmSegTarjetaAfiliacionExt.frx":0328
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox txtDetExtorno 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   240
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   1200
            Width           =   2415
         End
         Begin VB.CommandButton cmdExtContinuar 
            Caption         =   "&Continuar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   860
            TabIndex        =   20
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label lblExtCmb 
            BackStyle       =   0  'Transparent
            Caption         =   "Motivo"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Detalles del Extorno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   960
            Width           =   1575
         End
      End
      Begin VB.CommandButton CmdLecTarj 
         Caption         =   "Leer Tarjeta"
         Height          =   345
         Left            =   3960
         TabIndex        =   8
         Top             =   480
         Width           =   1290
      End
      Begin VB.Frame Frame2 
         Height          =   2295
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   5055
         Begin VB.Label lblDOI 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1320
            TabIndex        =   15
            Top             =   675
            Width           =   1335
         End
         Begin VB.Label lblTitular 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1320
            TabIndex        =   14
            Top             =   315
            Width           =   3495
         End
         Begin VB.Label Label9 
            Caption         =   "Titular:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "DOI :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblCuenta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1320
            TabIndex        =   11
            Top             =   1395
            Width           =   2055
         End
         Begin VB.Label lblNumCert 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1320
            TabIndex        =   10
            Top             =   1035
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Certificado:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Importe Prima:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   1785
            Width           =   1095
         End
         Begin VB.Label lblImporte 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1320
            TabIndex        =   4
            Top             =   1755
            Width           =   930
         End
         Begin VB.Label Label8 
            Caption         =   "Cuenta Debitar:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   1440
            Width           =   1215
         End
      End
      Begin VB.Label lblNumTarjeta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   960
         TabIndex        =   9
         Top             =   480
         Width           =   2850
      End
      Begin VB.Label Label1 
         Caption         =   "Tarjeta :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   535
         Width           =   735
      End
   End
   Begin OCXTarjeta.CtrlTarjeta Tarjeta 
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   3600
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmSegTarjetaAfiliacionExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmSegTarjetaAfiliacion
'** Descripción : Formulario para extornar afiliaciones seguros a las tarjetas de débito creado
'**               segun TI-ERS029-2013
'** Creación : JUEZ, 20130523 09:00:00 AM
'**********************************************************************************************

Option Explicit
Dim oNSegTar As COMNCaptaGenerales.NCOMSeguros
Dim oDSegTar As COMDCaptaGenerales.DCOMSeguros
Dim rs As ADODB.Recordset
Dim sOpeCod As String
Dim nMovNro As Long
Dim nProd As String
Dim sOpeCodExt As String
Dim sNumTarj As String
Dim sPVV As String
Dim sPVVOrig As String

Public Sub Inicia(ByVal psOpeCod As CaptacOperacion)
sOpeCod = psOpeCod
Me.Show 1
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
End Sub

Private Sub cmdExtContinuar_Click()
Dim lbResultadoVisto As Boolean
Dim loVistoElectronico As frmVistoElectronico
Set loVistoElectronico = New frmVistoElectronico
Dim oNCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim oNContFunc As COMNContabilidad.NCOMContFunciones
Dim sMovNro As String
Dim lsBoleta As String
'APRI20171027 ERS028-2017
Dim clsCap As COMDCaptaGenerales.DCOMCaptaGenerales
Dim lnMovNro As Long
'END APRI

    '***CTI3 (FERIMORO)   02102018
    If cmbMotivos.ListIndex = -1 Or Len(txtDetExtorno.Text) <= 0 Then
        MsgBox "Debe ingresar el motivo y/o detalle del Extorno", vbInformation, "Aviso"
        Exit Sub
    End If

    lbResultadoVisto = loVistoElectronico.Inicio(3, sOpeCod, , , nMovNro)
    If Not lbResultadoVisto Then
        cmbMotivos.ListIndex = -1
        txtDetExtorno.Text = ""
        frmMotExtorno.Visible = False
        Exit Sub
    End If
    
    If MsgBox("¿Desea extornar la operación?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Set oNContFunc = New COMNContabilidad.NCOMContFunciones
        sMovNro = oNContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set oNContFunc = Nothing
        
        '***CTI3 (FERIMORO)   02102018
        Dim DatosExtorna(1) As String
        
        frmMotExtorno.Visible = False
        DatosExtorna(0) = cmbMotivos.Text
        DatosExtorna(1) = txtDetExtorno.Text
        
        Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
        If Mid(Trim(lblCuenta.Caption), 6, 3) = gCapAhorros Then
            oNCapMov.CapExtornoCargoAho nMovNro, gCapExtCargoAfilSegTarjetaAho, Trim(lblCuenta.Caption), sMovNro, "Cuenta = " & lblCuenta.Caption & ", Tarjeta = " & lblNumTarjeta.Caption, CDbl(lblImporte.Caption), , , , gsNomAge, sLpt, gsCodCMAC, , , , , lsBoleta, , , , gbImpTMU, , , , , , , DatosExtorna 'CTI3 30102018
        Else
            oNCapMov.CapExtornoCargoCTS nMovNro, gCapExtCargoAfilSegTarjetaCTS, Trim(lblCuenta.Caption), sMovNro, "Cuenta = " & lblCuenta.Caption & ", Tarjeta = " & lblNumTarjeta.Caption, CDbl(lblImporte.Caption), , , gsNomAge, sLpt, gsCodCMAC, gsCodAge, lsBoleta, gbImpTMU, , , DatosExtorna 'CTI3 30102018
        End If
        Set oNCapMov = Nothing
        
        Set oNSegTar = New COMNCaptaGenerales.NCOMSeguros
            'Call oNSegTar.ActualizaEstadoSegTarjetaAfiliacion(nMovNro, 503)
        
         'APRI20171027 ERS028-2017
        Set clsCap = New COMDCaptaGenerales.DCOMCaptaGenerales
        lnMovNro = clsCap.GetnMovNro(sMovNro)
        Call oNSegTar.ActualizaEstadoSegTarjetaAfiliacion(Trim(lblNumCert.Caption), lnMovNro, 503)
        'END APRI
        
        Set oNSegTar = Nothing
        
        If Trim(lsBoleta) <> "" Then
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, lsBoleta
                Print #nFicSal, ""
            Close #nFicSal
        End If
        MsgBox "Extorno finalizado", vbInformation, "Aviso"
        Limpiar
    End If
End Sub

Private Sub cmdExtornar_Click()
CmdLecTarj.Enabled = False
frmMotExtorno.Visible = True
cmdExtornar.Enabled = False
End Sub

Private Sub CmdLecTarj_Click()
Me.Caption = "Extorno de Afiliación - PASE LA TARJETA"

sNumTarj = Mid(Tarjeta.LeerTarjeta("PASE LA TARJETA", gnTipoPinPad, gnPinPadPuerto, gnTimeOutAg), 2, 16)
lblNumTarjeta.Caption = Left(sNumTarj, 6) & " - - - - - - " & Right(sNumTarj, 4)
Me.Caption = "Extorno de Afiliación de Seguro de Tarjeta de Débito"

If sNumTarj = "" Then
    lblNumTarjeta.Caption = ""
    MsgBox "No hay conexion con el PINPAD o no paso la tarjeta, Intente otra vez", vbInformation, "Aviso"
    Limpiar
    Exit Sub
End If

If Not frmATMCargaCuentas.ExisteTarjeta_Selec(sNumTarj) Then
    lblNumTarjeta.Caption = ""
    MsgBox "La Tarjeta N° " & sNumTarj & " no Existe, Intente otra vez", vbInformation, "Aviso"
    Limpiar
    Exit Sub
End If

If Not frmATMCargaCuentas.ValidaEstadoTarjeta_Selec(sNumTarj) Then
    lblNumTarjeta.Caption = ""
    MsgBox "La Tarjeta no esta activa", vbInformation, "MENSAJE DEL SISTEMA"
    Limpiar
    Exit Sub
End If

If Left(sNumTarj, 3) <> "ERR" Then
    sPVV = frmATMCargaCuentas.RecuperaPVV(sNumTarj)
    sPVVOrig = frmATMCargaCuentas.RecuperaPVVOrig(sNumTarj)
    Call CargaDatos
End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Limpiar()
lblNumTarjeta.Caption = ""
lblTitular.Caption = ""
lblDOI.Caption = ""
lblNumCert.Caption = ""
lblCuenta.Caption = ""
lblImporte.Caption = ""
cmdExtornar.Enabled = False
frmMotExtorno.Visible = False
End Sub

Private Sub CargaDatos()
Set oDSegTar = New COMDCaptaGenerales.DCOMSeguros
    Set rs = oDSegTar.RecuperaSegAfiliacionTarjetaExtorno(sNumTarj, gdFecSis)
    If Not (rs.EOF And rs.BOF) Then
        nMovNro = rs!nMovNroReg
        lblTitular.Caption = rs!cPersNombre
        lblDOI.Caption = rs!cPersIDnro
        lblNumCert.Caption = rs!cNumCertificado
        lblCuenta.Caption = rs!cCtaCodDebito
        lblImporte.Caption = Format(rs!nMontoPrima, "#,##0.00")
        cmdExtornar.Enabled = True
    Else
        MsgBox "No se encontraron datos", vbInformation, "Aviso"
    End If
Set oDSegTar = Nothing
End Sub
'******CTI3 (ferimoro) 18102018
Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

Set oCons = New COMDConstantes.DCOMConstantes

Set R = oCons.ObtenerConstanteExtornoMotivo

Set oCons = Nothing
Call Llenar_Combo_MotivoExtorno(R, cmbMotivos)

End Sub

Private Sub Form_Load()
Call CargaControles 'cti3
End Sub


Private Sub txtDetExtorno_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then SendKeys "{TAB}": Exit Sub
End Sub

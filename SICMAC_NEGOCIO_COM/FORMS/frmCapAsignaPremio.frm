VERSION 5.00
Begin VB.Form frmCapAsignaPremio 
   Caption         =   "Asignación de Premio"
   ClientHeight    =   4260
   ClientLeft      =   4380
   ClientTop       =   1845
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5340
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Premio"
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
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   5055
      Begin VB.ComboBox cboPre 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   4095
      End
      Begin VB.ComboBox cboCamp 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label lblCost 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2880
         TabIndex        =   23
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Precio"
         Height          =   255
         Left            =   2400
         TabIndex        =   22
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblCant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1560
         TabIndex        =   21
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Premio"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Campaña"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cuenta"
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
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3735
         _extentx        =   6588
         _extenty        =   661
         texto           =   "Cuenta N°"
         enabledcmac     =   -1  'True
         enabledcta      =   -1  'True
         enabledprod     =   -1  'True
         enabledage      =   -1  'True
      End
      Begin VB.Label lblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   600
         TabIndex        =   17
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label lblDias 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4080
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Días"
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblPlazo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4080
         TabIndex        =   13
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Plazo (días)"
         Height          =   255
         Left            =   3120
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblVencimiento 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1080
         TabIndex        =   11
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Vencimiento"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblApertura 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1080
         TabIndex        =   9
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Apertura"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Titular"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmCapAsignaPremio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nIdCamp, nIdPre As Integer
'By capi 21012009
Dim objPista As COMManejador.Pista
'End by

Public Sub Inicia(ByVal pcCodPro As Producto)
    'Inicio de objetos
    lblCliente.Caption = ""
    lblApertura.Caption = ""
    lblVencimiento.Caption = ""
    lblPlazo.Caption = ""
    lblDias.Caption = ""
    cmdCancelar.Enabled = False
    cmdGrabar.Enabled = False
    cboCamp.Enabled = False
    cboPre.Enabled = False
    txtCuenta.CMAC = gsCodCMAC
    txtCuenta.Age = ""
    txtCuenta.Prod = Trim(Str(pcCodPro))
    txtCuenta.cuenta = ""
    txtCuenta.EnabledCMAC = False
    txtCuenta.EnabledProd = False
End Sub
'

Private Sub cboCamp_Click()
    Dim objPre As COMDCaptaGenerales.DCOMCampanas
    Dim rs As ADODB.Recordset
    Set objPre = New COMDCaptaGenerales.DCOMCampanas
    Set rs = New ADODB.Recordset
    
    'Llena segundo combo
    If cboCamp.Text <> "" Then
        nIdCamp = CInt(Right(cboCamp.Text, 3))
        Set rs = objPre.GetCapCampanaPremio(nIdCamp)
           If rs.EOF <> True And rs.BOF <> True Then
                With cboPre
                    .Clear
                    While Not rs.EOF
                        .AddItem rs.Fields("cDescripcion") & Space(100) & "|" & rs.Fields("nTipoPremio") & "|" & rs.Fields("nMontoPremio")
                        rs.MoveNext
                    Wend
                End With
            End If
     End If
    Set objPre = Nothing
End Sub

Private Sub cboPre_Click()
    Dim sDatos() As String
    'Obtenemos datos del premio seleccionado
    If cboPre.Text <> "" Then
        sDatos = Split(cboPre.Text, "|")
        lblCant.Caption = 1
        lblCost.Caption = Format(sDatos(2), "###,###.00")
        nIdPre = CInt(sDatos(1))
    End If
End Sub

'
Private Sub cmdCancelar_Click()
    cboCamp.Clear
    cboPre.Clear
    lblCant.Caption = ""
    lblCost.Caption = ""
    Inicia gCapPlazoFijo
    txtCuenta.SetFocusAge
End Sub
'
Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub CmdGrabar_Click()
    Dim objReg As COMDCaptaGenerales.DCOMCampanas
    If txtCuenta.Texto = "" Then
        MsgBox "Debe de ingresar la Cuenta", vbCritical, "SICMACM - Aviso"
        txtCuenta.SetFocusAge
        Exit Sub
    End If
    If cboCamp.Text = "" Then
        MsgBox "Debe seleccionar la campaña", vbCritical, "SICMACM - Aviso"
        cboCamp.SetFocus
        Exit Sub
    End If
    If cboPre.Text = "" Then
        MsgBox "Debe seleccionar el premio", vbCritical, "SICMACM - Aviso"
        cboPre.SetFocus
        Exit Sub
    End If
    If MsgBox("¿Esta seguro de guardar la información?", vbQuestion + vbYesNo, "SICMACM - Confirmar") = vbYes Then
        Set objReg = New COMDCaptaGenerales.DCOMCampanas
        objReg.RegCtaPremio txtCuenta.GetCuenta, gdFecSis, nIdCamp, nIdPre, gsCodUser
        'By Capi 21012009
        objPista.InsertarPista gsOpeCod, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, , Trim(txtCuenta.NroCuenta), gCodigoCuenta
        'End by
        cmdCancelar_Click
    End If
    Set objReg = Nothing
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    'By Capi 20012009
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCapAsignaPremio
    'End By

End Sub
'
Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sCta As String
        sCta = txtCuenta.NroCuenta
        ObtieneDatosCuenta sCta
    End If
End Sub
'
Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim rsCta As ADODB.Recordset, rsRel As New ADODB.Recordset
    Dim nEstado As COMDConstantes.CaptacEstado
    Dim nRow As Long
    Dim sMsg As String, sMoneda As String, sPersona As String
    Dim dUltRetInt As Date
    Dim bGarantia As Boolean
    Dim dRenovacion As Date, dApeReal As Date

    'Obtenemos la cuenta.
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)
    Set clsMant = Nothing
    If Not (rsCta.EOF And rsCta.BOF) Then
        lblApertura.Caption = Format$(rsCta("dApertura"), "dd mmm yyyy hh:mm")
        lblPlazo.Caption = Format$(rsCta("nPlazo"), "#,##0")
        lblVencimiento.Caption = Format(DateAdd("d", rsCta("nPlazo"), rsCta("dRenovacion")), "dd mmm yyyy")

        Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
        dUltRetInt = clsCap.GetFechaUltimoRetiroIntPF(sCuenta)
        lblDias = Format$(DateDiff("d", dUltRetInt, gdFecSis), "#0")
        Set clsCap = Nothing

        Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
        Set clsMant = Nothing

        If Not (rsRel.EOF And rsRel.BOF) Then
            Me.lblCliente.Caption = UCase(PstaNombre(rsRel("Nombre")))
        End If
        cmdCancelar.Enabled = True
        cmdGrabar.Enabled = True
        cboCamp.Enabled = True
        cboPre.Enabled = True
        LlenarListaCampanas
        cboCamp.SetFocus
    Else
        MsgBox "Nro de Cuenta Incorrecta", vbInformation, "SICMACM - Aviso"
    End If
End Sub

Sub LlenarListaCampanas()
    Dim objCamp As COMDCaptaGenerales.DCOMCampanas
    Dim rs As ADODB.Recordset
    
    Set objCamp = New COMDCaptaGenerales.DCOMCampanas
    Set rs = New ADODB.Recordset
    
    Set rs = objCamp.GetCapCampanas("A")
    
    If rs.EOF <> True And rs.BOF <> True Then
        With cboCamp
            .Clear
            While Not rs.EOF
                .AddItem rs.Fields("cDescripcion") & Space(100) & rs.Fields("IdCampana")
                rs.MoveNext
            Wend
        End With
    End If
    rs.Close
    Set rs = Nothing
    Set objCamp = Nothing
End Sub



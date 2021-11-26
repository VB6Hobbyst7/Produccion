VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCapServParametros 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "frmCapServParametros.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2610
      TabIndex        =   4
      Top             =   2160
      Width           =   915
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   1620
      TabIndex        =   3
      Top             =   2160
      Width           =   915
   End
   Begin VB.Frame fraParametro 
      Caption         =   "Parámetros"
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
      Height          =   1965
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4935
      Begin VB.PictureBox pctImg 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   4005
         ScaleHeight     =   750
         ScaleWidth      =   705
         TabIndex        =   8
         Top             =   315
         Width           =   735
      End
      Begin SICMACT.EditMoney txtMonto 
         Height          =   375
         Left            =   1710
         TabIndex        =   2
         Top             =   1350
         Width           =   2175
         _extentx        =   3836
         _extenty        =   661
         font            =   "frmCapServParametros.frx":030A
         text            =   "0"
         enabled         =   -1  'True
      End
      Begin VB.ComboBox cboTipoComision 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   945
         Width           =   2175
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   420
         Left            =   225
         TabIndex        =   0
         Top             =   315
         Width           =   3660
         _extentx        =   6456
         _extenty        =   741
         texto           =   "Cuenta N°:"
         enabledcmac     =   -1  'True
         enabledcta      =   -1  'True
         enabledprod     =   -1  'True
         enabledage      =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   225
         TabIndex        =   7
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Comisión :"
         Height          =   195
         Left            =   225
         TabIndex        =   6
         Top             =   1005
         Width           =   1080
      End
   End
   Begin MSComctlLib.ImageList imglst 
      Left            =   45
      Top             =   2115
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCapServParametros.frx":0332
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCapServParametros.frx":0784
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCapServParametros.frx":0A9E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCapServParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nServicio As COMDConstantes.CaptacInstServicios
Dim bNuevo As Boolean

Private Sub CargaDatosParametros()
Dim rsServ As ADODB.Recordset
Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
Dim oGen As COMDConstSistema.DCOMGeneral
Set oGen = New COMDConstSistema.DCOMGeneral
    Set rsServ = oGen.GetConstante(gCapServTipoComision)
Set oGen = Nothing

Do While Not rsServ.EOF
    cboTipoComision.AddItem rsServ("cDescripcion") & Space(50) & rsServ("nConsValor")
    rsServ.MoveNext
Loop
'-----------------comentado por avmm------------------------
'-----------------------------------------------------------

'Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
'    Set rsServ = clsServ.GetServicioParametros(nServicio)
'Set clsServ = Nothing
'If Not (rsServ.EOF And rsServ.BOF) Then
'    txtCuenta.NroCuenta = rsServ("cCtaCodAbono")
'    txtMonto.Text = Format$(rsServ("nComision"), "#,##0.00")
'    cboTipoComision.ListIndex = rsServ("nTipoComision") - 1
'    bNuevo = False
'Else
'    bNuevo = True
'End If
TxtCuenta.CMAC = gsCodCMAC
TxtCuenta.Prod = gCapAhorros
TxtCuenta.EnabledCMAC = False
TxtCuenta.EnabledProd = False
rsServ.Close
Set rsServ = Nothing
End Sub

Public Sub Inicia(ByVal nServ As COMDConstantes.CaptacInstServicios)
nServicio = nServ
Select Case nServicio
    Case gCapServSedalib
        Me.Caption = "Captaciones - Servicios - Parámetros - SEDALIB"
        Set pctImg.Picture = imglst.ListImages(1).Picture
    Case gCapServHidrandina
        Me.Caption = "Captaciones - Servicios - Parámetros - HIDRANDINA"
        Set pctImg.Picture = imglst.ListImages(3).Picture
    Case gCapServFideicomiso
        Me.Caption = "Captaciones - Servicios - Parámetros - FIDEICOMISO"
        Set pctImg.Picture = imglst.ListImages(2).Picture
    Case gCapServEdelnor
        Me.Caption = "Captaciones - Servicios - Parámetros - EDELNOR"
        Set pctImg.Picture = imglst.ListImages(3).Picture
End Select
CargaDatosParametros
Me.Show 1
End Sub

Private Sub cboTipoComision_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMonto.SetFocus
End If
End Sub

Private Sub cmdGrabar_Click()
Dim nTipoComision As COMDConstantes.CapServTipoComision
Dim nComision As Double
Dim sCuenta As String

If cboTipoComision.Text = "" Then
    MsgBox "Debe escoger un tipo de comision", vbInformation, "Aviso"
    cboTipoComision.SetFocus
    Exit Sub
End If
nComision = txtMonto.value
sCuenta = TxtCuenta.NroCuenta
nTipoComision = CLng(Trim(Right(cboTipoComision.Text, 2)))

If MsgBox("¿ Desea grabar la información ?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
    Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
    If bNuevo Then
        clsServ.AgregaServicioParametro nServicio, sCuenta, nComision, nTipoComision
    Else
        clsServ.ActualizaServicioParametro nServicio, sCuenta, nComision, nTipoComision
    End If
    Set clsServ = Nothing
    Unload Me
End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    CargaDatosParametros
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cboTipoComision.SetFocus
End If
End Sub

Private Sub txtMonto_GotFocus()
txtMonto.MarcaTexto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
End If
End Sub

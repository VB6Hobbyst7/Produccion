VERSION 5.00
Begin VB.Form frmCajaGenRegTransp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Comprobante de Traslado de Valores"
   ClientHeight    =   3585
   ClientLeft      =   375
   ClientTop       =   3870
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCajaGenRegTransp.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5985
      TabIndex        =   6
      Top             =   3105
      Width           =   1380
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4620
      TabIndex        =   5
      Top             =   3105
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos principales"
      Height          =   2970
      Left            =   75
      TabIndex        =   7
      Top             =   90
      Width           =   7305
      Begin VB.ComboBox cboEmpresa 
         Height          =   330
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1125
         Width           =   4545
      End
      Begin VB.TextBox txtNetoServ 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   1605
         Width           =   1590
      End
      Begin VB.TextBox txtMontoServ 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   2490
         Width           =   1590
      End
      Begin VB.TextBox txtNroComp 
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   5325
         MaxLength       =   15
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cboTipoTrans 
         Height          =   330
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   2790
      End
      Begin VB.Line Line1 
         X1              =   3840
         X2              =   7155
         Y1              =   2385
         Y2              =   2385
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Neto Servicio :"
         Height          =   210
         Left            =   3960
         TabIndex        =   16
         Top             =   1665
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Impuesto :"
         Height          =   210
         Left            =   3960
         TabIndex        =   15
         Top             =   2055
         Width           =   735
      End
      Begin VB.Label lblImpuesto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5265
         TabIndex        =   14
         Top             =   1995
         Width           =   1590
      End
      Begin VB.Label lblMontoTrans 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   2085
         TabIndex        =   13
         Top             =   1605
         Width           =   1620
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total Servicio :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3975
         TabIndex        =   12
         Top             =   2565
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Monto Transportado :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   210
         TabIndex        =   11
         Top             =   1665
         Width           =   1785
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "N° :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   4905
         TabIndex        =   10
         Top             =   285
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Empresa :"
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   1170
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo :"
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   735
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmCajaGenRegTransp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnMontoTransp As Currency
Dim lbOk As Boolean
Dim lsTipoTransp As CGTipoTransporte
Dim lsNroComp As String
Dim lsCodTransp As String
Dim lsDesTransp As String
Dim lnTotalServicio As Currency
Dim oTransp As NCajaGenTransp

Public Sub Inicio(ByVal pnMontoTransp As Currency)
lnMontoTransp = pnMontoTransp
Me.Show 1
End Sub

Private Sub cboEmpresa_Click()
txtNetoServ = Format(oTransp.GetMontosServicio(Right(cboEmpresa, 4), CCur(lblMontoTrans), NetoServicio, Mid(gsOpeCod, 3, 1), gnTipCambio), "#,#0.00")
lblImpuesto = Format(oTransp.GetMontosServicio(Right(cboEmpresa, 4), CCur(lblMontoTrans), IGV, Mid(gsOpeCod, 3, 1), gnTipCambio), "#,#0.00")
txtMontoServ = Format(oTransp.GetMontosServicio(Right(cboEmpresa, 4), CCur(lblMontoTrans), MontoServicio, Mid(gsOpeCod, 3, 1), gnTipCambio), "#,#0.00")
If Val(txtNetoServ) = 0 And Right(cboTipoTrans, 1) <> CGTipoTransportePropio Then
    txtNetoServ.Locked = False
Else
    txtNetoServ.Locked = True
End If
End Sub
Private Sub cboEmpresa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNetoServ.SetFocus
End If
End Sub
Private Sub cboTipoTrans_Click()
    CargaCombo cboEmpresa, oTransp.GetTranspTipo(Right(cboTipoTrans, 1))
    Select Case Right(cboTipoTrans, 1)
        Case CGTipoTransportePropio, CGTipoTransporteAlquilado
            txtNroComp = ""
            txtNroComp.Enabled = False
        Case Else
            txtNroComp.Enabled = True
    End Select
    If cboEmpresa.ListCount > 0 Then
        cboEmpresa.ListIndex = 0
    End If
End Sub
Private Sub cboTipoTrans_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboEmpresa.SetFocus
End If
End Sub
Private Sub cmdAceptar_Click()
If Valida = False Then Exit Sub
'If MsgBox("Desea realizar el Registro de Transporte de Valores??", vbYesNo + vbQuestion) = vbYes Then
    lbOk = True
    AsignaValores
    Unload Me
'End If
End Sub
Function Valida() As Boolean
Valida = True
If cboTipoTrans = "" Then
    MsgBox "Tipo de Transporte no Seleccionado", vbInformation, "Aviso"
    Valida = False
    cboTipoTrans.SetFocus
    Exit Function
End If
If Right(cboTipoTrans, 1) = CGTipoTransporteBlindado Then
    If Len(Trim(txtNroComp)) = 0 Then
        MsgBox "Nro de Comprobante no Ingresado", vbInformation, "Aviso"
        Valida = False
        txtNroComp.SetFocus
        Exit Function
    End If
Else
    If Len(Trim(txtNroComp)) = 0 Then
        If MsgBox("Nro de Comprobante no Ingresado" & vbCrLf & "Desea Continuar??", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            Valida = False
            If txtNroComp.Enabled Then txtNroComp.SetFocus
            Exit Function
        End If
    End If
End If
If cboEmpresa = "" Then
    MsgBox "Empresa de Transporte de Valores no Seleccionada", vbInformation, "Aviso"
    Valida = False
    cboEmpresa.SetFocus
    Exit Function
End If
If Right(cboTipoTrans, 1) <> CGTipoTransportePropio Then
    If Val(txtNetoServ) = 0 Then
        MsgBox "Monto Neto de Servicio no Ingresado", vbInformation, "Aviso"
        Valida = False
        txtNetoServ.SetFocus
        Exit Function
    End If
    If Val(txtMontoServ) = 0 Then
        MsgBox "Monto total de Servicio no ha sido Calculado", vbExclamation, "Aviso"
        Valida = False
        txtNetoServ.SetFocus
        Exit Function
    End If
    If CCur(txtMontoServ) > CCur(lblMontoTrans) Then
        MsgBox "Monto de Servicio no puede ser mayor que el Transportado", vbExclamation, "Aviso"
        Valida = False
        txtNetoServ.SetFocus
        Exit Function
    End If
End If
    



End Function
Sub AsignaValores(Optional ByVal pbAsigna As Boolean = True)
If pbAsigna Then
    lsTipoTransp = Val(Right(cboTipoTrans, 1))
    lsNroComp = Trim(txtNroComp)
    lsCodTransp = Trim(Right(cboEmpresa, 4))
    lsDesTransp = Trim(Left(cboEmpresa, 50))
    lnTotalServicio = txtMontoServ
Else
    lsTipoTransp = -1
    lsNroComp = ""
    lsCodTransp = ""
    lsDesTransp = ""
    lnTotalServicio = 0
End If
End Sub
Private Sub CmdSalir_Click()
lbOk = False
AsignaValores False
Unload Me
End Sub

Private Sub Form_Load()
Dim oGen As dgeneral
CentraForm Me

Set oGen = New dgeneral
Set oTransp = New NCajaGenTransp

CambiaTamañoCombo Me.cboTipoTrans
CargaCombo cboTipoTrans, oGen.GetConstante(gCGTipoTransporte)

lblMontoTrans = Format(lnMontoTransp, "#,#0.00")
Set oGen = Nothing

End Sub
Public Property Get Ok() As Boolean
Ok = lbOk
End Property
Public Property Let Ok(ByVal vNewValue As Boolean)
lbOk = vNewValue
End Property
Public Property Get TipoTransp() As CGTipoTransporte
TipoTransp = lsTipoTransp
End Property
Public Property Let TipoTransp(ByVal vNewValue As CGTipoTransporte)
lsTipoTransp = vNewValue
End Property
Public Property Get NroComp() As String
NroComp = lsNroComp
End Property
Public Property Let NroComp(ByVal vNewValue As String)
lsNroComp = vNewValue
End Property
Public Property Get CodTransp() As String
CodTransp = lsCodTransp
End Property
Public Property Let CodTransp(ByVal vNewValue As String)
lsCodTransp = vNewValue
End Property
Public Property Get DesTransp() As String
DesTransp = lsDesTransp
End Property
Public Property Let DesTransp(ByVal vNewValue As String)
lsDesTransp = vNewValue
End Property
Public Property Get TotalServicio() As Currency
TotalServicio = lnTotalServicio
End Property
Public Property Let TotalServicio(ByVal vNewValue As Currency)
lnTotalServicio = vNewValue
End Property


Private Sub txtMontoServ_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAceptar.SetFocus
End If
End Sub

Private Sub txtNetoServ_GotFocus()
fEnfoque txtNetoServ
End Sub

Private Sub txtNetoServ_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtNetoServ, KeyAscii, 15, 2)
If KeyAscii = 13 Then
    If txtNetoServ.Locked = False Then
        txtMontoServ = Format(txtNetoServ, "#,#0.00")
    End If
    txtMontoServ.SetFocus
End If
End Sub
Private Sub txtNetoServ_LostFocus()
    txtNetoServ = Format(txtNetoServ, "#,#0.00")
End Sub

Private Sub txtNroComp_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    cboTipoTrans.SetFocus
End If
End Sub

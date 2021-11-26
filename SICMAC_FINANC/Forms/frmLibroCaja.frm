VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLibroCaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LIBRO CAJA AGENCIA"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   Icon            =   "frmLibroCaja.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTodos 
      Caption         =   "&Todos"
      Height          =   240
      Left            =   1410
      TabIndex        =   5
      Top             =   150
      Width           =   975
   End
   Begin MSMask.MaskEdBox mskFI 
      Height          =   300
      Left            =   4425
      TabIndex        =   4
      Top             =   60
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   3030
      TabIndex        =   3
      Top             =   885
      Width           =   1065
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1890
      TabIndex        =   2
      Top             =   885
      Width           =   1065
   End
   Begin Sicmact.TxtBuscar txtAge 
      Height          =   345
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   609
      Appearance      =   1
      BackColor       =   12648447
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
   Begin VB.Label lblAge 
      Height          =   270
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   5685
   End
End
Attribute VB_Name = "frmLibroCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chktodos_Click()
    Me.txtAge.Text = ""
    Me.lblAge.Caption = ""
End Sub

Private Sub chkTodos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.mskFI.SetFocus
End Sub

Private Sub cmdAceptar_Click()
    Dim oPrevio As PrevioFinan.clsPrevioFinan
    Set oPrevio = New PrevioFinan.clsPrevioFinan
    Dim lsCadena As String
    Dim oCaj As nCajero
    Set oCaj = New nCajero
    
    If Not IsDate(Me.mskFI.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFI.SetFocus
        Exit Sub
    End If
    If gbBitCentral Then
        lsCadena = oCaj.GetLibroCaja1(Trim(gsEmpresaCompleto), Me.txtAge.Text, Me.lblAge.Caption, CDate(Me.mskFI), gMonedaNacional, gbBitCentral) & oImpresora.gPrnSaltoPagina & oCaj.GetLibroCaja1(Trim(gsEmpresaCompleto), Me.txtAge.Text, Me.lblAge.Caption, CDate(Me.mskFI), gMonedaExtranjera, gbBitCentral)
    Else
        lsCadena = oCaj.GetLibroCaja(Trim("CAJA MUNCIPAL DE AHORRO Y CREDITO DE TRUJILLO S.A."), gsRUC, Me.txtAge.Text, Me.lblAge.Caption, CDate(Me.mskFI), gMonedaNacional, gbBitCentral) & oImpresora.gPrnSaltoPagina & oCaj.GetLibroCaja(Trim("CAJA MUNCIPAL DE AHORRO Y CREDITO DE TRUJILLO S.A."), gsRUC, Me.txtAge.Text, Me.lblAge.Caption, CDate(Me.mskFI), gMonedaExtranjera, gbBitCentral)
    End If
    
    'By Capi 07102008
    'oPrevio.Show lsCadena, Caption, True
    EnviaPrevio lsCadena, Caption, 66, True

    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oAge As New DActualizaDatosArea
    Me.txtAge.rs = oAge.GetAgencias()
    If Not gbBitCentral Then
        Me.chkTodos.Visible = False
    End If
    
End Sub

Private Sub mskFF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.cmdAceptar.SetFocus
End Sub

Private Sub mskFI_GotFocus()
    mskFI.SelStart = 0
    mskFI.SelLength = 50
End Sub

Private Sub mskFI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.cmdAceptar.SetFocus
End Sub

Private Sub txtAge_EmiteDatos()
    Me.lblAge.Caption = txtAge.psDescripcion
End Sub

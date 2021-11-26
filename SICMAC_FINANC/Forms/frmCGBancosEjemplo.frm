VERSION 5.00
Begin VB.Form frmCGBancosEjemplo 
   Caption         =   "Form2"
   ClientHeight    =   1785
   ClientLeft      =   1665
   ClientTop       =   2055
   ClientWidth     =   7725
   LinkTopic       =   "Form2"
   ScaleHeight     =   1785
   ScaleWidth      =   7725
   Begin Sicmact.TxtBuscar txtBuscaEntidad 
      Height          =   360
      Left            =   90
      TabIndex        =   1
      Top             =   450
      Width           =   2580
      _ExtentX        =   4551
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
   Begin VB.Label lblDescIfTransf 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2730
      TabIndex        =   3
      Top             =   450
      Width           =   4290
   End
   Begin VB.Label lblDesCtaIfTransf 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   90
      TabIndex        =   2
      Top             =   870
      Width           =   6930
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta Banco :"
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
      Left            =   300
      TabIndex        =   0
      Top             =   210
      Width           =   1215
   End
End
Attribute VB_Name = "frmCGBancosEjemplo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Necesitas
'    DCajaCtasIF
'    NCajaCtaIF
'    DConecta

Dim oCtaIf As DCajaCtasIF

Private Sub Form_Load()
Set oCtaIf = New DCajaCtasIF
Dim lsFiltro As String
lsFiltro = "_1_[12]%"
txtBuscaEntidad.rs = oCtaIf.CargaCtasIF(gMonedaNacional, lsFiltro, MuestraCuentas)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oCtaIf = Nothing
End Sub

Private Sub txtBuscaEntidad_EmiteDatos()
Dim oNCtaIF As New NCajaCtaIF

lblDescIfTransf = oCtaIf.NombreIF(Mid(txtBuscaEntidad.Text, 4, 13))
lblDesCtaIfTransf = oNCtaIF.EmiteTipoCuentaIF(Mid(txtBuscaEntidad.Text, 18, Len(txtBuscaEntidad.Text))) & " " & txtBuscaEntidad.psDescripcion
Set oNCtaIF = Nothing
End Sub


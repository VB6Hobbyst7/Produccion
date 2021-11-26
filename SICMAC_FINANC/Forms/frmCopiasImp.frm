VERSION 5.00
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCopiasImp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copias de Cartas"
   ClientHeight    =   1725
   ClientLeft      =   3900
   ClientTop       =   2985
   ClientWidth     =   3810
   Icon            =   "frmCopiasImp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1170
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   3675
      Begin Spinner.uSpinner spnAsientos 
         Height          =   330
         Left            =   2700
         TabIndex        =   5
         Top             =   645
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin Spinner.uSpinner spnCartas 
         Height          =   345
         Left            =   2700
         TabIndex        =   4
         Top             =   240
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Copias de Carta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   3
         Top             =   285
         Width           =   1485
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Copias de Asiento Contable :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   675
         Width           =   2475
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1245
      Width           =   1770
   End
End
Attribute VB_Name = "frmCopiasImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnCopiasCartas As Integer
Dim lnCopiasAsientos As Integer

Private Sub cmdAceptar_Click()
lnCopiasCartas = spnCartas.Valor
lnCopiasAsientos = spnAsientos.Valor
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
spnCartas.Valor = 2
spnAsientos.Valor = 2
End Sub
Public Property Get CopiasCartas() As Integer
CopiasCartas = lnCopiasCartas
End Property
Public Property Let CopiasCartas(ByVal vNewValue As Integer)
lnCopiasCartas = vNewValue
End Property
Public Property Get CopiasAsientos() As Integer
CopiasAsientos = lnCopiasAsientos
End Property
Public Property Let CopiasAsientos(ByVal vNewValue As Integer)
lnCopiasAsientos = vNewValue
End Property

Private Sub spnAsientos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cmdAceptar.SetFocus
End If
End Sub

Private Sub spnCartas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.spnAsientos.SetFocus
End If
End Sub

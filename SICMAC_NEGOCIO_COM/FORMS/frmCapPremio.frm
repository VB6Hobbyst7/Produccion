VERSION 5.00
Begin VB.Form frmCapPremio 
   Caption         =   "Registra Premio"
   ClientHeight    =   2445
   ClientLeft      =   5340
   ClientTop       =   3570
   ClientWidth     =   4725
   FillColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4725
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Premio"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin SICMACT.EditMoney txtCosto 
         Height          =   315
         Left            =   480
         TabIndex        =   2
         Top             =   1080
         Width           =   800
         _ExtentX        =   1296
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12582912
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtRefe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3480
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtCantidad 
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtDesPre 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label5 
         Caption         =   "Referencia"
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Costo"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCapPremio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Inicia()
    txtDesPre.Text = ""
    txtCantidad.Text = ""
    txtRefe.Text = ""
    cmdAceptar.Enabled = False
    'txtDesPre.SetFocus
End Sub
'
Private Sub CmdAceptar_Click()
    Dim objPre As COMDCaptaGenerales.DCOMCampanas
    If txtDesPre.Text = "" Then
        MsgBox "Debe ingresar la descripciíon del premio.", vbCritical, "SICMACM"
        txtDesPre.SetFocus
        Exit Sub
    End If
    If txtCosto.Text = "0" Then
        MsgBox "Debe ingresar el costo del premio.", vbCritical, "SICMACM"
        txtCosto.SetFocus
        Exit Sub
    End If
    If txtCantidad.Text = "0" Then
        MsgBox "Debe ingresar el Stock.", vbCritical, "SICMACM"
        txtCantidad.SetFocus
        Exit Sub
    End If
    If txtRefe.Text = "0" Then
        MsgBox "Debe ingresar la referencia de entrega.", vbCritical, "SICMACM"
        txtRefe.SetFocus
        Exit Sub
    End If
    If MsgBox("Esta seguro de guardar la información.", vbInformation + vbYesNo, "SICMACM") = vbYes Then
    'ARCV 24-01-2007
        Set objPre = New COMDCaptaGenerales.DCOMCampanas
        objPre.RegPremio RTrim(txtDesPre.Text), CDbl(txtCosto.Text), CInt(txtCantidad.Text), CInt(txtRefe.Text)
        Me.Inicia
        Set objPre = Nothing
        If MsgBox("Desea ingresar otro premio.", vbInformation + vbYesNo, "SICMACM") = vbNo Then
            Unload Me
        End If
    End If
End Sub
'
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
'
Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub
'
Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtRefe.SetFocus
    Else
        KeyAscii = NumerosEnteros(KeyAscii, False)
    End If
End Sub
'
Private Sub txtDesPre_Change()
    If txtDesPre.Text <> "" Then
        cmdAceptar.Enabled = True
    End If
End Sub
'
Private Sub txtDesPre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCosto.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub
'
Private Sub txtRefe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    Else
        KeyAscii = NumerosEnteros(KeyAscii, False)
    End If
End Sub

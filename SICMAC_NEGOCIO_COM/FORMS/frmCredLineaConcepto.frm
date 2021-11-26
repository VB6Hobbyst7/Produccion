VERSION 5.00
Begin VB.Form frmCredLineaConcepto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Linea de Credito"
   ClientHeight    =   2475
   ClientLeft      =   3420
   ClientTop       =   3120
   ClientWidth     =   5340
   Icon            =   "frmCredLineaConcepto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   2745
      TabIndex        =   10
      Top             =   1935
      Width           =   1380
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   1305
      TabIndex        =   9
      Top             =   1935
      Width           =   1380
   End
   Begin VB.Frame fraSubFondo 
      Caption         =   "Sub Fondo de Linea de Credito "
      Height          =   1695
      Left            =   165
      TabIndex        =   11
      Top             =   135
      Visible         =   0   'False
      Width           =   5040
      Begin VB.TextBox TxtSubFondo 
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   720
         Width           =   3000
      End
      Begin VB.TextBox TxtAbrevSubFondo 
         Height          =   285
         Left            =   1815
         MaxLength       =   5
         TabIndex        =   6
         Top             =   1140
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Descripcion :"
         Height          =   240
         Left            =   135
         TabIndex        =   15
         Top             =   765
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Abreviatura :"
         Height          =   240
         Left            =   150
         TabIndex        =   14
         Top             =   1170
         Width           =   1005
      End
      Begin VB.Label Label5 
         Caption         =   "Codigo :"
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   345
         Width           =   900
      End
      Begin VB.Label LblsubFondo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1815
         TabIndex        =   12
         Top             =   330
         Width           =   465
      End
   End
   Begin VB.Frame fraFondo 
      Caption         =   "Fondo de Linea de Credito"
      Height          =   1695
      Left            =   165
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   5040
      Begin VB.TextBox TxtAbrev 
         Height          =   285
         Left            =   1815
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1140
         Width           =   735
      End
      Begin VB.ComboBox CmbFondo 
         Height          =   315
         Left            =   1815
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   3120
      End
      Begin VB.Label LblCodigoFondo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1815
         TabIndex        =   8
         Top             =   330
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Codigo :"
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   345
         Width           =   900
      End
      Begin VB.Label Label2 
         Caption         =   "Abreviatura :"
         Height          =   240
         Left            =   150
         TabIndex        =   4
         Top             =   1170
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Institucion Financiera :"
         Height          =   240
         Left            =   135
         TabIndex        =   1
         Top             =   765
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmCredLineaConcepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sResult As String
Dim sPersCod As String
Private Sub CargaCombo()
Dim oLinea As COMDCredito.DCOMLineaCredito
Dim R As adodb.Recordset
    CmbFondo.Clear
    Set oLinea = New COMDCredito.DCOMLineaCredito
    Set R = oLinea.RecuperaInstitucionesFinancieras
    Do While Not R.EOF
        CmbFondo.AddItem Trim(R!cPersNombre) & Space(100) & R!cPersCod
        R.MoveNext
    Loop
    Set oLinea = Nothing

End Sub

Public Function Fondo(ByVal psCodigoLinea As String, ByRef sPersCodTemp As String) As String
    fraFondo.Visible = True
    fraSubFondo.Visible = False
    LblCodigoFondo.Caption = psCodigoLinea
    TxtAbrev.Text = ""
    Me.Show 1
    Fondo = sResult
    sPersCodTemp = sPersCod
End Function

Public Function SubFondo(ByVal psCodigoLinea As String) As String
    fraFondo.Visible = False
    fraSubFondo.Visible = True
    LblsubFondo.Caption = psCodigoLinea
    TxtSubFondo.Text = ""
    TxtAbrevSubFondo.Text = ""
    Me.Show 1
    SubFondo = sResult
End Function


Private Sub CmbFondo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtAbrev.SetFocus
    End If
End Sub

Private Sub cmdAceptar_Click()
    If fraSubFondo.Visible = True Then
        If Trim(TxtSubFondo.Text) = "" Then
            MsgBox "Ingrese una Descripcion", vbInformation, "Aviso"
            TxtSubFondo.SetFocus
            Exit Sub
        End If
        If Trim(TxtAbrevSubFondo.Text) = "" Then
            MsgBox "Ingrese una Abreviatura", vbInformation, "Aviso"
            TxtAbrevSubFondo.SetFocus
            Exit Sub
        End If
    End If
    If fraFondo.Visible Then
        If CmbFondo.ListIndex = -1 Then
            MsgBox "Seleccione un Fondo", vbInformation, "Aviso"
            CmbFondo.SetFocus
            Exit Sub
        End If
        If Trim(TxtAbrev.Text) = "" Then
            MsgBox "Ingrese una Abreviatura", vbInformation, "Aviso"
            TxtAbrev.SetFocus
            Exit Sub
        End If
    End If
    If fraFondo.Visible Then
        sResult = Trim(Left(CmbFondo.Text, 50)) & Space(100 - Len(Trim(Left(CmbFondo.Text, 50)))) & Trim(LblCodigoFondo.Caption) & Space(100) & Trim(TxtAbrev.Text)
        sPersCod = Trim(Right(CmbFondo.Text, 20))
    Else
        sResult = Trim(Left(TxtSubFondo.Text, 50)) & Space(100 - Len(Trim(Left(TxtSubFondo.Text, 50)))) & Trim(LblsubFondo.Caption) & Space(100) & Trim(TxtAbrevSubFondo.Text)
    End If
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    sResult = ""
    Unload Me
End Sub

Private Sub Form_Load()
    Call CargaCombo
End Sub

Private Sub TxtAbrev_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        CmdAceptar.SetFocus
    End If
End Sub

Private Sub TxtAbrevSubFondo_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        CmdAceptar.SetFocus
    End If
End Sub

Private Sub TxtSubFondo_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        TxtAbrevSubFondo.SetFocus
    End If
End Sub

VERSION 5.00
Begin VB.UserControl ActXCodCta_Ant 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2820
   ScaleHeight     =   405
   ScaleWidth      =   2820
   Begin VB.TextBox txtEstado 
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame NroCuenta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   0
      TabIndex        =   3
      Top             =   -120
      Width           =   2820
      Begin VB.TextBox TxtCuenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   7
         TabIndex        =   2
         Top             =   165
         Width           =   855
      End
      Begin VB.TextBox TxtProd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1485
         MaxLength       =   3
         TabIndex        =   1
         Top             =   165
         Width           =   435
      End
      Begin VB.TextBox TxtAge 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1140
         MaxLength       =   2
         TabIndex        =   0
         Top             =   165
         Width           =   345
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta N°"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   4
         Top             =   210
         Width           =   1095
      End
   End
End
Attribute VB_Name = "ActXCodCta_Ant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event KeyPress(KeyAscii As Integer)


Private Sub txtAge_Change()
If Len(TxtAge) = 2 Then
    If TxtProd.Enabled And TxtProd.Visible Then
        TxtProd.SetFocus
    End If
End If
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    Dim Cad As String

    If KeyAscii = 8 Then
        Exit Sub
    End If

    If Len(TxtAge) = 1 Then
       txtCuenta.SetFocus
    End If

    If NumerosEnteros(KeyAscii) = 0 Or Len(TxtAge) = 2 Then
       If TxtProd.Enabled Then
            TxtProd.SetFocus
       End If
    End If

End Sub

Private Sub TxtAge_LostFocus()
    psAge = TxtAge.Text
End Sub

Private Sub txtCuenta_Change()
    If Len(txtCuenta) = 7 Then
    'UserControl.SetFocus
        'RaiseEvent KeyPress
       Exit Sub
    End If
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(txtCuenta) < 7 Then
         MsgBox "Número de Cuenta Incompleto", vbInformation, "Aviso"
         txtCuenta = "": txtCuenta.SetFocus
      Else
         RaiseEvent KeyPress(KeyAscii)
      End If
   End If
   If KeyAscii = 8 Then
      If Len(txtCuenta) < 1 Then
         If TxtProd.Enabled Then
            TxtProd.SetFocus
         Else
            If txtCuenta.Enabled Then
               TxtAge.SetFocus
            End If
         End If
      Else
         Exit Sub
      End If
   End If

   If NumerosEnteros(KeyAscii) = 0 Then
       Exit Sub
   End If
End Sub

Private Sub txtCuenta_LostFocus()
    psCuenta = txtCuenta.Text
End Sub

Private Sub txtProd_Change()
    If Len(TxtProd) = 3 Then
        If txtCuenta.Visible Then
            txtCuenta.SetFocus
        End If
    End If
End Sub

Private Sub txtprod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If Len(TxtProd) = 1 Or Len(TxtProd) = 0 Then
            TxtAge.SetFocus
        Else
            Exit Sub
        End If
    End If
    
    If KeyAscii = 13 And Len(TxtProd) = 3 Then
        txtCuenta.SetFocus
    End If

    If NumerosEnteros(KeyAscii) = 0 Or Len(TxtProd) = 3 Then
        txtCuenta.SetFocus
    End If

End Sub

Public Property Get NroCuenta() As String
    NroCuenta = Trim(TxtAge.Text) + Trim(TxtProd) + Trim(txtCuenta)
End Property

Public Property Let NroCuenta(ByVal vNewValue As String)
    TxtAge.Text = Mid(vNewValue, 1, 2)
    TxtProd.Text = Mid(vNewValue, 3, 3)
    txtCuenta.Text = Mid(vNewValue, 6, 7)
    'NroCuenta = vNewValue
    PropertyChanged "NroCuenta"
End Property

Private Sub TxtProd_LostFocus()
    psProd = TxtProd.Text
End Sub

Private Sub UserControl_InitProperties()
    TxtAge.Text = ""
    TxtProd.Text = ""
    txtCuenta = ""
'    txtAge.SetFocus
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'NroCuenta = PropBag.ReadProperty("Caption", Extender.Name)
    'txtAge.SetFocus
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'PropBag.WriteProperty "NroCuenta", NroCuenta, Extender.Name
End Sub

Public Property Get psProd() As String
    psProd = TxtProd.Text
End Property

Public Property Let psProd(ByVal vNewValue As String)
    TxtProd.Text = vNewValue
End Property

Public Property Get psAge() As String
    psAge = TxtAge.Text
End Property

Public Property Let psAge(ByVal vNewValue As String)
    TxtAge.Text = vNewValue
End Property

Public Property Get psCuenta() As String
    psCuenta = txtCuenta.Text
End Property

Public Property Let psCuenta(ByVal vNewValue As String)
    txtCuenta = vNewValue
End Property

Public Function GetCuenta() As String
    GetCuenta = Trim(TxtAge.Text) + Trim(TxtProd) + Trim(txtCuenta)
End Function

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal vNewEnabled As Boolean)
    UserControl.Enabled = vNewEnabled
    PropertyChanged "Enabled"
End Property

Public Property Get pbEnabledAge() As Boolean
    pbEnabledAge = TxtAge.Enabled
End Property

Public Property Let pbEnabledAge(ByVal vNewValue As Boolean)
    TxtAge.Enabled = vNewValue
End Property

Public Property Get pbEnabledProd() As Boolean
    pbEnabledProd = TxtProd.Enabled
End Property

Public Property Let pbEnabledProd(ByVal vNewValue As Boolean)
    TxtProd.Enabled = vNewValue
End Property

Public Property Get pbEnabledCta() As Boolean
    pbEnabledCta = txtCuenta.Enabled
End Property

Public Property Let pbEnabledCta(ByVal vNewValue As Boolean)
    txtCuenta.Enabled = vNewValue
End Property

Public Sub SetFocusAge()
    If TxtAge.Enabled Then TxtAge.SetFocus
End Sub

Public Sub SetFocusProd()
    If TxtProd.Enabled Then TxtProd.SetFocus
End Sub

Public Sub SetFocusCuenta()
    If txtCuenta.Enabled Then txtCuenta.SetFocus
End Sub

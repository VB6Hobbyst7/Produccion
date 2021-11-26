VERSION 5.00
Begin VB.UserControl ActXCodCta 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
   ScaleHeight     =   405
   ScaleWidth      =   3585
   Begin VB.Frame NroCuenta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   0
      TabIndex        =   4
      Top             =   -120
      Width           =   3585
      Begin VB.TextBox txtCMAC 
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
         Left            =   1050
         MaxLength       =   3
         TabIndex        =   0
         Top             =   165
         Width           =   435
      End
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
         Left            =   2310
         MaxLength       =   10
         TabIndex        =   3
         Top             =   165
         Width           =   1200
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
         Left            =   1860
         MaxLength       =   3
         TabIndex        =   2
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
         Left            =   1500
         MaxLength       =   2
         TabIndex        =   1
         Top             =   165
         Width           =   345
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Texto"
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
         Left            =   45
         TabIndex        =   5
         Top             =   195
         Width           =   495
      End
   End
End
Attribute VB_Name = "ActXCodCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event Change()
Dim lbBakAge As Boolean
Dim lbBakPro As Boolean
Dim lbBakCta As Boolean

Private Sub txtAge_Change()
If Len(Trim(txtAge)) = 2 Then
    If TxtProd.Enabled And TxtProd.Visible Then
        TxtProd.SetFocus
    End If
End If
End Sub

Private Sub TxtAge_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyTab
             If Len(Trim(txtAge)) = 0 Then
                txtAge.SetFocus
             End If
        Case vbKeyBack
             lbBakAge = True
             If Len(Trim(txtAge)) = 0 Then
                If txtCMAC.Enabled Then
                    txtCMAC.SetFocus
                End If
             End If
             Exit Sub
        Case Else
             If Len(Trim(txtAge)) = 1 Then
                txtCuenta.SetFocus
             End If
    End Select
    If NumerosEnteros(KeyAscii) = 0 Or Len(Trim(txtAge)) = 2 Then
       If TxtProd.Enabled Then
            TxtProd.SetFocus
       End If
    End If
End Sub

Private Sub TxtAge_LostFocus()
    If lbBakAge Then
    Else
       psAge = txtAge.Text
    End If
    lbBakAge = False
End Sub

Private Sub txtCMAC_Change()
If Len(Trim(txtCMAC)) = 3 Then
    If txtAge.Enabled And txtAge.Visible Then
        txtAge.SetFocus
    End If
End If
End Sub

Private Sub txtCMAC_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtCMAC_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyTab
             If Len(Trim(txtCMAC)) = 0 Then
                txtCMAC.SetFocus
             End If
        Case vbKeyBack
             Exit Sub
        Case Else
             If Len(Trim(txtAge)) = 1 Then
                txtCuenta.SetFocus
             End If
    End Select
    If NumerosEnteros(KeyAscii) = 0 Or Len(Trim(txtCMAC)) = 3 Then
       If txtAge.Enabled Then
            txtAge.SetFocus
       End If
    End If
End Sub

Private Sub txtCuenta_Change()
If Len(Trim(txtCuenta)) = 10 Then
    RaiseEvent Change
End If
End Sub

Private Sub TxtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtcuenta_LostFocus()
    If lbBakCta Then
    Else
       psCuenta = txtCuenta.Text
    End If
    lbBakCta = False
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn Or vbKeyTab
             If Len(Trim(txtCuenta)) < 10 Then
                MsgBox "Número de Cuenta Incompleto", vbInformation, "Aviso"
                txtCuenta = ""
                txtCuenta.SetFocus
             Else
                RaiseEvent KeyPress(KeyAscii)
             End If
        Case vbKeyBack
             lbBakCta = True
             If Len(Trim(txtCuenta)) = 0 Then
                If TxtProd.Enabled Then
                   TxtProd.SetFocus
                Else
                   If txtAge.Enabled Then
                      txtAge.SetFocus
                   End If
               End If
            Else
               Exit Sub
            End If
    End Select
    If NumerosEnteros(KeyAscii) = 0 Then
        Exit Sub
    End If
End Sub

Private Sub txtProd_Change()
    If Len(Trim(TxtProd)) = 3 And txtCuenta.Visible And txtCuenta.Enabled Then
       txtCuenta.SetFocus
    End If
End Sub

Private Sub TxtProd_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtprod_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyTab
             If Len(Trim(TxtProd)) = 0 Then
                TxtProd.SetFocus
             End If
        Case vbKeyBack
             lbBakPro = True
             If Len(Trim(TxtProd)) = 0 Then
                If txtAge.Enabled Then
                   txtAge.SetFocus
                End If
             Else
                Exit Sub
             End If
        Case vbKeyReturn
             If Len(Trim(TxtProd)) = 3 Then
                txtCuenta.SetFocus
             End If
    End Select

    If NumerosEnteros(KeyAscii) = 0 Or Len(Trim(TxtProd)) = 3 Then
        txtCuenta.SetFocus
    End If

End Sub

Public Property Get NroCuenta() As String
    NroCuenta = Trim(txtCMAC) & Trim(txtAge) & Trim(TxtProd) & Trim(txtCuenta)
End Property

Public Property Let NroCuenta(ByVal vNewValue As String)
    txtCMAC.Text = Mid(vNewValue, 1, 3)
    txtAge.Text = Mid(vNewValue, 4, 2)
    TxtProd.Text = Mid(vNewValue, 6, 3)
    txtCuenta.Text = Mid(vNewValue, 9, 10)
    PropertyChanged "NroCuenta"
End Property

Private Sub TxtProd_LostFocus()
    If lbBakPro Then
    Else
       psProd = TxtProd.Text
    End If
    lbBakPro = False
End Sub

Private Sub UserControl_InitProperties()
    txtAge.Text = ""
    TxtProd.Text = ""
    txtCuenta = ""
    lbBakAge = False
    lbBakPro = False
    lbBakCta = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lblTexto.Caption = PropBag.ReadProperty("Texto", "Texto")
    txtCMAC.Enabled = PropBag.ReadProperty("EnabledCMAC", Verdadero)
    txtCuenta.Enabled = PropBag.ReadProperty("EnabledCta", Verdadero)
    TxtProd.Enabled = PropBag.ReadProperty("EnabledProd", Verdadero)
    txtAge.Enabled = PropBag.ReadProperty("EnabledAge", Verdadero)
    txtCuenta.Text = PropBag.ReadProperty("Cuenta", "")
    txtAge.Text = PropBag.ReadProperty("Age", "")
    TxtProd.Text = PropBag.ReadProperty("Prod", "")
    txtCMAC.Text = PropBag.ReadProperty("CMAC", "")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Texto", lblTexto.Caption, "Texto")
    Call PropBag.WriteProperty("EnabledCMAC", txtCMAC.Enabled, Verdadero)
    Call PropBag.WriteProperty("EnabledCta", txtCuenta.Enabled, Verdadero)
    Call PropBag.WriteProperty("EnabledProd", TxtProd.Enabled, Verdadero)
    Call PropBag.WriteProperty("EnabledAge", txtAge.Enabled, Verdadero)
    Call PropBag.WriteProperty("Cuenta", txtCuenta.Text, "")
    Call PropBag.WriteProperty("Age", txtAge.Text, "")
    Call PropBag.WriteProperty("Prod", TxtProd.Text, "")
    Call PropBag.WriteProperty("CMAC", txtCMAC.Text, "")
End Sub

Public Function GetCuenta() As String
    GetCuenta = Trim(txtCMAC) & Trim(txtAge) & Trim(TxtProd) & Trim(txtCuenta)
End Function

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal vNewEnabled As Boolean)
    UserControl.Enabled = vNewEnabled
    PropertyChanged "Enabled"
End Property

Public Sub SetFocusAge()
    If txtAge.Enabled Then txtAge.SetFocus
End Sub

Public Sub SetFocusProd()
    If TxtProd.Enabled Then TxtProd.SetFocus
End Sub

Public Sub SetFocusCuenta()
    If txtCuenta.Enabled Then txtCuenta.SetFocus
End Sub
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=lblTexto,lblTexto,-1,Caption
Public Property Get Texto() As String
Attribute Texto.VB_Description = "Devuelve o establece el texto mostrado en la barra de título de un objeto o bajo el icono de un objeto."
    Texto = lblTexto.Caption
End Property

Public Property Let Texto(ByVal New_Texto As String)
    lblTexto.Caption() = New_Texto
    PropertyChanged "Texto"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCMAC,txtCMAC,-1,Enabled
Public Property Get EnabledCMAC() As Boolean
Attribute EnabledCMAC.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    EnabledCMAC = txtCMAC.Enabled
End Property

Public Property Let EnabledCMAC(ByVal New_EnabledCMAC As Boolean)
    txtCMAC.Enabled() = New_EnabledCMAC
    PropertyChanged "EnabledCMAC"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=TxtCuenta,TxtCuenta,-1,Enabled
Public Property Get EnabledCta() As Boolean
Attribute EnabledCta.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    EnabledCta = txtCuenta.Enabled
End Property

Public Property Let EnabledCta(ByVal New_EnabledCta As Boolean)
    txtCuenta.Enabled() = New_EnabledCta
    PropertyChanged "EnabledCta"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=TxtProd,TxtProd,-1,Enabled
Public Property Get EnabledProd() As Boolean
Attribute EnabledProd.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    EnabledProd = TxtProd.Enabled
End Property

Public Property Let EnabledProd(ByVal New_EnabledProd As Boolean)
    TxtProd.Enabled() = New_EnabledProd
    PropertyChanged "EnabledProd"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=TxtAge,TxtAge,-1,Enabled
Public Property Get EnabledAge() As Boolean
Attribute EnabledAge.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    EnabledAge = txtAge.Enabled
End Property

Public Property Let EnabledAge(ByVal New_EnabledAge As Boolean)
    txtAge.Enabled() = New_EnabledAge
    PropertyChanged "EnabledAge"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=TxtCuenta,TxtCuenta,-1,Text
Public Property Get Cuenta() As String
Attribute Cuenta.VB_Description = "Devuelve o establece el texto contenido en el control."
    Cuenta = txtCuenta.Text
End Property

Public Property Let Cuenta(ByVal New_Cuenta As String)
    txtCuenta.Text() = New_Cuenta
    PropertyChanged "Cuenta"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=TxtAge,TxtAge,-1,Text
Public Property Get Age() As String
Attribute Age.VB_Description = "Devuelve o establece el texto contenido en el control."
    Age = txtAge.Text
End Property

Public Property Let Age(ByVal New_Age As String)
    txtAge.Text() = New_Age
    PropertyChanged "Age"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=TxtProd,TxtProd,-1,Text
Public Property Get Prod() As String
Attribute Prod.VB_Description = "Devuelve o establece el texto contenido en el control."
    Prod = TxtProd.Text
End Property

Public Property Let Prod(ByVal New_Prod As String)
    TxtProd.Text() = New_Prod
    PropertyChanged "Prod"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCMAC,txtCMAC,-1,Text
Public Property Get CMAC() As String
Attribute CMAC.VB_Description = "Devuelve o establece el texto contenido en el control."
    CMAC = txtCMAC.Text
End Property

Public Property Let CMAC(ByVal New_CMAC As String)
    txtCMAC.Text() = New_CMAC
    PropertyChanged "CMAC"
End Property


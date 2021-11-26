VERSION 5.00
Begin VB.UserControl ActXCodCta_New 
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2730
   ForeColor       =   &H00FF0000&
   ScaleHeight     =   705
   ScaleWidth      =   2730
   Begin VB.Frame lblTexto 
      Caption         =   "Texto"
      ForeColor       =   &H00FF0000&
      Height          =   680
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2680
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
         Left            =   570
         MaxLength       =   2
         TabIndex        =   4
         Top             =   240
         Width           =   345
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
         Left            =   930
         MaxLength       =   3
         TabIndex        =   3
         Top             =   240
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
         Left            =   1380
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1200
      End
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
         Left            =   120
         MaxLength       =   3
         TabIndex        =   1
         Top             =   240
         Width           =   435
      End
   End
End
Attribute VB_Name = "ActXCodCta_New"
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
If Len(Trim(TxtAge)) = 2 Then
    If TxtProd.Enabled And TxtProd.Visible Then
        TxtProd.SetFocus
    End If
End If
End Sub

Private Sub TxtAge_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub TxtAge_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyTab
             If Len(Trim(TxtAge)) = 0 Then
                TxtAge.SetFocus
             End If
        Case vbKeyBack
             lbBakAge = True
             If Len(Trim(TxtAge)) = 0 Then
                If txtCMAC.Enabled Then
                    txtCMAC.SetFocus
                End If
             End If
             Exit Sub
        Case Else
             If Len(Trim(TxtAge)) = 1 Then
                TxtCuenta.SetFocus
             End If
    End Select
    If NumerosEnteros(KeyAscii) = 0 Or Len(Trim(TxtAge)) = 2 Then
       If TxtProd.Enabled Then
            TxtProd.SetFocus
       End If
    End If
End Sub

Private Sub TxtAge_LostFocus()
    If lbBakAge Then
    Else
       psAge = TxtAge.Text
    End If
    lbBakAge = False
End Sub

Private Sub txtCMAC_Change()
If Len(Trim(txtCMAC)) = 3 Then
    If TxtAge.Enabled And TxtAge.Visible Then
        TxtAge.SetFocus
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
             If Len(Trim(TxtAge)) = 1 Then
                TxtCuenta.SetFocus
             End If
    End Select
    If NumerosEnteros(KeyAscii) = 0 Or Len(Trim(txtCMAC)) = 3 Then
       If TxtAge.Enabled Then
            TxtAge.SetFocus
       End If
    End If
End Sub

Private Sub txtCuenta_Change()
If Len(Trim(TxtCuenta)) = 10 Then
    RaiseEvent Change
End If
End Sub

Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtcuenta_LostFocus()
    If lbBakCta Then
    Else
       psCuenta = TxtCuenta.Text
    End If
    lbBakCta = False
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn Or vbKeyTab
             If Len(Trim(TxtCuenta)) < 10 Then
                MsgBox "Número de Cuenta Incompleto", vbInformation, "Aviso"
                TxtCuenta = ""
                TxtCuenta.SetFocus
             Else
                RaiseEvent KeyPress(KeyAscii)
             End If
        Case vbKeyBack
             lbBakCta = True
             If Len(Trim(TxtCuenta)) = 0 Then
                If TxtProd.Enabled Then
                   TxtProd.SetFocus
                Else
                   If TxtAge.Enabled Then
                      TxtAge.SetFocus
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
    If Len(Trim(TxtProd)) = 3 And TxtCuenta.Visible And TxtCuenta.Enabled Then
       TxtCuenta.SetFocus
    End If
End Sub

Private Sub TxtProd_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub TxtProd_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyTab
             If Len(Trim(TxtProd)) = 0 Then
                TxtProd.SetFocus
             End If
        Case vbKeyBack
             lbBakPro = True
             If Len(Trim(TxtProd)) = 0 Then
                If TxtAge.Enabled Then
                   TxtAge.SetFocus
                End If
             Else
                Exit Sub
             End If
        Case vbKeyReturn
             If Len(Trim(TxtProd)) = 3 Then
                TxtCuenta.SetFocus
             End If
    End Select

    If NumerosEnteros(KeyAscii) = 0 Or Len(Trim(TxtProd)) = 3 Then
        TxtCuenta.SetFocus
    End If

End Sub

Public Property Get NroCuenta() As String
    NroCuenta = Trim(txtCMAC) & Trim(TxtAge) & Trim(TxtProd) & Trim(TxtCuenta)
End Property

Public Property Let NroCuenta(ByVal vNewValue As String)
    txtCMAC.Text = Mid(vNewValue, 1, 3)
    TxtAge.Text = Mid(vNewValue, 4, 2)
    TxtProd.Text = Mid(vNewValue, 6, 3)
    TxtCuenta.Text = Mid(vNewValue, 9, 10)
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
    TxtAge.Text = ""
    TxtProd.Text = ""
    TxtCuenta = ""
    lbBakAge = False
    lbBakPro = False
    lbBakCta = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lblTexto.Caption = PropBag.ReadProperty("Texto", "Texto")
    txtCMAC.Enabled = PropBag.ReadProperty("EnabledCMAC", Verdadero)
    TxtCuenta.Enabled = PropBag.ReadProperty("EnabledCta", Verdadero)
    TxtProd.Enabled = PropBag.ReadProperty("EnabledProd", Verdadero)
    TxtAge.Enabled = PropBag.ReadProperty("EnabledAge", Verdadero)
    TxtCuenta.Text = PropBag.ReadProperty("Cuenta", "")
    TxtAge.Text = PropBag.ReadProperty("Age", "")
    TxtProd.Text = PropBag.ReadProperty("Prod", "")
    txtCMAC.Text = PropBag.ReadProperty("CMAC", "")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Texto", lblTexto.Caption, "Texto")
    Call PropBag.WriteProperty("EnabledCMAC", txtCMAC.Enabled, Verdadero)
    Call PropBag.WriteProperty("EnabledCta", TxtCuenta.Enabled, Verdadero)
    Call PropBag.WriteProperty("EnabledProd", TxtProd.Enabled, Verdadero)
    Call PropBag.WriteProperty("EnabledAge", TxtAge.Enabled, Verdadero)
    Call PropBag.WriteProperty("Cuenta", TxtCuenta.Text, "")
    Call PropBag.WriteProperty("Age", TxtAge.Text, "")
    Call PropBag.WriteProperty("Prod", TxtProd.Text, "")
    Call PropBag.WriteProperty("CMAC", txtCMAC.Text, "")
End Sub

Public Function GetCuenta() As String
    GetCuenta = Trim(txtCMAC) & Trim(TxtAge) & Trim(TxtProd) & Trim(TxtCuenta)
End Function

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal vNewEnabled As Boolean)
    UserControl.Enabled = vNewEnabled
    PropertyChanged "Enabled"
End Property

Public Sub SetFocusAge()
    If TxtAge.Enabled Then TxtAge.SetFocus
End Sub

Public Sub SetFocusProd()
    If TxtProd.Enabled Then TxtProd.SetFocus
End Sub

Public Sub SetFocusCuenta()
    If TxtCuenta.Enabled Then TxtCuenta.SetFocus
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=lblTexto,lblTexto,-1,Caption
Public Property Get texto() As String
    texto = lblTexto.Caption
End Property

Public Property Let texto(ByVal New_Texto As String)
    lblTexto.Caption() = New_Texto
    PropertyChanged "Texto"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCMAC,txtCMAC,-1,Enabled
Public Property Get EnabledCMAC() As Boolean
    EnabledCMAC = txtCMAC.Enabled
End Property

Public Property Let EnabledCMAC(ByVal New_EnabledCMAC As Boolean)
    txtCMAC.Enabled() = New_EnabledCMAC
    PropertyChanged "EnabledCMAC"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=TxtCuenta,TxtCuenta,-1,Enabled
Public Property Get EnabledCta() As Boolean
    EnabledCta = TxtCuenta.Enabled
End Property

Public Property Let EnabledCta(ByVal New_EnabledCta As Boolean)
    TxtCuenta.Enabled() = New_EnabledCta
    PropertyChanged "EnabledCta"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=TxtProd,TxtProd,-1,Enabled
Public Property Get EnabledProd() As Boolean
    EnabledProd = TxtProd.Enabled
End Property

Public Property Let EnabledProd(ByVal New_EnabledProd As Boolean)
    TxtProd.Enabled() = New_EnabledProd
    PropertyChanged "EnabledProd"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=TxtAge,TxtAge,-1,Enabled
Public Property Get EnabledAge() As Boolean
    EnabledAge = TxtAge.Enabled
End Property

Public Property Let EnabledAge(ByVal New_EnabledAge As Boolean)
    TxtAge.Enabled() = New_EnabledAge
    PropertyChanged "EnabledAge"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=TxtCuenta,TxtCuenta,-1,Text
Public Property Get Cuenta() As String
    Cuenta = TxtCuenta.Text
End Property

Public Property Let Cuenta(ByVal New_Cuenta As String)
    TxtCuenta.Text() = New_Cuenta
    PropertyChanged "Cuenta"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=TxtAge,TxtAge,-1,Text
Public Property Get Age() As String
    Age = TxtAge.Text
End Property

Public Property Let Age(ByVal New_Age As String)
    TxtAge.Text() = New_Age
    PropertyChanged "Age"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=TxtProd,TxtProd,-1,Text
Public Property Get Prod() As String
    Prod = TxtProd.Text
End Property

Public Property Let Prod(ByVal New_Prod As String)
    TxtProd.Text() = New_Prod
    PropertyChanged "Prod"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCMAC,txtCMAC,-1,Text
Public Property Get CMAC() As String
    CMAC = txtCMAC.Text
End Property

Public Property Let CMAC(ByVal New_CMAC As String)
    txtCMAC.Text() = New_CMAC
    PropertyChanged "CMAC"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=lblTexto,lblTexto,-1,Color
Public Property Get BackColor() As OLE_COLOR
    BackColor = lblTexto.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    lblTexto.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=lblTexto,lblTexto,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lblTexto.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblTexto.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property


VERSION 5.00
Begin VB.UserControl ActxCtaCred 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2895
   ScaleHeight     =   600
   ScaleWidth      =   2895
   ToolboxBitmap   =   "ACTXCT~1.ctx":0000
   Begin VB.Frame FraCredito 
      Height          =   525
      Left            =   30
      TabIndex        =   3
      Top             =   -15
      Width           =   2790
      Begin VB.TextBox txtcta 
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
         Index           =   0
         Left            =   1110
         MaxLength       =   2
         TabIndex        =   2
         Top             =   180
         Width           =   345
      End
      Begin VB.TextBox txtcta 
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
         Index           =   1
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   0
         Top             =   180
         Width           =   435
      End
      Begin VB.TextBox txtcta 
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
         Index           =   2
         Left            =   1860
         MaxLength       =   7
         TabIndex        =   1
         Top             =   180
         Width           =   855
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
         Left            =   90
         TabIndex        =   4
         Top             =   195
         Width           =   990
      End
   End
End
Attribute VB_Name = "ActxCtaCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event keypressEnter()
Public Event ERROR()

Public Property Let Enabled(Credito As Boolean)
   FraCredito.Enabled = Credito
End Property
Public Property Get Enabled() As Boolean
  Enabled = FraCredito.Enabled
End Property

Public Property Let EnabledAge(cAge As Boolean)
   txtcta(0).Enabled = cAge
End Property
Public Property Get EnabledAge() As Boolean
   EnabledAge = txtcta(0).Enabled
End Property

Public Property Let EnabledProd(cProd As Boolean)
   txtcta(1).Enabled = cProd
End Property
Public Property Get EnabledProd() As Boolean
   EnabledProd = txtcta(1).Enabled
End Property

Private Sub txtCta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
        Case 1
            If KeyCode = 8 And txtcta(Index).SelStart = 0 Then
                If txtcta(Index - 1).Enabled Then
                    If Len(Trim(txtcta(Index - 1).Text)) > 0 Then
                        txtcta(Index - 1).Text = Mid(txtcta(Index - 1).Text, 1, Len(Trim(txtcta(Index - 1).Text)) - 1)
                        txtcta(Index - 1).SelStart = Len(Trim(txtcta(Index - 1).Text))
                        txtcta(Index - 1).SetFocus
                    End If
                End If
            End If
        Case 2
            If KeyCode = 8 And txtcta(Index).SelStart = 0 Then
                If txtcta(Index - 1).Enabled Then
                    If Len(Trim(txtcta(Index - 1).Text)) > 0 Then
                        txtcta(Index - 1).Text = Mid(txtcta(Index - 1).Text, 1, Len(Trim(txtcta(Index - 1).Text)) - 1)
                        txtcta(Index - 1).SelStart = Len(Trim(txtcta(Index - 1).Text))
                        txtcta(Index - 1).SetFocus
                    End If
                Else
                    If txtcta(Index - 2).Enabled Then
                        If Len(Trim(txtcta(Index - 2).Text)) > 0 Then
                            txtcta(Index - 2).Text = Mid(txtcta(Index - 2).Text, 1, Len(Trim(txtcta(Index - 2).Text)) - 1)
                            txtcta(Index - 2).SelStart = Len(Trim(txtcta(Index - 2).Text))
                            txtcta(Index - 2).SetFocus
                        End If
                    End If
                End If
            End If
            'valido que no se mueva si casilla de moneda no es ni 1 ni 2
            If txtcta(Index).SelStart = 0 Then
                If Mid(txtcta(Index).Text, 1, 1) <> "1" And Mid(txtcta(Index).Text, 1, 1) <> "2" Then
                    KeyCode = 0
                End If
            End If
End Select

End Sub

Private Sub txtCta_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
        Case 0
            If txtcta(Index).SelStart = 2 Then  'flecha derecha
                If txtcta(Index + 1).Enabled = True Then
                    txtcta(Index + 1).SelStart = 0
                    txtcta(Index + 1).SetFocus
                Else
                    txtcta(Index + 2).SelStart = 0
                    txtcta(Index + 2).SetFocus
                End If
            End If
        Case 1
            If KeyCode = 37 And txtcta(Index).SelStart = 0 Then  'flecha izquierda
                If txtcta(Index - 1).Enabled Then
                    txtcta(Index - 1).SelStart = 1
                    txtcta(Index - 1).SetFocus
                End If
            End If
            If txtcta(Index).SelStart = 3 Then
                txtcta(Index + 1).SelStart = 0
                txtcta(Index + 1).SetFocus
            End If
        Case 2
            'con tecla izquierda saltar al anterior
            If KeyCode = 37 And txtcta(Index).SelStart = 0 Then  'flecha izquierda
                If txtcta(Index - 1).Enabled Then
                    txtcta(Index - 1).SelStart = 2
                    txtcta(Index - 1).SetFocus
                Else
                    If txtcta(Index - 2).Enabled Then
                        txtcta(Index - 2).SelStart = 2
                        txtcta(Index - 2).SetFocus
                    End If
                End If
            End If
    End Select
End Sub

Private Sub UserControl_Initialize()
    txtcta(0).Text = Right(Trim(gsCodAge), 2)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 EnabledAge = PropBag.ReadProperty("EnabledAge", "true")
 EnabledProd = PropBag.ReadProperty("EnabledProd", "true")
 Enabled = PropBag.ReadProperty("Enabled", "true")
 Caption = PropBag.ReadProperty("caption", "Caption")
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "enabledage", EnabledAge
    PropBag.WriteProperty "enabledprod", EnabledProd
    PropBag.WriteProperty "enabled", Enabled
    PropBag.WriteProperty "Caption", Caption
End Sub

Function completo() As Boolean
    If Len(txtcta(0).Text) = 2 And Len(txtcta(1).Text) = 3 And Len(txtcta(2).Text) = 7 Then
        completo = True
    Else
        completo = False
    End If
End Function
Public Property Get Text() As String
    Text = Trim(txtcta(0).Text & txtcta(1).Text & txtcta(2).Text)
End Property
Public Property Let Text(cCta As String)
    txtcta(0).Text = Mid(cCta, 1, 2)
    txtcta(1).Text = Mid(cCta, 3, 3)
    txtcta(2).Text = Mid(cCta, 6, 7)
End Property
Public Property Get Estado() As String
'Dim Reg As New ADODB.Recordset
'Dim SQL1 As String
'    If completo = True Then
'        If EsValido(Right(txtcta(2).Text, 6)) Then
'            SQL1 = "SELECT cEstado FROM Credito WHERE cCodCta = '" & Text & "'"
'            Reg.Open SQL1, dbCmact
'            If Not Reg.BOF And Not Reg.EOF Then
'                Estado = Reg.Fields(0)
'            Else
'                Estado = ""
'            End If
'            Reg.Close
'        Else
'            MsgBox "Nro de Cuenta No Valido", vbInformation, "Aviso"
'            Estado = ""
'        End If
'    Else
'        MsgBox "Numero de Cuenta Incompleto", vbInformation, "Aviso"
'        Estado = ""
'    End If
End Property

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
Dim texto As String
    texto = txtcta(Index).Text + Chr$(KeyAscii)
    'KeyAscii = intfNumEnt(KeyAscii)
    'valido que moneda sea solo el numero 1 o 2
    If Index = 2 And txtcta(Index).SelStart = 0 And KeyAscii <> 49 And KeyAscii <> 50 Then
        Beep
        KeyAscii = 0
    End If
    'que no escriba otros numeros si en la casilla de moneda esta un valor diferente de 1 o 2
    If Index = 2 And txtcta(Index).SelStart <> 0 Then
        If Mid(txtcta(Index).Text, 1, 1) <> "1" And Mid(txtcta(Index).Text, 1, 1) <> "2" Then
            Beep
            KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 And Index = 2 Then
        If completo Then
            If EsValido(Right(txtcta(2).Text, 6)) Then
                If ValidaAgencia Then
                    If ValidaTipoCredito Then
                        If ValidaSubTipo Then
                            RaiseEvent keypressEnter
                        Else
                            MsgBox "Sub Tipo de Credito", vbInformation, "Aviso"
                            RaiseEvent ERROR
                        End If
                    Else
                        MsgBox "Tipo de Credito No Valido", vbInformation, "Aviso"
                        RaiseEvent ERROR
                    End If
                Else
                    MsgBox "Numero de Agencia No Valido", vbInformation, "Aviso"
                    RaiseEvent ERROR
                End If
            Else
                MsgBox "Nro de Cuenta No Valido", vbInformation, "Aviso"
                RaiseEvent ERROR
            End If
        Else
            MsgBox "Nro de Cuenta Incompleta", vbInformation, "Aviso"
            RaiseEvent ERROR
        End If
    End If
End Sub

Public Function ValidaAgencia() As Boolean
'Dim Reg As New ADODB.Recordset
'Dim SQL1 As String
'    SQL1 = "select cCodTab from " & gcCentralCom & "TablaCod WHERE cCodTab = '47" & txtcta(0).Text & "'"
'    Reg.Open SQL1, dbCmact, adOpenStatic, adLockBatchOptimistic, adCmdText
'    If Reg.BOF And Reg.EOF Then
'        ValidaAgencia = False
'    Else
'        ValidaAgencia = True
'    End If
'Reg.Close
End Function
Public Function ValidaTipoCredito() As Boolean
'Dim Reg As New ADODB.Recordset
'Dim SQL1 As String
'    SQL1 = "select cCodTab from " & gcCentralCom & "TablaCod WHERE cCodTab = '6" & Left(txtcta(1).Text, 1) & "' and cValor = '" & Left(txtcta(1).Text, 1) & "'"
'    Reg.Open SQL1, dbCmact, adOpenStatic, adLockBatchOptimistic, adCmdText
'    If Reg.BOF And Reg.EOF Then
'        ValidaTipoCredito = False
'    Else
'        ValidaTipoCredito = True
'    End If
'Reg.Close
End Function
Public Function ValidaSubTipo() As Boolean
'Dim Reg As New ADODB.Recordset
'Dim SQL1 As String
'    SQL1 = "select cCodTab from " & gcCentralCom & "TablaCod WHERE cCodTab = '6" & txtcta(1).Text & "'"
'    Reg.Open SQL1, dbCmact, adOpenStatic, adLockBatchOptimistic, adCmdText
'    If Reg.BOF And Reg.EOF Then
'        ValidaSubTipo = False
'    Else
'        ValidaSubTipo = True
'    End If
'Reg.Close
End Function

Public Property Get Caption() As Variant
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal vNewValue As Variant)
    Label1.Caption = vNewValue
End Property

Public Sub Enfoque(ByVal VCaja As Integer)
If VCaja = 1 Or VCaja = 2 Or VCaja = 3 Then
    If txtcta(VCaja - 1).Enabled Then
        txtcta(VCaja - 1).SelStart = 0
        txtcta(VCaja - 1).SetFocus
    End If
End If
End Sub

VERSION 5.00
Begin VB.UserControl ActxCtaCred 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2895
   ScaleHeight     =   600
   ScaleWidth      =   2895
   ToolboxBitmap   =   "ActXCtaCred.ctx":0000
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
   TxtCta(0).Enabled = cAge
End Property
Public Property Get EnabledAge() As Boolean
   EnabledAge = TxtCta(0).Enabled
End Property

Public Property Let EnabledProd(cProd As Boolean)
   TxtCta(1).Enabled = cProd
End Property
Public Property Get EnabledProd() As Boolean
   EnabledProd = TxtCta(1).Enabled
End Property

Private Sub txtCta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
        Case 1
            If KeyCode = 8 And TxtCta(Index).SelStart = 0 Then
                If TxtCta(Index - 1).Enabled Then
                    If Len(Trim(TxtCta(Index - 1).Text)) > 0 Then
                        TxtCta(Index - 1).Text = Mid(TxtCta(Index - 1).Text, 1, Len(Trim(TxtCta(Index - 1).Text)) - 1)
                        TxtCta(Index - 1).SelStart = Len(Trim(TxtCta(Index - 1).Text))
                        TxtCta(Index - 1).SetFocus
                    End If
                End If
            End If
        Case 2
            If KeyCode = 8 And TxtCta(Index).SelStart = 0 Then
                If TxtCta(Index - 1).Enabled Then
                    If Len(Trim(TxtCta(Index - 1).Text)) > 0 Then
                        TxtCta(Index - 1).Text = Mid(TxtCta(Index - 1).Text, 1, Len(Trim(TxtCta(Index - 1).Text)) - 1)
                        TxtCta(Index - 1).SelStart = Len(Trim(TxtCta(Index - 1).Text))
                        TxtCta(Index - 1).SetFocus
                    End If
                Else
                    If TxtCta(Index - 2).Enabled Then
                        If Len(Trim(TxtCta(Index - 2).Text)) > 0 Then
                            TxtCta(Index - 2).Text = Mid(TxtCta(Index - 2).Text, 1, Len(Trim(TxtCta(Index - 2).Text)) - 1)
                            TxtCta(Index - 2).SelStart = Len(Trim(TxtCta(Index - 2).Text))
                            TxtCta(Index - 2).SetFocus
                        End If
                    End If
                End If
            End If
            'valido que no se mueva si casilla de moneda no es ni 1 ni 2
            If TxtCta(Index).SelStart = 0 Then
                If Mid(TxtCta(Index).Text, 1, 1) <> "1" And Mid(TxtCta(Index).Text, 1, 1) <> "2" Then
                    KeyCode = 0
                End If
            End If
End Select

End Sub

Private Sub txtCta_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
        Case 0
            If TxtCta(Index).SelStart = 2 Then  'flecha derecha
                If TxtCta(Index + 1).Enabled = True Then
                    TxtCta(Index + 1).SelStart = 0
                    TxtCta(Index + 1).SetFocus
                Else
                    TxtCta(Index + 2).SelStart = 0
                    TxtCta(Index + 2).SetFocus
                End If
            End If
        Case 1
            If KeyCode = 37 And TxtCta(Index).SelStart = 0 Then  'flecha izquierda
                If TxtCta(Index - 1).Enabled Then
                    TxtCta(Index - 1).SelStart = 1
                    TxtCta(Index - 1).SetFocus
                End If
            End If
            If TxtCta(Index).SelStart = 3 Then
                TxtCta(Index + 1).SelStart = 0
                TxtCta(Index + 1).SetFocus
            End If
        Case 2
            'con tecla izquierda saltar al anterior
            If KeyCode = 37 And TxtCta(Index).SelStart = 0 Then  'flecha izquierda
                If TxtCta(Index - 1).Enabled Then
                    TxtCta(Index - 1).SelStart = 2
                    TxtCta(Index - 1).SetFocus
                Else
                    If TxtCta(Index - 2).Enabled Then
                        TxtCta(Index - 2).SelStart = 2
                        TxtCta(Index - 2).SetFocus
                    End If
                End If
            End If
    End Select
End Sub

Private Sub UserControl_Initialize()
    TxtCta(0).Text = Right(Trim(gsCodAge), 2)
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
    If Len(TxtCta(0).Text) = 2 And Len(TxtCta(1).Text) = 3 And Len(TxtCta(2).Text) = 7 Then
        completo = True
    Else
        completo = False
    End If
End Function
Public Property Get Text() As String
    Text = Trim(TxtCta(0).Text & TxtCta(1).Text & TxtCta(2).Text)
End Property
Public Property Let Text(cCta As String)
    TxtCta(0).Text = Mid(cCta, 1, 2)
    TxtCta(1).Text = Mid(cCta, 3, 3)
    TxtCta(2).Text = Mid(cCta, 6, 7)
End Property
Public Property Get Estado() As String
Dim Reg As New ADODB.Recordset
Dim SQL1 As String
    If completo = True Then
        If EsValido(Right(TxtCta(2).Text, 6)) Then
            SQL1 = "SELECT cEstado FROM Credito WHERE cCodCta = '" & Text & "'"
            Reg.Open SQL1, dbCmact
            If Not Reg.BOF And Not Reg.EOF Then
                Estado = Reg.Fields(0)
            Else
                Estado = ""
            End If
            Reg.Close
        Else
            MsgBox "Nro de Cuenta No Valido", vbInformation, "Aviso"
            Estado = ""
        End If
    Else
        MsgBox "Numero de Cuenta Incompleto", vbInformation, "Aviso"
        Estado = ""
    End If
End Property

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
Dim texto As String
    texto = TxtCta(Index).Text + Chr$(KeyAscii)
    KeyAscii = intfNumEnt(KeyAscii)
    'valido que moneda sea solo el numero 1 o 2
    If Index = 2 And TxtCta(Index).SelStart = 0 And KeyAscii <> 49 And KeyAscii <> 50 Then
        Beep
        KeyAscii = 0
    End If
    'que no escriba otros numeros si en la casilla de moneda esta un valor diferente de 1 o 2
    If Index = 2 And TxtCta(Index).SelStart <> 0 Then
        If Mid(TxtCta(Index).Text, 1, 1) <> "1" And Mid(TxtCta(Index).Text, 1, 1) <> "2" Then
            Beep
            KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 And Index = 2 Then
        If completo Then
            If EsValido(Right(TxtCta(2).Text, 6)) Then
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
Dim Reg As New ADODB.Recordset
Dim SQL1 As String
    SQL1 = "select cCodTab from " & gcCentralCom & "TablaCod WHERE cCodTab = '47" & TxtCta(0).Text & "'"
    Reg.Open SQL1, dbCmact, adOpenStatic, adLockBatchOptimistic, adCmdText
    If Reg.BOF And Reg.EOF Then
        ValidaAgencia = False
    Else
        ValidaAgencia = True
    End If
Reg.Close
End Function
Public Function ValidaTipoCredito() As Boolean
Dim Reg As New ADODB.Recordset
Dim SQL1 As String
    SQL1 = "select cCodTab from " & gcCentralCom & "TablaCod WHERE cCodTab = '6" & Left(TxtCta(1).Text, 1) & "' and cValor = '" & Left(TxtCta(1).Text, 1) & "'"
    Reg.Open SQL1, dbCmact, adOpenStatic, adLockBatchOptimistic, adCmdText
    If Reg.BOF And Reg.EOF Then
        ValidaTipoCredito = False
    Else
        ValidaTipoCredito = True
    End If
Reg.Close
End Function
Public Function ValidaSubTipo() As Boolean
Dim Reg As New ADODB.Recordset
Dim SQL1 As String
    SQL1 = "select cCodTab from " & gcCentralCom & "TablaCod WHERE cCodTab = '6" & TxtCta(1).Text & "'"
    Reg.Open SQL1, dbCmact, adOpenStatic, adLockBatchOptimistic, adCmdText
    If Reg.BOF And Reg.EOF Then
        ValidaSubTipo = False
    Else
        ValidaSubTipo = True
    End If
Reg.Close
End Function

Public Property Get Caption() As Variant
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal vNewValue As Variant)
    Label1.Caption = vNewValue
End Property

Public Sub Enfoque(ByVal VCaja As Integer)
If VCaja = 1 Or VCaja = 2 Or VCaja = 3 Then
    If TxtCta(VCaja - 1).Enabled Then
        TxtCta(VCaja - 1).SelStart = 0
        TxtCta(VCaja - 1).SetFocus
    End If
End If
End Sub

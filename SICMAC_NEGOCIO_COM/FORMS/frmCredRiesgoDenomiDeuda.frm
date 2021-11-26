VERSION 5.00
Begin VB.Form frmCredRiesgoDenomiDeuda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Denominador de Deuda"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   Icon            =   "frmCredRiesgoDenomiDeuda.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   22
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Denominador de Deuda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.Frame Frame5 
         Caption         =   "Ponderado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2295
         Left            =   3480
         TabIndex        =   21
         Top             =   360
         Width           =   1575
         Begin VB.TextBox txtR12 
            Height          =   285
            Left            =   360
            MaxLength       =   2
            TabIndex        =   12
            Text            =   "0"
            Top             =   1900
            Width           =   735
         End
         Begin VB.TextBox txtR9 
            Height          =   285
            Left            =   360
            MaxLength       =   2
            TabIndex        =   9
            Text            =   "0"
            Top             =   1430
            Width           =   735
         End
         Begin VB.TextBox txtR6 
            Height          =   285
            Left            =   360
            MaxLength       =   2
            TabIndex        =   6
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtR3 
            Height          =   285
            Left            =   360
            MaxLength       =   2
            TabIndex        =   3
            Text            =   "0"
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Saldo Bruto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2295
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   3255
         Begin VB.TextBox txtR10 
            Height          =   285
            Left            =   480
            MaxLength       =   6
            TabIndex        =   10
            Text            =   "0"
            Top             =   1900
            Width           =   735
         End
         Begin VB.TextBox txtR11 
            Height          =   285
            Left            =   2160
            MaxLength       =   6
            TabIndex        =   11
            Text            =   "0"
            Top             =   1900
            Width           =   735
         End
         Begin VB.TextBox txtR8 
            Height          =   285
            Left            =   2160
            MaxLength       =   6
            TabIndex        =   8
            Text            =   "0"
            Top             =   1430
            Width           =   735
         End
         Begin VB.TextBox txtR5 
            Height          =   285
            Left            =   2160
            MaxLength       =   6
            TabIndex        =   5
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtR2 
            Height          =   285
            Left            =   2160
            MaxLength       =   6
            TabIndex        =   2
            Text            =   "0"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtR7 
            Height          =   285
            Left            =   480
            MaxLength       =   6
            TabIndex        =   7
            Text            =   "0"
            Top             =   1430
            Width           =   735
         End
         Begin VB.TextBox txtR4 
            Height          =   285
            Left            =   480
            MaxLength       =   6
            TabIndex        =   4
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtR1 
            Height          =   285
            Left            =   480
            MaxLength       =   6
            TabIndex        =   1
            Text            =   "0"
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "S/"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "S/"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   23
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label9 
            Caption         =   "S/"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   20
            Top             =   1455
            Width           =   375
         End
         Begin VB.Label Label10 
            Caption         =   "S/"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   19
            Top             =   990
            Width           =   375
         End
         Begin VB.Label Label11 
            Caption         =   "S/"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   18
            Top             =   555
            Width           =   375
         End
         Begin VB.Label Label12 
            Caption         =   "S/"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1455
            Width           =   375
         End
         Begin VB.Label Label13 
            Caption         =   "S/"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   990
            Width           =   375
         End
         Begin VB.Label Label14 
            Caption         =   "S/"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   15
            Top             =   510
            Width           =   375
         End
      End
   End
End
Attribute VB_Name = "frmCredRiesgoDenomiDeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MatRetornaDenomDeuda As Variant

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Public Sub Inicio(Optional ByRef MatRecibeDenomDeuda As Variant)
If IsArray(MatRecibeDenomDeuda) Then

    If UBound(MatRecibeDenomDeuda) Then
        MatRetornaDenomDeuda = MatRecibeDenomDeuda
        Call CargaDatos
    End If
Else
Set MatRetornaDenomDeuda = Nothing
End If
    Me.Show 1
    
If IsArray(MatRetornaDenomDeuda) Then
    MatRecibeDenomDeuda = MatRetornaDenomDeuda
End If

End Sub

Private Sub Cmdguardar_Click()
Dim i As Integer
ReDim MatRetornaDenomDeuda(4, 4)

For i = 1 To 4
    If i = 1 Then
        MatRetornaDenomDeuda(i, 1) = i
        MatRetornaDenomDeuda(i, 2) = txtR1.Text
        MatRetornaDenomDeuda(i, 3) = txtR2.Text
        MatRetornaDenomDeuda(i, 4) = txtR3.Text
    ElseIf i = 2 Then
        MatRetornaDenomDeuda(i, 1) = i
        MatRetornaDenomDeuda(i, 2) = txtR4.Text
        MatRetornaDenomDeuda(i, 3) = txtR5.Text
        MatRetornaDenomDeuda(i, 4) = txtR6.Text
    ElseIf i = 3 Then
        MatRetornaDenomDeuda(i, 1) = i
        MatRetornaDenomDeuda(i, 2) = txtR7.Text
        MatRetornaDenomDeuda(i, 3) = txtR8.Text
        MatRetornaDenomDeuda(i, 4) = txtR9.Text
    ElseIf i = 4 Then
        MatRetornaDenomDeuda(i, 1) = i
        MatRetornaDenomDeuda(i, 2) = txtR10.Text
        MatRetornaDenomDeuda(i, 3) = txtR11.Text
        MatRetornaDenomDeuda(i, 4) = txtR12.Text
    End If
Next i

    Unload Me
End Sub

Public Sub CargaDatos()
Dim i As Integer

If IsArray(MatRetornaDenomDeuda) Then
    If UBound(MatRetornaDenomDeuda) Then
        For i = 1 To UBound(MatRetornaDenomDeuda)
            If i = 1 Then
                txtR1.Text = MatRetornaDenomDeuda(i, 2)
                txtR2.Text = MatRetornaDenomDeuda(i, 3)
                txtR3.Text = MatRetornaDenomDeuda(i, 4)
            ElseIf i = 2 Then
                txtR4.Text = MatRetornaDenomDeuda(i, 2)
                txtR5.Text = MatRetornaDenomDeuda(i, 3)
                txtR6.Text = MatRetornaDenomDeuda(i, 4)
            ElseIf i = 3 Then
                txtR7.Text = MatRetornaDenomDeuda(i, 2)
                txtR8.Text = MatRetornaDenomDeuda(i, 3)
                txtR9.Text = MatRetornaDenomDeuda(i, 4)
            ElseIf i = 4 Then
                txtR10.Text = MatRetornaDenomDeuda(i, 2)
                txtR11.Text = MatRetornaDenomDeuda(i, 3)
                txtR12.Text = MatRetornaDenomDeuda(i, 4)
            End If
        Next i
    End If
End If
End Sub

Private Sub Form_Load()
CentraForm Me
    fEnfoque txtR1
End Sub

Private Sub txtR1_Change()
    If Not IsNumeric(txtR1.Text) Then
        txtR1.Text = 0
    End If
End Sub

Private Sub txtR1_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtR1, KeyAscii, 15)
    If KeyAscii = 13 Then
        txtR2.SetFocus
        txtR1.Text = Format(txtR1.Text, "#0.00")
        fEnfoque txtR2
    End If
End Sub

Private Sub txtR2_Change()
    If Not IsNumeric(txtR2.Text) Then
        txtR2.Text = 0
    End If
End Sub

Private Sub txtR2_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtR2, KeyAscii, 15)
    If KeyAscii = 13 Then
        txtR3.SetFocus
        txtR2.Text = Format(txtR2.Text, "#0.00")
        fEnfoque txtR3
    End If
End Sub

Private Sub txtR3_Change()
    If Not IsNumeric(txtR3.Text) Then
        txtR3.Text = 0
    End If
End Sub

Private Sub txtR3_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        txtR4.SetFocus
        fEnfoque txtR4
    End If
End Sub

Private Sub txtR4_Change()
    If Not IsNumeric(txtR4.Text) Then
        txtR4.Text = 0
    End If
End Sub

Private Sub txtR4_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtR4, KeyAscii, 15)
    If KeyAscii = 13 Then
        txtR5.SetFocus
        txtR4.Text = Format(txtR4.Text, "#0.00")
        fEnfoque txtR5
    End If
End Sub

Private Sub txtR5_Change()
    If Not IsNumeric(txtR5.Text) Then
        txtR5.Text = 0
    End If
End Sub

Private Sub txtR5_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtR5, KeyAscii, 15)
    If KeyAscii = 13 Then
        txtR6.SetFocus
        txtR5.Text = Format(txtR5.Text, "#0.00")
        fEnfoque txtR6
    End If
End Sub

Private Sub txtR6_Change()
    If Not IsNumeric(txtR6.Text) Then
        txtR6.Text = 0
    End If
End Sub

Private Sub txtR6_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        txtR7.SetFocus
        fEnfoque txtR7
    End If
End Sub

Private Sub txtR7_Change()
    If Not IsNumeric(txtR7.Text) Then
        txtR7.Text = 0
    End If
End Sub

Private Sub txtR7_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtR7, KeyAscii, 15)
    If KeyAscii = 13 Then
        txtR8.SetFocus
        txtR7.Text = Format(txtR7.Text, "#0.00")
        fEnfoque txtR8
    End If
End Sub

Private Sub txtR8_Change()
    If Not IsNumeric(txtR8.Text) Then
        txtR8.Text = 0
    End If
End Sub

Private Sub txtR8_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtR8, KeyAscii, 15)
    If KeyAscii = 13 Then
        txtR9.SetFocus
        txtR8.Text = Format(txtR8.Text, "#0.00")
        fEnfoque txtR9
    End If
End Sub

Private Sub txtR9_Change()
    If Not IsNumeric(txtR9.Text) Then
        txtR9.Text = 0
    End If
End Sub

Private Sub txtR9_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        txtR10.SetFocus
        fEnfoque txtR10
    End If
End Sub

Private Sub txtR10_Change()
    If Not IsNumeric(txtR10.Text) Then
        txtR10.Text = 0
    End If
End Sub

Private Sub txtR10_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtR10, KeyAscii, 15)
    If KeyAscii = 13 Then
        txtR11.SetFocus
        txtR10.Text = Format(txtR10.Text, "#0.00")
        fEnfoque txtR11
    End If
End Sub

Private Sub txtR11_Change()
    If Not IsNumeric(txtR11.Text) Then
        txtR11.Text = 0
    End If
End Sub

Private Sub txtR11_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtR11, KeyAscii, 15)
    If KeyAscii = 13 Then
        txtR12.SetFocus
        txtR11.Text = Format(txtR11.Text, "#0.00")
        fEnfoque txtR12
    End If
End Sub

Private Sub txtR12_Change()
    If Not IsNumeric(txtR12.Text) Then
        txtR12.Text = 0
    End If
End Sub

Private Sub txtR12_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
         CmdGuardar.SetFocus
    End If
End Sub

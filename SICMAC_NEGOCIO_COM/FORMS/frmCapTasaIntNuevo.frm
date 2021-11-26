VERSION 5.00
Begin VB.Form frmCapTasaIntNuevo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   Icon            =   "frmCapTasaIntNuevo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Datos"
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
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtTasa 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtMonFin 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   1
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtMonIni 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   0
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblNro 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   720
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tasa :"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1410
         Width           =   450
      End
      Begin VB.Label Label3 
         Caption         =   "Monto Final :"
         Height          =   435
         Left            =   2520
         TabIndex        =   8
         Top             =   810
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "Monto Inicial :"
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   810
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° :"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   450
         Width           =   270
      End
   End
End
Attribute VB_Name = "frmCapTasaIntNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bEditar As Boolean

Private Sub cmdAceptar_Click()
Dim dbCmact As DConecta
Dim clsTasa As nCapDefinicion
Dim L As ListItem

If bEditar Then
    For Each L In frmCapTasaInt.lstTarifario.ListItems
        If L.Text = lblNro Then
            If frmCapTasaInt.nProducto = gCapPlazoFijo Then
                L.SubItems(1) = Format$(CDbl(txtMonIni), "#,##0")
                L.SubItems(2) = Format$(CDbl(txtMonFin), "#,##0")
            Else
                L.SubItems(1) = Format$(CDbl(txtMonIni), "#,##0.00")
                L.SubItems(2) = Format$(CDbl(txtMonFin), "#,##0.00")
            End If
            L.SubItems(3) = Format$(CDbl(txtTasa), "#,##0.00")
            Set clsTasa = New nCapDefinicion
            clsTasa.ActualizaTasa frmCapTasaInt.nProducto, frmCapTasaInt.nMoneda, frmCapTasaInt.nTipoTasa, CInt(lblNro), CDbl(txtMonIni), CDbl(txtMonFin), CDbl(txtTasa)
            Set clsTasa = Nothing
            Exit For
        End If
    Next
Else
    Set L = frmCapTasaInt.lstTarifario.ListItems.Add(, , lblNro)
    If frmCapTasaInt.nProducto = gCapPlazoFijo Then
        L.SubItems(1) = Format$(CDbl(txtMonIni), "#,##0")
        L.SubItems(2) = Format$(CDbl(txtMonFin), "#,##0")
    Else
        L.SubItems(1) = Format$(CDbl(txtMonIni), "#,##0.00")
        L.SubItems(2) = Format$(CDbl(txtMonFin), "#,##0.00")
    End If
    L.SubItems(3) = Format$(CDbl(txtTasa), "#,##0.00")
    Set clsTasa = New nCapDefinicion
    clsTasa.NuevaTasa frmCapTasaInt.nProducto, frmCapTasaInt.nMoneda, frmCapTasaInt.nTipoTasa, CInt(lblNro), CDbl(txtMonIni), CDbl(txtMonFin), CDbl(txtTasa)
    Set clsTasa = Nothing
End If
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub txtMonFin_GotFocus()
With txtMonFin
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtMonFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtTasa.SetFocus
End If
If frmCapTasaInt.nProducto = gCapPlazoFijo Then
    KeyAscii = NumerosEnteros(KeyAscii)
Else
    KeyAscii = NumerosDecimales(txtMonFin, KeyAscii, 12, 2)
End If
End Sub

Private Sub txtMonIni_GotFocus()
With txtMonIni
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtMonIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMonFin.SetFocus
    Exit Sub
End If
If frmCapTasaInt.nProducto = gCapPlazoFijo Then
    KeyAscii = NumerosEnteros(KeyAscii)
Else
    KeyAscii = NumerosDecimales(txtMonIni, KeyAscii, 12, 2)
End If
End Sub


Private Sub txtTasa_GotFocus()
With txtTasa
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtTasa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAceptar.SetFocus
End If
KeyAscii = NumerosDecimales(txtTasa, KeyAscii, 12, 2)
End Sub

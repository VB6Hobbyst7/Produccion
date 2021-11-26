VERSION 5.00
Begin VB.Form frmAsignaOrdPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Órdenes de Pago"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   DrawStyle       =   3  'Dash-Dot
   FillStyle       =   0  'Solid
   Icon            =   "frmAsignaOrdPago.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2950
      TabIndex        =   17
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4200
      TabIndex        =   18
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame fraGlosa 
      Caption         =   "Glosa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   5295
      Begin VB.TextBox txtGlosa 
         Appearance      =   0  'Flat
         Height          =   1095
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame fraRango 
      Caption         =   "Rango de Ordenes de Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   5295
      Begin VB.TextBox txtRangFinalValue 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   680
         Width           =   3255
      End
      Begin VB.TextBox txtRangInicialValue 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   320
         Width           =   3255
      End
      Begin VB.Label lblFinalField 
         Caption         =   "Final:"
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
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblInicialField 
         Caption         =   "Inicial:"
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
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame fraCantidad 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   5295
      Begin VB.Label lblMillar 
         Caption         =   "(Millar)"
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
         Left            =   3000
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblCantidadValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCantidad"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1440
         TabIndex        =   8
         Top             =   210
         Width           =   1455
      End
      Begin VB.Label lblCantidadField 
         Caption         =   "Cantidad:"
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
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraArea 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   5295
      Begin VB.Label lblAreaField 
         Caption         =   "Area:"
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
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblAreaValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblArea"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1440
         TabIndex        =   5
         Top             =   210
         Width           =   3735
      End
   End
   Begin VB.Frame fraResponsable 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.Label lblResponsableValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblResponsable"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1440
         TabIndex        =   2
         Top             =   210
         Width           =   3735
      End
      Begin VB.Label lblResponsableField 
         Caption         =   "Responsable:"
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
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmAsignaOrdPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fsMatOrdPag() As String
Dim fnItem As Integer
Dim fsResponsable As String
Dim fsArea As String
Dim fnCantidad As Integer
Dim oGen As DLogGeneral
Public Function Inicio(ByVal pnItem As Integer, ByVal psResponsable As String, ByVal psArea As String, ByVal pnCantidad As Integer) As String()
    fnItem = pnItem
    fsResponsable = psResponsable
    fsArea = psArea
    fnCantidad = pnCantidad
    Me.Show 1
    Inicio = fsMatOrdPag
End Function
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Public Function ValidaDatos() As Boolean
    If Len(Trim(txtRangInicialValue.Text)) = 0 Then
        MsgBox "Asegure de ingresar el Rango Inicial. Verifique", vbInformation, "Aviso"
        txtRangInicialValue.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If Len(Trim(txtRangFinalValue.Text)) = 0 Then
        MsgBox "Asegure de ingresar el Rango Final. Verifique", vbInformation, "Aviso"
        txtRangFinalValue.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If Len(Trim(txtGlosa.Text)) = 0 Then
        MsgBox "Asegure de ingresar la Glosa. Verifique", vbInformation, "Aviso"
        txtGlosa.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If Not (CLng(lblCantidadValue.Caption) * 1000) = (CLng(txtRangFinalValue.Text) - CLng(txtRangInicialValue.Text)) + 1 Then
        MsgBox "El rango de Ordenes no coincide con la Cantidad. Verifique", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    If oGen.ObtieneExisteNOrden(Trim(txtRangInicialValue.Text), Trim(txtRangFinalValue.Text)) > 0 Then
        MsgBox "Los Rangos ingresados ya están siendo usados. Verifique", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    ValidaDatos = True
End Function
Private Sub cmdGrabar_Click()
    If Not ValidaDatos Then Exit Sub
    If MsgBox("¿Esta seguro de Continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
    fsMatOrdPag(1, 1) = fnItem
    fsMatOrdPag(2, 1) = lblCantidadValue.Caption
    fsMatOrdPag(3, 1) = txtRangInicialValue.Text
    fsMatOrdPag(4, 1) = txtRangFinalValue.Text
    fsMatOrdPag(5, 1) = txtGlosa.Text
    Unload Me
End Sub
Private Sub Form_Load()
    ReDim fsMatOrdPag(5, 1)
    lblResponsableValue.Caption = fsResponsable
    lblAreaValue.Caption = fsArea
    lblCantidadValue.Caption = fnCantidad
    Set oGen = New DLogGeneral
    AsignaNOrdenes
End Sub
Private Sub AsignaNOrdenes()
    Dim nNumOrd As Long
    nNumOrd = oGen.ObtieneMaxOrdPago()
    txtRangInicialValue.Text = nNumOrd
End Sub
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdGrabar.SetFocus
    End If
End Sub
Private Sub txtRangFinalValue_GotFocus()
    fEnfoque txtRangFinalValue
End Sub
Private Sub txtRangFinalValue_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        Me.txtGlosa.SetFocus
    End If
End Sub
Private Sub txtRangFinalValue_LostFocus()
    If Not Len(Trim(txtRangFinalValue.Text)) = 0 Then
        If Mid(txtRangFinalValue.Text, 1, 1) = "," Then
            txtRangFinalValue.Text = Mid(txtRangFinalValue.Text, 2, Len(txtRangFinalValue.Text) - 1)
        End If
        txtRangFinalValue.Text = Format(txtRangFinalValue.Text, "#,#0")
    End If
End Sub
Private Sub txtRangInicialValue_GotFocus()
    fEnfoque txtRangInicialValue
End Sub
Private Sub txtRangInicialValue_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        Me.txtRangFinalValue.SetFocus
    End If
End Sub
Private Sub txtRangInicialValue_LostFocus()
    If Not Len(Trim(txtRangInicialValue.Text)) = 0 Then
        If Mid(txtRangInicialValue.Text, 1, 1) = "," Then
            txtRangInicialValue.Text = Mid(txtRangInicialValue.Text, 2, Len(txtRangInicialValue.Text) - 1)
        End If
        txtRangInicialValue.Text = Format(txtRangInicialValue.Text, "#,#0")
    End If
End Sub

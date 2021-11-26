VERSION 5.00
Begin VB.Form frmCGSimuladorPlanillaEncaje 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simulador de Planilla de Encaje"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   Icon            =   "frmCGSimladorPlanillaEncaje.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFecha 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   2145
      TabIndex        =   0
      Top             =   210
      Width           =   315
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   354
      Left            =   5925
      TabIndex        =   9
      Top             =   4395
      Width           =   1215
   End
   Begin VB.TextBox txtFecha 
      Height          =   315
      Left            =   855
      MaxLength       =   10
      TabIndex        =   37
      Top             =   180
      Width           =   1635
   End
   Begin VB.Frame FrameDolares 
      Caption         =   "Dolares"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   10
      Top             =   570
      Width           =   7440
      Begin VB.TextBox txtTME 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   5670
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   405
         Width           =   1620
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   1500
         Width           =   1665
      End
      Begin VB.TextBox txtPlazoFijo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   2025
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   1125
         Width           =   1650
      End
      Begin VB.TextBox txtAhorros 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   2025
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   765
         Width           =   1650
      End
      Begin VB.TextBox txtObligaciones 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   405
         Width           =   1635
      End
      Begin VB.TextBox txtToseBasico 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5655
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   765
         Width           =   1620
      End
      Begin VB.TextBox txtTasaEBasico 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5655
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   1140
         Width           =   1620
      End
      Begin VB.TextBox txtTasaEMarginal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5655
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   1500
         Width           =   1620
      End
      Begin VB.Frame Frame4 
         Caption         =   "Calculos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   2025
         TabIndex        =   11
         Top             =   1905
         Width           =   5325
         Begin VB.TextBox txtEncaje 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   3570
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "0.00"
            Top             =   1335
            Width           =   1635
         End
         Begin VB.TextBox txtTotal_ToseBase 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "0.00"
            Top             =   240
            Width           =   1620
         End
         Begin VB.TextBox txtToseBaseXEBasico 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3585
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "0.00"
            Top             =   615
            Width           =   1620
         End
         Begin VB.TextBox txtTotal_ToseBaseXEMarginal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3585
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "0.00"
            Top             =   975
            Width           =   1620
         End
         Begin VB.Label Label4 
            Caption         =   "Encaje"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   2850
            TabIndex        =   19
            Top             =   1425
            Width           =   900
         End
         Begin VB.Label Label4 
            Caption         =   "Total - Tose Base"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   240
            Index           =   16
            Left            =   1935
            TabIndex        =   18
            Top             =   345
            Width           =   1725
         End
         Begin VB.Label Label4 
            Caption         =   "Tose Base x E. Basico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   270
            Index           =   17
            Left            =   1545
            TabIndex        =   17
            Top             =   690
            Width           =   1980
         End
         Begin VB.Label Label4 
            Caption         =   "Total - Tose Base x (E. Marginal)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   270
            Index           =   18
            Left            =   660
            TabIndex        =   16
            Top             =   1050
            Width           =   2865
         End
      End
      Begin VB.CommandButton cmdCalcular 
         Caption         =   "&Calcular"
         Height          =   354
         Index           =   1
         Left            =   150
         TabIndex        =   8
         Top             =   3270
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Total"
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
         Index           =   9
         Left            =   945
         TabIndex        =   32
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Plazo Fijo"
         Height          =   195
         Index           =   10
         Left            =   960
         TabIndex        =   31
         Top             =   1215
         Width           =   1020
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Ahorros"
         Height          =   195
         Index           =   11
         Left            =   990
         TabIndex        =   30
         Top             =   855
         Width           =   960
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Obligaciones Inmediatas"
         Height          =   255
         Index           =   12
         Left            =   195
         TabIndex        =   29
         Top             =   465
         Width           =   1785
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Tasa Minimo Encaje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   8
         Left            =   3855
         TabIndex        =   28
         Top             =   495
         Width           =   1725
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Tose Base"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Index           =   13
         Left            =   3885
         TabIndex        =   27
         Top             =   870
         Width           =   1725
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Tasa de E. Basico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   270
         Index           =   14
         Left            =   3870
         TabIndex        =   26
         Top             =   1215
         Width           =   1725
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Tasa de E. Marginal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   270
         Index           =   15
         Left            =   3870
         TabIndex        =   25
         Top             =   1575
         Width           =   1725
      End
   End
   Begin VB.Frame FrameSoles 
      Caption         =   "Nuevos Soles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3720
      Left            =   120
      TabIndex        =   33
      Top             =   585
      Width           =   7440
      Begin VB.Frame Frame1 
         Caption         =   "Calculos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   3390
         TabIndex        =   39
         Top             =   1905
         Width           =   3915
         Begin VB.TextBox txtEncaje 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   0
            Left            =   2145
            Locked          =   -1  'True
            TabIndex        =   42
            Text            =   "0.00"
            Top             =   1035
            Width           =   1635
         End
         Begin VB.TextBox txtTME 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   0
            Left            =   2145
            Locked          =   -1  'True
            TabIndex        =   41
            Text            =   "0.00"
            Top             =   675
            Width           =   1620
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   0
            Left            =   2145
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "0.00"
            Top             =   300
            Width           =   1635
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Tasa Minimo Encaje"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   240
            Index           =   5
            Left            =   45
            TabIndex        =   45
            Top             =   720
            Width           =   2040
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Encaje"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   1125
            TabIndex        =   44
            Top             =   1125
            Width           =   960
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Total"
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
            Index           =   4
            Left            =   1155
            TabIndex        =   43
            Top             =   420
            Width           =   900
         End
      End
      Begin VB.TextBox txtObligaciones 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1935
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   630
         Width           =   1635
      End
      Begin VB.TextBox txtAhorros 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   990
         Width           =   1650
      End
      Begin VB.TextBox txtPlazoFijo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   1335
         Width           =   1650
      End
      Begin VB.CommandButton cmdCalcular 
         Caption         =   "&Calcular"
         Height          =   354
         Index           =   0
         Left            =   210
         TabIndex        =   4
         Top             =   3225
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Obligaciones Inmediatas"
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   36
         Top             =   705
         Width           =   1740
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Ahorros"
         Height          =   195
         Index           =   2
         Left            =   855
         TabIndex        =   35
         Top             =   1065
         Width           =   960
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Plazo Fijo"
         Height          =   195
         Index           =   3
         Left            =   945
         TabIndex        =   34
         Top             =   1410
         Width           =   915
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
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
      Index           =   0
      Left            =   255
      TabIndex        =   38
      Top             =   255
      Width           =   540
   End
End
Attribute VB_Name = "frmCGSimuladorPlanillaEncaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gOpe As String

Public Sub Inicio(vOpe As String)
    gOpe = vOpe
    Me.Show 1
End Sub

Private Sub cmdCalcular_Click(Index As Integer)
On Error GoTo cmdCalcularErr
    txtObligaciones(Index) = FormatNumber(txtObligaciones(Index), 2)
    txtAhorros(Index) = FormatNumber(txtAhorros(Index), 2)
    txtPlazoFijo(Index) = FormatNumber(txtPlazoFijo(Index), 2)
    txtTotal(Index) = FormatNumber(CDbl(txtObligaciones(Index)) + CDbl(txtAhorros(Index)) + CDbl(txtPlazoFijo(Index)), 2)
    Select Case Index
        Case 0
            txtEncaje(0).Text = FormatNumber(CDbl(txtTotal(0).Text) * CDbl(txtTME(0).Text), 2)
        Case 1
            txtTotal_ToseBase = FormatNumber(CDbl(txtTotal(1)) - CDbl(txtToseBasico), 2)
            txtToseBaseXEBasico = FormatNumber(CDbl(txtToseBasico) * CDbl(txtTasaEBasico), 2)
            txtTotal_ToseBaseXEMarginal.Text = FormatNumber(CDbl(txtTotal_ToseBase) * CDbl(txtTasaEMarginal.Text), 2)
            txtEncaje(1).Text = FormatNumber(CDbl(txtTotal_ToseBaseXEMarginal.Text) + CDbl(txtToseBaseXEBasico.Text), 2)
    End Select
    Exit Sub
cmdCalcularErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdFecha_Click()
On Error GoTo cmdFechaErr
    Dim oCon As DConecta, sSql As String, rs As ADODB.Recordset
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSql = " Select nCodigo, cDescripcion, nvalor  from ParamEncaje A " & _
               " Where A.dFecha = (Select max(dFecha) from ParamEncaje B Where A.nCodigo = B.nCodigo And dFecha < '" & Format(txtFecha, "yyyy/mm/dd") & "') "
        Set rs = oCon.CargaRecordSet(sSql)
        Do While Not rs.EOF
            Select Case rs!nCodigo
                Case 1
                    txtTME(0) = FormatNumber(rs!nValor, 2)
                    txtTME(1) = FormatNumber(rs!nValor, 2)
                Case 2
                    txtToseBasico = FormatNumber(rs!nValor, 2)
                Case 3
                    txtTasaEBasico = FormatNumber(rs!nValor, 2)
                Case 4
                    txtTasaEMarginal = FormatNumber(rs!nValor, 2)
            End Select
            rs.MoveNext
        Loop
        oCon.CierraConexion
'        Select Case gOpe
'            Case "461416"
'                txtObligaciones(0).SetFocus
'            Case "462416"
'                txtObligaciones(1).SetFocus
'        End Select
    End If
    Exit Sub
cmdFechaErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
    txtFecha.Text = gdFecSis
    cmdFecha_Click
    Select Case gOpe
        Case "461416"
            FrameSoles.Visible = True
            FrameDolares.Visible = False
        Case "462416"
            FrameSoles.Visible = False
            FrameDolares.Visible = True
    End Select
End Sub

Private Sub txtAhorros_GotFocus(Index As Integer)
   SelTexto txtAhorros(Index)
End Sub

Private Sub txtAhorros_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAhorros(Index) = FormatNumber(txtAhorros(Index), 2)
        txtPlazoFijo(Index).SetFocus
        Exit Sub
    End If
    KeyAscii = NumerosDecimales(txtAhorros(Index), KeyAscii, 20, 2)
End Sub

Private Sub txtFecha_GotFocus()
    SelTexto txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
Dim nKeyAscii  As Integer
nKeyAscii = KeyAscii
KeyAscii = DigFecha(txtFecha, KeyAscii)
End Sub

Private Sub txtObligaciones_GotFocus(Index As Integer)
    SelTexto txtObligaciones(Index)
End Sub

Private Sub txtObligaciones_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtObligaciones(Index) = FormatNumber(txtObligaciones(Index), 2)
        txtAhorros(Index).SetFocus
        Exit Sub
    End If
    KeyAscii = NumerosDecimales(txtObligaciones(Index), KeyAscii, 20, 2)
End Sub

Private Sub txtPlazoFijo_GotFocus(Index As Integer)
    SelTexto txtPlazoFijo(Index)
End Sub

Private Sub txtPlazoFijo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPlazoFijo(Index) = FormatNumber(txtPlazoFijo(Index), 2)
        cmdCalcular(Index).SetFocus
        Exit Sub
    End If
    KeyAscii = NumerosDecimales(txtPlazoFijo(Index), KeyAscii, 20, 2)
End Sub

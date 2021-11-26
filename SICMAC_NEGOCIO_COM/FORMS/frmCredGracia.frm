VERSION 5.00
Begin VB.Form frmCredGracia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones de Periodo de Gracia"
   ClientHeight    =   2865
   ClientLeft      =   2235
   ClientTop       =   2835
   ClientWidth     =   7200
   Icon            =   "frmCredGracia.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin SICMACT.FlexEdit FEGracia 
      Height          =   2040
      Left            =   4575
      TabIndex        =   15
      Top             =   195
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   3598
      ScrollBars      =   2
      HighLight       =   1
      EncabezadosNombres=   "No-Monto"
      EncabezadosAnchos=   "400-1200"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-1"
      ListaControles  =   "0-0"
      EncabezadosAlineacion=   "C-R"
      FormatosEdit    =   "0-2"
      TextArray0      =   "No"
      SelectionMode   =   1
      lbEditarFlex    =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   1665
      Left            =   60
      TabIndex        =   7
      Top             =   1050
      Width           =   4425
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3090
         TabIndex        =   18
         Top             =   1170
         Width           =   1140
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Aplicar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   14
         Top             =   1155
         Width           =   1140
      End
      Begin VB.Frame Frame2 
         Height          =   870
         Left            =   105
         TabIndex        =   8
         Top             =   195
         Width           =   4155
         Begin VB.OptionButton OptTipGra 
            Caption         =   "Como Primera Cuota"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   150
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton OptTipGra 
            Caption         =   "Como Ultima Cuota"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   135
            TabIndex        =   12
            Top             =   420
            Width           =   1680
         End
         Begin VB.OptionButton OptTipGra 
            Caption         =   "Exonerar"
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   3120
            TabIndex        =   11
            Top             =   165
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.OptionButton OptTipGra 
            Caption         =   "Configurar"
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   2415
            TabIndex        =   10
            Top             =   420
            Width           =   1035
         End
         Begin VB.OptionButton OptTipGra 
            Caption         =   "Repartir"
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   2415
            TabIndex        =   9
            Top             =   150
            Width           =   990
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Datos"
      Height          =   915
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   4425
      Begin VB.Label lblInteres 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Left            =   3375
         TabIndex        =   6
         Top             =   540
         Width           =   900
      End
      Begin VB.Label lblDias 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   3810
         TabIndex        =   5
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Interes :"
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
         Left            =   2655
         TabIndex        =   4
         Top             =   525
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Dias :"
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
         Left            =   2865
         TabIndex        =   3
         Top             =   255
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Credito :"
         Height          =   195
         Left            =   135
         TabIndex        =   2
         Top             =   285
         Width           =   585
      End
      Begin VB.Label LblCodCta 
         AutoSize        =   -1  'True
         Caption         =   "072011015425"
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
         Left            =   780
         TabIndex        =   1
         Top             =   285
         Width           =   1275
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total :"
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
      Left            =   4605
      TabIndex        =   17
      Top             =   2325
      Width           =   570
   End
   Begin VB.Label LblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   5355
      TabIndex        =   16
      Top             =   2295
      Width           =   1350
   End
End
Attribute VB_Name = "frmCredGracia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nInteres As Double
Dim nDias As Integer
Dim nNrocuotas As Integer
Dim nTipoGracia As Integer
Dim MatGracia() As String

Public Function Inicio(ByVal pnDiasGracia As Integer, ByVal pnMontoInteres As Double, ByVal pnNroCuotas As Integer, ByRef pnTipoGracia As Integer, Optional psCtaCod As String = "", Optional pMatConfig As Variant = "", Optional pbSimulador As Boolean) As Variant 'ARLO20180723 ADD pbSimulador
Dim i As Integer

    lblDias.Caption = Trim(str(pnDiasGracia))
    lblInteres.Caption = Trim(str(pnMontoInteres))
    LblCodCta.Caption = psCtaCod
    nInteres = pnMontoInteres
    nDias = pnDiasGracia
    nNrocuotas = pnNroCuotas
    'Solo para los Tipos Anteriores(27-03)
    'MAVM Comentado Por Info 262-2012
    'If pnTipoGracia > 0 And pnTipoGracia < 6 Then
    '    Me.OptTipGra(pnTipoGracia - 1).value = True
    'End If
        
    If pnTipoGracia <> -1 Then
        Call cmdAplicar_Click
        If pnTipoGracia = 4 Then
            If IsArray(pMatConfig) Then
            
                For i = 0 To UBound(pMatConfig) - 1
                    FEGracia.TextMatrix(i + 1, 1) = pMatConfig(i)
                Next i
            End If
        End If
    End If
    'ARLO20180723 ERS037-2018
    If Not pbSimulador Then
        Me.Show 1
    End If
    'ARLO END
    pnTipoGracia = nTipoGracia
    Inicio = MatGracia
End Function

Private Function SumaGracia() As Double
Dim i As Integer
    SumaGracia = 0
    If Trim(FEGracia.TextMatrix(1, 1)) = "" Then
        SumaGracia = 0
    Else
        For i = 1 To FEGracia.rows - 1
            SumaGracia = SumaGracia + CDbl(FEGracia.TextMatrix(i, 1))
        Next i
    End If
    SumaGracia = CDbl(Format(SumaGracia, "#0.00"))
End Function

Private Sub GeneraGracia(ByVal pOptGracia As Integer)
Dim oCalendario As COMNCredito.NCOMCalendario
Dim MatGracia As Variant
Dim i As Integer
  
    Call LimpiaFlex(FEGracia)
    If pOptGracia <> 4 Then
        FEGracia.lbEditarFlex = False
        Set oCalendario = New COMNCredito.NCOMCalendario
        MatGracia = oCalendario.GeneraGracia(pOptGracia, nInteres, nNrocuotas)
        For i = 0 To UBound(MatGracia) - 1
            FEGracia.AdicionaFila
            FEGracia.TextMatrix(i + 1, 1) = MatGracia(i)
        Next i
        FEGracia.lbEditarFlex = False
    Else
        FEGracia.lbEditarFlex = True
        FEGracia.row = 0
        For i = 0 To nNrocuotas - 1
            FEGracia.AdicionaFila
            FEGracia.TextMatrix(i + 1, 1) = "0.00"
        Next i
        FEGracia.row = 1
        FEGracia.Col = 1
        If FEGracia.Visible And FEGracia.Enabled Then
            FEGracia.SetFocus
        End If
    End If
End Sub
Private Sub cmdAplicar_Click()
Dim i As Integer
Dim Index As Integer
 
    Index = 0
    For i = 0 To 4
        If OptTipGra(i).value Then
            Index = i + 1
            Exit For
        End If
    Next i
    Call GeneraGracia(Index)
    ReDim MatGracia(FEGracia.rows - 1)
    For i = 1 To FEGracia.rows - 1
        MatGracia(i - 1) = FEGracia.TextMatrix(i, 1)
    Next i
    For i = 0 To 4
        If OptTipGra(i).value Then
            nTipoGracia = i + 1
            Exit For
        End If
    Next i
    LblTotal.Caption = Format(SumaGracia, "#0.00")
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub FEGracia_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    LblTotal.Caption = Format(SumaGracia, "#0.00")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
    If FEGracia.rows = 2 And Trim(FEGracia.TextMatrix(1, 1)) = "" Then
        MsgBox "Genere el Tipo de Gracia Pulsando Aplicar", vbInformation, "Aviso"
        cmdAplicar.SetFocus
        Cancel = 1
        Exit Sub
    End If

    If SumaGracia <> CDbl(lblInteres.Caption) And OptTipGra(4).value = False Then
        MsgBox "Monto Total de Gracia Generada no Coincide con el Interes de Gracia", vbInformation, "Aviso"
        Cancel = 1
        Exit Sub
    End If
    
    If OptTipGra(3).value Then
        ReDim MatGracia(FEGracia.rows - 1)
        For i = 1 To FEGracia.rows - 1
            MatGracia(i - 1) = FEGracia.TextMatrix(i, 1)
        Next i
    End If
End Sub

Private Sub OptTipGra_Click(Index As Integer)
    Call LimpiaFlex(FEGracia)
End Sub

VERSION 5.00
Begin VB.Form frmMntIPM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Indice de Precios al por Mayor : "
   ClientHeight    =   4710
   ClientLeft      =   2985
   ClientTop       =   1845
   ClientWidth     =   7665
   Icon            =   "frmMntIPM.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMnt 
      Height          =   4530
      Left            =   5460
      TabIndex        =   13
      Top             =   30
      Width           =   2085
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "A&plicar"
         Height          =   360
         Left            =   345
         TabIndex        =   5
         Top             =   3255
         Width           =   1425
      End
      Begin VB.Frame Frame6 
         Height          =   555
         Left            =   105
         TabIndex        =   18
         Top             =   2595
         Width           =   1875
         Begin VB.ComboBox cboFacAnio 
            Height          =   315
            ItemData        =   "frmMntIPM.frx":030A
            Left            =   720
            List            =   "frmMntIPM.frx":0332
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   158
            Width           =   1080
         End
         Begin VB.TextBox TxtFacAnio 
            Height          =   315
            Left            =   75
            TabIndex        =   3
            Top             =   165
            Width           =   645
         End
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   420
         Left            =   195
         TabIndex        =   11
         Top             =   3930
         Width           =   1695
      End
      Begin VB.CommandButton cmdElim 
         Caption         =   "Eli&minar"
         Height          =   420
         Left            =   195
         TabIndex        =   2
         Top             =   1365
         Width           =   1695
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   420
         Left            =   195
         TabIndex        =   1
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Agregar"
         Height          =   420
         Left            =   195
         TabIndex        =   0
         Top             =   330
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Factor IPM a Fecha :"
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
         Left            =   150
         TabIndex        =   19
         Top             =   2370
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4530
      Left            =   120
      TabIndex        =   12
      Top             =   30
      Width           =   5340
      Begin VB.Frame fraDatos 
         Enabled         =   0   'False
         Height          =   1140
         Left            =   180
         TabIndex        =   14
         Top             =   3300
         Width           =   4965
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Aceptar"
            Height          =   360
            Left            =   3600
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   360
            Left            =   3600
            TabIndex        =   10
            Top             =   660
            Width           =   1215
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1800
            TabIndex        =   8
            Top             =   735
            Width           =   1635
         End
         Begin VB.Frame fraFecha 
            Height          =   525
            Left            =   1230
            TabIndex        =   16
            Top             =   120
            Width           =   2205
            Begin VB.ComboBox cboMes 
               Height          =   315
               ItemData        =   "frmMntIPM.frx":039B
               Left            =   720
               List            =   "frmMntIPM.frx":03C3
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   150
               Width           =   1440
            End
            Begin VB.TextBox txtAño 
               Height          =   300
               Left            =   45
               MaxLength       =   4
               TabIndex        =   6
               Top             =   150
               Width           =   645
            End
         End
         Begin VB.Label lblVar1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Ajuste :"
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
            Left            =   450
            TabIndex        =   17
            Top             =   765
            Width           =   1155
         End
         Begin VB.Label lblFecha 
            AutoSize        =   -1  'True
            Caption         =   "Fecha :"
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
            Left            =   450
            TabIndex        =   15
            Top             =   270
            Width           =   660
         End
      End
      Begin Sicmact.FlexEdit fg 
         Height          =   3075
         Left            =   180
         TabIndex        =   20
         Top             =   240
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   5424
         Rows            =   12
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "-Fecha-Valor de Ajuste-Factor"
         EncabezadosAnchos=   "380-1300-1760-1200"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-R-R"
         FormatosEdit    =   "0-0-0-0"
         SelectionMode   =   1
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   375
         RowHeight0      =   285
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmMntIPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nAccion    As Integer
Dim lbConsulta As Boolean

Dim oIPM As DAjusteCont
Dim rs   As ADODB.Recordset

Public Sub Inicio(pbConsulta As Boolean)
lbConsulta = pbConsulta
Me.Show 0, frmMdiMain
End Sub

Private Sub RefrescaDatos()
   TxtAño.Text = Year(CDate(fg.TextMatrix(fg.Row, 1)))
   cboMes.ListIndex = Month(CDate(fg.TextMatrix(fg.Row, 1))) - 1
   txtValor.Text = fg.TextMatrix(fg.Row, 2)
End Sub
Private Sub ActivaControles(plActiva As Boolean, Optional plNuevo As Boolean = False)
   fraDatos.Enabled = plActiva
   fraMnt.Enabled = Not plActiva
   fraFecha.Enabled = plNuevo
   cmdAceptar.Visible = plActiva
   cmdCancelar.Visible = plActiva
End Sub

Private Sub CargaDatos()
Dim sSql As String
    Set rs = oIPM.CargaIPM()
    Set fg.Recordset = rs
End Sub
Private Sub AsignaFactor()
Dim dFecha  As Date
Dim nPos    As Integer
Dim nRow    As Integer
Dim FA      As Double
Dim oBarra  As New clsProgressBar
   oBarra.ShowForm Me
   oBarra.CaptionSyle = eCap_CaptionPercent
   oBarra.Max = fg.Rows - 1
   dFecha = DateAdd("m", 1, CDate("01/" & Format(cboFacAnio.ListIndex + 1, "00") & "/" & TxtFacAnio)) - 1
   For nPos = 1 To fg.Rows - 1
      oBarra.Progress nPos, "INDICE DE PRECIOS AL POR MAYOR", "", "Procesando día " & fg.TextMatrix(nPos, 1) & "...", vbBlue
      If dFecha >= CDate(fg.TextMatrix(nPos, 1)) Then
          Dim oCont As New NContFunciones
          FA = oCont.FactorAjuste(fg.TextMatrix(nPos, 1), dFecha)
          Set oCont = Nothing
          fg.TextMatrix(nPos, 3) = Format(FA, gsFormatoNumeroView)
          nRow = nPos
      Else
          fg.TextMatrix(nPos, 3) = ""
      End If
   Next
   oBarra.CloseForm Me
   Set oBarra = Nothing
   fg.Row = nRow
End Sub

Private Sub CboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtValor.SetFocus
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim sSql As String
Dim dFecha As Date
Dim R As New ADODB.Recordset
Dim Enc As Boolean
Dim nPos As Long

   dFecha = DateAdd("m", 1, CDate("01/" & Format(cboMes.ListIndex + 1, "00") & "/" & TxtAño)) - 1
   If nAccion = 1 Then
      Enc = False
      Set R = oIPM.CargaIPM(Format(dFecha, gsFormatoFecha))
         If Not R.BOF And Not R.EOF Then
             Enc = True
         Else
             Enc = False
         End If
       R.Close
       If Not Enc Then
          oIPM.InsertaIPM Format(dFecha, gsFormatoFecha), Val(txtValor.Text)
       Else
           MsgBox "Valor de Ajuste a la Fecha ya Existe", vbInformation, "Aviso"
       End If
   Else
       oIPM.ActualizaIPM Format(dFecha, gsFormatoFecha), Val(txtValor.Text)
   End If
   ActivaControles False
   nPos = fg.Row
   CargaDatos
   If nAccion = 1 Then
      fg.Row = fg.Rows - 1
   Else
      fg.Row = nPos
   End If
   fg.SetFocus
End Sub

Private Sub CmdAdd_Click()
   ActivaControles True, True
   nAccion = 1
   TxtAño.Text = Trim(Str(Year(gdFecSis)))
   cboMes.ListIndex = Month(gdFecSis) - 1
   txtValor.Text = "0.00"
   TxtAño.SetFocus
End Sub

Private Sub cmdAplicar_Click()
    Call AsignaFactor
End Sub

Private Sub CmdCancelar_Click()
    ActivaControles False
End Sub

Private Sub cmdeditar_Click()
    ActivaControles True, False
    nAccion = 2
    txtValor.SetFocus
End Sub

Private Sub CmdElim_Click()
Dim sSql As String
    If Len(Trim(fg.TextMatrix(1, 0))) > 0 Then
        If MsgBox("Esta seguro que Desea Eliminar el Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
            oIPM.EliminaIPM Format(fg.TextMatrix(fg.Row, 1), gsFormatoFecha)
            Call CargaDatos
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fg_RowColChange()
RefrescaDatos
End Sub

Private Sub Form_Load()
   CentraForm Me
   frmMdiMain.Enabled = False
   If lbConsulta Then
      CmdAdd.Visible = False
      CmdElim.Visible = False
      CmdEditar.Visible = False
      Me.Caption = Me.Caption & "Consulta"
   Else
      Me.Caption = Me.Caption & "Mantenimiento"
   End If
    ActivaControles False
    Set oIPM = New DAjusteCont
    Call CargaDatos
    If fg.Rows > 1 Then
      RefrescaDatos
    End If
    cboFacAnio.ListIndex = Month(gdFecSis) - 1
    TxtFacAnio.Text = Trim(Str(Year(gdFecSis)))
End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rs
Set oIPM = Nothing
frmMdiMain.Enabled = True
End Sub

Private Sub TxtAño_GotFocus()
   fEnfoque TxtAño
End Sub

Private Sub TxtAño_KeyPress(KeyAscii As Integer)
   KeyAscii = NumerosEnteros(KeyAscii)
   If KeyAscii = 13 Then
       cboMes.SetFocus
   End If
End Sub

Private Sub txtValor_GotFocus()
   fEnfoque txtValor
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtValor, KeyAscii, 16, 6)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub


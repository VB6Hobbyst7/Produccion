VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMantParametros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manteninimiento de Parametros"
   ClientHeight    =   5565
   ClientLeft      =   2100
   ClientTop       =   1905
   ClientWidth     =   7485
   Icon            =   "frmMantParametros.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5520
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7440
      Begin VB.ComboBox CboTipoMod 
         Height          =   315
         Left            =   945
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   180
         Width           =   2295
      End
      Begin VB.PictureBox Picture1 
         Height          =   585
         Left            =   195
         ScaleHeight     =   525
         ScaleWidth      =   7050
         TabIndex        =   12
         Top             =   4845
         Width           =   7110
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "&Nuevo"
            Height          =   400
            Left            =   30
            TabIndex        =   5
            Top             =   75
            Width           =   1100
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   400
            Left            =   2250
            TabIndex        =   7
            Top             =   75
            Width           =   1100
         End
         Begin VB.CommandButton cmdModificar 
            Caption         =   "&Modificar"
            Height          =   400
            Left            =   1140
            TabIndex        =   6
            Top             =   75
            Width           =   1100
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Aceptar"
            Height          =   400
            Left            =   4770
            TabIndex        =   8
            Top             =   60
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.CommandButton cmdSalir 
            Cancel          =   -1  'True
            Caption         =   "&Salir"
            CausesValidation=   0   'False
            Height          =   400
            Left            =   5910
            TabIndex        =   9
            Top             =   60
            Width           =   1100
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            CausesValidation=   0   'False
            Height          =   400
            Left            =   5910
            TabIndex        =   10
            Top             =   60
            Visible         =   0   'False
            Width           =   1100
         End
      End
      Begin MSDataGridLib.DataGrid GrdParametros 
         Height          =   3615
         Left            =   210
         TabIndex        =   1
         Top             =   600
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   17
         RowDividerStyle =   6
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "cParamVar"
            Caption         =   "Variable"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cParamDesc"
            Caption         =   "Descripcion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "nParamValor"
            Caption         =   "Valor"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            ScrollBars      =   2
            BeginProperty Column00 
               Alignment       =   2
               DividerStyle    =   6
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3644.788
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1530.142
            EndProperty
         EndProperty
      End
      Begin VB.TextBox TxtVariable 
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   4350
         Width           =   1455
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   315
         Left            =   1830
         TabIndex        =   3
         Top             =   4350
         Width           =   3660
      End
      Begin VB.TextBox TxtValor 
         Height          =   315
         Left            =   5490
         TabIndex        =   4
         Top             =   4350
         Width           =   1590
      End
      Begin VB.Label Label1 
         Caption         =   "Modulo :"
         Height          =   285
         Left            =   225
         TabIndex        =   13
         Top             =   210
         Width           =   720
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   210
         Top             =   4275
         Width           =   7095
      End
   End
End
Attribute VB_Name = "frmMantParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim R As ADODB.Recordset
Dim CmdEjec As Integer '1: Nuevo 2: Modificar 3:Eliminar
Private Sub LimpiaControles()
    TxtVariable.Text = ""
    TxtDescripcion.Text = ""
    TxtValor.Text = ""
End Sub
Private Sub CargaDatos()
    If Not R.BOF And Not R.EOF Then
        TxtVariable.Text = Trim(R!cParamVar)
        TxtDescripcion.Text = Trim(R!cParamDesc)
        TxtValor.Text = Format(R!nParamValor, "#0.00")
    End If
End Sub
Public Sub InicioCosultar()
    cmdNuevo.Enabled = False
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    Me.Show 1
End Sub

Public Sub InicioActualizar()
    Me.Show 1
End Sub

Private Sub HabilitaIngreso(ByVal pbHabilita As Boolean)
    TxtVariable.Enabled = pbHabilita
    TxtDescripcion.Enabled = pbHabilita
    TxtValor.Enabled = pbHabilita
    GrdParametros.Height = IIf(pbHabilita, 3615, 4170)
    cmdNuevo.Visible = Not pbHabilita
    cmdModificar.Visible = Not pbHabilita
    cmdEliminar.Visible = Not pbHabilita
    cmdSalir.Visible = Not pbHabilita
    GrdParametros.Enabled = Not pbHabilita
    cmdAceptar.Visible = pbHabilita
    cmdCancelar.Visible = pbHabilita
    CboTipoMod.Enabled = Not pbHabilita
End Sub
Private Sub CboTipoMod_Click()
Dim oParam As DParametro

On Error GoTo ERRORCboTipoMod_Click

    Set oParam = New DParametro
    Set R = oParam.RecuperaDatos(CInt(Trim(Right(CboTipoMod.Text, 5))))
    Set GrdParametros.DataSource = R
    GrdParametros.Refresh
    Set oParam = Nothing
    Exit Sub

ERRORCboTipoMod_Click:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdAceptar_Click()
Dim oParam As DParametro
    Set oParam = New DParametro
    If CmdEjec = 1 Then
        Call oParam.NuevoParametro(Trim(TxtVariable.Text), Trim(TxtDescripcion.Text), CDbl(Format(TxtValor.Text, "#0.00")), CInt(Right(CboTipoMod.Text, 2)))
    Else
        Call oParam.ModificarParametro(Trim(TxtVariable.Text), Trim(TxtDescripcion.Text), CDbl(Format(TxtValor.Text, "#0.00")), CInt(Right(CboTipoMod.Text, 2)))
    End If
    Set oParam = Nothing
    Call HabilitaIngreso(False)
    Call CboTipoMod_Click
    R.Find "cParamVar = '" & Trim(TxtVariable.Text) & "'"
    GrdParametros.SetFocus
End Sub

Private Sub cmdCancelar_Click()
    Call HabilitaIngreso(False)
End Sub

Private Sub cmdEliminar_Click()
Dim oParam As DParametro

    If MsgBox("Desea Eliminar el Parametro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set oParam = New DParametro
        Call oParam.EliminarParametro(Trim(TxtVariable.Text), CInt(Right(CboTipoMod.Text, 2)))
        Set oParam = Nothing
        Call CboTipoMod_Click
    End If
    GrdParametros.SetFocus
End Sub

Private Sub cmdModificar_Click()
    Call HabilitaIngreso(True)
    TxtVariable.Enabled = False
    Call LimpiaControles
    Call CargaDatos
    CmdEjec = 2
    TxtDescripcion.SetFocus
End Sub

Private Sub cmdNuevo_Click()
    Call HabilitaIngreso(True)
    Call LimpiaControles
    CmdEjec = 1
    R.Sort = GrdParametros.Columns(ColIndex).DataField & " ASC"
    GrdParametros.Refresh
    R.MoveLast
    TxtVariable.Text = CInt(R!cParamVar) + 1
    TxtVariable.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CmdEjec = -1
    GrdParametros.Height = 4170
    CboTipoMod.AddItem "Captaciones" & Space(50) & "1"
    CboTipoMod.AddItem "Colocaciones" & Space(50) & "2"
    CboTipoMod.ListIndex = 0
End Sub



Private Sub GrdParametros_DblClick()
    If cmdModificar.Enabled = True Then
        Call cmdModificar_Click
    End If
End Sub

Private Sub GrdParametros_HeadClick(ByVal ColIndex As Integer)
    R.Sort = GrdParametros.Columns(ColIndex).DataField & " ASC"
    GrdParametros.Refresh
End Sub

Private Sub GrdParametros_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cmdModificar.Enabled = True Then
        Call cmdModificar_Click
    End If
End Sub

Private Sub TxtDescripcion_GotFocus()
    fEnfoque TxtDescripcion
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        TxtValor.SetFocus
    End If
End Sub

Private Sub TxtValor_GotFocus()
    fEnfoque TxtValor
End Sub

Private Sub TxtValor_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtValor, KeyAscii)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub

Private Sub TxtVariable_GotFocus()
    fEnfoque TxtVariable
End Sub

Private Sub TxtVariable_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        TxtDescripcion.SetFocus
    End If
End Sub

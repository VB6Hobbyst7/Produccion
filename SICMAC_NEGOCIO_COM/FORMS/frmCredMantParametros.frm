VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmCredMantParametros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manteninimiento de Parametros"
   ClientHeight    =   5940
   ClientLeft      =   2100
   ClientTop       =   1905
   ClientWidth     =   7485
   Icon            =   "frmCredMantParametros.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame gbControles 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1040
      Left            =   210
      TabIndex        =   9
      Top             =   3820
      Width           =   7095
      Begin VB.TextBox txComentario 
         Height          =   315
         Left            =   1600
         MaxLength       =   48
         TabIndex        =   13
         Top             =   600
         Width           =   5300
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   315
         Left            =   1600
         MaxLength       =   48
         TabIndex        =   12
         Top             =   240
         Width           =   5300
      End
      Begin VB.TextBox TxtValor 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TxtVariable 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5800
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7440
      Begin VB.PictureBox Picture1 
         Height          =   585
         Left            =   195
         ScaleHeight     =   525
         ScaleWidth      =   7050
         TabIndex        =   8
         Top             =   5000
         Width           =   7110
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "&Nuevo"
            Height          =   400
            Left            =   30
            TabIndex        =   1
            Top             =   75
            Width           =   1100
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   400
            Left            =   2250
            TabIndex        =   3
            Top             =   75
            Width           =   1100
         End
         Begin VB.CommandButton cmdModificar 
            Caption         =   "&Modificar"
            Height          =   400
            Left            =   1140
            TabIndex        =   2
            Top             =   75
            Width           =   1100
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Aceptar"
            Height          =   400
            Left            =   4770
            TabIndex        =   4
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
            TabIndex        =   5
            Top             =   60
            Width           =   1100
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            CausesValidation=   0   'False
            Height          =   400
            Left            =   5910
            TabIndex        =   6
            Top             =   60
            Visible         =   0   'False
            Width           =   1100
         End
      End
      Begin MSDataGridLib.DataGrid GrdParametros 
         Height          =   3600
         Left            =   210
         TabIndex        =   0
         Top             =   225
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   6350
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
            DataField       =   "cParamDesc"
            Caption         =   "Descripción"
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
            DataField       =   "nParamValor"
            Caption         =   "Valor"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cParamcom"
            Caption         =   "Comentario"
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
               DividerStyle    =   6
               ColumnWidth     =   3644.788
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3644.788
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCredMantParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim R As ADODB.Recordset
Dim CmdEjec As Integer '1: Nuevo 2: Modificar 3:Eliminar
Private Function ValidaDatos() As Boolean

    ValidaDatos = True
    If Len(Trim(TxtVariable.Text)) = 0 Then
        MsgBox "Ingrese el Codigo del Parametro", vbInformation, "Aviso"
        ValidaDatos = False
        If TxtVariable.Enabled Then
            TxtVariable.SetFocus
        End If
        Exit Function
    End If
    
    If Len(Trim(TxtDescripcion.Text)) = 0 Then
        MsgBox "Ingrese la Descripcion del Parametro", vbInformation, "Aviso"
        ValidaDatos = False
        TxtDescripcion.SetFocus
        Exit Function
    End If
        
    If Len(Trim(TxtValor.Text)) = 0 Then
        MsgBox "Ingrese el Valor del Parametro", vbInformation, "Aviso"
        ValidaDatos = False
        TxtValor.SetFocus
        Exit Function
    End If
End Function
Private Sub LimpiaControles()
    TxtVariable.Text = ""
    TxtDescripcion.Text = ""
    TxtValor.Text = ""
    
    '20060329
    'limipiamos el txComentario
    Me.txComentario.Text = Empty
    
End Sub
Private Sub CargaDatos()
    If Not R.BOF And Not R.EOF Then
        TxtVariable.Text = R!nParamVar
        TxtDescripcion.Text = Trim(R!cParamDesc)
        TxtValor.Text = Format(R!nParamValor, "#0.00")
        Me.txComentario.Text = Trim(R!cParamCom)
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
    
    '20060329
    'cambiamos la visibilidad del GroupBox
    Me.gbControles.Visible = pbHabilita
    
    TxtVariable.Enabled = pbHabilita
    TxtDescripcion.Enabled = pbHabilita
    TxtValor.Enabled = pbHabilita
    
    '20060329
    'habilitamos el txComentario
    Me.txComentario.Enabled = pbHabilita
    
    GrdParametros.Height = IIf(pbHabilita, 3600, 4700)
    cmdNuevo.Visible = Not pbHabilita
    cmdModificar.Visible = Not pbHabilita
    cmdEliminar.Visible = Not pbHabilita
    cmdSalir.Visible = Not pbHabilita
    GrdParametros.Enabled = Not pbHabilita
    cmdAceptar.Visible = pbHabilita
    cmdCancelar.Visible = pbHabilita
End Sub


Private Sub CmdAceptar_Click()
Dim oParam As COMDCredito.DCOMParametro
    If Not ValidaDatos Then
        Exit Sub
    End If
    Set oParam = New COMDCredito.DCOMParametro
    If CmdEjec = 1 Then
        If oParam.ExisteParametro(Trim(TxtVariable.Text)) Then
            Set oParam = Nothing
            MsgBox "Parametro ya Existe", vbInformation, "Aviso"
            Exit Sub
        End If
        Call oParam.NuevoParametro(Trim(TxtVariable.Text), Trim(TxtDescripcion.Text), CDbl(Format(TxtValor.Text, "#0.00")), Trim(Me.txComentario.Text))
    Else
        Call oParam.ModificarParametro(Trim(TxtVariable.Text), Trim(TxtDescripcion.Text), CDbl(Format(TxtValor.Text, "#0.00")), Trim(Me.txComentario.Text))
    End If
    Set R = oParam.RecuperaDatos
    Set GrdParametros.DataSource = R
    GrdParametros.Refresh
    Set oParam = Nothing
    Call HabilitaIngreso(False)
    R.Find "nParamVar = '" & Trim(TxtVariable.Text) & "'"
    GrdParametros.SetFocus
    
End Sub

Private Sub cmdcancelar_Click()
    Call HabilitaIngreso(False)
End Sub

Private Sub CmdEliminar_Click()
Dim oParam As COMDCredito.DCOMParametro

    If MsgBox("Desea Eliminar el Parametro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set oParam = New COMDCredito.DCOMParametro
        Call oParam.EliminarParametro(Trim(Str(R!nParamVar)))
        Set R = oParam.RecuperaDatos
        Set GrdParametros.DataSource = R
        GrdParametros.Refresh
        Set oParam = Nothing
    End If
    
    GrdParametros.SetFocus
End Sub

Private Sub CmdModificar_Click()
    Call HabilitaIngreso(True)
    TxtVariable.Enabled = False
    Call LimpiaControles
    Call CargaDatos
    CmdEjec = 2
    TxtDescripcion.SetFocus
End Sub

Private Sub CmdNuevo_Click()
    Call HabilitaIngreso(True)
    Call LimpiaControles
    CmdEjec = 1
    '20060330 código obsoleto de la caja
    'R.Sort = GrdParametros.Columns(GrdParametros.Col).DataField & " ASC"
    'R.Sort = GrdParametros.Columns(0).DataField & " ASC"
    
    '20060330
    'nuevo codigo
    R.Sort = "nParamVar ASC"
    GrdParametros.Refresh
    R.MoveLast
    TxtVariable.Text = R!nParamVar + 1
    TxtVariable.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

'20060330
'cambiamos el tamaño del datagrid
Private Sub Form_Load()
Dim oParam As COMDCredito.DCOMParametro

    On Error GoTo ErrorForm_Load
        CmdEjec = -1
        GrdParametros.Height = 4700
        Me.gbControles.Visible = False
        Set oParam = New COMDCredito.DCOMParametro
        Set R = oParam.RecuperaDatos
        Set GrdParametros.DataSource = R
        GrdParametros.Refresh
        Set oParam = Nothing
    Exit Sub
ErrorForm_Load:
        MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub GrdParametros_DblClick()
    If cmdModificar.Enabled = True Then
        Call CmdModificar_Click
    End If
End Sub

Private Sub GrdParametros_HeadClick(ByVal ColIndex As Integer)
    'R.Sort = GrdParametros.Columns(GrdParametros.Col).DataField & " ASC"
    R.Sort = GrdParametros.Columns(ColIndex).DataField & " ASC" 'JUEZ 20160510
    GrdParametros.Refresh
End Sub

Private Sub GrdParametros_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cmdModificar.Enabled = True Then
        Call CmdModificar_Click
    End If
End Sub

'20060329
'Agergado esta línea de código
'para la gestion del campo comentario
Private Sub txComentario_GotFocus()
    fEnfoque Me.txComentario
End Sub

'20060329
'convertimos a mayusculas
Private Sub txComentario_KeyPress(KeyAscii As Integer)

    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        Me.cmdAceptar.SetFocus
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
        Me.txComentario.SetFocus
    End If
End Sub

Private Sub TxtValor_LostFocus()
    If Trim(TxtValor.Text) = "" Then
        TxtValor.Text = "0.00"
    Else
        TxtValor.Text = Format(TxtValor.Text, "#0.00")
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

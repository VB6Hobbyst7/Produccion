VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmRHCurriculumTabla 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   Icon            =   "frmRHCurriculumTabla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7005
      TabIndex        =   6
      Top             =   3495
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   3510
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   3510
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   3510
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   3510
      Width           =   1095
   End
   Begin VB.Frame fraCurr 
      Caption         =   "Datos Curriculum"
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
      Height          =   3450
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   8100
      Begin VB.TextBox txtUltMov 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   5715
         TabIndex        =   11
         Top             =   3075
         Width           =   1905
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   870
         MaxLength       =   50
         TabIndex        =   10
         Top             =   3075
         Width           =   4860
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   105
         TabIndex        =   9
         Top             =   3075
         Width           =   780
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
         Height          =   2865
         Left            =   105
         TabIndex        =   1
         Top             =   195
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   5054
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   3510
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   3510
      Width           =   1095
   End
End
Attribute VB_Name = "frmRHCurriculumTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsCodigo As String
Dim lbEditado  As Boolean
Dim lnTipo As TipoOpe
'ALPA 20090122***********************
Dim objPista As COMManejador.Pista
'************************************

Private Sub cmdCancelar_Click()
    lbEditado = True
    Activa False
    Limpia
    CargaDatos
End Sub

Private Sub cmdEditar_Click()
    If Me.Flex.TextMatrix(Me.Flex.Row, 1) = "" Then Exit Sub
    Activa True
    lbEditado = True
    Flex_EnterCell
    Me.txtDescripcion.SetFocus
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Desea Eliminar los el registro.", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Dim oCur As NActualizaDatosCurriculum
    Set oCur = New NActualizaDatosCurriculum
    
    On Error GoTo ERROR
    'ALPA 20090122*************************************
    glsMovNro = GetMovNro(gsCodUser, gsCodAge)
    '**************************************************
    If Not oCur.TipoUsado(Me.Flex.TextMatrix(Flex.Row, 1)) Then
        oCur.EliminaCurriculumTabla Me.Flex.TextMatrix(Flex.Row, 1)
        'ALPA 20090122*************************************
        gsOpeCod = LogPistaMantenimientoCurriculumTabla
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3"
        '**************************************************
    Else
        MsgBox "No se puede eliminar, porque el tipo de curriculum esta siendo usado.", vbInformation, "Aviso"
    End If
    
    CargaDatos
    
    Exit Sub
ERROR:
    MsgBox "No se puede eliminar." & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdGrabar_Click()
    Dim oCur As NActualizaDatosCurriculum
    Set oCur = New NActualizaDatosCurriculum
    
    If Me.txtDescripcion.Text = "" Then
        MsgBox "Debe ingresar una descripcion.", vbInformation, "Aviso"
        Me.txtDescripcion.SetFocus
        Exit Sub
    End If
    'ALPA 20090122*************************************
    glsMovNro = GetMovNro(gsCodUser, gsCodAge)
    '**************************************************
    If lbEditado Then
        oCur.ModificaCurriculumTabla Me.txtCodigo.Text, Me.txtDescripcion.Text, glsMovNro
        'ALPA 20090122*************************************
        gsOpeCod = LogPistaMantenimientoCurriculumTabla
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "2"
        '**************************************************
    Else
        oCur.AgregaCurriculumTabla Me.txtCodigo.Text, Me.txtDescripcion.Text, glsMovNro
        'ALPA 20090122*************************************
        gsOpeCod = LogPistaMantenimientoCurriculumTabla
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1"
        '**************************************************
    End If
        
    lbEditado = True
    Activa False
    Limpia
    CargaDatos
End Sub

Private Sub cmdImprimir_Click()
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    Dim oCur As NActualizaDatosCurriculum
    Set oCur = New NActualizaDatosCurriculum
    Dim lsCadena As String
    
    lsCadena = oCur.GetReporteCurriculumTabla(gsNomAge, gsEmpresa, gdFecSis)
    
    oPrevio.Show lsCadena, Me.Caption, True, 66
    
    Set oPrevio = Nothing
End Sub

Private Sub cmdNuevo_Click()
    Limpia
    
    If Flex.TextMatrix(1, 1) = "" Then
        Me.txtCodigo.Text = "1"
    Else
        Me.txtCodigo.Text = Trim(Str(CCur(Me.Flex.TextMatrix(Flex.Rows - 1, 1) + 1)))
    End If
    
    lbEditado = False
    
    Activa True
    Me.txtDescripcion.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Activa(pbValor As Boolean)
    Me.txtDescripcion.Enabled = pbValor
    Me.Flex.Enabled = Not pbValor
    Me.cmdSalir.Enabled = Not pbValor
    Me.cmdNuevo.Visible = Not pbValor
    Me.cmdEditar.Visible = Not pbValor
    Me.cmdGrabar.Visible = pbValor
    Me.cmdCancelar.Visible = pbValor
End Sub

Private Sub Flex_EnterCell()
    Me.txtCodigo.Text = Me.Flex.TextMatrix(Flex.Row, 1)
    Me.txtDescripcion.Text = Me.Flex.TextMatrix(Flex.Row, 2)
    Me.txtUltMov.Text = Me.Flex.TextMatrix(Flex.Row, 3)
End Sub

Private Sub Form_Load()
    If lnTipo = gTipoOpeRegistro Then
        cmdNuevo_Click
        Me.cmdEliminar.Enabled = False
        Me.cmdEditar.Enabled = False
        Me.cmdImprimir.Enabled = False
    ElseIf lnTipo = gTipoOpeConsulta Then
        Me.cmdEliminar.Enabled = False
        Me.cmdNuevo.Enabled = False
        Me.cmdEditar.Enabled = False
        Me.cmdImprimir.Enabled = False
    ElseIf lnTipo = gTipoOpeReporte Then
        Me.cmdEliminar.Enabled = False
        Me.cmdNuevo.Enabled = False
        Me.cmdEditar.Enabled = False
    End If
    'ALPA 20090122 ***************************************************************************
    Set objPista = New COMManejador.Pista
    '*****************************************************************************************
    CargaDatos
End Sub

Private Sub Limpia()
    Me.txtCodigo.Text = ""
    Me.txtDescripcion.Text = ""
    Me.txtUltMov.Text = ""
End Sub

Private Sub CargaDatos()
    Dim oCur As DActualizaDatosCurriculum
    Set oCur = New DActualizaDatosCurriculum
    Dim rsCur As ADODB.Recordset
    Set rsCur = New ADODB.Recordset
    
    Set rsCur = oCur.GetCurriculumTabla
    
    If rsCur.EOF And rsCur.BOF Then
        Flex.Rows = 2
        Flex.Cols = 4
        Flex.TextMatrix(0, 1) = "Codigo"
        Flex.TextMatrix(0, 2) = "Descripcion"
        Flex.TextMatrix(0, 3) = "Actualizacion"
    Else
        Set Me.Flex.DataSource = rsCur
    End If
    
    Me.Flex.ColWidth(0) = 0
    Me.Flex.ColWidth(1) = 700
    Me.Flex.ColWidth(2) = 4800
    Me.Flex.ColWidth(3) = 1900
    
    Set oCur = Nothing
    Flex_EnterCell
End Sub

Private Sub txtDescripcion_GotFocus()
    txtDescripcion.SelStart = 0
    txtDescripcion.SelLength = 50
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdGrabar.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Public Sub Ini(pnTipo As TipoOpe, psCaption As String)
    lnTipo = pnTipo
    Caption = psCaption
    Me.Show 1
End Sub

VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRHMerDemTabla 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "frmRHMerDemTabla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMerDem 
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
      Left            =   30
      TabIndex        =   5
      Top             =   30
      Width           =   8100
      Begin Sicmact.TxtBuscar TxtBuscar 
         Height          =   285
         Left            =   4965
         TabIndex        =   12
         Top             =   3075
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         sTitulo         =   ""
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   105
         TabIndex        =   8
         Top             =   3075
         Width           =   780
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   870
         MaxLength       =   50
         TabIndex        =   7
         Top             =   3075
         Width           =   4110
      End
      Begin VB.TextBox txtUltMov 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   6075
         TabIndex        =   6
         Top             =   3075
         Width           =   1935
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
         Height          =   2865
         Left            =   105
         TabIndex        =   9
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
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   3555
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   3555
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   3555
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   3555
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7035
      TabIndex        =   0
      Top             =   3540
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   30
      TabIndex        =   10
      Top             =   3555
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   3555
      Width           =   1095
   End
End
Attribute VB_Name = "frmRHMerDemTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsCodigo As String
Dim lbEditado  As Boolean
Dim lnTipo As TipoOpe

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
    Dim oMer As NMeritosDemeritos
    Set oMer = New NMeritosDemeritos
    
    If Not oMer.TipoUsado(Me.Flex.TextMatrix(Flex.Row, 1)) Then
        oMer.EliminaMerDemTabla Me.Flex.TextMatrix(Flex.Row, 1)
    Else
        MsgBox "No puede borrarlo porque el tipo esta en uso.", vbInformation, "Aviso"
    End If
    CargaDatos
End Sub

Private Sub CmdGrabar_Click()
    Dim oMer As NMeritosDemeritos
    Set oMer = New NMeritosDemeritos
    
    If Me.txtDescripcion.Text = "" Then
        MsgBox "Debe ingresar una descripcion.", vbInformation, "Aviso"
        Me.txtDescripcion.SetFocus
        Exit Sub
    ElseIf Me.TxtBuscar.Text = "" Then
        MsgBox "Debe ingresar un tipo de merito / demerito valido.", vbInformation, "Aviso"
        Me.TxtBuscar.SetFocus
        Exit Sub
    End If
    
    If lbEditado Then
        oMer.ModificaMerDemTabla Me.txtCodigo.Text, Me.txtDescripcion.Text, Me.TxtBuscar.Text, GetMovNro(gsCodUser, gsCodAge)
    Else
        oMer.AgregaMerDemTabla Me.txtCodigo.Text, Me.txtDescripcion.Text, Me.TxtBuscar.Text, GetMovNro(gsCodUser, gsCodAge)
    End If
        
    lbEditado = True
    Activa False
    Limpia
    CargaDatos
End Sub

Private Sub cmdImprimir_Click()
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    Dim oMer As NMeritosDemeritos
    Set oMer = New NMeritosDemeritos
    Dim lsCadena As String
    
    lsCadena = oMer.GetReporteMerDemTabla(gsNomAge, gsEmpresa, gdFecSis)
    
    oPrevio.Show lsCadena, Me.Caption, True, 66
    
    Set oPrevio = Nothing
End Sub

Private Sub cmdNuevo_Click()
    Limpia
    
    If Flex.TextMatrix(1, 1) = "" Then
        Me.txtCodigo.Text = "01"
    Else
        Me.txtCodigo.Text = Format(Trim(Str(CCur(Me.Flex.TextMatrix(Flex.Rows - 1, 1) + 1))), "00")
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
    Me.TxtBuscar.Text = Left(Me.Flex.TextMatrix(Flex.Row, 3), 1)
    Me.txtUltMov.Text = Mid(Me.Flex.TextMatrix(Flex.Row, 3), 4)
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
    
    CargaDatos
End Sub

Private Sub Limpia()
    Me.txtCodigo.Text = ""
    Me.txtDescripcion.Text = ""
    Me.txtUltMov.Text = ""
    Me.TxtBuscar.Text = ""
End Sub

Private Sub CargaDatos()
    Dim oMer As DMeritosDemeritos
    Set oMer = New DMeritosDemeritos
    Dim rsMer As ADODB.Recordset
    Set rsMer = New ADODB.Recordset
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    
    Set rsMer = oMer.GetMerDemTabla
    
    If rsMer.EOF And rsMer.BOF Then
        Flex.Rows = 2
        Flex.Cols = 5
        Flex.TextMatrix(0, 1) = "Codigo"
        Flex.TextMatrix(0, 2) = "Descripcion"
        Flex.TextMatrix(0, 3) = "MER_DEM"
        Flex.TextMatrix(0, 4) = "Actualizacion"
    Else
        Set Flex.DataSource = rsMer
    End If
    
    Me.Flex.ColWidth(0) = 0
    Me.Flex.ColWidth(1) = 700
    Me.Flex.ColWidth(2) = 4000
    Me.Flex.ColWidth(3) = 1200
    Me.Flex.ColWidth(4) = 1900
    
    Me.TxtBuscar.rs = oCon.GetConstante(6035, , , True)
    
    Set oMer = Nothing
End Sub

Private Sub txtBuscar_EmiteDatos()
    Me.txtUltMov.Text = Me.TxtBuscar.psDescripcion
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGrabar.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub TxtBuscar_LostFocus()
    If TxtBuscar.Text <> "" Then
        If Not IsNumeric(TxtBuscar.Text) Then
            TxtBuscar.Text = ""
        End If
    End If
End Sub

Private Sub txtDescripcion_GotFocus()
    txtDescripcion.SelStart = 0
    txtDescripcion.SelLength = 50
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.TxtBuscar.SetFocus
    End If
End Sub

Public Sub Ini(pnTipo As TipoOpe, psCaption As String)
    lnTipo = pnTipo
    Caption = psCaption
    Me.Show 1
End Sub

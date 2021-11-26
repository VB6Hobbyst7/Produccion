VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRHAsistMedPrivada 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualiza Asistencia Medica"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   Icon            =   "frmRHAsistMedPrivivada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMontoDesc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6000
      TabIndex        =   12
      Text            =   "0"
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   11
      Text            =   "0"
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox txtMovNro 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   6840
      TabIndex        =   10
      Top             =   4560
      Width           =   2970
   End
   Begin VB.TextBox txtDescripcion 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      MaxLength       =   50
      TabIndex        =   9
      Top             =   4560
      Width           =   4455
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      TabIndex        =   8
      Top             =   4560
      Width           =   735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   7858
      _Version        =   393216
      FixedCols       =   0
      ForeColorFixed  =   8388608
      SelectionMode   =   1
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8700
      TabIndex        =   7
      Top             =   4905
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   4920
      Width           =   1095
   End
End
Attribute VB_Name = "frmRHAsistMedPrivada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lbEdita As Boolean

Private Sub cmdCancelar_Click()
    Activa False
    Flex_EnterCell
End Sub

Private Sub cmdEditar_Click()
    lbEdita = True
    Activa True
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("¿ Desea Eliminar el Registro ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Dim oAsistencia As NActualizaAsistMedicaPriv
    Set oAsistencia = New NActualizaAsistMedicaPriv
    
    If Not oAsistencia.TipoUsado(Me.Flex.TextMatrix(Me.Flex.Row, 1)) Then
        oAsistencia.EliminaAsisMedPriv Me.Flex.TextMatrix(Me.Flex.Row, 1)
        
        If Flex.Rows = 2 Then
            Flex.TextMatrix(0, 0) = ""
            Flex.TextMatrix(1, 0) = ""
            Flex.TextMatrix(2, 0) = ""
            Flex.TextMatrix(3, 0) = ""
            Flex.TextMatrix(4, 0) = ""
            Flex.TextMatrix(5, 0) = ""
        Else
            Flex.RemoveItem Flex.Row
        End If
    Else
        MsgBox "No se puede Eliminar porque este tipo de asistencia medica esta siendo usado.", vbInformation, "Aviso"
    End If
    
    Set oAsistencia = Nothing
End Sub

Private Sub cmdGrabar_Click()
    Dim oAsistencia As NActualizaAsistMedicaPriv
    Set oAsistencia = New NActualizaAsistMedicaPriv
    
    If Not Valida() Then Exit Sub
     
    If lbEdita Then
        If Not oAsistencia.ModificaAsisMedPriv(Me.txtCodigo.Text, Me.txtDescripcion.Text, Me.txtMonto.Text, Me.txtMontoDesc.Text, Me.txtMovNro.Text) Then
            MsgBox "Error de Grabacion", vbInformation, "Aviso"
        Else
            Me.Flex.TextMatrix(Me.Flex.Row, 1) = Me.txtCodigo.Text
            Me.Flex.TextMatrix(Me.Flex.Row, 2) = Me.txtDescripcion.Text
            Me.Flex.TextMatrix(Me.Flex.Row, 3) = Me.txtMonto.Text
            Me.Flex.TextMatrix(Me.Flex.Row, 4) = Me.txtMontoDesc.Text
            Me.Flex.TextMatrix(Me.Flex.Row, 5) = Me.txtMovNro.Text
        End If
    Else
        If Not oAsistencia.AgregaAsisMedPriv(Me.txtCodigo.Text, Me.txtDescripcion.Text, Me.txtMonto.Text, Me.txtMontoDesc.Text, Me.txtMovNro.Text) Then
            MsgBox "Error de Grabacion", vbInformation, "Aviso"
        Else
            Flex.Rows = Flex.Rows + 1
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 1) = Me.txtCodigo.Text
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 2) = Me.txtDescripcion.Text
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 3) = Me.txtMonto.Text
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 4) = Me.txtMontoDesc.Text
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 5) = Me.txtMovNro.Text
        End If
        
    End If
    
    Set oAsistencia = Nothing
    lbEdita = False
    Activa False
End Sub


Private Sub cmdImprimir_Click()
    Dim oAsistencia As NActualizaAsistMedicaPriv
    Set oAsistencia = New NActualizaAsistMedicaPriv
    Dim lsCadena As String
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    
    lsCadena = oAsistencia.GetReporte(gsNomAge, gsEmpresa, gdFecSis)
    
    oPrevio.Show lsCadena, "Asitencia Medica Privada", True, 66
    Set oPrevio = Nothing
End Sub

Private Sub CmdNuevo_Click()
    lbEdita = False
    Me.txtCodigo.Text = Format(CCur(Flex.TextMatrix(Flex.Rows - 1, 1)) + 1, "00")
    Me.txtDescripcion.Text = ""
    Me.txtMonto.Text = ""
    Me.txtMovNro.Text = GetMovNro(gsCodUser, gsCodAge)
    Activa True
    Me.txtDescripcion.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Flex_DblClick()
    cmdEditar_Click
End Sub

Private Sub Flex_EnterCell()
    Me.txtCodigo.Text = Me.Flex.TextMatrix(Me.Flex.Row, 1)
    Me.txtDescripcion.Text = Me.Flex.TextMatrix(Me.Flex.Row, 2)
    Me.txtMonto.Text = Me.Flex.TextMatrix(Me.Flex.Row, 3)
    Me.txtMontoDesc.Text = Me.Flex.TextMatrix(Me.Flex.Row, 4)
    Me.txtMovNro.Text = Me.Flex.TextMatrix(Me.Flex.Row, 5)
End Sub

Private Sub Form_Load()
    Dim oAsisMedica As DActualizaAsistMedicaPrivada
    Set oAsisMedica = New DActualizaAsistMedicaPrivada
    
    Set Flex.DataSource = oAsisMedica.GetAsisMedPriv()

    Flex.ColWidth(0) = 1
    Flex.ColWidth(1) = 700
    Flex.ColWidth(2) = 4450
    Flex.ColWidth(3) = 840
    Flex.ColWidth(4) = 840
    Flex.ColWidth(5) = 2500
    
    Flex.ColAlignment(1) = 7
    Flex.ColAlignment(3) = 7
    Flex.ColAlignment(4) = 7
    
    Set oAsisMedica = Nothing
End Sub

Private Sub Activa(pbvalor As Boolean)
    Me.txtDescripcion.Enabled = pbvalor
    Me.txtMonto.Enabled = pbvalor
    Me.txtMontoDesc.Enabled = pbvalor
    Me.Flex.Enabled = Not pbvalor
    Me.cmdNuevo.Visible = Not pbvalor
    Me.cmdEditar.Visible = Not pbvalor
    Me.cmdGrabar.Visible = pbvalor
    Me.cmdCancelar.Visible = pbvalor
End Sub

Private Sub txtDescripcion_GotFocus()
    txtDescripcion.SelStart = 0
    txtDescripcion.SelLength = 50
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMonto.SetFocus
    End If
End Sub

Private Sub txtMonto_GotFocus()
    txtMonto.SelStart = 0
    txtMonto.SelLength = 10
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtMontoDesc.SetFocus
    Else
        NumerosDecimales txtMonto, KeyAscii
    End If
End Sub

Private Function Valida() As Boolean
    If Trim(Me.txtDescripcion) = "" Then
        MsgBox "Debe Ingresar una descripción.", vbInformation, "Aviso"
        Valida = False
        txtDescripcion.SetFocus
    ElseIf Trim(Me.txtMonto.Text) = "" Then
        MsgBox "Debe Ingresar un monto valido para el Valor.", vbInformation, "Aviso"
        Valida = False
        txtMonto.SetFocus
    ElseIf Trim(Me.txtMontoDesc.Text) = "" Then
        MsgBox "Debe Ingresar un monto valido para el descuento.", vbInformation, "Aviso"
        Valida = False
        txtMontoDesc.SetFocus
    Else
        Valida = True
    End If
End Function

Public Sub Ini(psCaption As String)
    Caption = psCaption
    Me.Show 1
End Sub

Private Sub txtMontoDesc_GotFocus()
    txtMontoDesc.SelStart = 0
    txtMontoDesc.SelLength = 50
End Sub

Private Sub txtMontoDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.cmdGrabar.Visible Then
            Me.cmdGrabar.SetFocus
        End If
    Else
        NumerosDecimales txtMontoDesc, KeyAscii
    End If
End Sub

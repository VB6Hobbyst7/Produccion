VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRHAFP 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   Icon            =   "frmRHAFP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAbreviatura 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      MaxLength       =   3
      TabIndex        =   12
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtCodPers 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      MaxLength       =   50
      TabIndex        =   11
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtUltMov 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   5520
      MaxLength       =   50
      TabIndex        =   10
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   2865
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtDescripcion 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   2
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox txtMontoPrima 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      MaxLength       =   8
      TabIndex        =   1
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtMontoComVariable 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2520
      Width           =   735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   2415
      Left            =   0
      TabIndex        =   3
      Top             =   15
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4260
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
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   2880
      Width           =   975
   End
End
Attribute VB_Name = "frmRHAFP"
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
    Me.txtAbreviatura.SetFocus
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Desea Eliminar : " & Flex.TextMatrix(Flex.Row, 2), vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Dim oAFP As NActualizaDatosAFP
    Set oAFP = New NActualizaDatosAFP
    
    If Not oAFP.EliminaAFP(Flex.TextMatrix(Flex.Row, 2)) Then
        MsgBox "A ocurrido un error.", vbInformation, "Aviso"
    Else
        Flex.RemoveItem Flex.Row
    End If
    
    Set oAFP = Nothing
End Sub

Private Sub CmdGrabar_Click()
    If Not Valida() Then Exit Sub
    
    If lbEdita Then
        Dim oAFP As NActualizaDatosAFP
        Set oAFP = New NActualizaDatosAFP
        Me.txtUltMov.Text = GetMovNro(gsCodUser, gsCodAge)
        If Not oAFP.ModificaAFP(Me.txtCodPers.Text, Me.txtAbreviatura.Text, Me.txtMontoPrima.Text, Me.txtMontoComVariable.Text, Me.txtUltMov.Text) Then
            MsgBox "Error en la Actualizacion.", vbInformation, "Aviso"
        Else
            Me.Flex.TextMatrix(Flex.Row, 3) = Me.txtAbreviatura.Text
            Me.Flex.TextMatrix(Flex.Row, 4) = Me.txtMontoPrima.Text
            Me.Flex.TextMatrix(Flex.Row, 5) = Me.txtMontoComVariable.Text
            Me.Flex.TextMatrix(Flex.Row, 6) = Me.txtUltMov.Text
        End If
        Activa False
        Flex_EnterCell
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim oAFP As NActualizaDatosAFP
    Dim oPrevio As Previo.clsPrevio
    Dim lsCadena As String
    Set oAFP = New NActualizaDatosAFP
    Set oPrevio = New Previo.clsPrevio
    
    lsCadena = oAFP.ReporteAFP(gsNomAge, gsEmpresa, gdFecSis)
    
    oPrevio.Show lsCadena, "Asitencia Medica Privada", True, 66
    Set oPrevio = Nothing
End Sub

Private Sub cmdSalir_Click()
    Dim s As String
    Dim rs As Recordset
    Unload Me
End Sub

Private Sub Flex_DblClick()
    cmdEditar_Click
End Sub

Private Sub Flex_EnterCell()
    Me.txtCodPers.Text = Me.Flex.TextMatrix(Flex.Row, 1)
    Me.txtDescripcion.Text = Me.Flex.TextMatrix(Flex.Row, 2)
    Me.txtAbreviatura.Text = Me.Flex.TextMatrix(Flex.Row, 3)
    Me.txtMontoPrima.Text = Me.Flex.TextMatrix(Flex.Row, 4)
    Me.txtMontoComVariable.Text = Me.Flex.TextMatrix(Flex.Row, 5)
    Me.txtUltMov.Text = Me.Flex.TextMatrix(Flex.Row, 6)
End Sub

Private Sub Form_Load()
    Dim oAFP As DActualizaDatosAFP
    Set oAFP = New DActualizaDatosAFP
    
    Set Me.Flex.DataSource = oAFP.GetAFPs()
    Set oAFP = Nothing
    
    Flex.ColWidth(0) = 1
    Flex.ColWidth(1) = 1300
    Flex.ColWidth(2) = 2300
    Flex.ColWidth(3) = 500
    Flex.ColWidth(4) = 750
    Flex.ColWidth(5) = 750
    Flex.ColWidth(6) = 2400
    
    Flex.ColAlignment(4) = 7
    Flex.ColAlignment(5) = 7
End Sub

Private Sub Activa(pbValor As Boolean)
    Me.txtMontoPrima.Enabled = pbValor
    Me.txtMontoComVariable.Enabled = pbValor
    Me.txtAbreviatura.Enabled = pbValor
    Me.cmdEditar.Visible = Not pbValor
    Me.cmdGrabar.Enabled = pbValor
    Me.cmdCancelar.Visible = pbValor
    Me.Flex.Enabled = Not pbValor
End Sub

Private Function Valida() As Boolean
    If Not IsNumeric(Me.txtMontoPrima.Text) Then
        MsgBox "Debe ingresar un Monto Valido", vbInformation, "Aviso"
        Valida = False
        txtMontoPrima.SetFocus
    ElseIf Not IsNumeric(Me.txtMontoComVariable.Text) Then
        MsgBox "Debe ingresar un Monto Valido", vbInformation, "Aviso"
        Valida = False
        txtMontoComVariable.SetFocus
    Else
        Valida = True
    End If
End Function

Private Sub txtAbreviatura_GotFocus()
    txtAbreviatura.SelStart = 0
    txtAbreviatura.SelLength = 10
End Sub

Private Sub txtAbreviatura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtMontoPrima.SetFocus
    End If
End Sub

Private Sub txtMontoComVariable_GotFocus()
    txtMontoComVariable.SelStart = 0
    txtMontoComVariable.SelLength = 10
End Sub

Private Sub txtMontoComVariable_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdGrabar.SetFocus
    Else
        NumerosDecimales txtMontoComVariable, KeyAscii
    End If
End Sub

Private Sub txtMontoPrima_GotFocus()
    txtMontoPrima.SelStart = 0
    txtMontoPrima.SelLength = 10
End Sub

Private Sub txtMontoPrima_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtMontoComVariable.SetFocus
    Else
        NumerosDecimales txtMontoPrima, KeyAscii
    End If
End Sub

Public Sub Ini(psCaption As String)
    Caption = psCaption
    Me.Show 1
End Sub

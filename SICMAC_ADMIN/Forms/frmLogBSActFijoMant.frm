VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogBSActFijoMant 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Activo Fijo"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10575
   Icon            =   "frmLogBSActFijoMant.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuscaDescripSig 
      Caption         =   "B Desc Sig"
      Height          =   405
      Left            =   3645
      TabIndex        =   20
      Top             =   6900
      Width           =   1155
   End
   Begin VB.CommandButton cmdBuscaDescrip 
      Caption         =   "Busca Desc"
      Height          =   405
      Left            =   2445
      TabIndex        =   19
      Top             =   6900
      Width           =   1155
   End
   Begin VB.CommandButton cmdBuscaSerieSig 
      Caption         =   "B Serie Sig"
      Height          =   405
      Left            =   1245
      TabIndex        =   18
      Top             =   6900
      Width           =   1155
   End
   Begin VB.CommandButton cmdBuscarSerie 
      Caption         =   "Busca Serie"
      Height          =   405
      Left            =   45
      TabIndex        =   17
      Top             =   6900
      Width           =   1155
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   405
      Left            =   6915
      TabIndex        =   11
      Top             =   6900
      Width           =   1155
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   405
      Left            =   8130
      TabIndex        =   10
      Top             =   6900
      Width           =   1155
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   9360
      TabIndex        =   9
      Top             =   6900
      Width           =   1155
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   5130
      Left            =   45
      TabIndex        =   8
      Top             =   45
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   9049
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.Frame fraBS 
      Caption         =   "Actibo Fijo"
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
      Height          =   1620
      Left            =   0
      TabIndex        =   0
      Top             =   5220
      Width           =   10515
      Begin VB.TextBox txtPatrimonio 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Top             =   555
         Width           =   3420
      End
      Begin VB.TextBox txtSerie 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6960
         TabIndex        =   4
         Top             =   555
         Width           =   3420
      End
      Begin VB.TextBox txtDescripción 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   885
         Width           =   9180
      End
      Begin VB.TextBox txtBSG 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   210
         Width           =   9180
      End
      Begin Sicmact.TxtBuscar txtBS 
         Height          =   300
         Left            =   90
         TabIndex        =   1
         Top             =   210
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
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
      Begin Sicmact.TxtBuscar txtAgencia 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   1230
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
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
      Begin VB.Label lblAgencia 
         Caption         =   "Agencia :"
         Height          =   195
         Left            =   165
         TabIndex        =   15
         Top             =   1290
         Width           =   960
      End
      Begin VB.Label lbdDesc 
         Caption         =   "Descripcion :"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   930
         Width           =   930
      End
      Begin VB.Label lblCodPatrimonial 
         Caption         =   "Cod. Patrim. :"
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblAgeG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2295
         TabIndex        =   7
         Top             =   1245
         Width           =   8085
      End
      Begin VB.Label lblSerie 
         Caption         =   "Serie :"
         Height          =   225
         Left            =   6330
         TabIndex        =   5
         Top             =   585
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   405
      Left            =   6915
      TabIndex        =   16
      Top             =   6900
      Width           =   1155
   End
End
Attribute VB_Name = "frmLogBSActFijoMant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsCadSerie As String
Dim lsCadDesc As String
Dim lnI As Integer


Private Sub cmdBuscaDescrip_Click()
    lsCadDesc = InputBox("Ingrese serie :", "Aviso")
    
    For lnI = 1 To Me.Flex.Rows - 1
        If InStr(1, UCase(Flex.TextMatrix(lnI, 7)), UCase(lsCadDesc)) <> 0 Then
            Me.Flex.TopRow = lnI
            Flex.Row = lnI
            Exit Sub
        End If
    Next lnI
    
    MsgBox "Cadena no encontrada.", vbInformation, "Aviso"
End Sub

Private Sub cmdBuscaDescripSig_Click()
    For lnI = lnI + 1 To Me.Flex.Rows - 1
        If InStr(1, UCase(Flex.TextMatrix(lnI, 7)), UCase(lsCadDesc)) <> 0 Then
            Me.Flex.TopRow = lnI
            Flex.Row = lnI
            Exit Sub
        End If
    Next lnI
    
    MsgBox "Cadena no encontrada.", vbInformation, "Aviso"

End Sub

Private Sub cmdBuscarSerie_Click()
    lsCadSerie = InputBox("Ingrese serie :", "Aviso")
    
    For lnI = 1 To Me.Flex.Rows - 1
        If InStr(1, Flex.TextMatrix(lnI, 3), lsCadSerie) <> 0 Then
            Me.Flex.TopRow = lnI
            Flex.Row = lnI
            Exit Sub
        End If
    Next lnI
    
    MsgBox "Cadena no encontrada.", vbInformation, "Aviso"
End Sub

Private Sub cmdBuscaSerieSig_Click()
    For lnI = lnI + 1 To Me.Flex.Rows - 1
        If InStr(1, Flex.TextMatrix(lnI, 3), lsCadSerie) <> 0 Then
            Me.Flex.TopRow = lnI
            Flex.Row = lnI
            Exit Sub
        End If
    Next lnI
    
    MsgBox "Cadena no encontrada.", vbInformation, "Aviso"
End Sub

Private Sub cmdCancelar_Click()
    Me.fraBS.Enabled = False
    Me.Flex.Enabled = True
    Me.cmdEditar.Enabled = True
    Me.CmdCancelar.Enabled = False
    Me.cmdGrabar.Enabled = False
End Sub

Private Sub cmdEditar_Click()
    Me.fraBS.Enabled = True
    Me.Flex.Enabled = False
    Me.cmdEditar.Enabled = False
    Me.CmdCancelar.Enabled = True
    Me.cmdGrabar.Enabled = True
End Sub

Private Sub cmdGrabar_Click()
    Dim Sql As String
    Dim sqlM As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    If Me.txtBS.Text = "" Then
        MsgBox "Debe Ingresar un codigo de bien valido.", vbInformation, "Aviso"
        Me.txtBS.SetFocus
        Exit Sub
    ElseIf Me.txtSerie.Text = "" Then
        MsgBox "Debe Ingresar una serie de bien valida.", vbInformation, "Aviso"
        Me.txtSerie.SetFocus
        Exit Sub
    ElseIf Me.txtPatrimonio.Text = "" Then
        MsgBox "Debe Ingresar un codigo de patrimonio valido.", vbInformation, "Aviso"
        Me.txtPatrimonio.SetFocus
        Exit Sub
    ElseIf Me.txtDescripción.Text = "" Then
        MsgBox "Debe Ingresar un comentario valido.", vbInformation, "Aviso"
        Me.txtDescripción.SetFocus
        Exit Sub
    ElseIf Me.TxtAgencia.Text = "" Then
        MsgBox "Debe Ingresar una area agencia valida.", vbInformation, "Aviso"
        Me.TxtAgencia.SetFocus
        Exit Sub
    End If
    
    If MsgBox(" Desea Grabar los cambios ? ", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    oCon.AbreConexion
    
    sqlM = " Update MovBSAF" _
         & " Set cSerie = '" & Me.txtSerie.Text & "' , cBSCod = '" & Me.txtBS.Text & "'" _
         & " Where nAnio = " & Me.Flex.TextMatrix(Me.Flex.Row, 0) & " And  nMovNro  = " & Me.Flex.TextMatrix(Me.Flex.Row, 1) & " And cBSCod = '" & Me.Flex.TextMatrix(Me.Flex.Row, 2) & "' And cSerie = '" & Me.Flex.TextMatrix(Me.Flex.Row, 3) & "'"

    
    Sql = " Update BSActivoFijo " _
        & " Set cSerie = '" & Me.txtSerie.Text & "', cBSCod = '" & Me.txtBS.Text & "', cAreCod = '" & Left(Me.TxtAgencia.Text, 3) & "', cAgeCod = '" & Mid(Me.TxtAgencia.Text, 4, 2) & "' , cDescripcion = '" & Me.txtDescripción.Text & "' , cNroPatrimonio = '" & Me.txtPatrimonio.Text & "'" _
        & " Where nAnio = " & Me.Flex.TextMatrix(Me.Flex.Row, 0) & " And  nMovNro  = " & Me.Flex.TextMatrix(Me.Flex.Row, 1) & " And cBSCod = '" & Me.Flex.TextMatrix(Me.Flex.Row, 2) & "' And cSerie = '" & Me.Flex.TextMatrix(Me.Flex.Row, 3) & "'"
    
    
    
    oCon.Ejecutar sqlM
    oCon.Ejecutar Sql
    Carga
    
    cmdCancelar_Click
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Flex_Click()
    Me.txtBS.Text = Me.Flex.TextMatrix(Me.Flex.Row, 2)
    txtBS_EmiteDatos
    Me.txtPatrimonio.Text = Me.Flex.TextMatrix(Me.Flex.Row, 8)
    Me.txtSerie.Text = Me.Flex.TextMatrix(Me.Flex.Row, 3)
    Me.txtDescripción.Text = Me.Flex.TextMatrix(Me.Flex.Row, 7)
    
    Me.TxtAgencia.Text = Me.Flex.TextMatrix(Me.Flex.Row, 5) & Me.Flex.TextMatrix(Me.Flex.Row, 6)
    
    txtAgencia_EmiteDatos
End Sub

Private Sub Form_Load()
    Dim oArea As DActualizaDatosArea
    Set oArea = New DActualizaDatosArea
    Dim oDoc As DOperaciones
    Set oDoc = New DOperaciones
    
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    
    Me.TxtAgencia.rs = oArea.GetAgenciasAreas
    
    Me.txtBS.rs = oALmacen.GetBienesAlmacen(, "" & gnLogBSTpoBienConsumo & "','" & gnLogBSTpoBienFijo & "','" & gnLogBSTpoBienNoDepreciable & "")
    
    Carga
    
    cmdCancelar_Click
End Sub

Private Sub txtAgencia_EmiteDatos()
    Me.lblAgeG.Caption = TxtAgencia.psDescripcion
End Sub

Private Sub txtAgencia_GotFocus()
    TxtAgencia.SelStart = 0
    TxtAgencia.SelLength = 50
End Sub

Private Sub txtBS_EmiteDatos()
    Me.txtBSG.Text = txtBS.psDescripcion
End Sub

Private Sub Carga()
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim Sql As String
    
    Sql = " Select nAnio, nMovNro, BS.cBSCod, cSerie, cBSDescripcion, cAreCod, BAF.cAgeCod, cDescripcion , cNroPatrimonio" _
        & " From BSActivoFijo BAF" _
        & " Inner Join BienesServicios BS ON BAF.cBSCOd = BS.cBSCod" _
        & " Left Join Agencias AGE ON AGE.cAgeCod = BAF.cAgecod" _
        & " Order By BS.cBSCod, cSerie"

    oCon.AbreConexion
    
    Set Me.Flex.DataSource = oCon.CargaRecordSet(Sql)

    Me.Flex.ColWidth(0) = 0
    Me.Flex.ColWidth(1) = 0
    Me.Flex.ColWidth(2) = 1200
    Me.Flex.ColWidth(3) = 3000
    Me.Flex.ColWidth(4) = 3000
    Me.Flex.ColWidth(5) = 1000
    Me.Flex.ColWidth(6) = 1000
    Me.Flex.ColWidth(7) = 5000
    Me.Flex.ColWidth(8) = 3000
End Sub

Private Sub txtBS_GotFocus()
    txtBS.SelStart = 0
    txtBS.SelLength = 50
End Sub

Private Sub txtDescripción_GotFocus()
    txtDescripción.SelStart = 0
    txtDescripción.SelLength = 300
End Sub

Private Sub txtPatrimonio_GotFocus()
    txtPatrimonio.SelStart = 0
    txtPatrimonio.SelLength = 50
End Sub

Private Sub txtSerie_GotFocus()
    txtSerie.SelStart = 0
    txtSerie.SelLength = 50
End Sub

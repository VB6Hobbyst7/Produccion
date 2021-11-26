VERSION 5.00
Begin VB.Form frmGrupoOperacionesContable 
   Caption         =   "Operaciones Contables"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9975
   Icon            =   "frmGrupoOperacionesContable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRegistrarCta 
      Caption         =   "&Registrar"
      Height          =   375
      Left            =   8280
      TabIndex        =   30
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminarCta 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   8280
      TabIndex        =   29
      Top             =   4800
      Width           =   1455
   End
   Begin Sicmact.FlexEdit FEFiltro 
      Height          =   1935
      Left            =   120
      TabIndex        =   28
      Top             =   6480
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3413
      Cols0           =   6
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "-cPersCod-IFTipo-Cuenta-SubCuenta-CtaIfPagare"
      EncabezadosAnchos=   "400-2000-1200-2000-1200-1000"
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
      ColumnasAEditar =   "X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-L-L-C"
      FormatosEdit    =   "0-0-0-0-0-0"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "&Mostrar"
      Height          =   375
      Left            =   8280
      TabIndex        =   26
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdRegistrarFiltro 
      Caption         =   "&Registrar"
      Height          =   375
      Left            =   8280
      TabIndex        =   23
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8280
      TabIndex        =   22
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton CmdEliminarFiltro 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   8280
      TabIndex        =   21
      Top             =   6480
      Width           =   1455
   End
   Begin Sicmact.FlexEdit FEOperacion 
      Height          =   1455
      Left            =   120
      TabIndex        =   20
      Top             =   4800
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2566
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "-Operacion-Orden-Cuenta-DH"
      EncabezadosAnchos=   "400-1200-1200-2500-800"
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
      ColumnasAEditar =   "X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-C-L-L"
      FormatosEdit    =   "0-0-0-0-0"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.Frame Frame10 
         Caption         =   "Cuenta IF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4800
         TabIndex        =   18
         Top             =   1680
         Width           =   4935
         Begin VB.TextBox txtCtaIf 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Tipo IFI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   4575
         Begin VB.ComboBox cboTipoIFI 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Persona"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   9615
         Begin Sicmact.TxtBuscar txtBuscarPersona 
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
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
            TipoBusqueda    =   3
            sTitulo         =   ""
         End
         Begin VB.Label lblNomPersona 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2880
            TabIndex        =   15
            Top             =   240
            Width           =   6615
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Cuenta N"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4800
         TabIndex        =   11
         Top             =   3840
         Width           =   4935
         Begin VB.TextBox txtCuentaN 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Opcion"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   3840
         Width           =   4575
         Begin VB.ComboBox cboOpcion 
            Height          =   315
            ItemData        =   "frmGrupoOperacionesContable.frx":030A
            Left            =   120
            List            =   "frmGrupoOperacionesContable.frx":0311
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tipo de Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4800
         TabIndex        =   6
         Top             =   3120
         Width           =   4935
         Begin VB.ComboBox cboTipoCta 
            Height          =   315
            ItemData        =   "frmGrupoOperacionesContable.frx":0318
            Left            =   120
            List            =   "frmGrupoOperacionesContable.frx":0325
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Orden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   3120
         Width           =   4575
         Begin VB.ComboBox cboOrden 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cuenta Contable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   2400
         Width           =   9615
         Begin Sicmact.TxtBuscar txtSubCta 
            Height          =   375
            Left            =   4800
            TabIndex        =   27
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Sicmact.TxtBuscar txtCuentaContable 
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   661
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
         Begin VB.Label Label1 
            Caption         =   "Sub Cuenta"
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
            Left            =   3600
            TabIndex        =   8
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Operaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9615
         Begin Sicmact.TxtBuscar txtOperacion 
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            sTitulo         =   ""
         End
         Begin VB.Label lblOperacionDes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2880
            TabIndex        =   25
            Top             =   270
            Width           =   6615
         End
      End
   End
End
Attribute VB_Name = "frmGrupoOperacionesContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub cboOrden_Change()
'    Call LLenaFlexEdit
'End Sub
'
'Private Sub cboTipoCta_Change()
'    Call LLenaFlexEdit
'End Sub


Private Sub cmdEliminarCta_Click()
    Dim oOpe As New DOperacion
    Call oOpe.EliminaOpeCtaAdeu(FEOperacion.TextMatrix(FEOperacion.Row, 1), FEOperacion.TextMatrix(FEOperacion.Row, 2), FEOperacion.TextMatrix(FEOperacion.Row, 3), FEOperacion.TextMatrix(FEOperacion.Row, 4))
    MsgBox "Eliminación de cuenta se realizó de manera satisfactoria", vbApplicationModal
    Call LLenaFlexEditOpeCta

End Sub

Private Sub CmdEliminarFiltro_Click()
  Dim oOpe As New DOperacion
    Call oOpe.EliminaCtaIFFiltro(FEFiltro.TextMatrix(FEFiltro.Row, 3), FEFiltro.TextMatrix(FEFiltro.Row, 1), FEFiltro.TextMatrix(FEFiltro.Row, 4), txtCtaIf.Text)
    MsgBox "Eliminación de cuenta se realizó de manera satisfactoria", vbApplicationModal
    Call LLenaFlexEditCtaFiltro
End Sub

Private Sub cmdMostrar_Click()
 Call LLenaFlexEditOpeCta
End Sub

Private Sub cmdRegistrarCta_Click()
Dim cOrden As String
 Dim psMovNro As String
 If Trim(Right(cboOrden.Text, 3)) = "10" Then
    cOrden = "A"
 ElseIf Trim(Right(cboOrden.Text, 3)) = "11" Then
    cOrden = "B"
 Else
    cOrden = Trim(Right(cboOrden.Text, 3))
 End If

Call Validar


If Mid(txtOperacion.Text, 3, 1) = Mid(txtCuentaContable.Text, 3, 1) Then
    Dim oOpe As New DOperacion
    Call oOpe.RegistraOpeCta(txtOperacion.Text, cOrden, txtCuentaContable.Text, Trim(Left(cboTipoCta.Text, 3)), Trim(Left(cboOpcion.Text, 3)), txtCuentaN.Text, psMovNro)
    MsgBox "Registro de cuenta se registró de manera satisfactoria", vbApplicationModal
Else
    MsgBox "Moneda de cuenta contable y de la operación debe ser la misma", vbCritical
End If
Call LLenaFlexEditOpeCta
End Sub

Private Sub cmdRegistrarFiltro_Click()

Call Validar
If Mid(txtOperacion.Text, 3, 1) = Mid(txtCuentaContable.Text, 3, 1) Then
    Dim oOpe As New DOperacion
    Call oOpe.RegistraCtaFiltro(txtCuentaContable.Text, txtBuscarPersona.Text, txtSubCta.Text, Trim(Right(cboTipoIFI.Text, 3)), txtCtaIf.Text)
    MsgBox "Registro de cuenta se registró de manera satisfactoria", vbApplicationModal
Else
    MsgBox "Moneda de cuenta contable y de la operación debe ser la misma", vbCritical
End If
Call LLenaFlexEditCtaFiltro

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Validar()
If Trim(txtOperacion.Text) = "" Then
    MsgBox "Ingresar la operacion", vbCritical
    txtOperacion.SetFocus
End If
If Trim(cboOrden.Text) = "" Then
    MsgBox "Seleccionar la Orden", vbCritical
    cboOrden.SetFocus
End If
If Trim(txtCuentaContable.Text) = "" Then
    MsgBox "Ingresar la Cuenta Contable", vbCritical
    txtOperacion.SetFocus
End If

If Trim(cboTipoCta.Text) = "" Then
    MsgBox "Ingresar el Tipo de Cuenta Contable", vbCritical
    cboTipoCta.SetFocus
End If

If Trim(cboOpcion.Text) = "" Then
    MsgBox "Ingresar el Tipo de Cuenta Contable", vbCritical
    cboOpcion.SetFocus
End If


If Trim(txtBuscarPersona.Text) = "" Then
    MsgBox "Ingresar el Tipo de Cuenta Contable", vbCritical
    txtBuscarPersona.SetFocus
End If

If Trim(cboTipoIFI.Text) = "" Then
    MsgBox "Ingresar el Tipo de Cuenta Contable", vbCritical
    cboTipoIFI.SetFocus
End If

End Sub

Private Sub FEOperacion_DblClick()
    Call LLenaFlexEditCtaFiltro
End Sub

Private Sub Form_Load()
    Dim nOrden As ConstanteCabecera
    Dim oOpe As New DOperacion
    nOrden = 1189
    Call LLenaCombo(cboTipoIFI, gCGTipoIF)
    Call LLenaCombo(cboOrden, nOrden)
    txtCuentaContable.rs = oOpe.CargaCtaCont()
    txtOperacion.rs = oOpe.CargaOperacionxCtaContOperacion()
    cboOpcion.ListIndex = IndiceListaCombo(cboOpcion, 1)
End Sub

Private Sub txtBuscarPersona_EmiteDatos()
    lblNomPersona.Caption = txtBuscarPersona.psDescripcion
    'Call LLenaFlexEditOpeCta
End Sub

Private Sub LLenaCombo(ByRef oCombo As ComboBox, ByVal nConsCod As ConstanteCabecera)
    Dim oConstante As DConstante
    Set oConstante = New DConstante
    
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset

    Set oRs = oConstante.CargaConstante(nConsCod)

    Do While Not oRs.EOF
        oCombo.AddItem oRs!cConsDescripcion & Space(500) & IIf(nConsCod = gCGTipoIF, "0", "") & CStr(oRs!nConsValor)
        oRs.MoveNext
    Loop
    oRs.Close
End Sub

Private Sub txtCuentaContable_EmiteDatos()
    Dim oOpe As New DOperacion
    txtSubCta.rs = oOpe.CargaCtaContSubCta(txtCuentaContable.Text)
End Sub

Private Sub txtOperacion_EmiteDatos()
    lblOperacionDes.Caption = txtOperacion.psDescripcion
    'Call LLenaFlexEditCtaFiltro
End Sub
Private Sub LLenaFlexEditOpeCta()
 Dim oOpe As New DOperacion
 Dim oRs As New ADODB.Recordset
 Dim cOrden As String

 If Trim(Right(cboOrden.Text, 3)) = "10" Then
    cOrden = "A"
 ElseIf Trim(Right(cboOrden.Text, 3)) = "11" Then
    cOrden = "B"
 Else
    cOrden = Trim(Right(cboOrden.Text, 3))
 End If

 
 Set oRs = oOpe.CargaListaOperacionxOpeCta(txtOperacion.Text, Trim(Left(cboTipoCta.Text, 3)), cOrden)
 LimpiaFlex FEOperacion
 Do While Not oRs.EOF
   FEOperacion.AdicionaFila
   FEOperacion.TextMatrix(oRs.Bookmark, 1) = oRs!cOpeCod
   FEOperacion.TextMatrix(oRs.Bookmark, 3) = oRs!cCtaContCod
   FEOperacion.TextMatrix(oRs.Bookmark, 4) = oRs!cOpeCtaDH
   FEOperacion.TextMatrix(oRs.Bookmark, 2) = cOrden
   
   oRs.MoveNext
 Loop
End Sub

Private Sub LLenaFlexEditCtaFiltro()
 Dim oOpe As New DOperacion
 Dim oRs As New ADODB.Recordset
 Dim cOrden As String
 
 Set oRs = oOpe.CargaListaOperacionxCtaContOperacion(txtBuscarPersona.Text, FEOperacion.TextMatrix(FEOperacion.Row, 3), txtCtaIf.Text)
 LimpiaFlex FEFiltro
 Do While Not oRs.EOF

   FEFiltro.AdicionaFila
   FEFiltro.TextMatrix(oRs.Bookmark, 1) = txtBuscarPersona.Text
   FEFiltro.TextMatrix(oRs.Bookmark, 2) = oRs!cIFTpo
   FEFiltro.TextMatrix(oRs.Bookmark, 3) = oRs!cCtaContCod
   FEFiltro.TextMatrix(oRs.Bookmark, 4) = oRs!cCtaIFSubCta
   FEFiltro.TextMatrix(oRs.Bookmark, 5) = oRs!cCtaIfCod
   oRs.MoveNext
 Loop
End Sub

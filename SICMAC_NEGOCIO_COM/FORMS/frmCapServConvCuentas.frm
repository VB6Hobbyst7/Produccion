VERSION 5.00
Begin VB.Form frmCapServConvCuentas 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frmCapServConvCuentas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   60
      TabIndex        =   7
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdEliminar 
      Height          =   435
      Left            =   2880
      Picture         =   "frmCapServConvCuentas.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2220
      Width           =   555
   End
   Begin VB.CommandButton cmdAgregar 
      Height          =   435
      Left            =   2880
      Picture         =   "frmCapServConvCuentas.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1740
      Width           =   555
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8310
      TabIndex        =   6
      Top             =   3660
      Width           =   915
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   3660
      Width           =   915
   End
   Begin VB.Frame fraCuentasReg 
      Caption         =   "Cuentas Registradas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2535
      Left            =   3540
      TabIndex        =   11
      Top             =   1020
      Width           =   5700
      Begin SICMACT.FlexEdit grdCuentasReg 
         Height          =   2115
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   3731
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Cuenta-Tipo Cuenta-Monto-Estado"
         EncabezadosAnchos=   "350-1900-1900-1200-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-3-X"
         TextStyleFixed  =   2
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-R-L"
         FormatosEdit    =   "0-0-0-2-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame fraCuentasDisp 
      Caption         =   "Cuentas Disponibles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2595
      Left            =   60
      TabIndex        =   9
      Top             =   1020
      Width           =   2715
      Begin SICMACT.FlexEdit grdCuentasDisp 
         Height          =   2175
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   3836
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Cuenta"
         EncabezadosAnchos=   "350-1900"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X"
         TextStyleFixed  =   2
         ListaControles  =   "0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C"
         FormatosEdit    =   "0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame fraConvenio 
      Caption         =   "Institucion Covnenio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   60
      TabIndex        =   8
      Top             =   120
      Width           =   9150
      Begin SICMACT.TxtBuscar txtCodigo 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   1875
         _ExtentX        =   3307
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
      Begin VB.Label lblInstitucion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   300
         Width           =   6435
      End
   End
End
Attribute VB_Name = "frmCapServConvCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sPersonaCod As String

Private Sub HabilitaBotones()
If grdCuentasDisp.Rows = 2 And grdCuentasDisp.TextMatrix(1, 1) = "" Then
    cmdAgregar.Enabled = False
Else
    cmdAgregar.Enabled = True
End If
If grdCuentasReg.Rows = 2 Then
    If grdCuentasReg.TextMatrix(1, 1) = "" Then
        cmdEliminar.Enabled = False
    Else
        cmdEliminar.Enabled = True
    End If
End If
End Sub

Public Sub Inicia(Optional sPersona As String = "", Optional sPersonaDesc As String = "", Optional nTipoCOnv As Integer)
Dim RSTEMP As ADODB.Recordset, i As Integer
sPersonaCod = sPersona
txtCodigo.sTitulo = "Instituciones - Convenio"
If sPersonaCod = "" Then
    Dim clsCap As COMNCaptaServicios.NCOMCaptaServicios
    Dim rsCap As ADODB.Recordset
    Set clsCap = New COMNCaptaServicios.NCOMCaptaServicios
     Set rsCap = clsCap.GetServConveniosArbol()
    Set clsCap = Nothing
    txtCodigo.rs = rsCap
    cmdCancelar.Enabled = False
    fraConvenio.Enabled = True
    cmdAgregar.Enabled = False
    cmdEliminar.Enabled = False
      
Else
    txtCodigo.Text = sPersona
    lblInstitucion = sPersonaDesc
    fraConvenio.Enabled = False
    
    Set clsCap = New COMNCaptaServicios.NCOMCaptaServicios
        Set RSTEMP = clsCap.GetConvenioCuenta2(sPersona, nTipoCOnv)
    Set clsCap = Nothing
    If Not RSTEMP.EOF Then
        Set grdCuentasReg.Recordset = RSTEMP
    End If
    
    txtCodigo_EmiteDatos
     
    
End If
Me.Caption = "Captaciones - Servicio - Convenios - Cuentas"
Me.Show 1
End Sub

Private Sub CmdAgregar_Click()
If grdCuentasDisp.TextMatrix(grdCuentasDisp.Row, 1) <> "" Then
    grdCuentasReg.AdicionaFila , , True
    grdCuentasReg.TextMatrix(grdCuentasReg.Row, 1) = grdCuentasDisp.TextMatrix(grdCuentasDisp.Row, 1)
    grdCuentasReg.TextMatrix(grdCuentasReg.Row, 2) = "1.- Pension"
    grdCuentasReg.TextMatrix(grdCuentasReg.Row, 3) = "0.00"
    
    grdCuentasReg.Col = 2
    grdCuentasReg.SetFocus
    grdCuentasDisp.EliminaFila grdCuentasDisp.Row
    HabilitaBotones
End If
End Sub

Private Sub cmdCancelar_Click()
txtCodigo.Text = ""
lblInstitucion = ""
grdCuentasDisp.Clear
grdCuentasDisp.Rows = 2
grdCuentasDisp.FormaCabecera
grdCuentasReg.Clear
grdCuentasReg.Rows = 2
grdCuentasReg.FormaCabecera
fraCuentasReg.Enabled = False
fraCuentasDisp.Enabled = False
fraConvenio.Enabled = True
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
txtCodigo.SetFocus
End Sub

Private Sub cmdeliminar_Click()
If grdCuentasReg.Rows - 1 >= 1 Then
  Dim nnumrow As Long
    If grdCuentasReg.TextMatrix(grdCuentasReg.Row, 4) = "NO" Then
        MsgBox "Cuenta Cancelada o Anulada", vbOKOnly + vbInformation, "Aviso"
        Exit Sub
    End If
    nnumrow = grdCuentasReg.Row
    grdCuentasDisp.AdicionaFila
    
    grdCuentasDisp.TextMatrix(grdCuentasDisp.Row, 1) = grdCuentasReg.TextMatrix(grdCuentasReg.Row, 1)
    If nnumrow > 1 And grdCuentasReg.Rows - 1 > 1 Then
        grdCuentasReg.EliminaFila nnumrow
    ElseIf nnumrow = 1 And grdCuentasReg.Rows - 1 > 1 Then
        grdCuentasReg.EliminaFila nnumrow
    ElseIf nnumrow = 1 And grdCuentasReg.Rows - 1 = 1 Then
        grdCuentasReg.TextMatrix(nnumrow, 0) = ""
        grdCuentasReg.TextMatrix(nnumrow, 1) = ""
        grdCuentasReg.TextMatrix(nnumrow, 2) = ""
        grdCuentasReg.TextMatrix(nnumrow, 3) = ""
        grdCuentasReg.TextMatrix(nnumrow, 4) = ""
    End If
    'grdCuentasReg.EliminaFila grdCuentasDisp.Row
    
    HabilitaBotones
End If
End Sub

Private Sub cmdGrabar_Click()
Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
If grdCuentasReg.TextMatrix(1, 1) <> "" Then
    If MsgBox("¿Desea grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        clsServ.ActualizaServConvCuentas txtCodigo.Text, grdCuentasReg.GetRsNew
        cmdGrabar.Enabled = False
        cmdCancelar.Enabled = False
    End If
Else
    MsgBox ""
End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub
Public Function EstaConvenio(ByVal sCtaCod As String) As Boolean
 Dim i As Integer
  EstaConvenio = False
    If grdCuentasReg.Rows > 1 Then
      For i = 1 To grdCuentasReg.Rows - 1
        If grdCuentasReg.TextMatrix(i, 1) = sCtaCod Then
                EstaConvenio = True
                Exit Function
        End If
     Next i
    End If
        
End Function
Private Sub txtCodigo_EmiteDatos()
Dim sCodigo As String, sCuenta As String
sCodigo = txtCodigo.Text
If sCodigo <> "" Then
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rsCtas As ADODB.Recordset
    If txtCodigo.psDescripcion <> "" Then lblInstitucion = txtCodigo.psDescripcion
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCtas = clsCap.GetCuentasPersona(sCodigo, gCapAhorros, True)
    If Not (rsCtas.EOF And rsCtas.BOF) Then
        Do While Not rsCtas.EOF
           If rsCtas("nPrdPersRelac") = gCapRelPersTitular Then
                       
                sCuenta = rsCtas("cCtaCod")
             If EstaConvenio(sCuenta) = False Then
                
                grdCuentasDisp.AdicionaFila
                grdCuentasDisp.TextMatrix(grdCuentasDisp.Row, 1) = sCuenta
                
                If CLng(Mid(sCuenta, 9, 1)) = gMonedaExtranjera Then
                
                    grdCuentasDisp.BackColorRow (&HC0FFC0)
                    
                End If
                
             End If
                
           End If
            rsCtas.MoveNext
        Loop
        cmdGrabar.Enabled = True
        
        fraCuentasDisp.Enabled = True
        fraCuentasReg.Enabled = True
    Else
        MsgBox "La Persona NO POSEE CUENTAS DISPONIBLES. Aperture alguna cuenta y vuelva a intentar.", vbInformation, "Aviso"
        cmdGrabar.Enabled = False
        fraCuentasDisp.Enabled = False
        fraCuentasReg.Enabled = False
        cmdCancelar.Enabled = True
    End If
    fraConvenio.Enabled = False
End If
HabilitaBotones
End Sub

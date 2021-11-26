VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{A7C47A80-96CC-11CF-8B85-0020AFE89883}#4.0#0"; "SigBox.OCX"
Begin VB.Form frmCapTarjetaRelacionLOTE 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   FillColor       =   &H00008000&
   Icon            =   "frmCapTarjetaRelacionLOTE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegCard 
      Caption         =   "Registrar &Tarjeta"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3660
      TabIndex        =   24
      Top             =   6990
      Width           =   1590
   End
   Begin VB.Frame fraRegCard 
      Caption         =   " Clave"
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
      Height          =   750
      Left            =   4275
      TabIndex        =   20
      Top             =   0
      Width           =   3300
      Begin VB.TextBox lblpassw2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2145
         PasswordChar    =   "*"
         TabIndex        =   22
         Top             =   225
         Width           =   810
      End
      Begin VB.TextBox lblpassw1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1095
         PasswordChar    =   "*"
         TabIndex        =   21
         Top             =   225
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Password :"
         Height          =   165
         Left            =   165
         TabIndex        =   23
         Top             =   330
         Width           =   780
      End
   End
   Begin VB.Frame FraMiembros 
      Caption         =   "Otros Miembros"
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
      ForeColor       =   &H00800000&
      Height          =   2100
      Left            =   105
      TabIndex        =   11
      Top             =   4770
      Width           =   7575
      Begin SICMACT.FlexEdit grdMiembros 
         Height          =   1695
         Left            =   75
         TabIndex        =   12
         Top             =   255
         Width           =   7320
         _ExtentX        =   12912
         _ExtentY        =   2990
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Codigo-Nombre-Relacion"
         EncabezadosAnchos=   "250-1800-3900-1200"
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
         ColumnasAEditar =   "X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0"
         EncabezadosAlineacion=   "C-L-C-C"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "#"
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   300
      End
   End
   Begin VB.CommandButton cmdConvenio 
      Caption         =   "Con&venio Tarjeta"
      Height          =   375
      Left            =   5325
      TabIndex        =   5
      Top             =   6975
      Width           =   2115
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   105
      TabIndex        =   6
      Top             =   6975
      Width           =   960
   End
   Begin VB.Frame fraPersona 
      Caption         =   "Persona Titular"
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
      Height          =   1560
      Left            =   105
      TabIndex        =   10
      Top             =   825
      Width           =   7515
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2220
         TabIndex        =   13
         Top             =   345
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   390
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   90
         TabIndex        =   18
         Top             =   900
         Width           =   600
      End
      Begin VB.Label lblNomCli 
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
         ForeColor       =   &H00800000&
         Height          =   540
         Left            =   810
         TabIndex        =   17
         Top             =   840
         Width           =   6465
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Identidad:"
         Height          =   195
         Left            =   3795
         TabIndex        =   16
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label lbllCodigoCli 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   825
         TabIndex        =   15
         Top             =   330
         Width           =   1380
      End
      Begin VB.Label lblDICli 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4980
         TabIndex        =   14
         Top             =   345
         Width           =   2250
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   1095
      TabIndex        =   4
      Top             =   6975
      Width           =   960
   End
   Begin VB.Frame fraRelacion 
      Caption         =   "Relación NO ASOCIADAS"
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
      ForeColor       =   &H00800000&
      Height          =   2280
      Left            =   105
      TabIndex        =   9
      Top             =   2430
      Width           =   7530
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   6495
         Picture         =   "frmCapTarjetaRelacionLOTE.frx":030A
         TabIndex        =   3
         Top             =   1785
         Width           =   960
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   5445
         Picture         =   "frmCapTarjetaRelacionLOTE.frx":064C
         TabIndex        =   2
         Top             =   1785
         Width           =   960
      End
      Begin SICMACT.FlexEdit grdClienteTarj 
         Height          =   1380
         Left            =   75
         TabIndex        =   1
         Top             =   315
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   2434
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Cuenta-Producto-Moneda-Estado-Personeria"
         EncabezadosAnchos=   "250-1800-2000-1000-2100-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "#"
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   300
      End
   End
   Begin VB.Frame fraTarjeta 
      Enabled         =   0   'False
      Height          =   750
      Left            =   105
      TabIndex        =   7
      Top             =   0
      Width           =   4005
      Begin MSMask.MaskEdBox txtTarjeta 
         Height          =   375
         Left            =   945
         TabIndex        =   0
         Top             =   210
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "####-####-####-####"
         Mask            =   "####-####-####-####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   300
         Width           =   735
      End
   End
   Begin SigBoxLib.SigBox boxFirma 
      Height          =   615
      Left            =   2445
      TabIndex        =   25
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
      _Version        =   262144
      _ExtentX        =   1720
      _ExtentY        =   1085
      _StockProps     =   233
      Appearance      =   1
      TitleText       =   ""
      PromptText      =   ""
      ConnectToPad    =   0
      Picture         =   "frmCapTarjetaRelacionLOTE.frx":098E
      DebugFileName   =   "SigBox1.TXT"
   End
   Begin RichTextLib.RichTextBox rtfCartas 
      Height          =   360
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   635
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmCapTarjetaRelacionLOTE.frx":09AA
   End
End
Attribute VB_Name = "frmCapTarjetaRelacionLOTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bConsulta As Boolean

Public Sub LimpiaPantalla()
'grdCliente.Clear
'grdCliente.Rows = 2
'SetupGridCliente
Me.lblDICli.Caption = ""
Me.lbllCodigoCli.Caption = ""
Me.lblNomCli.Caption = ""

Me.lblpassw1.Text = ""
Me.lblpassw2.Text = ""

grdClienteTarj.Clear
grdClienteTarj.Rows = 2
grdClienteTarj.FormaCabecera

grdMiembros.Clear
grdMiembros.Rows = 2
grdMiembros.FormaCabecera


txtTarjeta.Text = "____-____-____-____"
fraTarjeta.Enabled = False
Me.fraPersona.Enabled = True
Me.fraRelacion.Enabled = False
Me.FraMiembros.Enabled = False
Me.cmdRegCard.Enabled = False
cmdConvenio.Enabled = False

End Sub

Private Function ExisteCuentaRelacion(ByVal sCuenta As String) As Boolean
Dim i As Integer, nFila As Integer, nCol As Integer
Dim bExiste As Boolean
nFila = grdClienteTarj.Row
nCol = grdClienteTarj.Col
bExiste = False
For i = 1 To grdClienteTarj.Rows - 1
    If grdClienteTarj.TextMatrix(i, 1) = sCuenta Then
        bExiste = True
        Exit For
    End If
Next i
grdClienteTarj.Row = nFila
grdClienteTarj.Col = nCol
ExisteCuentaRelacion = bExiste
End Function

Public Sub SetupGridCliente()
'Dim I As Integer
'For I = 1 To grdCliente.Rows - 1
'    grdCliente.MergeCol(I) = True
'Next I
'grdCliente.MergeCells = flexMergeFree
'grdCliente.BandExpandable(0) = True
'grdCliente.Cols = 9
'grdCliente.ColWidth(0) = 100
'grdCliente.ColWidth(1) = 3500
'grdCliente.ColWidth(2) = 3500
'grdCliente.ColWidth(3) = 1500
'grdCliente.ColWidth(4) = 1000
'grdCliente.ColWidth(5) = 600
'grdCliente.ColWidth(6) = 1500
'grdCliente.ColWidth(7) = 0
'grdCliente.ColWidth(8) = 0
'grdCliente.TextMatrix(0, 1) = "Nombre"
'grdCliente.TextMatrix(0, 2) = "Dirección"
'grdCliente.TextMatrix(0, 3) = "Zona"
'grdCliente.TextMatrix(0, 4) = "Fono"
'grdCliente.TextMatrix(0, 5) = "ID"
'grdCliente.TextMatrix(0, 6) = "ID N°"
End Sub

Public Sub ObtieneDatosTarjeta()
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim rsTarj As New ADODB.Recordset
Dim sTarjeta As String, sPersona As String
sTarjeta = Replace(txtTarjeta, "-", "", 1, , vbTextCompare)

Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsTarj = clsMant.GetTarjetaCuentas(sTarjeta)
If rsTarj.EOF And rsTarj.BOF Then
    fraPersona.Enabled = True
    fraRelacion.Enabled = True
    FraMiembros.Enabled = True
        'ObtieneCuentasNoAsociadas
Else
        
    fraPersona.Enabled = False
    fraRelacion.Enabled = False
    FraMiembros.Enabled = False
    
    MsgBox "Esta tarjeta, ya posee cuentas relacionadas." & vbCrLf & " Ingrese a la Opcion de Relacion x Cuenta para relacionar una cuenta a esta tarjeta.", vbInformation, "Aviso"
    Exit Sub


End If
Set rsTarj = Nothing
Set clsMant = Nothing
End Sub
Private Sub ObtieneCuentasNoAsociadas()
 Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales, rstmp As ADODB.Recordset
 Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
 Set rstmp = New ADODB.Recordset
 
Set rstmp = clsMant.GetCuentasNoAsociadas(Me.lbllCodigoCli.Caption)
  
If Not (rstmp.EOF Or rstmp.BOF) Then
        Set grdClienteTarj.Recordset = rstmp
        If grdClienteTarj.TextMatrix(1, 1) <> "" Then
           Call ObtieneInfoOtrosMiembros(grdClienteTarj.TextMatrix(1, 1), Me.lbllCodigoCli, grdClienteTarj.TextMatrix(1, 5))
        End If
        cmdRegCard.Enabled = True
Else
        MsgBox "No existen cuentas no asociadas para este cliente.", vbOKOnly + vbInformation, "AVISO"
End If
 
Set rstmp = Nothing
 
Set clsMant = Nothing
 
End Sub

Public Sub Inicia(Optional bCons As Boolean = False)
bConsulta = bCons
If bConsulta Then
'    cmdGrabar.Visible = False
    cmdAgregar.Visible = False
    cmdEliminar.Visible = False
    cmdCancelar.Visible = False
Else
'    cmdGrabar.Visible = True
    cmdAgregar.Visible = True
    cmdEliminar.Visible = True
    cmdCancelar.Visible = True
'    cmdGrabar.Enabled = False
    cmdAgregar.Enabled = False
    cmdEliminar.Enabled = False
    cmdCancelar.Enabled = False
End If
cmdConvenio.Enabled = False
'SetupGridCliente
Me.Caption = "Captaciones - Tarjeta - Relación"
Me.Show 1
End Sub

Private Sub cmdAgregar_Click()
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsCta As ADODB.Recordset
Dim sPersona As String
sPersona = Me.lbllCodigoCli
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
'Set rsCta = clsmant.GetCuentasPersona(sPersona, , True)
Set rsCta = clsMant.GetCuentasNoAsociadas(sPersona)
If Not (rsCta.EOF And rsCta.BOF) Then
    Dim clsCap As UCapCuenta
    Dim sCuenta As String
    frmCapMantenimientoCtas.lstCuentas.Clear
    Do While Not rsCta.EOF
        sCuenta = rsCta("CUENTA")
        If Not ExisteCuentaRelacion(sCuenta) Then
            frmCapMantenimientoCtas.lstCuentas.AddItem sCuenta & Space(2) & rsCta("Producto") & Space(2) & rsCta("Moneda") & Space(2) & rsCta("Estado") & Space(2) & CStr(rsCta("Personeria"))
        End If
        rsCta.MoveNext
    Loop
    Set clsCap = frmCapMantenimientoCtas.Inicia(1)
    If Not clsCap Is Nothing Then
        sCuenta = clsCap.sCtaCod
        If sCuenta <> "" Then
            Dim nItem As Long
            grdClienteTarj.AdicionaFila
            nItem = grdClienteTarj.Rows - 1
            grdClienteTarj.TextMatrix(nItem, 1) = sCuenta
            grdClienteTarj.TextMatrix(nItem, 2) = clsCap.sProducto
            grdClienteTarj.TextMatrix(nItem, 3) = clsCap.sMoneda
            grdClienteTarj.TextMatrix(nItem, 4) = UCase(clsCap.sEstado)
            grdClienteTarj.TextMatrix(nItem, 5) = clsCap.sPersoneria
            
            'cmdGrabar.Enabled = True
            cmdCancelar.Enabled = True
        End If
    End If
Else
End If
Set clsMant = Nothing
End Sub

Private Sub cmdCancelar_Click()
cmdCancelar.Enabled = False
LimpiaPantalla
End Sub

Private Sub cmdConsultar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String, lsDni As String
Dim lsEstados As String


On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
    lsDni = loPers.sPersIdnroDNI

If lsPersCod <> "" Then

    Me.lbllCodigoCli.Caption = lsPersCod
    Me.lblNomCli.Caption = lsPersNombre
    Me.lblDICli.Caption = lsDni
    
    ObtieneCuentasNoAsociadas
    'fraTarjeta.Enabled =false
    Me.fraRelacion.Enabled = True
    cmdCancelar.Enabled = True
Else
    LimpiaPantalla
End If

Set loPers = Nothing
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdConvenio_Click()
Dim clsPrev As previo.clsPrevio
Dim sConvenio As String, sCuenta As String
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim nItem As Long
Dim lsCadImp As String
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales

nItem = grdClienteTarj.Row
sCuenta = grdClienteTarj.TextMatrix(nItem, 1)

rtfCartas.Filename = App.path & "\FormatoCarta\ConTar.txt"
lsCadImp = clsMant.GeneraConvenioTarjeta(rtfCartas.Text, sCuenta, gdFecSis)

Set clsMant = Nothing
Set clsPrev = New previo.clsPrevio
clsPrev.Show lsCadImp, "Convenio Tarjeta Magnética", True, , gImpresora
Set clsPrev = Nothing
End Sub

Private Sub cmdeliminar_Click()
If MsgBox("¿Está seguro de eliminar la cuenta relacionada?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim nItem As Long
    nItem = grdClienteTarj.Row
    grdClienteTarj.EliminaFila nItem
    If grdClienteTarj.Rows <= 2 Then
        cmdEliminar.Enabled = True
    End If
End If
End Sub

Private Sub cmdRegCard_Click()
Dim nCOM As TipoPuertoSerial
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim sMaquina As String
sMaquina = GetComputerName


'opciones de validacion
Set clsGen = New COMDConstSistema.DCOMGeneral
nCOM = clsGen.GetPuertoPeriferico(gPerifPENWARE, sMaquina)
If nCOM = -1 Then
    nCOM = clsGen.GetPuertoPeriferico(gPerifPINPAD, sMaquina)
    IniciaPinPad nCOM
    GrabaTarjetaPINPAD
    fraPersona.Enabled = False
    cmdAgregar.Enabled = False
    cmdEliminar.Enabled = False
    FraMiembros.Enabled = False
    fraTarjeta.Enabled = False
    
Else
    'If ConectaPad Then
       ' Habilitar
       ' MuestraPantalla
   ' End If
End If
Set clsGen = Nothing
End Sub

Private Sub AgregaTarjeta()
Dim clsMov As COMNContabilidad.NCOMContFunciones 'NContFunciones
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim nItem As Long
Dim sMovNro As String, sPersona As String, sTarjeta As String
Dim sCuenta As String, sClave As String
Dim lsCadImp As String
Dim loPrevio As previo.clsPrevio

Dim CLSSERV As COMNCaptaServicios.NCOMCaptaServicios 'NCapServicios
Set CLSSERV = New COMNCaptaServicios.NCOMCaptaServicios
sClave = Encripta(Trim(lblpassw1.Text), True)
'sClave = Encripta(sClave, False)   prueba desemcripta

Set clsMov = New COMNContabilidad.NCOMContFunciones
sMovNro = clsMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
Set clsMov = Nothing
'nItem = grdCliente.Row
sTarjeta = Trim(Replace(Me.txtTarjeta.Text, "-", ""))


Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
'If clsmant.AgregaTarjeta(sTarjeta, sClave, sMovNro, sPersona, sCuenta, DateAdd("y", 1, gdFecSis)) Then
If clsMant.AgregaTarjetaLOTE(sTarjeta, sClave, sMovNro, grdClienteTarj.GetRsNew, DateAdd("y", 1, gdFecSis), Trim(Me.lbllCodigoCli.Caption)) Then
    lsCadImp = CLSSERV.ImprimeBolTarjeta("REGISTRO TARJETA", _
                                    Trim(lblNomCli.Caption), txtTarjeta.Text, _
                                    "TARJEA MAGNETICA", gdFecSis, gsNomAge, _
                                    gsCodUser, sLpt)
    Do
       Set loPrevio = New previo.clsPrevio
            loPrevio.PrintSpool sLpt, lsCadImp, False
            loPrevio.PrintSpool sLpt, Chr(10) & Chr(10) & Chr(10) & Chr(10) & lsCadImp, False
       Set loPrevio = Nothing
    Loop Until MsgBox("DESEA REIMPRIMIR BOLETA?", vbYesNo, "AVISO") = vbNo

    cmdConvenio.Enabled = True
    cmdCancelar.Enabled = True
'    cmdCancelar_Click
Else
    'InHabilitar
    cmdCancelar.Enabled = True
    'lblTrack1.Caption = ""
    lblpassw1.Text = ""
    lblpassw2.Text = ""
    cmdRegCard.Enabled = True
    'cmdCancelaCard.Enabled = True
    'UnSetupPad
End If
Set CLSSERV = Nothing
Set clsMant = Nothing
End Sub


Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Me.Icon = LoadPicture(App.path & gsRutaIcono)
'If KeyCode = vbKeyF11 Then 'F11
'    Dim nCOM As TipoPuertoSerial
'    Dim clsGen As DGeneral
'
'    cmdCancelar.Enabled = True
'    'opciones de validacion
'    Set clsGen = New DGeneral
'    nCOM = clsGen.GetPuertoPeriferico(gPerifPENWARE)
'    If nCOM = -1 Then
'        nCOM = clsGen.GetPuertoPeriferico(gPerifPINPAD)
'        'ppoa Modificacion
'        If Not IniciaPinPad(nCOM) Then
'            MsgBox "No Inicio Dispositivo" & ". Consulte con Servicio Tecnico.", vbInformation, "Aviso"
'            Exit Sub
'        End If
'        If Not GrabaTarjetaPINPAD Then cmdCancelar_Click
'
'    Else
''        If ConectaPad Then
''            boxFirma.MagCardEnabled = False
''            MuestraPantalla
''        End If
'    End If
'End If
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub grdClienteTarj_Click()
If Me.cmdRegCard.Enabled Then
    cmdAgregar.Enabled = True
    If Me.grdClienteTarj.Rows > 2 Then
       If Me.grdClienteTarj.TextMatrix(2, 1) <> "" Then
           cmdEliminar.Enabled = True
       End If
    End If
End If
End Sub

Private Sub grdClienteTarj_OnRowChange(pnRow As Long, pnCol As Long)
  Call ObtieneInfoOtrosMiembros(grdClienteTarj.TextMatrix(pnRow, 1), Me.lbllCodigoCli, grdClienteTarj.TextMatrix(pnRow, 5))
End Sub

Private Sub ObtieneInfoOtrosMiembros(ByVal sCuenta As String, ByVal sPersCod As String, ByVal pnPersoneria As Integer)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales, rstmp As New ADODB.Recordset
  Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
  
  grdMiembros.Clear
  grdMiembros.Rows = 2
  grdMiembros.FormaCabecera
  
  Set rstmp = clsMant.GetPersonasCuentaNA(sCuenta, sPersCod, pnPersoneria)
  If Not (rstmp.EOF Or rstmp.BOF) Then
  
      Set grdMiembros.Recordset = rstmp
      
  End If
    
End Sub

Private Function GrabaTarjetaPINPAD()

Dim sNumTar As String
Dim sClaveTar As String
Dim lnErr As Long
Dim lnNumOp As Integer, lsCodTar As String, lblTrack1 As String
Dim sTitulo As String

sTitulo = Me.Caption

'ppoa Modificacion
If Not WriteToLcd("Pase su Tarjeta por la Lectora.") Then
    FinalizaPinPad
    MsgBox "No se Realizó Envío", vbInformation, "Aviso"
    Exit Function
End If

Me.Caption = "Lectura de Tarjeta Activada. Pase la tarjeta por la Lectora."

'ppoa Modificacion
sNumTar = GetNumTarjeta


lsCodTar = sNumTar
sNumTar = Replace(sNumTar, "-", "", 1, , vbTextCompare)
'txtTarjeta.Text = sNumTar

If Len(sNumTar) <> 16 Then
    MsgBox "Error en la Lectura de Tarjeta.", vbInformation, "Aviso"
    FinalizaPinPad
    Exit Function
End If

Me.Caption = "Ingrese la Clave de la Tarjeta."

'If Not WriteToLcd("                                       ") Then
'    FinalizaPinPad
'    MsgBox "No se Realizó Envío", vbInformation, "Aviso"
'    Exit Function
'End If


fraTarjeta.Enabled = True
txtTarjeta.Text = Format(sNumTar, "0000-0000-0000-0000")
fraTarjeta.Enabled = False
        
'ppoa Modificacion
sClaveTar = GetClaveTarjeta
                        
                        
If sClaveTar = "" Then
    MsgBox "Debe Ingresar una Clave Valida.", vbInformation, "Aviso"
    lblTrack1 = ""
    Exit Function
End If

lblpassw1 = sClaveTar
sClaveTar = ""
lnNumOp = 0

While lnNumOp < 3 And lblpassw1 <> sClaveTar
    sClaveTar = GetClaveTarjeta
    lnNumOp = lnNumOp + 1
    If lblpassw1 <> sClaveTar And lnNumOp < 3 Then
        MsgBox "La clave es errada. Re-Ingrese su Clave.", vbInformation, "Aviso"
    End If
Wend

lblpassw2 = sClaveTar
If lblpassw1 = lblpassw2 Then
    If MsgBox("Desea Registrar la Tarjeta ? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        AgregaTarjeta
    Else
       cmdConvenio.Enabled = False
    End If
Else
    MsgBox "La clave no ha sido reconocida, el proceso sera Cancelado.", vbInformation, "Aviso"
End If

FinalizaPinPad
cmdRegCard.Enabled = False
Me.Caption = sTitulo
End Function

Private Sub txtTarjeta_GotFocus()
With txtTarjeta
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtTarjeta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ObtieneDatosTarjeta
End If
End Sub

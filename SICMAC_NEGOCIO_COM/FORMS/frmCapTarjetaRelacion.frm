VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCapTarjetaRelacion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   FillColor       =   &H00008000&
   Icon            =   "frmCapTarjetaRelacion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConvenio 
      Caption         =   "Con&venio Tarjeta"
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   4725
      Width           =   2115
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   105
      TabIndex        =   9
      Top             =   4725
      Width           =   960
   End
   Begin VB.Frame fraPersona 
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
      ForeColor       =   &H00800000&
      Height          =   1485
      Left            =   105
      TabIndex        =   13
      Top             =   840
      Width           =   7155
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCliente 
         Height          =   1170
         Left            =   105
         TabIndex        =   1
         Top             =   210
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   2064
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2205
      TabIndex        =   7
      Top             =   4725
      Width           =   960
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   1155
      TabIndex        =   6
      Top             =   4725
      Width           =   960
   End
   Begin VB.Frame fraRelacion 
      Caption         =   "Relación"
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
      Height          =   2220
      Left            =   105
      TabIndex        =   12
      Top             =   2415
      Width           =   7155
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   5985
         Picture         =   "frmCapTarjetaRelacion.frx":030A
         TabIndex        =   5
         Top             =   1785
         Width           =   960
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   4935
         Picture         =   "frmCapTarjetaRelacion.frx":064C
         TabIndex        =   4
         Top             =   1785
         Width           =   960
      End
      Begin SICMACT.FlexEdit grdClienteTarj 
         Height          =   1380
         Left            =   105
         TabIndex        =   3
         Top             =   315
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   2434
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Cuenta-Producto-Moneda-cPersCod-Relacion-cEstado"
         EncabezadosAnchos=   "250-1800-2000-1500-0-1200-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   285
      End
   End
   Begin VB.Frame fraTarjeta 
      Height          =   750
      Left            =   105
      TabIndex        =   10
      Top             =   0
      Width           =   4005
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   210
         Width           =   435
      End
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
         TabIndex        =   11
         Top             =   300
         Width           =   735
      End
   End
   Begin RichTextLib.RichTextBox rtfCartas 
      Height          =   330
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   582
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmCapTarjetaRelacion.frx":098E
   End
End
Attribute VB_Name = "frmCapTarjetaRelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bConsulta As Boolean

Public Sub LimpiaPantalla()
grdCliente.Clear
grdCliente.Rows = 2
SetupGridCliente
grdClienteTarj.Clear
grdClienteTarj.Rows = 2
grdClienteTarj.FormaCabecera
txtTarjeta.Text = "____-____-____-____"
fraTarjeta.Enabled = True
txtTarjeta.SetFocus
cmdConvenio.Enabled = False
cmdGrabar.Enabled = False
End Sub

Private Function ExisteCuentaRelacion(ByVal sCuenta As String) As Boolean
Dim I As Integer, nFila As Integer, nCol As Integer
Dim bExiste As Boolean
nFila = grdClienteTarj.Row
nCol = grdClienteTarj.Col
bExiste = False
For I = 1 To grdClienteTarj.Rows - 1
    If grdClienteTarj.TextMatrix(I, 1) = sCuenta Then
        bExiste = True
        Exit For
    End If
Next I
grdClienteTarj.Row = nFila
grdClienteTarj.Col = nCol
ExisteCuentaRelacion = bExiste
End Function

Public Sub SetupGridCliente()
Dim I As Integer
For I = 1 To grdCliente.Rows - 1
    grdCliente.MergeCol(I) = True
Next I
grdCliente.MergeCells = flexMergeFree
grdCliente.BandExpandable(0) = True
grdCliente.Cols = 9
grdCliente.ColWidth(0) = 100
grdCliente.ColWidth(1) = 3500
grdCliente.ColWidth(2) = 3500
grdCliente.ColWidth(3) = 1500
grdCliente.ColWidth(4) = 1000
grdCliente.ColWidth(5) = 600
grdCliente.ColWidth(6) = 1500
grdCliente.ColWidth(7) = 0
grdCliente.ColWidth(8) = 0
grdCliente.TextMatrix(0, 1) = "Nombre"
grdCliente.TextMatrix(0, 2) = "Dirección"
grdCliente.TextMatrix(0, 3) = "Zona"
grdCliente.TextMatrix(0, 4) = "Fono"
grdCliente.TextMatrix(0, 5) = "ID"
grdCliente.TextMatrix(0, 6) = "ID N°"
End Sub

Public Sub ObtieneDatosTarjeta()
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsTarj As New ADODB.Recordset
Dim sTarjeta As String, sPersona As String
sTarjeta = Replace(txtTarjeta, "-", "", 1, , vbTextCompare)

Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsTarj = clsMant.GetTarjetaCuentas(sTarjeta)
If rsTarj.EOF And rsTarj.BOF Then
    MsgBox "Tarjeta no posee ninguna relación con cuentas activas o Tarjeta no activa.", vbInformation, "Aviso"
    Exit Sub
Else
    sPersona = rsTarj("cPersCod")
    Set grdClienteTarj.Recordset = rsTarj
    Set rsTarj = clsMant.GetDatosPersona(sPersona)
    Set grdCliente.Recordset = rsTarj
    SetupGridCliente
    If Not bConsulta Then
        cmdEliminar.Enabled = True
        cmdGrabar.Enabled = True
        cmdAgregar.Enabled = True
        cmdCancelar.Enabled = True
        fraTarjeta.Enabled = False
    End If
    cmdConvenio.Enabled = True
End If
Set rsTarj = Nothing
Set clsMant = Nothing
End Sub

Public Sub Inicia(Optional bCons As Boolean = False)
bConsulta = bCons
If bConsulta Then
    cmdGrabar.Visible = False
    cmdAgregar.Visible = False
    cmdEliminar.Visible = False
    cmdCancelar.Visible = False
Else
    cmdGrabar.Visible = True
    cmdAgregar.Visible = True
    cmdEliminar.Visible = True
    cmdCancelar.Visible = True
    cmdGrabar.Enabled = False
    cmdAgregar.Enabled = False
    cmdEliminar.Enabled = False
    cmdCancelar.Enabled = False
End If
cmdConvenio.Enabled = False
SetupGridCliente
Me.Caption = "Captaciones - Tarjeta - Relación"
Me.Show 1
End Sub

Private Sub cmdAgregar_Click()
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsCta As New ADODB.Recordset
Dim sPersona As String
sPersona = grdCliente.TextMatrix(1, 7)
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsCta = clsMant.GetCuentasPersona(sPersona, , True)
If Not (rsCta.EOF And rsCta.BOF) Then
    Dim clsCap As UCapCuenta
    Dim sCuenta As String
    frmCapMantenimientoCtas.lstCuentas.Clear
    Do While Not rsCta.EOF
        sCuenta = rsCta("cCtaCod")
        If Not ExisteCuentaRelacion(sCuenta) Then
            frmCapMantenimientoCtas.lstCuentas.AddItem sCuenta & Space(2) & rsCta("cProducto") & Space(2) & rsCta("cMoneda") & Space(2) & rsCta("cRelacion")
        End If
        rsCta.MoveNext
    Loop
    Set clsCap = frmCapMantenimientoCtas.Inicia
    If Not clsCap Is Nothing Then
        sCuenta = clsCap.sCtaCod
        If sCuenta <> "" Then
            Dim nItem As Long
            grdClienteTarj.AdicionaFila
            nItem = grdClienteTarj.Rows - 1
            grdClienteTarj.TextMatrix(nItem, 1) = sCuenta
            grdClienteTarj.TextMatrix(nItem, 2) = clsCap.sProducto
            grdClienteTarj.TextMatrix(nItem, 3) = clsCap.sMoneda
            grdClienteTarj.TextMatrix(nItem, 5) = UCase(clsCap.sRelacion)
            cmdGrabar.Enabled = True
            cmdCancelar.Enabled = True
        End If
    End If
Else
End If
End Sub

Private Sub CmdEditar_Click()
cmdAgregar.Enabled = True
cmdEliminar.Enabled = True
cmdCancelar.Enabled = True
cmdGrabar.Enabled = True
cmdAgregar.SetFocus
End Sub

Private Sub cmdBuscar_Click()
Dim clsPers As COMDPersona.UCOMPersona
Set clsPers = New COMDPersona.UCOMPersona
Set clsPers = frmBuscaPersona.Inicio

If Not clsPers Is Nothing Then
    Dim sPers As String, sTarj As String, sTarjBus
    Dim rsPers As New ADODB.Recordset
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
    Dim clsTarj As UCapCuenta
    sPers = clsPers.sPersCod
    Set clsPers = Nothing
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsPers = clsCap.GetPersonaTarj(sPers)
    
    If Not (rsPers.EOF And rsPers.EOF) Then
        sTarj = ""
        sTarjBus = Replace(txtTarjeta, "-", "", 1, , vbTextCompare)
        Do While Not rsPers.EOF
            If sTarj <> rsPers("cTarjCod") And rsPers("cTarjCod") <> sTarjBus Then
                sTarj = rsPers("cTarjCod")
                frmCapMantenimientoCtas.lstCuentas.AddItem sTarj & Space(2)
            End If
            rsPers.MoveNext
        Loop
        Set clsTarj = frmCapMantenimientoCtas.Inicia
    End If
    If Not clsTarj Is Nothing Then
        sTarj = Trim(clsTarj.sCtaCod)
        Set clsTarj = Nothing
        grdClienteTarj.Clear
        grdClienteTarj.FormaCabecera
        grdClienteTarj.Rows = 2
        Set rsPers = clsCap.GetTarjetaCuentas(sTarj)
        If Not (rsPers.EOF And rsPers.EOF) Then
            txtTarjeta.Text = Format$(sTarj, "0###-####-####-####")
            Set grdClienteTarj.Recordset = rsPers
            Set rsPers = clsCap.GetDatosPersona(sPers)
            Set grdCliente.Recordset = rsPers
            SetupGridCliente
            cmdAgregar.Enabled = True
            cmdEliminar.Enabled = True
        End If
    Else
        MsgBox "Persona NO posee tarjetas relacionadas", vbInformation, "Aviso"
    End If
    Set clsCap = Nothing
    rsPers.Close
    Set rsPers = Nothing
Else
    cmdBuscar.SetFocus
End If
End Sub

Private Sub cmdCancelar_Click()
cmdCancelar.Enabled = False
LimpiaPantalla
End Sub

Private Sub cmdConvenio_Click()
Dim clsPrev As previo.clsprevio
Dim sConvenio As String, sCuenta As String
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales   'NCapMantenimiento
Dim nItem As Long
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
nItem = grdClienteTarj.Row
sCuenta = grdClienteTarj.TextMatrix(nItem, 1)

rtfCartas.Filename = App.path & "\FormatoCarta\ConTar.txt"

sConvenio = clsMant.GeneraConvenioTarjeta(rtfCartas.Text, sCuenta, gdFecSis)

Set clsMant = Nothing
Set clsPrev = New previo.clsprevio
'ALPA 20100202*****************************
'clsPrev.Show sConvenio, "Convenio Tarjeta Magnética", True
clsPrev.Show sConvenio, "Convenio Tarjeta Magnética", True, , gImpresora
Set clsPrev = Nothing
End Sub

Private Sub cmdEliminar_Click()
If MsgBox("¿Está seguro de eliminar la cuenta relacionada?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim nItem As Long
    nItem = grdClienteTarj.Row
    grdClienteTarj.EliminaFila nItem
    If grdClienteTarj.Rows <= 2 Then
        cmdEliminar.Enabled = True
    End If
End If
End Sub

Private Sub cmdGrabar_Click()
Dim srelctas As String, I As Integer
If MsgBox("¿Está seguro de grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim CLSSERV As COMNCaptaServicios.NCOMCaptaServicios
    Set CLSSERV = New COMNCaptaServicios.NCOMCaptaServicios
    Dim rsTarj As New ADODB.Recordset
    Dim sTarjeta As String, sPersona As String
    Dim lscadimp As String
    Dim loPrevio As previo.clsprevio
    Set rsTarj = grdClienteTarj.GetRsNew
    sPersona = grdCliente.TextMatrix(1, 7)
    sTarjeta = Replace(txtTarjeta.Text, "-", "", 1, , vbTextCompare)
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    If clsMant.ActualizaCuentaTarj(sTarjeta, sPersona, rsTarj) Then
    
            
'        For i = 1 To grdClienteTarj.Rows - 1
'            If i Mod 3 = 0 Then
'                srelctas = srelctas & "," & Chr(10)
'                srelctas = srelctas & grdClienteTarj.TextMatrix(i, 1)
'            Else
'                srelctas = IIf(i > 1, ",", "") & grdClienteTarj.TextMatrix(i, 1)
'            End
'        Next i
    
        lscadimp = CLSSERV.ImprimeBolTarjeta("RELACIONA TARJETA-CUENTAS", _
                                    Trim(grdCliente.TextMatrix(1, 1)), txtTarjeta.Text, _
                                    "TARJEA MAGNETICA", gdFecSis, gsNomAge, _
                                    gsCodUser, sLpt)
        Do
            Set loPrevio = New previo.clsprevio
                loPrevio.PrintSpool sLpt, lscadimp, False
                loPrevio.PrintSpool sLpt, Chr(10) & Chr(10) & Chr(10) & Chr(10) & lscadimp, False
            Set loPrevio = Nothing
        Loop Until MsgBox("DESEA REIMPRIMIR BOLETA?", vbYesNo, "AVISO") = vbNo

        LimpiaPantalla
    End If
    Set clsMant = Nothing
    Set CLSSERV = Nothing
End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

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

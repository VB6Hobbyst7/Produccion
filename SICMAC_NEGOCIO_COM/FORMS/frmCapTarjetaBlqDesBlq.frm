VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmCapTarjetaBlqDesBlq 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "frmCapTarjetaBlqDesBlq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   75
      TabIndex        =   9
      Top             =   5460
      Width           =   960
   End
   Begin VB.Frame fraHistoria 
      Caption         =   "Historia"
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
      Height          =   1905
      Left            =   3675
      TabIndex        =   16
      Top             =   2100
      Width           =   3585
      Begin SICMACT.FlexEdit grdTarjetaEstado 
         Height          =   1380
         Left            =   105
         TabIndex        =   4
         Top             =   315
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   2434
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Fecha-Estado-Comentario-Usu"
         EncabezadosAnchos=   "250-1000-2000-3500-600"
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
         EncabezadosAlineacion=   "C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   285
      End
   End
   Begin VB.Frame fraEstado 
      Caption         =   "Estado"
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
      Height          =   1380
      Left            =   105
      TabIndex        =   14
      Top             =   3990
      Width           =   7155
      Begin VB.CheckBox chkBloqueo 
         Alignment       =   1  'Right Justify
         Caption         =   "&Bloqueada"
         Height          =   225
         Left            =   210
         TabIndex        =   5
         Top             =   315
         Width           =   1275
      End
      Begin VB.TextBox txtGlosa 
         Appearance      =   0  'Flat
         Height          =   540
         Left            =   1260
         TabIndex        =   6
         Top             =   630
         Width           =   5685
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   210
         TabIndex        =   15
         Top             =   630
         Width           =   495
      End
   End
   Begin VB.Frame fraTarjeta 
      Height          =   750
      Left            =   105
      TabIndex        =   12
      Top             =   0
      Width           =   4005
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   375
         Left            =   3360
         TabIndex        =   1
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
         TabIndex        =   13
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Frame fraRelacion 
      Caption         =   "Cuentas Relacionadas"
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
      Height          =   1905
      Left            =   105
      TabIndex        =   3
      Top             =   2100
      Width           =   3480
      Begin SICMACT.FlexEdit grdClienteTarj 
         Height          =   1380
         Left            =   105
         TabIndex        =   11
         Top             =   315
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   2434
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Cuenta-Relación"
         EncabezadosAnchos=   "250-1800-1100"
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
         ColumnasAEditar =   "X-X-X"
         ListaControles  =   "0-0-0"
         EncabezadosAlineacion=   "C-C-C"
         FormatosEdit    =   "0-0-0"
         TextArray0      =   "#"
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   285
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5265
      TabIndex        =   7
      Top             =   5460
      Width           =   960
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6315
      TabIndex        =   8
      Top             =   5460
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
      Height          =   1275
      Left            =   105
      TabIndex        =   10
      Top             =   735
      Width           =   7155
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCliente 
         Height          =   960
         Left            =   105
         TabIndex        =   2
         Top             =   210
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   1693
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "frmCapTarjetaBlqDesBlq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub LimpiaPantalla()
grdCliente.Clear
grdCliente.Rows = 2
SetupGridCliente
grdClienteTarj.Clear
grdClienteTarj.Rows = 2
grdClienteTarj.FormaCabecera
grdTarjetaEstado.Clear
grdTarjetaEstado.Rows = 2
grdTarjetaEstado.FormaCabecera
txtTarjeta.Text = "____-____-____-____"
fraTarjeta.Enabled = True
fraEstado.Enabled = False
cmdCancelar.Enabled = False
cmdGrabar.Enabled = False
txtGlosa = ""
chkBloqueo.value = 0
End Sub

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
Dim nEstado As COMDConstantes.CaptacTarjetaEstado

sTarjeta = Replace(txtTarjeta, "-", "", 1, , vbTextCompare)
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsTarj = clsMant.GetTarjetaCuentas(sTarjeta)
If rsTarj.EOF And rsTarj.BOF Then
    MsgBox "Tarjeta no posee ninguna relación con cuentas activas o Tarjeta no activa.", vbInformation, "Aviso"
    Exit Sub
Else
    Dim nItem As Integer
    sPersona = rsTarj("cPersCod")
    nEstado = rsTarj("nEstado")
    chkBloqueo.value = IIf(nEstado = gCapTarjEstBloqueada, 1, 0)
    Do While Not rsTarj.EOF
        grdClienteTarj.AdicionaFila
        nItem = grdClienteTarj.Rows - 1
        grdClienteTarj.TextMatrix(nItem, 1) = rsTarj("cCtaCod")
        grdClienteTarj.TextMatrix(nItem, 2) = rsTarj("Relacion")
        rsTarj.MoveNext
    Loop
    
    rsTarj.Close
    Set rsTarj = clsMant.GetDatosPersona(sPersona)
    Set grdCliente.Recordset = rsTarj
    SetupGridCliente
    
    rsTarj.Close
    Set rsTarj = clsMant.GetTarjetaEstadoHist(sTarjeta)
    Set grdTarjetaEstado.Recordset = rsTarj
    cmdGrabar.Enabled = True
    cmdCancelar.Enabled = True
    fraTarjeta.Enabled = False
    fraEstado.Enabled = True
    
End If
Set rsTarj = Nothing
Set clsMant = Nothing
End Sub

Private Sub chkBloqueo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtGlosa.SetFocus
End If
End Sub

Private Sub cmdBuscar_Click()
Dim clsPers As COMDPersona.UCOMPersona
Set clsPers = New COMDPersona.UCOMPersona
Set clsPers = frmBuscaPersona.Inicio

If Not clsPers Is Nothing Then
    Dim sPers As String, sTarj As String, sTarjBus
    Dim rsPers As New ADODB.Recordset
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim clsTarj As UCapCuenta
    sPers = clsPers.sPersCod
    Set clsPers = Nothing
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsPers = clsCap.GetPersonaTarj(sPers)
    Set clsCap = Nothing
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
        rsPers.Close
        Set rsPers = Nothing
        Set clsTarj = frmCapMantenimientoCtas.Inicia
    Else
        MsgBox "Persona no posee tarjetas activas.", vbInformation, "Aviso"
        txtTarjeta.SetFocus
        rsPers.Close
        Set rsPers = Nothing
        Exit Sub
    End If
    sTarj = Trim(clsTarj.sCtaCod)
    Set clsTarj = Nothing
    grdClienteTarj.Clear
    grdClienteTarj.FormaCabecera
    grdClienteTarj.Rows = 2
    If sTarj = "" Then
        txtTarjeta.Text = "____-____-____-____"
    Else
        txtTarjeta.Text = Format$(sTarj, "@@@@-@@@@-@@@@-@@@@")
        txtTarjeta.SetFocus
        SendKeys "{Enter}"
    End If
    
Else
    cmdBuscar.SetFocus
End If
End Sub

Private Sub cmdCancelar_Click()
LimpiaPantalla
txtTarjeta.SetFocus
End Sub

Private Sub cmdGrabar_Click()
If MsgBox("¿Desea grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    If Trim(txtGlosa) = "" Then
        MsgBox "Debe colocar la glosa correspondiente.", vbInformation, "Aviso"
        txtGlosa.SetFocus
        Exit Sub
    End If
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim sTarjeta As String, sMovNro As String
    Dim nEstado As COMDConstantes.CaptacTarjetaEstado 'CaptacTarjetaEstado
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    Dim CLSSERV As COMNCaptaServicios.NCOMCaptaServicios
    
    Dim lscadimp As String
    Dim loPrevio As previo.clsPrevio
    
    Set CLSSERV = New COMNCaptaServicios.NCOMCaptaServicios
    
    Set clsMov = New COMNContabilidad.NCOMContFunciones
        sMovNro = clsMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    Set clsMov = Nothing
    
    sTarjeta = Replace(txtTarjeta, "-", "", 1, , vbTextCompare)
        nEstado = IIf(chkBloqueo.value = 1, gCapTarjEstBloqueada, gCapTarjEstActiva)
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    
    If clsMant.ActualizaTarjetaEstado(sTarjeta, sMovNro, nEstado, Trim(txtGlosa)) Then
        lscadimp = CLSSERV.ImprimeBolTarjeta(IIf(nEstado = gCapTarjEstBloqueada, "BLOQUEO DE TARJETA", "DESBLOQUEO DE TARJETA"), _
                                                Trim(grdCliente.TextMatrix(1, 1)), txtTarjeta.Text, _
                                                "TARJEA MAGNETICA", gdFecSis, gsNomAge, _
                                                gsCodUser, sLpt)
        Do
           Set loPrevio = New previo.clsPrevio
                loPrevio.PrintSpool sLpt, lscadimp, False
                loPrevio.PrintSpool sLpt, Chr(10) & Chr(10) & Chr(10) & Chr(10) & lscadimp, False
           Set loPrevio = Nothing
        Loop Until MsgBox("DESEA REIMPRIMIR BOLETA?", vbYesNo, "AVISO") = vbNo
    Else
        cmdCancelar_Click
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
Me.Caption = "Captaciones - Tarjeta - Bloqueo/Desbloqueo"
LimpiaPantalla
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
End If
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

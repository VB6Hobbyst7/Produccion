VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmTarjetaBloqueo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bloqueo de Tarjeta"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   7350
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CMDbLOQUEO 
      Caption         =   "&Bloqueo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
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
      Height          =   1170
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   7155
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCliente 
         Height          =   855
         Left            =   105
         TabIndex        =   4
         Top             =   210
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   1508
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame fraTarjeta 
      Height          =   750
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3405
      Begin MSMask.MaskEdBox txtTarjeta 
         Height          =   375
         Left            =   945
         TabIndex        =   1
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
         TabIndex        =   2
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
      Caption         =   "EL PROCESO DE BLOQUEO DE LA TARJETA  ES IRREVERSIBLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3795
      TabIndex        =   5
      Top             =   135
      Width           =   3495
   End
End
Attribute VB_Name = "frmTarjetaBloqueo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CMDbLOQUEO_Click()
Dim clsMov As COMNContabilidad.NCOMContFunciones 'NContFunciones
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim sMovNro As String
Dim sTarjeta As String
Dim lblTrack1 As String
Dim ObjTarj As COMNCaptaServicios.NCOMCaptaTarjeta

sTarjeta = Trim(Replace(lblTrack1, "-", ""))
If Trim(sTarjeta) = "" Then
    MsgBox "Ingrese el Nùmero de Tarjeta a Bloquear", vbInformation, "AVISO"
    Exit Sub
End If

Set ObjTarj = New COMNCaptaServicios.NCOMCaptaTarjeta
Set clsMov = New COMNContabilidad.NCOMContFunciones
sMovNro = clsMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
Set clsMov = Nothing

If Not ObjTarj.VerificaTarjetaActiva(sTarjeta) Then
    MsgBox "Esta Tarjeta no esta Activa", vbInformation, "AVISO"
    Set ObjTarj = Nothing
    Exit Sub
End If

If MsgBox("Esta Seguro de Realizar el Bloqueo de la Tarjeta, este proceso es irreversible", vbInformation + vbYesNo, "AVISO") = vbNo Then
    Set ObjTarj = Nothing
    Exit Sub
End If
Call ObjTarj.BloqueoTarjeta(sTarjeta, sMovNro)

MsgBox "Tarjeta Bloqueada", vbInformation, "AVISO"
Set ObjTarj = Nothing
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub
Public Sub LimpiaPantalla()

grdCliente.Clear
grdCliente.Rows = 2

SetupGridCliente
'grdTarjetaEstado.Clear
'grdTarjetaEstado.Rows = 2
'grdTarjetaEstado.FormaCabecera
txtTarjeta.Text = "____-____-____-____"
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
Private Sub txtTarjeta_KeyPress(KeyAscii As Integer)
Dim rsTarj As New ADODB.Recordset
Dim ObjTarj As COMNCaptaServicios.NCOMCaptaTarjeta
Dim sPersona As String
Dim cTarjeta As String
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
If KeyAscii = 13 Then
    
    cTarjeta = Trim(Replace(Me.txtTarjeta.Text, "-", ""))
    If Trim(cTarjeta) = "" Then
        MsgBox "Nro de Tarjeta Incorrecta", vbInformation, "AVISO"
        LimpiaPantalla
        Exit Sub
    End If
    Set ObjTarj = New COMNCaptaServicios.NCOMCaptaTarjeta
    If Not ObjTarj.VerificaTarjetaActiva(cTarjeta) Then
        MsgBox "Tarjeta Invalida", vbInformation, "AVISO"
        Set ObjTarj = Nothing
        LimpiaPantalla
        Exit Sub
    End If
    Set rsTarj = New ADODB.Recordset
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    sPersona = ObjTarj.Get_Tarj_Cod_Titular(cTarjeta)
    If Trim(sPersona) <> "" Then
        Set rsTarj = clsMant.GetDatosPersona(sPersona)
        Set grdCliente.Recordset = rsTarj
        Set rsTarj = ObjTarj.Get_Tarj_HistorialEst(cTarjeta)
        'Set grdTarjetaEstado.Recordset = rsTarj
    End If
    Set clsMant = Nothing
    SetupGridCliente
    Set rsTarj = Nothing
    Set ObjTarj = Nothing
    CMDbLOQUEO.Enabled = True
End If
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCapServConvenioOpe 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10965
   Icon            =   "frmCapServConvenioOpe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5775
      TabIndex        =   2
      Top             =   3990
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4575
      TabIndex        =   1
      Top             =   3990
      Width           =   1095
   End
   Begin VB.Frame fraConvenio 
      Caption         =   "Entidades Convenio"
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
      Height          =   3825
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   10710
      Begin SICMACT.FlexEdit grdListado 
         Height          =   2535
         Left            =   75
         TabIndex        =   6
         Top             =   1110
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   4471
         Cols0           =   3
         ScrollBars      =   2
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Codigo-Institucion"
         EncabezadosAnchos=   "400-1600-8200"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X"
         ListaControles  =   "0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L"
         FormatosEdit    =   "0-0-0"
         TextArray0      =   "#"
         SelectionMode   =   1
         lbUltimaInstancia=   -1  'True
         lbOrdenaCol     =   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.OptionButton opt 
         Caption         =   "Colegios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   3630
         TabIndex        =   4
         Top             =   450
         Width           =   1185
      End
      Begin VB.OptionButton opt 
         Caption         =   "Otros Ingresos con Personas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   0
         Left            =   945
         TabIndex        =   3
         Top             =   375
         Value           =   -1  'True
         Width           =   1785
      End
   End
   Begin MSComctlLib.ImageList imglstFiguras 
      Left            =   90
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCapServConvenioOpe.frx":030A
            Key             =   "Padre"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCapServConvenioOpe.frx":065C
            Key             =   "Hijo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCapServConvenioOpe.frx":09AE
            Key             =   "Hijito"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCapServConvenioOpe.frx":0D00
            Key             =   "Bebe"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwConvenio 
      Height          =   2565
      Left            =   345
      TabIndex        =   5
      Top             =   990
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4524
      _Version        =   393217
      LabelEdit       =   1
      Style           =   1
      FullRowSelect   =   -1  'True
      ImageList       =   "imglstFiguras"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCapServConvenioOpe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub EjecutaOperacion(ByVal npersona As CaptacConvenios, ByVal sDescOperacion As String, _
        ByVal sPersCod As String)
'Select Case nPersona
'    Case gCapConvUNT    '**101
'        'frmCapServOpeUNT.Show 1
'         frmCapServConvAbonoInst.Inicia grdListado.TextMatrix(grdListado.Row, 1), grdListado.TextMatrix(grdListado.Row, 2)
'
'    Case gCapConvNarvaez  '**102
'        frmCapServConvAbono.Inicia nPersona
'End Select
End Sub

Private Sub CmdAceptar_Click()
'Dim nodOpe As Node
'Dim sDesc As String, sPersCod As String
'Set nodOpe = tvwConvenio.SelectedItem
'If Not nodOpe Is Nothing Then
'    sDesc = Mid(nodOpe.Text, 15, Len(nodOpe.Text) - 14)
'    sPersCod = Left(nodOpe.Text, 13)
'   EjecutaOperacion CLng(nodOpe.Tag), sDesc, sPersCod
'End If
'Set nodOpe = Nothing
Dim npersona As COMDConstantes.CaptacConvenios

If opt(0).value = True Then
     frmCapServConvAbonoInst.Inicia grdListado.TextMatrix(grdListado.Row, 1), grdListado.TextMatrix(grdListado.Row, 2), grdListado.TextMatrix(grdListado.Row, 3)
Else
     frmCapServConvAbono.Inicia npersona
End If



End Sub

Private Sub cmdsalir_Click()
        Unload Me
End Sub

Private Sub Form_Load()
Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios 'NCapServicios
Dim rsServ As New ADODB.Recordset
Dim nodOpe As Node
Dim sPersCod As String, sOpePadre As String
Dim sNombre As String
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Me.Caption = "Captaciones - Operaciones - Depósitos Entidades Convenios"
Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
Set rsServ = clsServ.GetServConvenios(gCapConvUNT)
Set grdListado.Recordset = rsServ
Set clsServ = Nothing
Set rsServ = Nothing

'Do While Not rsServ.EOF
'    sPersCod = rsServ("nConvCod")
'    sNombre = sPersCod & " - " & UCase(rsServ("cPersNombre"))
'    'sOpePadre = "P" & sPersCod
'    'Set nodOpe = tvwConvenio.Nodes.Add(, , sOpePadre, sNombre, "Padre")
'    'nodOpe.Tag = sPersCod
'    rsServ.MoveNext
'Loop
'rsServ.Close

End Sub

Private Sub Option2_Click()

End Sub

Private Sub opt_Click(Index As Integer)
Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
Dim rsServ As ADODB.Recordset
Dim nodOpe As Node
Dim sPersCod As String, sOpePadre As String
Dim sNombre As String
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Me.Caption = "Captaciones - Operaciones - Depósitos Entidades Convenios"
Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios


Select Case Index
    Case 0
            Set rsServ = clsServ.GetServConvenios(gCapConvUNT)
    Case 1
            Set rsServ = clsServ.GetServConvenios(gCapConvNarvaez)
End Select

Set grdListado.Recordset = rsServ
Set clsServ = Nothing
Set rsServ = Nothing


End Sub

Private Sub tvwConvenio_DblClick()
Dim nodOpe As Node
Dim sDesc As String, sPersCod As String
Set nodOpe = tvwConvenio.SelectedItem
If Not nodOpe Is Nothing Then
    sDesc = Mid(nodOpe.Text, 15, Len(nodOpe.Text) - 14)
    sPersCod = Left(nodOpe.Text, 13)
    EjecutaOperacion CLng(nodOpe.Tag), sDesc, sPersCod
End If
Set nodOpe = Nothing
End Sub

Private Sub tvwConvenio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim nodOpe As Node
    Dim sDesc As String, sPersCod As String
    Set nodOpe = tvwConvenio.SelectedItem
    If Not nodOpe Is Nothing Then
        sDesc = Mid(nodOpe.Text, 15, Len(nodOpe.Text) - 14)
        sPersCod = Left(nodOpe.Text, 13)
        EjecutaOperacion CLng(nodOpe.Tag), sDesc, sPersCod
    End If
    Set nodOpe = Nothing
End If
End Sub



VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCFHojasRecepcion 
   Caption         =   "Hojas CF: Recepcionar"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTCartasFolios 
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Cartas Fianza"
      TabPicture(0)   =   "frmCFHojasRecepcion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FeRemesas"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdConfirmar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCerrar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdActualizar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "A&ctualizar"
         Height          =   375
         Left            =   3480
         TabIndex        =   0
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   6120
         TabIndex        =   2
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   4800
         TabIndex        =   1
         Top             =   2880
         Width           =   1095
      End
      Begin SICMACT.FlexEdit FeRemesas 
         Height          =   2115
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   3731
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Agencia-Fecha Envío-Desde-Hasta-Sin Folio-CodEnvio-Estado"
         EncabezadosAnchos=   "400-2000-1200-1000-1000-1200-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remesas pendientes de Confirmación"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   2670
      End
   End
End
Attribute VB_Name = "frmCFHojasRecepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdActualizar_Click()
LlenarGridRemesas
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdConfirmar_Click()
Dim oCartaFianza As COMNCartaFianza.NCOMCartaFianzaValida
Set oCartaFianza = New COMNCartaFianza.NCOMCartaFianzaValida
If Trim(Me.FeRemesas.TextMatrix(Me.FeRemesas.Row, 6)) = "" Then
    MsgBox "No a selecionado.", vbInformation, "Aviso"
Else
    If MsgBox("Esta seguro de confirmar la recepción de las Hojas de CF?", vbInformation + vbYesNo, "Remesar Folios de CF") = vbYes Then
        Call oCartaFianza.ConfirmarEnvioFolios(Trim(Me.FeRemesas.TextMatrix(Me.FeRemesas.Row, 6)), gdFecSis)
        Call oCartaFianza.ActualizarEnvioFolios(Trim(Me.FeRemesas.TextMatrix(Me.FeRemesas.Row, 6)), 1)
        LlenarGridRemesas
    End If
End If
End Sub

Private Sub Form_Load()
Call CentraForm(Me)
Me.Icon = LoadPicture(App.path & gsRutaIcono)
LlenarGridRemesas
End Sub

Private Sub LlenarGridRemesas()
Dim oCartaFianza As COMNCartaFianza.NCOMCartaFianzaValida
Dim rsCartaFianza As ADODB.Recordset
Dim i As Integer
Set oCartaFianza = New COMNCartaFianza.NCOMCartaFianzaValida
Set rsCartaFianza = oCartaFianza.ObtenerEnvioFolios("0", gsCodAge)

Call LimpiaFlex(FeRemesas)
If rsCartaFianza.RecordCount > 0 Then
    If Not (rsCartaFianza.EOF Or rsCartaFianza.BOF) Then
        For i = 0 To rsCartaFianza.RecordCount - 1
            FeRemesas.AdicionaFila
            Me.FeRemesas.TextMatrix(i + 1, 0) = i + 1
            Me.FeRemesas.TextMatrix(i + 1, 1) = rsCartaFianza!cAgeDescripcion
            Me.FeRemesas.TextMatrix(i + 1, 2) = rsCartaFianza!dFechaEnvio
            Me.FeRemesas.TextMatrix(i + 1, 3) = rsCartaFianza!nDesde
            Me.FeRemesas.TextMatrix(i + 1, 4) = rsCartaFianza!nHasta
            Me.FeRemesas.TextMatrix(i + 1, 5) = rsCartaFianza!nSFolio
            Me.FeRemesas.TextMatrix(i + 1, 6) = rsCartaFianza!nCodEnvio
            Me.FeRemesas.TextMatrix(i + 1, 7) = rsCartaFianza!nEstado
            rsCartaFianza.MoveNext
        Next i
    End If
End If
End Sub


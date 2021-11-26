VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCFHojasConsultarFolio 
   Caption         =   "Folios de Cartas Fianza Por Envio"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTCartasFolios 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Cartas Fianza"
      TabPicture(0)   =   "frmCFHojasConsultarFolio.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FeRemesas"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCerrar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   8160
         TabIndex        =   0
         Top             =   5040
         Width           =   1095
      End
      Begin SICMACT.FlexEdit FeRemesas 
         Height          =   4155
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   9165
         _extentx        =   16166
         _extenty        =   7329
         cols0           =   7
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Agencia-Nº Folio-Nº Cta. Fianza-Fecha-Estado-CodEnvio"
         encabezadosanchos=   "400-2000-1300-2000-1200-1600-0"
         font            =   "frmCFHojasConsultarFolio.frx":001C
         font            =   "frmCFHojasConsultarFolio.frx":0048
         font            =   "frmCFHojasConsultarFolio.frx":0074
         font            =   "frmCFHojasConsultarFolio.frx":00A0
         fontfixed       =   "frmCFHojasConsultarFolio.frx":00CC
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-C-C-C-C-C"
         formatosedit    =   "0-0-0-0-0-0-0"
         textarray0      =   "#"
         colwidth0       =   405
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Folios Encontrados:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmCFHojasConsultarFolio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fsCodEnvio As String
Private Sub cmdCerrar_Click()
Unload Me
End Sub
Private Sub LlenarGridRemesas()
Dim oCartaFianza As COMNCartaFianza.NCOMCartaFianzaValida
Dim rsCartaFianza As ADODB.Recordset
Dim i As Integer
Set oCartaFianza = New COMNCartaFianza.NCOMCartaFianzaValida
Set rsCartaFianza = oCartaFianza.ConsultarEnviosFolios(2, , , , , fsCodEnvio)

Call LimpiaFlex(FeRemesas)
If rsCartaFianza.RecordCount > 0 Then
    If Not (rsCartaFianza.EOF Or rsCartaFianza.BOF) Then
        For i = 0 To rsCartaFianza.RecordCount - 1
            FeRemesas.AdicionaFila
            Me.FeRemesas.TextMatrix(i + 1, 0) = i + 1
            Me.FeRemesas.TextMatrix(i + 1, 1) = rsCartaFianza!cAgeDescripcion
            Me.FeRemesas.TextMatrix(i + 1, 2) = rsCartaFianza!nNumFolio
            Me.FeRemesas.TextMatrix(i + 1, 3) = rsCartaFianza!cCtaCod
            Me.FeRemesas.TextMatrix(i + 1, 4) = rsCartaFianza!dEstado
            Me.FeRemesas.TextMatrix(i + 1, 5) = rsCartaFianza!Estado
            Me.FeRemesas.TextMatrix(i + 1, 6) = rsCartaFianza!nCodEnvio
            rsCartaFianza.MoveNext
        Next i
    End If
Else
    MsgBox "No exiten datos a mostrar.", vbInformation, "Aviso"
End If
End Sub

Public Sub Inicio(ByVal psCodEnvio As String)
fsCodEnvio = psCodEnvio
Me.Show 1
End Sub

Private Sub Form_Load()
Call CentraForm(Me)
Me.Icon = LoadPicture(App.path & gsRutaIcono)
LlenarGridRemesas
End Sub

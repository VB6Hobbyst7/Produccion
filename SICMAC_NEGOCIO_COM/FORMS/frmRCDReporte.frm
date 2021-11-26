VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2E37A9B8-9906-4590-9F3A-E67DB58122F5}#7.0#0"; "OcxLabelX.ocx"
Begin VB.Form frmRCDReporte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes RCD"
   ClientHeight    =   3975
   ClientLeft      =   5055
   ClientTop       =   3630
   ClientWidth     =   4440
   Icon            =   "frmRCDReporte.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   360
      Left            =   525
      TabIndex        =   6
      Top             =   4380
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame FrameOperaciones 
      Caption         =   "Lista de Reportes"
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4215
      Begin OcxLabelX.LabelX LabelX 
         Height          =   555
         Left            =   975
         TabIndex        =   5
         Top             =   2235
         Visible         =   0   'False
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   979
         FondoBlanco     =   0   'False
         Resalte         =   0
         Caption         =   "Procesando...!"
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin MSComctlLib.TreeView tvwReporte 
         Height          =   2775
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4895
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imglstFiguras"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList imglstFiguras 
         Left            =   240
         Top             =   4320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRCDReporte.frx":030A
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRCDReporte.frx":0624
               Key             =   "Hijo"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fracontroles1 
      Height          =   765
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   4245
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmRCDReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fnRepoSelec As Long
Dim Progress As clsProgressBar
'Dim WithEvents loRep As nRcdReportes
Dim loRep As COMNCredito.NCOMRCD

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
Dim lsCad As String
Dim lsServerCons As String

Dim sImpresion As String
Dim sMensaje As String
'Dim oPrevio As Previo.clsPrevio
Dim NumeroArchivo As Integer
Dim lsNombreArchivo  As String

If fnRepoSelec = "179100" Then Exit Sub

fracontroles1.Enabled = False
LabelX.Visible = True

Set loRep = New COMNCredito.NCOMRCD
    lsServerCons = loRep.GetServerConsol

Select Case fnRepoSelec
    Case 179101: sImpresion = loRep.nRepo179101_ArchivoRCD(lsServerCons, gdFecDataFM, sMensaje)
    Case 179102: 'Call loRep.nRepo_IBM_Excel(lsServerCons, gdFecDataFM, gsNomCmac, gdFecDataFM)
    'Case 179103: sImpresion = loRep.nRepo179103_IBM(lsServerCons, gdFecSis, gsNomCmac, gsCodCMAC)
    'Case 179104: Call loRep.nRepo179104_ColocacionFoncodes(lsServerCons, gsNomCmac, gdFecDataFM)
    'Case 179105: Call loRep.nRepo179105_CarteraMensualFoncodes(gsNomCmac, gdFecDataFM)
End Select

If sMensaje <> "" Then
    MsgBox sMensaje, vbInformation, "Mensaje"
    Exit Sub
End If

'Se debe enviar a un Archivo...El Previo fue solo para Pruebas
'Set oPrevio = New Previo.clsPrevio
'oPrevio.Show sImpresion, "Reporte 179101 de Archivo RCD"
'Set oPrevio = Nothing

Select Case fnRepoSelec
    Case 179101:
        NumeroArchivo = FreeFile
        lsNombreArchivo = App.path & "\Spooler\RCD" & Format(gdFecSis, "yyyymm") & ".106" '".108" ' 106 CUSCO
        Open lsNombreArchivo For Output As #NumeroArchivo
        Print #NumeroArchivo, sImpresion
        Close #NumeroArchivo
        MsgBox "Se ha generado el Archivo RCDvc00.106 Satisfactoriamente, Termino : " & Time(), vbInformation, "Mensaje"
        
    Case 179102:
    Case 179103:
        NumeroArchivo = FreeFile
        lsNombreArchivo = App.path & "\Spooler\ICS01.112"
        Open lsNombreArchivo For Output As #NumeroArchivo
        Print #NumeroArchivo, sImpresion
        Close #NumeroArchivo
        MsgBox "Reporte Generado Satisfactoriamente en " & lsNombreArchivo, vbInformation, "Aviso"
    'Case 179104: Call loRep.nRepo179104_ColocacionFoncodes(lsServerCons, gsNomCmac, gdFecDataFM)
    'Case 179105: Call loRep.nRepo179105_CarteraMensualFoncodes(gsNomCmac, gdFecDataFM)
End Select


fracontroles1.Enabled = True
LabelX.Visible = False

End Sub

'Private Sub Command1_Click()
'Dim SQL As String
'Dim oCon As DConecta
'Dim rs As ADODB.Recordset
'Dim lsNombreArchivo As String
'Dim NúmeroArchivo As Integer
'
'SQL = "SELECT ccodsbs, isnull(cperscod,'') as cCodUni, canio, cmes, ctipopers, isnull(ctipodoc,'') ctipodoc , isnull(cnumdoc,'') as cnumdoc, isnull(cnroruc,'') as cnroruc, " _
'    & " isnull(capepat,'') as capepat , isnull(capemat,'') as capemat, isnull(capecas,'') as capecas, isnull(cprinom,'') as cprinom, isnull(nSegNom,'') as nSegNom  " _
'    & " FROM DBRCC..INFOCLISBS"
'
'Set oCon = New DConecta
'oCon.AbreConexion
'Set rs = oCon.CargaRecordSet(SQL)
'oCon.CierraConexion
'
'NúmeroArchivo = FreeFile
'lsNombreArchivo = "C:\APL\INFOCMACICA1.TXT"
'Open lsNombreArchivo For Output As #NúmeroArchivo
'Do While Not rs.EOF
'    'Print #NúmeroArchivo, ImpreFormat(rs!ccodsbs, 10, 0) _
'                        & ImpreFormat(rs!ccoduni, 20, 0) _
'                        & ImpreFormat(rs!cAnio, 4, 0) _
'                        & ImpreFormat(Format(rs!cmes, "00"), 2, 0) _
'                        & ImpreFormat(rs!ctipopers, 1, 0) _
'                        & ImpreFormat(rs!ctipodoc, 3, 0) _
'                        & ImpreFormat(rs!cnumdoc, 15, 0) _
'                        & ImpreFormat(rs!cnroruc, 11, 0) _
'                        & ImpreFormat(rs!capepat, 120, 0) _
'                        & ImpreFormat(rs!capemat, 40, 0) _
'                        & ImpreFormat(rs!capecas, 40, 0) _
'                        & ImpreFormat(rs!cprinom, 40, 0) _
'                        & ImpreFormat(rs!nSegNom, 40, 0) _
'                        & Chr(13) & Chr(10);
'    Print #NúmeroArchivo, ImpreFormat(rs!ccodsbs, 10, 0) _
'                        & ImpreFormat(rs!ccoduni, 20, 0) _
'                        & Chr(13) & Chr(10);
'    rs.MoveNext
'Loop
'
'Close #NúmeroArchivo   ' Cierra el archivo.
'MsgBox "Se ha generado el Archivo RCDvc00.108 Satisfactoriamente, Termino : " & Time()
'
'End Sub

Private Sub Form_Load()
    Set Progress = New clsProgressBar
    CargaMenu
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub CargaMenu()
Dim clsGen As DGeneral 'COMDConstSistema.DCOMGeneral
Dim rsUsu As Recordset
Dim sOperacion As String
Dim sOpeCod As String
Dim sOpePadre As String
Dim sOpeHijo As String
Dim sOpeHijito As String
Dim nodOpe As Node
Dim lsTipREP As String
lsTipREP = "1791"
Set clsGen = New DGeneral 'COMDConstSistema.DCOMGeneral
'ARCV 20-07-2006
'Set rsUsu = clsGen.GetOperacionesUsuario(gsCodUser, lsTipREP, MatOperac, NroRegOpe)
Set rsUsu = clsGen.GetOperacionesUsuario_NEW(lsTipREP, , gRsOpeRepo)

Set clsGen = Nothing
Do While Not rsUsu.EOF
    sOpeCod = rsUsu("cOpeCod")
    sOperacion = sOpeCod & " - " & UCase(rsUsu("cOpeDesc"))
    Select Case rsUsu("nOpeNiv")
        Case "1"
            sOpePadre = "P" & sOpeCod
         Set nodOpe = tvwReporte.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
            nodOpe.Tag = sOpeCod
        Case "2"
            sOpeHijo = "H" & sOpeCod
            Set nodOpe = tvwReporte.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
            nodOpe.Tag = sOpeCod
    End Select
    rsUsu.MoveNext
Loop
rsUsu.Close
Set rsUsu = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set loRep = Nothing
End Sub

'Private Sub loRep_CloseProgress()
'Progress.CloseForm Me
'End Sub

'Private Sub loRep_Progress(pnValor As Long, pnTotal As Long)
' Progress.Max = pnTotal
' Progress.Progress pnValor, "Generando Reporte"
' DoEvents
'End Sub

'Private Sub loRep_ShowProgress()
'Progress.ShowForm Me
'End Sub

Private Sub tvwReporte_Click()
Dim NodRep  As Node
Dim lsDesc As String
Set NodRep = tvwReporte.SelectedItem
If NodRep Is Nothing Then
   Exit Sub
End If
lsDesc = Mid(NodRep.Text, 8, Len(NodRep.Text) - 7)
fnRepoSelec = CLng(NodRep.Tag)
End Sub

Private Sub tvwReporte_DblClick()
    Call cmdImprimir_Click
End Sub

Private Sub tvwReporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdImprimir_Click
End Sub

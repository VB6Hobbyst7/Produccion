VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRCDReporte_NEW 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes RCD"
   ClientHeight    =   4620
   ClientLeft      =   5055
   ClientTop       =   3630
   ClientWidth     =   4890
   Icon            =   "frmRCDReporte_NEW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExportar 
      Cancel          =   -1  'True
      Caption         =   "&Exportar"
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
      Left            =   720
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.Frame FrameOperaciones 
      Caption         =   "Lista de Reportes"
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4695
      Begin VB.PictureBox LabelX 
         Height          =   555
         Left            =   720
         ScaleHeight     =   495
         ScaleWidth      =   3300
         TabIndex        =   5
         Top             =   1920
         Visible         =   0   'False
         Width           =   3360
      End
      Begin MSComctlLib.TreeView tvwReporte 
         Height          =   2775
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
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
               Picture         =   "frmRCDReporte_NEW.frx":030A
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRCDReporte_NEW.frx":0624
               Key             =   "Hijo"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fracontroles1 
      Height          =   765
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   4605
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
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   1300
      End
      Begin VB.CommandButton cmdCancelar 
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
         Width           =   1300
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   360
      Left            =   480
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblAvance 
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   4095
   End
End
Attribute VB_Name = "frmRCDReporte_NEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fnRepoSelec As Long
Dim Progress As clsProgressBar
'Dim WithEvents loRep As nRcdReportes
Dim lsArchivo As String 'LUCV20161007 'LUCV20170415
Dim loRep As COMNCredito.NCOMRCD

Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub cmdExportar_Click() 'LUCV20170415, Agregó

    Dim lsCad As String
    Dim lsServerCons As String
    
    Dim sImpresion As String
    Dim sMensaje As String
    
    Dim NumeroArchivo As Integer
    Dim lsNombreArchivo  As String
    
    If fnRepoSelec = "179100" Then Exit Sub
    
    fracontroles1.Enabled = False
    LabelX.Visible = True
    
    Set loRep = New COMNCredito.NCOMRCD
        lsServerCons = loRep.GetServerConsol
    
    Select Case fnRepoSelec
        Case 179101
            GeneraRptExcellRCD
        Case 179104
            ReporteArchivoRCA
        Case 179105
            ReporteArchivoRCM
        Case 179106 'JUEZ 20150113
            ReporteVerifRCDconRCC
        Case 179107 'JUEZ 20150310
            ReporteArchivoRCD True
    End Select
    
    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, "Mensaje"
        Exit Sub
    End If
    
    fracontroles1.Enabled = True
    LabelX.Visible = False
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
    Case 179101
        'ReporteArchivoRCD
        'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
        ReporteArchivoRCD 0
    'Case 179102 'Call loRep.nRepo_IBM_Excel(lsServerCons, gdFecDataFM, gsNomCmac, gdFecDataFM)
    Case 179104
        ReporteArchivoRCA
    Case 179105
        ReporteArchivoRCM
    Case 179106 'JUEZ 20150113
        ReporteVerifRCDconRCC
    Case 179107 'JUEZ 20150310
        'ReporteArchivoRCD True
        'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
        ReporteArchivoRCD 1
    'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
    Case 179108
        ReporteArchivoRCD 2
     'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
     
End Select

If sMensaje <> "" Then
    MsgBox sMensaje, vbInformation, "Mensaje"
    Exit Sub
End If

fracontroles1.Enabled = True
LabelX.Visible = False

End Sub

Private Sub Form_Load()
    Set Progress = New clsProgressBar
    CargaMenu
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub

Private Sub CargaMenu()
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim rsUsu As Recordset
Dim sOperacion As String
Dim sOpeCod As String
Dim sOpePadre As String
Dim sOpeHijo As String
Dim sOpeHijito As String
Dim nodOpe As Node
Dim lsTipREP As String
lsTipREP = "1791"
Set clsGen = New COMDConstSistema.DCOMGeneral
'Set rsUsu = clsGen.GetOperacionesUsuario(gsCodUser, lsTipREP, MatOperac, NroRegOpe)
'Set rsUsu = clsGen.GetOperacionesUsuario_NEW(lsTipREP, , gRsOpeRepo)
Set rsUsu = GetOperacionesUsuarioRCD(lsTipREP, , gRsOpeRepo)

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

Private Function GetOperacionesUsuarioRCD(ByVal sProducto As String, _
        Optional nMoneda As Moneda = 0, Optional ByVal prsOpeRep As ADODB.Recordset) As ADODB.Recordset

Dim ssql As String
Dim sFiltroMon As String
Dim rstemp As ADODB.Recordset

If nMoneda > 0 Then
    sFiltroMon = " AND O.cOpeCod NOT like '__" & Trim(nMoneda) & "%'"
End If

prsOpeRep.MoveFirst
Set rstemp = prsOpeRep.Clone
rstemp.Filter = "cOpeCod LIKE '" & sProducto & "%'" & sFiltroMon

Set GetOperacionesUsuarioRCD = rstemp
Set rstemp = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set loRep = Nothing
End Sub

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

'** GENERA EL ARCHIVO RCD *************************
'Private Sub ReporteArchivoRCD(Optional ByVal pbRCDTransf As Boolean = False)
'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
Private Sub ReporteArchivoRCD(Optional ByVal pbRCDTransf As Integer)
'JUEZ 20150310 Se agregó pbRCDTransf
Dim SQL1 As String
Dim contTotal As Long
Dim Cont As Long

Dim rs1 As New ADODB.Recordset

Dim NúmeroArchivo As Integer
Dim lsNombreArchivo  As String
Dim J As Long
Dim lscTidoci As String
Dim lscNudoCi As String
Dim lscTidoTr As String
Dim lscNudotr As String
Dim lsActividad As String
Dim lsCodSbs As String
Dim lsCalificacion As String
Dim lsCodPers As String
Dim lsNomTabla As String 'JUEZ 20150310
Dim lsNumEntidad As String 'JUEZ 20150310

NúmeroArchivo = FreeFile

Dim lsPersonaAct As String
Dim lnSaldo As Currency

Dim lsApePat As String, lsApeMat As String, lsApeCasada As String
Dim lsNomPri As String, lsNomSeg As String

Dim lsRiesCambiarioPersona As String
Dim lsCondEspCta As String
lsRiesCambiarioPersona = "0"

Dim lsIndicadorAtr As String
Dim lsCalifInterna As String
       
Dim gsCodEmpInfICC As String
Dim gnMontoMinimoICC As Integer
Dim lsCadMancomuno As String
Dim lsCadAval As String
Dim lsNumSecICC As String
Dim oConex As New COMConecta.DCOMConecta
Dim lsConsol As String
Dim lsFactorConversion As String
lsConsol = "dbconsolidada.."
oConex.AbreConexion

'lsNomTabla = IIf(pbRCDTransf, "RCDTvc", "RCDvc") 'JUEZ 20150310
'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
Select Case pbRCDTransf
    Case "1"
        lsNomTabla = "RCDTvc"
    Case "0"
        lsNomTabla = "RCDvc"
    Case "2"
        lsNomTabla = "RCDHvc"
    End Select

Select Case pbRCDTransf
    Case "1"
        lsNumEntidad = "399"
    Case "0"
        lsNumEntidad = "109"
    Case "2"
        lsNumEntidad = "583"
    End Select
'lsNumEntidad = IIf(pbRCDTransf, "399", "109") 'JUEZ 20150310
'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc

SQL1 = "select * from constsistema where nConsSisCod=80"
Set rs1 = oConex.CargaRecordSet(SQL1)
gsCodEmpInfICC = rs1!nConsSisValor

'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
'If pbRCDTransf Then gsCodEmpInfICC = "399"
Select Case pbRCDTransf
    Case "1"
        gsCodEmpInfICC = "399"
    Case "0"
        gsCodEmpInfICC = "109"
    Case "2"
        gsCodEmpInfICC = "583"
    End Select

SQL1 = "select top 1 * from dbconsolidada..rcdparametro  order by  cMEs desc"
Set rs1 = oConex.CargaRecordSet(SQL1)
gnMontoMinimoICC = rs1!nMontoMin


' Crea el nombre de archivo.
'EstatuBarICC.Panels(1).Text = "Generación de Reporte ICC en Archivo"
'lsNombreArchivo = App.path & "\Spooler\RCD" & Format(gdFecData, "yyyymm") & ".109"
lsNombreArchivo = App.Path & "\Spooler\RCD" & Format(gdFecData, "yyyymm") & "." & lsNumEntidad 'JUEZ 20150310
Open lsNombreArchivo For Output As #NúmeroArchivo


'**** Registro CABECERA
'ARCV 16-06-2007
Dim lsMontoMinimo As String
lsMontoMinimo = Trim(ImpreFormat(EliminaPunto(IIf(IsNull(gnMontoMinimoICC), 0, gnMontoMinimoICC)), 15, 0, False))
'------

Print #NúmeroArchivo, "0106" & "01" & FillNum(Trim(gsCodEmpInfICC), 5, "0") & Format(gdFecData, "yyyymmdd") & "012" _
        & Space(15) & String(15 - Len(lsMontoMinimo), "0") & lsMontoMinimo & Chr(13) & Chr(10);

'JUEZ 20150310 Se acambió "RCDvc" por lsNomTabla
SQL1 = " Select COUNT(r1.cPersCod) AS TOTAL From dbconsolidada.." & lsNomTabla & Format(gdFecData, "yyyymm") & "01 R1 " _
     & " Inner Join dbconsolidada.." & lsNomTabla & Format(gdFecData, "yyyymm") & "02 R2 ON R2.cPersCod = R1.cPersCod "

Set rs1 = oConex.CargaRecordSet(SQL1, adLockReadOnly)
If Not RSVacio(rs1) Then
    contTotal = IIf(IsNull(rs1!Total), 0, rs1!Total)
End If
rs1.Close
Set rs1 = Nothing

'JIPR20210408 RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
' If pbRCDTransf Then
'    SQL1 = "stp_sel_GeneraArchivoRCDTReporte '" & Format(gdFecData, "yyyymm") & "'" 'Agregado por LUCV20170227
' Else
'    SQL1 = "stp_sel_GeneraArchivoRCDReporte '" & Format(gdFecData, "yyyymm") & "'" 'Agregado por LUCV20170227
' End If
 
 If pbRCDTransf = 1 Then
    SQL1 = "stp_sel_GeneraArchivoRCDTReporte '" & Format(gdFecData, "yyyymm") & "'"
 ElseIf pbRCDTransf = 0 Then
    SQL1 = "stp_sel_GeneraArchivoRCDReporte '" & Format(gdFecData, "yyyymm") & "'"
 Else
    SQL1 = "stp_sel_GeneraArchivoRCDHReporte '" & Format(gdFecData, "yyyymm") & "'"
 End If
 
 
Set rs1 = oConex.CargaRecordSet(SQL1)

Cont = 0: J = 0
Screen.MousePointer = 11
If Not RSVacio(rs1) Then

    Do While Not rs1.EOF
        J = J + 1

        If lsPersonaAct <> rs1!cCodPers Then
            Cont = Cont + 1
            
            lsNumSecICC = rs1!cNumSec
            
            lsApePat = fgReemplazaCaracterEspecial(rs1!cApePat)
            lsApeMat = fgReemplazaCaracterEspecial(rs1!capemat)
            lsApeCasada = fgReemplazaCaracterEspecial(rs1!cApeCasada)
            lsNomPri = fgReemplazaCaracterEspecial(rs1!cNomPri)
            lsNomSeg = fgReemplazaCaracterEspecial(rs1!cNomSeg)
            
            Select Case Trim(rs1!cTipPers)
                Case "1", "3"
                    lscTidoci = IIf(IsNull(rs1!ctidoci), "", IIf(Len(Trim(rs1!ctidoci)) = 0, "", Trim(rs1!ctidoci)))
                    lscNudoCi = IIf(IsNull(rs1!cnudoci), "", IIf(Len(Trim(rs1!cnudoci)) = 0, "", Trim(rs1!cnudoci)))
                    lscTidoTr = ""
                    lscNudotr = ""
                Case Else
                    lscTidoci = ""
                    lscNudoCi = ""
                    lscTidoTr = IIf(IsNull(rs1!cTidoTr), "", IIf(Len(Trim(rs1!cTidoTr)) = 0, "", Trim(rs1!cTidoTr)))
                    lscNudotr = IIf(IsNull(rs1!cNudoTr), "", IIf(Len(Trim(rs1!cNudoTr)) = 0, "", Trim(rs1!cNudoTr)))
                    ' RUC DE 11 DIGITOS
                    lscTidoTr = IIf(lscTidoTr = "2", "3", lscTidoTr)
            End Select
            If Not IsNull(rs1!ccodsbs) Then
                lsCodSbs = Trim(rs1!ccodsbs)
            Else
                lsCodSbs = ""
            End If
            If Not IsNull(rs1!cActEcon) Then
                If Len(Trim(rs1!cActEcon)) > 0 Then
                    lsActividad = FillNum(rs1!cActEcon, 4, "0")
                Else
                    lsActividad = ""
                End If
            Else
                lsActividad = ""
            End If
            
            'LsCodPers = EmiteCodigoPersona(rs1!cCodPers)
            'If Len(Trim(lsCodPers)) = 0 Then
                lsCodPers = Trim(rs1!cCodPers)
            'End If
            
            lsCalificacion = Trim(rs1!cCalifica)
            lsRiesCambiarioPersona = Trim(rs1!cRCambiario)
            
            lsIndicadorAtr = Trim(rs1!cIndicadorAtraso)
            lsCalifInterna = Trim(IIf(IsNull(rs1!cCalifInterna), "X", rs1!cCalifInterna))
            If Trim(lsCalifInterna) = "X" Then
                lsCalifInterna = "XXXXX"
            End If
            
    '**ALPA 20100803,
      Dim cMagEmp As String
        'JACA 20110428************************************************************
        If Trim(rs1!cMagEmp) = "0" Then
            cMagEmp = "1"
        ElseIf Trim(rs1!cMagEmp) = "1" Then
            cMagEmp = "6"
        ElseIf Trim(rs1!cMagEmp) = "2" Then
            cMagEmp = "7"
        ElseIf Trim(rs1!cMagEmp) = "3" Then
            cMagEmp = "8"
        ElseIf Trim(rs1!cMagEmp) = "4" Then
            cMagEmp = "0"
        ElseIf Trim(rs1!cMagEmp) = "5" Then
            cMagEmp = "5"
        End If
 
            Print #NúmeroArchivo, rs1!cTipoFor & rs1!cTipoInf & FillNum(Trim(lsNumSecICC), 8, " ") & _
                    ImpreFormat(IIf(Len(Trim(lsCodSbs)) = 0, "0000000000", lsCodSbs), 10, 0) & _
                    ImpreFormat(Trim(rs1!cCodPers), 20, 0) & FillNum(lsActividad, 4, " ") & _
                    FillNum(Trim(IIf(IsNull(rs1!ccodregpub), "", rs1!ccodregpub)), 15, " ") & _
                    FillNum(IIf(Len(Trim(lscTidoTr)) = 0, "", lscTidoTr), 1, " ") & _
                    FillNum(IIf(Len(Trim(lscNudotr)) = 0, "", lscNudotr), 11, " ") & _
                    FillNum(IIf(Len(Trim(lscTidoci)) = 0, "", lscTidoci), 1, " ") & _
                    ImpreFormat(IIf(Len(Trim(lscNudoCi)) = 0, "", lscNudoCi), 12, 0) & _
                    FillNum(Trim(IIf(IsNull(rs1!cTipPers), "", rs1!cTipPers)), 1, " ") & FillNum(Trim(IIf(IsNull(rs1!cResid), "", rs1!cResid)), 1, " ") & _
                    FillNum(Trim(lsCalificacion), 1, " ") & _
                    FillNum(Trim(IIf(IsNull(cMagEmp), "", cMagEmp)), 1, " ") & _
                    FillNum(Trim(IIf(IsNull(rs1!cAccionista), "", rs1!cAccionista)), 1, " ") & FillNum(Trim(IIf(IsNull(rs1!cRelInst), "", rs1!cRelInst)), 1, " ") & _
                    FillNum(Trim(rs1!cPaisNac), 4, "  ") & FillNum(Trim(rs1!cGenero), 1, " ") & _
                    FillNum(Trim(rs1!cEstadoCiv), 1, " ") & _
                    ImpreFormat(Trim(IIf(IsNull(rs1!cSiglas), "", rs1!cSiglas)), 20, 0) & _
                    ImpreFormat(Trim(lsApePat), 120, 0) & ImpreFormat(Trim(lsApeMat), 40, 0) & _
                    ImpreFormat(Trim(lsApeCasada), 40, 0) & ImpreFormat(Trim(lsNomPri), 40, 0) & _
                    ImpreFormat(Trim(lsNomSeg), 40, 0) & _
                    FillNum(Trim(lsRiesCambiarioPersona), 1, " ") & FillNum(Trim(lsIndicadorAtr), 1, " ") & FillNum(lsCalifInterna, 5, " ") & _
                    FillNum(Trim(rs1!cCaliSinAlinea), 1, " ") & FillNum(Trim(rs1!cPersCodGrupEco), 20, " ") & _
                    ImpreFormat(IIf(Trim(IIf(IsNull(rs1!cTipPers), "1", rs1!cTipPers)) = "1", Trim(rs1!dFecNac), "        "), 8, 0) & _
                    FillNum(Trim(rs1!cTiDociComp), 2, "  ") & FillNum(Trim(rs1!cNuDociComp), IIf(Trim(rs1!cTiDociComp) = "05", 11, 12), IIf(Trim(rs1!cTiDociComp) = "05", "            ", "           ")) & IIf(Trim(rs1!cTiDociComp) = "05", " ", "") & Space(28) & FillNum(Trim(rs1!cCondEndeuda), 1, " ") & _
                    FillNum(Trim(rs1!cConPagCredRev), 1, " ") & FillNum(Trim(rs1!cConLineaRev), 1, " ") & _
                    ImpreFormat("", 2, 0) _
                    ; Chr(13) & Chr(10);
                    'NAGL 202102 Según ACTA N°010-2021 rs1!cConPagCredRev, rs1!cConLineaRev
                    'JUEZ 20130809 rs1!cCondEndeuda
                    'LUCV20170717, Modificó 14 = 2: ImpreFormat("", 14, 0) / Según Observación SBS
        End If
        
'        If Val(rs1!cCondEspCta) <= 3 Then
            lsCondEspCta = FillNum(Trim(rs1!cCondEspCta), 2, " ")
'        Else
'            Select Case lsRiesCambiarioPersona  ' 2006/03/15 susy
'                Case "0"
'                    lsCondEspCta = "13"
'                Case "1"
'                    lsCondEspCta = "11"
'                Case "2"
'                    lsCondEspCta = "12"
'            End Select
'        End If
lsFactorConversion = IIf(Mid(rs1!cCtaCnt, 1, 2) = "71", "02", "99")
        Print #NúmeroArchivo, rs1!cTipoFor2 & rs1!cTipoInf2 & FillNum(Trim(lsNumSecICC), 8, " ") & _
                FillNum(Trim(rs1!cCodAge), 4, "0") & FillNum(Trim(rs1!cUbicGeo), 6, " ") & _
                FormatoCtaContable(rs1!cCtaCnt) & FillNum(Trim(rs1!cTipoCred), 2, "0") & _
                ImpreFormat(EliminaPunto(IIf(IsNull(rs1!nSaldo), 0, rs1!nSaldo)), 18, 0, False) & _
                FillNum(Trim(rs1!nCondDias), 4, " ") & FillNum(Trim(lsCondEspCta), 2, " ") & _
                FillNum(Trim(rs1!cCondDisponib), 2, " ") & FillNum(lsFactorConversion, 2, "  ") & FillNum("", 382, " "); Chr(13) & Chr(10);

                
        lblAvance.Caption = "Avance :" & Format(J / contTotal * 100, "#0.000") & "%"
        
        lsPersonaAct = rs1!cCodPers
        
        rs1.MoveNext
        DoEvents
    Loop
End If
rs1.Close
Set rs1 = Nothing

'***************************************
'** Totales de la Empresa
'***************************************

Cont = Cont + 1
lsNumSecICC = FillNum(Trim(str(Cont)), 8, "0")

Print #NúmeroArchivo, "2" & "1" & FillNum(Trim(lsNumSecICC), 8, " ") & _
        Space(10) & Chr(13) & Chr(10);

'Actualiza la secuencia
'SQL1 = "Update dbconsolidada..RCDvc" & Format(gdFecData, "yyyymm") & "03 SET cNumSec ='" & lsNumSecICC & "' "
SQL1 = "Update dbconsolidada.." & lsNomTabla & Format(gdFecData, "yyyymm") & "03 SET cNumSec ='" & lsNumSecICC & "' " 'JUEZ 20150310
oConex.ejecutar (SQL1)

'Imprime totales
'JUEZ 20140409 ********************************************
'SQL1 = "Select * from dbconsolidada..RCDvc" & Format(gdFecData, "yyyymm") & "03 Order by cCtaCnt,cTipoCred,nCondDias "
SQL1 = "Select cTipoFor,cTipoInf,cNumSec,cCtaCnt,cTipoCred,nSaldo, "
SQL1 = SQL1 & "nCondDias = Case When nCondDias >= 10000 Then 9999 Else nCondDias End,cCondEspCta,cCondDisponib "
'SQL1 = SQL1 & "from dbconsolidada..RCDvc" & Format(gdFecData, "yyyymm") & "03 Order by cCtaCnt,cTipoCred,nCondDias "
SQL1 = SQL1 & "from dbconsolidada.." & lsNomTabla & Format(gdFecData, "yyyymm") & "03 Order by cCtaCnt,cTipoCred,nCondDias "
'END JUEZ *************************************************
Set rs1 = oConex.CargaRecordSet(SQL1)
'Dim lsFactorConversion As String
If Not RSVacio(rs1) Then
    Do While Not rs1.EOF
     lsFactorConversion = IIf(Mid(rs1!cCtaCnt, 1, 2) = "71", "02", "99")
     Print #NúmeroArchivo, "2" & "2" & FillNum(Trim(lsNumSecICC), 8, " ") & _
                FillNum(" ", 4, " ") & FillNum(" ", 6, " ") & FormatoCtaContable(rs1!cCtaCnt) & FillNum(Trim(rs1!cTipoCred), 2, "0") & _
                ImpreFormat(EliminaPunto(IIf(IsNull(rs1!nSaldo), 0, rs1!nSaldo)), 18, 0, False) & _
                FillNum(Trim(rs1!nCondDias), 4, " ") & FillNum(Trim(rs1!cCondEspCta), 2, " ") & _
                FillNum(Trim(rs1!cCondDisponib), 2, " ") & FillNum(lsFactorConversion, 2, "  ") & FillNum("", 382, " "); Chr(13) & Chr(10);

        rs1.MoveNext
    Loop
End If
rs1.Close
Set rs1 = Nothing

Close #NúmeroArchivo   ' Cierra el archivo.
Screen.MousePointer = 0
MsgBox "Se ha generado el archivo RCD" & Format(gdFecData, "yyyymm") & "." & lsNumEntidad & " satisfactoriamente. " & Chr(13) & "Terminó : " & Time(), vbInformation, "Aviso"
'NAGL 202007 Cambió de RCDvc00. a RCD" & Format(gdFecData, "yyyymm")
lblAvance.Caption = "" 'NAGL 202007
End Sub

'** GENERA EL ARCHIVO RCA *************************
Private Sub ReporteArchivoRCA()
Dim SQL1 As String
Dim contTotal As Long
Dim Cont As Long

Dim rs1 As New ADODB.Recordset
Dim lsPaisNac As String
Dim NúmeroArchivo As Integer
Dim lsNombreArchivo  As String
Dim J As Long
Dim lscTidoci As String
Dim lscNudoCi As String
Dim lscTidoTr As String
Dim lscNudotr As String
Dim lsActividad As String
Dim lsCodSbs As String
Dim lsCalificacion As String
Dim lsCodPers As String

NúmeroArchivo = FreeFile

Dim lsApePat As String, lsApeMat As String, lsApeCasada As String
Dim lsNomPri As String, lsNomSeg As String

Dim lsRiesCambiarioPersona As String
Dim lsCondEspCta As String

Dim lsCuentaTit As String
Dim lnSaldoTit As Double
       
Dim gsCodEmpInfICC As String
Dim gnMontoMinimoICC As Integer
Dim lsCadMancomuno As String
Dim lsCadAval As String
Dim lsNumSecICC As String
Dim oConex As New COMConecta.DCOMConecta
Dim sCtaTitular As String
oConex.AbreConexion

SQL1 = "select * from constsistema where nConsSisCod=80"
Set rs1 = oConex.CargaRecordSet(SQL1)
gsCodEmpInfICC = rs1!nConsSisValor

SQL1 = "select top 1 * from dbconsolidada..rcdparametro  order by  cMEs desc"
Set rs1 = oConex.CargaRecordSet(SQL1)
gnMontoMinimoICC = rs1!nMontoMin


' Crea el nombre de archivo.
lsNombreArchivo = App.Path & "\Spooler\RCA" & Format(gdFecData, "yyyymm") & ".109"
Open lsNombreArchivo For Output As #NúmeroArchivo

'**** Registro CABECERA
lsPaisNac = "PE"
'Print #NúmeroArchivo, "0106" & "01" & FillNum(Trim(gsCodEmpInfICC), 5, "0") & Format(gdFecData, "yyyymmdd") & "012" _
'        & Space(15) & ImpreFormat(EliminaPunto(IIf(IsNull(gnMontoMinimoICC), 0, gnMontoMinimoICC)), 15, 0, False) & Chr(13) & Chr(10);
'JUEZ 20130510 Se comentó la cabecera

SQL1 = " Select COUNT(r1.cPersCod) AS TOTAL From dbconsolidada..RCAvc" & Format(gdFecData, "yyyymm") & "01 R1 "
Set rs1 = oConex.CargaRecordSet(SQL1, adLockReadOnly)
If Not RSVacio(rs1) Then
    contTotal = IIf(IsNull(rs1!Total), 0, rs1!Total)
End If
rs1.Close
Set rs1 = Nothing

'Indexa la Tabla
'sql1 = "CREATE  INDEX cNumSec ON dbo.RCDvc" & Format(gdFecData, "yyyymm") & "01 ([cNumSec])"
'dbCmact.Execute sql1

'SQL1 = "Select R.cNumSec,  R.cCodSBS, R.cPersCod cCodPers, R.cActEcon, " _
'    & " R.cCodRegPub, R.cTidoTr, R.cNudoTr, R.cTiDoci, R.cNuDoci, R.cCodUnico, " _
'    & " R.cTipPers, R.cResid, R.cCalifica, R.cMagEmp, R.cAccionista, R.cRelInst,  " _
'    & " R.cPaisNac, R.cSiglas, R.cPersNom cNomPers, " _
'    & " R.cPersGenero cGenero,R.cPersEstado cEstadoCiv, " _
'    & " R.cApePat, R.cApeMat, R.cApeCasada, R.cNombre1 cNomPri, R.cNombre2 cNomSeg, " _
'    & " R.cCuentaTit, R.nSKTit, R.cCodAge, R.cUbicGeo " _
'    & " From dbconsolidada..RCAvc" & Format(gdFecData, "yyyymm") & "01 R Where R.cNumSec is not null " _
'    & " Order by R.cNumSec  "
'LUCV20170227, Comentó
'SQL1 = "Select  R.cNumSec,  R.cCodSBS, R.cPersCod cCodPers, R.cActEcon,"
'SQL1 = SQL1 & "     R.cCodRegPub, R.cTidoTr, R.cNudoTr, R.cTiDoci, R.cNuDoci, R.cCodUnico,"
'SQL1 = SQL1 & "     R.cTipPers, R.cResid, R.cCalifica, R.cMagEmp, R.cAccionista, R.cRelInst,"
'SQL1 = SQL1 & "     R.cPaisNac, R.cSiglas, R.cPersNom cNomPers,"
'SQL1 = SQL1 & "     R.cPersGenero cGenero,R.cPersEstado cEstadoCiv,"
'SQL1 = SQL1 & "     R.cApePat, R.cApeMat, R.cApeCasada, R.cNombre1 cNomPri, R.cNombre2 cNomSeg,"
''ALPA 20120427************************************************
'SQL1 = SQL1 & "     D.cCtaCnt,R.cCuentaTit , D.nSaldo, R.cCodAge, R.cUbicGeo, d.cTipoCred,"
''SQL1 = SQL1 & "     R.cCuentaTit , R.nSKTit, R.cCodAge, R.cUbicGeo, d.cTipoCred,"
''**************************************************************
'SQL1 = SQL1 & "     isnull(P.dPersNacCreac,'1990-01-01') dPersNacimiento " ' ALPA 20120327************************************
'SQL1 = SQL1 & " From (select distinct * from Dbconsolidada..RCAvc" & Format(gdFecData, "yyyymm") & "01) R inner join Dbconsolidada..RCDvc" & Format(gdFecData, "yyyymm") & "02 D on R.cCuentaTit=D.cCtaCod"
'SQL1 = SQL1 & "     inner join DBCmacMaynas..Persona P on R.cTitular=P.cPersCod " ' ALPA 20120327************************************
'SQL1 = SQL1 & " Where R.cNumSec is not null and D.cCtaCnt like '84_410%' "
'Fin LUCV20170227

SQL1 = "stp_sel_GeneraArchivoRCAReporte '" & Format(gdFecData, "yyyymm") & "'" 'Agregado por LUCV20170227
Set rs1 = oConex.CargaRecordSet(SQL1)

Cont = 0: J = 0
Screen.MousePointer = 11
If Not RSVacio(rs1) Then
    contTotal = rs1.RecordCount 'NAGL 202007
    Do While Not rs1.EOF
        J = J + 1
            If rs1!cPaisNac = "4028" Then
                lsPaisNac = "PE"
            End If
            sCtaTitular = rs1!cCuentaTit
            lsApePat = fgReemplazaCaracterEspecial(rs1!cApePat)
            lsApeMat = fgReemplazaCaracterEspecial(rs1!capemat)
            lsApeCasada = fgReemplazaCaracterEspecial(rs1!cApeCasada)
            lsNomPri = fgReemplazaCaracterEspecial(rs1!cNomPri)
            lsNomSeg = fgReemplazaCaracterEspecial(rs1!cNomSeg)
            
            Select Case Trim(rs1!cTipPers)
                Case "1", "3"
                    lscTidoci = IIf(IsNull(rs1!ctidoci), "", IIf(Len(Trim(rs1!ctidoci)) = 0, "", Trim(rs1!ctidoci)))
                    lscNudoCi = IIf(IsNull(rs1!cnudoci), "", IIf(Len(Trim(rs1!cnudoci)) = 0, "", Trim(rs1!cnudoci)))
                    lscTidoTr = ""
                    lscNudotr = ""
                Case Else
                    lscTidoci = ""
                    lscNudoCi = ""
                    lscTidoTr = IIf(IsNull(rs1!cTidoTr), "", IIf(Len(Trim(rs1!cTidoTr)) = 0, "", Trim(rs1!cTidoTr)))
                    lscNudotr = IIf(IsNull(rs1!cNudoTr), "", IIf(Len(Trim(rs1!cNudoTr)) = 0, "", Trim(rs1!cNudoTr)))
                    ' RUC DE 11 DIGITOS
                    lscTidoTr = IIf(lscTidoTr = "2", "3", lscTidoTr)
            End Select
            If Not IsNull(rs1!ccodsbs) Then
                lsCodSbs = Trim(rs1!ccodsbs)
            Else
                lsCodSbs = ""
            End If
            If Not IsNull(rs1!cActEcon) Then
                If Len(Trim(rs1!cActEcon)) > 0 Then
                    lsActividad = FillNum(rs1!cActEcon, 4, "0")
                Else
                    lsActividad = ""
                End If
            Else
                lsActividad = ""
            End If
            'LsCodPers = EmiteCodigoPersona(rs1!cCodPers)
            'If Len(Trim(lsCodPers)) = 0 Then
                lsCodPers = Trim(rs1!cCodPers)
            'End If
            lsCuentaTit = rs1!cCuentaTit
            lnSaldoTit = rs1!nSaldo 'rs1!nSKTit 'ALPA 20120427
            
            
        'JACA 20110428************************************************************
            Dim cMagEmp As String
            If Trim(rs1!cMagEmp) = "0" Then
                cMagEmp = "1"
            ElseIf Trim(rs1!cMagEmp) = "1" Then
                cMagEmp = "6"
            ElseIf Trim(rs1!cMagEmp) = "2" Then
                cMagEmp = "7"
            ElseIf Trim(rs1!cMagEmp) = "3" Then
                cMagEmp = "8"
            ElseIf Trim(rs1!cMagEmp) = "4" Then
                cMagEmp = "0"
            ElseIf Trim(rs1!cMagEmp) = "5" Then
                cMagEmp = "5"
            End If
        'JACA END***************************************************************
            
            '**DAOR 20070914, Por indicaciones de la Jefatura y Gerencia Reemplazar cCodUnico por cPersCod
            
            'JACA 20110428, SE REEMPLAZO LA SIGTE INSTRUCCION DEL PRINT******************
            'FillNum(Trim(IIf(IsNull(rs1!cMagEmp), "", rs1!cMagEmp)), 1, " ") & _
            ' POR ESTE
            'FillNum(Trim(IIf(IsNull(cMagEmp), "", cMagEmp)), 1, " ") & _

            Print #NúmeroArchivo, "1" & "4" & FillNum(Trim(rs1!cNumSec), 8, " ") & _
                    ImpreFormat(IIf(Len(Trim(lsCodSbs)) = 0, "", lsCodSbs), 10, 0) & _
                    ImpreFormat(Trim(rs1!cCodPers), 20, 0) & FillNum(lsActividad, 4, " ") & _
                    FillNum(Trim(IIf(IsNull(rs1!ccodregpub), "", rs1!ccodregpub)), 15, " ") & _
                    FillNum(IIf(Len(Trim(lscTidoTr)) = 0, "", lscTidoTr), 1, " ") & _
                    FillNum(IIf(Len(Trim(lscNudotr)) = 0, "", lscNudotr), 11, " ") & _
                    FillNum(IIf(Len(Trim(lscTidoci)) = 0, "", lscTidoci), 1, " ") & _
                    ImpreFormat(IIf(Len(Trim(lscNudoCi)) = 0, "", lscNudoCi), 12, 0) & _
                    FillNum(Trim(IIf(IsNull(rs1!cTipPers), "", rs1!cTipPers)), 1, " ") & _
                    FillNum(Trim(IIf(IsNull(rs1!cResid), "", rs1!cResid)), 1, " ") & _
                    Space(1) & _
                    FillNum(Trim(IIf(IsNull(cMagEmp), "", cMagEmp)), 1, " ") & _
                    FillNum(Trim(IIf(IsNull(rs1!cAccionista), "", rs1!cAccionista)), 1, " ") & FillNum(Trim(IIf(IsNull(rs1!cRelInst), "", rs1!cRelInst)), 1, " ") & _
                    FillNum(Trim(lsPaisNac), 4, " ") & FillNum(Trim(rs1!cGenero), 1, " ") & _
                    FillNum(Trim(rs1!cEstadoCiv), 1, " ") & _
                    Space(20) & _
                    ImpreFormat(Trim(lsApePat), 120, 0) & ImpreFormat(Trim(lsApeMat), 40, 0) & _
                    ImpreFormat(Trim(lsApeCasada), 40, 0) & ImpreFormat(Trim(lsNomPri), 40, 0) & _
                    ImpreFormat(Trim(lsNomSeg), 40, 0) & _
                    Space(28) & _
                    CStr(Format(rs1!dPersNacimiento, "YYYYMMDD")) & _
                    Space(14); Chr(13) & Chr(10);
        'JACA*******************************************************************
            Print #NúmeroArchivo, "1" & "5" & FillNum(Trim(rs1!cNumSec), 8, " ") & _
                FillNum(Trim(rs1!cCodAge), 4, "0") & FillNum(Trim(rs1!cUbicGeo), 6, " ") & _
                FormatoCtaContableRCANew("84" & Mid(IIf(IsNull(rs1!cCtaCnt), "109011111", rs1!cCtaCnt), 3, 1) & "41000000000") & FillNum(Trim(rs1!cTipoCred), 2, "  ") & _
                Replace(ImpreFormat(EliminaPunto(IIf(IsNull(lnSaldoTit), 0, lnSaldoTit)), 18, 0, False), " ", "0") & _
                Space(420); Chr(13) & Chr(10);
                'JUEZ 20130401 Se agregó funcion Replace en los lnSaldoTit para reemplazar los espacios en blanco con 0

        lblAvance.Caption = "Avance :" & Format(J / contTotal * 100, "#0.000") & "%"
        
        rs1.MoveNext
        DoEvents
    Loop
End If
rs1.Close
Set rs1 = Nothing


Close #NúmeroArchivo   ' Cierra el archivo.
Screen.MousePointer = 0
MsgBox "Se ha generado el Archivo RCA" & Format(gdFecData, "yyyymm") & "." & "109 Satisfactoriamente, Termino : " & Time(), vbInformation, "Aviso"
'NAGL 202007 Cambió de RCAvc00. a RCA & Format(gdFecData, "yyyymm")
lblAvance.Caption = "" 'NAGL 202007
End Sub

'** GENERA EL ARCHIVO RCM *************************
Private Sub ReporteArchivoRCM()
Dim SQL1 As String
Dim contTotal As Long
Dim Cont As Long

Dim rs1 As New ADODB.Recordset

Dim NúmeroArchivo As Integer
Dim lsNombreArchivo  As String
Dim J As Long
Dim lscTidoci As String
Dim lscNudoCi As String
Dim lscTidoTr As String
Dim lscNudotr As String
Dim lsActividad As String
Dim lsCodSbs As String
Dim lsCalificacion As String
Dim lsCodPers As String

NúmeroArchivo = FreeFile

Dim lsPersonaAct As String
Dim lnSaldo As Currency

Dim lsApePat As String, lsApeMat As String, lsApeCasada As String
Dim lsNomPri As String, lsNomSeg As String

Dim lsRiesCambiarioPersona As String
Dim lsCondEspCta As String
lsRiesCambiarioPersona = "0"

Dim lsIndicadorAtr As String
Dim lsCalifInterna As String
       
Dim gsCodEmpInfICC As String
Dim gnMontoMinimoICC As Integer
Dim lsCadMancomuno As String
Dim lsCadAval As String
Dim lsNumSecICC As String
Dim oConex As New COMConecta.DCOMConecta
oConex.AbreConexion

SQL1 = "select * from constsistema where nConsSisCod=80"
Set rs1 = oConex.CargaRecordSet(SQL1)
gsCodEmpInfICC = rs1!nConsSisValor

SQL1 = "select top 1 * from dbconsolidada..rcdparametro  order by  cMEs desc"
Set rs1 = oConex.CargaRecordSet(SQL1)
gnMontoMinimoICC = rs1!nMontoMin


' Crea el nombre de archivo.
'EstatuBarICC.Panels(1).Text = "Generación de Reporte ICC en Archivo"
lsNombreArchivo = App.Path & "\Spooler\RCM" & Format(gdFecData, "yyyymm") & ".109"
Open lsNombreArchivo For Output As #NúmeroArchivo


'**** Registro CABECERA
Print #NúmeroArchivo, "0106" & "01" & FillNum(Trim(gsCodEmpInfICC), 5, "0") & Format(gdFecData, "yyyymmdd") & "012" _
        & Space(15) & ImpreFormat(EliminaPunto(IIf(IsNull(gnMontoMinimoICC), 0, gnMontoMinimoICC)), 15, 0, False) & Chr(13) & Chr(10);


SQL1 = " Select COUNT(r1.cPersCod) AS TOTAL From dbconsolidada..RCDvc" & Format(gdFecData, "yyyymm") & "01 R1 " _
     & " Inner Join dbconsolidada..RCDvc" & Format(gdFecData, "yyyymm") & "02 R2 ON R2.cPersCod = R1.cPersCod "

Set rs1 = oConex.CargaRecordSet(SQL1, adLockReadOnly)
If Not RSVacio(rs1) Then
    contTotal = IIf(IsNull(rs1!Total), 0, rs1!Total)
End If
rs1.Close
Set rs1 = Nothing

'Indexa la Tabla
'sql1 = "CREATE  INDEX cNumSec ON dbo.RCDvc" & Format(gdFecData, "yyyymm") & "01 ([cNumSec])"
'dbCmact.Execute sql1

SQL1 = "Select R1.cTipoFor, R1.cTipoInf, R1.cNumSec, R1.cCodSBS, R1.cPersCod cCodPers, R1.cActEcon, " _
    & " R1.cCodRegPub, R1.cTidoTr, R1.cNudoTr, R1.cTiDoci, R1.cNuDoci, R1.cCodUnico, " _
    & " R1.cTipPers, R1.cResid, R1.cCalifica, R1.cMagEmp, R1.cAccionista, R1.cRelInst,  " _
    & " R1.cPaisNac, R1.cSiglas, R1.cPersNom cNomPers, " _
    & " R1.cPersGenero cGenero,R1.cPersEstado cEstadoCiv,  " _
    & " R1.cIndRCambiario cRCambiario, R1.cIndAtrasoDeudor cIndicadorAtraso, R1.cClasifInterna cCalifInterna, " _
    & " R1.cApePat, R1.cApeMat, R1.cApeCasada, R1.cNombre1 cNomPri, R1.cNombre2 cNomSeg " _
    & " From dbconsolidada..RCDvc" & Format(gdFecData, "yyyymm") & "01 R1 " _
    & " ORDER BY R1.cNumSec " _
    
Set rs1 = oConex.CargaRecordSet(SQL1)

Cont = 0: J = 0
Screen.MousePointer = 11
If Not RSVacio(rs1) Then

    Do While Not rs1.EOF
        J = J + 1

        lsNumSecICC = rs1!cNumSec
        'Mancomunos
        lsCadMancomuno = fCadenaMancomuno(lsNumSecICC)
        If Len(lsCadMancomuno) > 0 Then
            
            Cont = Cont + 1
            lsApePat = fgReemplazaCaracterEspecial(rs1!cApePat)
            lsApeMat = fgReemplazaCaracterEspecial(rs1!capemat)
            lsApeCasada = fgReemplazaCaracterEspecial(rs1!cApeCasada)
            lsNomPri = fgReemplazaCaracterEspecial(rs1!cNomPri)
            lsNomSeg = fgReemplazaCaracterEspecial(rs1!cNomSeg)
            
            Select Case Trim(rs1!cTipPers)
                Case "1", "3"
                    lscTidoci = IIf(IsNull(rs1!ctidoci), "", IIf(Len(Trim(rs1!ctidoci)) = 0, "", Trim(rs1!ctidoci)))
                    lscNudoCi = IIf(IsNull(rs1!cnudoci), "", IIf(Len(Trim(rs1!cnudoci)) = 0, "", Trim(rs1!cnudoci)))
                    lscTidoTr = ""
                    lscNudotr = ""
                Case Else
                    lscTidoci = ""
                    lscNudoCi = ""
                    lscTidoTr = IIf(IsNull(rs1!cTidoTr), "", IIf(Len(Trim(rs1!cTidoTr)) = 0, "", Trim(rs1!cTidoTr)))
                    lscNudotr = IIf(IsNull(rs1!cNudoTr), "", IIf(Len(Trim(rs1!cNudoTr)) = 0, "", Trim(rs1!cNudoTr)))
                    ' RUC DE 11 DIGITOS
                    lscTidoTr = IIf(lscTidoTr = "2", "3", lscTidoTr)
            End Select
            If Not IsNull(rs1!ccodsbs) Then
                lsCodSbs = Trim(rs1!ccodsbs)
            Else
                lsCodSbs = ""
            End If
            If Not IsNull(rs1!cActEcon) Then
                If Len(Trim(rs1!cActEcon)) > 0 Then
                    lsActividad = FillNum(rs1!cActEcon, 4, "0")
                Else
                    lsActividad = ""
                End If
            Else
                lsActividad = ""
            End If
            
            'LsCodPers = EmiteCodigoPersona(rs1!cCodPers)
            'If Len(Trim(lsCodPers)) = 0 Then
                lsCodPers = Trim(rs1!cCodPers)
            'End If
            
            lsCalificacion = Trim(rs1!cCalifica)
            lsRiesCambiarioPersona = Trim(rs1!cRCambiario)
            
            lsIndicadorAtr = Trim(rs1!cIndicadorAtraso)
            lsCalifInterna = Trim(IIf(IsNull(rs1!cCalifInterna), "X", rs1!cCalifInterna))
            If Trim(lsCalifInterna) = "X" Then
                lsCalifInterna = "XXXXX"
            End If

            Print #NúmeroArchivo, rs1!cTipoFor & rs1!cTipoInf & FillNum(Trim(lsNumSecICC), 8, " ") & _
                    ImpreFormat(IIf(Len(Trim(lsCodSbs)) = 0, "", lsCodSbs), 10, 0) & _
                    ImpreFormat(Trim(rs1!cCodUnico), 20, 0) & FillNum(lsActividad, 4, " ") & _
                    FillNum(Trim(IIf(IsNull(rs1!ccodregpub), "", rs1!ccodregpub)), 15, " ") & _
                    FillNum(IIf(Len(Trim(lscTidoTr)) = 0, "", lscTidoTr), 1, " ") & _
                    FillNum(IIf(Len(Trim(lscNudotr)) = 0, "", lscNudotr), 11, " ") & _
                    FillNum(IIf(Len(Trim(lscTidoci)) = 0, "", lscTidoci), 1, " ") & _
                    ImpreFormat(IIf(Len(Trim(lscNudoCi)) = 0, "", lscNudoCi), 12, 0) & _
                    FillNum(Trim(IIf(IsNull(rs1!cTipPers), "", rs1!cTipPers)), 1, " ") & FillNum(Trim(IIf(IsNull(rs1!cResid), "", rs1!cResid)), 1, " ") & _
                    FillNum(Trim(lsCalificacion), 1, " ") & _
                    FillNum(Trim(IIf(IsNull(rs1!cMagEmp), "", rs1!cMagEmp)), 1, " ") & _
                    FillNum(Trim(IIf(IsNull(rs1!cAccionista), "", rs1!cAccionista)), 1, " ") & FillNum(Trim(IIf(IsNull(rs1!cRelInst), "", rs1!cRelInst)), 1, " ") & _
                    FillNum(Trim(rs1!cPaisNac), 4, " ") & FillNum(Trim(rs1!cGenero), 1, " ") & _
                    FillNum(Trim(rs1!cEstadoCiv), 1, " ") & _
                    ImpreFormat(Trim(IIf(IsNull(rs1!cSiglas), "", rs1!cSiglas)), 20, 0) & _
                    ImpreFormat(Trim(lsApePat), 120, 0) & ImpreFormat(Trim(lsApeMat), 40, 0) & _
                    ImpreFormat(Trim(lsApeCasada), 40, 0) & ImpreFormat(Trim(lsNomPri), 40, 0) & _
                    ImpreFormat(Trim(lsNomSeg), 40, 0) & _
                    lsRiesCambiarioPersona & lsIndicadorAtr & FillNum(lsCalifInterna, 5, " ") _
                    ; Chr(13) & Chr(10);
                    
            'Mancomuno
            Print #NúmeroArchivo, lsCadMancomuno
                    
        End If
        lblAvance.Caption = "Avance :" & Format(J / contTotal * 100, "#0.000") & "%"
        rs1.MoveNext
        DoEvents
    Loop
End If
rs1.Close
Set rs1 = Nothing

Close #NúmeroArchivo   ' Cierra el archivo.
Screen.MousePointer = 0
MsgBox "Se ha generado el Archivo RCMvc00.109 Satisfactoriamente, Termino : " & Time()
End Sub

Private Function fCadenaMancomuno(ByVal psNumSec As String) As String
Dim lsCadena As String
Dim SQL1 As String
Dim rs1 As New ADODB.Recordset

Dim lscTidoci As String
Dim lscNudoCi As String
Dim lscTidoTr As String
Dim lscNudotr As String
Dim lsActividad As String
Dim lsCodSbs As String
Dim lsCalificacion As String
Dim lsCodPers As String
Dim lsApePat As String, lsApeMat As String, lsApeCasada As String
Dim lsNomPri As String, lsNomSeg As String
Dim oConex As New COMConecta.DCOMConecta
oConex.AbreConexion
    SQL1 = "Select R.cCodSBS, R.cPersCod , R.cActEcon, R.cCodUnico, " _
        & " R.cTipPers, R.cResid, R.cCalifica, R.cMagEmp, R.cAccionista, R.cRelInst, " _
        & " R.cCodRegPub, R.cTidoTr, R.cNudoTr, R.cTiDoci, R.cNuDoci, " _
        & " R.cPaisNac, R.cSiglas, R.cPersNom cNomPers, " _
        & " R.cPersGenero cGenero, R.cPersEstado cEstadoCiv,  " _
        & " R.cApePat, R.cApeMat, R.cApeCasada, R.cNombre1 cNomPri, R.cNombre2 cNomSeg " _
        & " From dbConsolidada..RCMvc" & Format(gdFecData, "yyyymm") & "01 R " _
        & " Where R.cNumSec = '" & psNumSec & "' " _
        & " Order by R.cApePat, R.cApeMat, R.cApeCasada, R.cNombre1, R.cNombre2 "
    
    Set rs1 = oConex.CargaRecordSet(SQL1)
    If rs1.BOF And rs1.EOF Then
        lsCadena = ""
    Else
        lsApePat = fgReemplazaCaracterEspecial(rs1!cApePat)
        lsApeMat = fgReemplazaCaracterEspecial(rs1!capemat)
        lsApeCasada = fgReemplazaCaracterEspecial(rs1!cApeCasada)
        lsNomPri = fgReemplazaCaracterEspecial(rs1!cNomPri)
        lsNomSeg = fgReemplazaCaracterEspecial(rs1!cNomSeg)
        
        Select Case Trim(rs1!cTipPers)
            Case "1", "3"
                lscTidoci = IIf(IsNull(rs1!ctidoci), "", IIf(Len(Trim(rs1!ctidoci)) = 0, "", Trim(rs1!ctidoci)))
                lscNudoCi = IIf(IsNull(rs1!cnudoci), "", IIf(Len(Trim(rs1!cnudoci)) = 0, "", Trim(rs1!cnudoci)))
                lscTidoTr = ""
                lscNudotr = ""
            Case Else
                lscTidoci = ""
                lscNudoCi = ""
                lscTidoTr = IIf(IsNull(rs1!cTidoTr), "", IIf(Len(Trim(rs1!cTidoTr)) = 0, "", Trim(rs1!cTidoTr)))
                lscNudotr = IIf(IsNull(rs1!cNudoTr), "", IIf(Len(Trim(rs1!cNudoTr)) = 0, "", Trim(rs1!cNudoTr)))
                ' RUC DE 11 DIGITOS
                lscTidoTr = IIf(lscTidoTr = "2", "3", lscTidoTr)
        End Select
        If Not IsNull(rs1!ccodsbs) Then
            lsCodSbs = Trim(rs1!ccodsbs)
        Else
            lsCodSbs = ""
        End If
        If Not IsNull(rs1!cActEcon) Then
            If Len(Trim(rs1!cActEcon)) > 0 Then
                lsActividad = FillNum(rs1!cActEcon, 4, "0")
            Else
                lsActividad = ""
            End If
        Else
            lsActividad = ""
        End If
        
        lsCodPers = rs1!cPersCod
        If Len(Trim(lsCodPers)) = 0 Then
            lsCodPers = Trim(rs1!cCodPers)
        End If
        
        lsCalificacion = Trim(rs1!cCalifica)
        
        lsCadena = "1" & "3" & FillNum(Trim(psNumSec), 8, " ") & _
                ImpreFormat(IIf(Len(Trim(lsCodSbs)) = 0, "", lsCodSbs), 10, 0) & ImpreFormat(Trim(rs1!cCodUnico), 20, 0) & FillNum(lsActividad, 4, " ") & _
                FillNum(Trim(IIf(IsNull(rs1!ccodregpub), "", rs1!ccodregpub)), 15, " ") & _
                FillNum(IIf(Len(Trim(lscTidoTr)) = 0, "", lscTidoTr), 1, " ") & _
                FillNum(IIf(Len(Trim(lscNudotr)) = 0, "", lscNudotr), 11, " ") & _
                FillNum(IIf(Len(Trim(lscTidoci)) = 0, "", lscTidoci), 1, " ") & _
                ImpreFormat(IIf(Len(Trim(lscNudoCi)) = 0, "", lscNudoCi), 12, 0) & _
                FillNum(Trim(IIf(IsNull(rs1!cTipPers), "", rs1!cTipPers)), 1, " ") & FillNum(Trim(IIf(IsNull(rs1!cResid), "", rs1!cResid)), 1, " ") & _
                FillNum(Trim(lsCalificacion), 1, " ") & _
                FillNum(Trim(IIf(IsNull(rs1!cMagEmp), "", rs1!cMagEmp)), 1, " ") & _
                FillNum(Trim(IIf(IsNull(rs1!cAccionista), "", rs1!cAccionista)), 1, " ") & FillNum(Trim(IIf(IsNull(rs1!cRelInst), "", rs1!cRelInst)), 1, " ") & _
                FillNum(Trim(rs1!cPaisNac), 4, " ") & FillNum(Trim(rs1!cGenero), 1, " ") & _
                FillNum(Trim(rs1!cEstadoCiv), 1, " ") & _
                Space(20) & _
                ImpreFormat(Trim(lsApePat), 120, 0) & ImpreFormat(Trim(lsApeMat), 40, 0) & _
                ImpreFormat(Trim(lsApeCasada), 40, 0) & ImpreFormat(Trim(lsNomPri), 40, 0) & _
                ImpreFormat(Trim(lsNomSeg), 40, 0) '& Chr(13) & Chr(10)
    End If
    Set rs1 = Nothing
    fCadenaMancomuno = lsCadena
End Function


Private Function fgReemplazaCaracterEspecial(ByVal psNom As String) As String
Dim lsNombrePers As String
            
            lsNombrePers = CadDerecha(Trim(Replace(psNom, "-", "", , , vbTextCompare)), 80)
            lsNombrePers = CadDerecha(Trim(Replace(psNom, ".", " ", , , vbTextCompare)), 80)
            lsNombrePers = CadDerecha(Trim(Replace(psNom, "Ñ", "#", , , vbTextCompare)), 80)
            lsNombrePers = CadDerecha(Trim(Replace(psNom, "ñ", "#", , , vbTextCompare)), 80)
fgReemplazaCaracterEspecial = lsNombrePers
End Function

Public Function FormatoCtaContable(ByVal pCuentaCnt As String) As String
  FormatoCtaContable = Trim(pCuentaCnt) & String(14 - Len(Trim(pCuentaCnt)), "0")
End Function
'ALPA 20110711 **********************
Public Function FormatoCtaContableRCANew(ByVal pCuentaCnt As String) As String
  FormatoCtaContableRCANew = Trim(pCuentaCnt) & String(14 - Len(Trim(pCuentaCnt)), "0")
End Function

Public Function FormatoCtaContableRCA(ByVal pCuentaCnt As String) As String
  FormatoCtaContableRCA = Trim(pCuentaCnt) & String(18 - Len(Trim(pCuentaCnt)), "0")
End Function
'************************************
'JUEZ 20150113 ***************************************************
Private Sub ReporteVerifRCDconRCC()
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim lbLibroOpen As Boolean
Dim lsArchivo As String
Dim lsHoja As String
Dim xlHoja1 As Excel.Worksheet
Dim xlHoja2 As Excel.Worksheet
Dim nLin As Long
Dim nMaximo As Long
Dim J As Long
Dim oConex As New COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim ssql As String
    
    oConex.AbreConexion
    ssql = "stp_sel_179106_ReporteVerificacionRCDconRCC '" & Format(gdFecData, "yyyymm") & "'"
    
    Set rs = oConex.CargaRecordSet(ssql)
    If rs.BOF And rs.EOF Then
        MsgBox "No hay datos para el reporte", vbInformation, "Aviso"
    Else
        lsArchivo = App.Path & "\SPOOLER\VerificacionRCD_RCC_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".xls"
        lbLibroOpen = gFunGeneral.ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
        If Not lbLibroOpen Then
            Exit Sub
        End If
        nLin = 1
        lsHoja = "Reporte"
        gFunGeneral.ExcelAddHoja lsHoja, xlLibro, xlHoja1
        
        xlHoja1.Range("A1:Y1").EntireColumn.Font.FontStyle = "Arial"
        xlHoja1.PageSetup.CenterHorizontally = True
        xlHoja1.PageSetup.Zoom = 75
        xlHoja1.PageSetup.TopMargin = 2
        
        xlHoja1.Range("A1:A1").RowHeight = 18
        xlHoja1.Range("A2:A2").RowHeight = 50
        xlHoja1.Range("A1:B1").ColumnWidth = 12
        xlHoja1.Range("C1:C1").ColumnWidth = 50
        xlHoja1.Range("D1:E1").ColumnWidth = 12
        xlHoja1.Range("F1:F1").ColumnWidth = 50
        xlHoja1.Range("G1:G1").ColumnWidth = 12
        
        xlHoja1.Cells(nLin, 1) = "RCD"
        xlHoja1.Range("A1", "E1").MergeCells = True
        xlHoja1.Cells(nLin, 7) = "RCC"
        xlHoja1.Range("F1", "G1").MergeCells = True
        nLin = nLin + 1
        xlHoja1.Cells(nLin, 1) = "Núm." & Chr(10) & "Secuencia"
        xlHoja1.Cells(nLin, 2) = "Cód. SICMAC"
        xlHoja1.Cells(nLin, 3) = "Cliente"
        xlHoja1.Cells(nLin, 4) = "Documento"
        xlHoja1.Cells(nLin, 5) = "Cód. SBS" & Chr(10) & "SICMAC"
        xlHoja1.Cells(nLin, 6) = "Nombre Deudor"
        xlHoja1.Cells(nLin, 7) = "Cód. SBS RCC" & Chr(10) & "(cambiarlo por" & Chr(10) & "este número)"
        
        xlHoja1.Range("G2:G2").EntireColumn.AutoFit
        xlHoja1.Range("A1:G2").Font.Bold = True
        xlHoja1.Range("A1:G2").HorizontalAlignment = xlHAlignCenter
        xlHoja1.Range("A1:G2").VerticalAlignment = xlVAlignCenter
        xlHoja1.Range("A1:G2").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
        xlHoja1.Range("A1:G2").Borders.LineStyle = xlContinuous
        xlHoja1.Range("A1:G2").Borders.Color = vbBlack
        xlHoja1.Range("A1:G2").Interior.Color = RGB(255, 50, 50)
        xlHoja1.Range("A1:G2").Font.Color = RGB(255, 255, 255)
        
        With xlHoja1.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = ""
        
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = xlPrintNoComments
            .CenterHorizontally = True
            .CenterVertically = False
            .Draft = False
            .FirstPageNumber = xlAutomatic
            .Order = xlDownThenOver
            .BlackAndWhite = False
            .Zoom = 55
        End With
    
        nLin = nLin + 1
        J = 1
        nMaximo = rs.RecordCount
        Do While Not rs.EOF
            xlHoja1.Range("A" & nLin & ":G" & nLin).Borders.LineStyle = xlContinuous
            xlHoja1.Range("A" & nLin & ":G" & nLin).Borders.Color = vbBlack
            xlHoja1.Range("G" & nLin & ":G" & nLin).Interior.Color = RGB(255, 250, 0)
            xlHoja1.Cells(nLin, 1) = "'" & rs!cNumSec
            xlHoja1.Cells(nLin, 2) = "'" & rs!cPersCod
            xlHoja1.Cells(nLin, 3) = rs!cPersNom
            xlHoja1.Cells(nLin, 4) = "'" & RTrim(rs!cnudoci)
            xlHoja1.Cells(nLin, 5) = "'" & rs!ccodsbs
            xlHoja1.Cells(nLin, 6) = rs!Nom_Deu
            xlHoja1.Cells(nLin, 7) = "'" & rs!Cod_Edu
            
            lblAvance.Caption = "Avance :" & Format(J / nMaximo * 100, "#0.000") & "%"
            nLin = nLin + 1
            J = J + 1
            rs.MoveNext
            DoEvents
        Loop
        Set rs = Nothing
    
        gFunGeneral.ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
        gFunGeneral.CargaArchivo lsArchivo, App.Path & "\SPOOLER\"
        MsgBox "Se ha generado el archivo VerificacionRCD_RCC_" & Format(gdFecSis, "yyyymmdd"), vbInformation, "Aviso" 'NAGL 202007
        lblAvance.Caption = "" 'NAGL 202007
    End If
End Sub


'***** LUCV20170415, Agregó *****
Private Sub GeneraRptExcellRCD()
 Dim nMES As Integer
    'Dim nAnio As Integer
    Dim dFecha As Date
    
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet
    'Dim ldFecha As Date
    'dFecha = Format(gdFecData, "yyyymm")
On Error GoTo GeneraRptGarantiaError
    'nMes = cboMes.ListIndex + 1
    'nAnio = txtAnio
    'dFecha = DateAdd("m", 1, "01/" & Format(nMes, "00") & "/" & Format(txtAnio, "0000")) - 1

    'Generacion
    lsArchivo = "\spooler\RptRCD" & UCase(gsCodUser) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    'ldFecha = dFecha
    
    Set xlsLibro = xlsAplicacion.Workbooks.Add
    'HOJA RCD
    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = "RCD"
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    Call GeneraRptHojaRCD(xlsHoja)
    
    'HOJA GARANTIAS
'    Set xlsHoja = xlsLibro.Worksheets.Add
'    xlsHoja.Name = "GARANTIAS"
'    xlsHoja.Cells.Font.Name = "Arial"
'    xlsHoja.Cells.Font.Size = 9
'    Call GeneraHojaGarantiaRpt(ldFecha, xlsHoja)
    
    MsgBox "Se ha generado satisfactoriamente el reporte de garantias", vbInformation, "Aviso"
    xlsHoja.SaveAs App.Path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True

    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlsHoja = Nothing
    
    Exit Sub
    'fin generacion
GeneraRptGarantiaError:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub GeneraRptHojaRCD(ByRef xlsHoja As Worksheet, Optional ByVal pbRCDTransf As Boolean = False)
    Dim SQL1 As String
    Dim contTotal As Long
    Dim Cont As Long
    
    Dim rs1 As New ADODB.Recordset
    
    Dim NúmeroArchivo As Integer
    Dim lsNombreArchivo  As String
    Dim J As Long
    Dim lscTidoci As String
    Dim lscNudoCi As String
    Dim lscTidoTr As String
    Dim lscNudotr As String
    Dim lsActividad As String
    Dim lsCodSbs As String
    Dim lsCalificacion As String
    Dim lsCodPers As String
    Dim lsNomTabla As String
    Dim lsNumEntidad As String
    
    NúmeroArchivo = FreeFile
    
    Dim lsPersonaAct As String
    Dim lnSaldo As Currency
    
    Dim lsApePat As String, lsApeMat As String, lsApeCasada As String
    Dim lsNomPri As String, lsNomSeg As String
    
    Dim lsRiesCambiarioPersona As String
    Dim lsCondEspCta As String
    lsRiesCambiarioPersona = "0"
    
    Dim lsIndicadorAtr As String
    Dim lsCalifInterna As String
           
    Dim gsCodEmpInfICC As String
    Dim gnMontoMinimoICC As Integer
    Dim lsCadMancomuno As String
    Dim lsCadAval As String
    Dim lsNumSecICC As String
    Dim oConex As New COMConecta.DCOMConecta
    Dim lsConsol As String
    Dim lsFactorConversion As String
    
    lsConsol = "dbconsolidada.."
    oConex.AbreConexion
    
    lsNomTabla = IIf(pbRCDTransf, "RCDTvc", "RCDvc")
    lsNumEntidad = IIf(pbRCDTransf, "399", "109")
    
    SQL1 = "select * from constsistema where nConsSisCod=80"
    Set rs1 = oConex.CargaRecordSet(SQL1)
    gsCodEmpInfICC = rs1!nConsSisValor
    If pbRCDTransf Then gsCodEmpInfICC = "399"
    
    SQL1 = "select top 1 * from dbconsolidada..rcdparametro  order by  cMEs desc"
    Set rs1 = oConex.CargaRecordSet(SQL1)
    gnMontoMinimoICC = rs1!nMontoMin
    
    ' Crea el nombre de archivo.
    'lsNombreArchivo = App.Path & "\Spooler\RCD" & Format(gdFecData, "yyyymm") & "." & lsNumEntidad 'JUEZ 20150310
    'Open lsNombreArchivo For Output As #NúmeroArchivo
    
    '**** Registro CABECERA
    Dim lsMontoMinimo As String
    lsMontoMinimo = Trim(ImpreFormat(EliminaPunto(IIf(IsNull(gnMontoMinimoICC), 0, gnMontoMinimoICC)), 15, 0, False))
    

'    Print #NúmeroArchivo, "0106" & "01" & FillNum(Trim(gsCodEmpInfICC), 5, "0") & Format(gdFecData, "yyyymmdd") & "012" _
'            & Space(15) & String(15 - Len(lsMontoMinimo), "0") & lsMontoMinimo & Chr(13) & Chr(10);
    
        ' lucv inicio
        SQL1 = " Select COUNT(r1.cPersCod) AS TOTAL From dbconsolidada.." & lsNomTabla & Format(gdFecData, "yyyymm") & "01 R1 " _
     & " Inner Join dbconsolidada.." & lsNomTabla & Format(gdFecData, "yyyymm") & "02 R2 ON R2.cPersCod = R1.cPersCod "

Set rs1 = oConex.CargaRecordSet(SQL1, adLockReadOnly)
If Not RSVacio(rs1) Then
    contTotal = IIf(IsNull(rs1!Total), 0, rs1!Total)
End If
rs1.Close
Set rs1 = Nothing

 
 SQL1 = "stp_sel_GeneraArchivoRCDReporte '" & Format(gdFecData, "yyyymm") & "'" 'Agregado por LUCV20170227
Set rs1 = oConex.CargaRecordSet(SQL1)

Cont = 0: J = 0
Screen.MousePointer = 11
If Not RSVacio(rs1) Then
'Cabecera Descripcion
        Dim I As Integer
        Dim nFinal As Integer
        
        Dim lnPosCabecera As Integer 'LUCV20170415, Agregó
        Dim lnPosIdentificador As Integer
        Dim lnPosSaldos As Integer
        Dim lnPosActual As Integer
        lnPosCabecera = 1
        lnPosIdentificador = 4
        lnPosSaldos = 5
        I = 1
        nFinal = 480
        
        For I = 1 To nFinal
            xlsHoja.Cells(lnPosCabecera, I) = I
        Next I
        xlsHoja.Cells.Font.Name = "Arial"
        xlsHoja.Cells.Font.Size = 9
        xlsHoja.Range("A2", "AZ" & 2).Borders.LineStyle = xlContinuous
        xlsHoja.Range("A2", "AY" & 2).Borders.Weight = xlThin
        xlsHoja.Range("A2", "AY" & 2).Borders.ColorIndex = xlAutomatic
        xlsHoja.Range("A2", "AY" & 2).Interior.Color = RGB(221, 235, 247)
        xlsHoja.Range("A2", "AY" & 2).Font.Color = RGB(0, 32, 96)
        xlsHoja.Range(xlsHoja.Cells(2, 1), xlsHoja.Cells(2, nFinal)).HorizontalAlignment = xlCenter
        xlsHoja.Range(xlsHoja.Cells(2, 1), xlsHoja.Cells(2, nFinal)).Font.Bold = True
        xlsHoja.Range("A2", "RL2").ColumnWidth = 2
                
        lnPosCabecera = lnPosCabecera + 1
        
        xlsHoja.Cells(lnPosCabecera, 1) = "Cod. For."
        xlsHoja.Range("A2", "D2").MergeCells = True
        
        xlsHoja.Cells(lnPosCabecera, 5) = "CA"
        xlsHoja.Range("E2", "F2").MergeCells = True
         
        xlsHoja.Cells(lnPosCabecera, 7) = "Cod. Empre."
        xlsHoja.Range("G2", "K2").MergeCells = True
        
        xlsHoja.Cells(lnPosCabecera, 12) = "Fecha Periodo Rep"
        xlsHoja.Range("L2", "S2").MergeCells = True
        
        xlsHoja.Cells(lnPosCabecera, 20) = "CEM"
        xlsHoja.Range("T2", "V2").MergeCells = True
        
        xlsHoja.Cells(lnPosCabecera, 23) = "Datos de Control"
        xlsHoja.Range("W2", "AK2").MergeCells = True
        
        xlsHoja.Cells(lnPosCabecera, 38) = "Monto Minimo"
        xlsHoja.Range("AL2", "AZ2").MergeCells = True
    'Cabecera Datos
        lnPosCabecera = lnPosCabecera + 1
        xlsHoja.Range(xlsHoja.Cells(3, 1), xlsHoja.Cells(3, nFinal)).HorizontalAlignment = xlRight
        xlsHoja.Cells(lnPosCabecera, 1).NumberFormat = "@"
        xlsHoja.Range("A3", "D3").MergeCells = True
        xlsHoja.Cells(lnPosCabecera, 1) = "0106"
        
        xlsHoja.Cells(lnPosCabecera, 5).NumberFormat = "@"
        xlsHoja.Range("E3", "F3").MergeCells = True
        xlsHoja.Cells(lnPosCabecera, 5) = "01"
        
        xlsHoja.Cells(lnPosCabecera, 7).NumberFormat = "@"
        xlsHoja.Range("G3", "K3").MergeCells = True
        xlsHoja.Cells(lnPosCabecera, 7) = FillNum(Trim(gsCodEmpInfICC), 5, "0")
                
        xlsHoja.Cells(lnPosCabecera, 12).NumberFormat = "@"
        xlsHoja.Range("L3", "S3").MergeCells = True
        xlsHoja.Cells(lnPosCabecera, 12) = Format(gdFecData, "yyyymmdd")
        
        xlsHoja.Cells(lnPosCabecera, 20).NumberFormat = "@"
        xlsHoja.Range("T3", "V3").MergeCells = True
        xlsHoja.Cells(lnPosCabecera, 20) = "012"
        
        xlsHoja.Cells(lnPosCabecera, 23).NumberFormat = "0"
        xlsHoja.Range("W3", "AK3").MergeCells = True
        xlsHoja.Cells(lnPosCabecera, 23) = ""
        
        xlsHoja.Cells(lnPosCabecera, 38).NumberFormat = "@"
        xlsHoja.Range("AL3", "AZ3").MergeCells = True
        xlsHoja.Cells(lnPosCabecera, 38) = String(15 - Len(lsMontoMinimo), "0") & lsMontoMinimo
        
        ' Descripción de Identificación y Saldos
        'Identificacion
        xlsHoja.Range("A4", "RG" & 4).Borders.LineStyle = xlContinuous
        xlsHoja.Range("A4", "RG" & 4).Borders.Weight = xlThin
        xlsHoja.Range("A4", "RG" & 4).Borders.ColorIndex = xlAutomatic
        xlsHoja.Range("A4", "RG" & 4).Interior.Color = RGB(221, 235, 247)
        xlsHoja.Range("A4", "RG" & 4).Font.Color = RGB(0, 32, 96)
        xlsHoja.Range(xlsHoja.Cells(4, 1), xlsHoja.Cells(4, nFinal)).HorizontalAlignment = xlCenter
        xlsHoja.Range(xlsHoja.Cells(4, 1), xlsHoja.Cells(4, nFinal)).Font.Bold = True
        
        xlsHoja.Cells(lnPosIdentificador, 1) = "TF"

        xlsHoja.Cells(lnPosIdentificador, 2) = "TI"

        xlsHoja.Cells(lnPosIdentificador, 3) = "Num. Secuencia"
        xlsHoja.Range("C4", "J4").MergeCells = True
        
        xlsHoja.Cells(lnPosIdentificador, 11) = "Código SBS"
        xlsHoja.Range("K4", "T4").MergeCells = True
        
        xlsHoja.Cells(lnPosIdentificador, 21) = "Código DE Persona CMAC"
        xlsHoja.Range("U4", "AN4").MergeCells = True
        
        xlsHoja.Cells(lnPosIdentificador, 41) = "CIIU"
        xlsHoja.Range("AO4", "AR4").MergeCells = True
        
        xlsHoja.Cells(lnPosIdentificador, 45) = "ZR"
        xlsHoja.Range("AS4", "AT4").MergeCells = True
        
        xlsHoja.Cells(lnPosIdentificador, 47) = "OR"
        xlsHoja.Range("AU4", "AV4").MergeCells = True
        
        xlsHoja.Cells(lnPosIdentificador, 49) = "TI"
        xlsHoja.Range("AW4", "AW4").MergeCells = True
        
        xlsHoja.Cells(lnPosIdentificador, 50) = "Numero de partida o Ficha"
        xlsHoja.Range("AX4", "BG4").MergeCells = True
        
        xlsHoja.Cells(lnPosIdentificador, 60) = "TD"
        
        xlsHoja.Cells(lnPosIdentificador, 61) = "Documento Tributario"
        xlsHoja.Range("BI4", "BS4").MergeCells = True
        
        xlsHoja.Cells(lnPosIdentificador, 72) = "TD"
        
        xlsHoja.Cells(lnPosIdentificador, 73) = "Documento de Identidad"
        xlsHoja.Range("BU4", "CF4").MergeCells = True
        
        xlsHoja.Cells(lnPosIdentificador, 85) = "TP"
        xlsHoja.Cells(lnPosIdentificador, 86) = "R"
        xlsHoja.Cells(lnPosIdentificador, 87) = "CD"
        xlsHoja.Cells(lnPosIdentificador, 88) = "M"
        xlsHoja.Cells(lnPosIdentificador, 89) = "A"
        xlsHoja.Cells(lnPosIdentificador, 90) = "RL"
        xlsHoja.Cells(lnPosIdentificador, 91) = "Pais Reside"
        xlsHoja.Range("CM4", "CP4").MergeCells = True
        xlsHoja.Cells(lnPosIdentificador, 95) = "G"
        xlsHoja.Cells(lnPosIdentificador, 96) = "EC"
        xlsHoja.Cells(lnPosIdentificador, 97) = "Sigla"
        xlsHoja.Range("CS4", "DL4").MergeCells = True
        xlsHoja.Cells(lnPosIdentificador, 117) = "Apellido Paterno o Razón Social"
        xlsHoja.Range("DM4", "IB4").MergeCells = True
        xlsHoja.Cells(lnPosIdentificador, 237) = "Apellido Materno"
        xlsHoja.Range("IC4", "JP4").MergeCells = True
        xlsHoja.Cells(lnPosIdentificador, 277) = "Apellido de Casada"
        xlsHoja.Range("JQ4", "LD4").MergeCells = True
        
        xlsHoja.Cells(lnPosIdentificador, 317) = "Primer Nombre"
        xlsHoja.Range("LE4", "MR4").MergeCells = True
        
        xlsHoja.Cells(lnPosIdentificador, 357) = "Segundo Nombre"
        xlsHoja.Range("MS4", "OF4").MergeCells = True
        
        xlsHoja.Cells(lnPosIdentificador, 397) = "RC"
        xlsHoja.Cells(lnPosIdentificador, 398) = "IA"
        
        xlsHoja.Cells(lnPosIdentificador, 399) = "Clasificacion"
        xlsHoja.Range("OI4", "OM4").MergeCells = True
        
        xlsHoja.Cells(lnPosIdentificador, 404) = "CA"

        xlsHoja.Cells(lnPosIdentificador, 405) = "Grupo Econonomico"
        xlsHoja.Range("OO4", "PH4").MergeCells = True
        
        xlsHoja.Cells(lnPosIdentificador, 425) = "Fecha Nacimiento"
        xlsHoja.Range("PI4", "PP4").MergeCells = True
        
        xlsHoja.Cells(lnPosIdentificador, 433) = "TI"
        xlsHoja.Range("PQ4", "PR4").MergeCells = True
        
        xlsHoja.Cells(lnPosIdentificador, 435) = "Documento Identidad Complementario"
        xlsHoja.Range("PS4", "RF4").MergeCells = True
        
        
        xlsHoja.Cells(lnPosIdentificador, 475) = "CE"
        'Saldos
        xlsHoja.Range("A5", "BL" & 5).Borders.LineStyle = xlContinuous
        xlsHoja.Range("A5", "BL" & 5).Borders.Weight = xlThin
        xlsHoja.Range("A5", "BL" & 5).Borders.ColorIndex = xlAutomatic
        xlsHoja.Range("A5", "BL" & 5).Interior.Color = RGB(221, 235, 247)
        xlsHoja.Range("A5", "BL" & 5).Font.Color = RGB(0, 32, 96)
        xlsHoja.Range(xlsHoja.Cells(5, 1), xlsHoja.Cells(5, 64)).HorizontalAlignment = xlCenter
        xlsHoja.Range(xlsHoja.Cells(5, 1), xlsHoja.Cells(5, 64)).Font.Bold = True
        
        xlsHoja.Cells(lnPosSaldos, 1) = "TF"
        xlsHoja.Cells(lnPosSaldos, 2) = "TI"
        xlsHoja.Cells(lnPosSaldos, 3) = "Num. Secuencia"
        xlsHoja.Range("C5", "J5").MergeCells = True
        xlsHoja.Cells(lnPosSaldos, 11) = "Codigo OE"
        xlsHoja.Range("K5", "N5").MergeCells = True
        xlsHoja.Cells(lnPosSaldos, 15) = "Ubicacion Geogra."
        xlsHoja.Range("O5", "T5").MergeCells = True
        xlsHoja.Cells(lnPosSaldos, 21) = "Codigo de Cuenta Contable"
        xlsHoja.Range("U5", "AH5").MergeCells = True
        xlsHoja.Cells(lnPosSaldos, 35) = "TC"
        xlsHoja.Range("AI5", "AJ5").MergeCells = True
        xlsHoja.Cells(lnPosSaldos, 37) = "Saldo"
        xlsHoja.Range("AK5", "BB5").MergeCells = True
        xlsHoja.Cells(lnPosSaldos, 55) = "Cond. Dias"
        xlsHoja.Range("BC5", "BF5").MergeCells = True
        xlsHoja.Cells(lnPosSaldos, 59) = "CEC"
        xlsHoja.Range("BG5", "BH5").MergeCells = True
        xlsHoja.Cells(lnPosSaldos, 61) = "CD"
        xlsHoja.Range("BI5", "BJ5").MergeCells = True
        xlsHoja.Cells(lnPosSaldos, 63) = "FCC"
        xlsHoja.Range("BK5", "BL5").MergeCells = True
' Fin de excell

    
    lnPosSaldos = lnPosSaldos + 1
    lnPosIdentificador = lnPosIdentificador + 1
    Do While Not rs1.EOF
        J = J + 1

        lnPosSaldos = lnPosSaldos + 1
        If lsPersonaAct <> rs1!cCodPers Then
            lnPosIdentificador = lnPosIdentificador + 1
            If lnPosIdentificador > 6 Then
            lnPosSaldos = lnPosSaldos + 1
            End If
            Cont = Cont + 1
            
            lsNumSecICC = rs1!cNumSec
            
            lsApePat = fgReemplazaCaracterEspecial(rs1!cApePat)
            lsApeMat = fgReemplazaCaracterEspecial(rs1!capemat)
            lsApeCasada = fgReemplazaCaracterEspecial(rs1!cApeCasada)
            lsNomPri = fgReemplazaCaracterEspecial(rs1!cNomPri)
            lsNomSeg = fgReemplazaCaracterEspecial(rs1!cNomSeg)
            
            Select Case Trim(rs1!cTipPers)
                Case "1", "3"
                    lscTidoci = IIf(IsNull(rs1!ctidoci), "", IIf(Len(Trim(rs1!ctidoci)) = 0, "", Trim(rs1!ctidoci)))
                    lscNudoCi = IIf(IsNull(rs1!cnudoci), "", IIf(Len(Trim(rs1!cnudoci)) = 0, "", Trim(rs1!cnudoci)))
                    lscTidoTr = ""
                    lscNudotr = ""
                Case Else
                    lscTidoci = ""
                    lscNudoCi = ""
                    lscTidoTr = IIf(IsNull(rs1!cTidoTr), "", IIf(Len(Trim(rs1!cTidoTr)) = 0, "", Trim(rs1!cTidoTr)))
                    lscNudotr = IIf(IsNull(rs1!cNudoTr), "", IIf(Len(Trim(rs1!cNudoTr)) = 0, "", Trim(rs1!cNudoTr)))
                    ' RUC DE 11 DIGITOS
                    lscTidoTr = IIf(lscTidoTr = "2", "3", lscTidoTr)
            End Select
            If Not IsNull(rs1!ccodsbs) Then
                lsCodSbs = Trim(rs1!ccodsbs)
            Else
                lsCodSbs = ""
            End If
            If Not IsNull(rs1!cActEcon) Then
                If Len(Trim(rs1!cActEcon)) > 0 Then
                    lsActividad = FillNum(rs1!cActEcon, 4, "0")
                Else
                    lsActividad = ""
                End If
            Else
                lsActividad = ""
            End If
     
            lsCodPers = Trim(rs1!cCodPers)

            lsCalificacion = Trim(rs1!cCalifica)
            lsRiesCambiarioPersona = Trim(rs1!cRCambiario)
            
            lsIndicadorAtr = Trim(rs1!cIndicadorAtraso)
            lsCalifInterna = Trim(IIf(IsNull(rs1!cCalifInterna), "X", rs1!cCalifInterna))
            If Trim(lsCalifInterna) = "X" Then
                lsCalifInterna = "XXXXX"
            End If
            
      Dim cMagEmp As String
        If Trim(rs1!cMagEmp) = "0" Then
            cMagEmp = "1"
        ElseIf Trim(rs1!cMagEmp) = "1" Then
            cMagEmp = "6"
        ElseIf Trim(rs1!cMagEmp) = "2" Then
            cMagEmp = "7"
        ElseIf Trim(rs1!cMagEmp) = "3" Then
            cMagEmp = "8"
        ElseIf Trim(rs1!cMagEmp) = "4" Then
            cMagEmp = "0"
        ElseIf Trim(rs1!cMagEmp) = "5" Then
            cMagEmp = "5"
        End If
                    'rs1!cTipoFor
                    'rs1!cTipoInf
                    'FillNum(Trim(lsNumSecICC), 8, " ")
                    'ImpreFormat(IIf(Len(Trim(lsCodSbs)) = 0, "0000000000", lsCodSbs), 10, 0)
                    'ImpreFormat(Trim(rs1!cCodPers), 20, 0) & FillNum(lsActividad, 4, " ")
                    'FillNum(Trim(IIf(IsNull(rs1!ccodregpub), "", rs1!ccodregpub)), 15, " ")
                    'FillNum(IIf(Len(Trim(lscTidoTr)) = 0, "", lscTidoTr), 1, " ")
                    'FillNum(IIf(Len(Trim(lscNudotr)) = 0, "", lscNudotr), 11, " ")
                    'FillNum(IIf(Len(Trim(lscTidoci)) = 0, "", lscTidoci), 1, " ")
                    'ImpreFormat(IIf(Len(Trim(lscNudoCi)) = 0, "", lscNudoCi), 12, 0)
                    'FillNum(Trim(IIf(IsNull(rs1!cTipPers), "", rs1!cTipPers)), 1, " ")
                    'FillNum(Trim(IIf(IsNull(rs1!cResid), "", rs1!cResid)), 1, " ")
                    'FillNum(Trim(lsCalificacion), 1, " ")
                    'FillNum(Trim(IIf(IsNull(cMagEmp), "", cMagEmp)), 1, " ")
                    'FillNum(Trim(IIf(IsNull(rs1!cAccionista), "", rs1!cAccionista)), 1, " ")
                    'FillNum(Trim(IIf(IsNull(rs1!cRelInst), "", rs1!cRelInst)), 1, " ")
                    'FillNum(Trim(rs1!cPaisNac), 4, "  ")
                    'FillNum(Trim(rs1!cGenero), 1, " ")
                    'FillNum(Trim(rs1!cEstadoCiv), 1, " ")
                    'ImpreFormat(Trim(IIf(IsNull(rs1!cSiglas), "", rs1!cSiglas)), 20, 0)
                    'ImpreFormat(Trim(lsApePat), 120, 0)
                    'ImpreFormat(Trim(lsApeMat), 40, 0)
                    'ImpreFormat(Trim(lsApeCasada), 40, 0)
                    'ImpreFormat(Trim(lsNomPri), 40, 0)
                    'ImpreFormat(Trim(lsNomSeg), 40, 0)
                    'FillNum(Trim(lsRiesCambiarioPersona), 1, " ")
                    'FillNum(Trim(lsIndicadorAtr), 1, " ")
                    'FillNum(lsCalifInterna, 5, " ")
                    'FillNum(Trim(rs1!cCaliSinAlinea), 1, " ")
                    'FillNum(Trim(rs1!cPersCodGrupEco), 20, " ")
                    'ImpreFormat(IIf(Trim(IIf(IsNull(rs1!cTipPers), "1", rs1!cTipPers)) = "1", Trim(rs1!dFecNac), "        "), 8, 0)
                    'FillNum(Trim(rs1!cTiDociComp), 2, "  ")
                    'FillNum(Trim(rs1!cNuDociComp), IIf(Trim(rs1!cTiDociComp) = "05", 11, 12), IIf(Trim(rs1!cTiDociComp) = "05", "            ", "           "))
                    'IIf(Trim(rs1!cTiDociComp) = "05", " ", "") & Space(28) &
                    'FillNum(Trim(rs1!cCondEndeuda), 1, " ")
                    'ImpreFormat("", 14, 0)
                    
                    '; Chr(13) & Chr(10); 'JUEZ 20130809 rs1!cCondEndeuda
                    
                    
                    'CODIGO DEL EXCEL
                    xlsHoja.Cells(lnPosIdentificador, 1) = rs1!cTipoFor '"TF"

                    xlsHoja.Cells(lnPosIdentificador, 2) = rs1!cTipoInf '"TI"
                       
                    xlsHoja.Cells(lnPosIdentificador, 3).NumberFormat = "@"
                    xlsHoja.Cells(lnPosIdentificador, 3) = FillNum(Trim(lsNumSecICC), 8, " ") '"Num. Secuencia"
                    xlsHoja.Range("C" & lnPosIdentificador & "", "J" & lnPosIdentificador & "").MergeCells = True
                    
                    xlsHoja.Cells(lnPosIdentificador, 11).NumberFormat = "@"
                    xlsHoja.Cells(lnPosIdentificador, 11) = ImpreFormat(IIf(Len(Trim(lsCodSbs)) = 0, "0000000000", lsCodSbs), 10, 0) '"Código SBS"
                    xlsHoja.Range("K" & lnPosIdentificador & "", "T" & lnPosIdentificador & "").MergeCells = True
                    
                    xlsHoja.Cells(lnPosIdentificador, 21).NumberFormat = "@"
                    xlsHoja.Cells(lnPosIdentificador, 21) = ImpreFormat(Trim(rs1!cCodPers), 20, 0)  '"Código DE Persona CMAC"
                    xlsHoja.Range("U" & lnPosIdentificador & "", "AN" & lnPosIdentificador & "").MergeCells = True
                    
                    xlsHoja.Cells(lnPosIdentificador, 41).NumberFormat = "@"
                    xlsHoja.Cells(lnPosIdentificador, 41) = FillNum(lsActividad, 4, " ") '"CIIU"
                    xlsHoja.Range("AO" & lnPosIdentificador & "", "AR" & lnPosIdentificador & "").MergeCells = True
                    
                    xlsHoja.Cells(lnPosIdentificador, 45).NumberFormat = "@" '*** Inicio->Codigo Registro de personas Juridicas 45 - 59
                    xlsHoja.Cells(lnPosIdentificador, 45) = FillNum(Trim(IIf(IsNull(rs1!ccodregpub), "", rs1!ccodregpub)), 15, " ") '"ZR"
                    xlsHoja.Range("AS" & lnPosIdentificador & "", "BG" & lnPosIdentificador & "").MergeCells = True
                    
                    'xlsHoja.Cells(lnPosIdentificador, 47) = ""  '"OR"
                    'xlsHoja.Range("AU" & lnPosIdentificador & "", "AV" & lnPosIdentificador & "").MergeCells = True
                    
                    'xlsHoja.Cells(lnPosIdentificador, 49) = ""  '"TI"
                    'xlsHoja.Range("AW" & lnPosIdentificador & "", "AW" & lnPosIdentificador & "").MergeCells = True
                    
                    'xlsHoja.Cells(lnPosIdentificador, 50).NumberFormat = "@" '*** Fin->Codigo Registro de personas Juridicas 45 - 59
                    'xlsHoja.Cells(lnPosIdentificador, 50) = "" '"Numero de partida o Ficha"
                    'xlsHoja.Range("AX4", "BG4").MergeCells = True


                    xlsHoja.Cells(lnPosIdentificador, 60) = FillNum(IIf(Len(Trim(lscTidoTr)) = 0, "", lscTidoTr), 1, " ") '"TD" - Tipo Documento Tributario

                    xlsHoja.Cells(lnPosIdentificador, 61) = FillNum(IIf(Len(Trim(lscNudotr)) = 0, "", lscNudotr), 11, " ") '"Documento Tributario"
                    xlsHoja.Range("BI" & lnPosIdentificador & "", "BS" & lnPosIdentificador & "").MergeCells = True

                    xlsHoja.Cells(lnPosIdentificador, 72) = FillNum(IIf(Len(Trim(lscTidoci)) = 0, "", lscTidoci), 1, " ") '"TD"

                    xlsHoja.Cells(lnPosIdentificador, 73) = ImpreFormat(IIf(Len(Trim(lscNudoCi)) = 0, "", lscNudoCi), 12, 0) '"Documento de Identidad"
                    xlsHoja.Range("BU" & lnPosIdentificador & "", "CF" & lnPosIdentificador & "").MergeCells = True

                    xlsHoja.Cells(lnPosIdentificador, 85) = FillNum(Trim(IIf(IsNull(rs1!cTipPers), "", rs1!cTipPers)), 1, " ") '"TP"
                    xlsHoja.Cells(lnPosIdentificador, 86) = FillNum(Trim(IIf(IsNull(rs1!cResid), "", rs1!cResid)), 1, " ") '"R"
                    xlsHoja.Cells(lnPosIdentificador, 87) = FillNum(Trim(lsCalificacion), 1, " ") '"CD"
                    xlsHoja.Cells(lnPosIdentificador, 88) = FillNum(Trim(IIf(IsNull(cMagEmp), "", cMagEmp)), 1, " ") '"M"
                    xlsHoja.Cells(lnPosIdentificador, 89) = FillNum(Trim(IIf(IsNull(rs1!cAccionista), "", rs1!cAccionista)), 1, " ") '"A"
                    xlsHoja.Cells(lnPosIdentificador, 90) = FillNum(Trim(IIf(IsNull(rs1!cRelInst), "", rs1!cRelInst)), 1, " ") '"RL"
                    xlsHoja.Cells(lnPosIdentificador, 91) = FillNum(Trim(rs1!cPaisNac), 4, "  ") ' "Pais Reside"
                    xlsHoja.Range("CM" & lnPosIdentificador & "", "CP" & lnPosIdentificador & "").MergeCells = True
                    xlsHoja.Cells(lnPosIdentificador, 95) = FillNum(Trim(rs1!cGenero), 1, " ") '"G"
                    xlsHoja.Cells(lnPosIdentificador, 96) = FillNum(Trim(rs1!cEstadoCiv), 1, " ") '"EC"
                    xlsHoja.Cells(lnPosIdentificador, 97) = ImpreFormat(Trim(IIf(IsNull(rs1!cSiglas), "", rs1!cSiglas)), 20, 0) '"Sigla"
                    xlsHoja.Range("CS" & lnPosIdentificador & "", "DL" & lnPosIdentificador & "").MergeCells = True
                    xlsHoja.Cells(lnPosIdentificador, 117) = ImpreFormat(Trim(lsApePat), 120, 0) '"Apellido Paterno o Razón Social"
                    xlsHoja.Range("DM" & lnPosIdentificador & "", "IB" & lnPosIdentificador & "").MergeCells = True
                    xlsHoja.Cells(lnPosIdentificador, 237) = ImpreFormat(Trim(lsApeMat), 40, 0) '"Apellido Materno"
                    xlsHoja.Range("IC" & lnPosIdentificador & "", "JP" & lnPosIdentificador & "").MergeCells = True
                    xlsHoja.Cells(lnPosIdentificador, 277) = ImpreFormat(Trim(lsApeCasada), 40, 0) '"Apellido de Casada"
                    xlsHoja.Range("JQ" & lnPosIdentificador & "", "LD" & lnPosIdentificador & "").MergeCells = True

                    xlsHoja.Cells(lnPosIdentificador, 317) = ImpreFormat(Trim(lsNomPri), 40, 0) '"Primer Nombre"
                    xlsHoja.Range("LE" & lnPosIdentificador & "", "MR" & lnPosIdentificador & "").MergeCells = True
                    xlsHoja.Cells(lnPosIdentificador, 357) = ImpreFormat(Trim(lsNomSeg), 40, 0) ' "Segundo Nombre"
                    xlsHoja.Range("MS" & lnPosIdentificador & "", "OF" & lnPosIdentificador & "").MergeCells = True
                    xlsHoja.Cells(lnPosIdentificador, 397) = FillNum(Trim(lsRiesCambiarioPersona), 1, " ") '"RC"
                    xlsHoja.Cells(lnPosIdentificador, 398) = FillNum(Trim(lsIndicadorAtr), 1, " ") '"IA"
                    
                    xlsHoja.Cells(lnPosIdentificador, 399) = FillNum(lsCalifInterna, 5, " ") '"Clasificacion"
                    xlsHoja.Range("OI" & lnPosIdentificador & "", "OM" & lnPosIdentificador & "").MergeCells = True
                    
                    xlsHoja.Cells(lnPosIdentificador, 404) = FillNum(Trim(rs1!cCaliSinAlinea), 1, " ") '"CA"
                    
                    
                    xlsHoja.Cells(lnPosIdentificador, 405) = FillNum(Trim(rs1!cPersCodGrupEco), 20, " ") '"Grupo Econonomico"
                    xlsHoja.Range("OO" & lnPosIdentificador & "", "PH" & lnPosIdentificador & "").MergeCells = True
                    
                    xlsHoja.Cells(lnPosIdentificador, 425) = ImpreFormat(IIf(Trim(IIf(IsNull(rs1!cTipPers), "1", rs1!cTipPers)) = "1", Trim(rs1!dFecNac), "        "), 8, 0) '"Fecha Nacimiento"
                    xlsHoja.Range("PI" & lnPosIdentificador & "", "PP" & lnPosIdentificador & "").MergeCells = True
                    
                    xlsHoja.Cells(lnPosIdentificador, 433) = FillNum(Trim(rs1!cTiDociComp), 2, "  ") '"TI"
                    xlsHoja.Range("PQ" & lnPosIdentificador & "", "PR" & lnPosIdentificador & "").MergeCells = True
                    
                    xlsHoja.Cells(lnPosIdentificador, 435) = FillNum(Trim(rs1!cNuDociComp), IIf(Trim(rs1!cTiDociComp) = "05", 11, 12), IIf(Trim(rs1!cTiDociComp) = "05", "            ", "           ")) '"Documento Identidad Complementario"
                    xlsHoja.Range("PS" & lnPosIdentificador & "", "RF" & lnPosIdentificador & "").MergeCells = True
                    
                    xlsHoja.Cells(lnPosIdentificador, 475) = FillNum(Trim(rs1!cCondEndeuda), 1, " ") '"CE"
                    ' CODIGO DEL EXCELL
        End If
        
'        If Val(rs1!cCondEspCta) <= 3 Then
            lsCondEspCta = FillNum(Trim(rs1!cCondEspCta), 2, " ")
'        Else
'            Select Case lsRiesCambiarioPersona  ' 2006/03/15 susy
'                Case "0"
'                    lsCondEspCta = "13"
'                Case "1"
'                    lsCondEspCta = "11"
'                Case "2"
'                    lsCondEspCta = "12"
'            End Select
'        End If

                lsFactorConversion = IIf(Mid(rs1!cCtaCnt, 1, 2) = "71", "02", "99")
                'rs1!cTipoFor2
                'rs1!cTipoInf2 &
                'FillNum(Trim(lsNumSecICC), 8, " ") & _
                'FillNum(Trim(rs1!cCodAge), 4, "0")
                'FillNum(Trim(rs1!cUbicGeo), 6, " ") & _
                'FormatoCtaContable(rs1!cctacnt)
                'FillNum(Trim(rs1!cTipoCred), 2, "0") & _
                'ImpreFormat(EliminaPunto(IIf(IsNull(rs1!nSaldo), 0, rs1!nSaldo)), 18, 0, False) & _
                'FillNum(Trim(rs1!nCondDias), 4, " ")
                'FillNum(Trim(lsCondEspCta), 2, " ") & _
                'FillNum(Trim(rs1!cCondDisponib), 2, " ")
                'FillNum(lsFactorConversion, 2, "  ")
                'FillNum("", 382, " "); Chr(13) & Chr(10);

                    'CODIGO EXCEL PARA SALDOS
                    xlsHoja.Cells(lnPosSaldos, 1) = rs1!cTipoFor2 '"TF"
                    xlsHoja.Cells(lnPosSaldos, 2) = rs1!cTipoInf2 '"TI"
                    xlsHoja.Cells(lnPosSaldos, 3).NumberFormat = "@"
                    xlsHoja.Cells(lnPosSaldos, 3) = FillNum(Trim(lsNumSecICC), 8, " ") '"Num. Secuencia"
                    xlsHoja.Range("C" & lnPosSaldos & "", "J" & lnPosSaldos & "").MergeCells = True
                    xlsHoja.Cells(lnPosSaldos, 11).NumberFormat = "@"
                    xlsHoja.Cells(lnPosSaldos, 11) = FillNum(Trim(rs1!cCodAge), 4, "0") '"Codigo OE"
                    xlsHoja.Range("K" & lnPosSaldos & "", "N" & lnPosSaldos & "").MergeCells = True
                    xlsHoja.Cells(lnPosSaldos, 15).NumberFormat = "@"
                    xlsHoja.Cells(lnPosSaldos, 15) = FillNum(Trim(rs1!cUbicGeo), 6, " ") '"Ubicacion Geogra."
                    xlsHoja.Range("O" & lnPosSaldos & "", "T" & lnPosSaldos & "").MergeCells = True
                    xlsHoja.Cells(lnPosSaldos, 21).NumberFormat = "@"
                    xlsHoja.Cells(lnPosSaldos, 21) = FormatoCtaContable(rs1!cCtaCnt) '"Codigo de Cuenta Contable"
                    xlsHoja.Range("U" & lnPosSaldos & "", "AH" & lnPosSaldos & "").MergeCells = True
                    xlsHoja.Cells(lnPosSaldos, 35) = FillNum(Trim(rs1!cTipoCred), 2, "0") '"TC"
                    xlsHoja.Range("AI" & lnPosSaldos & "", "AJ" & lnPosSaldos & "").MergeCells = True
                    xlsHoja.Cells(lnPosSaldos, 37) = ImpreFormat(EliminaPunto(IIf(IsNull(rs1!nSaldo), 0, rs1!nSaldo)), 18, 0, False) '"Saldo"
                    xlsHoja.Range("AK" & lnPosSaldos & "", "BB" & lnPosSaldos & "").MergeCells = True
                    xlsHoja.Cells(lnPosSaldos, 55) = FillNum(Trim(rs1!nCondDias), 4, " ") '"Cond. Dias"
                    xlsHoja.Range("BC" & lnPosSaldos & "", "BF" & lnPosSaldos & "").MergeCells = True
                    xlsHoja.Cells(lnPosSaldos, 59) = FillNum(Trim(lsCondEspCta), 2, " ") '"CEC"
                    xlsHoja.Range("BG" & lnPosSaldos & "", "BH" & lnPosSaldos & "").MergeCells = True
                    
                    xlsHoja.Cells(lnPosSaldos, 61) = FillNum(Trim(rs1!cCondDisponib), 2, " ") '"CD"
                    xlsHoja.Range("BI" & lnPosSaldos & "", "BJ" & lnPosSaldos & "").MergeCells = True
                    
                    xlsHoja.Cells(lnPosSaldos, 63) = FillNum(lsFactorConversion, 2, "  ") '"FC"
                    xlsHoja.Range("BK" & lnPosSaldos & "", "BL" & lnPosSaldos & "").MergeCells = True
                    ' FIN DE CODIGOS

        lnPosIdentificador = lnPosIdentificador + 1

        lblAvance.Caption = "Avance :" & Format(J / contTotal * 100, "#0.000") & "%"
        
        lsPersonaAct = rs1!cCodPers
        
        rs1.MoveNext
        DoEvents
    Loop
End If
rs1.Close
Set rs1 = Nothing

'***************************************
'** Totales de la Empresa
'***************************************

'cont = cont + 1
'lsNumSecICC = FillNum(Trim(str(cont)), 8, "0")
'
'Print #NúmeroArchivo, "2" & "1" & FillNum(Trim(lsNumSecICC), 8, " ") & _
'        Space(10) & Chr(13) & Chr(10);
'
''Actualiza la secuencia
''SQL1 = "Update dbconsolidada..RCDvc" & Format(gdFecData, "yyyymm") & "03 SET cNumSec ='" & lsNumSecICC & "' "
'SQL1 = "Update dbconsolidada.." & lsNomTabla & Format(gdFecData, "yyyymm") & "03 SET cNumSec ='" & lsNumSecICC & "' " 'JUEZ 20150310
'oConex.Ejecutar (SQL1)
'
''Imprime totales
''JUEZ 20140409 ********************************************
''SQL1 = "Select * from dbconsolidada..RCDvc" & Format(gdFecData, "yyyymm") & "03 Order by cCtaCnt,cTipoCred,nCondDias "
'SQL1 = "Select cTipoFor,cTipoInf,cNumSec,cCtaCnt,cTipoCred,nSaldo, "
'SQL1 = SQL1 & "nCondDias = Case When nCondDias >= 10000 Then 9999 Else nCondDias End,cCondEspCta,cCondDisponib "
''SQL1 = SQL1 & "from dbconsolidada..RCDvc" & Format(gdFecData, "yyyymm") & "03 Order by cCtaCnt,cTipoCred,nCondDias "
'SQL1 = SQL1 & "from dbconsolidada.." & lsNomTabla & Format(gdFecData, "yyyymm") & "03 Order by cCtaCnt,cTipoCred,nCondDias "
''END JUEZ *************************************************
'Set rs1 = oConex.CargaRecordSet(SQL1)
''Dim lsFactorConversion As String
'If Not RSVacio(rs1) Then
'    Do While Not rs1.EOF
'     lsFactorConversion = IIf(Mid(rs1!cctacnt, 1, 2) = "71", "02", "99")
'     Print #NúmeroArchivo, "2" & "2" & FillNum(Trim(lsNumSecICC), 8, " ") & _
'                FillNum(" ", 4, " ") & FillNum(" ", 6, " ") & FormatoCtaContable(rs1!cctacnt) & FillNum(Trim(rs1!cTipoCred), 2, "0") & _
'                ImpreFormat(EliminaPunto(IIf(IsNull(rs1!nSaldo), 0, rs1!nSaldo)), 18, 0, False) & _
'                FillNum(Trim(rs1!nCondDias), 4, " ") & FillNum(Trim(rs1!cCondEspCta), 2, " ") & _
'                FillNum(Trim(rs1!cCondDisponib), 2, " ") & FillNum(lsFactorConversion, 2, "  ") & FillNum("", 382, " "); Chr(13) & Chr(10);
''        Print #NúmeroArchivo, "2" & "2" & FillNum(Trim(lsNumSecICC), 8, " ") & _
''                FillNum(" ", 4, " ") & FillNum(" ", 6, " ") & FormatoCtaContable(rs1!cctacnt) & FillNum(Trim(rs1!cTipoCred), 1, " ") & _
''                ImpreFormat(EliminaPunto(IIf(IsNull(rs1!nSaldo), 0, rs1!nSaldo)), 15, 0, False) & _
''                FillNum(Trim(rs1!nCondDias), 4, " ") & FillNum(Trim(rs1!cCondEspCta), 2, " ") & _
''                FillNum(Trim(rs1!cCondDisponib), 2, " "); Chr(13) & Chr(10);
'        rs1.MoveNext
'    Loop
'End If
'rs1.Close
'Set rs1 = Nothing
'
'Close #NúmeroArchivo   ' Cierra el archivo.
'Screen.MousePointer = 0
MsgBox "Se ha generado el archivo RCDvc00." & lsNumEntidad & " satisfactoriamente. " & Chr(13) & "Terminó : " & Time()
        ' lucv fin
End Sub


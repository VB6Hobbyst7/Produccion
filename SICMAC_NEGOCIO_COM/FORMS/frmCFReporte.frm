VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCFReporte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carta Fianza - Reportes"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmNuevos 
      Caption         =   "Condicion"
      Height          =   675
      Left            =   4920
      TabIndex        =   23
      Top             =   3720
      Visible         =   0   'False
      Width           =   3315
      Begin VB.OptionButton OptNuevos 
         Caption         =   "Solo Nuevos"
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   25
         Top             =   300
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptNuevos 
         Caption         =   "Todos"
         Height          =   360
         Index           =   2
         Left            =   1920
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CheckBox ckConsolidado 
      Caption         =   "Consolidada?"
      Height          =   255
      Left            =   6600
      TabIndex        =   22
      Top             =   4485
      Width           =   1695
   End
   Begin VB.CheckBox chkEnviaExcel 
      Caption         =   "Enviar a Excel"
      Height          =   195
      Left            =   5040
      TabIndex        =   18
      Top             =   4485
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame FrameAgencias 
      Caption         =   "Agencias"
      Height          =   675
      Left            =   4920
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   3315
      Begin VB.OptionButton OptAgencias 
         Caption         =   "Ag. Remotas"
         Height          =   360
         Index           =   1
         Left            =   1200
         TabIndex        =   17
         Top             =   225
         Width           =   1215
      End
      Begin VB.OptionButton OptAgencias 
         Caption         =   "Ag. Local"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   16
         Top             =   300
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame FrameMoneda 
      Caption         =   "Moneda"
      Height          =   675
      Left            =   4920
      TabIndex        =   10
      Top             =   2925
      Visible         =   0   'False
      Width           =   3315
      Begin VB.CheckBox ChkMoneda 
         Caption         =   "Extranjera"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox ChkMoneda 
         Caption         =   "Nacional"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame FrameTipo 
      Caption         =   "Tipo"
      Height          =   1770
      Left            =   4920
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   3315
      Begin VB.CheckBox ChkTipo 
         Caption         =   "Microempresas"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   21
         Top             =   1380
         Width           =   1695
      End
      Begin VB.CheckBox ChkTipo 
         Caption         =   "Pequeñas empresas"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   1095
         Width           =   1875
      End
      Begin VB.CheckBox ChkTipo 
         Caption         =   "Medianas empresas"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   810
         Width           =   1995
      End
      Begin VB.CheckBox ChkTipo 
         Caption         =   "Grandes empresas"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   540
         Width           =   1695
      End
      Begin VB.CheckBox ChkTipo 
         Caption         =   "Corporativos"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.ImageList imglstFiguras 
      Left            =   420
      Top             =   3960
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
            Picture         =   "frmCFReporte.frx":0000
            Key             =   "Padre"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCFReporte.frx":005E
            Key             =   "Hijo"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FramePeriodo 
      Caption         =   "Período"
      Height          =   975
      Left            =   4920
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   3315
      Begin MSMask.MaskEdBox mskPeriodo1Al 
         Height          =   330
         Left            =   480
         TabIndex        =   5
         Top             =   600
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPeriodo1Del 
         Height          =   315
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Del "
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Al "
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   270
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   420
      Left            =   6180
      TabIndex        =   3
      Top             =   4890
      Width           =   1140
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   420
      Left            =   4920
      TabIndex        =   2
      Top             =   4890
      Width           =   1155
   End
   Begin VB.Frame FrameOperaciones 
      Caption         =   "Lista de Reportes"
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4695
      Begin MSComctlLib.TreeView tvwReporte 
         Height          =   4995
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   8811
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
   End
End
Attribute VB_Name = "frmCFReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fnRepoSelec As Long
Dim Agencias As String
Dim nConsolidada As Integer

Private Sub ckConsolidado_Click()
    If ckConsolidado.value = 1 Then
        nConsolidada = 0
    Else
        nConsolidada = 1
    End If
End Sub

Private Sub cmdImprimir_Click()
Dim NodRep  As Node
Dim lsDesc As String
Dim lsSql As String
Dim lsMsj As String

lsMsj = ValidaDatos

If lsMsj <> "" Then
    MsgBox lsMsj, vbInformation, "Alerta"
    Exit Sub
End If

If Not ValidaFecha Then
 If VerificaCheck Then
   Set NodRep = tvwReporte.SelectedItem
   If NodRep Is Nothing Then
     Exit Sub
   End If
   lsDesc = Mid(NodRep.Text, 8, Len(NodRep.Text) - 7)
   fnRepoSelec = CLng(NodRep.Tag)
   Screen.MousePointer = 11
   EjecutaReporte CLng(NodRep.Tag), lsDesc
   Screen.MousePointer = 0
   Set NodRep = Nothing
 Else
    MsgBox "No se ha seleccionado las condiciones", vbInformation, "CMACT"
 End If
Else
    MsgBox "El Periodo de Fecha AL, no puede ser menor", vbInformation, "CMACT"
    mskPeriodo1Al.SetFocus
End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub CleanOption()
    mskPeriodo1Al.Mask = "__/__/____"
    mskPeriodo1Del.Mask = "__/__/____"
    ChkMoneda(0).value = 0
    ChkMoneda(1).value = 0
    ChkTipo(1).value = 0
    ChkTipo(0).value = 0
    OptAgencias(0).value = True
End Sub

Private Function DameAgencias() As String
Dim Agencias As String
Dim lnAge As Integer
Dim est As Integer
est = 0
Agencias = ""
For lnAge = 1 To frmSelectAgencias.List1.ListCount
 If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
   est = est + 1
   If est = 1 Then
    Agencias = "'" & Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2) & "'"
   Else
    Agencias = Agencias & ", " & "'" & Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2) & "'"
   End If
 End If
Next lnAge
DameAgencias = Agencias
End Function


Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Me.mskPeriodo1Del.Text = Format$(gdFecSis, "dd/mm/yyyy")
    Me.mskPeriodo1Al.Text = Format$(gdFecSis, "dd/mm/yyyy")
    nConsolidada = 1 'JACA 20110702
    CargaMenu
End Sub

Private Function ValidaFecha() As Boolean
    ValidaFecha = IIf(Format(Me.mskPeriodo1Del.Text, "dd/mm/yyyy") > Format(Me.mskPeriodo1Al.Text, "dd/mm/yyyy"), True, False)
End Function

Private Sub OptAgencias_Click(index As Integer)
If index = 1 Then
    frmSelectAgencias.Inicio Me
    frmSelectAgencias.Show vbModal
    'Formulario q escoge otra agencia
End If
End Sub

Private Sub tvwReporte_NodeClick(ByVal Node As MSComctlLib.Node)
Dim NodRep  As Node
Dim lsDesc As String

Set NodRep = tvwReporte.SelectedItem
             
If NodRep Is Nothing Then
   Exit Sub
End If

lsDesc = Mid(NodRep.Text, 8, Len(NodRep.Text) - 7)
fnRepoSelec = CLng(NodRep.Tag)
CmdImprimir.Enabled = True

chkEnviaExcel.value = 0
chkEnviaExcel.Enabled = True
Select Case fnRepoSelec
    Case 148001
        Call HabilitaFrame(True, True, True, True, True) 'FRHU20131119
    Case 148002
        Call HabilitaFrame(True, True, True, True, True, True)
    Case 148003
        Call HabilitaFrame(True, True, True, True, True) 'FRHU20131119
    Case 148004
        Call HabilitaFrame(True, True, True, True)
    Case 148005
        Call HabilitaFrame(True, True, True, True)
    'By Capi Acta 035-2007
    Case 148006
        Call HabilitaFrame(True, True, True, True)
    Case 148007
        Call HabilitaFrame(False, True, True, True, True, , True)
    Case 148008 '*****RECO 2013-07-18*********
        Call HabilitaFrame(True, True, True, True, True)
        chkEnviaExcel.value = 1
        chkEnviaExcel.Enabled = False '******END RECO********
    Case 148009
        Call HabilitaFrame(True, True, True, True, True)
        chkEnviaExcel.Enabled = False
Case Else
    Call HabilitaFrame(False, False, False, False)
    CmdImprimir.Enabled = False
End Select

Set NodRep = Nothing
End Sub

Private Sub CargaMenu()
Dim clsGen As DGeneral  'COMDConstSistema.DCOMGeneral
Dim rsUsu As New ADODB.Recordset
Dim sOperacion As String
Dim sOpeCod As String
Dim sOpePadre As String
Dim sOpeHijo As String
Dim sOpeHijito As String
Dim nodOpe As Node
Dim lsTipREP As String
lsTipREP = "14800"
Set clsGen = New DGeneral  'COMDConstSistema.DCOMGeneral 'ARCV 25-10-2006
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

Public Sub HabilitaFrame(ByRef v1 As Boolean, ByRef v2 As Boolean, _
                         ByRef v3 As Boolean, ByRef v4 As Boolean, Optional ByRef v5 As Boolean = False, Optional ByRef v6 As Boolean = False, Optional ByRef v7 As Boolean = False)
With Me
    .FrameAgencias.Visible = v1
    .FrameMoneda.Visible = v2
    .FramePeriodo.Visible = v3
    .FrameTipo.Visible = v4
    .chkEnviaExcel.Visible = v5
    .ckConsolidado.Visible = v6
    .frmNuevos.Visible = v7
End With
End Sub
Private Sub EjecutaReporte(ByVal pnEstado As ColocEstado, ByVal psDescOperacion As String, _
   Optional Mn As Moneda)

Dim loRep As COMNCartaFianza.NCOMCartaFianzaReporte 'NCartaFianzaReporte
Dim lsDestino As String  ' P= Previo // I = Impresora // A = Archivo // E = Excel
Dim lsCadImp As String   'cadena q forma
Dim X As Integer
Dim loPrevio As previo.clsprevio
Dim Agencias As String
Dim lnAge As Integer

Dim lsmensaje As String

Agencias = IIf(OptAgencias(0).value = True, gsCodAge, frmSelectAgencias.RecupAgencias)
If gsCodAge = "" Then
    MsgBox "Usted no es usuario de Esta Agencia...Comuniquese con RRHH", vbInformation, "AVISO"
    Exit Sub
End If
Set loRep = New COMNCartaFianza.NCOMCartaFianzaReporte
    loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis


Select Case pnEstado
        'Reporte Cartas Fianza Emitidas
    Case 148001
        'gsCodAge
        If Me.chkEnviaExcel.Visible And Me.chkEnviaExcel.value = 1 Then ' FRHU 20131119
            If OptAgencias(0).value Then
              
                Call ImprimeCartaFianzaRegistradas(pnEstado, gsCodAge, gdFecSis, gsCodUser, _
                        ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, Agencias, , nConsolidada, Format(mskPeriodo1Del.Text, "yyyy/MM/dd"), Format(mskPeriodo1Al.Text, "yyyy/MM/dd"))
    
            Else
                Call ImprimeCartaFianzaRegistradas(pnEstado, Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), gdFecSis, gsCodUser, _
                        ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, DameAgencias, , nConsolidada, Format(mskPeriodo1Del.Text, "yyyy/MM/dd"), Format(mskPeriodo1Al.Text, "yyyy/MM/dd"))

            End If
                      
        Else
             MsgBox "Marcar la opcion: Enviar a Excel", vbInformation, "Aviso"
             Me.chkEnviaExcel.SetFocus
             Exit Sub
             'FRHU20131126
'             If OptAgencias(0).value Then
'              lsCadImp = loRep.nRepo148001_CartasFianzaEmitidas(pnEstado, gsCodAge, Format(mskPeriodo1Del.Text, "dd/mm/yy"), Format(mskPeriodo1Al.Text, "dd/mm/yy"), gsCodUser, _
'              ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, Agencias, lsmensaje, gImpresora)
'              If Trim(lsmensaje) <> "" Then
'                MsgBox lsmensaje, vbInformation, "Aviso"
'                Exit Sub
'              End If
'            Else
'                lsCadImp = loRep.nRepo148001_CartasFianzaEmitidas(pnEstado, Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), Format(mskPeriodo1Del.Text, "dd/mm/yy"), Format(mskPeriodo1Al.Text, "dd/mm/yy"), gsCodUser, _
'                ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, Agencias, lsmensaje, gImpresora)
'                If Trim(lsmensaje) <> "" Then
'                    MsgBox lsmensaje, vbInformation, "Aviso"
'                Exit Sub
'              End If
'            End If
'            lsDestino = "P"
            
        End If
        
    Case 148002
         If Me.chkEnviaExcel.Visible And Me.chkEnviaExcel.value = 1 Then
            If OptAgencias(0).value Then
                '*************RECO 2013-07-18***********************
                Call ImprimeCartaFianzaVigentes(pnEstado, gsCodAge, gdFecSis, gsCodUser, _
                        ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, Agencias, , nConsolidada, Format(mskPeriodo1Del.Text, "yyyy/MM/dd"), Format(mskPeriodo1Al.Text, "yyyy/MM/dd"))
                '*************END RECO******************************

                'Call ImprimeCartaFianzaVigentes(pnEstado, gsCodAge, gdFecSis, gsCodUser, _
                '        ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, Agencias, , nConsolidada)
            Else
                '**************RECO 2013-07-18***************
                Call ImprimeCartaFianzaVigentes(pnEstado, Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), gdFecSis, gsCodUser, _
                        ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, DameAgencias, , nConsolidada, Format(mskPeriodo1Del.Text, "yyyy/MM/dd"), Format(mskPeriodo1Al.Text, "yyyy/MM/dd"))
                '***************END RECO**********************

                'Call ImprimeCartaFianzaVigentes(pnEstado, Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), gdFecSis, gsCodUser, _
                '       ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, DameAgencias, , nConsolidada)
            End If
        Else
            If OptAgencias(0).value Then
              '*******RECO 2013-07-18******
              lsCadImp = loRep.nRepo148002_CartasFianzaVigentes(pnEstado, gsCodAge, gdFecSis, gsCodUser, _
              ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, Agencias, lsmensaje, gImpresora, Format(mskPeriodo1Del.Text, "yyyy/MM/dd"), Format(mskPeriodo1Al.Text, "yyyy/MM/dd"))
              '*******END RECO *********

            
              'lscadimp = loRep.nRepo148002_CartasFianzaVigentes(pnEstado, gsCodAge, gdFecSis, gsCodUser, _
              'ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, Agencias, lsmensaje, gImpresora)
              If Trim(lsmensaje) <> "" Then
                MsgBox lsmensaje, vbInformation, "Aviso"
                Exit Sub
              End If
            Else
            
                '*******RECO 2013-07-18******
                lsCadImp = loRep.nRepo148002_CartasFianzaVigentes(pnEstado, Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), gdFecSis, gsCodUser, _
                ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, DameAgencias, lsmensaje, gImpresora, Format(mskPeriodo1Del.Text, "yyyy/MM/dd"), Format(mskPeriodo1Al.Text, "yyyy/MM/dd"))
                '*******END RECO *********

              'lscadimp = loRep.nRepo148002_CartasFianzaVigentes(pnEstado, Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), gdFecSis, gsCodUser, _
              'ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, DameAgencias, lsmensaje, gImpresora)
              If Trim(lsmensaje) <> "" Then
                MsgBox lsmensaje, vbInformation, "Aviso"
                Exit Sub
              End If
            End If
            lsDestino = "P"
        End If
    Case 148003
        If Me.chkEnviaExcel.Visible And Me.chkEnviaExcel.value = 1 Then ' FRHU 20131119
            If OptAgencias(0).value Then
              
                Call ImprimeCartaFianzaVigentesPorVencer(pnEstado, gsCodAge, gdFecSis, gsCodUser, _
                        ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, Agencias, , nConsolidada, Format(mskPeriodo1Del.Text, "yyyy/MM/dd"), Format(mskPeriodo1Al.Text, "yyyy/MM/dd"))
    
            Else
                Call ImprimeCartaFianzaVigentesPorVencer(pnEstado, Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), gdFecSis, gsCodUser, _
                        ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, DameAgencias, , nConsolidada, Format(mskPeriodo1Del.Text, "yyyy/MM/dd"), Format(mskPeriodo1Al.Text, "yyyy/MM/dd"))

            End If
                      
        Else
             MsgBox "Marcar la opcion: Enviar a Excel", vbInformation, "Aviso"
             Me.chkEnviaExcel.SetFocus
             Exit Sub
             'FRHU20131126
'                If OptAgencias(0).value Then
'                  lsCadImp = loRep.nRepo148003_CartasFianzasVigentes_x_Vencer(pnEstado, gsCodAge, Format(mskPeriodo1Del.Text, "dd/mm/yy"), Format(mskPeriodo1Al.Text, "dd/mm/yy"), gsCodUser, _
'                  ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, gdFecSis, OptAgencias(0).value, Agencias, lsmensaje, gImpresora)
'                  If Trim(lsmensaje) <> "" Then
'                    MsgBox lsmensaje, vbInformation, "Aviso"
'                    Exit Sub
'                  End If
'                Else
'                  lsCadImp = loRep.nRepo148003_CartasFianzasVigentes_x_Vencer(pnEstado, Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), Format(mskPeriodo1Del.Text, "dd/mm/yy"), Format(mskPeriodo1Al.Text, "dd/mm/yy"), gsCodUser, _
'                  ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, gdFecSis, OptAgencias(0).value, DameAgencias, lsmensaje, gImpresora)
'                  If Trim(lsmensaje) <> "" Then
'                    MsgBox lsmensaje, vbInformation, "Aviso"
'                    Exit Sub
'                  End If
'                End If
'                lsDestino = "P"
        End If
        
    Case 148004
        If OptAgencias(0).value Then
          lsCadImp = loRep.nRepo148004_CartasFianzaHonradas(pnEstado, gsCodAge, Format(mskPeriodo1Del.Text, "dd/mm/yy"), Format(mskPeriodo1Al.Text, "dd/mm/yy"), gsCodUser, _
          ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, Agencias, lsmensaje, gImpresora)
        Else
          lsCadImp = loRep.nRepo148004_CartasFianzaHonradas(pnEstado, Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), Format(mskPeriodo1Del.Text, "dd/mm/yy"), Format(mskPeriodo1Al.Text, "dd/mm/yy"), gsCodUser, _
          ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, DameAgencias, lsmensaje, gImpresora)
        End If
        lsDestino = "P"
            
    Case 148005
        If OptAgencias(0).value Then
          lsCadImp = loRep.nRepo148005_CartasFianzaDevueltas(pnEstado, gsCodAge, Format(mskPeriodo1Del.Text, "dd/mm/yy"), Format(mskPeriodo1Al.Text, "dd/mm/yy"), gsCodUser, _
          ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, Agencias, lsmensaje, gImpresora)
          If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
          End If
        Else
          lsCadImp = loRep.nRepo148005_CartasFianzaDevueltas(pnEstado, gsCodAge, Format(mskPeriodo1Del.Text, "dd/mm/yy"), Format(mskPeriodo1Al.Text, "dd/mm/yy"), gsCodUser, _
          ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, DameAgencias, lsmensaje, gImpresora)
          If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
          End If
        End If
        lsDestino = "P"
    
    'By Capi Acta 035-2007
    Case 148006
        If OptAgencias(0).value Then
          '*** PEAC - se agrego los parametros Format(mskPeriodo1Del.Text, "yyyymmdd"), Format(mskPeriodo1Al.Text, "yyyymmdd")
          lsCadImp = loRep.nRepo148006_CartasFianzaCanceladas(pnEstado, gsCodAge, gdFecSis, gsCodUser, _
          ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, Agencias, lsmensaje, gImpresora, CDate(mskPeriodo1Del.Text), CDate(mskPeriodo1Al.Text))
          
          If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
          End If
        Else
          '*** PEAC - se agrego los parametros Format(mskPeriodo1Del.Text, "yyyymmdd"), Format(mskPeriodo1Al.Text, "yyyymmdd")
          lsCadImp = loRep.nRepo148006_CartasFianzaCanceladas(pnEstado, Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), gdFecSis, gsCodUser, _
          ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, DameAgencias, lsmensaje, gImpresora, CDate(mskPeriodo1Del.Text), CDate(mskPeriodo1Al.Text))
          If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
          End If
        End If
        lsDestino = "P"
        
    'MADM 20111226
     Case 148007
         If Me.chkEnviaExcel.Visible And Me.chkEnviaExcel.value = 1 Then
           Call ImprimeCartaFianzaNuevas(pnEstado, gsCodAge, gdFecSis, gsCodUser, 1, 1, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, IIf(OptNuevos.item(3).value = True, True, False), , lsmensaje, , CDate(mskPeriodo1Del.Text), CDate(mskPeriodo1Al.Text))
         Else
          lsCadImp = loRep.nRepo148007_CartasFianzanuevas(pnEstado, gsCodAge, gdFecSis, gsCodUser, _
          1, 1, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, IIf(OptNuevos.item(3).value = True, True, False), , lsmensaje, gImpresora, CDate(mskPeriodo1Del.Text), CDate(mskPeriodo1Al.Text))
          If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
          End If
       End If
       lsDestino = "P"
    '*******RECO 2013-07-18******
    Case 148008
        If Me.chkEnviaExcel.Visible And Me.chkEnviaExcel.value = 1 Then
            If OptAgencias(0).value Then
                Call ImprimeCartaFianzaConsolidada(pnEstado, gsCodAge, gdFecSis, gsCodUser, _
                        ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, Agencias, , nConsolidada, CDate(mskPeriodo1Del.Text), CDate(mskPeriodo1Al.Text))
            Else
                Call ImprimeCartaFianzaConsolidada(pnEstado, Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), gdFecSis, gsCodUser, _
                        ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, DameAgencias, , nConsolidada, CDate(mskPeriodo1Del.Text), CDate(mskPeriodo1Al.Text))
            End If
        Else
            If OptAgencias(0).value Then
              lsCadImp = loRep.nRepo148008_CartasFianzaConsolidado(pnEstado, gsCodAge, gdFecSis, gsCodUser, _
              ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, Agencias, lsmensaje, gImpresora, CDate(mskPeriodo1Del.Text), CDate(mskPeriodo1Al.Text))
              If Trim(lsmensaje) <> "" Then
                MsgBox lsmensaje, vbInformation, "Aviso"
                Exit Sub
              End If
            Else
              lsCadImp = loRep.nRepo148008_CartasFianzaConsolidado(pnEstado, Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), gdFecSis, gsCodUser, _
              ChkMoneda(0).value, ChkMoneda(1).value, ChkTipo(0).value, ChkTipo(1).value, ChkTipo(2).value, ChkTipo(3).value, ChkTipo(4).value, OptAgencias(0).value, DameAgencias, lsmensaje, gImpresora, CDate(mskPeriodo1Del.Text), CDate(mskPeriodo1Al.Text))
              If Trim(lsmensaje) <> "" Then
                MsgBox lsmensaje, vbInformation, "Aviso"
                Exit Sub
              End If
            End If
            lsDestino = "P"
        End If
    '*******END RECO******
    Case 148009 'RECO20160120 ERS001-2016
            Dim rsDatos As ADODB.Recordset
            Set rsDatos = loRep.nRepo148009_CartasAvisoVencimineto(gsCodAge, ObtieneMoneda, Agencias, ObtieneTipoCred, Format(mskPeriodo1Del.Text, "yyyy/MM/dd"), Format(mskPeriodo1Al.Text, "yyyy/MM/dd"))
            Call GeneraCartaAvisoVencimiento(rsDatos)
            
End Select
    
Set loRep = Nothing
If lsDestino = "P" Then
    If Len(Trim(lsCadImp)) > 0 Then
        Set loPrevio = New previo.clsprevio
            loPrevio.Show Chr$(27) & Chr$(77) & lsCadImp, psDescOperacion, True, , gImpresora
        Set loPrevio = Nothing
    Else
        MsgBox "No Existen Datos para el reporte ", vbInformation, "Aviso"
    End If
ElseIf lsDestino = "A" Then
End If
End Sub

Public Function VerificaCheck() As Boolean

'*** PEAC 20090615

'VerificaCheck = True

If Not IsDate(mskPeriodo1Del.Text) Or Not IsDate(mskPeriodo1Al.Text) Then
    VerificaCheck = False
Else
    VerificaCheck = True
End If
    
    'VerificaCheck = IIf((ChkTipo(0) = 1 Or ChkTipo(1) = 1) And (ChkMoneda(0) = 1 Or ChkMoneda(1) = 1), True, False)
End Function

'*** PEAC 20090615
Private Sub ImprimeCartaFianzaVigentes(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdCMACT As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
        ByRef psMonedaE As Integer, ByRef psTipoCor As Integer, ByRef psTipoGraEmp As Integer, _
        ByRef psTipoMedEmp As Integer, ByRef psTipoPeqEmp As Integer, ByRef psTipoMicEmp As Integer, _
        ByRef OptAge As Boolean, Optional psListaAgencias As String, Optional ByRef psMensaje As String, Optional nConsolidada As Integer = 1 _
        , Optional ByRef dFecIni As String, Optional ByRef dFecFin As String)
        '***RECO 2013-07-18:SE AGREGO DOS PARAMETROS NUEVOS (dFecIni,dFecFin)*

Dim lsSql As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim oFun As New COMFunciones.FCOMImpresion

Dim P As String, sAges As String, sTipo As String, sMone As String
Dim lcAge As String, lcMone As String, lnTot As Double

Dim i As Integer, j As Integer

    If OptAge = True Then
        sAges = psAgencia
    Else
        sAges = Replace(Replace(psListaAgencias, "'", ""), " ", "")
    End If
      
'    If psTipoC = 0 And psTipoM = 1 Then
'        sTipo = "2"
'    ElseIf psTipoC = 1 And psTipoM = 0 Then
'        sTipo = "1"
'    Else
'        sTipo = "1,2"
'    End If

'*** BRGO BASILEA II
    If psTipoCor = 0 And psTipoGraEmp = 0 And psTipoMedEmp = 0 And psTipoPeqEmp = 0 And psTipoMicEmp = 0 Then
        sTipo = "1,2,3,4,5"
    Else
        If psTipoCor = 1 Then
          sTipo = sTipo & ",1"
        End If
        If psTipoGraEmp = 1 Then
          sTipo = sTipo & ",2"
        End If
        If psTipoMedEmp = 1 Then
          sTipo = sTipo & ",3"
        End If
        If psTipoPeqEmp = 1 Then
          sTipo = sTipo & ",4"
        End If
        If psTipoMicEmp = 1 Then
          sTipo = sTipo & ",5"
        End If
        sTipo = Mid(sTipo, 2, Len(sTipo) - 1)
    End If

'**********************

    If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
        sMone = "1"
    ElseIf psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
        sMone = "2"
    Else
        sMone = "1,2"
    End If
    '**********RECO 2013-07-18*********************
    lsSql = " exec stp_sel_ReporteCartasFianzasVigentesExcel '" & sAges & "','" & sMone & "','" & sTipo & "'," & nConsolidada & ",'" & dFecIni & "','" & dFecFin & "'"
    '**********END RECO****************************

    'lsSQL = " exec stp_sel_ReporteCartasFianzasVigentesExcel '" & sAges & "','" & sMone & "','" & sTipo & "'," & nConsolidada

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSql)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psMensaje = " No Existen Datos para el reporte en la Agencia "
        Exit Sub
    End If


'**************************************************************************************
    Dim ApExcel As Variant, lcTipGar As String, lnImporteSOL As Double, lnImporteDOL As Double
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    'ApExcel.Cells(2, 2).Formula = "psNomCmac"
    'ApExcel.Cells(3, 2).Formula = "psNomAge"
    ApExcel.Cells(2, 14).Formula = Date + Time()
    ApExcel.Cells(3, 14).Formula = psUsuarioRep
    ApExcel.Range("B2", "O8").Font.Bold = True
    
    ApExcel.Range("B2", "B3").HorizontalAlignment = xlLeft
    ApExcel.Range("N2", "O3").HorizontalAlignment = xlRight
    ApExcel.Range("B5", "O6").HorizontalAlignment = xlCenter
    ApExcel.Range("B9", "O10").VerticalAlignment = xlCenter
    ApExcel.Range("B9", "R10").Borders.LineStyle = 1
    
    ApExcel.Cells(5, 2).Formula = "CONTROL DE CARTAS FIANZAS VIGENTES"
'    ApExcel.Cells(6, 2).Formula = "AGENCIA - MONEDA"

    ApExcel.Range("B5", "O5").MergeCells = True
    ApExcel.Range("B6", "O6").MergeCells = True

    ApExcel.Range("B9", "B10").MergeCells = True
    ApExcel.Range("C9", "C10").MergeCells = True
    ApExcel.Range("D9", "D10").MergeCells = True
    ApExcel.Range("E9", "E10").MergeCells = True
    ApExcel.Range("F9", "F10").MergeCells = True
    ApExcel.Range("G9", "G10").MergeCells = True
    ApExcel.Range("H9", "H10").MergeCells = True
    ApExcel.Range("I9", "I10").MergeCells = True
    ApExcel.Range("J9", "J10").MergeCells = True
    ApExcel.Range("K9", "K10").MergeCells = True
    ApExcel.Range("L9", "L10").MergeCells = True
    ApExcel.Range("M9", "M9").MergeCells = True
    ApExcel.Range("N9", "O9").MergeCells = True
    ApExcel.Range("P9", "Q9").MergeCells = True
    ApExcel.Range("R9", "R9").MergeCells = True
    
    ApExcel.Cells(9, 2).Formula = "ITEM"
    ApExcel.Cells(9, 3).Formula = "AGENCIA"
    ApExcel.Cells(9, 4).Formula = "CUENTA"
    ApExcel.Cells(9, 5).Formula = "CLIENTE"
    ApExcel.Cells(9, 6).Formula = "TIPO CREDITO"
    ApExcel.Cells(9, 7).Formula = "Nº C.FIANZA"
    ApExcel.Cells(9, 8).Formula = "MODALIDAD CARTA FIANZA"
    ApExcel.Cells(9, 9).Formula = "INICIO"
    ApExcel.Cells(9, 10).Formula = "VENCIMIENTO"
    ApExcel.Cells(9, 11).Formula = "BENEFICIARIO"
    ApExcel.Cells(9, 12).Formula = "IMPORTE DE LA CARTA FIANZA"
    ApExcel.Cells(9, 13).Formula = "GARANTIA"
    
    ApExcel.Cells(9, 14).Formula = "VALOR REALIZACION DE LA GARANTIA"
    ApExcel.Cells(10, 14).Formula = "S/."
    ApExcel.Cells(10, 15).Formula = "US$."
    
    ApExcel.Cells(9, 16).Formula = "VALOR GRAVADO DE LA GARANTIA"
    ApExcel.Cells(10, 16).Formula = "S/."
    ApExcel.Cells(10, 17).Formula = "US$."
    ApExcel.Cells(9, 18).Formula = "RENOVADO"
    
    ApExcel.Range("B9", "r10").Font.Bold = True
    ApExcel.Range("B9", "r10").HorizontalAlignment = 3
    
    ApExcel.Range("A1:Z1000").Font.Size = 8
    
    Dim nTotal As Integer
    Dim nImporte As Double
    Dim nSaldoSoles As Currency
    Dim nSaldoDolares As Currency
    nSaldoSoles = 0
    nSaldoDolares = 0

    i = 10
    j = 0
    Dim nCFAnterior As String
    Do While Not lrDataRep.EOF
        If nCFAnterior = lrDataRep!Carta_Fianza Then
            nSaldoSoles = nSaldoSoles + IIf(lrDataRep!nMonedaCG = "1", lrDataRep!nGravadoCG, 0)
            nSaldoDolares = nSaldoDolares + IIf(lrDataRep!nMonedaCG = "2", lrDataRep!nGravadoCG, 0)
            
            ApExcel.Cells(i, 16).Formula = nSaldoSoles
            ApExcel.Cells(i, 17).Formula = nSaldoDolares
        Else
            nSaldoSoles = 0
            nSaldoDolares = 0
        End If

        If nCFAnterior <> lrDataRep!Carta_Fianza Then
            
            i = i + 1
            j = j + 1
            nCFAnterior = lrDataRep!Carta_Fianza
            
            ApExcel.Cells(i, 2).Formula = "'" & Format(j, "000")
            ApExcel.Cells(i, 3).Formula = lrDataRep!NombAgencia
            ApExcel.Cells(i, 4).Formula = lrDataRep!Carta_Fianza
            ApExcel.Cells(i, 5).Formula = lrDataRep!Cliente
            ApExcel.Cells(i, 6).Formula = lrDataRep!TpoCred
            ApExcel.Cells(i, 7).Formula = "'" & Format(lrDataRep!num_poliza, "00000")
            ApExcel.Cells(i, 8).Formula = lrDataRep!Modalidad
            ApExcel.Cells(i, 9).Formula = "'" & Format(lrDataRep!Fecha_Emision, "dd/mm/yyyy")
            ApExcel.Cells(i, 10).Formula = "'" & Format(lrDataRep!Fecha_Vencimiento, "dd/mm/yyyy")
            ApExcel.Cells(i, 11).Formula = lrDataRep!Acreedor
            ApExcel.Cells(i, 12).Formula = lrDataRep!Importe
            
            'RECO20150921 ERS034-2015 ***************************
            'If lrDataRep!nTpoGarantia = 4 Or lrDataRep!nTpoGarantia2 = 4 Then
            '    lcTipGar = "ART"
            'Else
            '    lcTipGar = IIf(Len(lrDataRep!ctapzofjo) = 18, "DPF", IIf(lrDataRep!valgravhipo > 0, "HIP", IIf(lrDataRep!valgravhipo = 0 And lrDataRep!dCertifGravamen = "01/01/1990" And (lrDataRep!nTpoGarantia <> 0 Or lrDataRep!nTpoGarantia2 <> 0), "EnTram", "SGD")))
            'End If
            lcTipGar = lrDataRep!cTipoValorizacion
            'RECO FIN********************************************
            lnImporteSOL = IIf(lcTipGar = "DPF" And Mid(lrDataRep!ctapzofjo, 9, 1) = "1", lrDataRep!saldopzofjo, IIf(lcTipGar = "HIP" And lrDataRep!nmonehipo = 1, lrDataRep!valgravhipo, 0))
            lnImporteDOL = IIf(lcTipGar = "DPF" And Mid(lrDataRep!ctapzofjo, 9, 1) = "2", lrDataRep!saldopzofjo, IIf(lcTipGar = "HIP" And lrDataRep!nmonehipo = 2, lrDataRep!valgravhipo, 0))
                
            ApExcel.Cells(i, 13).Formula = lcTipGar
            ApExcel.Cells(i, 14).Formula = lnImporteSOL
            ApExcel.Cells(i, 15).Formula = lnImporteDOL
            
            nSaldoSoles = IIf(lrDataRep!nMonedaCG = "1", lrDataRep!nGravadoCG, 0)
            nSaldoDolares = IIf(lrDataRep!nMonedaCG = "2", lrDataRep!nGravadoCG, 0)
            
            ApExcel.Cells(i, 16).Formula = nSaldoSoles
            ApExcel.Cells(i, 17).Formula = nSaldoDolares
            ApExcel.Cells(i, 18).Formula = lrDataRep!Num_Renovacion
            
            ApExcel.Range("K" & Trim(str(i)) & ":" & "Q" & Trim(str(i))).NumberFormat = "#,##0.00"
            ApExcel.Range("B" & Trim(str(i)) & ":" & "R" & Trim(str(i))).Borders.LineStyle = 1
            
            nTotal = nTotal + 1
            nImporte = nImporte + lrDataRep!Importe
        Else
        
            Dim liPosicion As Integer
            nCFAnterior = lrDataRep!Carta_Fianza
            
            If Len(lrDataRep!ctapzofjo) > 0 Then
                liPosicion = InStr(lcTipGar, "DPF")
                If liPosicion = 0 Then
                    lcTipGar = lcTipGar & "+" & "DPF"
                End If
            Else
                If lrDataRep!valgravhipo > 0 Then
                    liPosicion = InStr(lcTipGar, "HIP")
                    If liPosicion = 0 Then
                        lcTipGar = lcTipGar & "+" & "HIP"
                    End If
                Else
                
                    If lrDataRep!dCertifGravamen = "01/01/1990" Then
                        lcTipGar = lcTipGar & "+" & "EnTram"
                    Else
                        liPosicion = InStr(lcTipGar, "SGD")
                        If liPosicion = 0 Then
                            lcTipGar = lcTipGar & "+" & "SGD"
                        End If
                        
                    End If
                    
                End If
            End If
            
            'lcTipGar = IIf(Len(lrDataRep!ctapzofjo) > 0, lcTipGar & " " & "DPF", IIf(lrDataRep!valgravhipo > 0, lcTipGar, ""))
            
            lnImporteSOL = lnImporteSOL + IIf(Mid(lrDataRep!ctapzofjo, 9, 1) = "1", lrDataRep!saldopzofjo, IIf(lrDataRep!nmonehipo = 1, lrDataRep!valgravhipo, 0))
            lnImporteDOL = lnImporteDOL + IIf(Mid(lrDataRep!ctapzofjo, 9, 1) = "2", lrDataRep!saldopzofjo, IIf(lrDataRep!nmonehipo = 2, lrDataRep!valgravhipo, 0))
                    
            ApExcel.Cells(i, 13).Formula = lcTipGar
            ApExcel.Cells(i, 14).Formula = lnImporteSOL
            ApExcel.Cells(i, 15).Formula = lnImporteDOL
            
            ApExcel.Cells(i, 16).Formula = nSaldoSoles
            ApExcel.Cells(i, 17).Formula = nSaldoDolares
            ApExcel.Cells(i, 18).Formula = lrDataRep!Num_Renovacion 'MADM 20111226
                
            ApExcel.Range("K" & Trim(str(i)) & ":" & "Q" & Trim(str(i))).NumberFormat = "#,##0.00"
            ApExcel.Range("B" & Trim(str(i)) & ":" & "R" & Trim(str(i))).Borders.LineStyle = 1
    
        End If
'        If i = 51 Then
'        MsgBox "aaa"
'        End If
    lrDataRep.MoveNext
            
    If lrDataRep.EOF Then
        Exit Do
    End If
    Loop
    
    ApExcel.Cells(1 + i, 5).Formula = "Total de Cartas Fianzas"
    ApExcel.Cells(1 + i, 6).Formula = nTotal
    ApExcel.Cells(1 + i, 11).Formula = Format(nImporte, "#,##0.00")
    ApExcel.Cells(1 + i, 16).Formula = "=SUM(P11:P" & i & ")"
    ApExcel.Cells(1 + i, 17).Formula = "=SUM(Q11:Q" & i & ")"
      
    lrDataRep.Close
    Set lrDataRep = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("O:O").ColumnWidth = 17#
    ApExcel.Columns("N:N").ColumnWidth = 17#
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing
    
End Sub

'MADM 20111226
Private Sub ImprimeCartaFianzaNuevas(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdCMACT As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
        ByRef psMonedaE As Integer, ByRef psTipoCor As Integer, ByRef psTipoGraEmp As Integer, _
        ByRef psTipoMedEmp As Integer, ByRef psTipoPeqEmp As Integer, ByRef psTipoMicEmp As Integer, _
        ByRef OptAge As Boolean, Optional psListaAgencias As String, Optional ByRef psMensaje As String, Optional nConsolidada As Integer = 1, Optional ByVal psDel As Date, Optional ByVal psAl As Date)

Dim lsSql As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim oFun As New COMFunciones.FCOMImpresion

Dim P As String, sAges As String, sTipo As String, sMone As String
Dim lcAge As String, lcMone As String, lnTot As Double

Dim i As Integer, j As Integer

    If OptAge = True Then
        sAges = 1
    Else
        sAges = 2
    End If
      
'    If psTipoC = 0 And psTipoM = 1 Then
'        sTipo = "2"
'    ElseIf psTipoC = 1 And psTipoM = 0 Then
'        sTipo = "1"
'    Else
'        sTipo = "1,2"
'    End If

'*** BRGO BASILEA II
    If psTipoCor = 0 And psTipoGraEmp = 0 And psTipoMedEmp = 0 And psTipoPeqEmp = 0 And psTipoMicEmp = 0 Then
        sTipo = "1,2,3,4,5"
    Else
        If psTipoCor = 1 Then
          sTipo = sTipo & ",1"
        End If
        If psTipoGraEmp = 1 Then
          sTipo = sTipo & ",2"
        End If
        If psTipoMedEmp = 1 Then
          sTipo = sTipo & ",3"
        End If
        If psTipoPeqEmp = 1 Then
          sTipo = sTipo & ",4"
        End If
        If psTipoMicEmp = 1 Then
          sTipo = sTipo & ",5"
        End If
        sTipo = Mid(sTipo, 2, Len(sTipo) - 1)
    End If

'**********************

    If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
        sMone = "1"
    ElseIf psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
        sMone = "2"
    Else
        sMone = "1,2"
    End If

    'lsSql = " exec stp_sel_ReporteCartasFianzasVigentesExcel '" & sAges & "','" & sMone & "','" & sTipo & "'," & nConsolidada
    lsSql = " exec stp_sel_ReporteCartasFianzasNuevasExcel '" & sMone & "','" & sTipo & "','" & sAges & "','" & Format(psDel, "yyyymmdd") & "','" & Format(psAl, "yyyymmdd") & "'"
     
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSql)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psMensaje = " No Existen Datos para el reporte en la Agencia "
        Exit Sub
    End If


'**************************************************************************************
    Dim ApExcel As Variant, lcTipGar As String, lnImporteSOL As Double, lnImporteDOL As Double
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    'ApExcel.Cells(2, 2).Formula = "psNomCmac"
    'ApExcel.Cells(3, 2).Formula = "psNomAge"
    ApExcel.Cells(2, 14).Formula = Date + Time()
    ApExcel.Cells(3, 14).Formula = psUsuarioRep
    ApExcel.Range("B2", "O8").Font.Bold = True
    
    ApExcel.Range("B2", "B3").HorizontalAlignment = xlLeft
    ApExcel.Range("N2", "O3").HorizontalAlignment = xlRight
    ApExcel.Range("B5", "O6").HorizontalAlignment = xlCenter
    ApExcel.Range("B9", "O10").VerticalAlignment = xlCenter
    ApExcel.Range("B9", "S10").Borders.LineStyle = 1
    
    ApExcel.Cells(5, 2).Formula = "CONTROL DE CARTAS FIANZAS NUEVAS"
'    ApExcel.Cells(6, 2).Formula = "AGENCIA - MONEDA"

    ApExcel.Range("B5", "O5").MergeCells = True
    ApExcel.Range("B6", "O6").MergeCells = True

    ApExcel.Range("B9", "B10").MergeCells = True
    ApExcel.Range("C9", "C10").MergeCells = True
    ApExcel.Range("D9", "D10").MergeCells = True
    ApExcel.Range("E9", "E10").MergeCells = True
    ApExcel.Range("F9", "F10").MergeCells = True
    ApExcel.Range("G9", "G10").MergeCells = True
    ApExcel.Range("H9", "H10").MergeCells = True
    ApExcel.Range("I9", "I10").MergeCells = True
    ApExcel.Range("J9", "J10").MergeCells = True
    ApExcel.Range("K9", "K10").MergeCells = True
    ApExcel.Range("L9", "L10").MergeCells = True
    ApExcel.Range("M9", "M9").MergeCells = True
    ApExcel.Range("N9", "O9").MergeCells = True
    ApExcel.Range("P9", "Q9").MergeCells = True
    ApExcel.Range("R9", "R9").MergeCells = True
    ApExcel.Range("S9", "S9").MergeCells = True
    
    ApExcel.Cells(9, 2).Formula = "ITEM"
    ApExcel.Cells(9, 3).Formula = "AGENCIA"
    ApExcel.Cells(9, 4).Formula = "CUENTA"
    ApExcel.Cells(9, 5).Formula = "CLIENTE"
    ApExcel.Cells(9, 6).Formula = "TIPO CREDITO"
    ApExcel.Cells(9, 7).Formula = "Nº C.FIANZA"
    ApExcel.Cells(9, 8).Formula = "MODALIDAD CARTA FIANZA"
    ApExcel.Cells(9, 9).Formula = "INICIO"
    ApExcel.Cells(9, 10).Formula = "VENCIMIENTO"
    ApExcel.Cells(9, 11).Formula = "BENEFICIARIO"
    ApExcel.Cells(9, 12).Formula = "IMPORTE DE LA CARTA FIANZA"
    ApExcel.Cells(9, 13).Formula = "GARANTIA"
    
    ApExcel.Cells(9, 14).Formula = "VALOR REALIZACION DE LA GARANTIA"
    ApExcel.Cells(10, 14).Formula = "S/."
    ApExcel.Cells(10, 15).Formula = "US$."
    
    ApExcel.Cells(9, 16).Formula = "VALOR GRAVADO DE LA GARANTIA"
    ApExcel.Cells(10, 16).Formula = "S/."
    ApExcel.Cells(10, 17).Formula = "US$."
    ApExcel.Cells(9, 18).Formula = "RENOVADO"
    ApExcel.Cells(9, 19).Formula = "ESTADO"
    
    ApExcel.Range("B9", "S10").Font.Bold = True
    ApExcel.Range("B9", "S10").HorizontalAlignment = 3
    
    ApExcel.Range("A1:Z1000").Font.Size = 8
    
    Dim nTotal As Integer
    Dim nImporte As Double
    Dim nSaldoSoles As Currency
    Dim nSaldoDolares As Currency
    nSaldoSoles = 0
    nSaldoDolares = 0

    i = 10
    j = 0
    Dim nCFAnterior As String
    Do While Not lrDataRep.EOF
        If nCFAnterior = lrDataRep!Carta_Fianza Then
            nSaldoSoles = nSaldoSoles + IIf(lrDataRep!nMonedaCG = "1", lrDataRep!nGravadoCG, 0)
            nSaldoDolares = nSaldoDolares + IIf(lrDataRep!nMonedaCG = "2", lrDataRep!nGravadoCG, 0)
            
            ApExcel.Cells(i, 16).Formula = nSaldoSoles
            ApExcel.Cells(i, 17).Formula = nSaldoDolares
        Else
            nSaldoSoles = 0
            nSaldoDolares = 0
        End If

        If nCFAnterior <> lrDataRep!Carta_Fianza Then
            
            i = i + 1
            j = j + 1
            nCFAnterior = lrDataRep!Carta_Fianza
            
            ApExcel.Cells(i, 2).Formula = "'" & Format(j, "000")
            ApExcel.Cells(i, 3).Formula = lrDataRep!NombAgencia
            ApExcel.Cells(i, 4).Formula = lrDataRep!Carta_Fianza
            ApExcel.Cells(i, 5).Formula = lrDataRep!Cliente
            ApExcel.Cells(i, 6).Formula = lrDataRep!TpoCred
            ApExcel.Cells(i, 7).Formula = "'" & Format(lrDataRep!num_poliza, "00000")
            ApExcel.Cells(i, 8).Formula = lrDataRep!Modalidad
            ApExcel.Cells(i, 9).Formula = "'" & Format(lrDataRep!Fecha_Emision, "dd/mm/yyyy")
            ApExcel.Cells(i, 10).Formula = "'" & Format(lrDataRep!Fecha_Vencimiento, "dd/mm/yyyy")
            ApExcel.Cells(i, 11).Formula = lrDataRep!Acreedor
            ApExcel.Cells(i, 12).Formula = lrDataRep!Importe
            
            If lrDataRep!nTpoGarantia = 4 Or lrDataRep!nTpoGarantia2 = 4 Then
                lcTipGar = "ART"
            Else
                lcTipGar = IIf(Len(lrDataRep!ctapzofjo) = 18, "DPF", IIf(lrDataRep!valgravhipo > 0, "HIP", IIf(lrDataRep!valgravhipo = 0 And lrDataRep!dCertifGravamen = "01/01/1990" And (lrDataRep!nTpoGarantia <> 0 Or lrDataRep!nTpoGarantia2 <> 0), "EnTram", "SGD")))
            End If
            lnImporteSOL = IIf(lcTipGar = "DPF" And Mid(lrDataRep!ctapzofjo, 9, 1) = "1", lrDataRep!saldopzofjo, IIf(lcTipGar = "HIP" And lrDataRep!nmonehipo = 1, lrDataRep!valgravhipo, 0))
            lnImporteDOL = IIf(lcTipGar = "DPF" And Mid(lrDataRep!ctapzofjo, 9, 1) = "2", lrDataRep!saldopzofjo, IIf(lcTipGar = "HIP" And lrDataRep!nmonehipo = 2, lrDataRep!valgravhipo, 0))
                
            ApExcel.Cells(i, 13).Formula = lcTipGar
            ApExcel.Cells(i, 14).Formula = lnImporteSOL
            ApExcel.Cells(i, 15).Formula = lnImporteDOL
            
            nSaldoSoles = IIf(lrDataRep!nMonedaCG = "1", lrDataRep!nGravadoCG, 0)
            nSaldoDolares = IIf(lrDataRep!nMonedaCG = "2", lrDataRep!nGravadoCG, 0)
            
            ApExcel.Cells(i, 16).Formula = nSaldoSoles
            ApExcel.Cells(i, 17).Formula = nSaldoDolares
            ApExcel.Cells(i, 18).Formula = lrDataRep!Num_Renovacion
            ApExcel.Cells(i, 19).Formula = lrDataRep!EstCred
            
            ApExcel.Range("K" & Trim(str(i)) & ":" & "Q" & Trim(str(i))).NumberFormat = "#,##0.00"
            ApExcel.Range("B" & Trim(str(i)) & ":" & "S" & Trim(str(i))).Borders.LineStyle = 1
            
            nTotal = nTotal + 1
            nImporte = nImporte + lrDataRep!Importe
        Else
        
            Dim liPosicion As Integer
            nCFAnterior = lrDataRep!Carta_Fianza
            
            If Len(lrDataRep!ctapzofjo) > 0 Then
                liPosicion = InStr(lcTipGar, "DPF")
                If liPosicion = 0 Then
                    lcTipGar = lcTipGar & "+" & "DPF"
                End If
            Else
                If lrDataRep!valgravhipo > 0 Then
                    liPosicion = InStr(lcTipGar, "HIP")
                    If liPosicion = 0 Then
                        lcTipGar = lcTipGar & "+" & "HIP"
                    End If
                Else
                
                    If lrDataRep!dCertifGravamen = "01/01/1990" Then
                        lcTipGar = lcTipGar & "+" & "EnTram"
                    Else
                        liPosicion = InStr(lcTipGar, "SGD")
                        If liPosicion = 0 Then
                            lcTipGar = lcTipGar & "+" & "SGD"
                        End If
                        
                    End If
                    
                End If
            End If
            
            'lcTipGar = IIf(Len(lrDataRep!ctapzofjo) > 0, lcTipGar & " " & "DPF", IIf(lrDataRep!valgravhipo > 0, lcTipGar, ""))
            
            lnImporteSOL = lnImporteSOL + IIf(Mid(lrDataRep!ctapzofjo, 9, 1) = "1", lrDataRep!saldopzofjo, IIf(lrDataRep!nmonehipo = 1, lrDataRep!valgravhipo, 0))
            lnImporteDOL = lnImporteDOL + IIf(Mid(lrDataRep!ctapzofjo, 9, 1) = "2", lrDataRep!saldopzofjo, IIf(lrDataRep!nmonehipo = 2, lrDataRep!valgravhipo, 0))
                    
            ApExcel.Cells(i, 13).Formula = lcTipGar
            ApExcel.Cells(i, 14).Formula = lnImporteSOL
            ApExcel.Cells(i, 15).Formula = lnImporteDOL
            
            ApExcel.Cells(i, 16).Formula = nSaldoSoles
            ApExcel.Cells(i, 17).Formula = nSaldoDolares
            ApExcel.Cells(i, 18).Formula = lrDataRep!Num_Renovacion 'MADM 20111226
            ApExcel.Cells(i, 19).Formula = lrDataRep!EstCred
            
            ApExcel.Range("K" & Trim(str(i)) & ":" & "Q" & Trim(str(i))).NumberFormat = "#,##0.00"
            ApExcel.Range("B" & Trim(str(i)) & ":" & "S" & Trim(str(i))).Borders.LineStyle = 1
    
        End If
'        If i = 51 Then
'        MsgBox "aaa"
'        End If
    lrDataRep.MoveNext
            
    If lrDataRep.EOF Then
        Exit Do
    End If
    Loop
    
    ApExcel.Cells(1 + i, 5).Formula = "Total de Cartas Fianzas"
    ApExcel.Cells(1 + i, 6).Formula = nTotal
    ApExcel.Cells(1 + i, 11).Formula = Format(nImporte, "#,##0.00")
    ApExcel.Cells(1 + i, 16).Formula = "=SUM(P11:P" & i & ")"
    ApExcel.Cells(1 + i, 17).Formula = "=SUM(Q11:Q" & i & ")"
      
    lrDataRep.Close
    Set lrDataRep = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("O:O").ColumnWidth = 17#
    ApExcel.Columns("N:N").ColumnWidth = 17#
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing
    
End Sub

Private Sub ImprimeCartaFianzaConsolidada(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdCMACT As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
        ByRef psMonedaE As Integer, ByRef psTipoCor As Integer, ByRef psTipoGraEmp As Integer, _
        ByRef psTipoMedEmp As Integer, ByRef psTipoPeqEmp As Integer, ByRef psTipoMicEmp As Integer, _
        ByRef OptAge As Boolean, Optional psListaAgencias As String, Optional ByRef psMensaje As String, Optional nConsolidada As Integer = 1, _
        Optional ByRef psFecIni As String, Optional ByRef psFecFin As String)

Dim lsSql As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim oFun As New COMFunciones.FCOMImpresion
Dim lcTipGar As String

Dim P As String, sAges As String, sTipo As String, sMone As String
Dim lcAge As String, lcMone As String, lnTot As Double
'**agregado***
Dim lnImporteSOL As Double
Dim lnImporteDOL As Double
'***fin

Dim i As Integer, j As Integer

    If OptAge = True Then
        sAges = psAgencia
    Else
        sAges = Replace(Replace(psListaAgencias, "'", ""), " ", "")
    End If
'*** BRGO BASILEA II
    If psTipoCor = 0 And psTipoGraEmp = 0 And psTipoMedEmp = 0 And psTipoPeqEmp = 0 And psTipoMicEmp = 0 Then
        sTipo = "1,2,3,4,5"
    Else
        If psTipoCor = 1 Then
          sTipo = sTipo & ",1"
        End If
        If psTipoGraEmp = 1 Then
          sTipo = sTipo & ",2"
        End If
        If psTipoMedEmp = 1 Then
          sTipo = sTipo & ",3"
        End If
        If psTipoPeqEmp = 1 Then
          sTipo = sTipo & ",4"
        End If
        If psTipoMicEmp = 1 Then
          sTipo = sTipo & ",5"
        End If
        sTipo = Mid(sTipo, 2, Len(sTipo) - 1)
    End If
'**********************

    If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
        sMone = "1"
    ElseIf psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
        sMone = "2"
    Else
        sMone = "1,2"
    End If

    lsSql = " exec [stp_sel_ReporteCartasFianzasConsolidada] '" & sAges & "','" & sMone & "','" & sTipo & "','" & Format(psFecIni, "yyyy/MM/dd") & "','" & Format(psFecFin, "yyyy/MM/dd") & "'"

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSql)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psMensaje = " No Existen Datos para el reporte en la Agencia "
        Exit Sub
    End If
    
'**************************************************************************************
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsFile As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim IniTablas As Integer
    Dim oPersona As UPersona
        
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    Set oPersona = New UPersona
    
    Dim nTotal As Integer
    Dim nImporte As Double
    Dim nSaldoSoles As Currency
    Dim nSaldoDolares As Currency
    Dim nNumRenovacion As Integer
    nSaldoSoles = 0
    nSaldoDolares = 0
    i = 0:  j = 0
    Dim nCFAnterior As String
    '*****RECO**************
    lsNomHoja = "Hoja1"
    lsFile = "ReporteCartaFianzaConsolidado"
    
    lsArchivo = "\spooler\" & "Carta_Fianza_Consolidado" & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    
    '*****END RECO *********
    IniTablas = 4
    Do While Not lrDataRep.EOF
        If nCFAnterior = lrDataRep!Carta_Fianza Then
            nSaldoSoles = nSaldoSoles + IIf(lrDataRep!nMonedaCG = "1", lrDataRep!nGravadoCG, 0)
            nSaldoDolares = nSaldoDolares + IIf(lrDataRep!nMonedaCG = "2", lrDataRep!nGravadoCG, 0)
            
            xlHoja1.Cells(IniTablas + i - 1, 18) = nSaldoSoles
            xlHoja1.Cells(IniTablas + i - 1, 19) = nSaldoDolares
        Else
            nSaldoSoles = 0
            nSaldoDolares = 0
        End If

        If nCFAnterior <> lrDataRep!Carta_Fianza Then
            
            i = i + 1
            j = j + 1
            nCFAnterior = lrDataRep!Carta_Fianza
            
            xlHoja1.Cells(IniTablas + i - 1, 2) = "'" & Format(j, "000")
            xlHoja1.Cells(IniTablas + i - 1, 3) = lrDataRep!NombAgencia
            xlHoja1.Cells(IniTablas + i - 1, 4) = lrDataRep!Carta_Fianza
            xlHoja1.Cells(IniTablas + i - 1, 5) = lrDataRep!Cliente
            xlHoja1.Cells(IniTablas + i - 1, 6) = lrDataRep!TpoCred
            xlHoja1.Cells(IniTablas + i - 1, 7) = "'" & Format(lrDataRep!num_poliza, "00000")
            xlHoja1.Cells(IniTablas + i - 1, 8) = lrDataRep!Modalidad
            xlHoja1.Cells(IniTablas + i - 1, 9) = "'" & Format(lrDataRep!Fecha_Emision, "dd/mm/yyyy")
            xlHoja1.Cells(IniTablas + i - 1, 10) = "'" & Format(lrDataRep!Fecha_Vencimiento, "dd/mm/yyyy")
            xlHoja1.Cells(IniTablas + i - 1, 11) = lrDataRep!Acreedor
            xlHoja1.Cells(IniTablas + i - 1, 12) = UCase(lrDataRep!Analista)
            xlHoja1.Cells(IniTablas + i - 1, 13) = lrDataRep!Estado
            xlHoja1.Cells(IniTablas + i - 1, 14) = lrDataRep!Importe
            
            If lrDataRep!nTpoGarantia = 4 Or lrDataRep!nTpoGarantia2 = 4 Then
                lcTipGar = "ART"
            Else
                lcTipGar = IIf(Len(lrDataRep!ctapzofjo) = 18, "DPF", IIf(lrDataRep!valgravhipo > 0, "HIP", IIf(lrDataRep!valgravhipo = 0 And lrDataRep!dCertifGravamen = "01/01/1990" And (lrDataRep!nTpoGarantia <> 0 Or lrDataRep!nTpoGarantia2 <> 0), "EnTram", "SGD")))
            End If
            lnImporteSOL = IIf(lcTipGar = "DPF" And Mid(lrDataRep!ctapzofjo, 9, 1) = "1", lrDataRep!saldopzofjo, IIf(lcTipGar = "HIP" And lrDataRep!nmonehipo = 1, lrDataRep!valgravhipo, 0))
            lnImporteDOL = IIf(lcTipGar = "DPF" And Mid(lrDataRep!ctapzofjo, 9, 1) = "2", lrDataRep!saldopzofjo, IIf(lcTipGar = "HIP" And lrDataRep!nmonehipo = 2, lrDataRep!valgravhipo, 0))
                
            xlHoja1.Cells(IniTablas + i - 1, 15) = lcTipGar
            xlHoja1.Cells(IniTablas + i - 1, 16) = lnImporteSOL
            xlHoja1.Cells(IniTablas + i - 1, 17) = lnImporteDOL
            
            nSaldoSoles = IIf(lrDataRep!nMonedaCG = "1", lrDataRep!nGravadoCG, 0)
            nSaldoDolares = IIf(lrDataRep!nMonedaCG = "2", lrDataRep!nGravadoCG, 0)
            
            xlHoja1.Cells(IniTablas + i - 1, 18) = nSaldoSoles
            xlHoja1.Cells(IniTablas + i - 1, 19) = nSaldoDolares
            
            
            nNumRenovacion = lrDataRep!Num_Renovacion
            If nNumRenovacion > 0 Then
                xlHoja1.Cells(IniTablas + i - 1, 20) = "SI"
                xlHoja1.Cells(IniTablas + i - 1, 21) = lrDataRep!Num_Renovacion
            Else
                xlHoja1.Cells(IniTablas + i - 1, 20) = "NO"
                xlHoja1.Cells(IniTablas + i - 1, 21) = "--"
            End If
            nTotal = nTotal + 1
            nImporte = nImporte + lrDataRep!Importe
        Else
            Dim liPosicion As Integer
            nCFAnterior = lrDataRep!Carta_Fianza
            If Len(lrDataRep!ctapzofjo) > 0 Then
                liPosicion = InStr(lcTipGar, "DPF")
                If liPosicion = 0 Then
                    lcTipGar = lcTipGar & "+" & "DPF"
                End If
            Else
                If lrDataRep!valgravhipo > 0 Then
                    liPosicion = InStr(lcTipGar, "HIP")
                    If liPosicion = 0 Then
                        lcTipGar = lcTipGar & "+" & "HIP"
                    End If
                Else
                
                    If lrDataRep!dCertifGravamen = "01/01/1990" Then
                        lcTipGar = lcTipGar & "+" & "EnTram"
                    Else
                        liPosicion = InStr(lcTipGar, "SGD")
                        If liPosicion = 0 Then
                            lcTipGar = lcTipGar & "+" & "SGD"
                        End If
                        
                    End If
                    
                End If
            End If
            
            lnImporteSOL = lnImporteSOL + IIf(Mid(lrDataRep!ctapzofjo, 9, 1) = "1", lrDataRep!saldopzofjo, IIf(lrDataRep!nmonehipo = 1, lrDataRep!valgravhipo, 0))
            lnImporteDOL = lnImporteDOL + IIf(Mid(lrDataRep!ctapzofjo, 9, 1) = "2", lrDataRep!saldopzofjo, IIf(lrDataRep!nmonehipo = 2, lrDataRep!valgravhipo, 0))
    
        End If
    lrDataRep.MoveNext
            
    If lrDataRep.EOF Then
        Exit Do
    End If
    Loop
    lrDataRep.Close
    Set lrDataRep = Nothing
    Screen.MousePointer = 0
    
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(i + 3, 21)).Borders.LineStyle = 1
    Dim psArchivoAGrabarC As String
    
    xlHoja1.SaveAs App.path & lsArchivo
    psArchivoAGrabarC = App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    
End Sub
'*** FRHU 20131119
Private Sub ImprimeCartaFianzaRegistradas(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdCMACT As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
        ByRef psMonedaE As Integer, ByRef psTipoCor As Integer, ByRef psTipoGraEmp As Integer, _
        ByRef psTipoMedEmp As Integer, ByRef psTipoPeqEmp As Integer, ByRef psTipoMicEmp As Integer, _
        ByRef OptAge As Boolean, Optional psListaAgencias As String, Optional ByRef psMensaje As String, _
        Optional nConsolidada As Integer = 1, Optional ByRef dFecIni As String, Optional ByRef dFecFin As String)

Dim lsSql As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim oFun As New COMFunciones.FCOMImpresion

Dim P As String, sAges As String, sTipo As String, sMone As String
Dim lcAge As String, lcMone As String, lnTot As Double

Dim i As Integer, j As Integer

    If OptAge = True Then
        sAges = psAgencia
    Else
        sAges = Replace(Replace(psListaAgencias, "'", ""), " ", "")
    End If
      
'    If psTipoC = 0 And psTipoM = 1 Then
'        sTipo = "2"
'    ElseIf psTipoC = 1 And psTipoM = 0 Then
'        sTipo = "1"
'    Else
'        sTipo = "1,2"
'    End If

'*** BRGO BASILEA II
    If psTipoCor = 0 And psTipoGraEmp = 0 And psTipoMedEmp = 0 And psTipoPeqEmp = 0 And psTipoMicEmp = 0 Then
        sTipo = "1,2,3,4,5"
    Else
        If psTipoCor = 1 Then
          sTipo = sTipo & ",1"
        End If
        If psTipoGraEmp = 1 Then
          sTipo = sTipo & ",2"
        End If
        If psTipoMedEmp = 1 Then
          sTipo = sTipo & ",3"
        End If
        If psTipoPeqEmp = 1 Then
          sTipo = sTipo & ",4"
        End If
        If psTipoMicEmp = 1 Then
          sTipo = sTipo & ",5"
        End If
        sTipo = Mid(sTipo, 2, Len(sTipo) - 1)
    End If

'**********************

    If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
        sMone = "1"
    ElseIf psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
        sMone = "2"
    Else
        sMone = "1,2"
    End If
    
    lsSql = " exec stp_sel_ReporteCartasFianzasRegistradasExcel '" & sAges & "','" & sMone & "','" & sTipo & "','" & dFecIni & "','" & dFecFin & "'"
   
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSql)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psMensaje = " No Existen Datos para el reporte en la Agencia "
        Exit Sub
    End If


'**************************************************************************************
    Dim ApExcel As Variant, lcTipGar As String, lnImporteSOL As Double, lnImporteDOL As Double
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    'ApExcel.Cells(2, 2).Formula = "psNomCmac"
    'ApExcel.Cells(3, 2).Formula = "psNomAge"
    ApExcel.Cells(2, 14).Formula = Date + Time()
    ApExcel.Cells(3, 14).Formula = psUsuarioRep
    ApExcel.Range("B2", "O8").Font.Bold = True
    
    ApExcel.Range("B2", "B3").HorizontalAlignment = xlLeft
    ApExcel.Range("N2", "O3").HorizontalAlignment = xlRight
    ApExcel.Range("B5", "O6").HorizontalAlignment = xlCenter
    ApExcel.Range("B9", "N10").VerticalAlignment = xlCenter
    ApExcel.Range("B9", "N10").Borders.LineStyle = 1
    
    ApExcel.Cells(5, 2).Formula = "CONTROL DE CARTAS FIANZAS REGISTRADAS"
'    ApExcel.Cells(6, 2).Formula = "AGENCIA - MONEDA"

    ApExcel.Range("B5", "O5").MergeCells = True
    ApExcel.Range("B6", "O6").MergeCells = True

    ApExcel.Range("B9", "B10").MergeCells = True
    ApExcel.Range("C9", "C10").MergeCells = True
    ApExcel.Range("D9", "D10").MergeCells = True
    ApExcel.Range("E9", "E10").MergeCells = True
    ApExcel.Range("F9", "F10").MergeCells = True
    ApExcel.Range("G9", "G10").MergeCells = True
    ApExcel.Range("H9", "H10").MergeCells = True
    ApExcel.Range("I9", "I10").MergeCells = True
    ApExcel.Range("J9", "J10").MergeCells = True
    ApExcel.Range("K9", "K10").MergeCells = True
    ApExcel.Range("L9", "L10").MergeCells = True
    ApExcel.Range("M9", "M10").MergeCells = True
    ApExcel.Range("N9", "N10").MergeCells = True
    'ApExcel.Range("P9", "Q9").MergeCells = True
    
    ApExcel.Cells(9, 2).Formula = "ITEM"
    ApExcel.Cells(9, 3).Formula = "CARTA FIANZA"
    ApExcel.Cells(9, 4).Formula = "CLIENTE"
    ApExcel.Cells(9, 5).Formula = "EMISION"
    ApExcel.Cells(9, 6).Formula = "VENCIMIENTO"
    ApExcel.Cells(9, 7).Formula = "ACREEDOR"
    ApExcel.Cells(9, 8).Formula = "ANALISTA"
    ApExcel.Cells(9, 9).Formula = "ESTADO"
    ApExcel.Cells(9, 10).Formula = "TIP_CREDITO"
    ApExcel.Cells(9, 11).Formula = "AGENCIA"
    ApExcel.Cells(9, 12).Formula = "MONEDA"
    ApExcel.Cells(9, 13).Formula = "MONTO"
    
    ApExcel.Cells(9, 14).Formula = "Modalidad"
     
    ApExcel.Range("B9", "N10").Font.Bold = True
    ApExcel.Range("B9", "N10").HorizontalAlignment = 3
    
    ApExcel.Range("A1:Z1000").Font.Size = 8
    
    Dim nTotal As Integer
    Dim nImporte As Double
    Dim nSaldoSoles As Currency
    Dim nSaldoDolares As Currency
    nSaldoSoles = 0
    nSaldoDolares = 0

    i = 10
    j = 0
    Dim nCFAnterior As String
    Do While Not lrDataRep.EOF
        
        'If nCFAnterior <> lrDataRep!carta_fianza Then
            
            i = i + 1
            j = j + 1
            nCFAnterior = lrDataRep!Carta_Fianza
            
            ApExcel.Cells(i, 2).Formula = "'" & Format(j, "000")
            ApExcel.Cells(i, 3).Formula = lrDataRep!Carta_Fianza
            ApExcel.Cells(i, 4).Formula = lrDataRep!Cliente
            ApExcel.Cells(i, 5).Formula = "'" & Format(lrDataRep!Fecha_Emision, "dd/mm/yyyy")
            ApExcel.Cells(i, 6).Formula = "'" & Format(lrDataRep!Fecha_Vencimiento, "dd/mm/yyyy")
            ApExcel.Cells(i, 7).Formula = lrDataRep!Acreedor
            ApExcel.Cells(i, 8).Formula = lrDataRep!Analista
            ApExcel.Cells(i, 9).Formula = lrDataRep!Estado
            ApExcel.Cells(i, 10).Formula = lrDataRep!Tip_Credito
            ApExcel.Cells(i, 11).Formula = lrDataRep!NombAgencia
            ApExcel.Cells(i, 12).Formula = lrDataRep!cmoneda
                       
            ApExcel.Cells(i, 13).Formula = Format(lrDataRep!Importe, "#,##0.00")
            ApExcel.Cells(i, 14).Formula = lrDataRep!Modalidad
                                         
            nTotal = nTotal + 1
            nImporte = nImporte + lrDataRep!Importe
            
            ApExcel.Range("B" & Trim(str(i)) & ":" & "N" & Trim(str(i))).Borders.LineStyle = 1
        'Else
        
            nCFAnterior = lrDataRep!Carta_Fianza
           
        'End If
'        If i = 51 Then
'        MsgBox "aaa"
'        End If
    lrDataRep.MoveNext
            
    If lrDataRep.EOF Then
        Exit Do
    End If
    Loop
    
    ApExcel.Cells(1 + i, 5).Formula = "Total de Cartas Fianzas"
    ApExcel.Cells(1 + i, 6).Formula = nTotal
    'ApExcel.Cells(1 + i, 11).Formula = Format(nImporte, "#,##0.00")
    'ApExcel.Cells(1 + i, 16).Formula = "=SUM(P11:P" & i & ")"
    'ApExcel.Cells(1 + i, 17).Formula = "=SUM(Q11:Q" & i & ")"
      
    lrDataRep.Close
    Set lrDataRep = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("O:O").ColumnWidth = 30#
    ApExcel.Columns("N:N").ColumnWidth = 30#
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing
    
End Sub
Private Sub ImprimeCartaFianzaVigentesPorVencer(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdCMACT As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
        ByRef psMonedaE As Integer, ByRef psTipoCor As Integer, ByRef psTipoGraEmp As Integer, _
        ByRef psTipoMedEmp As Integer, ByRef psTipoPeqEmp As Integer, ByRef psTipoMicEmp As Integer, _
        ByRef OptAge As Boolean, Optional psListaAgencias As String, Optional ByRef psMensaje As String, _
        Optional nConsolidada As Integer = 1, Optional ByRef dFecIni As String, Optional ByRef dFecFin As String)

Dim lsSql As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim oFun As New COMFunciones.FCOMImpresion

Dim P As String, sAges As String, sTipo As String, sMone As String
Dim lcAge As String, lcMone As String, lnTot As Double

Dim i As Integer, j As Integer

    If OptAge = True Then
        sAges = psAgencia
    Else
        sAges = Replace(Replace(psListaAgencias, "'", ""), " ", "")
    End If
      
'    If psTipoC = 0 And psTipoM = 1 Then
'        sTipo = "2"
'    ElseIf psTipoC = 1 And psTipoM = 0 Then
'        sTipo = "1"
'    Else
'        sTipo = "1,2"
'    End If

'*** BRGO BASILEA II
    If psTipoCor = 0 And psTipoGraEmp = 0 And psTipoMedEmp = 0 And psTipoPeqEmp = 0 And psTipoMicEmp = 0 Then
        sTipo = "1,2,3,4,5"
    Else
        If psTipoCor = 1 Then
          sTipo = sTipo & ",1"
        End If
        If psTipoGraEmp = 1 Then
          sTipo = sTipo & ",2"
        End If
        If psTipoMedEmp = 1 Then
          sTipo = sTipo & ",3"
        End If
        If psTipoPeqEmp = 1 Then
          sTipo = sTipo & ",4"
        End If
        If psTipoMicEmp = 1 Then
          sTipo = sTipo & ",5"
        End If
        sTipo = Mid(sTipo, 2, Len(sTipo) - 1)
    End If

'**********************

    If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
        sMone = "1"
    ElseIf psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
        sMone = "2"
    Else
        sMone = "1,2"
    End If
    
    lsSql = " exec stp_sel_ReporteCartasFianzasPorVencerExcel '" & sAges & "','" & sMone & "','" & sTipo & "','" & dFecIni & "','" & dFecFin & "'"
   
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSql)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psMensaje = " No Existen Datos para el reporte en la Agencia "
        Exit Sub
    End If


'**************************************************************************************
    Dim ApExcel As Variant, lcTipGar As String, lnImporteSOL As Double, lnImporteDOL As Double
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    'ApExcel.Cells(2, 2).Formula = "psNomCmac"
    'ApExcel.Cells(3, 2).Formula = "psNomAge"
    ApExcel.Cells(2, 12).Formula = Date + Time()
    ApExcel.Cells(3, 12).Formula = psUsuarioRep
    ApExcel.Range("B2", "O8").Font.Bold = True
    
    ApExcel.Range("B2", "B3").HorizontalAlignment = xlLeft
    ApExcel.Range("N2", "O3").HorizontalAlignment = xlRight
    ApExcel.Range("B5", "O6").HorizontalAlignment = xlCenter
    ApExcel.Range("B9", "L10").VerticalAlignment = xlCenter
    ApExcel.Range("B9", "L10").Borders.LineStyle = 1
    
    ApExcel.Cells(5, 2).Formula = "CONTROL DE CARTAS FIANZAS VIGENTES POR VENCER"
'    ApExcel.Cells(6, 2).Formula = "AGENCIA - MONEDA"

    ApExcel.Range("B5", "O5").MergeCells = True
    ApExcel.Range("B6", "O6").MergeCells = True

    ApExcel.Range("B9", "B10").MergeCells = True
    ApExcel.Range("C9", "C10").MergeCells = True
    ApExcel.Range("D9", "D10").MergeCells = True
    ApExcel.Range("E9", "E10").MergeCells = True
    ApExcel.Range("F9", "F10").MergeCells = True
    ApExcel.Range("G9", "G10").MergeCells = True
    ApExcel.Range("H9", "H10").MergeCells = True
    ApExcel.Range("I9", "I10").MergeCells = True
    ApExcel.Range("J9", "J10").MergeCells = True
    ApExcel.Range("K9", "K10").MergeCells = True
    ApExcel.Range("L9", "L10").MergeCells = True
    ApExcel.Range("M9", "M10").MergeCells = True
    ApExcel.Range("N9", "N10").MergeCells = True
    'ApExcel.Range("P9", "Q9").MergeCells = True
    
    ApExcel.Cells(9, 2).Formula = "ITEM"
    ApExcel.Cells(9, 3).Formula = "CARTA FIANZA"
    ApExcel.Cells(9, 4).Formula = "CLIENTE"
    ApExcel.Cells(9, 5).Formula = "DIRECCION"
    ApExcel.Cells(9, 6).Formula = "EMISION"
    ApExcel.Cells(9, 7).Formula = "VENCIMIENTO"
    ApExcel.Cells(9, 8).Formula = "IMPORTE"
    ApExcel.Cells(9, 9).Formula = "ACREEDOR"
    ApExcel.Cells(9, 10).Formula = "ANALISTA"
    ApExcel.Cells(9, 11).Formula = "TIPO CREDITO"
    ApExcel.Cells(9, 12).Formula = "TIPO DE GARANTIA"
    'ApExcel.Cells(9, 13).Formula = "MONTO"
    
    'ApExcel.Cells(9, 14).Formula = "Modalidad"
     
    ApExcel.Range("B9", "N10").Font.Bold = True
    ApExcel.Range("B9", "N10").HorizontalAlignment = 3
    
    ApExcel.Range("A1:Z1000").Font.Size = 8
    
    Dim nTotal As Integer
    Dim nImporte As Double
    Dim nSaldoSoles As Currency
    Dim nSaldoDolares As Currency
    nSaldoSoles = 0
    nSaldoDolares = 0

    i = 10
    j = 0
    Dim nCFAnterior As String
    Do While Not lrDataRep.EOF
        
        'If nCFAnterior <> lrDataRep!carta_fianza Then
            
            i = i + 1
            j = j + 1
            nCFAnterior = lrDataRep!Carta_Fianza
            
            ApExcel.Cells(i, 2).Formula = "'" & Format(j, "000")
            ApExcel.Cells(i, 3).Formula = lrDataRep!Carta_Fianza
            ApExcel.Cells(i, 4).Formula = lrDataRep!Cliente
            ApExcel.Cells(i, 5).Formula = lrDataRep!cPersDireccDomicilio
            ApExcel.Cells(i, 6).Formula = "'" & Format(lrDataRep!Fecha_Emision, "dd/mm/yyyy")
            ApExcel.Cells(i, 7).Formula = "'" & Format(lrDataRep!Fecha_Vencimiento, "dd/mm/yyyy")
            ApExcel.Cells(i, 8).Formula = lrDataRep!Importe
            ApExcel.Cells(i, 9).Formula = lrDataRep!Acreedor
            ApExcel.Cells(i, 10).Formula = lrDataRep!Analista
            ApExcel.Cells(i, 11).Formula = lrDataRep!Tip_Credito
            
            If Len(lrDataRep!tipouno) > 0 And Len(lrDataRep!tipodos) > 0 Then
            
                ApExcel.Cells(i, 12).Formula = lrDataRep!tipouno + " / " + lrDataRep!tipodos
            
            Else
            
                If Len(lrDataRep!tipouno) > 0 Then
                
                ApExcel.Cells(i, 12).Formula = lrDataRep!tipouno
                
                End If
                
                If Len(lrDataRep!tipodos) > 0 Then
                
                ApExcel.Cells(i, 12).Formula = lrDataRep!tipodos
                            
                End If
                        
            End If
                                              
            'ApExcel.Cells(i, 13).Formula = Format(lrDataRep!Importe, "#,##0.00")
            'ApExcel.Cells(i, 14).Formula = lrDataRep!Modalidad
                                         
            nTotal = nTotal + 1
            nImporte = nImporte + lrDataRep!Importe
            
            ApExcel.Range("B" & Trim(str(i)) & ":" & "L" & Trim(str(i))).Borders.LineStyle = 1
        'Else
        
            nCFAnterior = lrDataRep!Carta_Fianza
           
        'End If
'        If i = 51 Then
'        MsgBox "aaa"
'        End If
    lrDataRep.MoveNext
            
    If lrDataRep.EOF Then
        Exit Do
    End If
    Loop
    
    ApExcel.Cells(1 + i, 5).Formula = "Total de Cartas Fianzas"
    ApExcel.Cells(1 + i, 6).Formula = nTotal
    'ApExcel.Cells(1 + i, 11).Formula = Format(nImporte, "#,##0.00")
    'ApExcel.Cells(1 + i, 16).Formula = "=SUM(P11:P" & i & ")"
    'ApExcel.Cells(1 + i, 17).Formula = "=SUM(Q11:Q" & i & ")"
      
    lrDataRep.Close
    Set lrDataRep = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("O:O").ColumnWidth = 30#
    ApExcel.Columns("N:N").ColumnWidth = 30#
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing
    
End Sub
'*** FIN FRHU 20131119
'RECO20160120 ERS001-2016***********************************************
Private Function ObtieneMoneda() As String
    If ChkMoneda(0).value = 1 Then ObtieneMoneda = "1"
    If ChkMoneda(1).value = 1 Then ObtieneMoneda = "2"
    If ChkMoneda(0).value = 1 And ChkMoneda(1).value = 1 Then ObtieneMoneda = "1,2"
End Function
Private Function ObtieneTipoCred() As String
    If ChkTipo(0).value = 1 Then ObtieneTipoCred = ObtieneTipoCred & "1,"
    If ChkTipo(1).value = 1 Then ObtieneTipoCred = ObtieneTipoCred & "2,"
    If ChkTipo(2).value = 1 Then ObtieneTipoCred = ObtieneTipoCred & "3,"
    If ChkTipo(3).value = 1 Then ObtieneTipoCred = ObtieneTipoCred & "4,"
    If ChkTipo(4).value = 1 Then ObtieneTipoCred = ObtieneTipoCred & "5,"
    ObtieneTipoCred = Mid(ObtieneTipoCred, 1, Len(ObtieneTipoCred) - 1)
End Function

Private Sub GeneraCartaAvisoVencimiento(ByVal prsDatosCartas As ADODB.Recordset)
    On Error GoTo ErrorImprimirPDF
    Dim oCF As New COMNCartaFianza.NCOMCartaFianzaReporte
    Dim oDoc  As New cPDF
    Dim nTipo As Integer, nPosicion As Integer, nIndice As Integer
    Dim sNroCarta As String
    Dim nTop As Integer
    Dim nIndexChar As Integer
    
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Cartas de Aviso por Vencimiento " & gsUser
    oDoc.Title = "Cartas de Aviso por Vencimiento " & gsUser
    
    If Not oDoc.PDFCreate(App.path & "\Spooler\" & IIf(nTipo = 1, "Previo", "") & "CF_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    If Not (prsDatosCartas.EOF And prsDatosCartas.BOF) Then
            oDoc.Fonts.Add "F1", "Times New Roman", TrueType, Normal, WinAnsiEncoding
            oDoc.Fonts.Add "F2", "Times New Roman", TrueType, Bold, WinAnsiEncoding
            oDoc.Fonts.Add "F3", "Arial", TrueType, Normal, WinAnsiEncoding
            oDoc.Fonts.Add "F4", "Arial", TrueType, Bold, WinAnsiEncoding
            oDoc.Fonts.Add "F", "Times New Roman", TrueType, Bold, WinAnsiEncoding
        For nIndice = 1 To prsDatosCartas.RecordCount
    
            sNroCarta = oCF.RegistraCartaAvisoVencimiento(prsDatosCartas!Carta_Fianza, gsCodUser)
    
            Dim nFTabla As Integer
            Dim nFTablaCabecera As Integer
            Dim lnFontSizeBody As Integer
        
            nFTablaCabecera = 7
            nFTabla = 7
            lnFontSizeBody = 7
            nTop = 40
            oDoc.NewPage A4_Vertical
            
            nIndexChar = InStr(prsDatosCartas!cUbiGeoDescripcion, "(")
            oDoc.WTextBox 40 + nTop, 70, 10, 450, "", "F3", 11, hLeft, , vbBlack
            oDoc.WTextBox 50 + nTop, 70, 10, 450, Mid(prsDatosCartas!cUbiGeoDescripcion, 1, 1) & LCase(Mid(prsDatosCartas!cUbiGeoDescripcion, 2, nIndexChar - 2)) & ", " & IIf(Len(Day(Date)) = 1, "0" & Day(Date), Day(Date)) & " de " & fgDameNombreMes(Month(Date)) & " del " & Year(Date), "F3", 11, hLeft, , vbBlack
            oDoc.WTextBox 80 + nTop, 70, 10, 450, "CARTA Nº " & sNroCarta & "-SO-GA-CMACM ", "F4", 11, hLeft, , vbBlack
            'oDoc.WTextBox 95, 70, 0, 185, "", "F4", 11, hLeft, , vbBlack, 1
            
            oDoc.WTextBox 110 + nTop, 70, 10, 450, "Señora", "F3", 11, hLeft, , vbBlack
            oDoc.WTextBox 125 + nTop, 70, 10, 450, prsDatosCartas!Cliente, "F4", 11, hLeft, , vbBlack
            oDoc.WTextBox 140 + nTop, 70, 10, 450, prsDatosCartas!cPersDireccDomicilio, "F3", 11, hLeft, , vbBlack
            oDoc.WTextBox 155 + nTop, 70, 10, 450, "Presente", "F3", 11, hLeft, , vbBlack
            'oDoc.WTextBox 170, 70, 0, 45, "", "F3", 11, hLeft, , vbBlack, 1
            
            oDoc.WTextBox 195 + nTop, 70, 10, 70, "Referencia: ", "F3", 11, hLeft, , vbBlack
            oDoc.WTextBox 195 + nTop, 130, 10, 450, "Carta Fianza Nº " & prsDatosCartas!nPoliza & " – " & prsDatosCartas!Acreedor, "F4", 11, hLeft, , vbBlack
            
            oDoc.WTextBox 225 + nTop, 70, 10, 450, "De nuestra consideración:", "F3", 11, hLeft, , vbBlack
            oDoc.WTextBox 245 + nTop, 70, 10, 450, "Es grato dirigirme a usted para saludarlo(a) y al mismo tiempo comunicarle que el " & IIf(Len(Day(prsDatosCartas!Fecha_Vencimiento)) = 1, "0" _
                                           & Day(prsDatosCartas!Fecha_Vencimiento), Day(prsDatosCartas!Fecha_Vencimiento)) & " de " & fgDameNombreMes(Month(prsDatosCartas!Fecha_Vencimiento)) _
                                           & " " & Year(prsDatosCartas!Fecha_Vencimiento) & " venció la Carta Fianza " & prsDatosCartas!Carta_Fianza & " , que nuestra institución emitió " _
                                           & "garantizando a su representada " & Format(prsDatosCartas!Fecha_Emision, "dd.MM.yyyy") & ", con un plazo de " & prsDatosCartas!nPeriodo & " días, a favor de " _
                                           & prsDatosCartas!Acreedor & " , por un monto de S/. " & prsDatosCartas!Importe, "F3", 11, hjustify, , vbBlack
            
            oDoc.WTextBox 310 + nTop, 70, 10, 450, "En tal sentido, si en un lapso de 15 días como máximo después de la fecha de su vencimiento no renueva la carta fianza de la referencia, se le solicita " _
                                            & "la devolución del documento original otorgado para su cancelación, caso contrario si la institución beneficiaria solicita la ejecución se estará " _
                                            & "procediendo a la ejecución respectiva.", "F3", 11, hjustify, , vbBlack
            
            oDoc.WTextBox 370 + nTop, 70, 10, 450, "La renovación de la Carta Fianza, se podrá realizar a partir de la fecha de vencimiento.", "F3", 11, hLeft, , vbBlack
            
            oDoc.WTextBox 395 + nTop, 70, 10, 450, "Finalmente, le indicamos que en caso de renovación deberá presentar la siguiente documentación:", "F3", 11, hjustify, , vbBlack
            oDoc.WTextBox 435 + nTop, 150, 10, 450, "Solicitud de la renovación de la Carta Fianza", "F3", 11, hLeft, , vbBlack
            oDoc.WTextBox 450 + nTop, 150, 10, 450, "Carta del acreedor indicando el estado del servicio", "F3", 11, hLeft, , vbBlack
            oDoc.WTextBox 465 + nTop, 150, 10, 450, "Pago de la comisión por la renovación.", "F3", 11, hLeft, , vbBlack
            oDoc.WTextBox 500 + nTop, 70, 10, 450, "Sin otro particular, nos suscribimos de usted.", "F3", 11, hLeft, , vbBlack
            oDoc.WTextBox 525 + nTop, 70, 10, 450, "Atentamente,", "F3", 11, hLeft, , vbBlack
            
            prsDatosCartas.MoveNext
        Next
    End If

    oDoc.PDFClose
    oDoc.Show
    Exit Sub
ErrorImprimirPDF:
    MsgBox err.Description, vbInformation, "Aviso"

End Sub
Private Function ValidaDatos() As String
    ValidaDatos = ""
    If ChkTipo(0).value = 0 And ChkTipo(1).value = 0 And ChkTipo(2).value = 0 And ChkTipo(3).value = 0 Then
        ValidaDatos = "Debe seleccionar por lo menos un tipo de crédito"
        Exit Function
    End If
    If ChkMoneda(0).value = 0 And ChkMoneda(1).value = 0 Then
        ValidaDatos = "Debe seleccionar por lo menos un tipo de moneda"
        Exit Function
    End If
    If mskPeriodo1Del.Text = "__/__/____" Or mskPeriodo1Al.Text = "__/__/____" Then
        ValidaDatos = "Debe ingresar la fecha"
        Exit Function
    End If
    
    ValidaDatos = gFunGeneral.ValidaFecha(mskPeriodo1Del.Text)
    If ValidaDatos <> "" Then
        Exit Function
    End If
    ValidaDatos = gFunGeneral.ValidaFecha(mskPeriodo1Al.Text)
     If ValidaDatos <> "" Then
        Exit Function
    End If
    
End Function
'RECO FIN***************************************************************

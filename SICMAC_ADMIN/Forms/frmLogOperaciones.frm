VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogOperaciones 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SICMACT"
   ClientHeight    =   6285
   ClientLeft      =   1920
   ClientTop       =   1830
   ClientWidth     =   7755
   ForeColor       =   &H00EFEFEF&
   HelpContextID   =   210
   Icon            =   "frmLogOperaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   7755
   Begin MSComctlLib.TreeView tvOpe 
      Height          =   4335
      Left            =   375
      TabIndex        =   0
      Top             =   1395
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   7646
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "imglstFiguras"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Seleccionar"
      Height          =   370
      Left            =   4650
      TabIndex        =   2
      Top             =   5820
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   370
      Left            =   6270
      TabIndex        =   1
      Top             =   5820
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imglstFiguras 
      Left            =   1425
      Top             =   7230
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selector de Operaciones"
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
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   1020
      Width           =   2115
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006B6BB8&
      FillColor       =   &H80000005&
      Height          =   4425
      Left            =   300
      Top             =   1320
      Width           =   7170
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   240
      Picture         =   "frmLogOperaciones.frx":08CA
      Top             =   120
      Width           =   1515
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   -300
      Picture         =   "frmLogOperaciones.frx":3C1C
      Top             =   3300
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   -300
      Picture         =   "frmLogOperaciones.frx":3F5E
      Top             =   3600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D7EDFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006B6BB8&
      FillColor       =   &H00C0E0FF&
      Height          =   435
      Left            =   300
      Top             =   900
      Width           =   7170
   End
End
Attribute VB_Name = "frmLogOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRaizOpe As String

Public Sub Inicio(pcRaizOpe As String)
cRaizOpe = pcRaizOpe
Me.Show 1
End Sub

Private Sub cmdAceptar_Click()
tvOpe_DblClick
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Logo.AutoPlay = True
'Logo.Open App.Path & "\videos\Loguito.avi"
'Logo.Open "E:\Archivos de programa\AVILOGO\logoX.avi"
imglstFiguras.ListImages.Add 1, "Padre", Image1
imglstFiguras.ListImages.Add 2, "Hijo", Image2
imglstFiguras.ListImages.Add 3, "Hijito", Image2
imglstFiguras.ListImages.Add 4, "Bebe", Image2
CentraForm Me
CargaOperaciones
End Sub

Sub CargaOperaciones()
Dim oConn As New DConecta, Rs As New ADODB.Recordset
Dim sSQL As String, cKey As String, cKeySup As String
Dim cDescripcion As String, cClave As String
Dim cCod As String, cUlt As String, cIMG As String
Dim cGrupos As String

If Len(Trim(gsCodUser)) <> 0 Then
   cGrupos = ObtenerGruposUsuario(gsCodUser, gsDominio)
End If
        
'sSQL = "select cName from Permiso where cGrupoUsu in (" & cGrupos & ") and len(rtrim(cName))=6 "
        
sSQL = "select * from OpeTpo " & _
       " where cOpeCod like '" & cRaizOpe & "%' and " & _
       "       cOpeVisible = '1' and " & _
       "       cOpeCod in (select distinct cName from Permiso where cGrupoUsu in (" & cGrupos & ") and len(rtrim(cName))=6) " & _
       " order by cOpeCod "

If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
End If

tvOpe.Nodes.Clear
If Rs.State = 0 Then Exit Sub
If Not Rs.EOF Then
   cClave = Rs!cOpeCod
   cDescripcion = Rs!cOpeDesc
   tvOpe.Nodes.Add , , "K", cClave + " - " + cDescripcion, "Padre"
   Rs.MoveNext
   Do While Not Rs.EOF
      cUlt = Right(Rs!cOpeCod, 1)
      If cUlt = "0" Then
         cCod = Left(Right(Rs!cOpeCod, 2), 1)
         cKeySup = "K"
      Else
         cCod = Right(Rs!cOpeCod, 2)
         cKeySup = "K" + Left(Right(Rs!cOpeCod, 2), 1)
      End If
      cKey = "K" + cCod
      'cKey = "K" + Replace(rs!cOpeCod, "0", "", 3)
      'cKeySup = "K"
      cClave = Rs!cOpeCod
      Select Case Rs!nOpeNiv
          Case 2
               cIMG = "Hijo"
          Case 3
               cIMG = "Hijito"
      End Select
      cDescripcion = Rs!cOpeDesc
      tvOpe.Nodes.Add cKeySup, tvwChild, cKey, cClave + " - " + cDescripcion, cIMG
      Rs.MoveNext
   Loop
   tvOpe.Nodes(1).Expanded = True
End If
End Sub

Private Sub tvOpe_DblClick()
Dim cOpeDescripcion As String
gsOpeCod = Trim(Mid(tvOpe.Nodes(tvOpe.SelectedItem.Index).Text, 1, InStr(tvOpe.Nodes(tvOpe.SelectedItem.Index).Text, "-") - 1))
cOpeDescripcion = Mid(tvOpe.Nodes(tvOpe.SelectedItem.Index).Text, 10, Len(tvOpe.Nodes(tvOpe.SelectedItem.Index).Text))
If Left(gsOpeCod, 2) = "53" Then
   Select Case gsOpeCod
    'Requerimientos ---------------------------
    Case "531011"
         frmLogPlanAnualReq.Show 1
    Case "531012"
         frmLogPlanAnualReqGen.Show 1
    Case "531013"
         frmLogPlanAnualReqConsolida.Show 1
    Case "531014"
         frmLogPlanAnualValorizacion.Show 1
    'Plan Anual -------------------------------
    Case "531021"
         frmLogPlanAnual.Show 1
    Case "531022"
         frmLogPlanAnualAprobacion.Show 1
    Case "531023"
         frmLogPlanAnualAprobacion.Show 1
    Case "531024"
         frmLogPlanAnualAprobacion.Show 1
    Case "531025"
         frmLogPlanAnualFormato.Show 1
    'Mantenimiento ----------------------------
    Case "531041"
         frmLogBSCatalogo.Show 1
         'frmLogBSCatalogo.Inicio 1
    Case "531042"
         frmLogMntBSGrupos.Show 1
    Case "531043"
         frmLogNiveles.Show 1
    Case "531044"
         frmLogMntBSPlanAnual.Show 1
  End Select
End If

If Left(gsOpeCod, 2) = "54" Then
   Select Case gsOpeCod
    'LOCACIÓN DE SERVICIOS ---------------------------
     Case "541011"
          frmLogLocadorRegistro.Show 1
     Case "541012"
          frmLogLocadorOperacion.Inicio gsOpeCod
     Case "541013"
          frmMntConstantes.Inicio 9131
     Case "541014"
          frmMntConstantes.Inicio 9132
     Case "541015"
          frmMntConstantes.Inicio 9133
          
     'PAGO DE SERVICIOS PUBLICOS ---------------------
     Case "541021"
          frmLogServiciosRegistro.Inicio gsOpeCod
     Case "541022"
          frmLogServiciosRegistro.Inicio gsOpeCod
     Case "541023"
          frmLogServiciosRegistro.Inicio gsOpeCod
          
     'PAGO DE SERVICIOS CORRESPONDENCIA ---------------------
     Case "541031"
          frmLogEnviosRegistro.Inicio gsOpeCod, True
     Case "541032"
          frmLogEnviosAprobacion.Inicio gsOpeCod, cOpeDescripcion
     Case "541033"
          frmLogEnviosAprobacion.Inicio gsOpeCod, cOpeDescripcion
     Case "541034"
          frmLogEnviosTarifas.Show 1
     Case "541035"
          'frmLogEnviosAprobacion.Inicio gsOpeCod, cOpeDescripcion
   End Select
End If

'Control Vehicular ---------------------------------
If Left(gsOpeCod, 2) = "55" Then
   Select Case gsOpeCod
    Case "551011"
         frmLogVehiculoMnt.Show 1
    Case "551012"
         frmLogVehiculoCond.Show 1
    Case "551013"
         frmLogVehiculoIncidencia.Show 1
    Case "551021"
         frmLogVehiculoSol.Modalidad 1
    Case "551022"
         frmLogVehiculoAprueba.CambiaEstado gcSolicitud, gcAprobado
    Case "551023"
         frmLogVehiculoRegistro.Estado gsCodPersUser, 3
    Case "551024"
         frmLogVehiculoRegistro.Estado gsCodPersUser, 4
    Case "551025"
         frmLogVehiculoAprueba.CambiaEstado gcAceptado, gcVistoBueno
    Case "551026"
         frmLogVehiculoSol.Modalidad 2
    Case "551031"
         'frmLogVehiculoRep.Show 1
         frmLogVehiculoMov.Show 1
    Case "551032"
         frmLogVehiculoRep.Show 1
'    Case "551041"
'         frmLogVehiculoTipoReg.Show 1
   End Select
End If

If Left(gsOpeCod, 4) = "5015" Then
   Select Case gsOpeCod
   
   'Requerimientos no Programados
    Case "501511"
        frmLogProSelReq.Show 1
    Case "501512"
        frmLogProSelReqConsolida.Show 1
    Case "501513"
        frmLogProSelValorizacion.Show 1
    Case "501514"
        frmLogProSelReqAprobacion.Inicio 1
    Case "501515"
        frmLogProSelReqAprobacion.Inicio 2
        
    'creacion y monitoreo
    Case "501521"
        frmLogProSelGenerarProcesoSeleccion.Show 1
    Case "501522"
         frmLogProSelReqAprobacion.Inicio 3
    Case "501523"
         frmLogProSelEjecucion.TipoFuncion 5, cOpeDescripcion
    Case "501524"
         frmLogProSelCnsProcesoSeleccion.Inicio 3
    Case "501525"
         frmLogProSelRptProcesos.Show 1
    
    'primera parte
    Case "501531"
         frmLogProSelEjecucion.TipoFuncion 3, cOpeDescripcion
    Case "501532"
         frmLogProSelEjecucion.TipoFuncion 1, cOpeDescripcion
    Case "501533"
         frmLogProSelEjecucion.TipoFuncion 2, cOpeDescripcion
    Case "501534"
         frmLogProSelEjecucion.TipoFuncion 6, cOpeDescripcion
    Case "501535"
         frmLogProSelEjecucion.TipoFuncion 7, cOpeDescripcion
    Case "501536"
         frmLogProSelAprobacion.TipoFuncion 1
    Case "501537"
         frmLogProSelAprobacion.TipoFuncion 2
         
    'final
    Case "501541"
          frmLogProSelEjecucion.TipoFuncion 8, cOpeDescripcion
    Case "501542"
         frmLogProSelEjecucion.TipoFuncion 9, cOpeDescripcion
    Case "501543"
         frmLogProSelEvaluacionValor.Inicio 1, cOpeDescripcion
         'frmLogEvaluacionValor.Inicio 1
    Case "501544"
         frmLogProSelEvaluacionValor.Inicio 2, cOpeDescripcion
    Case "501545"
         frmLogProSelEjecucion.TipoFuncion 4, cOpeDescripcion
    Case "501546"
         frmLogProSelEvaluacionValor.Inicio 3, cOpeDescripcion
        
    'contrato
    Case "501550"
         frmLogProSelContrato.Show 1
    
    'mantenimiento
    Case "501561"
         frmLogProSelTipos.Show 1
    Case "501562"
         frmLogProSelEtapas.Show 1
    Case "501563"
         frmLogProSelMntFactores.Show 1
    Case "501564"
         frmLogProSelMntFactoresEvaluacion.Show 1
  End Select
End If


End Sub

Private Sub tvOpe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tvOpe_DblClick
End If
End Sub

Function ObtenerGruposUsuario(psUsuCod As String, psDominio As String) As String
Dim oGroup As IADsGroup, oUser As IADsUser

On Error GoTo SalInfo

ObtenerGruposUsuario = ""
Set oUser = GetObject("WinNT://" & psDominio & "/" & psUsuCod & ",user")

For Each oGroup In oUser.Groups
    ObtenerGruposUsuario = ObtenerGruposUsuario + "'" + oGroup.Name + "',"
Next
'ObtenerGruposUsuario = Mid(ObtenerGruposUsuario, 1, Len(ObtenerGruposUsuario) - 1)
ObtenerGruposUsuario = ObtenerGruposUsuario + "'" + psUsuCod + "'"
Exit Function
SalInfo:
   ObtenerGruposUsuario = ""
   MsgBox "Usuario no válido o no hallado..." + Space(10), vbInformation, "Aviso"

End Function

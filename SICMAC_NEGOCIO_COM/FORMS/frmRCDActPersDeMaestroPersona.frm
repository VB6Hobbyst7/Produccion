VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRCDActPersDeMaestroPersona 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe RCD - Actualiza Personas desde RCDMaestroPersona"
   ClientHeight    =   1305
   ClientLeft      =   2175
   ClientTop       =   3555
   ClientWidth     =   6810
   Icon            =   "frmRCDActPersDeMaestroPersona.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   5400
      TabIndex        =   3
      Top             =   420
      Width           =   1245
   End
   Begin VB.CommandButton cmdActualiza 
      Caption         =   "&Actualiza Personas"
      Height          =   405
      Left            =   3120
      TabIndex        =   0
      Top             =   420
      Width           =   2085
   End
   Begin MSComctlLib.ProgressBar barraProgreso 
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   1020
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar EstatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7937
            MinWidth        =   7937
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRCDActPersDeMaestroPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
' RCD - Actualiza Personas desde el  RCDMaestroPersona
'LAYG   :  10/01/2003.
'Resumen:  Actualiza Personas desde el  RCDMaestroPersona

Option Explicit
Dim fsServConsol As String

Private Sub cmdActualiza_Click()
    
    Dim oRCD As COMDCredito.DCOMRCD
    Set oRCD = New COMDCredito.DCOMRCD
    Call oRCD.ActualizaPersonas_RCDMaestro(fsServConsol)
    Set oRCD = Nothing
    MsgBox "Proceso Finalizado", vbInformation, "Aviso"
'Dim lsSQL  As String
'Dim rs As ADODB.Recordset
'Dim rsCodigo As ADODB.Recordset
'Dim loBase As DConecta
'Dim lnContTotal  As Long, lnCont As Long
'Dim lsCodPers As String
'
'lsSQL = "Select * From " & fsServConsol & "RCDMaestroPersona "
'Set loBase = New DConecta
'    loBase.AbreConexion
'    Set rs = loBase.CargaRecordSet(lsSQL)
'
'    lnContTotal = rs.RecordCount
'    If Not RSVacio(rs) Then
'        Do While Not rs.EOF
'            lnCont = lnCont + 1
'
'            ' Emite Codigo de Persona
'            lsSQL = "Select * From " & fsServConsol & "RCDCodigoAux Where cCodAux='" & Trim(rs!cCodUnico) & "' "
'            Set rsCodigo = loBase.CargaRecordSet(lsSQL)
'            If rsCodigo.BOF And rsCodigo.EOF Then  ' No existe coge el codigo Persona
'                lsCodPers = Trim(rs!cPersCod)
'            Else ' Existe coge el Codigo de la tabla Auxiliar
'                lsCodPers = Trim(rsCodigo!cPersCod)
'            End If
'            rsCodigo.Close
'            Set rsCodigo = Nothing
'            ' Codigo SBS
'            lsSQL = "UPDATE PERSONA SET cPersCodSbs = '" & Trim(rs!cCodSBS) & "' " _
'                & " WHERE cPersCod = '" & Trim(lsCodPers) & "' "
'
'            loBase.Ejecutar (lsSQL)
'
'            barraProgreso.value = Int(lnCont / lnContTotal * 100)
'            EstatusBar.Panels(2).Text = "Avance : " & Format(lnCont / lnContTotal * 100, "#0.00") & "%"
'            DoEvents
'            rs.MoveNext
'        Loop
'    End If
'rs.Close
'Set rs = Nothing
'MsgBox "Proceso Finalizado", vbInformation, "Aviso"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim loConstSist As COMDConstSistema.NCOMConstSistema
    Set loConstSist = New COMDConstSistema.NCOMConstSistema
        fsServConsol = loConstSist.LeeConstSistema(gConstSistServCentralRiesgos)
    Set loConstSist = Nothing
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

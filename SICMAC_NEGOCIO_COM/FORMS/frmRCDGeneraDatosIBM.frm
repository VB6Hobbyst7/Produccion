VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRCDGeneraDatosIBM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe RCD - Genera Datos para Informe IBM"
   ClientHeight    =   2310
   ClientLeft      =   2265
   ClientTop       =   3405
   ClientWidth     =   6945
   Icon            =   "frmRCDGeneraDatosIBM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   6690
      Begin VB.CommandButton cmdSalir 
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
         Height          =   375
         Left            =   5145
         TabIndex        =   2
         Top             =   1515
         Width           =   1410
      End
      Begin VB.CommandButton cmdConsolida 
         Caption         =   "&Generar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   315
         TabIndex        =   1
         Top             =   1545
         Width           =   2325
      End
      Begin MSComctlLib.ProgressBar barra 
         Height          =   225
         Left            =   285
         TabIndex        =   3
         Top             =   1200
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblDescripcion 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   300
         TabIndex        =   7
         Top             =   930
         Width           =   2640
      End
      Begin VB.Label lblAvance 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4980
         TabIndex        =   6
         Top             =   960
         Width           =   1500
      End
      Begin VB.Label lblfecha 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2640
         TabIndex        =   5
         Top             =   420
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Consolidacion :"
         Height          =   195
         Left            =   420
         TabIndex        =   4
         Top             =   540
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frmRCDGeneraDatosIBM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fnMontoMinimoRCD As Double
Dim fnTipCambio As Currency
Dim fsServConsol As String

Private Sub cmdConsolida_Click()
If RegistroIBM = True Then
    MsgBox "El Proceso de Consolidacion ha culminado a las " & Now, vbInformation, "Aviso"
End If
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
Dim loRCDproc As COMNCredito.NCOMRCD
Dim lrPar As ADODB.Recordset

Dim loConstSist As COMDConstSistema.NCOMConstSistema
    Set loConstSist = New COMDConstSistema.NCOMConstSistema
        fsServConsol = loConstSist.LeeConstSistema(gConstSistServCentralRiesgos)
    Set loConstSist = Nothing
    
Set loRCDproc = New COMNCredito.NCOMRCD
    Set lrPar = loRCDproc.nCargaParametroRCD(Format(gdFecDataFM, "yyyymmdd"), fsServConsol)
Set loRCDproc = Nothing

    If lrPar.EOF And lrPar.BOF Then
        MsgBox "No se han ingresado Parametros de RCD", vbInformation, "Aviso"
        cmdConsolida.Enabled = False
    Else
        fnMontoMinimoRCD = lrPar!nMontoMin
        fnTipCambio = lrPar!nCambioFijo
    End If
Set lrPar = Nothing

Me.lblfecha.Caption = gdFecDataFM

End Sub

Private Function RegistroIBM() As Boolean
Dim oRCD As COMNCredito.NCOMRCD
Dim sMensaje As String
Dim bResp As Boolean

On Error GoTo ErrorRegistroIBM

Set oRCD = New COMNCredito.NCOMRCD

bResp = False

If MsgBox("Se Generaran los Saldos de los Clientes para el Informe IBM " & Chr(13), vbYesNo + vbQuestion, _
          "Aviso") = vbYes Then bResp = True

RegistroIBM = oRCD.GenerarDatosIBM(bResp, fnMontoMinimoRCD, gdFecDataFM, fnTipCambio, fsServConsol, sMensaje)
If sMensaje <> "" Then
    MsgBox sMensaje, vbInformation, "Mensaje"
End If
Set oRCD = Nothing
Exit Function
ErrorRegistroIBM:
    MsgBox "Error Nº [" & Err.Number & " ]" & Err.Description, vbInformation, "Aviso"
'Dim sql As String
'Dim rs As New ADODB.Recordset
'Dim rs1 As New ADODB.Recordset
'
'Dim lsCadConexon As String
'Dim loBase As COMConecta.DCOMConecta
'Dim CmdGenera As ADODB.Command
'Dim Param1 As ADODB.Parameter
'
'Dim loRCDProceso As COMNCredito.NCOMRCD
'
'Dim contTotal  As Long
'Dim j As Long
'Dim lsApellPat As String
'Dim lsApellMat As String
'Dim lsApellCony As String
'Dim lsNombre As String
'Dim Pos As Long
'Dim Pos1 As Long
'Dim Pos2 As Long
'Dim Cont As Long
'Dim lsPrefijos(6) As String
'Dim I As Long
'
'Dim lsCodPers As String
'Dim lsCadAux As String
'
'Dim lsTipoDoc As String
'Dim lsNumDoc As String
'Dim lsCalificacion As String
'Dim lsTipoPers As String
'Dim Total As Long
'Dim lsNombrePers As String
'Dim lvTipPers As String
'Dim lvTidoCi As String
'Dim lvNuDoCi  As String
'Dim lvTidoTr As String
'Dim lvNudoTr As String
'
'On Error GoTo ErrorIbm
'RegistroIBM = False
'Me.lblDescripcion = "Generando Saldos para Cada cliente por Favor espere"
'Me.lblAvance = ""
'Me.barra.value = 0
'If MsgBox("Se Generaran los Saldos de los Clientes para el Informe IBM " & Chr(13), vbYesNo + vbQuestion, _
'          "Aviso") = vbYes Then
'
'    Set loBase = New COMConecta.DCOMConecta
'
'    lsCadConexon = loBase.GetCadenaConexion(Right(gsCodAge, 2), "03")
'    'loBase.AbreConexionRemota "07", , False, "03"
'    loBase.AbreConexion lsCadConexon
'    loBase.ConexionActiva.BeginTrans
'
'    Set CmdGenera = New ADODB.Command
'    Set Param1 = New ADODB.Parameter
'
'    CmdGenera.CommandText = "RCDGeneraIBM"
'    CmdGenera.CommandType = adCmdStoredProc
'    CmdGenera.Name = "RCDGeneraIBM"
'    Set Param1 = CmdGenera.CreateParameter("MontoMin", adInteger, adParamInput)
'    CmdGenera.Parameters.Append Param1
'    Set Param1 = CmdGenera.CreateParameter("Fecha", adDate, adParamInput)
'    CmdGenera.Parameters.Append Param1
'    Set Param1 = CmdGenera.CreateParameter("TipoCambio", adCurrency, adParamInput)
'    CmdGenera.Parameters.Append Param1
'    Set CmdGenera.ActiveConnection = loBase.ConexionActiva
'    CmdGenera.CommandTimeout = 720
'    CmdGenera.Parameters.Refresh
'
'   loBase.ConexionActiva.RCDGeneraIBM fnMontoMinimoRCD, Format(gdFecDataFM, "mm/dd/yyyy"), fnTipCambio
'   loBase.ConexionActiva.CommitTrans
'
'   Set CmdGenera = Nothing
'   Set Param1 = Nothing
'
'End If
'
'sql = "Select cCodPers, cApellPat, cApellMat, cNombre, cTipPers, cTipDoc, cNumDoc, cCalifica " & _
'      "From " & fsServConsol & "IBM1 WHERE cApellPat IS Null "
'
'Set rs1 = loBase.CargaRecordSet(sql)
'contTotal = rs1.RecordCount
'
'If Not (rs1.BOF And rs1.EOF) Then
'    Do While Not rs1.EOF
'        Cont = Cont + 1
'        '  P.cTidotr cTidoTr, P.cNudotr cNudoTr, P.cTidoci cTidoCi, P.cNudoci cNudoCi,
'        sql = "SELECT P.cPersCod,P.cPersNombre cNomPers, P.nPersPersoneria cTipPers,  " _
'            & " R.cPersNom cNomPersRCD, R.cTipPers cTipPersRCD, R.cTidotr cTidoTrRCD, " _
'            & " R.cNudotr cNudoTrRCD, R.cTidoci cTidoCiRCD, R.cNudoci cNudoCiRCD " _
'            & "FROM Persona P LEFT JOIN " & fsServConsol & "RCDMaestroPersona R " _
'            & "ON P.cPersCod = R.cPersCod " _
'            & "WHERE P.cPersCod='" & Trim(rs1!cCodPers) & "'  "
'        Set rs = loBase.CargaRecordSet(sql)
'
'        If Not (rs.BOF And rs.EOF) Then
'                lsApellPat = "": lsApellMat = "": lsNombre = "": lsApellCony = ""
'                'lsNombrePers = NombreMaestro(Trim(rs!cCodPers))
'                lsNombrePers = IIf(IsNull(rs!cNomPersRCD), "", Trim(rs!cNomPersRCD))
'                If Len(Trim(lsNombrePers)) = 0 Then
'                    lsNombrePers = Trim(rs!cNomPers)
'                Else
'                    If InStr(1, lsNombrePers, "Y/O", vbTextCompare) = 0 And InStr(1, lsNombrePers, " Y ", vbTextCompare) = 0 And InStr(1, lsNombrePers, " O ", vbTextCompare) = 0 Then
'                        lsNombrePers = Trim(rs!cNomPers)
'                    End If
'                End If
'                lsNombrePers = Trim(Replace(lsNombrePers, "-", "", , , vbTextCompare))
'                lsNombrePers = Trim(Replace(lsNombrePers, ".", " ", , , vbTextCompare))
'                '-****Separa Nombre de Clientes
'                If InStr(1, lsNombrePers, "Y/O", vbTextCompare) = 0 And InStr(1, lsNombrePers, " Y ", vbTextCompare) = 0 And InStr(1, lsNombrePers, " O ", vbTextCompare) = 0 Then
'                    If Trim(rs!cTipPers) = "1" Then
'                        Pos1 = InStr(1, lsNombrePers, "/", vbTextCompare)
'                        If Pos1 > 0 Then
'                            lsApellPat = Mid(lsNombrePers, 1, Pos1 - 1)
'                            lsCadAux = Mid(lsNombrePers, Pos1 + 1, Len(lsNombrePers))
'
'                            Pos1 = InStr(1, lsCadAux, "\", vbTextCompare)
'                            If Pos1 > 0 Then
'                                lsApellMat = Trim(Mid(lsCadAux, 1, Pos1 - 1))
'                                lsCadAux = Mid(lsCadAux, Pos1 + 1, Len(lsCadAux))
'                                If Len(Trim(lsApellMat)) = 0 Then
'                                    Pos1 = InStr(1, lsCadAux, ",", vbTextCompare)
'                                    If Pos1 > 0 Then
'                                        lsApellCony = Trim(Mid(lsCadAux, 1, Pos1 - 1))
'                                        lsNombre = Trim(Mid(lsCadAux, Pos1 + 1, Len(lsCadAux)))
'                                    Else
'                                        lsApellCony = lsCadAux
'                                    End If
'                                Else
'                                    Pos1 = InStr(1, lsCadAux, ",", vbTextCompare)
'                                    If Pos1 > 0 Then
'                                        lsApellCony = Trim(Mid(lsCadAux, 1, Pos1 - 1))
'                                        lsNombre = Trim(Mid(lsCadAux, Pos1 + 1, Len(lsCadAux)))
'                                    Else
'                                        lsApellCony = lsCadAux
'                                    End If
'
'                                End If
'                            Else
'                                Pos1 = InStr(1, lsCadAux, ",", vbTextCompare)
'                                If Pos1 > 0 Then
'                                    lsApellMat = Trim(Mid(lsCadAux, 1, Pos1 - 1))
'                                    lsNombre = Trim(Mid(lsCadAux, Pos1 + 1, Len(lsCadAux)))
'                                End If
'                            End If
'                        Else
'                            lsApellPat = lsNombrePers
'                            lsApellMat = ""
'                            lsNombre = ""
'                        End If
'                    Else
'                        lsApellPat = lsNombrePers
'                        lsApellMat = ""
'                        lsNombre = ""
'                    End If
'                    If Len(Trim(lsApellCony)) > 0 Then
'                        lsCadAux = lsNombre
'                        lsNombre = lsApellPat & IIf(Len(Trim(lsApellMat)) = 0, " ", " " & Trim(lsApellMat) & " ") & "DE"
'                        lsApellMat = lsCadAux
'                        lsApellPat = lsApellCony
'                    End If
'                Else
'                    lsApellPat = lsNombrePers
'                    lsApellMat = ""
'                    lsNombre = ""
'                End If
'                '-******
'                'lvTipPers = IIf(IsNull(rs!cTipPersRCD), rs!cTipPers, rs!cTipPersRCD)
'                lvTipPers = Trim(rs!cTipPers)
'                'lvTidoCi = IIf(IsNull(rs!cTiDoCiRCD), IIf(IsNull(rs!ctidoci), "", rs!ctidoci), rs!cTiDoCiRCD)
'                'lvNuDoCi = IIf(IsNull(rs!cNudoCiRCD), IIf(IsNull(rs!cnudoci), "", rs!cnudoci), rs!cNudoCiRCD)
'                'lvTidoTr = IIf(IsNull(rs!cTiDoTrRCD), IIf(IsNull(rs!cTidoTr), "", rs!cTidoTr), rs!cTiDoTrRCD)
'                'lvNudoTr = IIf(IsNull(rs!cNudoTrRCD), IIf(IsNull(rs!cNudoTr), "", rs!cNudoTr), rs!cNudoTrRCD)
'
'                Select Case lvTipPers
'                    Case "1"
'                        lsTipoPers = "1"
'                        lsNumDoc = Trim(IIf(IsNull(lvNuDoCi), "", lvNuDoCi))
'                        Select Case Trim(lvTidoCi)
'                            Case "1"
'                                lsTipoDoc = "1"
'                            Case "3", "4" ' Carnet FFAA / FFPP
'                                lsTipoDoc = "2"
'                            Case "2", "5" ' Carnet Extranjeria / Pasaporte
'                                lsTipoDoc = "3"
'                        End Select
'                    Case Else
'                        If (InStr(1, lsApellPat, "Y/O", vbTextCompare) <> 0 Or InStr(1, lsApellPat, " O ", vbTextCompare) <> 0 Or InStr(1, lsApellPat, " Y ", vbTextCompare) <> 0) And (Len(Trim(lvNudoTr)) = 0) Then
'                            lsTipoPers = "1"
'                            lsNumDoc = Trim(IIf(IsNull(lvNuDoCi), "", lvNuDoCi))
'                            Select Case Trim(lvTidoCi)
'                                Case "1"
'                                    lsTipoDoc = "1"
'                                Case "3", "4" ' Carnet FFAA / FFPP
'                                    lsTipoDoc = "2"
'                                Case "2", "5" ' Carnet Extranjeria / Pasaporte
'                                    lsTipoDoc = "3"
'                            End Select
'                        Else
'                            lsTipoPers = "2"
'                            lsNumDoc = Trim(IIf(IsNull(lvNudoTr), "", lvNudoTr))
'                            'lsNumDoc = PersonaRUC11(rs!cCodPers)
'                            'lsNumDoc = IIf(Len(lsNumDoc) > 1, lsNumDoc, Trim(IIf(IsNull(rs!cNuDoTr), "", rs!cNuDoTr)))
'                        End If
'                        Select Case Trim(lvTidoTr)
'                            Case "2", "3"
'                                lsTipoDoc = "4"
'
'                        End Select
'                End Select
'
'                '******************** CALIFICACION DE LA PERSONA **************
'                Set loRCDProceso = New COMNCredito.NCOMRCD
'                    lsCalificacion = loRCDProceso.nObtieneCalificacionPersonaProcesada(rs!cPersCod, fsServConsol)
'                Set loRCDProceso = Nothing
'
'                If Trim(lsCalificacion) = "" Then
'                    MsgBox " OJO, no se esta considerando calificacion "
'                End If
'
'                'sql = "Update ibm1 set cApellpat='" & Replace(lsApellPat, "'", "''") & "',cApellMat='" _
'                    & Replace(lsApellMat, "'", "''") & "',cNombre='" & Replace(lsNombre, "'", "''") & "'," _
'                    & "cCalifica='" & Trim(lsCalificacion) & "'," _
'                    & "cTipDoc='" & lsTipoDoc & "',cNumdoc='" & lsNumDoc & "',cTipPers='" & lsTipoPers & "' where ccodpers='" & Trim(rs1!cCodPers) & "' "
'                'dbCmact.Execute sql
'                loBase.Ejecutar " RCDActualizaIBM '" & Replace(lsApellPat, "'", "''") & "','" & Replace(lsApellMat, "'", "''") & "','" & Replace(lsNombre, "'", "''") & "','" & Trim(lsCalificacion) & "','" & lsTipoDoc & "','" & lsNumDoc & "','" & lsTipoPers & "','" & Trim(rs1!cCodPers) & "'"
'        End If
'        rs.Close
'        Set rs = Nothing
'
'        'Me.lblDescripcion = "Cliente : " & Trim(rs1!cPersCod)
'        Me.barra.value = Int(Cont / contTotal * 100)
'        Me.lblAvance = "Avance : " & Format(Cont / contTotal * 100, "#0.000") & "%"
'        DoEvents
'        rs1.MoveNext
'    Loop
'End If
'rs1.Close
'Set rs1 = Nothing
'RegistroIBM = True
'MsgBox "Proceso de Generacion IBM ha culminado !!! ", vbInformation, "Aviso"
'Exit Function
'ErrorIbm:
'    MsgBox "Error Nº [" & Err.Number & " ]" & Err.Description, vbInformation, "Aviso"
'    RegistroIBM = False
End Function


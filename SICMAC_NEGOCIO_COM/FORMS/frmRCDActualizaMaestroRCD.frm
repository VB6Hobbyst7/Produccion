VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRCDActualizaRCDMaestroPersona 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe RCD - Actualiza Datos de RCDMaestroPersona"
   ClientHeight    =   2760
   ClientLeft      =   2400
   ClientTop       =   2940
   ClientWidth     =   5940
   Icon            =   "frmRCDActualizaMaestroRCD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Actualiza  RCDMaestroPersona"
      Height          =   2715
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   195
         Left            =   2880
         TabIndex        =   12
         Top             =   2400
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   195
         Left            =   1680
         TabIndex        =   11
         Top             =   2400
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   2400
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "..."
         Height          =   330
         Left            =   5220
         TabIndex        =   8
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox txtruta 
         Height          =   330
         Left            =   540
         TabIndex        =   7
         Top             =   240
         Width           =   4695
      End
      Begin VB.CommandButton cmdActualizaMesAnterior 
         Caption         =   "Actualiza del Mes Anterior"
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
         Left            =   240
         TabIndex        =   1
         Top             =   1320
         Width           =   4215
      End
      Begin MSComctlLib.ProgressBar Barra 
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdSalir 
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
         Height          =   360
         Left            =   4620
         TabIndex        =   0
         Top             =   2220
         Width           =   975
      End
      Begin VB.CommandButton cmdActualizaSBS 
         Caption         =   "Actualiza desde Archivo enviado por la SBS"
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
         Left            =   240
         TabIndex        =   2
         Top             =   1800
         Width           =   4215
      End
      Begin MSComDlg.CommonDialog cmdlOpen 
         Left            =   4860
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ruta  "
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   435
      End
      Begin VB.Label lblDato 
         AutoSize        =   -1  'True
         Caption         =   "lblDato"
         Height          =   195
         Left            =   1320
         TabIndex        =   6
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Actualizando :"
         Height          =   195
         Left            =   255
         TabIndex        =   4
         Top             =   600
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmRCDActualizaRCDMaestroPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
' RCD - Actualiza la tabla RCDMaestroPersona
'LAYG   :  10/01/2003.
'Resumen:  Nos permite actualizar la tabla RCDMaestroPersona

Option Explicit

Private Sub cmdActualizaMesAnterior_Click()
Dim lsSQL As String
Dim rs As ADODB.Recordset
Dim rsCodAux As ADODB.Recordset
Dim loBase As DConecta
Dim lnTotal As Long, j As Long
Dim lnNuevos As Long, lnModif As Long
Dim rs1 As ADODB.Recordset
Dim lbNuevoInsert As Boolean
Dim lsCodigoPersona As String
'On Error GoTo ErrorActMaAnt
lsSQL = "SELECT * FROM RCDvc" & Format(gdFecSis - Day(gdFecSis), "yyyymm") & "01"
Set loBase = New DConecta
    loBase.AbreConexion
    Set rs = loBase.CargaRecordSet(lsSQL)
    
    lnTotal = rs.RecordCount
    j = 0
    lnNuevos = 0
    lnModif = 0
    If Not (rs.BOF And rs.EOF) Then
        Do While Not rs.EOF
            j = j + 1
            '*** BUSCO EL CODIGO DE PERSONA (TABLA AUXILIAR)
            lsSQL = "SELECT * FROM RCDCodigoAux WHERE cCodAux='" & Trim(rs!cPersCod) & "'"
            
            Set rsCodAux = loBase.CargaRecordSet(lsSQL)
            If Not (rsCodAux.BOF And rsCodAux.EOF) Then
                lsCodigoPersona = Trim(rsCodAux!cPersCod)
            Else
                lsCodigoPersona = Trim(rs!cPersCod)
            End If
            rsCodAux.Close
            Set rsCodAux = Nothing
            '********
            
            lsSQL = "SELECT cPersCod FROM RCDMaestroPersona WHERE cPersCod ='" & lsCodigoPersona & "' "
            Set rs1 = loBase.CargaRecordSet(lsSQL)
            If rs1.BOF And rs1.EOF Then ' No existe
                lbNuevoInsert = True
            Else
                lbNuevoInsert = False
            End If
            rs1.Close
            Set rs1 = Nothing
            
            If lbNuevoInsert = True Then
                lnNuevos = lnNuevos + 1
                
                lsSQL = "INSERT INTO RCDMaestroPersona (cPersCod, cCodUnico, cCodSBS, cPersNom, " _
                    & " cActEcon, cCodRegPub, cTidoTr, cNudoTr, cTiDoci, cNuDoci, cTipPers, cResid, " _
                    & " cMagEmp, cAccionista, cRelInst, cPaisNac, cSiglas) " _
                    & " VALUES ('" & lsCodigoPersona & "','" & Trim(rs!cPersCod) & "','" & Trim(rs!cCodSBS) & "','" _
                    & Trim(Replace(rs!cPersNom, "'", "''")) & "','" & Trim(rs!cActEcon) & "','" & IIf(IsNull(rs!ccodregpub), "", rs!ccodregpub) & "','" _
                    & IIf(IsNull(rs!cTidoTr), "", IIf(Trim(rs!cTidoTr) = "4", "", Trim(rs!cTidoTr))) & "','" _
                    & IIf(IsNull(rs!cNudoTr), "", Trim(rs!cNudoTr)) & "','" _
                    & IIf(IsNull(rs!ctidoci), "", IIf(Trim(rs!ctidoci) = "9", "", Trim(rs!ctidoci))) & "','" _
                    & IIf(IsNull(rs!cnudoci), "", Trim(rs!cnudoci)) & "','" _
                    & Trim(rs!cTipPers) & "','" & Trim(rs!cResid) & "','" & IIf(IsNull(rs!cMagEmp), "", Trim(rs!cMagEmp)) & "','" _
                    & IIf(IsNull(rs!cAccionista), "", Trim(rs!cAccionista)) & "','" & IIf(IsNull(rs!cRelInst), "", Trim(rs!cRelInst)) & "','" _
                    & IIf(IsNull(rs!cPaisNac), "", Trim(rs!cPaisNac)) & "','" _
                    & IIf(IsNull(rs!cSiglas), "", Trim(rs!cSiglas)) & "' ) "
                
                loBase.Ejecutar (lsSQL)
            Else
                lnModif = lnModif + 1
                ' Modificar para que solo actualize el codigo SBS
                'lsSQL = "UPDATE RCDMaestroPersona SET cCodSBS ='" & Trim(rs!cCodSBS) & "', " _
                    & " cNomPers='" & Trim(Replace(rs!cNomPers, "'", "''")) & "', " _
                    & " cActEcon='" & IIf(IsNull(rs!cActEcon), "", Trim(rs!cActEcon)) & "'," _
                    & " cTidoTr='" & IIf(IsNull(rs!cTidoTr), "", IIf(Trim(rs!cTidoTr) = "4", "", Trim(rs!cTidoTr))) & "'," _
                    & " cNudoTr='" & IIf(IsNull(rs!cNudoTr), "", Trim(rs!cNudoTr)) & "'," _
                    & " cTiDoci='" & IIf(IsNull(rs!cTiDoci), "", IIf(Trim(rs!cTiDoci) = "9", "", Trim(rs!cTiDoci))) & "'," _
                    & " cNuDoci='" & IIf(IsNull(rs!cNuDoci), "", Trim(rs!cNuDoci)) & "'," _
                    & " cCodRegPub='" & IIf(IsNull(rs!ccodregpub), "", rs!ccodregpub) & "'," _
                    & " cSiglas='" & IIf(IsNull(rs!cSiglas), "", Trim(rs!cSiglas)) & "' " _
                    & " WHERE cCodPers='" & Trim(lsCodigoPersona) & "'"
                'loBase.Ejecutar (lsSQL)
                
            End If
            barra.Value = Int(j / lnTotal * 100)
            Me.lblDato.Caption = Trim(rs!cPersCod) & "  Nuevos :" & lnNuevos & " - Modificados :" & lnModif
            rs.MoveNext
            DoEvents
        Loop
    End If
    rs.Close
    Set rs = Nothing
Set loBase = Nothing

MsgBox "Actualización Finalizada con Exito", vbInformation, "Aviso"
Exit Sub
ErrorActMaAnt:
    MsgBox "Error Nº[" & Err.Number & " ] " & Err.Description, vbInformation, "Aviso"

End Sub

Private Sub cmdActualizaSBS_Click()
Dim lsSQL As String
Dim loBase As DConecta
Dim j As Long
Dim lnNuevos As Long, lnModif    As Long
Dim lsLinea As String
Dim lsArchivo As String

Dim lsCodSbs  As String, lsCodDEUDOR   As String
Dim lsNomPers As String

Const lnPosCodSBS = 9
Const lnPosCodDEUDOR = 19
Const lnPosNomPers = 135

On Error GoTo ErrorActMa

If Len(Trim(Me.txtRuta)) <= 0 Then
    MsgBox "Nombre de Tabla no válido", vbInformation, "Aviso"
    Exit Sub
End If
    
lsArchivo = Trim(txtRuta)
If Dir(lsArchivo) = "" Then Exit Sub

cmdActualizaSBS.Enabled = True
j = 0
lnNuevos = 0
lnModif = 0

Open lsArchivo For Input As #1   ' Abre el archivo.
Set loBase = New DConecta
    loBase.AbreConexion
    
    Do While Not EOF(1)   ' Repite el bucle hasta el final del archivo.
        Input #1, lsLinea
        If Len(Trim(lsLinea)) > 0 Then ' Linea tiene datos
            If Mid(lsLinea, 1, 2) = "11" Then ' Datos de cliente
    
                ' Limpio las variables
                lsCodSbs = "": lsCodDEUDOR = "": lsNomPers = ""
                
                lsCodSbs = Mid(lsLinea, lnPosCodSBS, 10)
                lsCodDEUDOR = Mid(lsLinea, lnPosCodDEUDOR, 10)
                lsNomPers = Mid(lsLinea, lnPosNomPers, 40)
    
                '  solo actualiza el codigo SBS
                lsSQL = "UPDATE RCDMaestroPersona SET cCodSBS ='" & Trim(lsCodSbs) & "' " _
                    & " WHERE cCodUnico='" & Trim(lsCodDEUDOR) & "' "
                
                loBase.Ejecutar (lsSQL)
                
                lnModif = lnModif + 1
                
                'Barra.Value = Int((Asc(lsNomPers) - 64) / (Asc("Z") - 35) * 100)
                Me.lblDato.Caption = " Modificados :" & lnModif & " - " & Mid(lsNomPers, 1, 20)
                DoEvents
            End If
        End If
        DoEvents
    Loop
Set loBase = Nothing
Close #1   ' Cierra el archivo.

MsgBox "Actualización Finalizada con Exito", vbInformation, "Aviso"
Exit Sub
ErrorActMa:
    MsgBox "Error Nº[" & Err.Number & " ] " & Err.Description, vbInformation, "Aviso"
 
End Sub

Private Sub cmdOpen_Click()
   ' Establecer CancelError a True
    cmdlOpen.CancelError = True
    On Error GoTo ErrHandler
    ' Establecer los indicadores
    cmdlOpen.Flags = cdlOFNHideReadOnly
    cmdlOpen.InitDir = App.Path
    ' Establecer los filtros
    cmdlOpen.Filter = "Todos los archivos (*.*)|*.*|Archivos de texto" & _
    "(*.txt)|*.txt|Archivos por lotes (*.bat)|*.bat|Archivos 112 (*.112)|*.112"
    ' Especificar el filtro predeterminado
    cmdlOpen.FilterIndex = 2
    ' Presentar el cuadro de diálogo Abrir
    cmdlOpen.ShowOpen
    ' Presentar el nombre del archivo seleccionado
    txtRuta = cmdlOpen.FileName
    Exit Sub
    
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
    Exit Sub
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
'Dim sql As String
'Dim rs As New ADODB.Recordset
'Dim rsCodAux As New ADODB.Recordset
'Dim lnTotal As Long
'Dim j As Long
'Dim lnNuevos As Long, lnModif   As Long
'
'Dim lsCodigoPersona As String
''On Error GoTo ErrorActMaAnt
''sql = "select * from RCDVC01062001 where ccodpers in " _
' & "( select ccodpers from rcdmaestropersona where cNudoTr ='' and ctidotr ='3' ) "
'
''sql = "select * from rcdmaestropersona where ccodpers in " _
' & " (select ccodpers from ibm1 where cnumdoc ='' and ctippers <> '1' ) "
'
''sql = "select * from rcdmaestropersona where ccodpers in " _
' & " (select ccodpers from RCDVC01072001 where cNudoTr ='' and ctidotr ='3'  ) "
'
''sql = "select * from persona where ccodpers in " _
' & " (select ccodpers from rcdmaestropersona where cNudoTr ='' and ctidotr ='3'  ) "
'
'sql = " select h.codsbs, h.coddeu, m.cCodPers, m.cNomPers, m.cCodSBS " _
'    & " from rcdvc10 h inner join rcdvc01112001 m " _
'    & " on h.CodDeu = m.cCodPers " _
'    & " where m.cCodsbs =''  "
'
'rs.CursorLocation = adUseClient
'rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'Set rs.ActiveConnection = Nothing
'
'lnTotal = rs.RecordCount
'j = 0
'Nuevos = 0
'Modif = 0
'If Not RSVacio(rs) Then
'    Do While Not rs.EOF
'        j = j + 1
'        '*** BUSCO EL CODIGO DE PERSONA
'        sql = "Select * from CodigoAux where cCodAux='" & Trim(rs!cCodPers) & "'"
'        rsCodAux.Open sql, dbCmact, adOpenForwardOnly, adLockOptimistic, adCmdText
'        If Not RSVacio(rsCodAux) Then
'            lsCodigoPersona = Trim(rsCodAux!cCodPers)
'        Else
'            lsCodigoPersona = Trim(rs!cCodPers)
'        End If
'        rsCodAux.Close
'        Set rsCodAux = Nothing
'        '********
'            Modif = Modif + 1
'            ' ACTUALIZAR EL NRO DOC TRIBUTARIO
'            sql = " UPDATE RCDvc01112001 SET " _
'                & " cCodSBS ='" & rs!CODSBS & "' " _
'                & " WHERE cCodPers='" & Trim(lsCodigoPersona) & "'"
'            'sql = "UPDATE RCDMaestroPersona SET " _
'                & " cNudoTr='" & IIf(IsNull(rs!cNudoTr), "", Trim(rs!cNudoTr)) & "' " _
'                & " WHERE cCodPers='" & Trim(lsCodigoPersona) & "'"
'            'sql = "UPDATE IBM1 SET " _
'                & " cNumDoc='" & IIf(IsNull(rs!cNudoTr), "", Trim(rs!cNudoTr)) & "' " _
'                & " WHERE cCodPers='" & Trim(lsCodigoPersona) & "'"
'            'sql = "UPDATE RCDvc01072001 SET " _
'                & " cNudoTr='" & IIf(IsNull(rs!cNudoTr), "", Trim(rs!cNudoTr)) & "' " _
'                & " WHERE cCodPers='" & Trim(lsCodigoPersona) & "'"
'
'            'sql = "UPDATE RCDvc01092001 SET " _
'                & " cNudoTr='" & IIf(IsNull(rs!cNudoTr), "", Trim(rs!cNudoTr)) & "' " _
'                & " WHERE cCodPers='" & Trim(lsCodigoPersona) & "'"
'
'            dbCmact.Execute sql
'
'
'        Barra.Value = Int(j / lnTotal * 100)
'        Me.lblDato.Caption = Trim(rs!cCodPers) & "  Nuevos :" & Nuevos & " - Modificados :" & Modif
'        rs.MoveNext
'        DoEvents
'    Loop
'End If
'rs.Close
'Set rs = Nothing
'
'MsgBox "Actualización Finalizada con Exito", vbInformation, "Aviso"
'Exit Sub
'ErrorActMaAnt:
'    MsgBox "Error Nº[" & Err.Number & " ] " & Err.Description, vbInformation, "Aviso"


End Sub

Private Sub Command2_Click()
'Dim sql As String
'Dim rs As New ADODB.Recordset
'Dim rsCodAux As New ADODB.Recordset
'Dim lnTotal As Long
'Dim j As Long
'Dim lnNuevos As Long, lnModif   As Long
'
'Dim lsCodigoPersona As String
''On Error GoTo ErrorActMaAnt
'
''sql = "select * from persona where ccodpers in " _
'     & " (select ccodpers from rcdmaestropersona where cNudoTr ='' and ctidotr ='3'  ) "
'
'sql = " select re.COD_SBS, re.APE_PAT, re.APE_MAT, re.Nombre,  " _
'    & " rcd.cCodPers, rcd.cNomPers, rcd.cCodSBS " _
'    & " from reniec re inner join rcdvc01122001 rcd " _
'    & " on rtrim(ltrim(re.num_doc)) = rtrim(ltrim(rcd.cNudoci)) " _
'    & " where rcd.cCodsbs =''  "
'
''select * from reniec re inner join
''rcdvc01112001  rcd on re.num_doc = rcd.cnudoci
''where rcd.cCodSBS = ''
'
'rs.CursorLocation = adUseClient
'rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'Set rs.ActiveConnection = Nothing
'
'lnTotal = rs.RecordCount
'j = 0
'Nuevos = 0
'Modif = 0
'If Not RSVacio(rs) Then
'    Do While Not rs.EOF
'        j = j + 1
'        '*** BUSCO EL CODIGO DE PERSONA
'        sql = "Select * from CodigoAux where cCodAux='" & Trim(rs!cCodPers) & "'"
'        rsCodAux.Open sql, dbCmact, adOpenForwardOnly, adLockOptimistic, adCmdText
'        If Not RSVacio(rsCodAux) Then
'            lsCodigoPersona = Trim(rsCodAux!cCodPers)
'        Else
'            lsCodigoPersona = Trim(rs!cCodPers)
'        End If
'        rsCodAux.Close
'        Set rsCodAux = Nothing
'        '********
'
'            'If MsgBox(rs!APE_PAT & " " & rs!APE_MAT & " " & rs!Nombre & " -- " & rs!cNomPers, vbYesNo, "Aviso") = vbNo Then
'            '    MsgBox " Aguanta"
'            'End If
'
'            Modif = Modif + 1
'            ' ACTUALIZAR CODIGO SBS
'            'sql = " UPDATE RCDvc01112001 SET " _
'                & " cCodSBS ='" & rs!Cod_SBS & "' " _
'                & " WHERE cCodPers='" & Trim(lsCodigoPersona) & "'"
'            sql = " UPDATE RCDvc01122001 SET  " _
'                & " cCodSBS ='" & rs!Cod_SBS & "' " _
'                & " WHERE cCodPers='" & Trim(lsCodigoPersona) & "'"
'            'sql = "UPDATE IBM1 SET " _
'                & " cNumDoc='" & IIf(IsNull(rs!cNudoTr), "", Trim(rs!cNudoTr)) & "' " _
'                & " WHERE cCodPers='" & Trim(lsCodigoPersona) & "'"
'
'            dbCmact.Execute sql
'
'
'        Barra.Value = Int(j / lnTotal * 100)
'        Me.lblDato.Caption = Trim(rs!cCodPers) & "  Nuevos :" & Nuevos & " - Modificados :" & Modif
'        rs.MoveNext
'        DoEvents
'    Loop
'End If
'rs.Close
'Set rs = Nothing
'
'MsgBox "Actualización Finalizada con Exito", vbInformation, "Aviso"
'Exit Sub
'ErrorActMaAnt:
'    MsgBox "Error Nº[" & Err.Number & " ] " & Err.Description, vbInformation, "Aviso"

End Sub

Private Sub Command3_Click()
' Actualiza el codigo de SBS desde el archivo RCC enviado por la SBS
' Y en el sengundo el rcdvc01 del mes
'Dim sql As String
'Dim rs As New ADODB.Recordset
'Dim rsCodAux As New ADODB.Recordset
'Dim lnTotal As Long
'Dim j As Long
'Dim lnNuevos As Long, lnModif   As Long
'
'Dim lsCodigoPersona As String
''On Error GoTo ErrorActMaAnt
'
''sql = "select * from persona where ccodpers in " _
'     & " (select ccodpers from rcdmaestropersona where cNudoTr ='' and ctidotr ='3'  ) "
'
''sql = " select  rcc.cCodSBS rcc_cCodSBS, rcc.cTiDoCi, rcc.cNuDoCi, rcc.cDeudor , " _
'    & " rm.cCodSBS rm_cCodSBS , rm.cNomPers, rm.cTidoci, rm.cNudoci, rM.cCodPers " _
'    & " from rcc rcc inner join rcdMaestroPersona rM on " _
'    & " rtrim(ltrim(rcc.cTidoci)) = rtrim(ltrim(rM.cTidoci)) and " _
'    & " RTrim(LTrim(rcc.cNudoci)) = RTrim(LTrim(rm.cNudoci)) " _
'    & " and rm.cCodSBS ='' "
'
'
'sql = " Select  rcc.cCodSBS rcc_cCodSBS, rcc.cTiDoCi, rcc.cNuDoCi, " _
'    & " rM.cCodSBS rM_cCodSBS , rM.cNomPers, rM.cTidoci, rM.cNudoci, rM.ccodpers " _
'    & " from rcdvc01122001 rcc inner join rcdMaestroPersona rM on " _
'    & " rtrim(ltrim(rcc.cTidoci)) = rtrim(ltrim(rM.cTidoci)) and " _
'    & " RTrim(LTrim(rcc.cNudoci)) = RTrim(LTrim(rM.cNudoci)) " _
'    & " and rcc.cCodSBS ='' and len(rM.cCodSBS) > 0 "
'
'rs.CursorLocation = adUseClient
'rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'Set rs.ActiveConnection = Nothing
'
'lnTotal = rs.RecordCount
'j = 0
'Nuevos = 0
'Modif = 0
'If Not RSVacio(rs) Then
'    Do While Not rs.EOF
'        j = j + 1
'        '*** BUSCO EL CODIGO DE PERSONA
'        sql = "Select * from CodigoAux where cCodAux='" & Trim(rs!cCodPers) & "'"
'        rsCodAux.Open sql, dbCmact, adOpenForwardOnly, adLockOptimistic, adCmdText
'        If Not RSVacio(rsCodAux) Then
'            lsCodigoPersona = Trim(rsCodAux!cCodPers)
'        Else
'            lsCodigoPersona = Trim(rs!cCodPers)
'        End If
'        rsCodAux.Close
'        Set rsCodAux = Nothing
'        '********
'
'            'If MsgBox(rs!APE_PAT & " " & rs!APE_MAT & " " & rs!Nombre & " -- " & rs!cNomPers, vbYesNo, "Aviso") = vbNo Then
'            '    MsgBox " Aguanta"
'            'End If
'
'            Modif = Modif + 1
'            ' ACTUALIZAR CODIGO SBS
'            'sql = " UPDATE RCDMaestroPersona SET " _
'                & " cCodSBS ='" & rs!rcc_cCodSBS & "' " _
'                & " WHERE cCodPers='" & Trim(lsCodigoPersona) & "'"
'            sql = " UPDATE RCDvc01122001 SET  " _
'                & " cCodSBS ='" & rs!rM_cCodSBS & "' " _
'                & " WHERE cCodPers='" & Trim(lsCodigoPersona) & "'"
'
'            dbCmact.Execute sql
'
'
'        Barra.Value = Int(j / lnTotal * 100)
'        Me.lblDato.Caption = Trim(rs!cCodPers) & "  Nuevos :" & Nuevos & " - Modificados :" & Modif
'        rs.MoveNext
'        DoEvents
'    Loop
'End If
'rs.Close
'Set rs = Nothing
'
'MsgBox "Actualización Finalizada con Exito", vbInformation, "Aviso"
'Exit Sub
'ErrorActMaAnt:
'    MsgBox "Error Nº[" & Err.Number & " ] " & Err.Description, vbInformation, "Aviso"

End Sub

Private Sub Form_Load()
    Me.lblDato = ""
    Me.txtRuta = App.Path
End Sub


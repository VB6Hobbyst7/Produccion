VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmITFGeneraArchivos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generacion de  Archivos"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar PB 
      Height          =   300
      Left            =   195
      TabIndex        =   13
      Top             =   3600
      Visible         =   0   'False
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame fraFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   630
      Left            =   450
      TabIndex        =   10
      Top             =   30
      Width           =   4095
      Begin MSMask.MaskEdBox txtFechaF 
         Height          =   300
         Left            =   2475
         TabIndex        =   1
         Top             =   225
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   300
         Left            =   480
         TabIndex        =   0
         Top             =   225
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Del:"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   278
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Al:"
         Height          =   195
         Left            =   2205
         TabIndex        =   11
         Top             =   285
         Width           =   180
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   2385
      TabIndex        =   8
      Top             =   3045
      Width           =   1755
   End
   Begin VB.CommandButton CmdProcesar 
      Caption         =   "&Procesar"
      Height          =   420
      Left            =   570
      TabIndex        =   7
      Top             =   3030
      Width           =   1755
   End
   Begin VB.Frame Frame1 
      Height          =   2160
      Left            =   105
      TabIndex        =   9
      Top             =   720
      Width           =   4920
      Begin VB.CheckBox chkOption 
         Caption         =   "Archivo de Sustentacion de Movimientos"
         Height          =   300
         Index           =   5
         Left            =   210
         TabIndex        =   14
         Top             =   1635
         Width           =   3255
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Archivo de Movimientos de Contra Asiento por Error"
         Height          =   300
         Index           =   3
         Left            =   195
         TabIndex        =   5
         Top             =   1050
         Width           =   4215
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Archivo de movimientos acumulados por contribuyente"
         Height          =   300
         Index           =   2
         Left            =   195
         TabIndex        =   4
         Top             =   760
         Width           =   4410
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Archivo de Exonerados"
         Height          =   300
         Index           =   4
         Left            =   195
         TabIndex        =   6
         Top             =   1320
         Width           =   3255
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Archivos de contribuyentes extranjeros"
         Height          =   300
         Index           =   1
         Left            =   195
         TabIndex        =   3
         Top             =   495
         Width           =   3480
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Archivo de contribuyentes"
         Height          =   300
         Index           =   0
         Left            =   195
         TabIndex        =   2
         Top             =   240
         Width           =   2190
      End
   End
   Begin VB.OLE OleExcel 
      Class           =   "Excel.Sheet.8"
      Height          =   870
      Left            =   105
      OleObjectBlob   =   "FrmITFGeneraArchivos.frx":0000
      TabIndex        =   15
      Top             =   1980
      Visible         =   0   'False
      Width           =   1800
   End
End
Attribute VB_Name = "FrmITFGeneraArchivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim J As Long
Dim nTotal As Long

Private Sub chkOption_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdProcesar.SetFocus
End If
End Sub

Private Sub cmdProcesar_Click()
Dim i As Integer
Dim Mfecha As String

If txtFecha <> "__/__/____" Then
    Mfecha = ValidaFecha(txtFecha.Text)
    If Mfecha <> "" Then
        MsgBox Mfecha, vbInformation, "Aviso"
        Me.txtFecha.SetFocus
        Exit Sub
    End If
End If
If txtFecha = "__/__/____" Then
    MsgBox "Por favor Ingrese una Fecha", vbInformation, "Aviso"
    txtFecha.SetFocus
    Exit Sub
End If
Mfecha = ValidaFecha(txtFechaF.Text)
If Mfecha <> "" Then
    MsgBox Mfecha, vbInformation, "Aviso"
    Me.txtFechaF.SetFocus
    Exit Sub
End If

If Me.txtFecha > Me.txtFechaF Then
    MsgBox "Rango de Fechas Incorrectas", vbInformation, "AVISO"
    Exit Sub
End If


For i = 0 To Me.chkOption.count - 1
    If chkOption.item(i).value = 1 Then
        Select Case i
            Case 0 ' Archivo de Contribuyentes
                GeneraContribuyentes
            Case 1 'Archivo de Contribuyentes
                GeneraContribuyentesExtranjeros
            Case 2 'Archivo de Movimientos Acumulados por Contribuyentes
                GeneraMovimientosAcumulados
                'GeneraMovimientosAcumuladosNew  'Gitu
            Case 3 ' Archivo de Movimientos de Contra Asiento por Error
                GeneraMovContraAsientoporError
            Case 4 ' Archivo de Exonerados
                GeneraExonerados
            Case 5
                GeneraSustentacionMov
        End Select
    End If
Next i
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub GeneraContribuyentes()
Dim sTipoDoc  As String * 2
Dim snumdoc  As String * 15
Dim sRazonSoc  As String * 40
Dim sApPaterno  As String * 20
Dim sApMaterno As String * 20
Dim sNombres  As String * 20
Dim sP As String * 1
Dim sCodPers As String
Dim sMes As String
Dim lsArc As String
Dim NumeroArchivo As Integer

Dim sLinea As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oITF As COMDCaptaGenerales.DCOMCaptaGenerales
Dim dFecha As Date

dFecha = gdFecSis & " " & GetHoraServer()
sMes = Format$(CDate(txtFechaF.Text), "yyyymm")
J = 0
sP = "|"
NumeroArchivo = FreeFile
lsArc = App.Path & "\Spooler\0695" & sMes & gcEmpresaRUC & ".con"

Set oITF = New COMDCaptaGenerales.DCOMCaptaGenerales
Set rs = oITF.GetITFReporteContribuyentes(CDate(txtFecha.Text), CDate(txtFechaF.Text))

If Not (rs.EOF And rs.BOF) Then
    Open lsArc For Output As #NumeroArchivo
    If LOF(1) > 0 Then
        If MsgBox("Existen Archivos  Anteriores en el Directorio, Desea Remplazarlos ?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
            Close #1
            Kill lsArc
            Open lsArc For Append As #1
        Else
            Exit Sub
        End If
    End If
    oITF.dbCmact.BeginTrans
    While Not rs.EOF
        J = rs.RecordCount
        If Trim(rs!nPersPersoneria) = gPersonaNat Then  ' Tipo de Persona Natural
            Select Case rs!cPersIDTpo
                Case "1": sTipoDoc = "01" 'DNI
                Case "2": sTipoDoc = "06" 'RUC by gitu porque hay clientes como persona natural pero su documento es ruc
                Case "3": sTipoDoc = "02" 'FFPP
                Case "16": sTipoDoc = "03" 'FFAA
                Case "11": sTipoDoc = "07" 'Pasaporte
                Case "10": sTipoDoc = "11" 'Partida de nacimiento
                Case "12": sTipoDoc = "08" 'Documento Provisional de Identidad
                Case "13": sTipoDoc = "08" 'Documento Provisional de Identidad
                Case Else: sTipoDoc = "00" 'Otros
            End Select
            If rs!cPersIDnro = "" Then
                snumdoc = ""
                sTipoDoc = "00"
            Else
                snumdoc = Trim(rs!cPersIDnro)
            End If
        Else
            sTipoDoc = "06" 'RUC
            If rs!cPersIDnro = "" Then
                snumdoc = ""
                sTipoDoc = "00"
            Else
                snumdoc = Trim(rs!cPersIDnro)
            End If
        End If
        sCodPers = Trim(rs!cPersCod)
        sLinea = Trim(sTipoDoc) & sP & Trim(snumdoc) & sP
        Print #NumeroArchivo, sLinea
        
     ''   oITF.AgregaITFControlContribuyentes dFecha, sMes, sCodPers, sTipoDoc, Trim(sNumDoc)
'        sql = "Insert Values  ('" & sTipoDoc & "','" & sNumDoc & "')"
'        .Execute sql
        rs.MoveNext
    Wend
    oITF.dbCmact.CommitTrans
    Set oITF = Nothing
    Close #NumeroArchivo   ' Cierra el archivo.
    nTotal = 0
    nTotal = rs.RecordCount
    MsgBox "Archivo Generado " & nTotal, vbInformation, "AVISO"
    
    rs.Close
    Set rs = Nothing
Else
    Set oITF = Nothing
    MsgBox "No existe Data para la exportacion", vbInformation, "AVISO"
End If
End Sub

Private Sub GeneraExonerados()
Dim sTipoDoc  As String * 2
Dim snumdoc  As String * 15
Dim sCodCta  As String * 20
Dim sTipoDecl  As String * 1
Dim sCodExon As String * 2
Dim sTipoOpe  As String * 1
Dim sFechaSol As String * 8
Dim sHoraSol  As String * 6
Dim sP As String * 1
Dim lsArc As String
Dim NumeroArchivo As Integer
Dim sLinea As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oITF As COMDCaptaGenerales.DCOMCaptaGenerales
Dim dFecha As Date
Dim sMes As String


dFecha = gdFecSis & " " & GetHoraServer()
sMes = Format$(CDate(txtFechaF.Text), "YYYYMM")
J = 0
sP = "|"

'AbreConeccion "07", , , "02"
NumeroArchivo = FreeFile
lsArc = App.Path & "\Spooler\0695" & sMes & gcEmpresaRUC & ".exo"

Set oITF = New COMDCaptaGenerales.DCOMCaptaGenerales
Set rs = oITF.GetITFExonerados(CDate(txtFecha.Text), CDate(txtFechaF.Text))

If Not (rs.EOF And rs.BOF) Then
    J = rs.RecordCount
    Open lsArc For Output As #NumeroArchivo
    If LOF(1) > 0 Then
        If MsgBox("Existen Archivos  Anteriores en el Directorio, Desea Remplazarlos ?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
            Close #1
            Kill lsArc
            Open lsArc For Append As #1
        Else
            Exit Sub
        End If
    End If
    oITF.dbCmact.BeginTrans
    While Not rs.EOF
        If Trim(rs!nPersPersoneria) = gPersonaNat And rs!cPersIDTpo <> 2 Then   ' Tipo de Persona Natural
            Select Case rs!cPersIDTpo
                Case "1": sTipoDoc = "01" 'DNI
                Case "3": sTipoDoc = "02" 'FFPP
                Case "16": sTipoDoc = "03" 'FFAA
                Case "11": sTipoDoc = "07" 'Pasaporte
                Case "10": sTipoDoc = "11" 'Partida de nacimiento
                Case "12": sTipoDoc = "08" 'Documento Provisional de Identidad
                Case "13": sTipoDoc = "08" 'Documento Provisional de Identidad
                
                Case Else: sTipoDoc = "00" 'Otros
                
            End Select
        Else
            sTipoDoc = "06" 'RUC
        End If
'        If rs!cPersIDnro = "" Then
'           sNumDoc = ""
'           sTipoDoc = "00"
'        Else
           snumdoc = Trim(rs!cPersIDnro)
'        End If
        sCodCta = Trim(rs!cCtaCod)
        sTipoDecl = Trim(rs!TipoDecla)
        sCodExon = LCase(Trim(rs!CodOpe) & "")
        sTipoOpe = Trim(rs!TipoOper)
        sFechaSol = Left(rs!cRegistro, 8)
        sHoraSol = Mid(rs!cRegistro, 9, 6)

        sLinea = Trim(sTipoDoc) & sP & Trim(snumdoc) & sP & Trim(sCodCta) & sP & Trim(sTipoDecl) & sP & Trim(sCodExon) & sP & Trim(sTipoOpe) & sP & Trim(sFechaSol) & sP & Trim(sHoraSol) & sP
        Print #NumeroArchivo, sLinea
    '    oITF.AgregaITFControlExonerados dFecha, sMes, rs("cPersCod"), Trim(sTipoDoc), Trim(sNumDoc), Trim(sCodCta), Trim(sTipoDecl), Trim(sCodExon), Trim(sTipoOpe), rs("cRegistro")
'        sql = "Insert DBITF..ArchExonerados values  ('" & sTipoDoc & "','" & sNumDoc & "','" & sCodCta & "','" & sTipoDecl & "','" & sCodExon & "','" & sTipoOpe & "','" & sFechaSol & "','" & sHoraSol & "')"
'        'sql = "Insert DBITF..ArchExoneradosConsol values  ('" & sTipoDoc & "','" & sNumDoc & "','" & sCodCta & "','" & sTipoDecl & "','" & sCodExon & "','" & sTipoOpe & "','" & sFechaSol & "','" & sHoraSol & "','" & Format(txtFechaF.Text, "yyyy/mm/dd") & "')"
'        dbCmactN.Execute sql
        rs.MoveNext
    Wend
    oITF.dbCmact.CommitTrans
    Close #NumeroArchivo   ' Cierra el archivo.
    nTotal = 0
    nTotal = rs.RecordCount
    MsgBox "Archivo Generado " & nTotal, vbInformation, "AVISO"
'    CierraConeccion
Else
    MsgBox "No existe Data para la exportacion", vbInformation, "AVISO"
End If
Set oITF = Nothing
rs.Close
Set rs = Nothing
End Sub
Private Sub GeneraMovContraAsientoporError()
Dim sTipoDoc  As String * 2
Dim snumdoc  As String * 15
Dim sPeriodo  As String * 6
Dim sTipoOpe  As String * 2
Dim sTipoMov  As String * 2
Dim sCodOpe  As String * 2
Dim sModOpe  As String * 1
Dim sMontoBase As String * 15
Dim sMontoImpuesto As String * 15
Dim sP As String * 1
Dim lsArc As String
Dim NumeroArchivo As Integer
Dim sLinea As String
Dim sql As String
Dim rs As New ADODB.Recordset
Dim sMes As String

sMes = Format$(CDate(txtFechaF.Text), "YYYYMM")
J = 0
sP = "|"
NumeroArchivo = FreeFile
lsArc = App.Path & "\Spooler\0695" & sMes & gcEmpresaRUC & ".ext"

 Open lsArc For Output As #NumeroArchivo
    If LOF(1) > 0 Then
        If MsgBox("Existen Archivos  Anteriores en el Directorio, Desea Remplazarlos ?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
            Close #1
            Kill lsArc
            Open lsArc For Append As #1
        Else
            Exit Sub
        End If
    End If
    Close #NumeroArchivo   ' Cierra el archivo.
    'MsgBox J
    MsgBox "Archivo Generado", vbInformation, "AVISO"

End Sub

Private Sub GeneraMovimientosAcumulados()
Dim sTipoDoc  As String * 2
Dim snumdoc  As String * 15
Dim sTipoOpe  As String * 2
Dim sTipoMov  As String * 2
Dim sCodOpe  As String * 2
Dim sMontoBase As String * 15
Dim sMontoImpuesto As String * 15
Dim sP As String * 1
Dim lsArc As String
Dim NumeroArchivo As Integer
Dim sLinea As String
Dim sql As String, sMes As String
Dim rs As ADODB.Recordset
Dim oITF As COMDCaptaGenerales.DCOMCaptaGenerales
Dim oMov As COMDMov.DCOMMov
Dim fechaini As String, fechafin As String
Dim oConConsol As DConecta
Set oConConsol = New DConecta
oConConsol.AbreConexion
Set oITF = New COMDCaptaGenerales.DCOMCaptaGenerales
J = 0
sP = "|"
'sMes = Format$(gdFecSis, "YYYYMM")
sMes = Right(txtFecha.Text, 4) & Mid(txtFecha.Text, 4, 2)
NumeroArchivo = FreeFile
lsArc = App.Path & "\Spooler\0695" & sMes & gcEmpresaRUC & ".mov"

fechaini = Right(txtFecha.Text, 4) & Mid(txtFecha.Text, 4, 2) & Left(txtFecha.Text, 2)

fechafin = Right(txtFechaF.Text, 4) & Mid(txtFechaF.Text, 4, 2) & Left(txtFechaF.Text, 2)

sql = "stp_sel_MovimientosAcumulados '" & fechaini & "', '" & fechafin & "'"
Set rs = oConConsol.CargaRecordSet(sql)
'**************Comentado por ALPA Nr 01
'Set oMov = New COMDMov.DCOMMov
'    Set rs = oMov.ObtenerMovimientosAculuados(fechaini, fechafin)
'Set oMov = Nothing
'**************Fin de Comentario 01
Set rs.ActiveConnection = Nothing

If Not (rs.EOF And rs.BOF) Then
    Open lsArc For Output As #NumeroArchivo
    If LOF(1) > 0 Then
        If MsgBox("Existen Archivos  Anteriores en el Directorio, Desea Remplazarlos ?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
            Close #1
            Kill lsArc
            Open lsArc For Append As #1
        Else
            Exit Sub
        End If
    End If
    PB.Visible = True
    PB.value = 0
    PB.Min = 0
    PB.Max = rs.RecordCount + 10
    Dim TpDoc As String
    While Not rs.EOF
        TpDoc = rs!TpoDoc
        If TpDoc = "00" Then
        TpDoc = "00"
        End If
        
        'sLinea = TpDoc & sP & Trim(rs!Cdoc) & sP & rs!CodPaisE & sP & rs!tipoOpe & sP & rs!TipoMov & sP & Trim(rs!CodOpe) & sP & rs!codPias & sP & rs!MBase & sP & rs!Monto & sP '*** PEAC 20110720
        'sLinea = TpDoc & sP & Trim(rs!Cdoc) & sP & rs!CodPaisE & sP & rs!tipoOpe & sP & rs!TipoMov & sP & Trim(rs!CodOpe) & sP & rs!codPias & sP & rs!MBase & sP & rs!Monto & sP & rs!MontoCalc & sP 'Comentado por LUCV20161117
        sLinea = TpDoc & sP & Trim(rs!Cdoc) & sP & rs!CodPaisE & sP & rs!tipoOpe & sP & rs!TipoMov & sP & Trim(rs!CodOpe) & sP & rs!codPias & sP & rs!MBase & sP & rs!Monto & sP & rs!MontoCalc & sP & rs!nIndicador & sP & rs!TipoDocExtran & sP & rs!NroDocExtran & sP & rs!FechaNac & sP & rs!TipoDirecLegal & sP & rs!DirecLegal & sP
              
        Print #NumeroArchivo, sLinea
        'INSERTAR
            'Set oMov = New COMDMov.DCOMMov
            'oMov.InsertarArchMov sMes, IIf(IsNull(rs!CodPers), "", rs!CodPers), TpDoc, rs!Cdoc, rs!afec, rs!CTA, rs!CodOpe, rs!Monto
            'Set oMov = Nothing '***Comentado by NAGL 20190321
            
        Me.Caption = Str(PB.value)
        DoEvents
        rs.MoveNext
        PB.value = PB.value + 1
    Wend
    PB.Visible = False
    Close #NumeroArchivo   ' Cierra el archivo.
    nTotal = 0
    nTotal = rs.RecordCount
    MsgBox "Archivo Generado " & nTotal, vbInformation, "AVISO"
Else
    MsgBox "No existe Data para la exportacion", vbInformation, "AVISO"
End If
End Sub


Private Sub GeneraContribuyentesExtranjeros()
Dim sTipoDoc  As String * 2
Dim snumdoc  As String * 15
Dim sApPaterno  As String * 20
Dim sApMaterno As String * 20
Dim sNombres  As String * 20
Dim sCodPais As String * 4
Dim sP As String * 1
Dim lsArc As String, sPersCod As String
Dim NumeroArchivo As Integer
Dim sLinea As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oITF As COMDCaptaGenerales.DCOMCaptaGenerales
Dim sMes As String
Dim dFecha As Date

dFecha = gdFecSis & " " & GetHoraServer()
sMes = Format$(CDate(txtFechaF.Text), "yyyymm")


J = 0
sP = "|"

'AbreConeccion "07", , , "02"
NumeroArchivo = FreeFile
lsArc = App.Path & "\Spooler\0695" & sMes & gcEmpresaRUC & ".cex"

' sql = "Select distinct T.cCodPers, T.cTipPers, T.cTidoci , T.cNudoci, T.cTidotr, T.cNudotr from"
' sql = sql & " ( Select"
' sql = sql & " cCodPers, cTipPers,"
' sql = sql & " case when cTidoci in ('2') and (cNudoci <> null Or cNudoci<>'' ) then"
' sql = sql & " cNomPers else '' end NomPers,"
' sql = sql & " cTidoci , cNudoci, cTidotr, cNudotr"
' sql = sql & " from [128.107.2.3].DBPersona.dbo.Persona where cTipPers = '1' and CTidoci = '2' and cCodPers in"
' sql = sql & " (Select distinct cCodPers from DBItf.dbo.ITFConsolAsiento  ITFCA"
' sql = sql & " Inner Join DBItf.dbo.PersCuentaITFConsol PersITF on PersITF.cCodCta = ITFCA.cCodCta"
' sql = sql & " where datediff(month, ITFCA.dFecTran, '" & Format(Me.txtFechaF, "yyyy/mm/dd") & "') = 0) "
' sql = sql & " Union"
' sql = sql & " Select"
' sql = sql & " cCodPers, cTipPers,"
' sql = sql & " case when cTidoci in ('2') and (cNudoci <> null Or cNudoci<>'' ) then"
' sql = sql & " cNomPers else '' end NomPers,"
' sql = sql & " cTidoci , cNudoci, cTidotr, cNudotr"
' sql = sql & " from [128.107.2.3].DBPersona.dbo.Persona where cTipPers <> '1' and CTidoci = '2' and cCodPers in"
' sql = sql & " (Select distinct cCodPers from DBItf.dbo.ITFConsolAsiento  ITFCA"
' sql = sql & " Inner Join DBItf.dbo.PersCuentaITFConsol PersITF on PersITF.cCodCta = ITFCA.cCodCta"
' sql = sql & " where datediff(month, ITFCA.dFecTran, '" & Format(Me.txtFechaF, "yyyy/mm/dd") & "') = 0)) T "
'
'sCodPais = "9239"
'Rs.Open sql, dbCmactN, adOpenStatic, adLockReadOnly, adCmdText

Set oITF = New COMDCaptaGenerales.DCOMCaptaGenerales
Set rs = oITF.GetITFReporteContribuyentes(CDate(txtFecha.Text), CDate(txtFechaF.Text), False)
sCodPais = "9239"
If Not (rs.EOF And rs.BOF) Then
    Open lsArc For Output As #NumeroArchivo
    If LOF(1) > 0 Then
        If MsgBox("Existen Archivos  Anteriores en el Directorio, Desea Remplazarlos ?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
            Close #1
            Kill lsArc
            Open lsArc For Append As #1
        Else
            Exit Sub
        End If
    End If
    oITF.dbCmact.BeginTrans
    While Not rs.EOF
        J = rs.RecordCount
        If Trim(rs!nPersPersoneria) = gPersonaNat Then  ' Tipo de Persona Natural
            sTipoDoc = "04"
            If IsNull(Trim(rs!cPersIDnro)) Then
                snumdoc = ""
                sTipoDoc = "00"
            Else
                snumdoc = Trim(rs!cPersIDnro)
            End If
        Else
            sTipoDoc = "06" 'RUC
            If IsNull(Trim(rs!cPersIDnro)) Then
                snumdoc = ""
                sTipoDoc = "00"
            Else
                snumdoc = Trim(rs!cPersIDnro)
            End If
        End If
        sLinea = Trim(sTipoDoc) & sP & Trim(snumdoc) & sP & Trim(sCodPais) & sP
        sPersCod = rs("cPersCod")
        Print #NumeroArchivo, sLinea
   ''     oITF.AgregaITFControlContribuyentes dFecha, sMes, sPersCod, Trim(sTipoDoc), Trim(sNumDoc)
'        sql = "Insert dbitf..ArchContExtranj values ('" & sTipoDoc & "','" & sNumDoc & "','" & sCodPais & "')"
'        sql = "Insert dbitf..ArchContExtranjConsol values ('" & sTipoDoc & "','" & sNumDoc & "','" & sCodPais & "','" & Format(txtFechaF.Text, "yyyy/mm/dd") & "')"
'        dbCmactN.Execute sql
        rs.MoveNext
    Wend
    oITF.dbCmact.CommitTrans
    Close #NumeroArchivo   ' Cierra el archivo.
    nTotal = 0
    nTotal = rs.RecordCount
    MsgBox "Archivo Generado " & nTotal, vbInformation, "AVISO"
'    CierraConeccion
Else
    MsgBox "No existe Data para la exportacion", vbInformation, "AVISO"
End If
Set oITF = Nothing
rs.Close
Set rs = Nothing
End Sub

Private Sub GeneraSustentacionMov()
Dim sTipoDoc  As String * 2
Dim snumdoc  As String * 15
Dim sTipoOpe  As String * 2
Dim sTipoMov  As String * 2
Dim sCodOpe  As String * 2
Dim sMontoBase As String * 15
Dim sMontoImpuesto As String * 15
Dim sP As String * 1
Dim lsArc As String
Dim NumeroArchivo As Integer
Dim sLinea As String
Dim rs As ADODB.Recordset
Dim oITF As COMDCaptaGenerales.DCOMCaptaGenerales
Dim fechaini As String, fechafin As String

        
Set oITF = New COMDCaptaGenerales.DCOMCaptaGenerales
Set rs = oITF.GetITFReporteSustentacionMov(CDate(txtFecha.Text), CDate(txtFechaF.Text))

If Not (rs.EOF And rs.BOF) Then


Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim nFila As Long, i As Long
Dim lsArchivoN As String, lbLibroOpen As Boolean
Dim Total As Double
Dim sRep As String

lsArchivoN = App.Path & "\Spooler\Rep" & "SUSTMOV" & Format(gdFecSis & " " & Time, "yyyymmddhhmmss") & gsCodUser & ".xls"
OleExcel.Class = "ExcelWorkSheet"
lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
If lbLibroOpen Then
  Screen.MousePointer = vbHourglass

        Set xlHoja1 = xlLibro.Worksheets(1)
        ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
    
        nFila = 1
            
        xlHoja1.Cells(nFila, 1) = gsNomCmac
        nFila = 2
        xlHoja1.Cells(nFila, 1) = gsNomAge
        xlHoja1.Range("F2:H2").MergeCells = True
        xlHoja1.Cells(nFila, 6) = Format(gdFecSis, "Long Date")
    
        nFila = 3
        xlHoja1.Cells(nFila, 1) = "REPORTE DE SUSTENTACION DE MOVIMIENTOS ACUMULADOS POR AGENCIA Y OPERACION DEL" & txtFecha.Text & " AL " & txtFechaF.Text
        
        xlHoja1.Range("A1:M5").Font.Bold = True
            
        xlHoja1.Range("A3:M3").MergeCells = True
        xlHoja1.Range("A3:A3").HorizontalAlignment = xlCenter
        xlHoja1.Range("A5:M5").HorizontalAlignment = xlCenter
    
        nFila = 5
            
            xlHoja1.Cells(nFila, 1) = "FECHA"
            xlHoja1.Cells(nFila, 2) = "CODAGE"
            xlHoja1.Cells(nFila, 3) = "CODOPE"
            xlHoja1.Cells(nFila, 4) = "OPERACION"
            xlHoja1.Cells(nFila, 5) = "CODCONCEPTO"
            xlHoja1.Cells(nFila, 6) = "CONCEPTO"
            xlHoja1.Cells(nFila, 7) = "MONTO"
       
        Total = 0
        
    
    While Not rs.EOF
     nFila = nFila + 1
       
                xlHoja1.Cells(nFila, 1) = rs!cfecha
                xlHoja1.Cells(nFila, 2) = rs!cCodAge
                xlHoja1.Cells(nFila, 3) = rs!cOpeCod
                xlHoja1.Cells(nFila, 4) = rs!cOpedesc
                xlHoja1.Cells(nFila, 5) = rs!nPrdConceptoCod
                xlHoja1.Cells(nFila, 6) = rs!cDescripcion
                xlHoja1.Cells(nFila, 7) = rs!NMONTDET
                Total = Total + rs!NMONTDET
        rs.MoveNext
        
    Wend
    
        nFila = nFila + 1
        xlHoja1.Range("A" & nFila & ":G" & nFila).Font.Bold = True
        xlHoja1.Range("A" & nFila & ":F" & nFila).MergeCells = True
        xlHoja1.Cells(nFila, 7) = Total
    
        xlHoja1.Cells.Select
        xlHoja1.Cells.Font.Name = "Arial"
        xlHoja1.Cells.Font.Size = 9
        xlHoja1.Cells.EntireColumn.AutoFit
    
    'Cierro...
            OleExcel.Class = "ExcelWorkSheet"
            ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
            OleExcel.SourceDoc = lsArchivoN
            OleExcel.Verb = 1
            OleExcel.Action = 1
            OleExcel.DoVerb -1
        
    
End If
    
    nTotal = 0
    nTotal = rs.RecordCount
    MsgBox "Archivo Generado " & nTotal, vbInformation, "AVISO"
'    CierraConeccion
Else
    MsgBox "No existe Data para la exportacion", vbInformation, "AVISO"
End If

Screen.MousePointer = vbDefault

Set oITF = Nothing
'rs.Close
Set rs = Nothing




End Sub


Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtFechaF.SetFocus
End If
End Sub

Private Sub txtFechaF_GotFocus()
fEnfoque txtFechaF
End Sub

Private Sub Form_Load()
Me.Caption = "Reportes ITF SUNAT"
End Sub

Private Sub txtFechaF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    chkOption(0).SetFocus
End If
End Sub

'***********************************************************
' Inicia Trabajo con EXCEL, crea variable Aplicacion y Libro
'***********************************************************
Private Function ExcelBegin(psArchivo As String, _
        xlAplicacion As Excel.Application, _
        xlLibro As Excel.Workbook, Optional pbBorraExiste As Boolean = True) As Boolean
        
Dim fs As New Scripting.FileSystemObject
On Error GoTo ErrBegin
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application

If fs.FileExists(psArchivo) Then
   If pbBorraExiste Then
      fs.DeleteFile psArchivo, True
      Set xlLibro = xlAplicacion.Workbooks.Add
   Else
      Set xlLibro = xlAplicacion.Workbooks.Open(psArchivo)
   End If
Else
   Set xlLibro = xlAplicacion.Workbooks.Add
End If
ExcelBegin = True
Exit Function
ErrBegin:
  MsgBox err.Description, vbInformation, "Aviso"
  ExcelBegin = False
End Function

'***********************************************************
' Final de Trabajo con EXCEL, graba Libro
'***********************************************************
Private Sub ExcelEnd(psArchivo As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional plSave As Boolean = True)
On Error GoTo ErrEnd
   If plSave Then
        xlHoja1.SaveAs psArchivo
   End If
   xlLibro.Close
   xlAplicacion.Quit
   Set xlAplicacion = Nothing
   Set xlLibro = Nothing
   Set xlHoja1 = Nothing
Exit Sub
ErrEnd:
   MsgBox err.Description, vbInformation, "Aviso"
End Sub


'********************************
' Adiciona Hoja a LibroExcel
'********************************
Private Sub ExcelAddHoja(psHojName As String, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet)
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = psHojName Then
       xlHoja1.Delete
       Exit For
    End If
Next
Set xlHoja1 = xlLibro.Worksheets.Add
xlHoja1.Name = psHojName
End Sub

Private Sub GeneraMovimientosAcumuladosNew()
Dim sTipoDoc  As String * 2
Dim snumdoc  As String * 15
Dim sTipoOpe  As String * 2
Dim sTipoMov  As String * 2
Dim sCodOpe  As String * 2
Dim sMontoBase As String ' * 15
Dim sMontoImpuesto As String * 15
'----------------------------------
Dim sTipoDoc_1  As String * 2
Dim snumdoc_1  As String * 15
Dim sTipoOpe_1  As String * 2
Dim sTipoMov_1  As String * 2
Dim sCodOpe_1  As String * 2
Dim sMontoBase_1 As String '* 15
Dim sMontoImpuesto_1 As String * 15
Dim sNumDocAnt As String * 15
Dim lsArcLogITF  As String
Dim nCorrelativo As Integer
Dim lsCodPaisCont As String ' * 4
Dim lsCodPaisOpe As String '* 4
'----------------------------------
Dim sP As String * 1
Dim lsArc As String
Dim NumeroArchivo As Integer
Dim sLinea As String
Dim sql As String
Dim oITF As COMDCaptaGenerales.DCOMCaptaGenerales
Dim oMov As COMDMov.DCOMMov
Dim rs As ADODB.Recordset
Dim V As Integer
Dim sMes As String
Dim fechaini As String, fechafin As String

Set oITF = New COMDCaptaGenerales.DCOMCaptaGenerales
J = 0
sP = "|"

'sMes = Format$(gdFecSis, "YYYYMM")
sMes = Right(txtFecha.Text, 4) & Mid(txtFecha.Text, 4, 2)
NumeroArchivo = FreeFile
lsArc = App.Path & "\Spooler\0695" & sMes & gcEmpresaRUC & ".mov"
lsArcLogITF = App.Path & "\Spooler\0695" & Format(Me.txtFechaF, "YYYYMM") & gsRUCCmac & ".txt"

fechaini = Right(txtFecha.Text, 4) & Mid(txtFecha.Text, 4, 2) & Left(txtFecha.Text, 2)

fechafin = Right(txtFechaF.Text, 4) & Mid(txtFechaF.Text, 4, 2) & Left(txtFechaF.Text, 2)

Set oMov = New COMDMov.DCOMMov
    Set rs = oMov.ObtenerMovimientosAculuados(fechaini, fechafin)
Set oMov = Nothing


Set rs = New ADODB.Recordset

'--------------------- Validar Datos----------------------------------------
    rs.MoveFirst
'----------------------------------------------------------------------------------------------
If Not (rs.EOF And rs.BOF) Then
    Open lsArc For Output As #NumeroArchivo
    If LOF(NumeroArchivo) > 0 Then
        If MsgBox("Existen Archivos  Anteriores en el Directorio, Desea Remplazarlos ?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
            Close #NumeroArchivo
            Kill lsArc
            Open lsArc For Append As #NumeroArchivo '#1
        Else
            Exit Sub
        End If
    End If
    PB.Visible = True
    PB.value = 0
    PB.Min = 0
    PB.Max = rs.RecordCount + 10
    
    sNumDocAnt = ""
    While Not rs.EOF
    
        sTipoDoc = Trim(rs("sTipoDoc"))
        sTipoDoc_1 = Trim(rs("sTipoDoc"))
        snumdoc = Trim(rs("snumdoc"))
        snumdoc_1 = Trim(rs("snumdoc"))
        
        sTipoOpe = Trim(rs("TipoOpe"))
        sTipoOpe_1 = Trim(rs("TipoOpe"))
        sTipoMov = Trim(rs("TipoMov"))
        sTipoMov_1 = Trim(rs("TipoMov"))
        sCodOpe = Trim(rs("CodOpe"))
        sCodOpe_1 = Trim(rs("CodOpe"))
        lsCodPaisCont = Trim(rs("CodPaisCont"))
        lsCodPaisOpe = Trim(rs("CodPaisOpe"))
        
        
        Select Case sCodOpe
         Case "12", "13", "14", "15"
             sMontoBase = ""
             sMontoBase_1 = ""
         Case Else
             sMontoBase = Trim(Format(Abs(rs("nMonto")), "###########0.00"))
             sMontoBase_1 = Trim(Format(Abs(rs("nMonto")), "###########0.00"))
        End Select
                
        
        sMontoImpuesto = Trim(Format(Abs(rs("nImpuesto")), "###########0.00"))
        sMontoImpuesto_1 = Trim(Format(Abs(rs("nImpuesto")), "###########0.00"))
        
        'AAI----------------------------------------------------------
           
           
          If (sTipoDoc = "00" And Trim(snumdoc) = "") Then
            nCorrelativo = nCorrelativo + 1
            snumdoc_1 = "0000000" & CStr(nCorrelativo)
            sTipoDoc_1 = "01"
           
         End If
         
        ' Cuando la Tipo de Operacion es Exonerada(02) pero Genera ITF
        If (sTipoDoc = "00" And Trim(snumdoc) = "" And Trim(sTipoOpe) = "02") Then
           nCorrelativo = nCorrelativo + 1
           snumdoc_1 = "0000000" & CStr(nCorrelativo)
           sTipoOpe_1 = "01"
           sTipoDoc_1 = "01"
        End If

        
         If (sTipoDoc = "06" And Len(Trim(snumdoc)) < 6) Then
            'nCorrelativo = nCorrelativo + 1
            'snumdoc = "00" & CStr(nCorrelativo)
            sTipoDoc_1 = "01"
         End If

          If (sTipoDoc = "06" And Trim(snumdoc) = "00000000") Then
            'nCorrelativo = nCorrelativo + 1
            'snumdoc = "00" & CStr(nCorrelativo)
            sTipoDoc_1 = "01"
          End If
          
         ' AAI Tipo el Caso Operacion Invalida
         ' que es cuando Un RUC aparece Dos veces
         ' en este caso se elimina el segundo Item porq no genera ITF
         If Trim(snumdoc) = Trim(sNumDocAnt) And Trim(sMontoBase) <> "" And Trim(sTipoOpe) = "02" And Trim(sMontoImpuesto) = "0.00" Then
           ' Cuando es igual NO Imprime en el Archivo Texto
           'Stop
         Else
             
             sLinea = Trim(sTipoDoc_1) & sP & Trim(snumdoc_1) & sP & lsCodPaisCont & sP & Trim(sTipoOpe_1) & sP & Trim(sTipoMov_1) & sP & Trim(sCodOpe_1) & sP & lsCodPaisOpe & sP & Trim(sMontoBase_1) & sP & Trim(sMontoImpuesto_1) & sP
             Print #NumeroArchivo, sLinea
             sNumDocAnt = snumdoc
            
         End If
        '---- El execute irá fuera si se insertan todos los registros los que generan y no ITF----
            'SQL = "Insert dbitf..ArchMovimientos values ('" & sTipoDoc & "','" & snumdoc & "','" & Rs("CodPaisCont") & "','" & sTipoOpe & "','" & sTipoMov & "','" & sCodOpe & "','" & lsCodPaisOpe & "'," & IIf(Trim(sMontoBase) = "", 0, sMontoBase) & "," & sMontoImpuesto & ")"
            sql = "Insert dbitf..ArchMovimientos values ('" & sTipoDoc_1 & "','" & snumdoc_1 & "','" & lsCodPaisCont & "','" & sTipoOpe_1 & "','" & sTipoMov_1 & "','" & sCodOpe_1 & "','" & lsCodPaisOpe & "'," & IIf(Trim(sMontoBase_1) = "", 0, sMontoBase_1) & "," & sMontoImpuesto_1 & ")"
            
            Sleep 20
            'dbCmactN.Execute Sql
        '------------------------------------------------------------------------------------------
        Me.Caption = Str(PB.value)
        rs.MoveNext
        PB.value = PB.value + 1
    Wend
    PB.Visible = False
    Close #NumeroArchivo   ' Cierra el archivo.
    nTotal = 0
    nTotal = rs.RecordCount
 '  MsgBox "Archivo Generado " & nTotal, vbInformation, "AVISO"
 '   CierraConeccion
Else
    MsgBox "No existe Data para la exportacion", vbInformation, "AVISO"
End If
End Sub

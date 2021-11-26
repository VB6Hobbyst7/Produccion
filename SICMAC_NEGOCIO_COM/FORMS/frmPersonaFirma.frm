VERSION 5.00
Object = "{160AE063-3670-11D5-8214-000103686C75}#6.0#0"; "PryOcxExplorer.ocx"
Begin VB.Form frmPersonaFirma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Firmas por Persona"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   Icon            =   "frmPersonaFirma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin SICMACT.ImageDB IDBFirma 
      Height          =   5790
      Left            =   75
      TabIndex        =   2
      Top             =   75
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   10213
      Enabled         =   0   'False
   End
   Begin VB.CommandButton CmdActFirma 
      Caption         =   "&Actualizar Firma"
      Height          =   375
      Left            =   7125
      TabIndex        =   0
      Top             =   5925
      Width           =   1380
   End
   Begin PryOcxExplorer.OcxCdlgExplorer CdlgImg 
      Height          =   480
      Left            =   75
      TabIndex        =   1
      Top             =   5850
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   847
      Filtro          =   "Archivos bmp|*.bmp|Archivos jpg|*.jpg|Archivos gif|*.gif|Todos los Archivos|*.*"
      Altura          =   280
   End
End
Attribute VB_Name = "frmPersonaFirma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsFirma As ADODB.Recordset
Dim sCadena As String
Dim oConFirma As COMConecta.DCOMConecta
Dim sPersCod As String
'FRHU 20141115 MEMO-2766-2014
Dim objPista As COMManejador.Pista
Dim sCtaCod As String
'FIN FRHU 20141115

Public Function Inicio(ByVal psPersCod As String, _
                        ByVal psCodAge As String, _
                        Optional ByVal pbPermiteActualizar As Boolean = True, _
                        Optional ByVal pbPermitePresentar As Boolean = False, _
                        Optional ByVal pnSoloPersMant As Integer = 0, _
                        Optional ByVal psCtaCod As String = "") 'FRHU 20141115 MEMO-2766-2014: Se agrego psCtaCod

    Dim rs As ADODB.Recordset
    Dim ssql As String
    Dim oIni As COMConecta.DCOMClasIni

    Dim lsProvider As String
    Dim lsUser As String
    Dim lsPassword As String
    Dim lsArchivo As String
    Dim lsTabla As String
    Dim ClsPersona As COMDPersona.DCOMPersonas ' ADD BY JATO 20210215
    Dim bPersona As Boolean ' ADD BY JATO 20210215

    sPersCod = psPersCod
    sCtaCod = psCtaCod 'FRHU 20141115 MEMO-2766-2014
    
    Set oConFirma = New COMConecta.DCOMConecta
    Set oIni = New COMConecta.DCOMClasIni


    '
    'lsArchivo = App.path & "\SICMACT.INI"
    'lsProvider = oIni.LeerArchivoIni(oIni.Encripta("SICMACT"), oIni.Encripta("Provider"), lsArchivo)
    'lsUser = oIni.LeerArchivoIni(oIni.Encripta("SICMACT"), oIni.Encripta("User"), lsArchivo)
    'lsPassword = oIni.LeerArchivoIni(oIni.Encripta("SICMACT"), oIni.Encripta("Password"), lsArchivo)

    'ARCV 30-10-2006
    'sSql = "SELECT cServidor, cBaseDatos, cTabla FROM RutaImagenes WHERE cAgeCod='" & psCodAge & "'"
    'Debe cargar la ruta de Agencia actual...no la de la que pertenece la Persona
    ssql = "SELECT cServidor, cBaseDatos, cTabla FROM RutaImagenes WHERE cAgeCod='" & gsCodAge & "'"
    '-------------------------

    Call oConFirma.AbreConexion
    Set rs = oConFirma.CargaRecordSet(ssql)

    If Not rs.EOF Then
        Dim nPos As Integer
        Dim i As Integer
        Dim sUsuSql As String
        Dim sPassSql As String
    
        'sCadena = Replace(sCadena, "dbcmac", Trim(rs!cBaseDatos))
        'sCadena = Replace(sCadena, "Virtualsql\sqlnodo01", Trim(rs!cServidor))
        'sCadena = "PROVIDER=SQLOLEDB.1;User ID=sa;Password=migrasa;INITIAL CATALOG=" & Trim(rs!cBaseDatos) & ";DATA SOURCE=" & Trim(rs!cServidor) & ""
        sCadena = UCase(oConFirma.CadenaConexion)
    
        'Obtiene usuario
        nPos = InStr(1, sCadena, "USER ID")
        i = nPos
        sUsuSql = ""
        Do While nPos > 0
            If Mid(sCadena, i, 1) = "=" Then
                i = i + 1
                Do While Mid(sCadena, i, 1) <> ";"
                    sUsuSql = sUsuSql & Mid(sCadena, i, 1)
                    i = i + 1
                Loop
                Exit Do
            End If
            i = i + 1
        Loop
    
        'Obtiene Password
        nPos = InStr(1, sCadena, "PASSWORD")
        i = nPos
        sPassSql = ""
        Do While nPos > 0
            If Mid(sCadena, i, 1) = "=" Then
                i = i + 1
                Do While Mid(sCadena, i, 1) <> ";"
                    sPassSql = sPassSql & Mid(sCadena, i, 1)
                    i = i + 1
                Loop
                Exit Do
            End If
            i = i + 1
        Loop
    
        sCadena = "PROVIDER=SQLOLEDB.1;User ID=" & sUsuSql & ";Password=" & LCase(sPassSql) & ";INITIAL CATALOG=" & Trim(rs!cBaseDatos) & ";DATA SOURCE=" & Trim(rs!cServidor) & ""
    
        lsTabla = rs!cTabla
    Else
        sCadena = oConFirma.CadenaConexion
        lsTabla = "PersImagen"
    End If
    rs.Close
    Set oIni = Nothing
    oConFirma.CierraConexion


    ssql = "Select cPersCod,iPersFirma,cUltimaActualizacion from " & lsTabla & " Where cPersCod = '" & sPersCod & "'"
    Set rsFirma = New ADODB.Recordset
    'Set oConFirma = New COMConecta.DCOMConecta

    Dim o As ADODB.Connection
    Set o = New ADODB.Connection
    o.CursorLocation = adUseClient
    o.Open sCadena

    'Set oConFirma.ConexionActiva = O


    'Call oConFirma.AbreConexion(sCadena)


    'rsFirma.Open sSql, oConFirma.ConexionActiva, adOpenKeyset, adLockOptimistic
    rsFirma.Open ssql, o, adOpenKeyset, adLockOptimistic

    '************ firma madm 20090928

    Set ClsPersona = New COMDPersona.DCOMPersonas 'ADD BY JATO 20210215
    bPersona = ClsPersona.GetPuedeGenerar(gsCodUser) 'ADD BY JATO 20210215

    If rsFirma.RecordCount > 0 And pbPermiteActualizar And pbPermitePresentar Then
        Call IDBFirma.CargarFirma(rsFirma)
        'If (gsCodCargo = "006024" Or gsCodCargo = "007012") Then 'COM BY JATO 20210215
        If bPersona = True Then 'ADD BY JATO 20210215
            CmdActFirma.Enabled = False
        Else
            CmdActFirma.Enabled = True
        End If
        Me.Show 1
    ElseIf rsFirma.RecordCount > 0 And pbPermiteActualizar = False And pbPermitePresentar = False Then
    
    ElseIf rsFirma.RecordCount > 0 And pbPermiteActualizar = False And pbPermitePresentar = True Then
        'ADD BY JATO 20210315
        Call IDBFirma.CargarFirma(rsFirma)
        CmdActFirma.Enabled = pbPermiteActualizar
        Me.Show 1
    ElseIf rsFirma.RecordCount > 0 And pbPermiteActualizar = False And pbPermitePresentar = True And pnSoloPersMant = 1 Then
        Call IDBFirma.CargarFirma(rsFirma)
        CmdActFirma.Enabled = pbPermiteActualizar
        Me.Show 1
    ElseIf rsFirma.RecordCount = 0 And pbPermiteActualizar = False And pbPermitePresentar = False Then
        'Call IDBFirma.CargarFirma(rsFirma)
'        CmdActFirma.Enabled = pbPermiteActualizar
'        CmdActFirma.Enabled = True
'        Me.Show 1
    ElseIf rsFirma.RecordCount = 0 And pbPermiteActualizar = False And pbPermitePresentar = True Then
        Call IDBFirma.CargarFirma(rsFirma)
'        CmdActFirma.Enabled = pbPermiteActualizar
'        CmdActFirma.Enabled = True
'        Me.Show 1
    Else
        Call IDBFirma.CargarFirma(rsFirma)
        'Add By GITU 2011-04-26
        ' If (gsCodCargo = "006024" Or gsCodCargo = "007012") And rsFirma.RecordCount > 0 Then ' ADD BY JATO 20210215
        If bPersona = True And rsFirma.RecordCount > 0 Then ' ADD BY JATO 20210215
            CmdActFirma.Enabled = False
        Else
            CmdActFirma.Enabled = True
        End If
        'CmdActFirma.Enabled = True
        'End GITU
        Me.Show 1
   End If
    '******************

'    Call IDBFirma.CargarFirma(rsFirma)
''
'    CmdActFirma.Enabled = pbPermiteActualizar
''    Me.Show 1
    End Function


'Public Function Inicio(ByVal psPersCod As String, _
'                        ByVal psCodAge As String, _
'                        Optional ByVal pbPermiteActualizar As Boolean = True)
'
'Dim rs As ADODB.Recordset
'Dim sSql As String
'Dim oIni As COMConecta.DCOMClasIni
'
'Dim lsProvider As String
'Dim lsUser As String
'Dim lsPassword As String
'Dim lsArchivo As String
'Dim lsTabla As String
'
'sPersCod = psPersCod
'
'Set oConFirma = New COMConecta.DCOMConecta
'Set oIni = New COMConecta.DCOMClasIni
'
'
''
''lsArchivo = App.path & "\SICMACT.INI"
''lsProvider = oIni.LeerArchivoIni(oIni.Encripta("SICMACT"), oIni.Encripta("Provider"), lsArchivo)
''lsUser = oIni.LeerArchivoIni(oIni.Encripta("SICMACT"), oIni.Encripta("User"), lsArchivo)
''lsPassword = oIni.LeerArchivoIni(oIni.Encripta("SICMACT"), oIni.Encripta("Password"), lsArchivo)
'
''ARCV 30-10-2006
''sSql = "SELECT cServidor, cBaseDatos, cTabla FROM RutaImagenes WHERE cAgeCod='" & psCodAge & "'"
''Debe cargar la ruta de Agencia actual...no la de la que pertenece la Persona
'sSql = "SELECT cServidor, cBaseDatos, cTabla FROM RutaImagenes WHERE cAgeCod='" & gsCodAge & "'"
''-------------------------
'
'Call oConFirma.AbreConexion
'Set rs = oConFirma.CargaRecordSet(sSql)
'
'
'If Not rs.EOF Then
'    'sCadena = Replace(sCadena, "dbcmac", Trim(rs!cBaseDatos))
'    'sCadena = Replace(sCadena, "Virtualsql\sqlnodo01", Trim(rs!cServidor))
'
'    'sCadena = "PROVIDER=SQLOLEDB.1;User ID=sa;Password=migrasa;INITIAL CATALOG=" & Trim(rs!cBaseDatos) & ";DATA SOURCE=" & Trim(rs!cServidor) & ""
'    sCadena = "PROVIDER=SQLOLEDB.1;User ID=sa;Password=paraque;INITIAL CATALOG=" & Trim(rs!cBaseDatos) & ";DATA SOURCE=" & Trim(rs!cServidor) & "" 'MAYNAS
'
'    lsTabla = rs!cTabla
'Else
'    sCadena = oConFirma.CadenaConexion
'    lsTabla = "PersImagen"
'End If
'Set oIni = Nothing
'oConFirma.CierraConexion
'
'
'sSql = "Select cPersCod,iPersFirma,cUltimaActualizacion from " & lsTabla & " Where cPersCod = '" & sPersCod & "'"
'Set rsFirma = New ADODB.Recordset
''Set oConFirma = New COMConecta.DCOMConecta
'
'Dim O As ADODB.Connection
'Set O = New ADODB.Connection
'    O.CursorLocation = adUseClient
'  O.Open sCadena
'
' 'Set oConFirma.ConexionActiva = O
'
'
''Call oConFirma.AbreConexion(sCadena)
'
'
''rsFirma.Open sSql, oConFirma.ConexionActiva, adOpenKeyset, adLockOptimistic
'rsFirma.Open sSql, O, adOpenKeyset, adLockOptimistic
'
'
'Call IDBFirma.CargarFirma(rsFirma)
'
'
'CmdActFirma.Enabled = pbPermiteActualizar
'
'Me.Show 1
'End Function

Private Sub CmdActFirma_Click()
    Dim sRuta As String
    Dim rs As ADODB.Recordset

    CdlgImg.nHwd = Me.hwnd
    CdlgImg.Show
    sRuta = CdlgImg.Ruta
    If Len(Trim(sRuta)) > 0 Then
        IDBFirma.RutaImagen = sRuta
        Call IDBFirma.GrabarFirma(rsFirma, sPersCod, "")
        gbFirmaActualizada = True 'ande 20171011 hay actualizacion de firma
        rsFirma.Update  'MADM 20090928 Graba en Persona
        'FRHU 20141115 MEMO-2766-2014
        If sCtaCod <> "" Then 'Por el momento solo se inserta la pista cuando hacen la operacion de cancelar una cuenta de ahorro
            Set objPista = New COMManejador.Pista
            objPista.InsertarPista gMantenimientoFirma, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gConsultar, "Modifica Firma", sCtaCod, gCodigoCuenta
            Set objPista = Nothing
        End If
        'FIN FRHU
    End If
    
    oConFirma.CierraConexion
    Unload Me
End Sub

Private Sub Form_Load()
    'FRHU 20141115 MEMO-2766-2014
    If sCtaCod <> "" Then 'Por el momento solo se inserta la pista cuando quieren hacer la operacion de cancelar una cuenta de ahorro
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gConsultarFirma, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gConsultar, "Consulta Firma", sCtaCod, gCodigoCuenta
        Set objPista = Nothing
    End If
    'FIN FRHU
    Call CentraForm(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsFirma = Nothing
    Set oConFirma = Nothing
End Sub

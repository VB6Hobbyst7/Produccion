VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogCalculaSaldos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre Mensual - Operaciones de Almacén"
   ClientHeight    =   2715
   ClientLeft      =   2985
   ClientTop       =   3735
   ClientWidth     =   5160
   Icon            =   "frmLogCalculaSaldos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCalcula 
      Height          =   1245
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4935
      Begin VB.ComboBox cboTpoAlm 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   345
         Width           =   3345
      End
      Begin Sicmact.TxtBuscar txtAlmacen 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         sTitulo         =   ""
      End
      Begin VB.Label lblEtiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Cierre para"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   420
         Width           =   945
      End
      Begin VB.Label lblAlmacen 
         AutoSize        =   -1  'True
         Caption         =   "Almacen"
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
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   780
         Width           =   735
      End
      Begin VB.Label lblAlmacenG 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2160
         TabIndex        =   7
         Top             =   720
         Width           =   2490
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3900
      TabIndex        =   5
      Top             =   2220
      Width           =   1155
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   2220
      Width           =   1155
   End
   Begin VB.Frame fraMes 
      Height          =   795
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   4935
      Begin VB.TextBox txtAnio 
         Height          =   315
         Left            =   3840
         MaxLength       =   4
         TabIndex        =   3
         Top             =   300
         Width           =   795
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   2535
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Mes y Año"
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
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.Frame fraBarra 
      Height          =   840
      Left            =   120
      TabIndex        =   12
      Top             =   1280
      Visible         =   0   'False
      Width           =   4935
      Begin MSComctlLib.ProgressBar prgBar 
         Height          =   315
         Left            =   300
         TabIndex        =   13
         Top             =   420
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Procesando"
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
         Left            =   300
         TabIndex        =   14
         Top             =   210
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmLogCalculaSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsAlmacen As ADODB.Recordset
Dim rsProductos As ADODB.Recordset
Dim lsTitulo As Boolean
Dim lsCadena As String
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtAnio.Visible And txtAnio.Enabled Then txtAnio.SetFocus
    End If
End Sub

Private Sub cboTpoAlm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtAlmacen.Visible And txtAlmacen.Enabled Then txtAlmacen.SetFocus
    End If
End Sub

'Private Sub cmdProcesar_Click()
'    Dim oALmacen As DLogAlmacen
'    Set oALmacen = New DLogAlmacen
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    Dim lnSaldoIni As Double
'    Dim lnIng As Double
'    Dim lnSal As Double
'    Dim lnStockIni As Long
'    Dim ldFecIni As Date
'    Dim lnMontoIni As Currency
'    Dim lnI As Integer
'
'    Dim nMes As Integer, nAnio As Integer
'    Dim cFecIni As String, cFecFin As String
'
'    lsCadena = Caption
'
'    If cboMes.ListIndex <= 0 Then
'       MsgBox "Debe indicar un mes válido..." + Space(10), vbInformation, "Aviso"
'       Exit Sub
'    End If
'
'    'If Len(txtAnio) = 0 Or Val(txtAnio) < Year(Date) Then
'    If Len(txtAnio) = 0 Then
'       MsgBox "Debe indicar un año válido..." + Space(10), vbInformation, "Aviso"
'       Exit Sub
'    End If
'
'    If Not IsNumeric(txtAlmacen.Text) Then 'EJVG 20111031
'       MsgBox "Seleccione un Almacen a procesar..." + Space(10), vbInformation, "Aviso"
'       Exit Sub
'    End If
'
'    nMes = cboMes.ListIndex
'    nAnio = CInt(txtAnio)
'
'    cFecIni = DateSerial(nAnio, nMes + 0, 1)
'    cFecFin = DateSerial(nAnio, nMes + 1, 0)
'
'    If YaHaySaldos(nMes, nAnio, Me.txtAlmacen.Text) Then
'       MsgBox "Ya se ha procesado " & Me.lblAlmacenG.Caption & " - " & cboMes.Text & " - " & txtAnio & "..." + Space(10), vbInformation, "Aviso"
'       Exit Sub
'    End If
''
''    If Not IsDate(Me.mskFecIni.Text) Then
''        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
''        mskFecIni.SetFocus
''        Exit Sub
''    ElseIf Not IsDate(Me.mskFecFin.Text) Then
''        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
''        mskFecFin.SetFocus
''        Exit Sub
''    End If
'
'    If MsgBox("¿ Está seguro de realizar el cierre de Almacén " & Me.lblAlmacenG.Caption + Space(10) & vbCrLf & "     para el mes de " + cboMes.Text + " - " + txtAnio & " ?" & vbCrLf & vbCrLf & "  Este proceso puede durar varios minutos", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
'       fraMes.Visible = False
'       fraBarra.Visible = True
'
'    'Procesar
''    If Me.txtAlmacen.Text <> "" And Me.txtProducto.Text <> "" Then
''        ldFecIni = CDate(Me.mskFecIni.Text)
''        lnStockIni = oALmacen.GetStock(Me.txtAlmacen.Text, Me.txtProducto.Text, Val(Right(Me.cboTpoAlm.Text, 5)), ldFecIni, lnMontoIni)
''        Me.prgBar.Max = Abs(DateDiff("d", CDate(Me.mskFecFin.Text), CDate(Me.mskFecIni.Text))) + 2
''        prgBar.Value = 0
''
''        lnI = 0
''
''        While DateAdd("d", lnI, ldFecIni) <= CDate(Me.mskFecFin.Text)
''            Caption = Format(DateAdd("d", lnI, ldFecIni), gsFormatoFechaView)
''            Set rs = oALmacen.GetLogMovimientoDia(Me.txtAlmacen.Text, Me.txtProducto.Text, DateAdd("d", lnI, CDate(Me.mskFecIni.Text)))
''            oALmacen.InsertaBSSaldos Me.txtAlmacen.Text, Me.txtProducto.Text, DateAdd("d", lnI, CDate(Me.mskFecIni.Text)), rs.Fields(0) + lnStockIni, rs.Fields(1) + lnMontoIni, Val(Right(Me.cboTpoAlm.Text, 5))
''            lnStockIni = rs.Fields(0) + lnStockIni
''            lnMontoIni = rs.Fields(1) + lnMontoIni
''            lnI = lnI + 1
''            prgBar.Value = prgBar.Value + 1
''        Wend
''    ElseIf Me.txtAlmacen.Text = "" And Me.txtProducto.Text <> "" Then
''        rsAlmacen.MoveFirst
''        Me.prgBar.Max = Abs((DateDiff("d", CDate(Me.mskFecFin.Text), CDate(Me.mskFecIni.Text)) + 2) * rsAlmacen.RecordCount)
''        prgBar.Value = 0
''
''        While rsAlmacen.EOF
''            ldFecIni = CDate(Me.mskFecIni.Text)
''            lnStockIni = oALmacen.GetStock(rsAlmacen.Fields(0), Me.txtProducto.Text, Val(Right(Me.cboTpoAlm.Text, 5)), ldFecIni, lnMontoIni)
''            lnI = 0
''            While DateAdd("d", lnI, ldFecIni) <= CDate(Me.mskFecFin.Text)
''                Set rs = oALmacen.GetLogMovimientoDia(rsAlmacen.Fields(0), Me.txtProducto.Text, DateAdd("d", lnI, CDate(Me.mskFecIni.Text)))
''                oALmacen.InsertaBSSaldos rsAlmacen.Fields(0), Me.txtProducto.Text, DateAdd("d", lnI, CDate(Me.mskFecIni.Text)), rs.Fields(0) + lnStockIni, rs.Fields(1) + lnMontoIni, Val(Right(Me.cboTpoAlm.Text, 5))
''                lnStockIni = rs.Fields(0) + lnStockIni
''                lnMontoIni = rs.Fields(1) + lnMontoIni
''                lnI = lnI + 1
''
''                prgBar.Value = prgBar.Value + 1
''            Wend
''            rsAlmacen.MoveNext
''        Wend
'
'     If Me.txtAlmacen.Text <> "" Then
'        rsProductos.MoveFirst
'
'        Me.prgBar.Max = Abs((DateDiff("d", CDate(cFecFin), CDate(cFecIni)) + 1) * rsProductos.RecordCount)
'        prgBar.value = 0
'
'        While DateAdd("d", lnI, ldFecIni) <= CDate(cFecFin)
'            While Not rsProductos.EOF
'                'If rsProductos.Fields(0) = "1110201229" Then
'                '    MsgBox "hola"
'                'End If
'
'                ldFecIni = CDate(cFecFin)
'                lnStockIni = oALmacen.GetStock(Me.txtAlmacen.Text, rsProductos.Fields(0), Val(Right(Me.cboTpoAlm.Text, 5)), ldFecIni, lnMontoIni)
'                lnI = 0
'                Set rs = oALmacen.GetLogMovimientoDiaRango(Me.txtAlmacen.Text, rsProductos.Fields(0), Val(Right(Me.cboTpoAlm.Text, 5)), DateAdd("d", 1, ldFecIni), DateAdd("d", lnI, CDate(cFecFin)))
'                oALmacen.InsertaBSSaldos Me.txtAlmacen.Text, rsProductos.Fields(0), CDate(cFecFin), rs.Fields(0) + lnStockIni, rs.Fields(1) + lnMontoIni, Val(Right(Me.cboTpoAlm.Text, 5))
'                lnStockIni = rs.Fields(0) + lnStockIni
'                lnMontoIni = rs.Fields(1) + lnMontoIni
'                Call oALmacen.LimpiaSaldoBSSaldoCierreMes(rsProductos.Fields(0), txtAlmacen.Text, Val(Right(Me.cboTpoAlm.Text, 5)), CDate(cFecFin)) 'EJVG20120102
'                If prgBar.value < Me.prgBar.Max Then prgBar.value = prgBar.value + 1
'                DoEvents
'                'rsProductos.MovePrevious
'                rsProductos.MoveNext
'            Wend
'            lnI = lnI + 1
'        Wend
'    Else
'        rsAlmacen.MoveFirst
'        While Not rsAlmacen.EOF
'            rsProductos.MoveFirst
'            While Not rsProductos.EOF
'                ldFecIni = CDate(cFecIni)
'                lnStockIni = oALmacen.GetStock(rsAlmacen.Fields(0), rsProductos.Fields(0), Val(Right(Me.cboTpoAlm.Text, 5)), ldFecIni, lnMontoIni)
'                lnI = 0
'
'                While DateAdd("d", lnI, ldFecIni) <= CDate(cFecFin)
'                    Set rs = oALmacen.GetLogMovimientoDia(rsAlmacen.Fields(0), rsProductos.Fields(0), DateAdd("d", lnI, CDate(cFecIni)))
'                    oALmacen.InsertaBSSaldos rsAlmacen.Fields(0), rsProductos.Fields(0), DateAdd("d", lnI, CDate(cFecIni)), rs.Fields(0) + lnStockIni, rs.Fields(1) + lnMontoIni, Val(Right(Me.cboTpoAlm.Text, 5))
'                    lnStockIni = rs.Fields(0) + lnStockIni
'                    lnMontoIni = rs.Fields(1) + lnMontoIni
'                    lnI = lnI + 1
'                Wend
'                rsProductos.MoveNext
'            Wend
'            rsProductos.MoveNext
'        Wend
'        rsAlmacen.MoveNext
'    End If
'
'    MsgBox "Proceso Finalizado", vbInformation, "Aviso"
'    Caption = lsCadena
'    Unload Me
'End Sub
Private Sub cmdProcesar_Click()
    Dim oAlmacen As DLogMov
    Dim nMes As Integer, nAnio As Integer
    Dim cFecIni As String, cFecFin As String
    Dim lnAlmTpo As Integer, lnAlmCod As Integer
    Dim ldFecha As Date
    Dim bTrans As Boolean
    Dim ldFechaUltCierre As Date, ldFechaNextUltCierre As Date
    Dim lbHaySaldoDespues As Boolean

    On Error GoTo ErrProcesar
    If cboTpoAlm.ListIndex = -1 Then
       MsgBox "Ud. debe seleccionar un Tipo de Almacen para continuar", vbInformation, "Aviso"
       If cboTpoAlm.Visible And cboTpoAlm.Enabled Then cboTpoAlm.SetFocus
       Exit Sub
    End If
    If cboMes.ListIndex <= 0 Then
       MsgBox "Ud. debe seleccionar un mes para continuar", vbInformation, "Aviso"
       If cboMes.Visible And cboMes.Enabled Then cboMes.SetFocus
       Exit Sub
    End If
    If Val(txtAnio.Text) = 0 Then
       MsgBox "Ud. debe indicar un año válido para continuar", vbInformation, "Aviso"
       If txtAnio.Visible And txtAnio.Enabled Then txtAnio.SetFocus
       Exit Sub
    End If
    If Val(txtAlmacen.Text) = 0 Then
       MsgBox "Ud. debe seleccionar un Almacen para continuar", vbInformation, "Aviso"
       If txtAlmacen.Visible And txtAlmacen.Enabled Then txtAlmacen.SetFocus
       Exit Sub
    End If
    
    nMes = cboMes.ListIndex
    nAnio = Val(txtAnio.Text)
    lnAlmTpo = Val(Right(cboTpoAlm.Text, 5))
    lnAlmCod = Val(txtAlmacen.Text)
    ldFecha = obtenerFechaFinMes(nMes, nAnio)
    ldFechaUltCierre = ObtenerUltimoCierre(lnAlmTpo, lnAlmCod, ldFecha)
    lbHaySaldoDespues = HaySaldosDespues(lnAlmTpo, lnAlmCod, ldFecha)
    
    If lbHaySaldoDespues Then
       MsgBox "No se puede continuar, ya que después de esta fecha ya realizaron cierre de Almacen", vbExclamation, "Aviso"
       Exit Sub
    End If
    If DateDiff("D", ldFechaUltCierre, CDate("01/01/1900")) = 0 Then 'Si en caso crean nuevos almacenes que pase nomas
    Else
        ldFechaNextUltCierre = DateAdd("M", 1, ldFechaUltCierre)
        ldFechaNextUltCierre = obtenerFechaFinMes(Month(ldFechaNextUltCierre), Year(ldFechaNextUltCierre))
        If DateDiff("D", ldFechaNextUltCierre, ldFecha) <> 0 Then
            MsgBox "No se puede continuar, ya que la sgte fecha de cierre del Almacen es en el Periodo " & Format(ldFechaNextUltCierre, "mmm-yyyy"), vbExclamation, "Aviso"
            Exit Sub
        End If
    End If
    
    If YaHaySaldos(Month(ldFecha), Year(ldFecha), lnAlmCod, lnAlmTpo) Then
        If MsgBox("El cierre del Almacén [" & UCase(lblAlmacenG.Caption) & "] para el mes de " & cboMes.Text + " - " & txtAnio & " ya se realizó." & vbNewLine & "¿Desea volver a procesarlo?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            Exit Sub
        End If
    Else
        If MsgBox("¿Está seguro de realizar el cierre de Almacén [" & UCase(lblAlmacenG.Caption) & "] para el mes de " & cboMes.Text & " - " & txtAnio.Text & "?" & vbNewLine & "Este proceso puede durar varios minutos", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            Exit Sub
        End If
    End If
    Screen.MousePointer = 11
    Set oAlmacen = New DLogMov
    oAlmacen.BeginTrans
    bTrans = True
    oAlmacen.CierreSaldosAlmacen lnAlmTpo, lnAlmCod, ldFecha
    oAlmacen.CommitTrans
    bTrans = False
    Screen.MousePointer = 0
    MsgBox "Se ha realizado el cierre de Almacen satisfactoriamente", vbInformation, "Aviso"
    Set oAlmacen = Nothing
        'ARLO 20160126 ***
        gsOpeCod = LogPistaInventarioAlmacen
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Para : " & Left(cboTpoAlm.Text, 15) & " | A la Fecha : " & ldFechaNextUltCierre & " | Del Almacen" & UCase(lblAlmacenG.Caption)
        Set objPista = Nothing
        '**************
    Exit Sub
ErrProcesar:
    Screen.MousePointer = 0
    If bTrans Then
        oAlmacen.RollbackTrans
        Set oAlmacen = Nothing
    End If
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oDoc As DOperaciones
    Set oDoc = New DOperaciones
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set rs = oCon.GetConstante(5010, False)
    CargaCombo rs, cboTpoAlm
    cboTpoAlm.ListIndex = 0
    
    Set rsAlmacen = oDoc.GetAlmacenes
    Set rsProductos = oAlmacen.GetBienes
    
    Me.txtAlmacen.rs = oDoc.GetAlmacenes
    'Me.txtProducto.rs = oALmacen.GetBienes
    CentraForm Me
    txtAnio = Year(Date)
cboMes.AddItem "----- SELECCIONE MES -----"
cboMes.AddItem "ENERO"
cboMes.AddItem "FEBRERO"
cboMes.AddItem "MARZO"
cboMes.AddItem "ABRIL"
cboMes.AddItem "MAYO"
cboMes.AddItem "JUNIO"
cboMes.AddItem "JULIO"
cboMes.AddItem "AGOSTO"
cboMes.AddItem "SEPTIEMBRE"
cboMes.AddItem "OCTUBRE"
cboMes.AddItem "NOVIEMBRE"
cboMes.AddItem "DICIEMBRE"
cboMes.ListIndex = 0
End Sub

'------------------------------------------
'Para verificar la existencia de saldos
'------------------------------------------
Function YaHaySaldos(pnMes As Integer, pnAnio As Integer, pnAlmacen As Integer, Optional ByVal pnAlmTpo As Integer = 0) As Boolean
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta

YaHaySaldos = False
If oConn.AbreConexion Then
   'Set rs = oConn.CargaRecordSet("select top 1 cBSCod from BSSaldos where month(dSaldo)=" & pnMes & " and year(dSaldo)=" & pnAnio & " And nAlmCod = " & pnAlmacen & "")
   Set rs = oConn.CargaRecordSet("select top 1 cBSCod from BSSaldos where month(dSaldo)=" & pnMes & " and year(dSaldo)=" & pnAnio & " And nAlmCod = " & pnAlmacen & " And nAlmTpo = " & pnAlmTpo) 'EJVG20140924
   If Not rs.EOF Then
      YaHaySaldos = True
   End If
End If
End Function


Private Sub txtAlmacen_EmiteDatos()
    Me.lblAlmacenG.Caption = Me.txtAlmacen.psDescripcion
End Sub
'EJVG20140211 ***
Private Function HaySaldosDespues(ByVal pnAlmTpo As Integer, ByVal pnAlmCod As Integer, ByVal pdFecha As Date) As Boolean
    Dim rs As New ADODB.Recordset
    Dim oConn As New DConecta

    oConn.AbreConexion
    Set rs = oConn.CargaRecordSet("SELECT TOP 1 cBSCod FROM BSSaldos WHERE nAlmTpo = " & pnAlmTpo & " And nAlmCod = " & pnAlmCod & " And dSaldo > '" & Format(pdFecha, "yyyymmdd") & "'")
    If Not rs.EOF Then
       HaySaldosDespues = True
    End If
    oConn.CierraConexion
    Set rs = Nothing
    Set oConn = Nothing
End Function
Private Function ObtenerUltimoCierre(ByVal pnAlmTpo As Integer, ByVal pnAlmCod As Integer, ByVal pdFecha As Date) As Date
    Dim rs As New ADODB.Recordset
    Dim oConn As New DConecta

    ObtenerUltimoCierre = CDate("01/01/1900")
    oConn.AbreConexion
    Set rs = oConn.CargaRecordSet("SELECT TOP 1 dSaldo FROM BSSaldos WHERE nAlmTpo = " & pnAlmTpo & " And nAlmCod = " & pnAlmCod & " And dSaldo < '" & Format(pdFecha, "yyyymmdd") & "' ORDER BY dSaldo DESC")
    If Not rs.EOF Then
       ObtenerUltimoCierre = rs!dSaldo
    End If
    oConn.CierraConexion
    Set rs = Nothing
    Set oConn = Nothing
End Function
Private Function obtenerFechaFinMes(ByVal pnMes As Integer, ByVal pnAnio As Integer) As Date
    Dim sFecha  As Date
    sFecha = CDate("01/" & Format(pnMes, "00") & "/" & pnAnio)
    sFecha = DateAdd("m", 1, sFecha)
    sFecha = sFecha - 1
    obtenerFechaFinMes = sFecha
End Function
'END EJVG *******
Private Sub txtAlmacen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboMes.Visible And cboMes.Enabled Then cboMes.SetFocus
    End If
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    If cmdProcesar.Visible And cmdProcesar.Enabled Then cmdProcesar.SetFocus
End Sub

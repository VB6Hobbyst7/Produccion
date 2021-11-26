VERSION 5.00
Begin VB.Form frmARendirLista 
   Caption         =   "A rendir Cuenta: Pendientes de regularizar"
   ClientHeight    =   5715
   ClientLeft      =   225
   ClientTop       =   1695
   ClientWidth     =   11040
   Icon            =   "frmARendirExtRend.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   11040
   Begin VB.CommandButton cmdRegulariza 
      Caption         =   "S&ustentación"
      Height          =   405
      Left            =   7575
      TabIndex        =   15
      ToolTipText     =   "Regularizar con Documentos sustentatorios"
      Top             =   5160
      Width           =   1680
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "&Extornar"
      Height          =   405
      Left            =   7575
      TabIndex        =   14
      ToolTipText     =   "Ingresar Saldo a Caja General"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
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
      Left            =   8895
      TabIndex        =   2
      Top             =   735
      Width           =   1785
   End
   Begin Sicmact.Usuario usu 
      Left            =   885
      Top             =   5760
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CheckBox chkSelec 
      Caption         =   "&Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   315
      TabIndex        =   9
      Top             =   135
      Value           =   1  'Checked
      Width           =   900
   End
   Begin VB.Frame FraSeleccion 
      Enabled         =   0   'False
      Height          =   1110
      Left            =   150
      TabIndex        =   8
      Top             =   120
      Width           =   10800
      Begin VB.CheckBox chkTodo 
         Caption         =   "&Incluir A rendir sin Saldo"
         Height          =   285
         Left            =   8535
         TabIndex        =   1
         Top             =   270
         Width           =   1995
      End
      Begin Sicmact.TxtBuscar txtBuscarAgenciaArea 
         Height          =   330
         Left            =   1065
         TabIndex        =   0
         Top             =   247
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   582
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         lbUltimaInstancia=   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Agencia :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   675
         Width           =   825
      End
      Begin VB.Label lblAgenciArea 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2175
         TabIndex        =   12
         Top             =   262
         Width           =   5715
      End
      Begin VB.Label lblAgeDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1065
         TabIndex        =   11
         Top             =   622
         Width           =   6825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Area :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   10
         Top             =   315
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   405
      Left            =   9255
      TabIndex        =   6
      Top             =   5160
      Width           =   1680
   End
   Begin VB.CommandButton cmdRendicion 
      Caption         =   "Rendicion Caja General"
      Height          =   405
      Left            =   5565
      TabIndex        =   5
      ToolTipText     =   "Ingresar Saldo a Caja General"
      Top             =   5160
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.CommandButton cmdSaldoCh 
      Caption         =   "&Rendición de Saldo"
      Height          =   405
      Left            =   150
      TabIndex        =   4
      ToolTipText     =   "Ingresar Saldo a Caja Chica"
      Top             =   5145
      Visible         =   0   'False
      Width           =   1680
   End
   Begin Sicmact.FlexEdit fgAtenciones 
      Height          =   2835
      Left            =   180
      TabIndex        =   3
      Top             =   1305
      Width           =   10800
      _ExtentX        =   16616
      _ExtentY        =   2672
      Cols0           =   19
      HighLight       =   2
      AllowUserResizing=   3
      EncabezadosNombres=   $"frmARendirExtRend.frx":030A
      EncabezadosAnchos=   "350-450-1200-900-1200-900-3000-1200-0-0-0-0-1200-2000-0-0-0-2000-0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-C-L-C-L-R-C-R-L-L-R-L-L-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-2-0-2-0-0-2-0-0-0-0-0-0"
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      ColWidth0       =   345
      RowHeight0      =   285
   End
   Begin VB.TextBox txtMovDesc 
      Height          =   825
      Left            =   165
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4200
      Width           =   10800
   End
End
Attribute VB_Name = "frmARendirLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cOpeCod As String, cOpeDesc As String
Dim lTransActiva As Boolean
Dim lRindeCajaG As Boolean, lEsChica As Boolean
Dim lRindeCajaCh As Boolean, lRindeViaticos As Boolean
Dim sObjRendir As String
Dim sDocTpoRecibo As String


'************************************************************************************
'************************************************************************************
Dim oContFunc As NContFunciones
Dim oAreas As DActualizaDatosArea
Dim oNArendir As NARendir
Dim oOperacion As DOperacion

Dim lnTipoArendir As ArendirTipo
Dim lbEsChica  As Boolean
Dim lsCtaArendir As String
Dim lsCtaPendiente As String
Dim lsDocTpoRecibo As String

Dim lsTpoDocVoucher  As String
Dim lSalir As Boolean
Dim lsMovNroSolicitud As String
Dim lnArendirFase As ARendirFases
Public Sub Inicio(ByVal pnTipoArendir As ArendirTipo, ByVal pnArendirFase As ARendirFases, Optional pbEsCajaChica As Boolean = False)
lnArendirFase = pnArendirFase
lnTipoArendir = pnTipoArendir
lbEsChica = pbEsCajaChica
Me.Show 1
End Sub

Private Function GetReciboEgreso() As Boolean
Dim lnFila As Long
Dim rs As ADODB.Recordset
GetReciboEgreso = False
lSalir = False
Set rs = New ADODB.Recordset
fgAtenciones.Clear
fgAtenciones.FormaCabecera
fgAtenciones.Rows = 2
Me.MousePointer = 11
If chkSelec.Value = 0 Then
    If txtBuscarAgenciaArea = "" Then
        MsgBox "Ingrese el Area a la cual Pertenece el Arendir", vbInformation, "Aviso"
        txtBuscarAgenciaArea.SetFocus
        Exit Function
    End If
End If
Select Case lnArendirFase
     Case ArendirExtornoAtencion
        Set rs = oNArendir.GetAtencionSinSustentacion(chkSelec.Value, Mid(txtBuscarAgenciaArea.Text, 4, 2), Mid(txtBuscarAgenciaArea.Text, 1, 3), lnTipoArendir, lsCtaArendir)
     Case ArendirExtornoRendicion
        Set rs = oNArendir.GetAtencionPendArendir(chkSelec.Value, Mid(txtBuscarAgenciaArea.Text, 4, 2), Mid(txtBuscarAgenciaArea.Text, 1, 3), lnTipoArendir, lsCtaArendir)
     Case Else
        Set rs = oNArendir.GetAtencionPendArendir(chkSelec.Value, Mid(txtBuscarAgenciaArea.Text, 4, 2), Mid(txtBuscarAgenciaArea.Text, 1, 3), lnTipoArendir, lsCtaArendir)
End Select
If Not rs.EOF And Not rs.BOF Then
    Do While Not rs.EOF
       fgAtenciones.AdicionaFila
       lnFila = fgAtenciones.Row
       fgAtenciones.TextMatrix(lnFila, 1) = Trim(rs!cAbreDocAtenc)
       fgAtenciones.TextMatrix(lnFila, 2) = Trim(rs!cNroDocAtenc)
       fgAtenciones.TextMatrix(lnFila, 3) = IIf(IsNull(rs!dDocAtencion), "", rs!dDocAtencion)
       fgAtenciones.TextMatrix(lnFila, 4) = Trim(rs!cDocNro)
       fgAtenciones.TextMatrix(lnFila, 5) = rs!dDocSolicitud
       fgAtenciones.TextMatrix(lnFila, 6) = PstaNombre(rs!cPersNombre)
       fgAtenciones.TextMatrix(lnFila, 7) = Format(rs!nImporte, gsFormatoNumeroView)
       fgAtenciones.TextMatrix(lnFila, 8) = rs!cMovDesc
       fgAtenciones.TextMatrix(lnFila, 9) = rs!cPersCod
       fgAtenciones.TextMatrix(lnFila, 10) = rs!cMovNroAtenc
       fgAtenciones.TextMatrix(lnFila, 11) = rs!cTpoDocAtenc
       fgAtenciones.TextMatrix(lnFila, 12) = Format(rs!nMovSaldo, gsFormatoNumeroView)
       fgAtenciones.TextMatrix(lnFila, 13) = rs!cAreaDescripcion
       fgAtenciones.TextMatrix(lnFila, 14) = rs!cAreaCod
       fgAtenciones.TextMatrix(lnFila, 15) = rs!cDocDesc
       fgAtenciones.TextMatrix(lnFila, 16) = rs!cMovNroSol
       fgAtenciones.TextMatrix(lnFila, 17) = rs!cAgeDescripcion
       fgAtenciones.TextMatrix(lnFila, 18) = rs!cAgecod
       
       If lnTipoArendir = gArendirTipoCajaChica Then
          'lvItem.SubItems(4) = Format(rs!nMovImporte * -1, gsFormatoNumeroView)
          'lvItem.SubItems(9) = Format(rs!nSaldo * -1, gsFormatoNumeroView)
          'lvItem.SubItems(10) = rs!AgeDes
          'lvItem.SubItems(11) = rs!AgeCod
       Else
          'lvItem.SubItems(11) = txtAgeCod
          'lvItem.SubItems(10) = txtAgeDesc
       End If
       rs.MoveNext
    Loop
Else
    If lnTipoArendir = gArendirTipoCajaChica Then
        MsgBox "Caja Chica sin egresos pendientes de A rendir", vbInformation, "Aviso"
    Else
        MsgBox "Area funcional sin A rendir Cuenta Pendientes", vbInformation, "Aviso"
    End If
    'Exit Function
End If
rs.Close: Set rs = Nothing
GetReciboEgreso = True
Me.MousePointer = 0
End Function
Private Function ValidaAgencia(sAgeCod As String) As Boolean
ValidaAgencia = False
If lEsChica Then
   SSQL = "SELECT a.cObjetoCod, a.cObjetoDesc, a.nObjetoNiv " _
     & "FROM   " & gcCentralCom & "Objeto a, varsistema b " _
     & "WHERE  b.cCodProd = 'CON' and substring(b.cNomVar,1,4) = 'cCCH' and " _
     & "       substring(b.cNomVar,5,5) = a.cObjetoCod " _
     & "   " & IIf(sAgeCod = "", "", " and a.cObjetoCod = '" & sAgeCod & "' ")
Else
   SSQL = "SELECT h.cObjetoCod, h.cObjetoDesc, " _
     & "  h.nObjetoNiv FROM   " & gcCentralCom & "Objeto h WHERE  " _
     & "  cObjetoCod like '" & IIf(sAgeCod = "", "11", sAgeCod & "%' ")
End If
lvRecibo.ListItems.Clear
If rs.State = adStateOpen Then rs.Close
rs.Open SSQL, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
If rs.EOF Then
   If Not lEsChica Then
      If sAgeCod = "" Then
         MsgBox "No existen Areas con A rendir cuenta Pendiente", vbCritical, "Error"
      Else
         MsgBox "Area funcional no tiene A rendir Cuenta pendiente", vbCritical, "Error"
      End If
   Else
      If sAgeCod = "" Then
         MsgBox "No existen Areas asignadas como Caja Chica ", vbCritical, "Error"
      Else
         MsgBox "Area funcional no asignada como Caja Chica", vbCritical, "Error"
      End If
   End If
   rs.Close: Set rs = Nothing
Else
   ValidaAgencia = True
End If
End Function
Private Sub chkSelec_Click()
If chkSelec.Value = 0 Then
    FraSeleccion.Enabled = True
    txtBuscarAgenciaArea.SetFocus
Else
    FraSeleccion.Enabled = False
    txtBuscarAgenciaArea.Text = ""
    lblAgeDesc = ""
    lblAgenciArea = ""
    chkTodo.Value = 0
End If
End Sub

Private Sub cmdExtornar_Click()
Dim ldFechaAtenc As Date
Dim lsMovAtenc As String
Dim lnImporte As Currency
Dim lsDocTpo As String
Dim lsOpeDoc As String
Dim lsMovNro As String
If fgAtenciones.TextMatrix(1, 0) = "" Then Exit Sub
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Ingrese la descripcion del extorno ", vbInformation, "Aviso"
    txtMovDesc.SetFocus
    Exit Sub
End If
lsOpeDoc = gsOpeCod
If MsgBox("Desea Realizar el extorno??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Select Case lnArendirFase
        Case ArendirExtornoAtencion
            lsMovAtenc = fgAtenciones.TextMatrix(fgAtenciones.Row, 10)
            lsMovNroSolicitud = fgAtenciones.TextMatrix(fgAtenciones.Row, 16)
            lnImporte = CCur(fgAtenciones.TextMatrix(fgAtenciones.Row, 7))
            ldFechaAtenc = CDate(Mid(lsMovAtenc, 7, 2) & "/" & Mid(lsMovAtenc, 5, 2) & "/" & Mid(lsMovAtenc, 1, 4))
            lsDocTpo = IIf(fgAtenciones.TextMatrix(fgAtenciones.Row, 11) = "", "-1", fgAtenciones.TextMatrix(fgAtenciones.Row, 11))
            lsOpeDoc = oNArendir.GetOpeRendicion(Mid(gsOpeCod, 1, 5), lsDocTpo, lsCtaArendir, lsCtaPendiente)
            lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            If ldFechaAtenc <> gdFecSis Then
                oNArendir.ExtornaAtencionArendir lsMovNro, lsOpeDoc, txtMovDesc, _
                            lsMovAtenc, lsMovNroSolicitud, lnTipoArendir, lnImporte, False
                            
                ImprimeAsientoContable lsMovNro
            Else
                oNArendir.EliminaAtencionArendir lsMovNro, lsOpeDoc, txtMovDesc, lsMovAtenc, lsMovNroSolicitud, lnTipoArendir, lnImporte, False
            End If
        Case ArendirExtornoRendicion
    End Select
    fgAtenciones.EliminaFila fgAtenciones.Row
End If
End Sub

Private Sub cmdProcesar_Click()
If GetReciboEgreso Then
    fgAtenciones.SetFocus
Else
    txtBuscarAgenciaArea.Text = ""
    lblAgeDesc = ""
    lblAgenciArea = ""
    txtBuscarAgenciaArea.SetFocus
End If
End Sub

Private Sub cmdRegulariza_Click()
Dim sRecEstado As String
Dim lsNroArendir As String
Dim lsNroDoc As String
Dim lsFechaDoc As String
Dim lsPersCod As String
Dim lsPersNomb As String
Dim lsAreaCod As String
Dim lsAreaDesc As String
Dim lsDescDoc As String
Dim lnImporte As Currency
Dim lnSaldo As Currency
Dim lsMovNroAtenc As String
Dim lsMovNroSolicitud As String
Dim lsAgeCod As String
Dim lsAgeDesc As String

If fgAtenciones.TextMatrix(1, 0) = "" Then
   Exit Sub
End If
lsNroArendir = fgAtenciones.TextMatrix(fgAtenciones.Row, 4)
lsNroDoc = fgAtenciones.TextMatrix(fgAtenciones.Row, 2)
lsFechaDoc = fgAtenciones.TextMatrix(fgAtenciones.Row, 5)
lsPersCod = fgAtenciones.TextMatrix(fgAtenciones.Row, 9)
lsPersNomb = fgAtenciones.TextMatrix(fgAtenciones.Row, 6)
lsAreaCod = fgAtenciones.TextMatrix(fgAtenciones.Row, 14)
lsAreaDesc = fgAtenciones.TextMatrix(fgAtenciones.Row, 13)

lsDescDoc = fgAtenciones.TextMatrix(fgAtenciones.Row, 15)
lnImporte = CCur(fgAtenciones.TextMatrix(fgAtenciones.Row, 7))
lnSaldo = CCur(fgAtenciones.TextMatrix(fgAtenciones.Row, 12))
lsMovNroAtenc = fgAtenciones.TextMatrix(fgAtenciones.Row, 10)
lsMovNroSolicitud = fgAtenciones.TextMatrix(fgAtenciones.Row, 16)
lsAgeDesc = fgAtenciones.TextMatrix(fgAtenciones.Row, 17)
lsAgeCod = fgAtenciones.TextMatrix(fgAtenciones.Row, 18)

frmOpeRegDocs.Inicio lnArendirFase, lnTipoArendir, False, lsNroArendir, lsNroDoc, lsFechaDoc, lsPersCod, _
                     lsPersNomb, lsAreaCod, lsAreaDesc, lsAgeCod, lsAgeDesc, lsDescDoc, lsMovNroAtenc, lnImporte, lsCtaArendir, _
                     lsCtaPendiente, lnSaldo, lsMovNroSolicitud

fgAtenciones.TextMatrix(fgAtenciones.Row, 12) = Format(frmOpeRegDocs.lnSaldo, gsFormatoNumeroView)
'If Val(fgAtenciones.TextMatrix(fgAtenciones.Row, 12)) = 0 Then  'Doc. Regularizado
'    fgAtenciones.EliminaFila fgAtenciones.Row
'End If
fgAtenciones.SetFocus
End Sub
Private Sub cmdRendicion_Click()
Dim sRecEstado As String
Dim lsNroArendir As String
Dim lsNroDoc As String
Dim lsFechaDoc As String
Dim lsPersCod As String
Dim lsPersNomb As String
Dim lsAreaCod As String
Dim lsAreaDesc As String
Dim lsDescDoc As String
Dim lnImporte As Currency
Dim lnSaldo As Currency
Dim lsMovNroAtenc As String
Dim lsMovNroSolicitud As String
Dim lsAgeCod As String
Dim lsAgeDesc As String

If fgAtenciones.TextMatrix(1, 0) = "" Then
    MsgBox "No existen Atenciones Pendientes", vbInformation, "Aviso"
    Exit Sub
End If

lsNroArendir = fgAtenciones.TextMatrix(fgAtenciones.Row, 4)
lsNroDoc = fgAtenciones.TextMatrix(fgAtenciones.Row, 2)
lsFechaDoc = fgAtenciones.TextMatrix(fgAtenciones.Row, 5)
lsPersCod = fgAtenciones.TextMatrix(fgAtenciones.Row, 9)
lsPersNomb = fgAtenciones.TextMatrix(fgAtenciones.Row, 6)
lsAreaCod = fgAtenciones.TextMatrix(fgAtenciones.Row, 14)
lsAreaDesc = fgAtenciones.TextMatrix(fgAtenciones.Row, 13)

lsDescDoc = fgAtenciones.TextMatrix(fgAtenciones.Row, 15)
lnImporte = CCur(fgAtenciones.TextMatrix(fgAtenciones.Row, 7))
lnSaldo = CCur(fgAtenciones.TextMatrix(fgAtenciones.Row, 12))
lsMovNroAtenc = fgAtenciones.TextMatrix(fgAtenciones.Row, 10)
lsMovNroSolicitud = fgAtenciones.TextMatrix(fgAtenciones.Row, 16)
lsAgeDesc = fgAtenciones.TextMatrix(fgAtenciones.Row, 17)
lsAgeCod = fgAtenciones.TextMatrix(fgAtenciones.Row, 18)

frmArendirRendicion.Inicio lnArendirFase, lnTipoArendir, lsNroArendir, lsNroDoc, lsFechaDoc, lsPersCod, _
                     lsPersNomb, lsAreaCod, lsAreaDesc, lsAgeCod, lsAgeDesc, lsDescDoc, lsMovNroAtenc, lnImporte, lsCtaArendir, _
                     lsCtaPendiente, lnSaldo, lsMovNroSolicitud, txtMovDesc

If frmArendirRendicion.vbOk Then
    fgAtenciones.EliminaFila fgAtenciones.Row
End If
End Sub


Private Sub cmdSaldoCh_Click()
Dim sMovNro As String
Dim sAgeCod As String
Dim nItem   As Integer
Dim sCtaCod As String
If lvRecibo.ListItems.Count = 0 Then
   Exit Sub
End If
gnImporte = Val(Format(lvRecibo.SelectedItem.SubItems(9), gsFormatoNumeroDato))
sMovNro = lvRecibo.SelectedItem.SubItems(7)
If MsgBox("Caja Chica " & IIf(gnImporte < 0, "reintegrará", "recibirá") & " saldo de " & gcMN & " " & Format(gnImporte * IIf(gnImporte < 0, -1, 1), gsFormatoNumeroView) & Chr(10) & " ¿ Seguro de Grabar Operación ? ", vbQuestion + vbOKCancel, "Confirmación") = vbOk Then
   SSQL = "SELECT a.cCtaContCod, b.cCtaContDesc, a.cOpeCtaDH " _
        & "FROM  " & gcCentralCom & "OpeCta a,  " & gcCentralCom & "CtaCont b " _
        & "WHERE  a.cOpeCod = '" & gsOpeCod & "' and b.cCtaContCod = a.cCtaContCod and a.cOpeCtaTpo = '0' "
   Set rs = CargaRecord(SSQL)
   If rs.RecordCount <> 2 Then
      MsgBox "Operación sólo trabaja con Cuentas de Caja Chica y A rendir", vbCritical, "Error"
      Exit Sub
   End If
   If lTransActiva Then
      dbCmact.RollbackTrans
      lTransActiva = False
   End If
   dbCmact.BeginTrans
   lTransActiva = True
   txtMovNro = GeneraMovNro
   If gnImporte > 0 Then
      gcGlosa = "Ingreso de efectivo correspondiente al " & lvRecibo.SelectedItem.Text & " Nro. " & lvRecibo.SelectedItem.SubItems(1)
   Else
      gcGlosa = "Desembolso de efectivo correspondiente al " & lvRecibo.SelectedItem.Text & " Nro. " & lvRecibo.SelectedItem.SubItems(1)
   End If
   SSQL = "INSERT INTO Mov (cMovNro, cOpeCod, cMovDesc, cMovEstado) VALUES ('" & txtMovNro & "', '" & gsOpeCod & "', '" & gsOpeCod & "', '0')"
   dbCmact.Execute SSQL
   nItem = 0
   Do While Not rs.EOF
      nItem = nItem + 1

      If rs!cOpeCtaDH = "D" Then
         SSQL = "INSERT INTO  MovObj VALUES ('" & txtMovNro & "', '" & Format(nItem, "000") & "', '1', '" _
              & txtAgeCod & "')"
         dbCmact.Execute SSQL
         sCtaCod = GetCtaObjFiltro(rs!cCtaContCod, txtAgeCod)
      Else
         If gnImporte > 0 Then
            SSQL = "INSERT INTO  MovObj VALUES ('" & txtMovNro & "', '" & Format(nItem, "000") & "', '1', '" & sObjRendir & "')"
            dbCmact.Execute SSQL
            sCtaCod = GetCtaObjFiltro(rs!cCtaContCod, sObjRendir)
            SSQL = "INSERT INTO  MovObj VALUES ('" & txtMovNro & "', '" & Format(nItem, "000") & "', '2', '" & lvRecibo.SelectedItem.SubItems(11) & "')"
            dbCmact.Execute SSQL
            sCtaCod = sCtaCod & GetCtaObjFiltro(rs!cCtaContCod, lvRecibo.SelectedItem.SubItems(11), False)
            SSQL = "INSERT INTO  MovObj VALUES ('" & txtMovNro & "', '" & Format(nItem, "000") & "', '3', '" _
                 & lvRecibo.SelectedItem.SubItems(6) & "')"
            dbCmact.Execute SSQL
            sCtaCod = sCtaCod & GetCtaObjFiltro(rs!cCtaContCod, lvRecibo.SelectedItem.SubItems(6), False)
         Else
            sCtaCod = sCtaPendiente
         End If
      End If
      SSQL = "INSERT INTO  MovCta VALUES ('" & txtMovNro & "', '" & Format(nItem, "000") & "', '" _
           & sCtaCod & " ', " & gnImporte * IIf(rs!cOpeCtaDH = "D", 1, -1) & ")"
      dbCmact.Execute SSQL
      rs.MoveNext
   Loop
   SSQL = "INSERT INTO MovRef VALUES ('" & txtMovNro & "', '" & sMovNro & "')"
   dbCmact.Execute SSQL
   dbCmact.CommitTrans
   lTransActiva = False
   If gnImporte > 0 Then
      ImprimeReciboIngreso
   End If
   lvRecibo.ListItems.Remove lvRecibo.SelectedItem.Index
End If
End Sub
Private Sub ImprimeReciboIngreso()
Dim sTexto As String
Dim nAncho As Integer, n As Integer
rtxtAsiento.Text = ""
nAncho = gnColPage
  gcMovNro = txtMovNro
  sTexto = ""
  Lin1 sTexto, ImpreCabAsiento(" R E C I B O   D E   I N G R E S O ", , , , 1)
  Lin1 sTexto, "  CAJA CHICA   : " & BON & txtAgeCod & "  " & ImpreCarEsp(txtAgeDesc) & BOFF
  Lin1 sTexto, "  Area/Agencia : " & BON & lvRecibo.SelectedItem.SubItems(11) & "  " & ImpreCarEsp(lvRecibo.SelectedItem.SubItems(10)) & BOFF
  Lin1 sTexto, "  Persona      : " & BON & lvRecibo.SelectedItem.SubItems(6) & "  " & ImpreCarEsp(lvRecibo.SelectedItem.SubItems(2)) & BOFF
  Lin1 sTexto, "  Cargo        : "
  Lin1 sTexto, "  Importe      : " & BON & ConvNumLet(gnImporte, False) & BOFF
  Lin1 sTexto, ImpreGlosa("  Concepto     : ")
  Lin1 sTexto, "", 6
  Lin1 sTexto, BON & Space(20) & "___________________________"
  Lin1 sTexto, Space(20) & CON & Centra(ImpreCarEsp("Firma y Sello de Encargado"), 46) & COFF & BOFF, 3
  rtxtAsiento.Text = sTexto & Chr(10) & Chr(10) & IIf(gnLinPage = gnLinHori, Chr$(12), "") & sTexto
  frmPrevio.Previo rtxtAsiento, "Recibo de Ingreso", False, gnLinPage
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub fgAtenciones_Click()
If lnArendirFase <> ArendirExtornoAtencion And lnArendirFase <> ArendirExtornoRendicion Then
        txtMovDesc = fgAtenciones.TextMatrix(fgAtenciones.Row, 8)
    Else
        txtMovDesc = ""
    End If
End Sub
Private Sub fgAtenciones_GotFocus()
If lnArendirFase <> ArendirExtornoAtencion And lnArendirFase <> ArendirExtornoRendicion Then
        txtMovDesc = fgAtenciones.TextMatrix(fgAtenciones.Row, 8)
    Else
        txtMovDesc = ""
    End If

End Sub
Private Sub fgAtenciones_OnRowChange(pnRow As Long, pnCol As Long)
If fgAtenciones.TextMatrix(1, 0) <> "" Then
    If lnArendirFase <> ArendirExtornoAtencion And lnArendirFase <> ArendirExtornoRendicion Then
        txtMovDesc = fgAtenciones.TextMatrix(fgAtenciones.Row, 8)
    Else
        txtMovDesc = ""
    End If
End If

End Sub

Private Sub Form_Activate()
If lSalir Then
   Unload Me
End If
End Sub

Private Sub Form_Load()
Dim lvItem As ListItem
Dim rsPer As New ADODB.Recordset
Dim sOpeCod As String
 

Set oContFunc = New NContFunciones
Set oNArendir = New NARendir
Set oAreas = New DActualizaDatosArea
Set oOperacion = New DOperacion
lSalir = False
Me.Caption = gsOpeDesc

CentraForm Me

If Mid(gsOpeCod, 3, 1) = gMonedaNacional Then
   gsSimbolo = gcMN
Else
   gsSimbolo = gcME
End If
lsTpoDocVoucher = oOperacion.EmiteDocOpe(gsOpeCod, OpcionalDebeExistir, Autogenerado)
lsCtaArendir = oOperacion.EmiteOpeCta(gsOpeCod, "H", "0")
If lsCtaArendir = "" Then
   MsgBox "Faltan asignar Cuentas Contables a Operación." & Chr(10) & "Por favor consultar con Sistemas", vbInformation, "Aviso"
   lSalir = True
   Exit Sub
End If
lsCtaPendiente = oOperacion.EmiteOpeCta(gsOpeCod, "H", "1")
If lsCtaPendiente = "" Then
   MsgBox "Falta asignar Cuenta de Pendiente a Operación." & Chr(10) & "Por favor consultar con Sistemas", vbInformation, "Aviso"
   lSalir = True
   Exit Sub
End If
If lnTipoArendir = gArendirTipoCajaChica Then
    txtBuscarAgenciaArea.psRaiz = "CAJAS CHICAS"
    txtBuscarAgenciaArea.rs = oNArendir.EmiteCajasChicas
Else
    txtBuscarAgenciaArea.rs = oAreas.GetAgenciasAreas
End If
chkTodo.Visible = True
Select Case lnArendirFase
    Case ArendirSustentacion
        cmdRendicion.Visible = False
        cmdExtornar.Visible = False
    Case ArendirRendicion
        cmdRendicion.Visible = True
        cmdExtornar.Visible = False
    Case ArendirExtornoAtencion, ArendirExtornoRendicion
        chkTodo.Visible = False
        cmdRendicion.Visible = False
        cmdRegulariza.Visible = False
        txtMovDesc.Locked = False
        cmdExtornar.Visible = True
End Select
Select Case lnTipoArendir
    Case gArendirTipoCajaChica
         Me.Height = 5550
         cmdRendicion.Top = 5030 - cmdRendicion.Height
         cmdSaldoCh.Top = 5030 - cmdSaldoCh.Height
         cmdRegulariza.Top = 5030 - cmdRegulariza.Height
         cmdSalir.Top = 5030 - cmdSalir.Height
         chkTodo.Top = 5030 - cmdSalir.Height
         cmdRendicion.Visible = True
         lsDocTpoRecibo = oOperacion.EmiteDocOpe(gsOpeCod, ObligatorioDebeExistir, Digitado)
         If lsDocTpoRecibo = "" Then
            MsgBox "No se asignó Tipo de Documento Recibo de A rendir a Operación", vbCritical, "Error"
            lSalir = True
            Exit Sub
         End If
         cmdSaldoCh.Visible = IIf(lnTipoArendir = gArendirTipoCajaChica, True, False)
         fgAtenciones.ColWidth(2) = fgAtenciones.ColWidth(2) - 300
         fgAtenciones.ColWidth(3) = fgAtenciones.ColWidth(3) + 300
End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set oContFunc = Nothing
Set oNArendir = Nothing
Set oAreas = Nothing
End Sub
Public Property Get sPendiente() As String
sPendiente = sCtaPendiente
End Property
Public Property Let sPendiente(ByVal vNewValue As String)
sCtaPendiente = sPendiente
End Property
Private Sub txtBuscarAgenciaArea_EmiteDatos()
lblAgenciArea = oAreas.GetNombreAreas(Mid(txtBuscarAgenciaArea, 1, 3))
lblAgeDesc = oAreas.GetNombreAgencia(Mid(txtBuscarAgenciaArea, 4, 2))
If chkTodo.Visible Then
    chkTodo.SetFocus
Else
    cmdProcesar.SetFocus
End If
End Sub

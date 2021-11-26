VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmContabMigracion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Migración de Asientos Contables Agencias"
   ClientHeight    =   3090
   ClientLeft      =   4305
   ClientTop       =   3420
   ClientWidth     =   6255
   Icon            =   "frmContabMigracion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6255
   Begin VB.CommandButton cmdMigracion 
      Caption         =   "Migracion SIAFC"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Height          =   795
      Left            =   120
      TabIndex        =   6
      Top             =   990
      Width           =   5895
      Begin MSMask.MaskEdBox txtFecha2 
         Height          =   345
         Left            =   4560
         TabIndex        =   3
         Top             =   270
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   345
         Left            =   2310
         TabIndex        =   2
         Top             =   270
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label lblFecha 
         Alignment       =   2  'Center
         Caption         =   "Fecha de Proceso         DEL"
         Height          =   255
         Left            =   90
         TabIndex        =   8
         Top             =   345
         Width           =   2025
      End
      Begin VB.Label lblFecha2 
         Alignment       =   2  'Center
         Caption         =   "AL"
         Height          =   255
         Left            =   3960
         TabIndex        =   7
         Top             =   345
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   3750
      TabIndex        =   4
      Top             =   1920
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   4830
      TabIndex        =   5
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Caption         =   "Agencia"
      Height          =   795
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txtAgeDesc 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1560
         TabIndex        =   1
         Top             =   300
         Width           =   4125
      End
      Begin Sicmact.TxtBuscar txtAgeCod 
         Height          =   345
         Left            =   150
         TabIndex        =   0
         Top             =   300
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
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
      End
   End
   Begin VB.Label lblMsg 
      Caption         =   " Procesando ...por favor espere un momento"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   150
      TabIndex        =   10
      Top             =   1980
      Visible         =   0   'False
      Width           =   3525
   End
End
Attribute VB_Name = "frmContabMigracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql As String
Dim rs As New ADODB.Recordset
Dim lTransActiva As Boolean
Dim aAsientoS() As String
Dim aAsientoD() As String
Dim aAsientoV() As String
'Dim aAsientoZ() As String
Dim sObjetoCod  As String
Dim lOk         As Boolean
Dim sAsientoT   As String
Dim oCon        As DConecta
Dim oConN       As DConecta
Dim lsDts As String
Dim lsDtsa As String
Dim lsDtsb As String
Dim lsDtsc As String
Dim s As Variant
Dim X As Variant
Dim Y As Variant
Dim z As Variant

Private Sub cmdAceptar_Click()
Dim dFecha As Date
Dim nPos   As Integer
'By Capi 04032008
Dim lrs As ADODB.Recordset

sAsientoT = "AsientoDN"

If txtAgeCod = "" Then
   MsgBox "No se definió Agencia...!", vbInformation, "Aviso"
   Exit Sub
End If
If txtFecha > txtFecha2 Then
   MsgBox "Fecha inicial debe ser menor que Fecha Final", vbInformation, "Aviso"
   Exit Sub
End If
Me.Enabled = False
MousePointer = 11

Set oCon = New DConecta
Set oConN = New DConecta

oCon.AbreConexion
If gbBitCentral Then
    oConN.AbreConexion
Else
    If Not oConN.AbreConexion Then 'Remota(Right(txtAgeCod, 2), True)
       MousePointer = 0
       Me.Enabled = True
       Exit Sub
    End If
End If
lblMsg.Visible = True

For dFecha = CDate(txtFecha) To CDate(txtFecha2)
   DoEvents
   ReDim aAsientoS(1 To 3, 0 To 0)
   ReDim aAsientoD(1 To 5, 0 To 0)
   ReDim aAsientoV(1 To 5, 0 To 0)

   
   gdFecha = dFecha
   sSql = "SELECT * FROM AsientoValida WHERE dAsientoModif = (" _
        & "SELECT MAX(dAsientoModif) FROM AsientoValida WHERE datediff(dd,dAsientoFecha,'" & Format(gdFecha, "mm/dd/yyyy") & "')=0 and cAsientoTipo = '" & IIf(sAsientoT = "ASIENTOD", 1, 2) & "')"
   Set rs = oConN.CargaRecordSet(sSql)
   If Not rs.EOF Then
      If rs!cAsientoEstado = "0" Then
         MsgBox "Asiento Generado posee Observaciones. " & Chr(10) & "Por favor solicitar nueva generación de Asiento", vbInformation, "Aviso"
         Me.Enabled = True
         Exit Sub
      End If
   End If
          
   sSql = " SELECT cCtaCnt, cTipo, Round(SUM(nDebe),2) as nDebe, Round(SUM(nHaber),2) as nHaber " _
        & " FROM " & sAsientoT & " WHERE   datediff(dd,dfecha,'" & Format(dFecha, "mm/dd/yyyy") & "')= 0 " _
        & " GROUP BY cCtaCnt, cTipo ORDER BY cCtaCnt "
   Set rs = oConN.CargaRecordSet(sSql)
   If rs.EOF Then
      MsgBox "No se registraron Movimientos en Agencia el día " & dFecha & "...!", vbCritical, "Error"
   Else
      'By Capi 04032008
      Dim lbTodosCuadran As Boolean
      sSql = " SELECT cTipo, Round(SUM(nDebe),2) as nDebe, Round(SUM(nHaber),2) as nHaber " _
        & " FROM " & sAsientoT & " WHERE   datediff(dd,dfecha,'" & Format(dFecha, "mm/dd/yyyy") & "')= 0 " _
        & " GROUP BY cTipo Order By cTipo "
      Set lrs = oConN.CargaRecordSet(sSql)
      lbTodosCuadran = True
      Do While Not lrs.EOF
        If lrs!nDebe <> lrs!nHaber And lrs!cTipo <> "2" Then
            MsgBox "Debe " & lrs!nDebe & " y Haber " & lrs!nHaber & " no Cuadran...Verifique Asiento Tipo " & lrs!cTipo
            lbTodosCuadran = False
        End If
        lrs.MoveNext
      Loop
      If Not lbTodosCuadran Then Exit For
      'End By
      Do While Not rs.EOF
         Select Case rs!cTipo
         Case "0"
            If Mid(rs!cCtaCnt, 3, 1) = "1" Then
               ReDim Preserve aAsientoS(1 To 3, 0 To UBound(aAsientoS, 2) + 1)
               nPos = UBound(aAsientoS, 2)
               aAsientoS(1, nPos) = rs!cCtaCnt
               aAsientoS(2, nPos) = Round(rs!nDebe, 2)
               aAsientoS(3, nPos) = Round(rs!nHaber, 2)
            Else
               nPos = BuscaMatriz(aAsientoD, rs!cCtaCnt, 2)
               If nPos = -1 Then
                  ReDim Preserve aAsientoD(1 To 5, 0 To UBound(aAsientoD, 2) + 1)
                  nPos = UBound(aAsientoD, 2)
               End If
               aAsientoD(1, nPos) = rs!cCtaCnt
               aAsientoD(4, nPos) = Round(rs!nDebe, 2)
               aAsientoD(5, nPos) = Round(rs!nHaber, 2)
            End If
         Case "1"
            nPos = BuscaMatriz(aAsientoV, rs!cCtaCnt, 2)
            If nPos = -1 Then
               ReDim Preserve aAsientoV(1 To 5, 0 To UBound(aAsientoV, 2) + 1)
               nPos = UBound(aAsientoV, 2)
            End If
            aAsientoV(1, nPos) = rs!cCtaCnt
            aAsientoV(2, nPos) = Round(rs!nDebe, 2)
            aAsientoV(3, nPos) = Round(rs!nHaber, 2)
         Case "2"
            nPos = BuscaMatriz(aAsientoV, rs!cCtaCnt, 2)
            If nPos = -1 Then
               ReDim Preserve aAsientoV(1 To 5, 0 To UBound(aAsientoV, 2) + 1)
               nPos = UBound(aAsientoV, 2)
            End If
            aAsientoV(1, nPos) = rs!cCtaCnt
            aAsientoV(4, nPos) = Round(rs!nDebe, 2)
            aAsientoV(5, nPos) = Round(rs!nHaber, 2)
         Case "3"
            nPos = BuscaMatriz(aAsientoD, rs!cCtaCnt, 2)

            If nPos = -1 Then
               ReDim Preserve aAsientoD(1 To 5, 0 To UBound(aAsientoD, 2) + 1)
               nPos = UBound(aAsientoD, 2)
            End If
            aAsientoD(1, nPos) = rs!cCtaCnt
            aAsientoD(2, nPos) = Round(rs!nDebe, 2)
            aAsientoD(3, nPos) = Round(rs!nHaber, 2)
            
         End Select
         rs.MoveNext

      Loop
      lOk = True
      If UBound(aAsientoS, 2) > 0 Then
         GrabaAsientoMigra aAsientoS, True, , 1
         If Not lOk Then
            lblMsg.Visible = False
            Exit Sub
         End If
      End If
      If UBound(aAsientoD, 2) > 0 Then
         GrabaAsientoMigra aAsientoD, False, , 2
         If Not lOk Then
            lblMsg.Visible = False
            Exit Sub
         End If
      End If

      If UBound(aAsientoV, 2) > 0 Then
         GrabaAsientoMigra aAsientoV, False, "Compra - Venta de M.E.", 3
      End If
   End If
   RSClose rs
Next
oCon.CierraConexion
oConN.CierraConexion
Set oCon = Nothing
Set oConN = Nothing

lblMsg.Visible = False
MousePointer = 0
Me.Enabled = True
If lbTodosCuadran Then
    MsgBox "Proceso Terminado...!", vbInformation, "Aviso"
Else
    MsgBox "Proceso Cancelado...!", vbInformation, "Aviso"
End If
End Sub

Private Sub GrabaAsientoMigra(paAsiento() As String, plMN As Boolean, Optional psMsg As String = "", Optional pnTipo As Integer = 0)
Dim N As Integer
Dim sCtaCod As String
Dim nItem As Integer
On Error GoTo ErrMigra

gsMovNro = Format(gdFecha, "yyyymmdd") & String(6, "0") & gsCodCMAC & Right(txtAgeCod, 2) & "00XXX" & pnTipo
If Len(gsMovNro) <> 25 Then
   MsgBox "Error en definición de Agencia...!", vbInformation, "Aviso"
   Exit Sub
End If
sSql = "SELECT cMovNro FROM Mov WHERE cMovNro like '" & Mid(gsMovNro, 1, 19) & "__" & Right(gsMovNro, 4) & "' and nMovEstado = " & gMovEstContabMovContable & " and nMovFlag <> " & gMovFlagEliminado & " "
Set rs = oCon.CargaRecordSet(sSql)
If Not rs.EOF Then
   MsgBox "Ya se realizó la Migración de Asiento de " & txtAgeDesc & Chr(10) & "Por favor verificar...!", vbInformation, "Advertencia"
   Exit Sub
End If
If pnTipo = 3 Then
   gsOpeCod = "701108"
Else
   gsOpeCod = "701107"
End If
Dim oMov As New DMov
gsMovNro = oMov.GeneraMovNro(, , , gsMovNro)
gsGlosa = "Asiento Contable Consolidado de Operaciones en " & IIf(plMN, "M.N.", "M.E.") & " del " & gdFecha & " " & psMsg
lTransActiva = True

oMov.BeginTrans      'Iniciamos Transaccion
oMov.InsertaMov gsMovNro, gsOpeCod, gsGlosa, gMovEstContabMovContable, gMovFlagVigente
gnMovNro = oMov.GetnMovNro(gsMovNro)
nItem = 0
For N = 1 To UBound(paAsiento, 2)
   'Verificamos la existencia de Cuenta Contable
   sCtaCod = VerificaCuenta(paAsiento(1, N))
   'Grabamos MovCta
   If sCtaCod <> "" Then
      If Val(paAsiento(2, N)) > 0 Then
         nItem = nItem + 1
         oMov.InsertaMovCta gnMovNro, nItem, sCtaCod, nVal(paAsiento(2, N))
         If sObjetoCod <> "" Then  'Grabamos MovObj
            oMov.InsertaMovObj gnMovNro, nItem, "1", Format(TpoObjetos.ObjEntidadesFinancieras, "00")
            oMov.InsertaMovObjIF gnMovNro, nItem, "1", Mid(sObjetoCod, 4, 13), Left(sObjetoCod, 2), Mid(sObjetoCod, 18, 7)
         End If
         If Not plMN Then
            If Val(paAsiento(4, N)) > 0 Then
                oMov.InsertaMovMe gnMovNro, nItem, nVal(paAsiento(4, N))
            End If
         End If
      End If
      If Val(paAsiento(3, N)) > 0 Then
         nItem = nItem + 1
         oMov.InsertaMovCta gnMovNro, nItem, sCtaCod, nVal(paAsiento(3, N)) * -1
         If sObjetoCod <> "" Then  'Grabamos MovObj
            oMov.InsertaMovObj gnMovNro, nItem, "1", Format(TpoObjetos.ObjEntidadesFinancieras, "00")
            oMov.InsertaMovObjIF gnMovNro, nItem, "1", Mid(sObjetoCod, 4, 13), Left(sObjetoCod, 2), Mid(sObjetoCod, 18, 7)
         End If
         If Not plMN Then
            If Val(paAsiento(5, N)) > 0 Then
                oMov.InsertaMovMe gnMovNro, nItem, nVal(paAsiento(5, N)) * -1
            End If
         End If
      End If
   Else
      MsgBox "Cuenta Contable " & paAsiento(1, N) & " del día " & gdFecha & " Tipo " & Format(pnTipo, "00") & " no existe. Por favor verificar", vbInformation, "Error"
      oMov.RollbackTrans
      lTransActiva = False
      lOk = False
      Exit Sub
   End If
Next
oMov.ActualizaSaldoMovimiento gsMovNro, "+"
oMov.CommitTrans
lTransActiva = False
Exit Sub

ErrMigra:
   If lTransActiva Then
      oMov.RollbackTrans
      lTransActiva = False
   End If
   MsgBox TextErr(Err.Description), vbInformation, "Error"
End Sub

Private Function VerificaCuenta(sCtaCod As String) As String
Dim lsCta  As String
Dim CtaCor As String
Dim SubCta As String
Dim lrs    As New ADODB.Recordset
lsCta = sCtaCod
sObjetoCod = ""
sSql = "SELECT cCtaContCod, MAX(cIFTpo+'.'+cPersCod+'.'+cCtaIFCod) as cObjetoCod FROM CtaIFFiltro WHERE bUsoAgencia = 1 and cCtaContCod + cCtaIFSubCta = '" & sCtaCod & "' GROUP BY cCtaContCod "
Set lrs = oCon.CargaRecordSet(sSql)
If Not lrs.EOF Then
   sObjetoCod = lrs!cObjetoCod
End If
RSClose lrs
VerificaCuenta = lsCta
End Function

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdMigracion_Click()

'    frmContabMigraAsientoCont.Show 1
MousePointer = 11
Me.Enabled = False

'Inserta en la tabla Financ_Asientos
' clsMov.InsertaAsientoRuta (Text1.Text)

'LLAMAR EL DTS migracion local
'dtsrun /Sserver_name /Uuser_nName /Ppassword /Npackage_name /Mpackage_password
'lsDts = "dtsrun /S01SRVSICMAC01 /Usa /Pcmacica /NDTS_IMPORTAFINAC /M"
's = Shell(lsDts, vbMaximizedFocus)
'Migra Asientos
'lsDtsa = "dtsrun /S01SRVSICMAC01 /Usa /Pcmacica /NFINANC_Migra_Asientos_18022004 /M"
'x = Shell(lsDtsa, vbMaximizedFocus)
'Migra Cuentas
'lsDtsb = "dtsrun /S01SRVSICMAC01 /Usa /Pcmacica /NFINANC_Migra_Cuentas_18022004 /M"
'y = Shell(lsDtsb, vbMaximizedFocus)
'Migra Saldos
'lsDtsc = "dtsrun /S01SRVSICMAC01 /Usa /Pcmacica /NFINANC_Migra_saldos_18022004 /M"
'z = Shell(lsDtsc, vbMaximizedFocus)

'Llamar DTS Migracion Red

'dtsrun /Sserver_name /Uuser_nName /Ppassword /Npackage_name /Mpackage_password
lsDts = "dtsrun /S01SRVSICMAC01 /Usa /Pcmacica /NDTS_IMPORTAFINAC /M"
s = Shell(lsDts, vbMaximizedFocus)
'Migra Asientos
lsDtsa = "dtsrun /S01SRVSICMAC01 /Usa /Pcmacica /NFINANC_Migra_Asientos_28022004 /M"
X = Shell(lsDtsa, vbMaximizedFocus)
'Migra Cuentas
lsDtsb = "dtsrun /S01SRVSICMAC01 /Usa /Pcmacica /NFINANC_Migra_Cuentas_28022004 /M"
Y = Shell(lsDtsb, vbMaximizedFocus)
'Migra Saldos
lsDtsc = "dtsrun /S01SRVSICMAC01 /Usa /Pcmacica /NFINANC_Migra_saldos_28022004 /M"
z = Shell(lsDtsc, vbMaximizedFocus)

MousePointer = 0
Me.Enabled = True

MsgBox "Proceso Terminado...!", vbInformation, "Aviso"
    
End Sub

Private Sub Form_Load()
    CentraForm Me
    lTransActiva = False
    txtFecha = gdFecSis
    txtFecha2 = gdFecSis
    Dim clsRHArea As New DActualizaDatosArea
    txtAgeCod.rs = clsRHArea.GetAgencias
    Set clsRHArea = Nothing
End Sub

Private Sub txtAgeCod_EmiteDatos()
txtAgeDesc = txtAgeCod.psDescripcion
If txtAgeDesc <> "" Then
    txtFecha.SetFocus
End If
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtFecha2.SetFocus
End If
End Sub

Private Sub txtFecha_LostFocus()
If ValidaFecha(txtFecha) <> "" Then
   MsgBox "Fecha no válida...!", vbInformation, "Error"
   txtFecha.SetFocus
End If
End Sub

Private Sub txtFecha2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdAceptar.SetFocus
End If
End Sub

Private Sub txtFecha2_LostFocus()
If ValidaFecha(txtFecha2) <> "" Then
   MsgBox "Fecha no válida...!", vbInformation, "Error"
   txtFecha2.SetFocus
End If
End Sub

Private Function BuscaMatriz(paM() As String, psDato As String, Optional pnDimen As Integer = 1) As Integer
Dim N As Integer
BuscaMatriz = -1
For N = LBound(paM, pnDimen) To UBound(paM, pnDimen)
   If paM(1, N) = psDato Then
      BuscaMatriz = N
      Exit Function
   End If
Next
End Function

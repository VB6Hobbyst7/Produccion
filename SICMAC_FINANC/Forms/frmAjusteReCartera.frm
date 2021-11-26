VERSION 5.00
Begin VB.Form frmAjusteReCartera 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Colocaciones: Cartera: Reclasificación"
   ClientHeight    =   3210
   ClientLeft      =   5340
   ClientTop       =   4005
   ClientWidth     =   4410
   Icon            =   "frmAjusteReCartera.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4410
   Begin VB.Frame Frame1 
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   795
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4170
      Begin VB.ComboBox CboMes 
         Height          =   315
         ItemData        =   "frmAjusteReCartera.frx":030A
         Left            =   690
         List            =   "frmAjusteReCartera.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   315
         Width           =   1830
      End
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   1
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   165
         TabIndex        =   11
         Top             =   390
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   2640
         TabIndex        =   10
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame frmMoneda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   720
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   4170
      Begin VB.OptionButton optMoneda 
         Caption         =   "A&justado"
         Height          =   255
         Index           =   3
         Left            =   4500
         TabIndex        =   7
         Top             =   330
         Width           =   1005
      End
      Begin VB.TextBox txtTipCambio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2010
         MaxLength       =   16
         TabIndex        =   2
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Cambio :"
         Height          =   315
         Left            =   615
         TabIndex        =   8
         Top             =   300
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdAsiento 
      Caption         =   "Grabar &Asiento Contable"
      Enabled         =   0   'False
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
      Left            =   135
      TabIndex        =   3
      Top             =   2280
      Width           =   4155
   End
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
      Left            =   135
      TabIndex        =   5
      Top             =   2700
      Width           =   4155
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar Cuadro de Reclasificación"
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
      Left            =   135
      TabIndex        =   4
      Top             =   1860
      Width           =   4155
   End
End
Attribute VB_Name = "frmAjusteReCartera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lTransActiva As Boolean
Dim WithEvents oImp As NContImprimir
Attribute oImp.VB_VarHelpID = -1
Dim oBarra As New clsProgressBar

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Private Sub GeneraReporte(psClase As String)
Dim nMes As Integer
Dim nAnio As Integer
Dim dFecha As Date
Dim sImpre As String
Dim oCont As New NContFunciones
On Error GoTo ReCarteraErr
nMes = CboMes.ListIndex + 1
nAnio = txtAnio
dFecha = DateAdd("m", 1, "01/" & Format(nMes, "00") & "/" & Format(txtAnio, "0000")) - 1

Me.Enabled = False
sImpre = oImp.ImprimeCuadroReclasificacion("C", dFecha, CInt(Mid(gsOpeCod, 3, 1)), psClase, nVal(txtTipCambio), gnLinPage)
EnviaPrevio sImpre, "CUADRO DE RECLASIFICACION DE CARTERA DE COLACIONES", gnLinPage, False
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Genero el Cuadro de Reclasificacion al Cierre de " & dFecha & " con el Tipo de Cambio : " & txtTipCambio.Text
            Set objPista = Nothing
            '*******
Set oCont = Nothing
Me.Enabled = True
cmdAsiento.Enabled = True
cmdAsiento.SetFocus
Exit Sub
ReCarteraErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso¡!"
End Sub

Private Sub CboMes_Click()
txtTipCambio = TipoCambioCierre(nVal(txtAnio), CboMes.ListIndex + 1, False)
End Sub

Private Sub CboMes_Validate(Cancel As Boolean)
If CboMes.ListIndex <> Val(CboMes.Tag) Then
   cmdAsiento.Enabled = False
   CboMes.Tag = CboMes.ListIndex
End If
End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtTipCambio = TipoCambioCierre(nVal(txtAnio), CboMes.ListIndex + 1, False)
   txtAnio.SetFocus
End If
End Sub

Private Sub cmdAsiento_Click()
Dim rs As ADODB.Recordset
Dim nTotDif As Currency

Dim nMes     As Integer
Dim nAnio    As Integer
Dim dFecha   As Date
Dim psClase  As String
Dim sAsiento As String
Dim nItem    As String

On Error GoTo AsientoErr

If MsgBox("¿ Seguro que desea generar Asiento de Reclasificación ? ", vbQuestion + vbYesNo, "!Confirmación!") = vbNo Then
   Exit Sub
End If
psClase = "14"
nMes = CboMes.ListIndex + 1
nAnio = txtAnio
dFecha = DateAdd("m", 1, "01/" & Format(nMes, "00") & "/" & Format(txtAnio, "0000")) - 1

Dim oCont As New NContFunciones
Dim oMov  As New DMov
Dim oAju  As New DAjusteCont

If Not oCont.PermiteModificarAsiento(Format(dFecha, gsFormatoMovFecha), False) Then
   MsgBox "Mes ya cerrado. Imposible generar Asiento de Reclasificación", vbInformation, "!Aviso!"
   Exit Sub
End If

gsMovNro = Format(dFecha, "yyyymmdd")
If oCont.ExisteMovimiento(gsMovNro, gsOpeCod) Then
   MsgBox "Asiento de Reclasificación de Cartera ya fue Generado!", vbInformation, "¡Aviso!"
   Exit Sub
End If

Me.Enabled = False
Set rs = oAju.AjusteReclasificaCartera(Format(dFecha, gsFormatoFecha), CInt(Mid(gsOpeCod, 3, 1)), psClase, nVal(txtTipCambio))
If rs.EOF Then
   MsgBox "No existen diferencias entre Estadísticas y Saldos Contables ", vbInformation, "!Aviso!"
Else
   oImp_BarraShow rs.RecordCount
   
   oMov.BeginTrans
   gsGlosa = "Asiento de Reclasificación de Cartera al " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
   gsMovNro = oMov.GeneraMovNro(dFecha, gsCodAge, gsCodUser)
   lTransActiva = True
   oMov.InsertaMov gsMovNro, gsOpeCod, gsGlosa, gMovEstContabMovContable, gMovFlagVigente
   gnMovNro = oMov.GetnMovNro(gsMovNro)
   nItem = 0
   Do While Not rs.EOF
        If rs!nSaldo - rs!nCtaSaldoImporte <> 0 Then
            nItem = nItem + 1
            oMov.InsertaMovCta gnMovNro, nItem, IIf(IsNull(rs!Cta1), rs!Cta2, rs!Cta1), rs!nSaldo - rs!nCtaSaldoImporte
            If Mid(gsOpeCod, 3, 1) = "2" Then
                If Round((rs!nSaldo - rs!nCtaSaldoImporte) / nVal(Me.txtTipCambio), 3) <> 0 Then
                    oMov.InsertaMovMe gnMovNro, nItem, Round((rs!nSaldo - rs!nCtaSaldoImporte) / nVal(Me.txtTipCambio), 3)
                End If
            End If
        End If
        oImp_BarraProgress rs.Bookmark, "ASIENTO DE RECLASIFICACION", "", "Grabando...", vbBlue
        rs.MoveNext
   Loop
   If dFecha < gdFecSis Then
      oMov.ActualizaSaldoMovimiento gsMovNro, "+"
   End If
   oMov.CommitTrans
   lTransActiva = False
   oImp_BarraClose
End If
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaReclasifiCartera
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Se Grabo Asiento Contable al Cierre de " & dFecha & " con el Tipo de Cambio : " & txtTipCambio.Text
            Set objPista = Nothing
RSClose rs
Me.Enabled = False
ImprimeAsientoContable gsMovNro, , , , , , , , , , , , 1, "ASIENTO DE RECLASIFICACION DE CARTERA"
Set oMov = Nothing
Set oCont = Nothing
Set oAju = Nothing
Me.Enabled = True

Exit Sub
AsientoErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
   Me.Enabled = True
   If lTransActiva Then
        oMov.RollbackTrans
   End If
End Sub

Private Sub cmdGenerar_Click()
If Not ValidaDatos Then
   Exit Sub
End If
If MsgBox("¿ Seguro que desea generar Cuadro de Reclasificación ? ", vbQuestion + vbYesNo, "!Confirmación!") = vbNo Then
   Exit Sub
End If
GeneraReporte "14"
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
frmOperaciones.Enabled = False
Set oImp = New NContImprimir
CboMes.ListIndex = Month(gdFecSis) - 1
txtAnio = Year(gdFecSis)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oImp = Nothing
frmOperaciones.Enabled = True
End Sub

Private Sub txtAnio_Change()
cmdAsiento.Enabled = False
End Sub

Private Sub txtAnio_GotFocus()
fEnfoque txtAnio
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtTipCambio = TipoCambioCierre(txtAnio, CboMes.ListIndex + 1, False)
   txtTipCambio.SetFocus
End If
End Sub

Private Sub txtTipCambio_GotFocus()
fEnfoque txtTipCambio
End Sub

Private Sub txtTipCambio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTipCambio, KeyAscii, 14, 5)
If KeyAscii = 13 Then
   CmdGenerar.SetFocus
End If
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If CboMes.ListIndex = -1 Then
   MsgBox "Debe seleccionarse mes de proceso", vbInformation, "!Aviso!"
   CboMes.SetFocus
   Exit Function
End If
If Not ValidaAnio(txtAnio) Then
   txtAnio.SetFocus
   Exit Function
End If
If nVal(txtTipCambio) = 0 Then
   MsgBox "Debe indicar Tipo de Cambio", vbInformation, "!Aviso!"
   txtTipCambio.SetFocus
   Exit Function
End If
ValidaDatos = True
End Function

Private Sub oImp_BarraClose()
oBarra.CloseForm Me
End Sub

Private Sub oImp_BarraProgress(value As Variant, psTitulo As String, psSubTitulo As String, psTituloBarra As String, ColorLetras As ColorConstants)
oBarra.Progress value, psTitulo, psSubTitulo, psTituloBarra, ColorLetras
End Sub

Private Sub oImp_BarraShow(pnMax As Variant)
oBarra.ShowForm Me
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.Max = pnMax
End Sub

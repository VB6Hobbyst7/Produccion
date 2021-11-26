VERSION 5.00
Begin VB.Form frmAjusteIntDevengado 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4440
   ClientLeft      =   4755
   ClientTop       =   3540
   ClientWidth     =   4440
   Icon            =   "frmAjusteIntDevengado.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   4440
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Mostrar Comparación Estadística - Contable"
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
      TabIndex        =   11
      Top             =   3270
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
      TabIndex        =   10
      Top             =   4040
      Width           =   4155
   End
   Begin VB.CommandButton cmdAsiento 
      Caption         =   "&Grabar Asiento de Ajuste"
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
      TabIndex        =   9
      Top             =   3650
      Width           =   4155
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
      Height          =   2190
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   4170
      Begin VB.TextBox txtTipCambioCompra 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2220
         MaxLength       =   16
         TabIndex        =   15
         Top             =   1440
         Width           =   1425
      End
      Begin VB.TextBox txtTipCambioVenta 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2220
         MaxLength       =   16
         TabIndex        =   12
         Top             =   840
         Width           =   1425
      End
      Begin VB.TextBox txtTipCambio 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2220
         MaxLength       =   16
         TabIndex        =   7
         Top             =   240
         Width           =   1425
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "A&justado"
         Height          =   255
         Index           =   3
         Left            =   4500
         TabIndex        =   6
         Top             =   330
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Cambio Venta"
         Height          =   315
         Left            =   360
         TabIndex        =   14
         Top             =   960
         Width           =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Cambio Compra"
         Height          =   315
         Left            =   360
         TabIndex        =   13
         Top             =   1560
         Width           =   1920
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Cambio Fijo"
         Height          =   315
         Left            =   345
         TabIndex        =   8
         Top             =   300
         Width           =   1680
      End
   End
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
      TabIndex        =   0
      Top             =   90
      Width           =   4170
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   2
         Top             =   300
         Width           =   855
      End
      Begin VB.ComboBox CboMes 
         Height          =   315
         ItemData        =   "frmAjusteIntDevengado.frx":030A
         Left            =   690
         List            =   "frmAjusteIntDevengado.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   315
         Width           =   1830
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   165
         TabIndex        =   3
         Top             =   390
         Width           =   390
      End
   End
   Begin VB.CommandButton cmdGenerarDet 
      Caption         =   "M&ostrar Estadística Detallado"
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
      TabIndex        =   16
      Top             =   3650
      Visible         =   0   'False
      Width           =   4155
   End
End
Attribute VB_Name = "frmAjusteIntDevengado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnTipoAsiento As Integer
Dim aAsiento() As String
Dim nCta As Integer
Dim dFecha As Date
Dim sCtaDebe  As String
Dim sCtaHaber As String

Dim WithEvents oImp As NContImprimir
Attribute oImp.VB_VarHelpID = -1
Dim WithEvents oImp1 As NContImpreReg '*** PEAC 20130111
Attribute oImp1.VB_VarHelpID = -1

Dim oBarra As New clsProgressBar
Dim rsAjusteContab As ADODB.Recordset 'PASI20170424
Dim oAjusteCont As DAjusteCont 'PASI20170424
'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub inicio(pnTipoAsiento As Integer)
lnTipoAsiento = pnTipoAsiento
Me.Show 0, frmMdiMain

End Sub

Private Sub GeneraReporte()
Dim nMes As Integer
Dim nAnio As Integer
Dim dFecha As Date
Dim sImpre As String
Dim oCont As New NContFunciones
Dim psRep As String 'NAGL 202102
On Error GoTo ErrGeneraReporte

nMes = cboMes.ListIndex + 1
nAnio = txtAnio
dFecha = DateAdd("m", 1, "01/" & Format(nMes, "00") & "/" & Format(txtAnio, "0000")) - 1
psRep = "" 'NAGL 202102
If Not oCont.PermiteModificarAsiento(Format(dFecha, gsFormatoMovFecha), False) Then
   Set oCont = Nothing
   MsgBox "Mes ya cerrado. Imposible generar Cuadro de Comparación", vbInformation, "!Aviso!"
   Exit Sub
End If

If lnTipoAsiento = 3 Then
    gsMovNro = Format(dFecha, "yyyymmdd")
    If Not oCont.ExisteMovProvEspecificaCOVID(gsMovNro, Mid(gsOpeCod, 3, 1)) Then
        Set oCont = Nothing
        If MsgBox("Aún no se ha generado la provisión de cartera reprogramada Covid en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & ", Desea continuar?", vbInformation + vbYesNo, "Atención") = vbNo Then
            Exit Sub
        End If
    End If
ElseIf lnTipoAsiento = 34 Then
    psRep = "RepGen"
End If 'NAGL 202102 Según ACTA N°017-2021

Me.Enabled = False
If lnTipoAsiento = 11 Then
   'ALPA 20120402****************
    sImpre = oImp.ImprimeCuadroReclasificacion(IIf(lnTipoAsiento = 1, "D", IIf(lnTipoAsiento = 2, "S", IIf(lnTipoAsiento = 3, "P", Trim(Str(lnTipoAsiento))))), dFecha, CInt(Mid(gsOpeCod, 3, 1)), sCtaHaber, nVal(txtTipCambio), gnLinPage, sCtaHaber)
    'sImpre = oImp.ImprimeCuadroReclasificacion(IIf(lnTipoAsiento = 1, "D", IIf(lnTipoAsiento = 2, "S", IIf(lnTipoAsiento = 3, "P", Trim(Str(lnTipoAsiento))))), dFecha, CInt(Mid(gsOpeCod, 3, 1)), sCtaHaber, nVal(txtTipCambio), gnLinPage, sCtaHaber, nVal(txtTipCambioVenta), nVal(txtTipCambioCompra))
   '*****************************
Else
   'ALPA 20120402****************
    sImpre = oImp.ImprimeCuadroReclasificacion(IIf(lnTipoAsiento = 1, "D", IIf(lnTipoAsiento = 2, "S", IIf(lnTipoAsiento = 3, "P", Trim(Str(lnTipoAsiento))))), dFecha, CInt(Mid(gsOpeCod, 3, 1)), sCtaDebe, nVal(txtTipCambio), gnLinPage, sCtaDebe, psRep)
    'psRep NAGL 202102 Agregó
    'sImpre = oImp.ImprimeCuadroReclasificacion(IIf(lnTipoAsiento = 1, "D", IIf(lnTipoAsiento = 2, "S", IIf(lnTipoAsiento = 3, "P", Trim(Str(lnTipoAsiento))))), dFecha, CInt(Mid(gsOpeCod, 3, 1)), sCtaDebe, nVal(txtTipCambio), gnLinPage, sCtaDebe, nVal(txtTipCambioVenta), nVal(txtTipCambioCompra))
   '*****************************
    'MADM 20111026********************
    'If Trim(Mid(sImpre, (Len(sImpre) - 10), 10)) = "0.00" Then
    'ALPA20140130*******************
    If Trim(sImpre) = "" Then
        Exit Sub
    End If
    '*******************************
    If Trim(Mid(sImpre, (Len(sImpre) - 10), 10)) = "0.00" And lnTipoAsiento = 20 Then
    '*********************************
        MsgBox "No han habido cambios con respecto al Mes Anterior", vbInformation, "¡Aviso!"
        cmdAsiento.Enabled = False
        Set oCont = Nothing

        Exit Sub
    End If
End If
'PASI20170424 ********
Set oAjusteCont = New DAjusteCont
Select Case lnTipoAsiento
    Case 28
        Set rsAjusteContab = oAjusteCont.CargaAjusteDeContabilidadxCredito(lnTipoAsiento, dFecha, CInt(Mid(gsOpeCod, 3, 1)), nVal(txtTipCambio))
End Select
'PASI END*************

EnviaPrevio sImpre, "COMPARACION ESTADISTICA - CONTABLE DE INTERESES DE COLOCACIONES", gnLinPage, False

If lnTipoAsiento = 31 Then
    If MsgBox("Desea generar el Reporte Detallado de Cartera Reprogramada al " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & " ?", vbInformation + vbYesNo, "Atención") = vbYes Then
       Call ObtieneReporteDetalle(lnTipoAsiento) 'NAGL 20200902
    End If
End If 'NAGL 202007 Según ACTA N°049-2020
               
    'ARLO20170208
    Set objPista = New COMManejador.Pista
    'gsOpeCod = LogPistaMantDocumento
    objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", Me.Caption & "Se Mostro la Comparacion Estadistica-Contable al Cierre " & dFecha & " con el Tipo de Cambio Fijo : " & txtTipCambio.Text _
    & " Tpo Cambio venta : " & txtTipCambioVenta & " Tpo Cambio Venta : " & txtTipCambioCompra
    Set objPista = Nothing
    '*******
Set oCont = Nothing
Me.Enabled = True
cmdAsiento.Enabled = True
cmdAsiento.SetFocus
Exit Sub
ErrGeneraReporte:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub CboMes_Click()
'If CboMes.ListIndex > -1 And txtAnio <> "" Then
'    txtTipCambio = TipoCambioCierre(txtAnio, CboMes.ListIndex + 1, False)
'    txtTipCambio.SetFocus
'End If
If cboMes.ListIndex > -1 And txtAnio <> "" Then
    txtTipCambio = TiposCambiosCierreMensual(txtAnio, cboMes.ListIndex + 1, False, 1)
    txtTipCambioVenta = TiposCambiosCierreMensual(txtAnio, cboMes.ListIndex + 1, False, 2)
    txtTipCambioCompra = TiposCambiosCierreMensual(txtAnio, cboMes.ListIndex + 1, False, 3)
    txtTipCambioCompra.SetFocus
End If
End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    txtTipCambio = TipoCambioCierre(txtAnio, CboMes.ListIndex + 1, False)
'   txtAnio.SetFocus
'End If
If KeyAscii = 13 Then
    If cboMes.ListIndex > -1 And txtAnio <> "" Then
        txtTipCambio = TiposCambiosCierreMensual(txtAnio, cboMes.ListIndex + 1, False, 1)
        txtTipCambioVenta = TiposCambiosCierreMensual(txtAnio, cboMes.ListIndex + 1, False, 2)
        txtTipCambioCompra = TiposCambiosCierreMensual(txtAnio, cboMes.ListIndex + 1, False, 3)
        txtAnio.SetFocus
    End If
 End If
End Sub

Private Sub CboMes_Validate(Cancel As Boolean)
If cboMes.ListIndex <> Val(cboMes.Tag) Then
   cmdAsiento.Enabled = False
   cboMes.Tag = cboMes.ListIndex
End If
End Sub

Private Sub cmdAsiento_Click()
Dim rs       As ADODB.Recordset
Dim rsIntDev As ADODB.Recordset '*** PEAC 20130111
Dim nTotal   As Currency
Dim nItem    As Integer
Dim lTransActiva As Boolean
Dim nMes     As Integer, nAnio As Integer, dFecha As Date
Dim lsCtaGarantCF As String
Dim lsIntTipo As String
Dim lsContraCta1 As String
Dim lsTituloImpre As String 'EJVG20130322
Dim psRep As String 'NAGL 202102
On Error GoTo AsientoErr

If MsgBox("¿ Seguro desea Grabar Asiento ? ", vbQuestion + vbYesNo + vbDefaultButton2, "¡Confirmación¡") = vbNo Then
   Exit Sub
End If

nMes = cboMes.ListIndex + 1
nAnio = txtAnio
dFecha = DateAdd("m", 1, "01/" & Format(nMes, "00") & "/" & Format(txtAnio, "0000")) - 1
psRep = "" 'NAGL 202102

Dim oCont As New NContFunciones
Dim oMov  As New DMov
Dim oAju  As New DAjusteCont
Dim lnDiferenciaAjsute As Currency

If Not oCont.PermiteModificarAsiento(Format(dFecha, gsFormatoMovFecha), False) Then
   MsgBox "Mes ya cerrado. Imposible generar Asiento de Reclasificación", vbInformation, "!Aviso!"
   Exit Sub
End If
gsMovNro = Format(dFecha, "yyyymmdd")
If oCont.ExisteMovimiento(Left(gsMovNro, 6), gsOpeCod) Then
   MsgBox "Asiento ya generado", vbInformation, "¡Aviso!"
   Exit Sub
End If

If lnTipoAsiento = 3 Then
    If Not oCont.ExisteMovProvEspecificaCOVID(gsMovNro, Mid(gsOpeCod, 3, 1)) Then
        Set oCont = Nothing
        If MsgBox("Aún no se ha generado la provisión de cartera reprogramada Covid en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & ", Desea continuar?", vbInformation + vbYesNo, "Atención") = vbNo Then
            Exit Sub
        End If
    End If
ElseIf lnTipoAsiento = 34 Then
    psRep = "Asi"
End If 'NAGL 202102 Según ACTA N°017-2021

Me.Enabled = False

If lnTipoAsiento = 11 Then
    Set rs = oAju.AjusteInteresesColocaciones(lnTipoAsiento, Format(dFecha, gsFormatoFecha), CInt(Mid(gsOpeCod, 3, 1)), sCtaHaber, nVal(txtTipCambio), sCtaHaber)
ElseIf lnTipoAsiento = 32 Or lnTipoAsiento = 33 Then 'NAGL 202008 Según Acta N°063-2020
    Call GeneraAsientoOtrosProcesos(lnTipoAsiento)
    Exit Sub
Else
    If lnTipoAsiento = 7 Or lnTipoAsiento = 28 Then 'PASI 20170417 Tipo Asiento = 28
        Set rs = oAju.AjusteInteresesColocaciones(lnTipoAsiento, Format(dFecha, gsFormatoFecha), CInt(Mid(gsOpeCod, 3, 1)), sCtaDebe, nVal(txtTipCambio), sCtaDebe, True)
    Else
        Set rs = oAju.AjusteInteresesColocaciones(lnTipoAsiento, Format(dFecha, gsFormatoFecha), CInt(Mid(gsOpeCod, 3, 1)), sCtaDebe, nVal(txtTipCambio), sCtaDebe, , , , psRep)
        'NAGL 202102 Agregó psRep
        If lnTipoAsiento = 14 Then
            Set rsIntDev = oAju.AjusteIntDevCaptaciones(lnTipoAsiento, Format(dFecha, gsFormatoFecha), CInt(Mid(gsOpeCod, 3, 1)), sCtaDebe, nVal(txtTipCambio), sCtaDebe)
        End If
    End If
End If
If rs.EOF Then
   MsgBox "No existen diferencias entre Estadísticas y Saldos Contables ", vbInformation, "!Aviso!"
Else
    'ALPA 20150131*****************************
     If lnTipoAsiento = 19 Then
        If (nVal(Format(rs!nSaldo, "#.00")) - rs!nCtaSaldoImporte) = 0 Then
            MsgBox "No existen diferencias entre Estadísticas y Saldos Contables ", vbInformation, "!Aviso!"
'            lTransActiva = False
            oImp_BarraClose
            Exit Sub
        End If
     End If
    '*****************************************
   oImp_BarraShow rs.RecordCount
   oMov.BeginTrans
   lTransActiva = True
   
   If lnTipoAsiento = 7 Then
      gsGlosa = "ASIENTO DE REVERSION DE INTERESES DEVENGADOS AL " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
   ElseIf lnTipoAsiento = 3 Then
      gsGlosa = "ASIENTO DE PROVISION DE CARTERA AL " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
   ElseIf lnTipoAsiento = 5 Then
      gsGlosa = "ASIENTO DE CAPITAL DE CREDITOS CASTIGADOS AL " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
   ElseIf lnTipoAsiento = 6 Then
      gsGlosa = "ASIENTO DE INTERESES DE CREDITOS CASTIGADOS AL " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
   ElseIf lnTipoAsiento = 8 Then
      gsGlosa = "ASIENTO DE PROVISION DE CARTAS FIANZA AL " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
   ElseIf lnTipoAsiento = 9 Then
      gsGlosa = "ASIENTO DE CALIFICACION DE CREDITOS CONTINGENTES AL " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
   ElseIf lnTipoAsiento = 10 Then
      gsGlosa = "ASIENTO DE CREDITOS EN GARANTIA DE FINANCIAMIENTO AL " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
    ElseIf lnTipoAsiento = 11 Then
      gsGlosa = "ASIENTO DE CREDITOS DE PROVISION CARTAFIANZA AL " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
    ElseIf lnTipoAsiento = 12 Then
      gsGlosa = "ASIENTO DE AJUSTE POR CALIFICACION DE CARTERA " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
    ElseIf lnTipoAsiento = 13 Then
      gsGlosa = "ASIENTO DE AJUSTE DE RIESGO PONDERADO " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
    ElseIf lnTipoAsiento = 15 Then
      gsGlosa = "ASIENTO DE REVERSION DE PROVISIONES  " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
    'ALPA 20090529******************************************
    ElseIf lnTipoAsiento = 17 Then
      gsGlosa = "ASIENTO FINAL  " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
    'ALPA 20110425******************************************
    ElseIf lnTipoAsiento = 18 Then
      gsGlosa = "ASIENTO INTERESES DIFERIDOS PAGADOS " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
    ElseIf lnTipoAsiento = 19 Then
      gsGlosa = "ASIENTO DE GARANTIAS DE CARTAS FIANZAS  " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
    ElseIf lnTipoAsiento = 20 Then
      'gsGlosa = "ASIENTO DE SALDOS DE GARANTIAS DE CARTAS FIANZAS  " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
      gsGlosa = "ASIENTO DE SALDOS DE CARTAS FIANZAS  " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.") 'EJVG20130322
    ElseIf lnTipoAsiento = 21 Then
      gsGlosa = "ASIENTO INTERESES DIFERIDOS TRANSFERIDOS " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
    'JUEZ 20130116 *******************************************
    ElseIf lnTipoAsiento = 22 Then
      gsGlosa = "ASIENTO DE GASTOS DE CREDITOS CASTIGADOS AL " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
    'END JUEZ ************************************************
    ElseIf lnTipoAsiento = 23 Then 'EJVG20130307
      gsGlosa = "ASIENTO DE COMISION CARTA FIANZA " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
    ElseIf lnTipoAsiento = 24 Then 'ALPA20131129
      gsGlosa = "ASIENTO DE AJUSTE INTERES DIFERIDOS-SALDOS FINALES " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
    ElseIf lnTipoAsiento = 25 Then 'ALPA20140925
      gsGlosa = "ASIENTO DE AJUSTE COMPRA DE CARTERA " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
    ElseIf lnTipoAsiento = 26 Then 'ALPA20140925
      gsGlosa = "ASIENTO DE RECLASIFICACION DE PROVISIONES PROCICLICAS " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
    ElseIf lnTipoAsiento = 27 Then 'ALPA20150813
      gsGlosa = "ASIENTO DE RECLASIFICACION DE INTERESES DIFERIDOS DE CREDITOS VIGENTES" & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
     'NAGL 202007 *****************************************
    ElseIf lnTipoAsiento = 29 Then
     gsGlosa = "ASIENTO DE CIERRE FAE - MYPE " & dFecha & " EN " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
    ElseIf lnTipoAsiento = 30 Then
     gsGlosa = "ASIENTO DE CIERRE REACTIVA " & dFecha & " EN " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
    ElseIf lnTipoAsiento = 31 Then
     gsGlosa = "ASIENTO DE CIERRE CARTERA REPROGRAMADA " & dFecha & " EN " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
    'END NAGL 202007 Según ACTA N°049-2020******************
    'NAGL 202102 *****************************************
    ElseIf lnTipoAsiento = 34 Then
     gsGlosa = "ASIENTO DE CIERRE PROVISIONES DE CARTERA REPROGRAMADA COVID " & dFecha & " EN " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
    'END NAGL 202102 Según ACTA N°017-2021******************
    Else
      gsGlosa = "Asiento de " & Me.Caption & " al " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
    End If

    gsMovNro = oMov.GeneraMovNro(dFecha, gsCodAge, gsCodUser)
    oMov.InsertaMov gsMovNro, gsOpeCod, gsGlosa, gMovEstContabMovContable, gMovFlagVigente
    gnMovNro = oMov.GetnMovNro(gsMovNro)
    nItem = 0

    Dim lsCtaCod As String
    Dim lsCta As String
   
    Do While Not rs.EOF
'      lsContraCta2 = ""
      lsIntTipo = ""
      lsContraCta1 = ""
'      gnImporte2 = 0
      If lnTipoAsiento <> 7 And lnTipoAsiento <> 34 Then 'NAGL 202102 Agregó lnTipoAsiento <> 34
        lsCtaCod = IIf(IsNull(rs!Cta1), rs!Cta2, rs!Cta1)
                
        If lnTipoAsiento = 19 Then
            lsCtaGarantCF = "83" + Mid(lsCtaCod, 3, 1) + "1"

        ElseIf lnTipoAsiento = 20 Then
            lsCtaGarantCF = "72" + Mid(lsCtaCod, 3, 1) + "2"
        End If
      Else
        lsCtaCod = Trim(rs!Cta1)
      End If
      
      If lnTipoAsiento = 34 Then
         lsContraCta1 = rs!cCtaContraParte
         gnImporte = rs!nSaldoAsiento
      Else
        gnImporte = nVal(Format(rs!nSaldo, "#.00")) - rs!nCtaSaldoImporte
      End If 'NAGL 202102 Según ACTA N°017-2021
      
      If lnTipoAsiento = 18 Or lnTipoAsiento = 21 Or lnTipoAsiento = 27 Then
        'lsIntTipo = rs!cTipo
        lsContraCta1 = rs!cContraCuenta
        gnImporte = rs!nIntDiferido
      End If
      'ALPA 20140925*************************************
      If lnTipoAsiento = 25 Then
        'lsIntTipo = rs!cTipo
        lsContraCta1 = rs!Cta2
        gnImporte = (rs!nSaldo - rs!nCtaSaldoImporte) * -1
      End If
      'ALPA 20141206*************************************
      If lnTipoAsiento = 26 Then
        gnImporte = rs!nSaldo
      End If
      '**************************************************
'      If lsCtaCod = "14190402010601010137" Or lsCtaCod = "14190402012301010101" Or lsCtaCod = "141909010601010101" Then
'      MsgBox "1"
'      MsgBox "2"
'      End If
      lnDiferenciaAjsute = lnDiferenciaAjsute + gnImporte
      
      If gnImporte <> 0 Then
        nItem = nItem + 1
        'If lnTipoAsiento <> 17 Then
        If lnTipoAsiento = 11 Or lnTipoAsiento = 17 Or lnTipoAsiento = 20 Or lnTipoAsiento = 29 Or lnTipoAsiento = 30 Or lnTipoAsiento = 31 Or lnTipoAsiento = 34 Then
        'NAGL 202007 Agregó lnTipoAsiento (29,30,31)
        'NAGL 202102 Agregó lnTipoAsiento = 34
            If lnTipoAsiento = 11 Then
                'oMov.InsertaMovCta gnMovNro, nItem, lsCtaCod, Abs(gnImporte) * IIf(gnImporte > 0, -1, 1)
                oMov.InsertaMovCta gnMovNro, nItem, lsCtaCod, DevolverSaldoTC(lsCtaCod, nVal(txtTipCambio.Text), nVal(txtTipCambioVenta.Text), nVal(txtTipCambioCompra.Text), Abs(gnImporte) * IIf(gnImporte > 0, -1, 1), Mid(lsCtaCod, 3, 1))
            Else
                'oMov.InsertaMovCta gnMovNro, nItem, lsCtaCod, Abs(gnImporte) * IIf(gnImporte > 0, 1, -1)
                oMov.InsertaMovCta gnMovNro, nItem, lsCtaCod, DevolverSaldoTC(lsCtaCod, nVal(txtTipCambio.Text), nVal(txtTipCambioVenta.Text), nVal(txtTipCambioCompra.Text), Abs(gnImporte) * IIf(gnImporte > 0, 1, -1), Mid(lsCtaCod, 3, 1))
            End If
        Else
            'MADM 20110805 - 20
            'oMov.InsertaMovCta gnMovNro, nItem, lsCtaCod, gnImporte * IIf(lnTipoAsiento = 3 Or lnTipoAsiento = 8 Or lnTipoAsiento = 14 Or lnTipoAsiento = 19, -1, 1)
            'oMov.InsertaMovCta gnMovNro, nItem, lsCtaCod, gnImporte * IIf(lnTipoAsiento = 3 Or lnTipoAsiento = 8 Or lnTipoAsiento = 14 Or lnTipoAsiento = 19 Or lnTipoAsiento = 21, -1, 1)
            '        nItem = nItem + 1
            oMov.InsertaMovCta gnMovNro, nItem, lsCtaCod, DevolverSaldoTC(lsCtaCod, nVal(txtTipCambio.Text), nVal(txtTipCambioVenta.Text), nVal(txtTipCambioCompra.Text), gnImporte * IIf(lnTipoAsiento = 3 Or lnTipoAsiento = 8 Or lnTipoAsiento = 14 Or lnTipoAsiento = 19 Or lnTipoAsiento = 21 Or lnTipoAsiento = 23 Or lnTipoAsiento = 24 Or lnTipoAsiento = 25 Or lnTipoAsiento = 26, -1, 1), Mid(lsCtaCod, 3, 1))
        End If
        'End If
        If lnTipoAsiento <> 26 Then
        'ALPA 20090108***************************************************************
        If lnTipoAsiento = 3 Then
           lsCta = lsCtaCod
           If gsCodCMAC = "112" Then
                lsCta = sCtaHaber & Mid(lsCta, 5, 4) & IIf(Mid(lsCta, 7, 2) = "01", "09", "01") & Mid(lsCta, 9, 2) & Mid(lsCta, 13, 20)
           ElseIf gsCodCMAC = "106" Then
                lsCta = sCtaHaber & Mid(lsCta, 5, 4) & IIf(Mid(lsCta, 7, 2) = "01", "09", "01") & Mid(lsCta, 9, 2) & Mid(lsCta, 13, 20)
           Else
               If Mid(lsCtaCod, 5, 2) = "03" And Mid(lsCtaCod, 7, 2) <> "02" And (Mid(lsCtaCod, 9, 2) = "13" Or Mid(lsCtaCod, 9, 2) = "20") Then
                'ALPA 20110301
                 If Mid(lsCtaCod, 7, 2) = "01" Then
                    lsCta = sCtaHaber & Mid(lsCtaCod, 5, 4) & "060101" & Right(lsCtaCod, 2)
                 Else
                    lsCta = sCtaHaber & Mid(lsCtaCod, 5, 4) & "010101" & Right(lsCtaCod, 2)
                 End If
                  
               ElseIf Mid(lsCtaCod, 5, 2) = "03" And Mid(lsCtaCod, 7, 2) = "02" And (Mid(lsCtaCod, 11, 2) = "13" Or Mid(lsCtaCod, 11, 2) = "20") Then
                    If Mid(lsCtaCod, 9, 2) = "01" Then
                        lsCta = sCtaHaber & Mid(lsCtaCod, 5, 4) & "010101" & Right(lsCtaCod, 2)
                    Else
                        lsCta = sCtaHaber & Mid(lsCtaCod, 5, 4) & "020101" & Right(lsCtaCod, 2)
                    End If
               Else
                    If Mid(lsCtaCod, 7, 2) = "02" Then
                        If Mid(lsCtaCod, 9, 2) = "01" Then
                            lsCta = sCtaHaber & Mid(lsCtaCod, 5, 4) & "01" & Right(lsCtaCod, 6)
                        Else
                            lsCta = sCtaHaber & Mid(lsCtaCod, 5, 4) & "02" & Right(lsCtaCod, 6)
                        End If
                    Else
                        'ALPA 20110301
                        If Mid(lsCtaCod, 5, 2) = "02" Or Mid(lsCtaCod, 5, 2) = "03" Or Mid(lsCtaCod, 5, 2) = "04" Or Mid(lsCtaCod, 5, 2) = "11" Or Mid(lsCtaCod, 5, 2) = "12" Or Mid(lsCtaCod, 5, 2) = "13" Or Mid(lsCtaCod, 5, 2) = "09" Then
                        'If Mid(lsCtaCod, 5, 2) = "02" Or Mid(lsCtaCod, 5, 2) = "03" Or Mid(lsCtaCod, 5, 2) = "04" Or Mid(lsCtaCod, 5, 2) = "11" Or Mid(lsCtaCod, 5, 2) = "12" Or Mid(lsCtaCod, 5, 2) = "13" Or Mid(lsCtaCod, 5, 2) = "09" Then
                                                                                    
                            'PEAC 20210305
                            If Mid(lsCtaCod, 7, 2) = "07" Then
                                lsCta = sCtaHaber & Mid(lsCtaCod, 5, 4) & Right(lsCtaCod, 4)
                            Else
                                lsCta = sCtaHaber & Mid(lsCtaCod, 5, 4) & "0601" & Right(lsCtaCod, 4)
                            End If
                            'lsCta = sCtaHaber & Mid(lsCtaCod, 5, 4) & "0601" & Right(lsCtaCod, 4)
                        Else
                            lsCta = sCtaHaber & Mid(lsCtaCod, 5, 4) & "01" & Right(lsCtaCod, 6)
                        End If
                       
                        'lsCta = sCtaHaber & Mid(lsCtaCod, 5, 4) & "01" & Right(lsCtaCod, 6)
                    End If
               End If
           End If
        ElseIf lnTipoAsiento = 8 Then
           lsCta = lsCtaCod
           If gsCodCMAC = "112" Then
                lsCta = sCtaHaber & IIf(Mid(lsCta, 5, 2) = "01", "02", "01") & IIf(Mid(lsCta, 5, 2) = "01", "09", "01") & Mid(lsCta, 7, 22)
           Else
           End If
        ElseIf lnTipoAsiento = 15 Then
            lsCta = rs.Fields(1)
        Else
           If lnTipoAsiento = 7 Or lnTipoAsiento = 7 Or lnTipoAsiento = 28 Then 'PASI 20170417 Tipo Asiento = 28
                lsCta = rs!Cta2
'              If Mid(lsCtaCod, 7, 2) = "03" Then
'                 If Mid(lsCtaCod, 9, 2) = "13" Then
'                    lsCta = sCtaHaber & Mid(lsCtaCod, 7, 2) & "06" & "01" & Mid(lsCtaCod, 9, 20)
'                 ElseIf Mid(lsCtaCod, 9, 2) = "20" Then
'                    lsCta = sCtaHaber & Mid(lsCtaCod, 7, 2) & "20" & "01" & Mid(lsCtaCod, 11, 20)
'                 Else
'                    lsCta = sCtaHaber & Mid(lsCtaCod, 7, 2) & "06" & "01" & Mid(lsCtaCod, 11, 20)
'                 End If
'              Else
'                 lsCta = sCtaHaber & Mid(lsCtaCod, 7, 2) & "06" & "01" & Mid(lsCtaCod, 11, 20)
'              End If

           ElseIf lnTipoAsiento = 6 Or lnTipoAsiento = 22 Then 'JUEZ 20130116 Se agregó lnTipoAsiento = 22
                lsCta = sCtaHaber
           ElseIf lnTipoAsiento = 11 Then
                lsCta = sCtaDebe & Mid(lsCtaCod, 5, 2) & "01" & Mid(lsCtaCod, 5, 2) & "01" & Right(lsCtaCod, 2)
                'lsCta = sCtaDebe & "010101" & Right(lsCta, 2)
                If gnImporte < 0 Then
                    lsCta = "64" & Mid(lsCta, 3, 1) & "1040201" & Right(lsCta, 2)
                End If
            ElseIf lnTipoAsiento = 12 Or lnTipoAsiento = 14 Then
                lsCta = sCtaHaber
            ElseIf lnTipoAsiento = 13 Or lnTipoAsiento = 16 Then
                lsCta = sCtaHaber
            ElseIf lnTipoAsiento = 17 Or lnTipoAsiento = 29 Or lnTipoAsiento = 30 Or lnTipoAsiento = 31 Then 'NAGL 202007 Agregó lnTipoAsiento (29,30,31)
                lsCta = sCtaHaber
            ElseIf lnTipoAsiento = 18 Or lnTipoAsiento = 21 Or lnTipoAsiento = 25 Or lnTipoAsiento = 34 Then 'NAGL 202102 Agregó lnTipoAsiento = 34
                lsCta = lsContraCta1 'Mid(lsCtaCod, 3, 1) + "4" + IIf(Mid(lsCtaCod, 7, 2) = "04", "01", "06") + Mid(lsCtaCod, 9, 4) + "0101" + Right(lsCtaCod, 2)
            ElseIf lnTipoAsiento = 19 Then
                lsCta = "83" + Mid(lsCtaCod, 3, 1) + "1"
            'MADM 20110805
            ElseIf lnTipoAsiento = 20 Then
                lsCta = "72" + Mid(lsCtaCod, 3, 1) + "2"
            'END MADM
            ElseIf lnTipoAsiento = 23 Then 'EJVG20130307
                lsCta = "52" + Mid(lsCtaCod, 3, 1) + "102" & Right(lsCtaCod, 2)
            ElseIf lnTipoAsiento = 24 Then 'ALPA20131129
                lsCta = IIf(IsNull(rs!cCtaContraCta), "", rs!cCtaContraCta)
           Else
            If gsCodCMAC = "102" Then
                lsCta = sCtaHaber & Mid(lsCtaCod, 5, Len(Mid(lsCtaCod, 5, 22)) - 2)
            Else
                If Mid(lsCtaCod, 7, 2) = "20" Then
                    lsCta = sCtaHaber & Mid(lsCtaCod, 5, 6) & Right(lsCtaCod, 2)
                Else
                    If Mid(lsCtaCod, 7, 2) = "13" Or Mid(lsCtaCod, 7, 2) = "20" Then
                        'lsCta = sCtaHaber & Mid(lsCtaCod, 5, 4) & "01" & Mid(lsCtaCod, 9, 20)
                        If Left(sCtaHaber, 4) = "5114" And Mid(lsCtaCod, 7, 2) = "13" And lnTipoAsiento = 1 Then
                            lsCta = sCtaHaber & Mid(lsCtaCod, 5, 4) & "02" & Right(lsCtaCod, 2) 'ALPA20140303
                        Else
                            lsCta = sCtaHaber & Mid(lsCtaCod, 5, 4) & "01" & Right(lsCtaCod, 2) 'EJVG20130207
                        End If
                    Else
                        If Mid(lsCtaCod, 11, 2) = "13" Then
                            lsCta = sCtaHaber & Mid(lsCtaCod, 5, 2) & Mid(lsCtaCod, 11, 20)
                        Else
                            'lsCta = sCtaHaber & Mid(lsCtaCod, 5, 4) & IIf(Mid(lsCtaCod, 5, 2) = "03", "09", "02") & Mid(lsCtaCod, 9, 20)
                            lsCta = sCtaHaber & Mid(lsCtaCod, 5, 2) & IIf(Mid(lsCtaCod, 5, 2) = "09", Mid(lsCtaCod, 11, 2) & Mid(lsCtaCod, 7, 4), Mid(lsCtaCod, 7, 2)) & IIf(Mid(lsCtaCod, 5, 2) = "09", Mid(lsCtaCod, 15, 10), Mid(lsCtaCod, 11, 10))
                            'lsCta = sCtaHaber & Mid(lsCtaCod, 5, 4) & Mid(lsCtaCod, 11, 20)
                        End If
                    End If
                End If
            End If
           End If
        End If
        
        'If lnTipoAsiento = 1 Or lnTipoAsiento = 3 Or lnTipoAsiento = 7 Or lnTipoAsiento = 8 Or lnTipoAsiento = 11 Or lnTipoAsiento = 12 Or lnTipoAsiento = 13 Or lnTipoAsiento = 15 Then
        'MADM 20110805 - 20
        'If lnTipoAsiento = 1 Or lnTipoAsiento = 3 Or lnTipoAsiento = 7 Or lnTipoAsiento = 8 Or lnTipoAsiento = 11 Or lnTipoAsiento = 12 Or lnTipoAsiento = 13 Or lnTipoAsiento = 15 Or lnTipoAsiento = 17 Or lnTipoAsiento = 18 Or lnTipoAsiento = 19 Or lnTipoAsiento = 20 Or lnTipoAsiento = 21 Then
        If lnTipoAsiento = 1 Or lnTipoAsiento = 3 Or lnTipoAsiento = 7 Or lnTipoAsiento = 8 Or lnTipoAsiento = 11 Or lnTipoAsiento = 12 Or lnTipoAsiento = 13 Or lnTipoAsiento = 15 Or lnTipoAsiento = 17 Or lnTipoAsiento = 18 Or lnTipoAsiento = 19 Or lnTipoAsiento = 20 Or lnTipoAsiento = 21 Or lnTipoAsiento = 23 Or lnTipoAsiento = 24 Or lnTipoAsiento = 25 Or lnTipoAsiento = 27 Or lnTipoAsiento = 28 Or lnTipoAsiento = 29 Or lnTipoAsiento = 30 Or lnTipoAsiento = 31 Or lnTipoAsiento = 34 Then 'PASI 20170417 Tipo Asiento = 28
            'NAGL 202007 lnTipoAsiento (29,30,31)
            'NAGL 202102 lnTipoAsiento = 34
            'nItem = nItem + 1
            If lnTipoAsiento = 11 Then
                'oMov.InsertaMovCta gnMovNro, nItem, lsCta, Abs(gnImporte) * IIf(gnImporte > 0, 1, -1)  ' si el importe es mayor cero
                If Mid(rs!ctaNuevaN, 1, 1) = "6" Then
                    nItem = nItem + 1
                    'oMov.InsertaMovCta gnMovNro, nItem, rs!ctaNuevaN, Abs(gnImporte) * IIf(gnImporte > 0, 1, -1)  ' si el importe es mayor cero
                    oMov.InsertaMovCta gnMovNro, nItem, rs!ctaNuevaN, DevolverSaldoTC(rs!ctaNuevaN, nVal(txtTipCambio.Text), nVal(txtTipCambioVenta.Text), nVal(txtTipCambioCompra.Text), Abs(gnImporte) * IIf(gnImporte > 0, 1, -1), Mid(rs!ctaNuevaN, 3, 1))
                Else
                    'If rs!nSaldoProvision <> 0 Then
                    If gnImporte <> 0 Then
                         nItem = nItem + 1
                        'oMov.InsertaMovCta gnMovNro, nItem, rs!ctaNuevaN, Abs(gnImporte) * IIf(gnImporte > 0, 1, -1) ' si el importe es mayor cero
                        oMov.InsertaMovCta gnMovNro, nItem, rs!ctaNuevaN, DevolverSaldoTC(rs!ctaNuevaN, nVal(txtTipCambio.Text), nVal(txtTipCambioVenta.Text), nVal(txtTipCambioCompra.Text), Abs(gnImporte) * IIf(gnImporte > 0, 1, -1), Mid(rs!ctaNuevaN, 3, 1))
                    End If
'                    If rs!nSaldoProvisionPro <> 0 Then
'                         nItem = nItem + 1
'                        oMov.InsertaMovCta gnMovNro, nItem, rs!ctaNuevaP, Abs(rs!nSaldoProvisionPro) * IIf(rs!nSaldoProvisionPro > 0, 1, -1)  ' si el importe es mayor cero
'                    End If
                End If
            ElseIf lnTipoAsiento = 17 Or lnTipoAsiento = 29 Or lnTipoAsiento = 30 Or lnTipoAsiento = 31 Or lnTipoAsiento = 34 Then
                   'NAGL 202007 Agregó lnTipoAsiento (29,30,31)
                   'NAGL 202102 Agregó lnTipoAsiento = 34
                    nItem = nItem + 1
                    'oMov.InsertaMovCta gnMovNro, nItem, lsCta, Abs(gnImporte) * IIf(gnImporte > 0, -1, 1)  ' si el importe es mayor cero
                    oMov.InsertaMovCta gnMovNro, nItem, lsCta, DevolverSaldoTC(lsCta, nVal(txtTipCambio.Text), nVal(txtTipCambioVenta.Text), nVal(txtTipCambioCompra.Text), Abs(gnImporte) * IIf(gnImporte > 0, -1, 1), Mid(lsCta, 3, 1))
            Else
                 nItem = nItem + 1
                 If lnTipoAsiento = 3 Then
                    If lsCta = "" Then
                    End If
                 End If
                'MADM 20110805 - 20
                'If lnTipoAsiento = 18 Then
                 '   If gnImporte > 0 Then
                        'oMov.InsertaMovCta gnMovNro, nItem, lsCta, gnImporte * IIf(lnTipoAsiento = 3 Or lnTipoAsiento = 8 Or lnTipoAsiento = 19 Or lnTipoAsiento = 21, 1, -1)
                        oMov.InsertaMovCta gnMovNro, nItem, lsCta, DevolverSaldoTC(lsCta, nVal(txtTipCambio.Text), nVal(txtTipCambioVenta.Text), nVal(txtTipCambioCompra.Text), gnImporte * IIf(lnTipoAsiento = 3 Or lnTipoAsiento = 8 Or lnTipoAsiento = 19 Or lnTipoAsiento = 21 Or lnTipoAsiento = 23 Or lnTipoAsiento = 24 Or lnTipoAsiento = 25, 1, -1), Mid(lsCta, 3, 1))

                  '  End If
               
                'Else
                 '   oMov.InsertaMovCta gnMovNro, nItem, lsCta, gnImporte * IIf(lnTipoAsiento = 3 Or lnTipoAsiento = 8 Or lnTipoAsiento = 19, 1, -1)
                'End If
                
            End If
            
           If lnTipoAsiento = 7 Then 'insertamos asiento de reversion de interes devengado a la cuenta de suspenso EJRS
                lsCta = rs!Cta3
                nItem = nItem + 1
                'oMov.InsertaMovCta gnMovNro, nItem, lsCta, gnImporte
                oMov.InsertaMovCta gnMovNro, nItem, lsCta, DevolverSaldoTC(lsCta, nVal(txtTipCambio.Text), nVal(txtTipCambioVenta.Text), nVal(txtTipCambioCompra.Text), gnImporte, Mid(lsCta, 3, 1))

                
                lsCta = rs!Cta4
                nItem = nItem + 1
                'oMov.InsertaMovCta gnMovNro, nItem, lsCta, gnImporte * -1
                oMov.InsertaMovCta gnMovNro, nItem, lsCta, DevolverSaldoTC(lsCta, nVal(txtTipCambio.Text), nVal(txtTipCambioVenta.Text), nVal(txtTipCambioCompra.Text), gnImporte * -1, Mid(lsCta, 3, 1))
           End If
'           If lnTipoAsiento = 17 Then
'                nItem = nItem + 1
'                oMov.InsertaMovCta gnMovNro, nItem, rs!cta2, rs!nCtaSaldoImporte ' * IIf(Mid(rs!cta2, 5, 4) = "0601", 1, -1)
'           End If
            
        Else
           nTotal = nTotal + gnImporte
        End If
        
      End If
      End If
      oImp_BarraProgress rs.Bookmark, "ASIENTO DE INTERESES", "", "Grabando...", vbBlue
      rs.MoveNext
    'ALPA 20141206*******************
   
    '********************************
   Loop
   
   'ALPA 20141206**********
   If lnTipoAsiento <> 26 Then
   '**********************
   'ALPA 20090525*************************
   'If lnTipoAsiento = 2 Or lnTipoAsiento = 5 Or lnTipoAsiento = 16 Or lnTipoAsiento = 6 Or lnTipoAsiento = 9 Or lnTipoAsiento = 10 Then
   If lnTipoAsiento = 2 Or lnTipoAsiento = 5 Or lnTipoAsiento = 16 Or lnTipoAsiento = 6 Or lnTipoAsiento = 9 Or lnTipoAsiento = 10 Or lnTipoAsiento = 22 Then 'JUEZ 20130116 Se agregó lnTipoAsiento = 22
      nItem = nItem + 1
      'oMov.InsertaMovCta gnMovNro, nItem, sCtaHaber, nTotal * -1
      oMov.InsertaMovCta gnMovNro, nItem, sCtaHaber, DevolverSaldoTC(sCtaHaber, nVal(txtTipCambio.Text), nVal(txtTipCambioVenta.Text), nVal(txtTipCambioCompra.Text), nTotal * -1, Mid(sCtaHaber, 3, 1))
      
   End If
   
   'ALPA 20090525*************************
   'If (lnTipoAsiento = 1 Or lnTipoAsiento = 2 Or lnTipoAsiento = 3 Or lnTipoAsiento = 5 Or lnTipoAsiento = 16 Or lnTipoAsiento = 6 Or lnTipoAsiento = 7 Or lnTipoAsiento = 8 Or lnTipoAsiento = 9 Or lnTipoAsiento = 10 Or lnTipoAsiento = 11 Or lnTipoAsiento = 12 Or lnTipoAsiento = 13 Or lnTipoAsiento = 14 Or lnTipoAsiento = 15) And Mid(gsOpeCod, 3, 1) = "2" Then
   
   '*** PEAC 20130114
    If lnTipoAsiento = 14 Then
        If Not rsIntDev.EOF Then
            oImp_BarraShow rsIntDev.RecordCount
            
            Dim lsCtaOtr As String
            nItem = nItem + 1
            
            Do While Not rsIntDev.EOF
                gnImporte = nVal(Format(rsIntDev!nSaldo, "#.00")) - rsIntDev!nCtaSaldoImporte
                lsCtaCod = IIf(IsNull(rsIntDev!Cta1), rsIntDev!Cta2, rsIntDev!Cta1)
                lsCtaOtr = "41" & Mid(lsCtaCod, 3, 1) & "1030301" & Right(lsCtaCod, 2)
                If gnImporte <> 0 Then
                    nItem = nItem + 1
                    If rsIntDev!nSaldo > rsIntDev!nCtaSaldoImporte Then ' si estadistico(monto del informe) es mayor que contable
                        oMov.InsertaMovCta gnMovNro, nItem, lsCtaOtr, DevolverSaldoTC(lsCtaOtr, nVal(txtTipCambio.Text), nVal(txtTipCambioVenta.Text), nVal(txtTipCambioCompra.Text), Abs(gnImporte), Mid(lsCtaOtr, 3, 1))
                        nItem = nItem + 1
                        oMov.InsertaMovCta gnMovNro, nItem, lsCtaCod, DevolverSaldoTC(lsCtaCod, nVal(txtTipCambio.Text), nVal(txtTipCambioVenta.Text), nVal(txtTipCambioCompra.Text), Abs(gnImporte) * -1, Mid(lsCtaCod, 3, 1))
                    Else
                        oMov.InsertaMovCta gnMovNro, nItem, lsCtaCod, DevolverSaldoTC(lsCtaCod, nVal(txtTipCambio.Text), nVal(txtTipCambioVenta.Text), nVal(txtTipCambioCompra.Text), Abs(gnImporte), Mid(lsCtaCod, 3, 1))
                        nItem = nItem + 1
                        oMov.InsertaMovCta gnMovNro, nItem, lsCtaOtr, DevolverSaldoTC(lsCtaOtr, nVal(txtTipCambio.Text), nVal(txtTipCambioVenta.Text), nVal(txtTipCambioCompra.Text), Abs(gnImporte) * -1, Mid(lsCtaOtr, 3, 1))
                    End If
                End If
                oImp_BarraProgress rsIntDev.Bookmark, "ASIENTO DE INTERESES DEVENGADOS", "", "Grabando...", vbBlue
                rsIntDev.MoveNext
            Loop
        End If
    End If
   '*** FIN PEAC
   'ALPA 20141205**********
    End If
    '**********************
    If (lnTipoAsiento = 1 Or lnTipoAsiento = 2 Or lnTipoAsiento = 3 Or lnTipoAsiento = 5 Or lnTipoAsiento = 16 Or lnTipoAsiento = 6 Or lnTipoAsiento = 7 Or lnTipoAsiento = 8 Or lnTipoAsiento = 9 Or lnTipoAsiento = 10 Or lnTipoAsiento = 11 Or lnTipoAsiento = 12 Or lnTipoAsiento = 13 Or lnTipoAsiento = 14 Or lnTipoAsiento = 15 Or lnTipoAsiento = 17 Or lnTipoAsiento = 19 Or lnTipoAsiento = 20 Or lnTipoAsiento = 18 Or lnTipoAsiento = 21 Or lnTipoAsiento = 22 Or lnTipoAsiento = 23 Or lnTipoAsiento = 24 Or lnTipoAsiento = 25 Or lnTipoAsiento = 26 Or lnTipoAsiento = 28 Or lnTipoAsiento = 29 Or lnTipoAsiento = 30 Or lnTipoAsiento = 31 Or lnTipoAsiento = 34) And Mid(gsOpeCod, 3, 1) = "2" Then 'JUEZ 20130116 Se agregó lnTipoAsiento = 22
    'PASI 20170417 Tipo Asiento = 28
    'NAGL 202007 Agregó lnTipoAsiento (29,30,31)
    'NAGL 202102 Agregó lnTipoAsiento = 34
        oMov.GeneraMovME gnMovNro, gsMovNro
    End If

   If dFecha < gdFecSis Then
      oMov.ActualizaSaldoMovimiento gsMovNro, "+"
   End If
   
  
   'PASI20170424****************
   Dim nIdAjuste As Integer
   If lnTipoAsiento = 28 Then 'VAPA 20170703
    If Not (rsAjusteContab.EOF And rsAjusteContab.BOF) Then
         nIdAjuste = oAjusteCont.RegistrarColocAjusteDeContabilidad(gsCodUser, gnMovNro)
         Do While Not rsAjusteContab.EOF
             oAjusteCont.RegistrarColocAjusteDeContabilidadDet nIdAjuste, rsAjusteContab!cCtaCod, rsAjusteContab!cCtaContDebe, rsAjusteContab!cCtaContHaber, rsAjusteContab!nSaldo
             rsAjusteContab.MoveNext
         Loop
    End If
   End If
   'PASI END*****

   oMov.CommitTrans
   lTransActiva = False
   oImp_BarraClose
End If
RSClose rs
'EJVG20130322 ***
If lnTipoAsiento = 20 Then
    lsTituloImpre = "ASIENTO DE CAPITAL"
ElseIf lnTipoAsiento = 29 Or lnTipoAsiento = 30 Or lnTipoAsiento = 31 Or lnTipoAsiento = 34 Then 'NAGL 202007 Agregó Condicional - lnTipoAsiento (29,30,31)
'NAGL 202102 Agregó lnTipoAsiento = 34
    lsTituloImpre = "ASIENTO DE CAPITAL E INTERESES"
Else
    lsTituloImpre = "ASIENTO DE INTERESES"
End If
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaMantDocumento
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & "Se Grabo el Asiento Contable al Cierre " & dFecha & " con el Tipo de Cambio Fijo : " & txtTipCambio.Text _
            & " Tpo Cambio venta : " & txtTipCambioVenta & " Tpo Cambio Venta : " & txtTipCambioCompra
            Set objPista = Nothing
            '*******
'ImprimeAsientoContable gsMovNro, , , , , , , , , , , , 1, "ASIENTO DE INTERESES"
ImprimeAsientoContable gsMovNro, , , , , , , , , , , , 1, lsTituloImpre
'END EJVG ********
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
If MsgBox("¿ Seguro que desea generar Cuadro ? ", vbQuestion + vbYesNo, "!Confirmación!") = vbNo Then
   Exit Sub
End If
GeneraReporte
End Sub
Private Sub cmdGenerarDet_Click()
Dim oAju As New DAjusteCont
Dim oCont As New NContFunciones
'***********NAGL 202102***************
Dim rsAjuste As ADODB.Recordset
Dim rsAjusteSubProc As ADODB.Recordset
Set rsAjuste = New ADODB.Recordset
Set rsAjusteSubProc = New ADODB.Recordset
'*************************************
Dim nMes As Integer, nAnio As Integer
Dim dFechaFin As Date

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim fs As Scripting.FileSystemObject
Dim lsPlantillaALeer As String
Dim lsNomHoja As String
Dim lbExisteHoja As Boolean
Dim lsArchivoAGuardar As String
Dim RutaArchivo As String
Dim lsCabeceraBarra As String
Dim iPos As Currency

If Not ValidaDatos Then
   Exit Sub
End If
If MsgBox("¿ Seguro que desea generar Cuadro ? ", vbQuestion + vbYesNo, "!Confirmación!") = vbNo Then
   Exit Sub
End If

nMes = cboMes.ListIndex + 1
nAnio = txtAnio.Text
dFechaFin = DateAdd("m", 1, "01/" & Format(nMes, "00") & "/" & Format(txtAnio, "0000")) - 1
If Not oCont.PermiteModificarAsiento(Format(dFechaFin, gsFormatoMovFecha), False) Then
   Set oCont = Nothing
   MsgBox "Mes ya cerrado. Imposible generar Cuadro de Comparación", vbInformation, "!Aviso!"
   Exit Sub
End If
If lnTipoAsiento = 1 Then
    Set rsAjuste = oAju.AjusteIntDevengadoDet("14", Mid(gsOpeCod, 3, 1), dFechaFin)
ElseIf lnTipoAsiento = 2 Then
    Set rsAjuste = oAju.AjusteIntSuspensoDet("81", Mid(gsOpeCod, 3, 1), dFechaFin)
ElseIf lnTipoAsiento = 7 Then
    Set rsAjuste = oAju.AjusteRevIntDevengadoDet("51", Mid(gsOpeCod, 3, 1), dFechaFin)
ElseIf lnTipoAsiento = 3 Then
    If lnTipoAsiento = 3 Then
        gsMovNro = Format(dFechaFin, "yyyymmdd")
        If Not oCont.ExisteMovProvEspecificaCOVID(gsMovNro, Mid(gsOpeCod, 3, 1)) Then
            Set oCont = Nothing
            If MsgBox("Aún no se ha generado la provisión de cartera reprogramada Covid en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & ", Desea continuar?", vbInformation + vbYesNo, "Atención") = vbNo Then
                Exit Sub
            End If
        End If
    End If 'NAGL 202102 Según ACTA N°017-2021
    Set rsAjuste = oAju.AjusteProvisionCarteraDet("14", Mid(gsOpeCod, 3, 1), dFechaFin)
'************************BEGIN NAGL 202008****************************'
ElseIf lnTipoAsiento = 32 Or lnTipoAsiento = 33 Or lnTipoAsiento = 34 Then 'NAGL 202102 Agregó lnTipoAsiento = 34
    Set rsAjuste = Nothing 'NAGL 202102
    Set rsAjusteSubProc = Nothing 'NAGL 202102
    If lnTipoAsiento = 32 Then
        Call oAju.InsertaInteresesDiferidos(dFechaFin, gsOpeCod)
        Set rsAjuste = oAju.CargaAjusteReporteDet(dFechaFin, lnTipoAsiento, gsOpeCod)
        Set rsAjusteSubProc = oAju.CargaAjusteReporteDet(dFechaFin, lnTipoAsiento, gsOpeCod, "SNeg")
    ElseIf lnTipoAsiento = 33 Then
        Set rsAjuste = oAju.CargaAjusteReporteDet(dFechaFin, lnTipoAsiento, gsOpeCod)
        Set rsAjusteSubProc = rsAjuste
    ElseIf lnTipoAsiento = 34 Then
        gsMovNro = Format(dFechaFin, "yyyymmdd")
        If oCont.ExisteMovimiento(Left(gsMovNro, 6), gsOpeCod) Then
            MsgBox "No existen diferencias en Provisiones de Cartera Reprogramada COVID - " & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & "..!!", vbInformation, "Aviso"
            Exit Sub
        Else
            Set rsAjuste = oAju.CargaAjusteReporteDet(dFechaFin, lnTipoAsiento, Mid(gsOpeCod, 3, 1))
            Set rsAjusteSubProc = rsAjuste
        End If 'NAGL 202102 Según Acta N°017-2021
    End If
    If (Not rsAjuste.BOF And Not rsAjuste.EOF) Or (Not rsAjusteSubProc.BOF And Not rsAjusteSubProc.EOF) Then
    'NAGL 202102 Agregó rsAjusteSubProc en la condición
        Call ObtieneReporteDetalle(lnTipoAsiento, rsAjuste, rsAjusteSubProc) 'NAGL 202102 Agregó rsAjust, rsAjusteSubProc
         If lnTipoAsiento = 34 Then
            cmdGenerar.Enabled = True
            cmdAsiento.Enabled = True
        Else
            cmdAsiento.Enabled = True
        End If 'NAGL 202102 Agregó Condición
        
        'ARLO20170208
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Mostro la Comparacion Estadistica Detallada al Cierre " & dFecha & " con el Tipo de Cambio Fijo : " & txtTipCambio.Text _
        & "Tpo Cambio venta : " & txtTipCambioVenta & " Tpo Cambio Venta : " & txtTipCambioCompra
        Set objPista = Nothing
        '*******
    Else
        If lnTipoAsiento = 32 Then
            MsgBox "No existen registros de Intereses Diferidos en Créditos Cancelados - " & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & "..!!", vbInformation, "Aviso"
        ElseIf lnTipoAsiento = 33 Then
            MsgBox "No existen diferencias en Cuentas de Orden, entre el Saldo Contable y el RCD - " & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & "..!!", vbInformation, "Aviso"
        ElseIf lnTipoAsiento = 34 Then
            MsgBox "No existen registros de Provisiones de Cartera Reprogramada COVID - " & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & "..!!", vbInformation, "Aviso"
        End If 'NAGL202102
    End If
    Exit Sub
End If '*******END NAGL 202008 Agregó esta sección Según Acta N°063-2020

If (rsAjuste.EOF And rsAjuste.BOF) Then
    MsgBox "No se ha podido generar el reporte. Favor de comunicarse con el Area de TI."
    Exit Sub
End If
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application
    
If lnTipoAsiento = 1 Or lnTipoAsiento = 2 Or lnTipoAsiento = 3 Then
    lsPlantillaALeer = "Ajuste"
ElseIf lnTipoAsiento = 7 Then
    lsPlantillaALeer = "AjusteRev"
End If

If lnTipoAsiento = 1 Then
    lsArchivoAGuardar = "\spooler\AjusteIntDev" & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & Format(dFechaFin, "yyyymmdd") & gsCodUser & ".xlsx"
ElseIf lnTipoAsiento = 2 Then
    lsArchivoAGuardar = "\spooler\AjusteIntSus" & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & Format(dFechaFin, "yyyymmdd") & gsCodUser & ".xlsx"
ElseIf lnTipoAsiento = 7 Then
    lsArchivoAGuardar = "\spooler\AjusteRevIntDev" & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & Format(dFechaFin, "yyyymmdd") & gsCodUser & ".xlsx"
ElseIf lnTipoAsiento = 3 Then
    lsArchivoAGuardar = "\spooler\AjusteProvisionCartera" & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & Format(dFechaFin, "yyyymmdd") & gsCodUser & ".xlsx"
End If
lsNomHoja = "AjusteDet"
RutaArchivo = App.path & "\FormatoCarta\" & lsPlantillaALeer & ".xlsx"
If fs.FileExists(RutaArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsPlantillaALeer & ".xlsx")
Else
    MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
    Exit Sub
End If

For Each xlHoja1 In xlLibro.Worksheets
   If xlHoja1.Name = lsNomHoja Then
        xlHoja1.Activate
        lbExisteHoja = True
        Exit For
   End If
Next
If lbExisteHoja = False Then
    Set xlHoja1 = xlLibro.Worksheets
    xlHoja1.Name = lsNomHoja
End If

If lnTipoAsiento = 1 Then
    xlHoja1.Cells(2, 2) = "INTERESES DEVENGADOS DETALLADO"
    lsCabeceraBarra = "CUADRO DE INTERESES DEVENGADOS"
ElseIf lnTipoAsiento = 2 Then
    xlHoja1.Cells(2, 2) = "INTERESES DE SUSPENSO DETALLADO"
    lsCabeceraBarra = "CUADRO DE INTERESES DE SUSPENSO"
ElseIf lnTipoAsiento = 7 Then
    xlHoja1.Cells(2, 2) = "REVERSIÓN DE INTERESES DEVENGADOS DETALLADO"
    lsCabeceraBarra = "CUADRO DE REVERSIÓN DE INTERESES DEVENGADOS"
ElseIf lnTipoAsiento = 3 Then
    xlHoja1.Cells(2, 2) = "PROVISIÓN DE CARTERA DE CRÉDITOS DETALLADO"
    lsCabeceraBarra = "CUADRO DE PROVISIÓN DE CARTERA DE CRÉDITOS"
End If
xlHoja1.Cells(3, 2) = "Fecha : " & dFechaFin & " - Moneda : " & IIf(Mid(gsOpeCod, 3, 1) = "1", "SOLES", "DOLARES")
iPos = 6
oImp_BarraShow rsAjuste.RecordCount

If lnTipoAsiento = 1 Or lnTipoAsiento = 2 Or lnTipoAsiento = 3 Then
    If lnTipoAsiento = 3 Then 'NAGL 202102 Según ACTA N°017-2021
        oImp_BarraProgress rsAjuste.RecordCount / 2, lsCabeceraBarra, "", "Generando...", vbBlue
        xlHoja1.Range(xlHoja1.Cells(iPos, 2), xlHoja1.Cells(iPos, 2)).CopyFromRecordset rsAjuste
        oImp_BarraProgress rsAjuste.RecordCount, lsCabeceraBarra, "", "Generando...", vbBlue
    Else
        Do While Not rsAjuste.EOF
            xlHoja1.Cells(iPos, 2) = rsAjuste!cCtaCod
            xlHoja1.Cells(iPos, 3) = rsAjuste!ctaContDebe
            xlHoja1.Cells(iPos, 4) = rsAjuste!ctaContHaber
            xlHoja1.Cells(iPos, 5) = rsAjuste!nMonto
            iPos = iPos + 1
            oImp_BarraProgress rsAjuste.Bookmark, lsCabeceraBarra, "", "Generando...", vbBlue
            rsAjuste.MoveNext
        Loop
   End If
ElseIf lnTipoAsiento = 7 Then
     Do While Not rsAjuste.EOF
        xlHoja1.Cells(iPos, 2) = rsAjuste!cCtaCod
        xlHoja1.Cells(iPos, 3) = rsAjuste!Cta1
        xlHoja1.Cells(iPos, 4) = rsAjuste!Cta2
        xlHoja1.Cells(iPos, 5) = rsAjuste!Cta3
        xlHoja1.Cells(iPos, 6) = rsAjuste!Cta4
        xlHoja1.Cells(iPos, 7) = rsAjuste!nSaldo
        iPos = iPos + 1
        oImp_BarraProgress rsAjuste.Bookmark, lsCabeceraBarra, "", "Generando...", vbBlue
        rsAjuste.MoveNext
    Loop
End If

oImp_BarraClose
xlHoja1.SaveAs App.path & lsArchivoAGuardar
xlAplicacion.Visible = True
xlAplicacion.Windows(1).Visible = True
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaMantDocumento
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Mostro la Comparacion Estadistica Detallada al Cierre " & dFecha & " con el Tipo de Cambio Fijo : " & txtTipCambio.Text _
            & "Tpo Cambio venta : " & txtTipCambioVenta & " Tpo Cambio Venta : " & txtTipCambioCompra
            Set objPista = Nothing
            '*******
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset
CentraForm Me
frmOperaciones.Enabled = False
Set rs = New ADODB.Recordset 'PASI20170424
If lnTipoAsiento = 1 Then
   Me.Caption = "Colocaciones: Intereses Devengados"
End If
If lnTipoAsiento = 2 Then
    Me.Caption = "Captaciones: Intereses en Suspenso"
End If
If lnTipoAsiento = 3 Then
    Me.Caption = "Colocaciones: Provisión de Cartera"
End If
If lnTipoAsiento = 5 Then
    Me.Caption = "Colocaciones: Capital de Creditos Castigados"
End If
If lnTipoAsiento = 6 Then
    Me.Caption = "Colocaciones: Intereses de Creditos Castigados"
End If
If lnTipoAsiento = 7 Then
    Me.Caption = "Colocaciones: Reversión de Intereses Devengados"
End If
If lnTipoAsiento = 11 Then
    Me.Caption = "Colocaciones: Provisión de Carta Fianza"
End If

If lnTipoAsiento = 12 Then
    Me.Caption = "Colocaciones: Ajuste por calificacion de cartera"
End If
If lnTipoAsiento = 13 Then
    Me.Caption = "Colocaciones: Asiento por Riesgo ponderado"
End If
'JUEZ 20130116 ****************************************************
If lnTipoAsiento = 22 Then
    Me.Caption = "Colocaciones: Gastos de Creditos Castigados"
End If
'END JUEZ *********************************************************
If lnTipoAsiento = 24 Then
    Me.Caption = "Colocaciones: Intereses Diferidos Final"
End If
If lnTipoAsiento = 25 Then
    Me.Caption = "Colocaciones: Compra de Deuda"
End If
If lnTipoAsiento = 26 Then
    Me.Caption = "Colocaciones: Reclasificación de Provisiones Prociclicas"
End If
If lnTipoAsiento = 27 Then
    Me.Caption = "Colocaciones: Reclasificación de Intereses Diferidos de Créditos Normales"
End If
'PASIERS0152017 20170104****************************
If lnTipoAsiento = 28 Then
    Me.Caption = "Colocaciones: Reversión de Intereses Devengados x Zona Inundada"
End If
'PASI END ******************************************
If Me.Caption = "" Then
   Me.Caption = gsOpeDesc
End If

'PASI20160223 ERS0072016
If (lnTipoAsiento = 1 Or lnTipoAsiento = 2 Or lnTipoAsiento = 7 Or lnTipoAsiento = 3 Or lnTipoAsiento = 34) Then
'NAGL 202102 Agregó lnTipoAsiento = 34
    Me.Height = 5265
    Me.cmdGenerarDet.Enabled = True
    Me.cmdGenerarDet.Visible = True
    Me.cmdAsiento.Top = 4020
    Me.cmdSalir.Top = 4400
    If lnTipoAsiento = 34 Then
        Me.cmdGenerarDet.Top = 3270
        Me.cmdGenerarDet.Caption = "Mostrar Estadístico Detallado"
        Me.cmdGenerar.Top = 3650
        Me.cmdGenerar.Caption = "Mostrar Comparación Estadística - Contable"
        Me.cmdGenerar.Enabled = False
    End If 'NAGL 202102
End If
'end PASI

If (lnTipoAsiento = 32 Or lnTipoAsiento = 33) Then
    cmdGenerar.Visible = False
    Me.cmdGenerarDet.Enabled = True
    Me.cmdGenerarDet.Visible = True
    Me.cmdGenerarDet.Caption = "Mostrar Estadístico Detallado"
    Me.cmdGenerarDet.Top = 3270
End If 'NAGL 202008

cboMes.ListIndex = Month(gdFecSis) - 1
txtAnio = Year(gdFecSis)

Dim oOpe As New DOperacion
Set rs = oOpe.CargaOpeCta(gsOpeCod, "D")
If Not rs.EOF Then
    sCtaDebe = rs!cCtaContCod
End If
Set rs = oOpe.CargaOpeCta(gsOpeCod, "H")
If Not rs.EOF Then
    sCtaHaber = rs!cCtaContCod
End If
RSClose rs
Set oOpe = Nothing
Set oImp = New NContImprimir
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
'If KeyAscii = 13 Then
'    If CboMes.ListIndex > -1 And txtAnio <> "" Then
'        txtTipCambio = TipoCambioCierre(txtAnio, CboMes.ListIndex + 1, False)
'        txtTipCambio.SetFocus
'    End If
'    txtTipCambio.SetFocus
'End If
If KeyAscii = 13 Then
    If cboMes.ListIndex > -1 And txtAnio <> "" Then
        txtTipCambio = TiposCambiosCierreMensual(txtAnio, cboMes.ListIndex + 1, False, 1)
        txtTipCambioVenta = TiposCambiosCierreMensual(txtAnio, cboMes.ListIndex + 1, False, 2)
        txtTipCambioCompra = TiposCambiosCierreMensual(txtAnio, cboMes.ListIndex + 1, False, 3)
        txtTipCambioCompra.SetFocus
    End If
    txtTipCambioCompra.SetFocus
End If
End Sub

Private Sub txtTipCambio_GotFocus()
    fEnfoque txtTipCambio
End Sub
'ALPA 20120402************************************
Private Sub txtTipCambioVenta_GotFocus()
    fEnfoque txtTipCambioVenta
End Sub

Private Sub txtTipCambioCompra_GotFocus()
    fEnfoque txtTipCambioCompra
End Sub
'**************************************************

Private Sub txtTipCambio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTipCambio, KeyAscii, 14, 5)
If KeyAscii = 13 Then
   cmdGenerar.SetFocus
End If
End Sub
'ALPA 20120402******************************************
Private Sub txtTipCambioVenta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTipCambioVenta, KeyAscii, 14, 5)
If KeyAscii = 13 Then
   cmdGenerar.SetFocus
End If
End Sub

Private Sub txtTipCambioCompra_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTipCambioCompra, KeyAscii, 14, 5)
If KeyAscii = 13 Then
   cmdGenerar.SetFocus
End If
End Sub
'*******************************************************

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If cboMes.ListIndex = -1 Then
   MsgBox "Debe seleccionarse mes de proceso", vbInformation, "!Aviso!"
   cboMes.SetFocus
   Exit Function
End If
If Val(txtAnio) = 0 Then
   MsgBox "Debe indicar año de proceso", vbInformation, "!Aviso!"
   txtAnio.SetFocus
   Exit Function
End If
If Val(txtTipCambio) = 0 Then
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

Private Sub ObtieneReporteDetalle(pnTipoAsiento As Integer, Optional rsAjuste As ADODB.Recordset = Nothing, Optional rsAjusteSubProc As ADODB.Recordset = Nothing)
'NAGL 202102 Agregó rsAjuste, rsAjusteSubProc
Dim oAjuste As New DAjusteCont
Dim oCambio As New nTipoCambio
'*****Comentado by NAGL 202102
'Dim rsAjuste As New ADODB.Recordset
'Dim rsAjusteSubProc As New ADODB.Recordset
'**********************************
'**********NAGL 202102************
Dim rsProc1 As New ADODB.Recordset
Dim rsProc2 As New ADODB.Recordset
Dim pnCol As Integer
Dim psFecActual As String
'*********************************
Dim oExcel As Object
Dim oBook As Object
Dim xlHoja1 As Object
Dim nRep As Integer, i As Integer 'NAGL 202008
Dim nLineasReg As Integer 'NAGL 202008
Dim psDescripDet As String 'NAGL 202008
Dim nMes As Integer
Dim nAnio As Integer
Dim dFecha As Date
Dim lsArchivoAGuardar As String
Dim psDescrip As String
Dim nLineas As Long

Dim oBarraDet As clsProgressBar
Dim TituloProgress As String
Dim MensajeProgress As String
Set oBarraDet = New clsProgressBar
oBarraDet.ShowForm Me
oBarraDet.Max = 3

If pnTipoAsiento = 31 Then 'NAGL 202008 Cambio a pnTipoAsiento
   psDescrip = "Cartera Reprogramada"
ElseIf pnTipoAsiento = 32 Then
   psDescrip = "Diferidos en Créditos Cancelados"
ElseIf pnTipoAsiento = 33 Then
   psDescrip = "Cuentas de Orden (Saldo Contable - RCD)"
ElseIf pnTipoAsiento = 34 Then 'NAGL 202002 Según Acta N°017-2021
   psDescrip = "Provisiones Covid Cartera Reprogramada"
End If 'NAGL 202008 Según Acta N°063-2020

psDescrip = "Reporte de Cierre " & psDescrip & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME")
oBarraDet.Progress 0, psDescrip, "GENERANDO EL ARCHIVO", "", vbBlue
TituloProgress = psDescrip
MensajeProgress = "GENERANDO EL ARCHIVO"

Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
Set xlHoja1 = oBook.Worksheets(1)
nMes = cboMes.ListIndex + 1
nAnio = txtAnio
dFecha = DateAdd("m", 1, "01/" & Format(nMes, "00") & "/" & Format(txtAnio, "0000")) - 1
nLineas = 3
i = 1 'NAGL 202008
nRep = 0 'NAGL 202008

If pnTipoAsiento = 31 Then 'Antes psTipo = "Reprog" 'NAGL 20200902
    lsArchivoAGuardar = "\spooler\CarteraReprogramada" & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & Format(dFecha, "yyyymmdd") & gsCodUser & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    Set rsAjuste = oAjuste.CargaAjusteReporteDet(dFecha, pnTipoAsiento, Mid(gsOpeCod, 3, 1))
    oBarraDet.Progress 2, TituloProgress, MensajeProgress, "", vbBlue
    If Mid(gsOpeCod, 3, 1) = "2" Then
        xlHoja1.Cells(1, 2) = "T.C"
        xlHoja1.Cells(1, 3) = Format(oCambio.EmiteTipoCambio(DateAdd("d", 1, dFecha), TCFijoDia), "#,##0.0000")
        xlHoja1.Range("B1:C1").Interior.ColorIndex = 39
    End If
    xlHoja1.Cells(2, 2) = "Nro"
    ExcelCuadro xlHoja1, 2, 2, 2, 2
    xlHoja1.Cells(2, 3) = "Crédito"
    ExcelCuadro xlHoja1, 3, 2, 3, 2
    xlHoja1.Cells(2, 4) = "Cliente"
    ExcelCuadro xlHoja1, 4, 2, 4, 2
    xlHoja1.Cells(2, 5) = "SaldoConsol"
    ExcelCuadro xlHoja1, 5, 2, 5, 2
    xlHoja1.Cells(2, 6) = "SaldoTC"
    ExcelCuadro xlHoja1, 6, 2, 6, 2
    xlHoja1.Cells(2, 7) = "Devengado"
    ExcelCuadro xlHoja1, 7, 2, 7, 2
    xlHoja1.Cells(2, 8) = "Suspenso"
    ExcelCuadro xlHoja1, 8, 2, 8, 2
    xlHoja1.Cells(2, 9) = "Moneda"
    ExcelCuadro xlHoja1, 9, 2, 9, 2
    xlHoja1.Cells(2, 10) = "TipoProducto"
    ExcelCuadro xlHoja1, 10, 2, 10, 2
    xlHoja1.Cells(2, 11) = "TipoCredito"
    ExcelCuadro xlHoja1, 11, 2, 11, 2
    xlHoja1.Cells(2, 12) = "TipoReprogramado"
    ExcelCuadro xlHoja1, 12, 2, 12, 2
    xlHoja1.Cells(2, 13) = "Mayorista"
    ExcelCuadro xlHoja1, 13, 2, 13, 2
    xlHoja1.Cells(2, 14) = "FecUltRepro"
    ExcelCuadro xlHoja1, 14, 2, 14, 2
    xlHoja1.Cells(2, 15) = "Agencia"
    ExcelCuadro xlHoja1, 15, 2, 15, 2
    xlHoja1.Cells(2, 16) = "Estado Cierre"
    ExcelCuadro xlHoja1, 16, 2, 16, 2
    
    xlHoja1.Range("B3").CopyFromRecordset rsAjuste
    xlHoja1.Range("B2:P2").Interior.Color = RGB(153, 153, 255)
    xlHoja1.Range("B1:P2").Font.Bold = True
    xlHoja1.Range("B2:P2").HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(rsAjuste.RecordCount + 2, 16)).Font.Name = "Calibri"
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(rsAjuste.RecordCount + 2, 16)).Font.Size = 10
    
    xlHoja1.Range("A1:A1").ColumnWidth = 4  '"A"
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(rsAjuste.RecordCount + 2, 3)).EntireColumn.AutoFit
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(rsAjuste.RecordCount + 2, 3)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(3, 4), xlHoja1.Cells(rsAjuste.RecordCount + 2, 4)).ColumnWidth = 32
    xlHoja1.Range(xlHoja1.Cells(3, 5), xlHoja1.Cells(rsAjuste.RecordCount + 2, 8)).HorizontalAlignment = xlRight
    xlHoja1.Range(xlHoja1.Cells(3, 5), xlHoja1.Cells(rsAjuste.RecordCount + 2, 8)).Style = "Comma"
    
    xlHoja1.Range(xlHoja1.Cells(3, 5), xlHoja1.Cells(rsAjuste.RecordCount + 2, 8)).NumberFormat = "#,##0.00"
    Do While Not rsAjuste.EOF
      xlHoja1.Cells(nLineas, 14) = Format(rsAjuste!FechaUltRepro, "mm/dd/yyyy")
      nLineas = nLineas + 1
      rsAjuste.MoveNext
    Loop
    xlHoja1.Range(xlHoja1.Cells(2, 9), xlHoja1.Cells(rsAjuste.RecordCount + 2, 16)).EntireColumn.AutoFit
    xlHoja1.Range(xlHoja1.Cells(3, 9), xlHoja1.Cells(rsAjuste.RecordCount + 2, 16)).HorizontalAlignment = xlCenter
    xlHoja1.Name = "ConsolidadoReprogramado"
    Set rsAjuste = Nothing

ElseIf pnTipoAsiento = 32 Then 'DifCancel 'NAGL 202008 Según Acta N°063-2020
    lsArchivoAGuardar = "\spooler\DiferidosCreditosCancelados" & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & Format(dFecha, "yyyymmdd") & gsCodUser & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    'Set rsAjuste = oAjuste.CargaAjusteReporteDet(dFecha, pnTipoAsiento, gsOpeCod)'Comentado by NAGL 202102
    If Not (rsAjuste.BOF Or rsAjuste.EOF) Then
        nRep = nRep + 1
        nLineas = 4
        nLineasReg = rsAjuste.RecordCount
        psDescripDet = "INTERESES DIFERIDOS EN CRÉDITOS CANCELADOS"
    End If
    oBarraDet.Progress 1, TituloProgress, MensajeProgress, "", vbBlue
    'Set rsAjusteSubProc = oAjuste.CargaAjusteReporteDet(dFecha, pnTipoAsiento, gsOpeCod, "SNeg")
    If Not (rsAjusteSubProc.BOF Or rsAjusteSubProc.EOF) Then
        nRep = nRep + 1
        If rsAjuste.RecordCount = 0 Then
            nLineas = 4
            nLineasReg = rsAjusteSubProc.RecordCount
            psDescripDet = "OTROS - DIFERIDOS SALDOS NEGATIVOS"
            Set rsAjuste = rsAjusteSubProc
        End If
    End If
    oBarraDet.Progress 2, TituloProgress, MensajeProgress, "", vbBlue
    Do While i <= nRep
        xlHoja1.Cells(nLineas - 2, 2) = psDescripDet
        xlHoja1.Cells(nLineas - 1, 2) = "Fecha"
        xlHoja1.Cells(nLineas - 1, 3) = "Crédito"
        xlHoja1.Cells(nLineas - 1, 4) = "Cód.Ope"
        xlHoja1.Cells(nLineas - 1, 5) = "Cuenta Cont."
        xlHoja1.Cells(nLineas - 1, 6) = "nIntDifAmp"
        
        xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).CopyFromRecordset rsAjuste
        
        xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 2), xlHoja1.Cells(nLineas - 1, 6)).Interior.ColorIndex = 3
        xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineas - 1, 6)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineas - 1, 6)).HorizontalAlignment = xlCenter
        
        xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineas - 2, 6)).Font.Color = vbBlack
        xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 2), xlHoja1.Cells(nLineas - 1, 6)).Font.Color = vbWhite
        xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineasReg + 3, 6)).Font.Name = "Calibri"
        xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineasReg + 3, 6)).Font.Size = 11
        
        xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 1), xlHoja1.Cells(nLineas - 2, 1)).ColumnWidth = 4  '"A"
        xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineasReg + 3, 2)).ColumnWidth = 10
        xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 3), xlHoja1.Cells(nLineasReg + 3, 3)).EntireColumn.AutoFit
        xlHoja1.Range(xlHoja1.Cells(nLineas, 4), xlHoja1.Cells(nLineasReg + 3, 4)).ColumnWidth = 10
        xlHoja1.Range(xlHoja1.Cells(nLineas, 5), xlHoja1.Cells(nLineasReg + 3, 5)).ColumnWidth = 12
        xlHoja1.Range(xlHoja1.Cells(nLineas, 6), xlHoja1.Cells(nLineasReg + 3, 6)).EntireColumn.AutoFit
        
        xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineas - 2, 6)).Merge True
        xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineasReg + 3, 2)).HorizontalAlignment = xlCenter
        xlHoja1.Range(xlHoja1.Cells(nLineas, 4), xlHoja1.Cells(nLineasReg + 3, 4)).HorizontalAlignment = xlCenter
        xlHoja1.Range(xlHoja1.Cells(nLineas, 6), xlHoja1.Cells(nLineasReg + 3, 6)).HorizontalAlignment = xlRight
        
        xlHoja1.Range(xlHoja1.Cells(nLineas, 6), xlHoja1.Cells(nLineasReg + 3, 6)).Style = "Comma"
        xlHoja1.Range(xlHoja1.Cells(nLineas, 6), xlHoja1.Cells(nLineasReg + 3, 6)).NumberFormat = "#,##0.00"
        
        xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineasReg + 3, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineasReg + 3, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineasReg + 3, 6)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineasReg + 3, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineasReg + 3, 6)).Borders(xlEdgeRight).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineasReg + 3, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        If nRep = 2 Then
            nLineas = nLineasReg + 8
            nLineasReg = nLineasReg + rsAjusteSubProc.RecordCount + 4
            psDescripDet = "OTROS - DIFERIDOS SALDOS NEGATIVOS"
            Set rsAjuste = rsAjusteSubProc
        End If
        i = i + 1
    Loop
    xlHoja1.Name = "ASIENTO_DIF"
    Set rsAjuste = Nothing
    Set rsAjusteSubProc = Nothing
    
ElseIf pnTipoAsiento = 33 Then 'NivCuentasOrden NAGL 202008 Según Acta N°063-2020
    lsArchivoAGuardar = "\spooler\NivelaciónCuentasOrden" & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & Format(dFecha, "yyyymmdd") & gsCodUser & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    oBarraDet.Progress 1, TituloProgress, MensajeProgress, "", vbBlue
    'Set rsAjuste = oAjuste.CargaAjusteReporteDet(dFecha, pnTipoAsiento, gsOpeCod)'Comentado by NAGL 202102
    nLineas = 4
    nLineasReg = rsAjuste.RecordCount
    psDescripDet = "CUENTAS DE ORDEN (SALDO CONTABLE - RCD)"
    oBarraDet.Progress 2, TituloProgress, MensajeProgress, "", vbBlue
    
    xlHoja1.Cells(nLineas - 2, 2) = psDescripDet
    xlHoja1.Cells(nLineas - 1, 2) = "Fecha"
    xlHoja1.Cells(nLineas - 1, 3) = "Crédito"
    xlHoja1.Cells(nLineas - 1, 4) = "Cuenta Cont."
    xlHoja1.Cells(nLineas - 1, 5) = "Cód.Ope"
    xlHoja1.Cells(nLineas - 1, 6) = "SaldoMesAnt."
    xlHoja1.Cells(nLineas - 1, 7) = "SaldoAsiento"
    xlHoja1.Cells(nLineas - 1, 8) = "SaldoContable"
    xlHoja1.Cells(nLineas - 1, 9) = "SaldoCartera_RCD"
    xlHoja1.Cells(nLineas - 1, 10) = "Diferencia(I-H)"
    
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).CopyFromRecordset rsAjuste
    
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 2), xlHoja1.Cells(nLineas - 1, 10)).Interior.ColorIndex = 3
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineas - 1, 10)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineas - 1, 10)).HorizontalAlignment = xlCenter
    
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineas - 2, 10)).Font.Color = vbBlack
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 2), xlHoja1.Cells(nLineas - 1, 10)).Font.Color = vbWhite
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineasReg + 3, 10)).Font.Name = "Calibri"
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineasReg + 3, 10)).Font.Size = 11
    
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 1), xlHoja1.Cells(nLineas - 2, 1)).ColumnWidth = 4  '"A"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineasReg + 3, 2)).ColumnWidth = 10
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 3), xlHoja1.Cells(nLineasReg + 3, 3)).EntireColumn.AutoFit
    xlHoja1.Range(xlHoja1.Cells(nLineas, 4), xlHoja1.Cells(nLineasReg + 3, 4)).ColumnWidth = 12
    xlHoja1.Range(xlHoja1.Cells(nLineas, 5), xlHoja1.Cells(nLineasReg + 3, 5)).ColumnWidth = 10
    xlHoja1.Range(xlHoja1.Cells(nLineas, 6), xlHoja1.Cells(nLineasReg + 3, 10)).EntireColumn.AutoFit
    
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineas - 2, 10)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineasReg + 3, 2)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(nLineas, 5), xlHoja1.Cells(nLineasReg + 3, 5)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(nLineas, 6), xlHoja1.Cells(nLineasReg + 3, 10)).HorizontalAlignment = xlRight
    xlHoja1.Range(xlHoja1.Cells(nLineas, 10), xlHoja1.Cells(nLineasReg + 3, 10)).Font.Bold = True
    
    xlHoja1.Range(xlHoja1.Cells(nLineas, 6), xlHoja1.Cells(nLineasReg + 3, 10)).Style = "Comma"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 6), xlHoja1.Cells(nLineasReg + 3, 10)).NumberFormat = "#,##0.00"
    
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineasReg + 3, 10)).Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineasReg + 3, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineasReg + 3, 10)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineasReg + 3, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineasReg + 3, 10)).Borders(xlEdgeRight).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 2), xlHoja1.Cells(nLineasReg + 3, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        
    xlHoja1.Name = "ASIENTO_NIV_CUENTAS_ORDEN"
    Set rsAjuste = Nothing

ElseIf pnTipoAsiento = 34 Then 'Provisión de Cartera Reprogramada NAGL 202102 Según Acta N°017-2021
    lsArchivoAGuardar = "\spooler\ProvCarteraReprogCOVID_" & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & Format(dFecha, "yyyymmdd") & gsCodUser & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    oBarraDet.Progress 1, TituloProgress, MensajeProgress, "", vbBlue
    Set xlHoja1 = oExcel.Worksheets.Add
    Set xlHoja1 = oBook.Sheets.Item(1)
    nLineas = 3
    pnCol = 17
    nLineasReg = rsAjuste.RecordCount
    
    Set rsProc1 = oAjuste.CargaListaFechasColReprogCOVID(dFecha, "ParamFec")
    psDescripDet = "REPORTE DE CRÉDITOS REPROGRAMADOS - COVID AL " & Format(dFecha, "dd/mm/yyyy") & " (CON EVALUACIÓN DE PAGOS DE " & rsProc1!dFecIni & " A " & rsProc1!dFecFin & ")"
    psFecActual = rsProc1!dFecActual
    xlHoja1.Name = "BD_" & psFecActual
    Set rsProc1 = Nothing
    oBarraDet.Progress 2, TituloProgress, MensajeProgress, "", vbBlue
    
    xlHoja1.Cells(nLineas - 2, 1) = psDescripDet
    xlHoja1.Cells(nLineas - 1, 1) = "Nro"
    xlHoja1.Cells(nLineas - 1, 2) = "Crédito"
    xlHoja1.Cells(nLineas - 1, 3) = "Cliente"
    xlHoja1.Cells(nLineas - 1, 4) = "Moneda"
    xlHoja1.Cells(nLineas - 1, 5) = "Saldo_FechaReprog"
    xlHoja1.Cells(nLineas - 1, 6) = "Int.Deveng"
    xlHoja1.Cells(nLineas - 1, 7) = "Int.Deveng_Prov"
    xlHoja1.Cells(nLineas - 1, 8) = "Provisión"
    xlHoja1.Cells(nLineas - 1, 9) = "SaldoCap.Ant"
    xlHoja1.Cells(nLineas - 1, 10) = "SaldoCap.Actual"
    xlHoja1.Cells(nLineas - 1, 11) = "Calif.Ant"
    xlHoja1.Cells(nLineas - 1, 12) = "Calif.Actual"
    xlHoja1.Cells(nLineas - 1, 13) = "Gar.Pref"
    xlHoja1.Cells(nLineas - 1, 14) = "Gar.No Pref"
    xlHoja1.Cells(nLineas - 1, 15) = "Gar.Autol"
    xlHoja1.Cells(nLineas - 1, 16) = "Nro Cuotas Pag."
    xlHoja1.Cells(nLineas - 1, 17) = "Importe Cuota"
    xlHoja1.Cells(nLineas - 1, 42) = "PagoCapitalCompleto"
    xlHoja1.Cells(nLineas - 1, 43) = "SC_CtaProv"
    xlHoja1.Cells(nLineas - 1, 44) = "SC_ContraPartida"
    xlHoja1.Cells(nLineas - 1, 45) = "ID_CtaProv"
    xlHoja1.Cells(nLineas - 1, 46) = "ID_ContraPartida"
    
    Set rsProc1 = oAjuste.CargaListaFechasColReprogCOVID(dFecha, "ListFec")
    Do While Not rsProc1.EOF
        xlHoja1.Cells(nLineas - 2, rsProc1!nPos) = rsProc1!cMES
        xlHoja1.Range(xlHoja1.Cells(nLineas - 2, rsProc1!nPos), xlHoja1.Cells(nLineas - 2, rsProc1!nPos + 3)).HorizontalAlignment = xlCenter
        xlHoja1.Range(xlHoja1.Cells(nLineas - 2, rsProc1!nPos), xlHoja1.Cells(nLineas - 2, rsProc1!nPos + 3)).Merge True
        xlHoja1.Range(xlHoja1.Cells(nLineas - 2, rsProc1!nPos), xlHoja1.Cells(nLineas - 2, rsProc1!nPos + 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(nLineas - 2, rsProc1!nPos), xlHoja1.Cells(nLineas - 2, rsProc1!nPos + 3)).Borders(xlEdgeRight).LineStyle = xlContinuous
        rsProc1.MoveNext
    Loop
    
    Set rsProc2 = oAjuste.CargaListaFechasColReprogCOVID(dFecha, "ColumDet")
    Do While Not rsProc2.EOF
        xlHoja1.Cells(nLineas - 1, pnCol + rsProc2!nOrdenDet) = rsProc2!cDescripColum
        rsProc2.MoveNext
    Loop
    
    xlHoja1.Cells(nLineas - 2, 43) = "Provisión Saldo Capital"
    xlHoja1.Cells(nLineas - 2, 45) = "Provisión Interés Devengado"
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 43), xlHoja1.Cells(nLineas - 2, 46)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 43), xlHoja1.Cells(nLineas - 2, 44)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 45), xlHoja1.Cells(nLineas - 2, 46)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 43), xlHoja1.Cells(nLineas - 2, 44)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 43), xlHoja1.Cells(nLineas - 2, 44)).Borders(xlEdgeRight).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 45), xlHoja1.Cells(nLineas - 2, 46)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 45), xlHoja1.Cells(nLineas - 2, 46)).Borders(xlEdgeRight).LineStyle = xlContinuous
     
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).CopyFromRecordset rsAjuste
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 1), xlHoja1.Cells(nLineas - 1, 17)).Interior.Color = 16764057
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 18), xlHoja1.Cells(nLineas - 1, 41)).Interior.ColorIndex = 24
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 43), xlHoja1.Cells(nLineas - 2, 46)).Interior.ColorIndex = 24
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 42), xlHoja1.Cells(nLineas - 1, 46)).Interior.Color = 16764057
    
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 1), xlHoja1.Cells(nLineas - 1, 46)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 1), xlHoja1.Cells(nLineas - 2, 17)).HorizontalAlignment = xlLeft
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 1), xlHoja1.Cells(nLineas - 2, 1)).RowHeight = 13.5
    
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 1), xlHoja1.Cells(nLineas - 1, 46)).Font.Color = vbBlack
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 1), xlHoja1.Cells(nLineasReg + 2, 46)).Font.Name = "Calibri"
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 1), xlHoja1.Cells(nLineasReg + 2, 46)).Font.Size = 9
    
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 1), xlHoja1.Cells(nLineas - 1, 46)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineasReg + 2, 1)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(nLineas, 4), xlHoja1.Cells(nLineasReg + 2, 4)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(nLineas, 11), xlHoja1.Cells(nLineasReg + 2, 12)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(nLineas, 42), xlHoja1.Cells(nLineasReg + 2, 42)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 1), xlHoja1.Cells(nLineasReg + 2, 46)).VerticalAlignment = xlCenter
    
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineasReg + 2, 1)).ColumnWidth = 7  '"A"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineasReg + 2, 2)).EntireColumn.AutoFit
    xlHoja1.Range(xlHoja1.Cells(nLineas, 3), xlHoja1.Cells(nLineasReg + 2, 3)).ColumnWidth = 22
    xlHoja1.Range(xlHoja1.Cells(nLineas, 4), xlHoja1.Cells(nLineasReg + 2, 46)).EntireColumn.AutoFit
    
    xlHoja1.Range(xlHoja1.Cells(nLineas, 5), xlHoja1.Cells(nLineasReg + 3, 10)).Style = "Comma"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 13), xlHoja1.Cells(nLineasReg + 3, 15)).Style = "Comma"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 17), xlHoja1.Cells(nLineasReg + 3, 41)).Style = "Comma"
    
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 1), xlHoja1.Cells(nLineas - 2, 17)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 1), xlHoja1.Cells(nLineas - 1, 46)).Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 1), xlHoja1.Cells(nLineas - 1, 46)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 1), xlHoja1.Cells(nLineas - 1, 46)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 1), xlHoja1.Cells(nLineas - 1, 46)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 1), xlHoja1.Cells(nLineas - 1, 46)).Borders(xlEdgeRight).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 1), xlHoja1.Cells(nLineas - 1, 46)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Set xlHoja1 = oBook.Sheets.Item(2)
    xlHoja1.Name = "PRE_ASIENTO"
    Set rsAjuste = oAjuste.AjusteInteresesColocaciones(pnTipoAsiento, Format(dFecha, gsFormatoFecha), Mid(gsOpeCod, 3, 1), sCtaDebe, nVal(txtTipCambio), sCtaDebe, , , , "RepDet")
    nLineasReg = rsAjuste.RecordCount
    psDescripDet = "REPORTE PRE_ASIENTOS - PROVISIONES COVID AL " & Format(dFecha, "dd/mm/yyyy")
    
    xlHoja1.Cells(nLineas - 2, 1) = psDescripDet
    xlHoja1.Cells(nLineas - 1, 1) = "Nro"
    xlHoja1.Cells(nLineas - 1, 2) = "Sección"
    xlHoja1.Cells(nLineas - 1, 3) = "Cuenta Contable"
    xlHoja1.Cells(nLineas - 1, 4) = "Descripcion"
    xlHoja1.Cells(nLineas - 1, 5) = "DEBE"
    xlHoja1.Cells(nLineas - 1, 6) = "HABER"
   
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).CopyFromRecordset rsAjuste
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 1), xlHoja1.Cells(nLineas - 1, 6)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 1), xlHoja1.Cells(nLineas - 1, 6)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 1), xlHoja1.Cells(nLineas - 2, 6)).HorizontalAlignment = xlLeft
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 1), xlHoja1.Cells(nLineas - 2, 6)).Font.Color = vbBlack
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 1), xlHoja1.Cells(nLineas - 1, 6)).Font.Color = vbWhite
    
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 1), xlHoja1.Cells(nLineas - 2, 6)).RowHeight = 13.5
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 1), xlHoja1.Cells(nLineas - 1, 6)).Interior.ColorIndex = 3
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 1), xlHoja1.Cells(nLineasReg + 2, 6)).Font.Name = "Calibri"
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 1), xlHoja1.Cells(nLineas - 1, 6)).Font.Size = 10
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineasReg + 2, 6)).Font.Size = 9
    
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineasReg + 2, 2)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(nLineas, 3), xlHoja1.Cells(nLineasReg + 2, 4)).HorizontalAlignment = xlLeft
    xlHoja1.Range(xlHoja1.Cells(nLineas, 5), xlHoja1.Cells(nLineasReg + 2, 6)).HorizontalAlignment = xlRight
    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 1), xlHoja1.Cells(nLineasReg + 2, 6)).VerticalAlignment = xlCenter

    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineasReg + 2, 1)).ColumnWidth = 7  '"A"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineasReg + 2, 6)).EntireColumn.AutoFit
    xlHoja1.Range(xlHoja1.Cells(nLineas, 5), xlHoja1.Cells(nLineasReg + 3, 6)).Style = "Comma"

    xlHoja1.Range(xlHoja1.Cells(nLineas - 2, 1), xlHoja1.Cells(nLineas - 2, 6)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 1), xlHoja1.Cells(nLineas - 1, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 1), xlHoja1.Cells(nLineas - 1, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 1), xlHoja1.Cells(nLineas - 1, 6)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 1), xlHoja1.Cells(nLineas - 1, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 1), xlHoja1.Cells(nLineas - 1, 6)).Borders(xlEdgeRight).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 1), xlHoja1.Cells(nLineas - 1, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Set rsAjuste = Nothing
    Set rsProc1 = Nothing
    Set rsProc2 = Nothing
End If

oBarraDet.Progress 3, TituloProgress, MensajeProgress, "", vbBlue
oBarraDet.CloseForm Me
Set oBarraDet = Nothing

xlHoja1.SaveAs App.path & lsArchivoAGuardar
oExcel.Visible = True
oExcel.Windows(1).Visible = True
Set oExcel = Nothing
Set oBook = Nothing
Set xlHoja1 = Nothing
End Sub 'NAGL 202007 Según ACTA N°049-2020

Private Sub GeneraAsientoOtrosProcesos(pnTipoAsiento As Integer)
Dim oAjuste As New DAjusteCont
Dim oMov  As New DMov
Dim rsAjuste As New ADODB.Recordset
Dim nMes As Integer
Dim nAnio As Integer
Dim dFecha As Date
Dim lsArchivoAGuardar As String, lsTituloImpre As String, psFecha As String
Dim psDescrip As String, psMovNroImp As String

Dim oBarraDet As clsProgressBar
Dim TituloProgress As String
Dim MensajeProgress As String
Set oBarraDet = New clsProgressBar
oBarraDet.ShowForm Me
oBarraDet.Max = 3

nMes = cboMes.ListIndex + 1
nAnio = txtAnio
dFecha = DateAdd("m", 1, "01/" & Format(nMes, "00") & "/" & Format(txtAnio, "0000")) - 1
psFecha = Format(dFecha, "yyyymmdd")

If pnTipoAsiento = 32 Then
   psDescrip = "Diferidos en Créditos Cancelados"
ElseIf pnTipoAsiento = 33 Then
   psDescrip = "Nivelación Cuentas de Orden (Saldo Contable - RCD)"
End If

psDescrip = "Asiento de Cierre " & psDescrip & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME")
oBarraDet.Progress 0, psDescrip, "GENERANDO EL ARCHIVO", "", vbBlue
TituloProgress = psDescrip
MensajeProgress = "GENERANDO EL ARCHIVO"
psMovNroImp = ""

If pnTipoAsiento = 32 Or pnTipoAsiento = 33 Then
   lsTituloImpre = "ASIENTO DE INTERESES"
   oBarraDet.Progress 1, TituloProgress, MensajeProgress, "", vbBlue
   gsMovNro = oMov.GeneraMovNro(dFecha, gsCodAge, gsCodUser)
   psMovNroImp = oAjuste.CargaAsientosOtrosProcesos(dFecha, gsOpeCod, gsMovNro, pnTipoAsiento)
   If psMovNroImp <> "" Then
        If Left(psMovNroImp, 8) = psFecha Then
            oBarraDet.Progress 2, TituloProgress, MensajeProgress, "", vbBlue
             If dFecha < gdFecSis Then
               oMov.ActualizaSaldoMovimiento psMovNroImp, "+"
            End If
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & "Se Grabo el Asiento Contable al Cierre " & dFecha & " con el Tipo de Cambio Fijo : " & txtTipCambio.Text _
            & " Tpo Cambio venta : " & txtTipCambioVenta & " Tpo Cambio Venta : " & txtTipCambioCompra
            Set objPista = Nothing
            '*******
            oBarraDet.Progress 3, TituloProgress, MensajeProgress, "", vbBlue
            oBarraDet.CloseForm Me
            ImprimeAsientoContable psMovNroImp, , , , , , , , , , , , 1, lsTituloImpre
        Else
            oBarraDet.CloseForm Me
            EnviaPrevio psMovNroImp, "OBSERVACIÓN CONTABLE DE INTERESES DIFERIDOS EN CRÉDITOS CANCELADOS", gnLinPage, False
        End If
   Else
        oBarraDet.CloseForm Me
   End If
End If
Me.Enabled = True
Set oBarraDet = Nothing
Exit Sub
End Sub 'NAGL 202008 Según ACTA N°063-2020

VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAjusteTipCambio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajuste de Tipo de Cambio"
   ClientHeight    =   2925
   ClientLeft      =   3180
   ClientTop       =   2610
   ClientWidth     =   4545
   Icon            =   "frmAjusteTipCambio.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1320
      Left            =   180
      TabIndex        =   6
      Top             =   1005
      Width           =   4155
      Begin VB.TextBox txtTipAnt 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2460
         MaxLength       =   16
         TabIndex        =   11
         Top             =   240
         Width           =   1395
      End
      Begin VB.TextBox txtTipCambio 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2460
         MaxLength       =   16
         TabIndex        =   2
         Top             =   645
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Cambio Anterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         TabIndex        =   9
         Top             =   270
         Width           =   2115
      End
      Begin VB.Label Label1 
         Caption         =   "Nuevo Tipo de Cambio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   165
         TabIndex        =   7
         Top             =   675
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   360
      Left            =   1620
      TabIndex        =   3
      Top             =   2460
      Width           =   1200
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   2850
      TabIndex        =   4
      Top             =   2460
      Width           =   1200
   End
   Begin VB.Frame Frame3 
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
      Left            =   180
      TabIndex        =   5
      Top             =   120
      Width           =   4170
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   1
         Top             =   300
         Width           =   855
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmAjusteTipCambio.frx":030A
         Left            =   690
         List            =   "frmAjusteTipCambio.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   315
         Width           =   1830
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   2640
         TabIndex        =   12
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   390
         Width           =   390
      End
   End
   Begin RichTextLib.RichTextBox rtxt 
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Top             =   1470
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAjusteTipCambio.frx":039A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAjusteTipCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs   As New ADODB.Recordset
Dim dFecTipCam As Date
Dim oBarra As clsProgressBar

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Private Function GetFecha() As Date
Dim nMes As Integer
    nMes = CboMes.ListIndex + 1
    GetFecha = DateAdd("m", 1, CDate("01/" & Format(nMes, "00") & "/" & Format(txtAnio, "0000"))) - 1
End Function
Private Sub CboMes_Click()
Dim nTipCambio As Currency
Dim ind As Integer
    nTipCambio = gnTipCambio
    dFecTipCam = GetFecha()
    Call GetTipCambio(dFecTipCam)
    txtTipAnt = Format(gnTipCambio, "#0.00###")
    gnTipCambio = nTipCambio
End Sub

Private Sub cmdProcesar_Click()
Dim lbActivaTrans As Boolean
Dim lsMsgErr      As String
Dim nItem      As Long
Dim nMes As Integer
Dim sFecha As String
Dim lsHoraAjuste As String
Dim nImporte As Currency
Dim nDif As Currency
Dim nDiferenciaCta13 As Currency

Dim oCont As New NContFunciones
Dim oMov  As New DMov
Dim rsO   As ADODB.Recordset
Set oBarra = New clsProgressBar

If nVal(txtTipAnt) = 0 Then
    MsgBox "Falta definir Tipo de Cambio del mes ", vbInformation, "¡Aviso!"
    If txtTipAnt.Enabled Then
        txtTipAnt.SetFocus
    End If
    Exit Sub
End If
If nVal(txtTipCambio) = 0 Then
    MsgBox "Falta definir Tipo de Cambio de Ajuste", vbInformation, "¡Aviso!"
    If txtTipCambio.Enabled Then
        txtTipCambio.SetFocus
    End If
    Exit Sub
End If

sFecha = GetFecha()
If Not oCont.PermiteModificarAsiento(Format(sFecha, gsFormatoMovFecha), False) Then
   MsgBox "Mes ya cerrado. Imposible generar Asiento de Ajuste", vbInformation, "!Aviso!"
   Exit Sub
End If
       
If MsgBox(" ¿ Seguro de Generar Asiento de Ajuste por Tipo de Cambio ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
   Exit Sub
End If
Screen.MousePointer = 11
oBarra.ShowForm frmMdiMain
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.Max = 1
oBarra.Progress 0, "Ajuste de Tipo de Cambio", "", "Validando datos...", vbBlue
lsHoraAjuste = "235900"
gsMovNro = Format(sFecha, gsFormatoMovFecha)
If oCont.ExisteMovimiento(gsMovNro, gsOpeCod) Then
    oBarra.CloseForm frmMdiMain
    MsgBox "Ajuste de Tipo de Cambio para este mes ya fue Realizado", vbInformation, "Aviso"
    Screen.MousePointer = 0
    Exit Sub
End If

gsMovNro = Format(sFecha, gsFormatoMovFecha) & lsHoraAjuste & gsCodCMAC & gsCodAge & "00" & gsCodUser
gsMovNro = oCont.GeneraMovNro(sFecha, , , gsMovNro)
oBarra.Progress 0, "Ajuste de Tipo de Cambio", "", "Obteniendo datos...", vbBlue

Dim oFun   As New ncontasientos
Set rs = oFun.GetAsientoAjusteTipoCambio(Format(sFecha, gsFormatoFecha))
Set rsO = oFun.GetAsientoAjusteTipoCambio(Format(sFecha, gsFormatoFecha), False, "8[12]")
If rs.EOF Or rs.BOF Then
   oBarra.CloseForm frmMdiMain
   MsgBox "No existen Cuentas por Ajustar...", vbInformation, "Advertencia"
   RSClose rs
   Screen.MousePointer = 0
   Exit Sub
End If

gsGlosa = "Asiento de Ajuste por Tipo de Cambio del Mes de " + CboMes.Text & "  Cambio de Ajuste : " & txtTipAnt & "  " & txtTipCambio
lbActivaTrans = True
oMov.BeginTrans
oMov.InsertaMov gsMovNro, gsOpeCod, gsGlosa, gMovEstContabMovContable, gMovFlagVigente
gnMovNro = oMov.GetnMovNro(gsMovNro)
oMov.InsertaMovCont gnMovNro, 0, 0, ""

nItem = 0
oBarra.Max = rs.RecordCount + rsO.RecordCount
Do While Not rs.EOF
   oBarra.Progress rs.Bookmark, "Ajuste de Tipo de Cambio", "", "Generando Ajuste...", vbBlue
   If rs!nCtaSaldoImporte <> 0 Then
      nImporte = Round((rs!nCtaSaldoImporte / Val(Format(txtTipAnt, "##0.00000"))) * Val(Format(txtTipCambio, "##0.00000")) - rs!nCtaSaldoImporte, 2)
      If Mid(rs!cCtaCaracter, 1, 1) = "A" Then 'Acreedora
         nImporte = nImporte * -1
      End If
      'ALPA 20090528********************************
      'nDif = nDif + nImporte
      If Mid(rs!cCtaContCod, 1, 2) <> "13" Then
        nDif = nDif + nImporte
      Else
        nDiferenciaCta13 = nDiferenciaCta13 + nImporte
      End If
      '*********************************************
      If nImporte <> 0 Then
         nItem = nItem + 1
         oMov.InsertaMovCta gnMovNro, nItem, rs!cCtaContCod, nImporte
      End If
   End If
   rs.MoveNext
Loop
If nDif <> 0 Then
    nItem = nItem + 1
    If nDif < 0 Then   'Perdida
       oMov.InsertaMovCta gnMovNro, nItem, gcConvMEDAjTC, nDif * -1
    Else               'Ganancia
       oMov.InsertaMovCta gnMovNro, nItem, gcConvMESAjTC, nDif * -1
    End If
End If
nDif = 0
'ALPA 20090528*******************
If nDiferenciaCta13 <> 0 Then
    nItem = nItem + 1
    If nDiferenciaCta13 < 0 Then   'Perdida
        'JUEZ 20130218 *********************************************************
        'oMov.InsertaMovCta gnMovNro, nItem, "51280403", nDiferenciaCta13 * -1
        oMov.InsertaMovCta gnMovNro, nItem, "41280403", nDiferenciaCta13 * -1
    Else               'Ganancia
        'oMov.InsertaMovCta gnMovNro, nItem, "41280403", nDiferenciaCta13 * -1
        oMov.InsertaMovCta gnMovNro, nItem, "51280403", nDiferenciaCta13 * -1
        'END JUEZ **************************************************************
    End If
End If
nDiferenciaCta13 = 0
'*********************************

If Not rsO.EOF Then
    Do While Not rsO.EOF
       oBarra.Progress rs.RecordCount + rsO.Bookmark, "Ajuste de Tipo de Cambio", "", "Generando Ajuste...", vbBlue
       If rsO!nCtaSaldoImporte <> 0 Then
          nImporte = Round((rsO!nCtaSaldoImporte / Val(Format(txtTipAnt, "##0.00000"))) * Val(Format(txtTipCambio, "##0.00000")) - rsO!nCtaSaldoImporte, 2)
          If Mid(rsO!cCtaCaracter, 1, 1) = "A" Then 'Acreedora
             nImporte = nImporte * -1
          End If
          nDif = nDif + nImporte
          If nImporte <> 0 Then
             nItem = nItem + 1
             oMov.InsertaMovCta gnMovNro, nItem, rsO!cCtaContCod, nImporte
          End If
       End If
       rsO.MoveNext
    Loop
End If
If nDif <> 0 Then
    nItem = nItem + 1
    oMov.InsertaMovCta gnMovNro, nItem, "8221", nDif * -1
End If

Set rsO = oFun.GetAsientoAjusteTipoCambio(Format(sFecha, gsFormatoFecha), False, "8[34]")
nDif = 0
If Not rsO.EOF Then
    Do While Not rsO.EOF
       oBarra.Progress rs.RecordCount + rsO.Bookmark, "Ajuste de Tipo de Cambio", "", "Generando Ajuste...", vbBlue
       If rsO!nCtaSaldoImporte <> 0 Then
          nImporte = Round((rsO!nCtaSaldoImporte / Val(Format(txtTipAnt, "##0.00000"))) * Val(Format(txtTipCambio, "##0.00000")) - rsO!nCtaSaldoImporte, 2)
          If Mid(rsO!cCtaCaracter, 1, 1) = "A" Then 'Acreedora
             nImporte = nImporte * -1
          End If
          nDif = nDif + nImporte
          If nImporte <> 0 Then
             nItem = nItem + 1
             oMov.InsertaMovCta gnMovNro, nItem, rsO!cCtaContCod, nImporte
          End If
       End If
       rsO.MoveNext
    Loop
End If
If nDif <> 0 Then
    nItem = nItem + 1
    oMov.InsertaMovCta gnMovNro, nItem, "8321", nDif * -1
End If

' 85 y 86
Set rsO = oFun.GetAsientoAjusteTipoCambio(Format(sFecha, gsFormatoFecha), False, "8[56]")
nDif = 0
If Not rsO.EOF Then
    Do While Not rsO.EOF
       oBarra.Progress rs.RecordCount + rsO.Bookmark, "Ajuste de Tipo de Cambio", "", "Generando Ajuste...", vbBlue
       If rsO!nCtaSaldoImporte <> 0 Then
          nImporte = Round((rsO!nCtaSaldoImporte / Val(Format(txtTipAnt, "##0.00000"))) * Val(Format(txtTipCambio, "##0.00000")) - rsO!nCtaSaldoImporte, 2)
          If Mid(rsO!cCtaCaracter, 1, 1) = "A" Then 'Acreedora
             nImporte = nImporte * -1
          End If
          nDif = nDif + nImporte
          If nImporte <> 0 Then
             nItem = nItem + 1
             oMov.InsertaMovCta gnMovNro, nItem, rsO!cCtaContCod, nImporte
          End If
       End If
       rsO.MoveNext
    Loop
End If
If nDif <> 0 Then
    nItem = nItem + 1
    oMov.InsertaMovCta gnMovNro, nItem, "8521", nDif * -1
End If
oMov.InsertaMovTpoCambio gnMovNro, Format(txtTipCambio, "########0.#####")
oMov.InsertaMovOtrosItem gnMovNro, 1, "TC1", Format(txtTipAnt, gsFormatoNumeroDato), ""
oMov.InsertaMovOtrosItem gnMovNro, 1, "TC2", Format(txtTipCambio, gsFormatoNumeroDato), ""


oMov.ActualizaSaldoMovimiento gsMovNro, "+"
'oMov.RollbackTrans
    
    'ARLO20170208
    Set objPista = New COMManejador.Pista
    gsOpeCod = LogPistaAjusteTpoCambio
    objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Se Ajusto el Tipo del Cambio | Mes : " & CboMes.Text & " | Año : " & txtAnio.Text _
    & " | Tipo Cambio Anterior : " & txtTipAnt.Text & " | Tipo Cambio Nuevo : " & txtTipCambio.Text
    Set objPista = Nothing
    
oMov.CommitTrans
lbActivaTrans = False
oBarra.CloseForm frmMdiMain
ImprimeAsientoContable gsMovNro, , , , , , , , , , , , 1
Screen.MousePointer = 0
ProcesarFin:
Set oCont = Nothing
Set oMov = Nothing
Set oBarra = Nothing
Exit Sub
ProcesarErr:
    lsMsgErr = Err.Description
    If lbActivaTrans Then
        oMov.RollbackTrans
        oBarra.CloseForm frmMdiMain
    End If
    MsgBox TextErr(lsMsgErr), vbInformation, "¡Aviso!"
   GoTo ProcesarFin
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
frmMdiMain.Enabled = False
txtAnio = Year(gdFecSis)
CboMes.ListIndex = Month(gdFecSis) - 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
frmMdiMain.Enabled = True
End Sub

Private Sub txtTipAnt_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTipAnt, KeyAscii, 14, 5)
If KeyAscii = 13 Then
   txtTipAnt = Format(txtTipAnt, "###,##0.00000")
   txtTipCambio.SetFocus
End If
End Sub

Private Sub txtTipCambio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTipCambio, KeyAscii, 14, 5)
If KeyAscii = 13 Then
   cmdProcesar.SetFocus
End If
End Sub

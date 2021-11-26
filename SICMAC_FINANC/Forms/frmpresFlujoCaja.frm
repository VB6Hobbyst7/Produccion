VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmpresFlujoCaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PRESUPUESTO: Flujo de Caja"
   ClientHeight    =   3735
   ClientLeft      =   1635
   ClientTop       =   1890
   ClientWidth     =   4050
   Icon            =   "frmpresFlujoCaja.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      ForeColor       =   &H00000080&
      Height          =   1365
      Left            =   180
      TabIndex        =   14
      Top             =   1395
      Width           =   3705
      Begin VB.TextBox txttipCambio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1980
         TabIndex        =   5
         Top             =   180
         Width           =   1290
      End
      Begin VB.ComboBox cboPlazo 
         Height          =   315
         ItemData        =   "frmpresFlujoCaja.frx":030A
         Left            =   1980
         List            =   "frmpresFlujoCaja.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   540
         Width           =   1320
      End
      Begin VB.CheckBox chkProyectado 
         Caption         =   "Proyectado"
         Height          =   255
         Left            =   1980
         TabIndex        =   8
         Top             =   990
         Width           =   1275
      End
      Begin VB.CheckBox chkEjecutado 
         Caption         =   "Ejecutado"
         Height          =   255
         Left            =   510
         TabIndex        =   7
         Top             =   990
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Cambio :"
         Height          =   195
         Left            =   480
         TabIndex        =   16
         Top             =   210
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Plazo                 :"
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Top             =   570
         Width           =   1200
      End
   End
   Begin VB.CommandButton cmdsalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   2100
      TabIndex        =   10
      Top             =   2895
      Width           =   1350
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   390
      Left            =   630
      TabIndex        =   9
      Top             =   2895
      Width           =   1350
   End
   Begin MSComctlLib.StatusBar barraEstado 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   13
      Top             =   3405
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12347
            MinWidth        =   12347
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ejecución de Presupuesto del :"
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
      Height          =   1140
      Left            =   180
      TabIndex        =   0
      Top             =   195
      Width           =   3705
      Begin VB.ComboBox cboAnioFin 
         Height          =   315
         Left            =   2745
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   660
         Width           =   825
      End
      Begin VB.ComboBox cboMesFin 
         Height          =   315
         ItemData        =   "frmpresFlujoCaja.frx":034C
         Left            =   810
         List            =   "frmpresFlujoCaja.frx":0374
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   645
         Width           =   1800
      End
      Begin VB.ComboBox cboAnioInI 
         Height          =   315
         Left            =   2745
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   825
      End
      Begin VB.ComboBox cboMesIni 
         Height          =   315
         ItemData        =   "frmpresFlujoCaja.frx":03DC
         Left            =   810
         List            =   "frmpresFlujoCaja.frx":0404
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   1800
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   135
         TabIndex        =   12
         Top             =   675
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   315
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmpresFlujoCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type ltFinanciamientos
    lnMes As Integer
    lnAnio As Integer
    lnDesembolso As Currency
    lnAmortizacion As Currency
    lnInteres As Currency
End Type
Dim ltFinanciamientos() As ltFinanciamientos
Dim ltEjecutado() As ltFinanciamientos
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim lsArchivo As String
Dim lbExcel As Boolean
Dim N As Integer, I  As Integer
Dim oCon As DConecta

Private Sub cboAnioFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     txtTipCambio.SetFocus
End If
End Sub
Private Sub cboAnioInI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cboMesFin.SetFocus
End If
End Sub
Private Sub cboMesFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cboAnioFin.SetFocus
End If
End Sub
Private Sub cboMesIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cboAnioInI.SetFocus
End If
End Sub

Private Sub cboPlazo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.chkEjecutado.SetFocus
End If
End Sub

Private Sub chkEjecutado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.chkProyectado.SetFocus
End If
End Sub

Private Sub cmdGenerar_Click()
Dim lsMsgErr As String
On Error GoTo ErrorGenerarFC
Me.cmdGenerar.Enabled = False
lbExcel = False
N = 0
If GeneraDatos(chkProyectado = vbChecked) Then
    GeneraReporteExcel
End If
If lsArchivo <> "" Then
    CargaArchivo lsArchivo, App.path & "\SPOOLER\"
End If
Me.cmdGenerar.Enabled = True
Exit Sub
ErrorGenerarFC:
    lsMsgErr = Err.Description
    Me.cmdGenerar.Enabled = True
    If lbExcel = True Then
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
    End If
    MsgBox TextErr(lsMsgErr), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim sql As String
Dim rs As New ADODB.Recordset
Dim I As Integer
Dim lnAnio As Integer
CentraForm Me

Set oCon = New DConecta
oCon.AbreConexion

cboPlazo.ListIndex = 0

txtTipCambio = gnTipCambio
For I = 0 To 20
    lnAnio = Year(gdFecSis) + I
    cboAnioFin.AddItem lnAnio
    cboAnioInI.AddItem lnAnio
Next I
Me.cboAnioInI.ListIndex = 0
Me.cboAnioFin.ListIndex = 0

cboMesIni.ListIndex = 0
cboMesFin.ListIndex = cboMesFin.ListCount - 1
RSClose rs

End Sub

Private Sub txtTipCambio_GotFocus()
fEnfoque txtTipCambio
End Sub
Private Sub txtTipCambio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTipCambio, KeyAscii)
If KeyAscii = 13 Then
    Me.cboPlazo.SetFocus
End If
End Sub
Private Sub txttipCambio_LostFocus()
If txtTipCambio = "" Then txtTipCambio = 0
txtTipCambio = Format(txtTipCambio, "#0.00")
End Sub

Private Function GeneraDatos(pbGetProyectado As Boolean) As Boolean
'*****************************************************************************************
'ltFinanciamientos(1, NumMeses) 'FINACIAMIENTO EXTERNO LARGO PLAZO
'ltFinanciamientos(4, NumMeses) 'FINACIAMIENTO EXTERNO CORTO PLAZO
'ltFinanciamientos(7, NumMeses) 'FINACIAMIENTO INTERNO LARGO PLAZO
'ltFinanciamientos(10,NumMeses) 'FINACIAMIENTO INTERNO CORTO PLAZO
'*****************************************************************************************
Dim lnNumMeses As Integer
Dim ldFechaIni As Date
Dim ldFechaFin As Date
Dim lnMesIni As Integer
Dim lnMesFin As Integer
Dim lnAnioIni As Integer
Dim lnAnioFin As Integer
Dim j As Integer
Dim lnMes As Integer
Dim lnAnio As Integer
Dim ldFecha As Date
Dim lnIntervMes As Integer

On Error GoTo ErrorGenerDatosFC

GeneraDatos = True

ldFechaIni = CDate("01/" & Format(cboMesIni.ListIndex + 1, "00") & "/" & Trim(cboAnioInI))
ldFechaFin = DateAdd("m", 1, CDate("01/" & Format(cboMesFin.ListIndex + 1, "00") & "/" & Trim(cboAnioFin))) - 1

If ldFechaIni > ldFechaFin Then
    MsgBox "Intervalos de Tiempos mal Ingresados por favor verifique", vbInformation, "Aviso"
    cboAnioInI.SetFocus
    GeneraDatos = False
    Exit Function
End If
 
lnNumMeses = Int((DateDiff("m", ldFechaIni, ldFechaFin) + 1) / Val(Right(cboPlazo, 1))) + IIf(Val(Right(cboPlazo, 1)) = 1, 0, 1)

ReDim ltFinanciamientos(4, lnNumMeses)
'INGRESANDO LOS MESES Y AÑOS RESPECTIVOS DE LOS FINANCIAMIENTOS
j = 0
ldFecha = ldFechaIni
lnIntervMes = Val(Right(Me.cboPlazo, 1))

For I = 0 To lnNumMeses - 1
    If I = 0 Then
        ldFecha = DateAdd("m", I, ldFechaIni)
    Else
        ldFecha = DateAdd("m", lnIntervMes * I, ldFechaIni) - IIf(Right(cboPlazo, 1) = "1", 0, 1)
        If Right(Me.cboPlazo, 1) = "3" And I = lnNumMeses - 1 Then
            ldFecha = DateAdd("m", lnIntervMes * I, ldFechaIni) - 1
        End If
    End If
    lnMes = Month(ldFecha)
    lnAnio = Year(ldFecha)
    j = j + 1
    ltFinanciamientos(1, j).lnMes = lnMes:     ltFinanciamientos(1, j).lnAnio = lnAnio
    ltFinanciamientos(2, j).lnMes = lnMes:     ltFinanciamientos(2, j).lnAnio = lnAnio
    ltFinanciamientos(3, j).lnMes = lnMes:     ltFinanciamientos(3, j).lnAnio = lnAnio
    ltFinanciamientos(4, j).lnMes = lnMes:     ltFinanciamientos(4, j).lnAnio = lnAnio
Next I

GeneraEjecutado cboMesIni.ListIndex + 1, cboMesFin.ListIndex + 1, Trim(Str(lnAnio - 1))

If Right(cboPlazo, 1) = "1" Then
    'Cargamos los financiamientos Externos largo plazo
    CargaInformacion cboMesIni.ListIndex + 1, cboMesFin.ListIndex + 1, Trim(Me.cboAnioInI), Trim(Me.cboAnioFin), False, "1", 1, pbGetProyectado
    'Cargamos los financiamientos Externos Corto plazo
    CargaInformacion cboMesIni.ListIndex + 1, cboMesFin.ListIndex + 1, Trim(Me.cboAnioInI), Trim(Me.cboAnioFin), True, "1", 2, pbGetProyectado
    'Cargamos los financiamientos Internos largo plazo
    CargaInformacion cboMesIni.ListIndex + 1, cboMesFin.ListIndex + 1, Trim(Me.cboAnioInI), Trim(Me.cboAnioFin), False, "0", 3, pbGetProyectado
    'Cargamos los financiamientos Internos Corto plazo
    CargaInformacion cboMesIni.ListIndex + 1, cboMesFin.ListIndex + 1, Trim(Me.cboAnioInI), Trim(Me.cboAnioFin), True, "0", 4, pbGetProyectado
Else
    For I = 1 To UBound(ltFinanciamientos, 2) - 1
        If I = 1 Then
            lnMesIni = ltFinanciamientos(1, I).lnMes
        Else
            If ltFinanciamientos(1, I).lnMes = 12 Then
                lnMesIni = 1
            Else
                lnMesIni = ltFinanciamientos(1, I).lnMes + 1
            End If
        End If
        lnMesFin = ltFinanciamientos(1, I + 1).lnMes
        If ltFinanciamientos(1, I).lnMes = 12 Then
            lnAnioIni = ltFinanciamientos(1, I).lnAnio + 1
        Else
            lnAnioIni = ltFinanciamientos(1, I).lnAnio
        End If
        lnAnioFin = ltFinanciamientos(1, I + 1).lnAnio
        'Cargamos los financiamientos Externos largo plazo
        CargaInformacion Trim(Str(lnMesIni)), Trim(Str(lnMesFin)), Trim(Str(lnAnioIni)), Trim(Str(lnAnioFin)), False, "1", 1, pbGetProyectado
        'Cargamos los financiamientos Externos Corto plazo
        CargaInformacion Trim(Str(lnMesIni)), Trim(Str(lnMesFin)), Trim(Str(lnAnioIni)), Trim(Str(lnAnioFin)), True, "1", 2, pbGetProyectado
        'Cargamos los financiamientos Internos largo plazo
        CargaInformacion Trim(Str(lnMesIni)), Trim(Str(lnMesFin)), Trim(Str(lnAnioIni)), Trim(Str(lnAnioFin)), False, "0", 3, pbGetProyectado
        'Cargamos los financiamientos Internos Corto plazo
        CargaInformacion Trim(Str(lnMesIni)), Trim(Str(lnMesFin)), Trim(Str(lnAnioIni)), Trim(Str(lnAnioFin)), True, "0", 4, pbGetProyectado
    Next I
End If
Exit Function
ErrorGenerDatosFC:
    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
    GeneraDatos = False
End Function

Private Sub CargaInformacion(lsMesIni As String, lsMesFin As String, lsAnioIni As String, lsAnioFin As String, lbCortoPlazo As Boolean, lsPlaza As String, lnItem As Integer, Optional lbGetProyectado As Boolean = False, Optional lbEjecutadoAnt As Boolean = False)
Dim sql As String
Dim rs As New ADODB.Recordset
Dim lsPlazo As String
Dim I As Integer
'lsPlaza = 0 Interna  , 1 Externa
If lbCortoPlazo Then
    lsPlazo = " AND (CA.nCtaIFCuotas * CB.nCtaIFPlazo)<=365 "
Else
    lsPlazo = " AND (CA.nCtaIFCuotas * CB.nCtaIFPlazo)> 365 "
End If
'PAGOS
sql = "SELECT T.MES, T.ANIO, SUM(CAPITAL) AS  CAPITAL, SUM(INTERES) AS INTERES " _
    & " FROM ( SELECT CAC.CPersCod, CAC.cIFTpo, CAC.cCtaIFCod, MONTH(CAC.dVencimiento) AS MES, YEAR(CAC.dVencimiento) AS  ANIO, " _
    & "               CAPITAL = CASE WHEN SUBSTRING(CAC.CCTAIFCOD,3,1)='1' THEN SUM(NCAPITAL) " _
    & "                              ELSE SUM ( ROUND(NCAPITAL * " & txtTipCambio & ",2) ) End, " _
    & "               INTERES = CASE WHEN SUBSTRING(CAC.CCTAIFCOD,3,1)='1' THEN SUM(CAC.NINTERES) " _
    & "                              ELSE SUM ( ROUND(CAC.NINTERES *" & txtTipCambio & ",2) ) End " _
    & "        FROM   CTAIF CB JOIN CtaIFAdeudados CA ON CA.cPersCod = CB.CPersCod and CA.cIFTpo = CB.cIFTpo and CA.cCtaIFCod = CB.cCtaIFCod " _
    & "               JOIN CtaIFCalendario CAC ON CAC.cPersCod = CA.CPersCod and CAC.cIFTpo = CA.cIFTpo and CAC.cCtaIFCod = CA.cCtaIFCod " _
    & "        WHERE  CB.cCtaIFEstado IN (" & gEstadoCtaIFActiva & "," & gEstadoCtaIFRegistrada & ") and YEAR(CAC.dVencimiento) BETWEEN " & lsAnioIni & " AND " & lsAnioFin & " " _
    & "               AND MONTH(CAC.dVencimiento) BETWEEN " & lsMesIni & " AND " & lsMesFin & "  " _
    & "               AND CAC.cTpoCuota ='" & gCGTipoCuotCalIFCuota & "' AND CA.cPlaza='" & lsPlaza & "' " & lsPlazo _
    & "               AND LEFT(CB.cCtaIFCod,2) in ('" & Format(gTpoCtaIFCtaAdeud, "00") & "'" & IIf(lbGetProyectado, ",'" & Format(gTpoCtaIFAdeudProyecta, "00") & "'", "") & ")" _
    & "        GROUP BY MONTH(CAC.dVencimiento) ,YEAR(CAC.dVencimiento), CAC.CPersCod, CAC.cIFTpo, CAC.cCtaIFCod ) AS T " _
    & " GROUP BY T.MES,T.ANIO " _
    & " ORDER BY T.ANIO, T.MES "
Set rs = oCon.CargaRecordSet(sql)
If Not RSVacio(rs) Then
    Do While Not rs.EOF
        If lbEjecutadoAnt = False Then
            For I = 1 To UBound(ltFinanciamientos, 2)
                If Right(Me.cboPlazo, 1) = "1" Then
                    If ltFinanciamientos(lnItem, I).lnAnio = rs!Anio And ltFinanciamientos(lnItem, I).lnMes = rs!MES Then
                        ltFinanciamientos(lnItem, I).lnAmortizacion = Round(rs!Capital)
                        ltFinanciamientos(lnItem, I).lnInteres = Round(rs!Interes)
                        Exit For
                    End If
                Else
                    'Primer mes del año
                    If rs!Anio > ltFinanciamientos(lnItem, I).lnAnio And (rs!MES >= 1 And rs!MES <= 3) And ltFinanciamientos(lnItem, I).lnMes = 12 Then
                        ltFinanciamientos(lnItem, I).lnAmortizacion = ltFinanciamientos(lnItem, I).lnAmortizacion + Round(rs!Capital)
                        ltFinanciamientos(lnItem, I).lnInteres = ltFinanciamientos(lnItem, I).lnInteres + Round(rs!Interes)
                        Exit For
                    Else
                        If ltFinanciamientos(lnItem, I).lnAnio = rs!Anio And (rs!MES >= ltFinanciamientos(lnItem, I).lnMes And rs!MES <= ltFinanciamientos(lnItem, I + 1).lnMes) Then
                            ltFinanciamientos(lnItem, I).lnAmortizacion = ltFinanciamientos(lnItem, I).lnAmortizacion + Round(rs!Capital)
                            ltFinanciamientos(lnItem, I).lnInteres = ltFinanciamientos(lnItem, I).lnInteres + Round(rs!Interes)
                            Exit For
                        End If
                    End If
                End If
            Next
        Else
            ltEjecutado(lnItem, 1).lnAmortizacion = ltEjecutado(lnItem, 1).lnAmortizacion + Round(rs!Capital)
            ltEjecutado(lnItem, 1).lnInteres = ltEjecutado(lnItem, 1).lnInteres + Round(rs!Interes)
        End If
        rs.MoveNext
    Loop
End If
RSClose rs

'DESEMBOLSOS
sql = "SELECT T.MES, T.ANIO, SUM(DESEMBOLSO) AS  CAPITAL " _
    & " FROM ( SELECT CA.CPersCod, CA.cIFTpo, CA.cCtaIFCod, MONTH(CB.dCtaIFAper) AS MES, YEAR(CB.dCtaIFAper) AS  ANIO, " _
    & "               DESEMBOLSO = CASE WHEN SUBSTRING(CA.CCTAIFCOD,3,1)='1' THEN CA.nMontoPrestado " _
    & "                                 ELSE ROUND(CA.nMontoPrestado * " & txtTipCambio & ",2) END " _
    & "        FROM   CTAIF CB JOIN CtaIFAdeudados CA ON CA.cPersCod = CB.CPersCod and CA.cIFTpo = CB.cIFTpo and CA.cCtaIFCod = CB.cCtaIFCod " _
    & "        WHERE  YEAR(CB.dCtaIFAper) BETWEEN " & lsAnioIni & " AND " & lsAnioFin & " " _
    & "               AND MONTH(CB.dCtaIFAper) BETWEEN " & lsMesIni & " AND " & lsMesFin & "  " _
    & "               AND CA.cPlaza ='" & lsPlaza & "' " & lsPlazo _
    & "        ) AS T " _
    & " GROUP BY T.MES,T.ANIO " _
    & " ORDER BY T.ANIO, T.MES "

Set rs = oCon.CargaRecordSet(sql)
If Not RSVacio(rs) Then
    Do While Not rs.EOF
        If lbEjecutadoAnt = False Then
            For I = 1 To UBound(ltFinanciamientos, 2)
                If Right(Me.cboPlazo, 1) = "1" Then
                    If ltFinanciamientos(lnItem, I).lnAnio = rs!Anio And ltFinanciamientos(lnItem, I).lnMes = rs!MES Then
                        ltFinanciamientos(lnItem, I).lnDesembolso = Round(rs!Capital)
                        Exit For
                    End If
                Else
                    'Primer mes del año
                    If rs!Anio > ltFinanciamientos(lnItem, I).lnAnio And (rs!MES >= 1 And rs!MES <= 3) And ltFinanciamientos(lnItem, I).lnMes = 12 Then
                        ltFinanciamientos(lnItem, I).lnDesembolso = Round(rs!Capital)
                        Exit For
                    Else
                        If ltFinanciamientos(lnItem, I).lnAnio = rs!Anio And (rs!MES >= ltFinanciamientos(lnItem, I).lnMes And rs!MES <= ltFinanciamientos(lnItem, I + 1).lnMes) Then
                            ltFinanciamientos(lnItem, I).lnDesembolso = ltFinanciamientos(lnItem, I).lnDesembolso + Round(rs!Capital)
                            Exit For
                        End If
                    End If
                End If
            Next
        Else
            ltEjecutado(lnItem, 1).lnDesembolso = ltEjecutado(lnItem, 1).lnDesembolso + Round(rs!Capital)
        End If
        rs.MoveNext
    Loop
End If
RSClose rs
End Sub
Private Function ValorRomano(I As Integer) As String
Select Case I
        Case 1: ValorRomano = "I"
        Case 2: ValorRomano = "II"
        Case 3: ValorRomano = "III"
        Case 4: ValorRomano = "IV"
        Case 5: ValorRomano = "V"
        Case 6: ValorRomano = "VI"
        Case 7: ValorRomano = "VII"
        Case 8: ValorRomano = "VIII"
        Case 9: ValorRomano = "IX"
        Case 10: ValorRomano = "X"
        Case 11: ValorRomano = "XI"
        Case 12: ValorRomano = "XII"
End Select
End Function
Private Sub GeneraReporteExcel()
Dim fs As New Scripting.FileSystemObject
Dim lbExisteHoja As Boolean
Dim lnFila As Integer, lnCol As Integer
Dim I As Integer
Dim lsTotal As String
Dim Y1 As Integer, Y2 As Integer, Y11 As Integer
Dim j As Integer, N As Integer

Dim lnFilaNeto As Integer

Dim lnFilaExtNeto As Integer
Dim lnFilaIntNeto As Integer

Dim lnFilaComun As Integer

Dim lnFilaServDeuda As Integer

Dim lsFilaNeto() As String
Dim lsFilaExtNeto() As String
Dim lsFilaIntNeto() As String
Dim lsFilaServDeuda() As String
Dim lsFilaComun() As String
Dim lsTotales() As String
    
ReDim lsFilaNeto(0)
ReDim lsFilaExtNeto(0)
ReDim lsFilaIntNeto(0)
ReDim lsFilaServDeuda(0)
ReDim lsFilaComun(0)
    
    
    lsArchivo = App.path & "\SPOOLER\PFlujoCaja_" & Me.cboAnioInI & "Flujo.XLS"
    lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)

    Me.barraEstado.Panels(1).Text = "Generando Reporte ..."
    Set xlHoja1 = xlLibro.Worksheets.Add
    
    xlHoja1.PageSetup.Zoom = 65
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlAplicacion.Range("A1:R100").Font.Size = 9
    
    xlHoja1.Range("A1").ColumnWidth = 7
    xlHoja1.Range("B1").ColumnWidth = 36
    xlHoja1.Range("C1:Z1").ColumnWidth = IIf(Right(cboPlazo, 1) = "1", 10, 15)
    
    lnFila = 3
    xlHoja1.Cells(lnFila, 2) = "ENTIDAD": xlHoja1.Cells(lnFila, 6) = gsNomCmac: xlHoja1.Cells(lnFila, 10) = "Area de Caja General"
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 6)).Font.Bold = True
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).Font.Size = 12
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "PRESUPUESTO " & Trim(Me.cboAnioInI): xlHoja1.Cells(lnFila, 6) = "FLUJO DE CAJA"
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 6)).Font.Bold = True
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).Font.Size = 12
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "EN NUEVOS SOLES ":     xlHoja1.Cells(lnFila, 10) = "Reporte  al :" & Format(gdFecSis, "dd mmmm yyyy")
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
    
    lnFila = lnFila + 2
    Y1 = lnFila
    xlHoja1.Cells(lnFila, 2) = "RUBROS":  xlHoja1.Cells(lnFila, 3) = "EJECUCION":
    lnCol = 3
    For I = 1 To UBound(ltFinanciamientos, 2) - IIf(Right(cboPlazo, 1) = "1", 0, 1)
        lnCol = lnCol + 1
        If Right(cboPlazo, 1) = "1" Then
            xlHoja1.Cells(lnFila, lnCol) = Format("01/" & Format(ltFinanciamientos(1, I).lnMes, "00") & "/" & ltFinanciamientos(1, I).lnAnio, "mmmm")
        Else
            xlHoja1.Cells(lnFila, lnCol) = ValorRomano(I) & " TRIMESTRE"
        End If
    Next
    lnCol = lnCol + 1
    xlHoja1.Cells(lnFila, lnCol) = "TOTAL"
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, lnCol)).HorizontalAlignment = xlCenter
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = " ":  xlHoja1.Cells(lnFila, 3) = "AÑO " & Val(Me.cboAnioInI) - 1:
    lnCol = 3
    For I = 1 To UBound(ltFinanciamientos, 2) - IIf(Right(cboPlazo, 1) = "1", 0, 1)
        lnCol = lnCol + 1
        xlHoja1.Cells(lnFila, lnCol) = ltFinanciamientos(1, I).lnAnio
    Next
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, lnCol)).HorizontalAlignment = xlCenter
    Y2 = lnFila
    CuadroExcel 2, Y1, lnCol + 1, Y2
    
    'lnFila = lnFila + 1
    lnCol = lnCol + 1
    Y11 = lnFila + 1
    For j = 1 To UBound(ltFinanciamientos, 1)
        Select Case j
            Case 1
                lnFila = lnFila + 1
                Y1 = lnFila
                lnFilaNeto = lnFila
                xlHoja1.Cells(lnFila, 2) = "FINANCIAMIENTO NETO":
                xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, lnCol + 1)).Font.Bold = True
                
                Y2 = lnFila
                CuadroExcel 2, Y1, lnCol, Y2
                
                lnFila = lnFila + 1
                Y1 = lnFila
                lnFilaExtNeto = lnFila
                xlHoja1.Cells(lnFila, 2) = " Financiamiento Externo Neto":
                xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, lnCol + 1)).Font.Bold = True
                Y2 = lnFila
                CuadroExcel 2, Y1, lnCol, Y2
                
                lnFila = lnFila + 1
                lnFilaComun = lnFila
                xlHoja1.Cells(lnFila, 2) = "   Financiamiento Largo Plazo":
                xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, lnCol + 1)).Font.Bold = True
                Y2 = lnFila
                CuadroExcel 2, Y2, lnCol, Y2
            Case 2
                lnFila = lnFila + 1
                Y1 = lnFila
                lnFilaComun = lnFila
                xlHoja1.Cells(lnFila, 2) = "   Financiamiento Corto Plazo":
                xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
                Y2 = lnFila
                CuadroExcel 2, Y1, lnCol, Y2
            Case 3
                lnFila = lnFila + 1
                Y1 = lnFila
                lnFilaIntNeto = lnFila
                xlHoja1.Cells(lnFila, 2) = " Financiamiento Interno Neto":
                xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
                Y2 = lnFila
                CuadroExcel 2, Y1, lnCol, Y2
                
                lnFila = lnFila + 1
                Y1 = lnFila
                lnFilaComun = lnFila
                xlHoja1.Cells(lnFila, 2) = "   Financiamiento Largo Plazo":
                xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
                Y2 = lnFila
                CuadroExcel 2, Y1, lnCol, Y2
            Case 4
                lnFila = lnFila + 1
                Y1 = lnFila
                lnFilaComun = lnFila
                xlHoja1.Cells(lnFila, 2) = "   Financiamiento Corto Plazo":
                xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
                Y2 = lnFila
                CuadroExcel 2, Y1, lnCol, Y2
        End Select
        ReDim Preserve lsTotales(1)
        lnFila = lnFila + 1
        xlHoja1.Cells(lnFila, 2) = "        Desembolsos ":
        
        lnCol = 3
        ReDim Preserve lsFilaComun(lnCol)
        xlHoja1.Cells(lnFila, lnCol) = ltEjecutado(j, 1).lnDesembolso
        lsFilaComun(lnCol) = xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False)
        For I = 1 To UBound(ltFinanciamientos, 2) - IIf(Right(cboPlazo, 1) = "1", 0, 1)
            lnCol = lnCol + 1
            ReDim Preserve lsFilaComun(lnCol)
            xlHoja1.Cells(lnFila, lnCol) = ltFinanciamientos(j, I).lnDesembolso
            xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, lnCol)).NumberFormat = "#,##0"
            lsFilaComun(lnCol) = xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False)
            If I = 1 Then
                lsTotales(1) = xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol - 1), xlHoja1.Cells(lnFila, lnCol - 1)).Address(False, False)
            End If
        Next
        lsTotales(1) = lsTotales(1) & ":" & xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False)
        lnCol = lnCol + 1
        xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Formula = "=SUM(" & lsTotales(1) & ")"
        
        lnFila = lnFila + 1
        lnFilaServDeuda = lnFila
        xlHoja1.Cells(lnFila, 2) = "          Servicio de Deuda ":
        xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
        
        lnFila = lnFila + 1
        xlHoja1.Cells(lnFila, 2) = "          Amortizaciones ":
        lnCol = 3
        ReDim Preserve lsFilaServDeuda(lnCol)
        xlHoja1.Cells(lnFila, lnCol) = ltEjecutado(j, 1).lnAmortizacion
        lsFilaServDeuda(lnCol) = xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False)
        For I = 1 To UBound(ltFinanciamientos, 2) - IIf(Right(cboPlazo, 1) = "1", 0, 1)
            lnCol = lnCol + 1
            ReDim Preserve lsFilaServDeuda(lnCol)
            xlHoja1.Cells(lnFila, lnCol) = ltFinanciamientos(j, I).lnAmortizacion
            xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, lnCol)).NumberFormat = "#,##0"
            lsFilaServDeuda(lnCol) = xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False)
            If I = 1 Then
                lsTotales(1) = xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol - 1), xlHoja1.Cells(lnFila, lnCol - 1)).Address(False, False)
            End If
        Next
        lsTotales(1) = lsTotales(1) & ":" & xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False)
        lnCol = lnCol + 1
        xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Formula = "=SUM(" & lsTotales(1) & ")"
        
        lnFila = lnFila + 1
        xlHoja1.Cells(lnFila, 2) = "          Intereses y comisiones de la deuda ":
        lnCol = 3
        xlHoja1.Cells(lnFila, lnCol) = ltEjecutado(j, 1).lnInteres
        lsFilaServDeuda(lnCol) = lsFilaServDeuda(lnCol) & ":" & xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False)
        For I = 1 To UBound(ltFinanciamientos, 2) - IIf(Right(cboPlazo, 1) = "1", 0, 1)
            lnCol = lnCol + 1
            xlHoja1.Cells(lnFila, lnCol) = ltFinanciamientos(j, I).lnInteres
            xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, lnCol)).NumberFormat = "#,##0"
            lsFilaServDeuda(lnCol) = lsFilaServDeuda(lnCol) & ":" & xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False)
            If I = 1 Then
                lsTotales(1) = xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol - 1), xlHoja1.Cells(lnFila, lnCol - 1)).Address(False, False)
            End If
        Next
        lsTotales(1) = lsTotales(1) & ":" & xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False)
        lnCol = lnCol + 1
        xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Formula = "=SUM(" & lsTotales(1) & ")"
        
        '***************** Impresion de la fila de montos de servicio a la deuda
        lnCol = 3
        lsFilaComun(lnCol) = lsFilaComun(lnCol) & "-" & xlAplicacion.Range(xlHoja1.Cells(lnFilaServDeuda, lnCol), xlHoja1.Cells(lnFilaServDeuda, lnCol)).Address(False, False)
        xlAplicacion.Range(xlHoja1.Cells(lnFilaServDeuda, lnCol), xlHoja1.Cells(lnFilaServDeuda, lnCol)).Formula = "=SUM(" & lsFilaServDeuda(lnCol) & ")"
        For I = 1 To UBound(ltFinanciamientos, 2) - IIf(Right(cboPlazo, 1) = "1", 0, 1)
            lnCol = lnCol + 1
            xlAplicacion.Range(xlHoja1.Cells(lnFilaServDeuda, lnCol), xlHoja1.Cells(lnFilaServDeuda, lnCol)).Formula = "=SUM(" & lsFilaServDeuda(lnCol) & ")"
            lsFilaComun(lnCol) = lsFilaComun(lnCol) & "-" & xlAplicacion.Range(xlHoja1.Cells(lnFilaServDeuda, lnCol), xlHoja1.Cells(lnFilaServDeuda, lnCol)).Address(False, False)
            If I = 1 Then
                lsTotales(1) = xlAplicacion.Range(xlHoja1.Cells(lnFilaServDeuda, lnCol - 1), xlHoja1.Cells(lnFilaServDeuda, lnCol - 1)).Address(False, False)
            End If
        Next
        lsTotales(1) = lsTotales(1) & ":" & xlAplicacion.Range(xlHoja1.Cells(lnFilaServDeuda, lnCol), xlHoja1.Cells(lnFilaServDeuda, lnCol)).Address(False, False)
        lnCol = lnCol + 1
        xlAplicacion.Range(xlHoja1.Cells(lnFilaServDeuda, lnCol), xlHoja1.Cells(lnFilaServDeuda, lnCol)).Formula = "=SUM(" & lsTotales(1) & ")"
        
        '-********************* Impresion de la fila de TOTALES DE FINANCIAMIENTOS TANTO A LARGO COMO ACORTO PLAZO
        lnCol = 3
        
        Select Case j
            Case 1
                ReDim Preserve lsFilaExtNeto(lnCol)
                lsFilaExtNeto(lnCol) = xlAplicacion.Range(xlHoja1.Cells(lnFilaComun, lnCol), xlHoja1.Cells(lnFilaComun, lnCol)).Address(False, False)
            Case 2
                lsFilaExtNeto(lnCol) = lsFilaExtNeto(lnCol) & "+" & xlAplicacion.Range(xlHoja1.Cells(lnFilaComun, lnCol), xlHoja1.Cells(lnFilaComun, lnCol)).Address(False, False)
            Case 3
                ReDim Preserve lsFilaIntNeto(lnCol)
                lsFilaIntNeto(lnCol) = xlAplicacion.Range(xlHoja1.Cells(lnFilaComun, lnCol), xlHoja1.Cells(lnFilaComun, lnCol)).Address(False, False)
            Case 4
                lsFilaIntNeto(lnCol) = lsFilaIntNeto(lnCol) & "+" & xlAplicacion.Range(xlHoja1.Cells(lnFilaComun, lnCol), xlHoja1.Cells(lnFilaComun, lnCol)).Address(False, False)
        End Select
        xlAplicacion.Range(xlHoja1.Cells(lnFilaComun, lnCol), xlHoja1.Cells(lnFilaComun, lnCol)).Formula = "=(" & lsFilaComun(lnCol) & ")"
        For I = 1 To UBound(ltFinanciamientos, 2) - IIf(Right(cboPlazo, 1) = "1", 0, 1)
            lnCol = lnCol + 1
            xlAplicacion.Range(xlHoja1.Cells(lnFilaComun, lnCol), xlHoja1.Cells(lnFilaComun, lnCol)).Formula = "=(" & lsFilaComun(lnCol) & ")"
            Select Case j
                Case 1
                    ReDim Preserve lsFilaExtNeto(lnCol)
                    lsFilaExtNeto(lnCol) = xlAplicacion.Range(xlHoja1.Cells(lnFilaComun, lnCol), xlHoja1.Cells(lnFilaComun, lnCol)).Address(False, False)
                Case 2
                    lsFilaExtNeto(lnCol) = lsFilaExtNeto(lnCol) & "+" & xlAplicacion.Range(xlHoja1.Cells(lnFilaComun, lnCol), xlHoja1.Cells(lnFilaComun, lnCol)).Address(False, False)
                Case 3
                    ReDim Preserve lsFilaIntNeto(lnCol)
                    lsFilaIntNeto(lnCol) = xlAplicacion.Range(xlHoja1.Cells(lnFilaComun, lnCol), xlHoja1.Cells(lnFilaComun, lnCol)).Address(False, False)
                Case 4
                    lsFilaIntNeto(lnCol) = lsFilaIntNeto(lnCol) & "+" & xlAplicacion.Range(xlHoja1.Cells(lnFilaComun, lnCol), xlHoja1.Cells(lnFilaComun, lnCol)).Address(False, False)
            End Select
            If I = 1 Then
                lsTotales(1) = xlAplicacion.Range(xlHoja1.Cells(lnFilaComun, lnCol - 1), xlHoja1.Cells(lnFilaComun, lnCol - 1)).Address(False, False)
            End If
        Next
        lsTotales(1) = lsTotales(1) & ":" & xlAplicacion.Range(xlHoja1.Cells(lnFilaComun, lnCol), xlHoja1.Cells(lnFilaComun, lnCol)).Address(False, False)
        lnCol = lnCol + 1
        xlAplicacion.Range(xlHoja1.Cells(lnFilaComun, lnCol), xlHoja1.Cells(lnFilaComun, lnCol)).Formula = "=SUM(" & lsTotales(1) & ")"
    Next
    Y2 = lnFila
    CuadroExcel 2, Y11, lnCol, Y2
    
    lnCol = 3
    ReDim Preserve lsFilaNeto(lnCol)
    lsFilaNeto(lnCol) = xlAplicacion.Range(xlHoja1.Cells(lnFilaExtNeto, lnCol), xlHoja1.Cells(lnFilaExtNeto, lnCol)).Address(False, False) & "+" & xlAplicacion.Range(xlHoja1.Cells(lnFilaIntNeto, lnCol), xlHoja1.Cells(lnFilaIntNeto, lnCol)).Address(False, False)
    xlAplicacion.Range(xlHoja1.Cells(lnFilaExtNeto, lnCol), xlHoja1.Cells(lnFilaExtNeto, lnCol)).Formula = "=(" & lsFilaExtNeto(lnCol) & ")"
    xlAplicacion.Range(xlHoja1.Cells(lnFilaIntNeto, lnCol), xlHoja1.Cells(lnFilaIntNeto, lnCol)).Formula = "=(" & lsFilaIntNeto(lnCol) & ")"
    For I = 1 To UBound(ltFinanciamientos, 2) - IIf(Right(cboPlazo, 1) = "1", 0, 1)
        lnCol = lnCol + 1
        xlAplicacion.Range(xlHoja1.Cells(lnFilaExtNeto, lnCol), xlHoja1.Cells(lnFilaExtNeto, lnCol)).Formula = "=(" & lsFilaExtNeto(lnCol) & ")"
        xlAplicacion.Range(xlHoja1.Cells(lnFilaIntNeto, lnCol), xlHoja1.Cells(lnFilaIntNeto, lnCol)).Formula = "=(" & lsFilaIntNeto(lnCol) & ")"
        ReDim Preserve lsFilaNeto(lnCol)
        lsFilaNeto(lnCol) = xlAplicacion.Range(xlHoja1.Cells(lnFilaExtNeto, lnCol), xlHoja1.Cells(lnFilaExtNeto, lnCol)).Address(False, False) & "+" & xlAplicacion.Range(xlHoja1.Cells(lnFilaIntNeto, lnCol), xlHoja1.Cells(lnFilaIntNeto, lnCol)).Address(False, False)
        If I = 1 Then
            lsTotales(1) = xlAplicacion.Range(xlHoja1.Cells(lnFilaExtNeto, lnCol - 1), xlHoja1.Cells(lnFilaExtNeto, lnCol - 1)).Address(False, False)
            lsTotales(0) = xlAplicacion.Range(xlHoja1.Cells(lnFilaIntNeto, lnCol - 1), xlHoja1.Cells(lnFilaIntNeto, lnCol - 1)).Address(False, False)
        End If
    Next
    lsTotales(1) = lsTotales(1) & ":" & xlAplicacion.Range(xlHoja1.Cells(lnFilaExtNeto, lnCol - 1), xlHoja1.Cells(lnFilaExtNeto, lnCol - 1)).Address(False, False)
    lsTotales(0) = lsTotales(0) & ":" & xlAplicacion.Range(xlHoja1.Cells(lnFilaIntNeto, lnCol), xlHoja1.Cells(lnFilaIntNeto, lnCol)).Address(False, False)
    lnCol = lnCol + 1
    xlAplicacion.Range(xlHoja1.Cells(lnFilaExtNeto, lnCol), xlHoja1.Cells(lnFilaExtNeto, lnCol)).Formula = "=SUM(" & lsTotales(1) & ")"
    xlAplicacion.Range(xlHoja1.Cells(lnFilaIntNeto, lnCol), xlHoja1.Cells(lnFilaIntNeto, lnCol)).Formula = "=SUM(" & lsTotales(0) & ")"
    
    lnCol = 3
    xlAplicacion.Range(xlHoja1.Cells(lnFilaNeto, lnCol), xlHoja1.Cells(lnFilaNeto, lnCol)).Formula = "=(" & lsFilaNeto(lnCol) & ")"
    For I = 1 To UBound(ltFinanciamientos, 2) - IIf(Right(cboPlazo, 1) = "1", 0, 1)
        lnCol = lnCol + 1
        xlAplicacion.Range(xlHoja1.Cells(lnFilaNeto, lnCol), xlHoja1.Cells(lnFilaNeto, lnCol)).Formula = "=(" & lsFilaNeto(lnCol) & ")"
        If I = 1 Then
            lsTotales(1) = xlAplicacion.Range(xlHoja1.Cells(lnFilaNeto, lnCol - 1), xlHoja1.Cells(lnFilaNeto, lnCol - 1)).Address(False, False)
        End If
    Next
    lsTotales(1) = lsTotales(1) & ":" & xlAplicacion.Range(xlHoja1.Cells(lnFilaNeto, lnCol), xlHoja1.Cells(lnFilaNeto, lnCol)).Address(False, False)
    lnCol = lnCol + 1
    xlHoja1.Range(xlHoja1.Cells(lnFilaNeto, lnCol), xlHoja1.Cells(lnFilaNeto, lnCol)).Formula = "=SUM(" & lsTotales(1) & ")"
    
    ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
    
    Me.barraEstado.Panels(1).Text = "Reporte Generado con Exito"
    lbExcel = False

End Sub
Private Sub CuadroExcel(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, Optional lbLineasVert As Boolean = False)
Dim I, j As Integer

For I = X1 To X2
    xlHoja1.Range(xlHoja1.Cells(Y1, I), xlHoja1.Cells(Y1, I)).Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(Y2, I), xlHoja1.Cells(Y2, I)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Next I
If lbLineasVert = False Then
    For I = X1 To X2
        For j = Y1 To Y2
            xlHoja1.Range(xlHoja1.Cells(j, I), xlHoja1.Cells(j, I)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Next j
    Next I
End If
If lbLineasVert Then
    For j = Y1 To Y2
        xlHoja1.Range(xlHoja1.Cells(j, X1), xlHoja1.Cells(j, X1)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Next j
End If

For j = Y1 To Y2
    xlHoja1.Range(xlHoja1.Cells(j, X2), xlHoja1.Cells(j, X2)).Borders(xlEdgeRight).LineStyle = xlContinuous
Next j
End Sub

'***** Ejecucion del anio anterior tanto
Private Sub GeneraEjecutado(lsMesIni As String, lsMesFin As String, lsAnio As String)

ReDim ltEjecutado(4, 1)
'Cargamos los financiamientos Externos Largo Plazo
CargaInformacion lsMesIni, lsMesFin, lsAnio, lsAnio, False, "1", 1, , True
'Cargamos los financiamientos Externos Corto plazo
CargaInformacion lsMesIni, lsMesFin, lsAnio, lsAnio, True, "1", 2, , True
'Cargamos los financiamientos Internos largo plazo
CargaInformacion lsMesIni, lsMesFin, lsAnio, lsAnio, False, "0", 3, , True
'Cargamos los financiamientos Internos Corto plazo
CargaInformacion lsMesIni, lsMesFin, lsAnio, lsAnio, True, "0", 4, , True

End Sub

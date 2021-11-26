VERSION 5.00
Begin VB.Form frmCredPagoServicios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago de Servicios"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7155
   Icon            =   "frmCredPagoServicios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDetalle 
      Appearance      =   0  'Flat
      Caption         =   "Detalle"
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
      Height          =   3345
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7095
      Begin VB.Frame frameCli 
         Caption         =   "Busca Cliente"
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   3855
         Begin VB.CommandButton CmdCliente 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   22
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblcodservicio 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   2460
         End
      End
      Begin VB.ComboBox cboInsti 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   5730
      End
      Begin SICMACT.EditMoney txtMonto 
         Height          =   345
         Left            =   4725
         TabIndex        =   1
         Top             =   2145
         Width           =   1800
         _extentx        =   3175
         _extenty        =   609
         font            =   "frmCredPagoServicios.frx":030A
         text            =   "0"
         enabled         =   -1  'True
      End
      Begin VB.Label lblmensaje 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   795
         Left            =   4080
         TabIndex        =   19
         Top             =   600
         Width           =   2880
      End
      Begin VB.Label lblcuenta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   1320
         Width           =   2820
      End
      Begin VB.Label Label5 
         Caption         =   "Cuenta :"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   1035
      End
      Begin VB.Label lblcomision 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   300
         Left            =   1200
         TabIndex        =   16
         Top             =   2160
         Width           =   1755
      End
      Begin VB.Label Label2 
         Caption         =   "Comisión :"
         ForeColor       =   &H80000007&
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblApe 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1200
         TabIndex        =   15
         Top             =   1680
         Width           =   5700
      End
      Begin VB.Label Label1 
         Caption         =   "Institucion :"
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   885
      End
      Begin VB.Label lblMonto 
         Caption         =   "Monto :"
         ForeColor       =   &H80000007&
         Height          =   180
         Left            =   4110
         TabIndex        =   10
         Top             =   2235
         Width           =   735
      End
      Begin VB.Label lblPers 
         Caption         =   "Persona :"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   915
      End
      Begin VB.Label lblSimbolo 
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6585
         TabIndex        =   8
         Top             =   2160
         Width           =   480
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   300
         Left            =   4725
         TabIndex        =   18
         Top             =   2970
         Width           =   1755
      End
      Begin VB.Label lblITF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   300
         Left            =   4725
         TabIndex        =   17
         Top             =   2595
         Width           =   1755
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         Height          =   195
         Left            =   4245
         TabIndex        =   7
         Top             =   3015
         Width           =   450
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "ITF :"
         Height          =   195
         Left            =   4275
         TabIndex        =   6
         Top             =   2625
         Width           =   330
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4875
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
End
Attribute VB_Name = "frmCredPagoServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nNroDeu As Long
Dim nValorDeuda As Currency
Dim sPersCodCMAC As String, sTipoCuenta As String
Dim sMensaje As String, sBoleta As String, sBoletaITF As String, smensajePer As String
Dim bgraba As Boolean
Dim LslistaCreditos As String
'Dim nPTipoBus As Integer
Dim nRedondeoITF As Double 'BRGO 20110914
Private MatCreditos() As String

Private Sub cboInsti_Click()
   Dim oPers As COMDPersona.DCOMPersonas
   Dim rs As ADODB.Recordset

If Not bgraba Then
    If Me.cboInsti.ListIndex = 0 Then
        MsgBox "Seleccione una institución para iniciar el proceso de pago", vbCritical, "Aviso"
        Exit Sub
    End If

    If Len(Me.cboInsti.Text) = 0 Then
        MsgBox "Seleccione una institución para iniciar el proceso de pago", vbCritical, "Aviso"
        Exit Sub
    End If

    Set oPers = New COMDPersona.DCOMPersonas
    Set rs = oPers.CargaDatosPagoServiciosxInstitucion(Trim(Right(Me.cboInsti.Text, 13)))
    Set oPers = Nothing
    'MADM 20110602
    If Not (rs.BOF And rs.EOF) Then
        lblMensaje.Caption = Trim(rs!cObsConv)
        'nPTipoBus = Trim(rs!nTipoBus)
        frameCli.Enabled = True

'        If nPTipoBus <> -1 Then
'            If nPTipoBus = 2 Then
'                frameId.Enabled = True
'                frameCli.Enabled = True
'            ElseIf nPTipoBus = 1 Then
'                frameId.Enabled = True
'                frameCli.Enabled = False
'                Me.txtcodservicio.Text = ""
'                lblcodservicio.Caption = ""
'            Else
'                frameId.Enabled = False
'                frameCli.Enabled = True
'                Me.txtcodservicio.Text = ""
'                lblcodservicio.Caption = ""
'            End If
'        End If

        If gdFecSis > CDate(rs!fVigencia) Then
            MsgBox "Ud. No podrá continuar el Convenio ya no tiene Vigencia", vbCritical, "Aviso"
            Exit Sub
        End If
     Else
        MsgBox "Ud. No podrá continuar debido a que el código de Servicio no se ha encontrado, Verifique", vbCritical, "Aviso"
        Exit Sub
    End If
    Set rs = Nothing
End If
End Sub

Private Sub cmdCancelar_Click()
    LimpiaControles
End Sub

Private Function DameCreditos() As String
Dim sCreditos As String
Dim iCre As Integer
sCreditos = ""

  If UBound(MatCreditos) > 0 Then
     For iCre = 0 To UBound(MatCreditos)
            If UBound(MatCreditos) = 0 Then
                sCreditos = "'" & MatCreditos(iCre) & "'"
            Else
                If MatCreditos(iCre) <> "" Then
                    sCreditos = sCreditos & "'" & MatCreditos(iCre) & "'" & ","
                End If
            End If
        Next iCre
   End If
   If Right(sCreditos, 1) = "," Then
        sCreditos = Mid(sCreditos, 1, Len(sCreditos) - 1)
   End If
   DameCreditos = sCreditos
End Function

Sub limpia_array()
ReDim MatCreditos(0)
MatCreditos(0) = ""
End Sub

Private Sub CmdCliente_Click()
Dim oPers As COMDPersona.DCOMPersonas
Dim rs As ADODB.Recordset
limpia_array
Set oPers = New COMDPersona.DCOMPersonas

    If Trim(Right(Me.cboInsti.Text, 13)) = "" Then
        MsgBox "Debe especificar una Institución para Iniciar búsqueda", vbInformation
        Exit Sub
    End If

MatCreditos = frmBuscaPersonaNegativaServiciosLista.Inicio(Trim(Right(Me.cboInsti.Text, 13)))
If MatCreditos(0) <> "" Then
LslistaCreditos = DameCreditos
    Set rs = oPers.CargaDatosPagoServiciosxDocLista(LslistaCreditos, gdFecSis)
        If Not (rs.BOF And rs.EOF) Then
            smensajePer = ""
            nValorDeuda = 0
            'nNroDeu = frmBuscaPersonaNegativaServicios.sNumdeuSist
            'lblcodservicio = frmBuscaPersonaNegativaServicios.sNumdeuInst
            lblcuenta = Trim(rs!cCuenta)
            lblApe = Trim(rs!cNombre)
            LblNumDoc = Trim(rs!cNumDoc)
            'lblcon.Caption = IIf(IsNull(rs!cConcepto), "", rs!cConcepto)
            'lblper.Caption = IIf(IsNull(rs!cPeriodo), "", rs!cPeriodo)
            'lblmon.Caption = rs!Moneda
            If rs!Moneda = "Soles" Then
                sMoneda = "MONEDA NACIONAL"
                txtMonto.BackColor = &HC0FFFF
'                lblmon.BackColor = &HC0FFFF
                lblComision.BackColor = &HC0FFFF
                lblSimbolo.Caption = "S/."
            Else
                sMoneda = "MONEDA EXTRANJERA"
                txtMonto.BackColor = &HC0FFC0
'                lblmon.BackColor = &HC0FFC0
                lblComision.BackColor = &HC0FFC0
                lblSimbolo.Caption = "$"
            End If
'            txtGlosa.SetFocus
            lblComision = Format(rs!nImporteComision, "#,##0.00")
            
            txtMonto.Text = IIf(IsNull(rs!nImporteCuota), "", rs!nImporteCuota)
            nValorDeuda = CCur(txtMonto.Text)
            smensajePer = IIf(IsNull(rs!cImprime), "", rs!cImprime)
            txtMonto.Enabled = True
            txtMonto.SetFocus
'            If CmdGrabar.Enabled = False Then
'                CmdGrabar.Enabled = True
'                Me.CmdGrabar.SetFocus
'            End If
  End If
End If

'MADM 20111123 - PARA UNA OPERACION
''Call frmBuscaPersonaNegativaServicios.Inicio(Trim(Right(Me.cboInsti.Text, 13)))
'''**********MADM 20110602
''If Not frmBuscaPersonaNegativaServicios.sNumdeuSist = "" Then
''    Set rs = oPers.CargaDatosPagoServiciosxDoc(frmBuscaPersonaNegativaServicios.sNumdeuSist, gdFecSis)
''    If Not (rs.BOF And rs.EOF) Then
''        smensajePer = ""
''        nValorDeuda = 0
''        nNroDeu = frmBuscaPersonaNegativaServicios.sNumdeuSist
''        lblcodservicio = frmBuscaPersonaNegativaServicios.sNumdeuInst
''        lblcuenta = Trim(rs!cCuenta)
''        lblApe = Trim(rs!cNombre)
''        lblNumdoc = Trim(rs!cNumDoc)
''        lblcon.Caption = IIf(IsNull(rs!cConcepto), "", rs!cConcepto)
''        lblper.Caption = IIf(IsNull(rs!cPeriodo), "", rs!cPeriodo)
''        If rs!Moneda = "Soles" Then
''            sMoneda = "MONEDA NACIONAL"
''            txtMonto.BackColor = &HC0FFFF
''            lblmon.BackColor = &HC0FFFF
''            lblcomision.BackColor = &HC0FFFF
''            lblSimbolo.Caption = "S/."
''        Else
''            sMoneda = "MONEDA EXTRANJERA"
''            txtMonto.BackColor = &HC0FFC0
''            lblmon.BackColor = &HC0FFC0
''            lblcomision.BackColor = &HC0FFC0
''            lblSimbolo.Caption = "$"
''        End If
''        txtGlosa.SetFocus
''        lblcomision = Format(rs!nImporteComision, "#,##0.00")
''        txtMonto.Text = IIf(IsNull(rs!nImporteCuota), "", rs!nImporteCuota)
''        nValorDeuda = CCur(txtMonto.Text)
''        smensajePer = IIf(IsNull(rs!cImprime), "", rs!cImprime)
''    End If
''End If
''''**********MADM 20101109
''Set oPers = Nothing
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'''Private Sub txtcodservicio_KeyPress(KeyAscii As Integer)
'''Dim oPers As COMDPersona.DCOMPersonas
'''Dim rs As ADODB.Recordset
'''Dim rsP As ADODB.Recordset
'''Dim pnMoneda  As Integer
'''
'''If KeyAscii = 13 Then
'''
'''    If Len(Me.txtcodservicio.Text) = 0 Then
'''        MsgBox "Ingrese Codigo de Servicio para devolver datos del pago", vbCritical, "Aviso"
'''        Exit Sub
'''    End If
'''
'''    Set oPers = New COMDPersona.DCOMPersonas
'''    If Trim(Right(Me.cboInsti.Text, 13)) = "" Then
'''        MsgBox "Debe especificar una Institución para Iniciar búsqueda", vbInformation
'''        Exit Sub
'''    End If
'''
'''    Set rs = oPers.CargaDatosPagoServiciosxCodServicio(Trim(Right(Me.cboInsti.Text, 13)), Trim(Me.txtcodservicio.Text), gdFecSis)
'''    If Not (rs.BOF And rs.EOF) Then
'''        smensajePer = ""
'''        nValorDeuda = 0
'''        nNroDeu = Trim(rs!nDeuNro)
'''        lblcuenta = Trim(rs!cCuenta)
'''        lblApe = Trim(rs!cNombre)
'''        lblNumdoc = Trim(rs!cNumDoc)
''''        lblcon.Caption = IIf(IsNull(rs!cConcepto), "", rs!cConcepto)
''''        lblper.Caption = IIf(IsNull(rs!cPeriodo), "", rs!cPeriodo)
'''
'''        pnMoneda = CLng(Mid(lblcuenta, 9, 1))
'''        lblmon.Caption = rs!Moneda
'''        If lblmon = "Soles" Then
'''            sMoneda = "MONEDA NACIONAL"
'''            txtMonto.BackColor = &HC0FFFF
'''            lblmon.BackColor = &HC0FFFF
'''            lblcomision.BackColor = &HC0FFFF
'''            lblSimbolo.Caption = "S/."
'''        Else
'''            sMoneda = "MONEDA EXTRANJERA"
'''            txtMonto.BackColor = &HC0FFC0
'''            lblmon.BackColor = &HC0FFC0
'''            lblcomision.BackColor = &HC0FFC0
'''            lblSimbolo.Caption = "$"
'''        End If
'''
'''        lblcomision = Format(rs!nImporteComision, "#,##0.00")
'''        txtMonto.SetFocus
'''        txtMonto.Text = IIf(IsNull(rs!nImporteCuota), "", rs!nImporteCuota)
'''        smensajePer = IIf(IsNull(rs!cImprime), "", rs!cImprime)
'''        nValorDeuda = CCur(txtMonto.Text)
'''    Else
'''        Set rsP = oPers.ValDatosPagoServiciosxCodServicioEsPagado(Trim(Right(Me.cboInsti.Text, 13)), Trim(Me.txtcodservicio.Text), gdFecSis)
'''        If Not (rsP.BOF And rsP.EOF) Then
'''            Dim sFecha As String
'''            sFecha = ""
'''            sFecha = Trim(rsP!fProceso)
'''            MsgBox "El código de Servicio ha sido Cancelado el : " & sFecha & " ", vbInformation, "Aviso"
'''            LimpiaControles False
'''        Else
'''            MsgBox "El código de Servicio no ha sido encontrado, Verifique", vbInformation, "Aviso"
'''            LimpiaControles False
'''        End If
'''  End If
'''    Set oPers = Nothing
'''    Set rs = Nothing
'''    Set rsP = Nothing
'''''            If lblmon.Caption = "Soles" Then
''''''                cGetValorOpe = ""
''''''                cGetValorOpe = GetMontoDescuento(2115, 1)
''''''                lblcomision = Format(cGetValorOpe, "#,##0.00")
'''''            Else
''''''                cGetValorOpe = ""
''''''                cGetValorOpe = GetMontoDescuento(2116, 1)
''''''                lblcomision = Format(cGetValorOpe, "#,##0.00")
''''            End If
'''End If
'''End Sub

'Private Sub txtGlosa_GotFocus()
'    txtGlosa.SelStart = 0
'    txtGlosa.SelLength = Len(txtGlosa.Text)
'End Sub

'Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
'    KeyAscii = fgIntfMayusculas(KeyAscii)
'    If KeyAscii = 13 Then
'        txtMonto.SetFocus
'    Else
'        KeyAscii = Letras(KeyAscii)
'    End If
'End Sub

Private Sub CmdGrabar_Click()
Dim oCredD As COMDCredito.DCOMCredito
Dim sGlosa As String, sMovNro As String, sMovNro1 As String
Dim clsMov As COMNContabilidad.NCOMContFunciones
Dim clsMov1 As COMNContabilidad.NCOMContFunciones
Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim clsCapM As COMDCaptaGenerales.DCOMCaptaMovimiento
Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
Dim oMov As COMDMov.DCOMMov
Dim rs1 As ADODB.Recordset

Dim nSaldo As Double
Dim nMonto As Double
Dim sCuenta As String
Dim nMovNro As Long
Dim nmovnro1 As Long
Dim FechaHora As String
Dim nNroDeuSistPago As String
Dim it As Integer
Dim conceptos As String
Dim montos As String
FechaHora = ""
nMonto = txtMonto.value
    If nMonto = 0 Then
        MsgBox "Monto debe ser mayor a cero", vbInformation, "Aviso"
        If txtMonto.Enabled Then txtMonto.SetFocus
        Exit Sub
    End If

'    If Len(Me.txtGlosa.Text) = 0 Or (Me.txtGlosa.Text = "") Then
'        MsgBox "Ingrese la Glosa. de la Operación", vbInformation, "Aviso"
'        Exit Sub
'    End If

    If (Me.lblApe.Caption) = "" Then
        MsgBox "Apellido no Válido, Ud. no podrá continuar ", vbInformation, "Aviso"
        Exit Sub
    End If

    If (Me.lblcuenta.Caption) = "" Then
        MsgBox "Cuenta no Válida, Ud. no podrá continuar", vbInformation, "Aviso"
        Exit Sub
    End If

'     If nPTipoBus = -1 Then
'        MsgBox "Parámetro de Búsqueda No Válido, Ud. no podrá continuar", vbInformation, "Aviso"
'        Exit Sub
'    End If

If MsgBox("Desea Grabar la Operacion??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
'    Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
'    sCtaAbono = oCap.GetCuentaPagoServicio(Trim(Right(Me.cboInsti.Text, 13)), IIf(lblmon = "Soles", 1, 2))
    sCuenta = Trim(Me.lblcuenta.Caption)
    conceptos = ""
    montos = ""
'    Set oCap = Nothing

'    If sCuenta = "" Then
'        MsgBox "No existen Cuentas Asociadas para el Pago de Servicios", vbInformation, "Aviso"
'        Exit Sub
'    End If
    Set oCredD = New COMDCredito.DCOMCredito
    'MADM 20110612
        Dim wint As Integer
        wint = 0
        Set rs1 = New ADODB.Recordset
        Set rs1 = oCredD.DevolverDatosConceptosPagados(LslistaCreditos)
        While (wint <= rs1.RecordCount - 1)
            If rs1.RecordCount = 0 Then
                conceptos = Trim(rs1!cConcepto)
                montos = Trim(Format(CStr(rs1!nImporteCuota), "###,###.00"))
            Else
                conceptos = conceptos + " XXX " + Trim(rs1!cConcepto) + String(35 - Len(Trim(rs1!cConcepto)), " ")
                montos = montos + " XXX " + Trim(Format(CStr(rs1!nImporteCuota), "###,###.00")) + String(10 - Len(Trim(rs1!nImporteCuota)), " ")
                rs1.MoveNext
            End If
            wint = wint + 1
        Wend
        conceptos = conceptos + Space(5) + CStr(wint)
    'end madm

    Set clsMov = New COMNContabilidad.NCOMContFunciones
    Set clsMov1 = New COMNContabilidad.NCOMContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov = Nothing

    Set clsCapM = New COMDCaptaGenerales.DCOMCaptaMovimiento
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    clsCap.IniciaImpresora gImpresora
    nSaldo = clsCap.CapAbonoCuentaAho(sCuenta, nMonto, gAhoDepEfec, sMovNro, "Pago Servcios", , , , , , , , , gsNomAge, sLpt, , True, , , , gsCodCMAC, , gbITFAplica, Me.LblItf.Caption, gbITFAsumidoAho, gITFCobroEfectivo, , , , , sBoleta, sBoletaITF, gbImpTMU, , , , , , , , True, Trim(Left(Me.cboInsti.Text, 35)), CDbl(Me.lblComision), smensajePer, conceptos, Mid(lblApe, 1, 30), montos)
    nMovNro = clsCapM.GetnMovNro(sMovNro)
    
    Set oMov = New COMDMov.DCOMMov
    '*** BRGO 20110914 *******************************
    'Call oMov.InsertaMovRedondeoITF(sMovNro, 1, CCur(lblITF.Caption) + nRedondeoITF, CCur(lblITF.Caption))
    '*** END BRGO
    
    sMovNro1 = clsMov1.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov1 = Nothing

    oMov.InsertaMov sMovNro1, gServCobServConv, "Comision Pago de Servicio", gMovEstContabMovContable, gMovFlagVigente
    nmovnro1 = clsCapM.GetnMovNro(sMovNro1)

    '20110812 - fechahora
    FechaHora = gdFecSis & " " & Format(Now, "hh:mm:SS")
    'fechahora = Replace(fechahora, Format(Now, "dd/MM/yyyy"), str(gdFecSis)) 'Format(gdFecSis, "dd/MM/yyyy") + " " + Right(fechahora, 8)
    For it = 0 To UBound(MatCreditos) - 1
        Call oCredD.InsertarDatosPagoVariosServiciosDeuda(MatCreditos(it), CCur(Me.lblComision), CCur(Me.lbltotal), FechaHora, nMovNro, nmovnro1, sCuenta)
    Next it
    'Inserta Numero Pago
    oMov.InsertaMovOpeVarias nmovnro1, "CodDeuda " & MatCreditos(0), "Comision Pago de Servicio", CDbl(Me.lblComision), Mid(sCuenta, 9, 1)
   
    bgraba = True
    Do
        If Trim(sBoleta) <> "" Then
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, sBoleta
                Print #nFicSal, ""
            Close #nFicSal
        End If
        Loop Until MsgBox("¿Desea Re-Imprimir Boletas ?", vbQuestion + vbYesNo, "Aviso") = vbNo
    Else
        MsgBox "No se pudo realizar el Pago de Servicio, Verifique los datos ingresados", vbCritical, "Aviso"
        Exit Sub
    End If

    Set oCredD = Nothing
    Set clsCap = Nothing
    sBoleta = ""
    cmdCancelar_Click
    bgraba = False
    limpia_array
End Sub
Private Sub LimpiaControles(Optional bopc As Boolean = True)
    txtglosa = ""
    lblComision = ""
'    Me.txtcodservicio.Text = ""
'    Me.lblcodservicio.Caption = ""
    txtMonto.value = 0
    LblItf = ""
    lbltotal = ""
    cmdGrabar.Enabled = False
    CmdCancelar.Enabled = False
'    lblper = ""
    lblcon = ""
    LblNumDoc = ""
    lblApe = ""
    lblMon = ""
    lblcuenta = ""
    lblMensaje = ""
        If bopc Then
            Me.cboInsti.ListIndex = -1
            Me.cboInsti.Enabled = True
        End If
    cboInsti.SetFocus
    nRedondeoITF = 0
End Sub

Private Sub Form_Load()
  bgraba = False
'  nPTipoBus = -1
  Call CargaInstitucion
End Sub

Private Sub txtMonto_GotFocus()
    With txtMonto
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    
    
    If KeyAscii = 13 Then
       Dim oITF As New COMDConstSistema.FCOMITF

            If txtMonto.value <> 0 And nValorDeuda = CCur(txtMonto.value) Then
                cmdGrabar.Enabled = True
            Else
                cmdGrabar.Enabled = False
            End If

            oITF.fgITFParametros
            If oITF.gbITFAplica Then
                Me.LblItf.Caption = Format(oITF.fgITFCalculaImpuesto(txtMonto.value), "#,##0.00")
                '*** BRGO 20110908 ************************************************
                    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.LblItf.Caption))
                    If nRedondeoITF > 0 Then
                       Me.LblItf.Caption = Format(CCur(Me.LblItf.Caption) - nRedondeoITF, "#,##0.00")
                    End If
                '*** END BRGO
            End If
            Me.lbltotal.Caption = Format(CCur(Me.LblItf.Caption) + txtMonto.value + lblComision.Caption, "#,##0.00")
       Set oITF = Nothing
       If cmdGrabar.Enabled Then Me.cmdGrabar.SetFocus
    End If
End Sub
'MADM 20110601 - stp_sel_LogInstPagoClientesVarios - fecha caducidad del archivo
Private Sub CargaInstitucion()
Dim rsCred As ADODB.Recordset
Dim oCredD As COMDCredito.DCOMCredito

    Set oCredD = New COMDCredito.DCOMCredito
    Set rsCred = New ADODB.Recordset
    Set rsCred = oCredD.GetInstitucionesConvenioPago(gsCodAge, gdFecSis)

    Call llenar_cbo(rsCred, Me.cboInsti)
    Set oGen = Nothing
    Set rsCred = Nothing
    Exit Sub
ERRORCargaInstitucion:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub

Sub llenar_cbo(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
pcboObjeto.Clear
Do While Not pRs.EOF
    pcboObjeto.AddItem Trim(pRs!cPersNombre) & Space(100) & Trim(str(pRs!cPersCod))
    pRs.MoveNext
Loop
pRs.Close
End Sub

'Private Function GetMontoDescuento(pnTipoDescuento As CaptacParametro, Optional pnCntPag As Integer = 0) As Double
'Dim oParam As COMNCaptaGenerales.NCOMCaptaDefinicion
'Dim rsPar As New ADODB.Recordset
'
'Set oParam = New COMNCaptaGenerales.NCOMCaptaDefinicion
'    Set rsPar = oParam.GetTarifaParametro(gAhoDepEfec, pnTipoDescuento, 1, pnTipoDescuento)
'Set oParam = Nothing
'
'    If rsPar.EOF And rsPar.BOF Then
'        GetMontoDescuento = 0
'    Else
'        GetMontoDescuento = rsPar("nParValor") * pnCntPag
'    End If
'rsPar.Close
'Set rsPar = Nothing
'End Function

'Bton busca persona
'Private Sub cmdexaminar_Click()
'Dim oPers As COMDPersona.DCOMPersonas
'Dim rs As ADODB.Recordset
'
'Set oPers = New COMDPersona.DCOMPersonas
'    If Trim(Right(Me.cboInsti.Text, 13)) = "" Then
'        MsgBox "Debe especificar una Institución para Iniciar búsqueda", vbInformation
'        Exit Sub
'    End If
'
'Call frmBuscaPersonaNegativaServicios.Inicio(Trim(Right(Me.cboInsti.Text, 13)))
''**********MADM 20101109
'If Not frmBuscaPersonaNegativaServicios.snumdoc = "" Then
'    Set rs = oPers.CargaDatosPagoServiciosxDoc(frmBuscaPersonaNegativaServicios.lnTipoDocId, frmBuscaPersonaNegativaServicios.snumdoc)
'    If Not (rs.BOF And rs.EOF) Then
'        nNroDeu = frmBuscaPersonaNegativaServicios.lnTipoDocId
'        lblApe = Trim(rs!cNombre)
'        lblNumdoc = Trim(rs!cNumDoc)
'        lblcon.Caption = IIf(IsNull(rs!cConcepto), "", rs!cConcepto)
'        lblper.Caption = IIf(IsNull(rs!cPeriodo), "", rs!cPeriodo)
'        lblMon.Caption = rs!Moneda
'        txtGlosa.SetFocus
'            If lblMon.Caption = "Soles" Then
'                cGetValorOpe = ""
'                cGetValorOpe = GetMontoDescuento(2115, 1)
'                lblComision = Format(cGetValorOpe, "#,##0.00")
'            Else
'                cGetValorOpe = ""
'                cGetValorOpe = GetMontoDescuento(2116, 1)
'                lblComision = Format(cGetValorOpe, "#,##0.00")
'            End If
'        txtMonto.Text = IIf(IsNull(rs!nImporteCuota), "", rs!nImporteCuota) + lblComision
'    End If
'End If
''**********MADM 20101109
'Set oPers = Nothing
'End Sub

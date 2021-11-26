VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmACGCierreDiario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  CIERRE DIARIO CONTABLE - FINANCIERO"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   Icon            =   "frmACGCierreDiario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   210
      Left            =   210
      TabIndex        =   17
      Top             =   4065
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame2 
      Height          =   3930
      Left            =   90
      TabIndex        =   10
      Top             =   75
      Width           =   5925
      Begin VB.Frame Frame5 
         Height          =   720
         Left            =   4095
         TabIndex        =   18
         Top             =   240
         Width           =   1710
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ultimo Cierre:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   105
            TabIndex        =   20
            Top             =   0
            Width           =   1155
         End
         Begin VB.Label lblFecha 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 99/99/9999 "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   300
            Left            =   165
            TabIndex        =   19
            Top             =   285
            Width           =   1380
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1995
         Left            =   120
         TabIndex        =   13
         Top             =   1005
         Width           =   5685
         Begin VB.ListBox lstLista 
            Height          =   1635
            Left            =   90
            Style           =   1  'Checkbox
            TabIndex        =   2
            Top             =   225
            Width           =   5490
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Rango de Fechas"
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
         Height          =   765
         Left            =   120
         TabIndex        =   12
         Top             =   210
         Width           =   3285
         Begin MSMask.MaskEdBox txtFechaDel 
            Height          =   345
            Left            =   570
            TabIndex        =   0
            Top             =   300
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtFechaAl 
            Height          =   345
            Left            =   2025
            TabIndex        =   1
            Top             =   285
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Al"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1770
            TabIndex        =   15
            Top             =   345
            Width           =   135
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Del"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   150
            TabIndex        =   14
            Top             =   345
            Width           =   225
         End
      End
      Begin VB.Frame Frame1 
         Height          =   765
         Left            =   120
         TabIndex        =   11
         Top             =   3015
         Width           =   5700
         Begin VB.CommandButton cmdConsolida 
            Caption         =   "3. Consolida"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2302
            TabIndex        =   5
            Top             =   240
            Width           =   1020
         End
         Begin VB.CommandButton cmdPreCuadre 
            Caption         =   "2.Pre Cuadre"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1181
            TabIndex        =   4
            Top             =   240
            Width           =   1080
         End
         Begin VB.CommandButton cmdActualizaSaldos 
            Caption         =   "1.Act Saldos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   60
            TabIndex        =   3
            Top             =   240
            Width           =   1080
         End
         Begin VB.CommandButton cmdReportes 
            Caption         =   "4. Reportes"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3363
            TabIndex        =   6
            Top             =   240
            Width           =   1080
         End
         Begin VB.CommandButton cmdSalir 
            Cancel          =   -1  'True
            Caption         =   "4. Salir"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4485
            TabIndex        =   7
            Top             =   240
            Width           =   1080
         End
      End
   End
   Begin VB.TextBox txtTipCambio 
      Alignment       =   1  'Right Justify
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
      Left            =   6435
      MaxLength       =   16
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   900
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txtTipCambio2 
      Alignment       =   1  'Right Justify
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
      Left            =   6420
      MaxLength       =   16
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   1380
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label lblReporte 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   165
      TabIndex        =   16
      Top             =   4050
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "frmACGCierreDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbValidaOk  As Boolean
Dim dFecIni As Date
Dim dFecFin As Date
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub cmdActualizaSaldos_Click()
    frmCierreContDia.Inicio True, False
End Sub


Private Function Valida() As Boolean
Valida = False
If ValidaFecha(txtFechaDel) <> "" Then
    Exit Function
End If
If ValidaFecha(txtFechaAl) <> "" Then
    Exit Function
End If
If CDate(txtFechaAl) < CDate(txtFechaDel) Then
    MsgBox "Fecha Inicial es menor que Fecha Final", vbInformation, "Aviso"
    Valida = False
    Exit Function
End If

If IsDate(lblFecha.Caption) Then
    If DateAdd("d", 1, CDate(lblFecha.Caption)) = CDate(txtFechaAl.Text) Then
    Else
        Valida = False
        MsgBox "Verifique la fecha de cierre que Ud. está colocando ... " & Chr(10) & "No corresponde al día siguiente del último cierre efectuado" & Chr(10) & Chr(10) & "==> ULTIMO CIERRE DE CAJA EFECTUADO: " & lblFecha.Caption & Chr(10) & Chr(10) & "==> CIERRE QUE UD. PRETENDE HACER: " & txtFechaAl.Text, vbInformation, "Aviso"
        Exit Function
    End If
Else
    MsgBox "No hay registrada una fecha anterior de cierre!!!", vbInformation, "Aviso"
    Valida = False
    Exit Function
End If
Valida = True
End Function

Private Sub cmdPreCuadre_Click()
Dim oBalance As New NBalanceCont
Dim sTexto  As String
Dim lnTipoCambio As Currency
Dim oTC As New nTipoCambio
On Error GoTo ErrorValida

cmdReportes.Enabled = False

lbValidaOk = False
If Valida() Then

    lblReporte.Visible = True
    PB1.Visible = True
    PB1.Max = 9
    PB1.Min = 0
    PB1.value = 0

    dFecIni = CDate(txtFechaDel.Text)
    dFecFin = CDate(txtFechaAl.Text)
    lblReporte.Caption = "Obteniendo Tipo de Cambio"
    DoEvents
    lnTipoCambio = oTC.EmiteTipoCambio(dFecFin, TCFijoMes)
    PB1.value = PB1.value + 1
    sTexto = ""
    lblReporte.Caption = "Validando asientos Descuadrados"
    DoEvents
    sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(dFecIni, gsFormatoMovFecha), Format(dFecFin, gsFormatoMovFecha), gnLinPage, gValidaCuadreAsiento, , True)
    PB1.value = PB1.value + 1
    lblReporte.Caption = "Validando asientos de ME"
    DoEvents
    sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(dFecIni, gsFormatoMovFecha), Format(dFecFin, gsFormatoMovFecha), gnLinPage, gValidaConvesionME, , True, lnTipoCambio)
    PB1.value = PB1.value + 1
    lblReporte.Caption = "Validando Cuentas Contables"
    DoEvents
    sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(dFecIni, gsFormatoMovFecha), Format(dFecFin, gsFormatoMovFecha), gnLinPage, gValidaCuentasNoExistentes, , True)
    PB1.value = PB1.value + 1
    lblReporte.Caption = "Validando Cuentas Contables"
    DoEvents
    sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(dFecIni, gsFormatoMovFecha), Format(dFecFin, gsFormatoMovFecha), gnLinPage, gValidaCuentasNoExistentes2, , True)
    PB1.value = PB1.value + 1
    lblReporte.Caption = "Validando asientos con Cuentas Analíticas"
    DoEvents
    sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(dFecIni, gsFormatoMovFecha), Format(dFecFin, gsFormatoMovFecha), gnLinPage, gValidaCuentasAnaliticas, , True)
    PB1.value = PB1.value + 1
    lblReporte.Caption = "Validando Cuentas de Orden por Agencia"
    DoEvents
    sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(dFecIni, gsFormatoMovFecha), Format(dFecFin, gsFormatoMovFecha), gnLinPage, gValidaCuentasDeOrden, , True)
    PB1.value = PB1.value + 1
    lblReporte.Caption = "Validando Saldos Contables"
    DoEvents
    sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(dFecIni, gsFormatoFecha), Format(dFecFin, gsFormatoFecha), gnLinPage, gValidaSaldosContables, , True)
    PB1.value = PB1.value + 1
    lblReporte.Caption = "Validando Cuentas Sin Padre"
    DoEvents
    sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(dFecIni, gsFormatoFecha), Format(dFecFin, gsFormatoFecha), gnLinPage, gValidaCuentasSinPadre, , True)
    PB1.value = PB1.value + 1
    DoEvents
    
    If sTexto = "" Then
        
        MsgBox "Asientos y Saldos registrados Correctamente", vbInformation, "¡Aviso!"
        lbValidaOk = True
        Exit Sub
    Else
        EnviaPrevio sTexto, "VALIDACION DE ASIENTOS Y SALDOS", gnLinPage, False
    End If

    lblReporte.Visible = True
    PB1.Visible = True
    
End If

Set oBalance = Nothing
Set oTC = Nothing

'If lbValidaOk Then
   cmdConsolida.Enabled = True
'End If
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Realizo el Precuadre "
                Set objPista = Nothing
                '****

Exit Sub

ErrorValida:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdConsolida_Click()

Dim oCon As New DConecta
Dim oAge As New DActualizaDatosArea
Dim sRuta As String
Dim nBan As Boolean
Dim sNombreTabla As String
Dim sNombreTabla2 As String

On Error GoTo Err_
Dim sSql As String
'Dim rAge As New ADODB.Recordset

Dim rTemp As New ADODB.Recordset
Dim sSqlT As String

 
Dim nIndiceVac As Double

    
    
    sNombreTabla = "ColLineaSaldoAdeud"
    sNombreTabla2 = "ColAdeudadoSaldo"
    
    sSqlT = "DELETE FROM " & sNombreTabla & " WHERE Convert(varchar(8), dfecha, 112)='" & Format(txtFechaAl.Text, "YYYYMMdd") & "'"
    oCon.AbreConexion
    oCon.Ejecutar (sSqlT)
''    oCon.CierraConexion
    
''    Set rAge = oAge.GetAgencias(, False, True)
''    Do While Not rAge.EOF
''        nBan = False
''        sSqlT = "select cnomser, cdatabase from servidor where ccodage='112" & rAge!codigo & "' and cnroser='02'"
        
''        oCon.AbreConexionRemota rAge!codigo
        
''        Set rTemp = oCon.CargaRecordSet(sSqlT)
''        If rTemp.BOF Then
            
''            nBan = False
''        Else
''            sRuta = "[" & Trim(rTemp!cNomSer) & "]." & Trim(rTemp!cDataBase) & ".dbo."
''            nBan = True
''        End If
''        Set rTemp = Nothing
        
''        If nBan = True Then
        
            'sSql = "Insert Into " & sServidorConsolidada & "ColLineaSaldoAdeud (cAgeCod,dFecha,cLineaCred,nProducto,nSaldo,nNumero) "
            
            sSql = "Insert Into " & sRuta & sNombreTabla & " (cAgeCod,dFecha,cLineaCred,nProducto,nSaldo,nNumero) "
            sSql = sSql & " Select cAgeCod, dFecha, Fondo, Producto, SUM(nSaldo) nSaldo, Sum(nNro) nNumero "
            sSql = sSql & " From ( "
''''            sSql = sSql & "         Select cAgeCod = '" & rAge!codigo & "', "
''''            sSql = sSql & "         dFecha, "
''''            sSql = sSql & "         Fondo = CASE WHEN LEN(CONVERT(Varchar(4),TC.nRanIniTab)) < 4 "
''''            sSql = sSql & "         THEN '0'+ CONVERT(Varchar(4),TC.nRanIniTab) + Substring(cCodLinCred,4,1) "
''''            sSql = sSql & "         ELSE CONVERT(Varchar(4),TC.nRanIniTab) + Substring(cCodLinCred,4,1) END, "
''''            sSql = sSql & "         Producto = LEFT(C.cCodLinCred,3), "
''''            sSql = sSql & "         nSaldoCap nSaldo, nNumSaldos nNro FROM "
''''            sSql = sSql & "         EstaddiaCred C JOIN DBComunes.dbo.TablaCod TC ON "
''''            sSql = sSql & "         Substring(C.cCodLinCred,6,1) = TC.cValor Where TC.cCodTab LIKE '22__'"
''''            sSql = sSql & "         and convert(varchar(8), dFecha,112)='" & Format(txtFechaAl.Text, "YYYYMMdd") & "' "
            
            sSql = sSql & " Select CEC.cCodAge as cAgeCod, CEC.dEstad as dFecha, "
            sSql = sSql & "    L0.cLineaCred as Fondo,"
            sSql = sSql & "    Substring(L1.cLineaCred,7,3) as Producto,"
            sSql = sSql & "    CEC.nSaldoCap as nSaldo, CEC.nNumSaldos nNro"
            sSql = sSql & " From"
            sSql = sSql & "    ColocLineaCredito L1"
            sSql = sSql & "        Inner Join ColocLineaCredito L0"
            sSql = sSql & "            On L0.cLineaCred = Left(L1.cLineaCred,4)"
            sSql = sSql & "        Inner Join ColocEstadDiaCred CEC on CEC.cLineaCred =L1.cLineaCred"
            sSql = sSql & " Where Convert(varchar(8), CEC.dEstad, 112)='" & Format(txtFechaAl.Text, "YYYYMMdd") & "'"

''''            sSql = sSql & "         Union "
''''            sSql = sSql & "         Select cAgeCod = '" & rAge!codigo & "', '" & Format(txtFechaAl, "MM/dd/YYYY") & "',  Fondo = CASE "
''''            sSql = sSql & "         WHEN Len(convert(VarChar(4), TC.nRanIniTab)) < 4"
''''            sSql = sSql & "         THEN '0'+ CONVERT(Varchar(4),TC.nRanIniTab) "
''''            sSql = sSql & "         ELSE CONVERT(Varchar(4),TC.nRanIniTab) "
''''            sSql = sSql & "         END,   Producto = '',"
''''            sSql = sSql & "         0 nSaldo, 0 nNro"
''''            sSql = sSql & "         from DBComunes.dbo.TablaCod TC"
''''            sSql = sSql & "         Where TC.cCodTab LIKE '22__'"
    
            sSql = sSql & " Union "
            sSql = sSql & " Select '00' as cAgeCod, '" & Format(txtFechaAl, "MM/dd/YYYY") & "', L0.cLineaCred as Fondo, "
            sSql = sSql & " '' as Producto, 0 as nSaldo, 0 as nNro "
            sSql = sSql & " From ColocLineaCredito L0 "
            sSql = sSql & " Where L0.cLineaCred Like '____' "
            
            sSql = sSql & "         ) A Group by cAgeCod, dFecha, Fondo, Producto  "
            
            oCon.Ejecutar sSql
''''        End If
        
''        oCon.CierraConexion
            
''        rAge.MoveNext
''    Loop
     
''    Set rAge = Nothing
oCon.CierraConexion

            
    'Indice VAC
    
    sSqlT = "Select nIndiceVac From IndiceVac Where "
    sSqlT = sSqlT & " dIndiceVac IN (Select MAX(dIndiceVac) FRom IndiceVac Where dIndiceVac < DateAdd(dd,1,'" & Format(txtFechaAl, "YYYY/MM/dd") & "'))"
    oCon.AbreConexion
    Set rTemp = oCon.CargaRecordSet(sSqlT)
    If rTemp.BOF Then
        nIndiceVac = 0
    Else
        nIndiceVac = rTemp!nIndiceVac
    End If
    rTemp.Close
    Set rTemp = Nothing
             
    sSqlT = "DELETE FROM " & sNombreTabla2 & " WHERE Convert(varchar(8), dfecha, 112)='" & Format(txtFechaAl.Text, "YYYYMMdd") & "'"
    oCon.AbreConexion
    oCon.Ejecutar (sSqlT)
    oCon.CierraConexion
     
    sSqlT = " Insert Into " & sNombreTabla2 & " (dFecha,cLineaCred,nSaldo) "
    sSqlT = sSqlT & " Select '" & Format(txtFechaAl, "YYYY/MM/dd") & "', cCodLinCred, SUM(nSaldoCap) nSaldoCap From ( "
    sSqlT = sSqlT & " SELECT CI.cIFTpo, CI.cPersCod, CI.cCtaIFCod, CI.cCtaIFDesc, CI.dCtaIFAper, dCtaIFVenc, "
    sSqlT = sSqlT & " cia.nMontoPrestado, ci.nCtaIFPlazo, cia.nCtaIFCuotas, cia.nPeriodoGracia, cic.nNroCuota, cic.nInteresPagado, "
    sSqlT = sSqlT & " cic.dVencimiento, Round(CIA.nSaldoCap * CASE WHEN SubString(CI.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' "
    sSqlT = sSqlT & " THEN " & nIndiceVac & " ELSE 1 END,2) nSaldoCap , ISNULL(cia.cCodLinCred,'') cCodLinCred, "
    sSqlT = sSqlT & " ISNULL(L.cDescripcion,'') cDesLinCred "
    sSqlT = sSqlT & " FROM CtaIF CI LEFT JOIN CtaIfAdeudados CIA ON CIA.cIFTpo = CI.cIFTpo And CIA.cPersCod = CI.cPersCod "
    sSqlT = sSqlT & " And CIA.cCtaIFCod = CI.cCtaIFCod JOIN ColocLineaCredito L ON L.cLineaCred = CIA.cCodLinCred "
    sSqlT = sSqlT & " LEFT JOIN CtaIFCalendario CIC ON CIC.cIFTpo = ci.cIFTpo and CIC.cPersCod = ci.cPersCod And "
    sSqlT = sSqlT & " CIC.cCtaIFCod = CI.cCtaIFCod And CIC.cTpoCuota = '2' And CIC.nNroCuota = (SELECT Min(nNroCuota) "
    sSqlT = sSqlT & " FROM CtaIFCalendario cic1 Where cic1.cIFTpo = CIC.cIFTpo And cic1.cPersCod = CIC.cPersCod And "
    sSqlT = sSqlT & " cic1.cCtaIFCod = cic.cCtaIFCod And cic1.cTpoCuota = CIC.cTpoCuota And cEstado = 0) "
    sSqlT = sSqlT & " WHERE ci.cCtaIFEstado IN (1,0) and  ci.cIFTpo+ci.cCtaIFCod LIKE '__05%' "
    sSqlT = sSqlT & " ) A Group by cCodLinCred "
            
    oCon.AbreConexion
    oCon.Ejecutar (sSqlT)
    oCon.CierraConexion
    
    cmdReportes.Enabled = True
    
    MsgBox "Consolidación Finalizada satisfactoriamente", vbInformation, "Aviso"
    
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Realizo la Consolidada "
                Set objPista = Nothing
                '****
    
Exit Sub
Err_:

End Sub

Private Sub cmdReportes_Click()
    
Dim oCon As New DConecta
Dim i As Integer
Dim nCant As Integer

    On Error GoTo ErrorReportes
    
    If Not lbValidaOk Then
       If MsgBox("Validación de Asientos no realizado o presenta observaciones." & Chr(10) & " ¿ Desea continuar ? ", vbQuestion + vbYesNo, "¡Aviso!") = vbNo Then
          Exit Sub
       End If
    End If
   
    nCant = 0
    For i = 0 To lstLista.ListCount - 1
        If lstLista.Selected(i) = True Then
            nCant = nCant + 1
        End If
    Next
    If nCant = 0 Then
        MsgBox "Ud. debe seleccionar al menos un reporte", vbInformation, "Aviso"
        Exit Sub
    End If
   
    If Me.lstLista.Selected(0) Then
        MsgBox "Reporte 1 generado satisfactoriamente", vbInformation, "Aviso"
    End If
    
    oCon.AbreConexion
    oCon.Ejecutar "update constsistema set nconssisvalor='" & Format(txtFechaAl.Text, "YYYYMMdd") & "' Where nconssiscod = 94"
    oCon.CierraConexion
    
    lblFecha.Caption = txtFechaAl.Text
    cmdReportes.Enabled = False
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", Me.Caption & " Se Genero Reportes "
                Set objPista = Nothing
                '****
Exit Sub
ErrorReportes:
    MsgBox TextErr(Err.Description) & Chr(13) & "Consulte al Area de Sistemas", vbInformation, "Aviso"
    Enabled = True

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
  

Private Sub txtFechaAl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFechaAl) = False Then Exit Sub
    If CDate(txtFechaDel) > CDate(txtFechaAl) Then
        MsgBox "Fecha Inicial no puede ser mayor que la Final", vbInformation, "Aviso"
        Exit Sub
    End If
    cmdPreCuadre.SetFocus
End If
End Sub

Private Sub txtFechaAl_LostFocus()
    

    Dim oTC As New nTipoCambio
    
    txtTipCambio.Text = Format(oTC.EmiteTipoCambio(txtFechaAl, TCFijoMes), "#,##0.00###")
    txtTipCambio2.Text = Format(oTC.EmiteTipoCambio(CDate(txtFechaAl) + 1, TCFijoMes), "#,##0.00###")
    If Val(txtTipCambio2.Text) = 0 Then
        txtTipCambio2.Text = txtTipCambio.Text
    End If

End Sub

Private Sub Form_Load()
   
   CentraForm Me
   
    Dim oCon As New DConecta
    Dim reg As New ADODB.Recordset
     
    
    
    oCon.AbreConexion
    Set reg = oCon.CargaRecordSet("Select nconssisvalor From constsistema Where nconssiscod = 94")
    If reg.BOF Then
        lblFecha.Caption = "  __/__/____  "
    Else
        'lblFecha.Caption = "  " & Mid(reg!nConsSisValor, 7, 2) & "/" & Mid(reg!nConsSisValor, 5, 2) & "/" & Mid(reg!nConsSisValor, 1, 4) & "  "
        lblFecha.Caption = reg!nConsSisValor
    End If
    reg.Close
    Set reg = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
    
    Me.lstLista.AddItem "[000] - Reporte 01"
        
    Me.txtFechaDel.Text = "01" & Format(gdFecSis, Mid(gsFormatoFechaView, 3))
    Me.txtFechaAl.Text = Format(gdFecSis, gsFormatoFechaView)
    Me.txtFechaDel.Enabled = False
    
End Sub

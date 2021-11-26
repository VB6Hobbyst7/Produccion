VERSION 5.00
Begin VB.Form frmUtilidadesTramaPago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Utilidades de Ex Trabajadores"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "frmUtilidadesTramaPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   315
      Left            =   4590
      TabIndex        =   20
      Top             =   3330
      Width           =   915
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   90
      TabIndex        =   19
      Top             =   3330
      Width           =   915
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   315
      Left            =   3600
      TabIndex        =   18
      Top             =   3330
      Width           =   915
   End
   Begin VB.Frame fraTipoPago 
      Caption         =   "Tipo de Pago"
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
      Height          =   1590
      Left            =   90
      TabIndex        =   14
      Top             =   1665
      Width           =   2670
      Begin VB.CheckBox ckAbonoCuenta 
         Caption         =   "Abono en Cuenta"
         Height          =   330
         Left            =   135
         TabIndex        =   17
         Top             =   270
         Width           =   1725
      End
      Begin VB.ComboBox cboCuenta 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   630
         Width           =   2220
      End
      Begin VB.TextBox txtSubProducto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   135
         TabIndex        =   15
         Top             =   990
         Width           =   2220
      End
   End
   Begin VB.Frame fraMontos 
      Caption         =   "Montos"
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
      Height          =   1590
      Left            =   2835
      TabIndex        =   7
      Top             =   1665
      Width           =   2670
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   315
         Width           =   1590
      End
      Begin VB.TextBox txtITF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   675
         Width           =   1590
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   1035
         Width           =   1590
      End
      Begin VB.Label Label4 
         Caption         =   "Importe:"
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   360
         Width           =   690
      End
      Begin VB.Label Label5 
         Caption         =   "I.F.T."
         Height          =   195
         Left            =   135
         TabIndex        =   12
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label6 
         Caption         =   "Total:"
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   1080
         Width           =   690
      End
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Datos del Ex Colaborador"
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
      Height          =   1545
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   5415
      Begin VB.TextBox txtPeriodo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "2015"
         Top             =   1080
         Width           =   690
      End
      Begin VB.TextBox txtDOI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "45581921"
         Top             =   1080
         Width           =   1050
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   690
         Width           =   4290
      End
      Begin SICMACT.TxtBuscar txtBuscaPersona 
         Height          =   315
         Left            =   900
         TabIndex        =   4
         Top             =   315
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
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
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label Label7 
         Caption         =   "Periodo:"
         Height          =   240
         Left            =   2115
         TabIndex        =   22
         Top             =   1125
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Doc:"
         Height          =   240
         Left            =   135
         TabIndex        =   3
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre:"
         Height          =   240
         Left            =   135
         TabIndex        =   2
         Top             =   690
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "Código:"
         Height          =   240
         Left            =   135
         TabIndex        =   1
         Top             =   360
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmUtilidadesTramaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************************************************************
'* NOMBRE         : "frmUtilidadesTramaPago"
'* DESCRIPCION    : Formulario que lista las utilidades de los ex trabajadores
'* CREACION       : RIRO, 20150528 10:00 AM
'************************************************************************************************************************************************

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private clsCap As COMNCaptaGenerales.NCOMCaptaGenerales
Private ClsMov As COMNCaptaGenerales.NCOMCaptaMovimiento
Private clsCont As COMNContabilidad.NCOMContFunciones
Private oPer As COMDPersona.DCOMPersonas
Private rsPers As ADODB.Recordset
Private bCerrar As Boolean
Private sOpeCod As String
Private sFicSal As String
Private oUtil As PagoUtilidades
Private sDireccion As String
Private nRetencionJudicial As Double
Private nOtrasRetenciones As Double

Public Sub inicio(ByVal psOpeCod As String)
sOpeCod = psOpeCod
Me.Show 1
txtPeriodo.Text = ""
txtDOI.Text = ""
nRetencionJudicial = 0
nOtrasRetenciones = 0
End Sub

Private Sub cboCuenta_Click()
    txtSubProducto.Text = Trim(Right(cboCuenta.Text, 50))
End Sub

Private Sub ckAbonoCuenta_Click()
    Dim nRedondeoITF As Double
    If ckAbonoCuenta.value Then
        
        cboCuenta.Enabled = True
        txtSubProducto.Enabled = True
        If cboCuenta.ListCount > 0 Then
            cboCuenta.ListIndex = 0
            txtSubProducto.Text = Trim(Right(cboCuenta.Text, 50))
        End If
        txtITF.Text = Format(fgITFCalculaImpuesto(txtImporte.Text), "#,##0.00")
        nRedondeoITF = fgDiferenciaRedondeoITF(CCur(txtITF.Text))
        txtITF.Text = Format(CCur(txtITF.Text) - nRedondeoITF, "#,##0.00")
        txtTotal.Text = Format(CDbl(txtImporte.Text) - CDbl(txtITF.Text), "#0.00")
        
    Else
        cboCuenta.Enabled = False
        txtSubProducto.Enabled = False
        
        cboCuenta.ListIndex = -1
        txtSubProducto.Text = ""
        txtITF.Text = "0.00"
        txtTotal.Text = Format(CDbl(txtImporte.Text) - CDbl(txtITF.Text), "#0.00")
    End If
End Sub
Private Sub cmdCancelar_Click()
LimpiarUtilidades
End Sub
Private Function Validar() As String
    
Dim sMensaje As String

If ckAbonoCuenta.value Then
    If cboCuenta.ListCount <= 0 Then
        sMensaje = sMensaje & "El cliente seleccionado debe contar con una cuenta Corriente o Soñada para continuar con la operación" & vbNewLine
    End If
    If Trim(txtBuscaPersona.psCodigoPersona) = "" Then
        sMensaje = sMensaje & "Debe seleccionar un ex trabajador para continuar con el proceso" & vbNewLine
    End If
    If Val(txtTotal.Text) = 0 Then
        sMensaje = sMensaje & "El monto a procesar debe ser mayor que cero" & vbNewLine
    End If
Else
    If Trim(txtBuscaPersona.psCodigoPersona) = "" Then
        sMensaje = sMensaje & "Debe seleccionar un ex trabajador para continuar con el proceso" & vbNewLine
    End If
    If Val(txtTotal.Text) <= 0 Then
        sMensaje = sMensaje & "El monto a procesar debe ser mayor que cero" & vbNewLine
    End If
End If
Validar = sMensaje
End Function

Private Sub cmdGuardar_Click()

On Error GoTo Error

    'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
    'fin Comprobacion si es RFIII

Dim sMensaje As String
Dim lsMov As String, sBoleta As String
Dim nMovNro As Long

sMensaje = Validar
nMovNro = 0

If Len(Trim(sMensaje)) <> 0 Then
    MsgBox "Se presentaron las siguientes observaciones: " & vbNewLine & sMensaje, vbExclamation, "Aviso"
    Exit Sub
Else

    Dim lbResultadoVisto As Boolean
    Dim sUsuario As String
    Dim loVistoElectronico As frmVistoElectronico
    Set loVistoElectronico = New frmVistoElectronico

    lbResultadoVisto = loVistoElectronico.inicio(13, IIf(ckAbonoCuenta.value = 1, gotropeDepUtilidadesTrans, gotrOpeDepUtilidadesEfect))
    
    If Not lbResultadoVisto Then
        Exit Sub
    End If
    If MsgBox("¿Deseas efectuar el pago de las Utilidades?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    Set ClsMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Set clsCont = New COMNContabilidad.NCOMContFunciones
    lsMov = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    sMensaje = ""
    If ckAbonoCuenta.value Then
        sMensaje = ClsMov.GrabarPagoUtilidades(Trim(Left(cboCuenta.Text, 30)), oUtil.nIdUtilidad, CDbl(txtImporte.Text), CDbl(txtITF.Text), _
                                         Not IIf(ckAbonoCuenta.value = 1, True, False), lsMov, "Pago de Utilidades con Abono a la Cuenta: " & Trim(Left(cboCuenta.Text, 30)), _
                                         sBoleta, Trim(txtDOI.Text), txtBuscaPersona.Text, Trim(txtNombre.Text), gsNomAge, gbImpTMU, sLpt, nMovNro, CInt(Val(txtPeriodo.Text)))
    Else
        sMensaje = ClsMov.GrabarPagoUtilidades(Trim(Left(cboCuenta.Text, 30)), oUtil.nIdUtilidad, CDbl(txtImporte.Text), CDbl(txtITF.Text), _
                                         Not IIf(ckAbonoCuenta.value = 1, True, False), lsMov, "Pago de Utilidades en Efectivo ", _
                                         sBoleta, Trim(txtDOI.Text), txtBuscaPersona.Text, Trim(txtNombre.Text), gsNomAge, gbImpTMU, sLpt, nMovNro, CInt(Val(txtPeriodo.Text)))
    End If
    If Trim(sMensaje) = "" Then
        
        loVistoElectronico.RegistraVistoElectronico (nMovNro)
        MsgBox "Se ha efectuado el pago de las utilidades correctamente ", vbInformation, "Aviso"
               
        MsgBox "Se imprimirá la liquidación de Utilidades de Ex Trabajadores", vbInformation, "Aviso"
               
        Dim oPDF As New cPDF
        DefinirPDF oPDF
        Pintar oPDF, 60, 0, oUtil
        Pintar oPDF, 60, 400, oUtil
        oPDF.PDFClose
        oPDF.Show
        Set oPDF = Nothing
        Shell "rundll32.exe url.dll,FileProtocolHandler " & sDireccion
        LimpiarUtilidades
        Do
           If Trim(sBoleta) <> "" Then
                sFicSal = FreeFile
                Open sLpt For Output As sFicSal
                    Print #sFicSal, sBoleta
                    Print #sFicSal, ""
                Close #sFicSal
           End If
        Loop Until MsgBox("¿Desea reimprimir la boleta?", vbQuestion + vbYesNo, "Aviso") = vbNo
        sFicSal = ""
        'INICIO JHCU ENCUESTA 16-10-2019
        Encuestas gsCodUser, gsCodAge, "ERS0292019", sOpeCod
        'FIN
        Exit Sub
    Else
        MsgBox sMensaje, vbInformation, "Aviso"
        sFicSal = ""
        Exit Sub
    End If
End If
Error:
MsgBox "Se presentó un error durante el proceso de pago", vbExclamation, "Aviso"
End Sub

Private Sub cmdsalir_Click()
    If MsgBox("¿Deseas salir de la formulario?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        bCerrar = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
bCerrar = False
ckAbonoCuenta.value = False
ckAbonoCuenta_Click
txtBuscaPersona.EnabledText = False
txtPeriodo.Text = ""
txtDOI.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not bCerrar Then
        If MsgBox("¿Deseas salir de la formulario?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            Cancel = 1
        Else
            bCerrar = True
        End If
    End If
End Sub
Private Sub LimpiarUtilidades()
Set clsCap = Nothing
Set oPer = Nothing
Set rsPers = Nothing
txtBuscaPersona.Text = ""
bCerrar = False
txtNombre.Text = ""
txtDOI.Text = ""
txtImporte.Text = "0.00"
txtITF.Text = "0.00"
txtTotal.Text = "0.00"
txtPeriodo.Text = ""
'nPeriodo = 0
cboCuenta.Clear
txtSubProducto.Text = ""
ckAbonoCuenta.value = False
'nImporte = 0
'nIdUtilidad = 0
limpiaType
'sDireccion = ""
ckAbonoCuenta_Click
If txtBuscaPersona.Enabled And txtBuscaPersona.Visible Then txtBuscaPersona.SetFocus
End Sub

Private Sub llenarType(ByVal pRs As ADODB.Recordset)
    
    oUtil.nPeriodo = pRs("nPeriodo")
    oUtil.nImporte = Format(pRs("nImporte"), "#,###0.00")
    oUtil.sMoneda = pRs("cMoneda")
    oUtil.nIdUtilidad = pRs("nIdUtilidades")
    oUtil.nIdTrama = pRs("nIdTrama")
    oUtil.sDoi = pRs("cDoi")
    oUtil.sArea = pRs("cArea")
    oUtil.sCargo = pRs("cCargo")
    oUtil.sNombre = pRs("cPersNombre")
    oUtil.dFechaIngreso = pRs("dFechaIngreso")
    oUtil.sCiudad = pRs("cCiudad")
    
    oUtil.C09_ParticipAdistribuir = pRs("C09_ParticipAdistribuir")
    oUtil.C10_DiasLaborTodosTrabAnio = pRs("C10_DiasLaborTodosTrabAnio")
    oUtil.C11_RemunPercibTodosTrabAnio = pRs("C11_RemunPercibTodosTrabAnio")
    oUtil.C12_MontDistribXdiasLabor = pRs("C12_MontDistribXdiasLabor")
    oUtil.C13_MontDistribXremunPercib = pRs("C13_MontDistribXremunPercib")
    oUtil.C14_TotDiasEfectivLabor = pRs("C14_TotDiasEfectivLabor")
    oUtil.C15_TotRemuneraciones = pRs("C15_TotRemuneraciones")
    oUtil.C16_ParticipXdiasLabor = pRs("C16_ParticipXdiasLabor")
    oUtil.C17_ParticipXremuneraciones = pRs("C17_ParticipXremuneraciones")
    oUtil.C18_TotParticipUtilidades = pRs("C18_TotParticipUtilidades")
    oUtil.C19_RetencionImpuestoRenta = pRs("C19_RetencionImpuestoRenta")
    oUtil.C20_TotalDescuento = pRs("C20_TotalDescuento")
    oUtil.C21_TotalPagar = pRs("C21_TotalPagar")
    
    nRetencionJudicial = pRs("nRetencionJudicial")
    nOtrasRetenciones = pRs("nOtrosDescuentos")
    
End Sub

Private Function limpiaType()
oUtil.nIdUtilidad = -1
oUtil.nPeriodo = -1
oUtil.sNombre = ""
oUtil.sDoi = ""
oUtil.sCargo = ""
oUtil.sArea = ""
oUtil.dFechaIngreso = "01/01/1900"
oUtil.nImporte = -1
oUtil.C09_ParticipAdistribuir = 0
oUtil.C10_DiasLaborTodosTrabAnio = 0
oUtil.C11_RemunPercibTodosTrabAnio = 0
oUtil.C12_MontDistribXdiasLabor = 0
oUtil.C13_MontDistribXremunPercib = 0
oUtil.C14_TotDiasEfectivLabor = 0
oUtil.C15_TotRemuneraciones = 0
oUtil.C16_ParticipXdiasLabor = 0
oUtil.C17_ParticipXremuneraciones = 0
oUtil.C18_TotParticipUtilidades = 0
oUtil.C19_RetencionImpuestoRenta = 0
oUtil.C20_TotalDescuento = 0
oUtil.C21_TotalPagar = 0
nRetencionJudicial = 0
nOtrasRetenciones = 0
End Function

Private Sub txtBuscaPersona_EmiteDatos()
       
    Dim bResultado As Boolean
    Dim oCta As frmProdPersona
    Dim loCuentas As COMDPersona.UCOMProdPersona
    
    LimpiarUtilidades
    bResultado = False
    If txtBuscaPersona.Text = gsCodPersUser Then
        MsgBox "No se puede registrar un Voucher de si mismo", vbInformation, "Aviso"
        txtBuscaPersona.Text = ""
        Exit Sub
    End If
    If Trim(txtBuscaPersona.psDescripcion) = "" Then
        LimpiarUtilidades
        Exit Sub
    Else
        'Verificando si persona seleccionada se encuentra en la trama de pagos
        Set oPer = New COMDPersona.DCOMPersonas
        Set rsPers = oPer.ObtenerUtilidadesExTrabajador(Trim(txtBuscaPersona.sPersNroDoc), gsCodAge)
        If Not rsPers Is Nothing Then
            If Not rsPers.EOF And Not rsPers.BOF Then
                bResultado = True
            End If
        End If
        
        'Si bResultado = true, significa que la persona seleccionada si está dentro de la trama para pagar las utilidades
        If bResultado Then
            If rsPers.RecordCount > 1 Then
                frmUtilidadesLista.inicia rsPers, Trim(txtBuscaPersona.psDescripcion)
                oUtil = frmUtilidadesLista.getPagoUtilidades
            Else
                llenarType rsPers
            End If
            'si nIdUtilidad es menor que "cero", significa que no ha seleccioado un elemento de la lista
            If oUtil.nIdUtilidad <= 0 Then
                MsgBox "Debe seleccionar un elemento de la lista", vbInformation, "Aviso"
                LimpiarUtilidades
                Exit Sub
            End If
            txtBuscaPersona.Text = txtBuscaPersona.psCodigoPersona
            txtNombre.Text = txtBuscaPersona.psDescripcion
            txtDOI.Text = Trim(txtBuscaPersona.sPersNroDoc)
            txtPeriodo.Text = oUtil.nPeriodo
            txtImporte.Text = Format(oUtil.nImporte, "#0.00")
            txtTotal.Text = Format(oUtil.nImporte - CDbl(txtITF.Text), "#,###0.00")
            Set rsPers = Nothing
            Set oPer = Nothing
            Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales
            Set rsPers = clsCap.GetCuentasPersona(txtBuscaPersona.psCodigoPersona, 232, True, False, gMonedaNacional, , , "0,5")
            rsPers.Filter = "nPrdPersRelac = 10" 'RIRO 20200310 Solucionar Incidente / solo debe listar a titulares
            
            If Not rsPers Is Nothing Then
                If Not rsPers.EOF And Not rsPers.BOF Then
                    bResultado = True
                End If
            End If
            If bResultado Then
                If rsPers.RecordCount > 0 Then
                    cargarComboUtilidad rsPers
                End If
            End If
            Set rsPers = Nothing
            Set clsCap = Nothing
            ckAbonoCuenta_Click
        Else
            MsgBox "No se tienen registros de pagos de la persona seleccionada", vbInformation, "Aviso"
            If txtBuscaPersona.Enabled And txtBuscaPersona.Visible Then txtBuscaPersona.SetFocus
        End If
        
        Exit Sub
    End If
End Sub
Private Sub cargarComboUtilidad(ByVal rs As ADODB.Recordset)
    cboCuenta.Clear
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            cboCuenta.AddItem rs!cCtaCod & Space(50) & rs!subproducto
            rs.MoveNext
        Loop
    End If
    If cboCuenta.ListCount > 0 Then
        cboCuenta.ListIndex = 0
    End If
End Sub

Private Sub DefinirPDF(ByRef poPdf As cPDF)

    poPdf.Author = "RIRO" 'gsCodUser
    poPdf.Creator = "SICMACT - NEGOCIO"
    poPdf.Producer = "CAJA MAYNAS" 'gsNomCmac
    poPdf.Subject = "CONSTANCIA DE OPERACIONES CON CHEQUE N° " '& R!cNroCheque
    poPdf.Title = poPdf.Subject

    If Not poPdf.PDFCreate(App.Path & "\spooler\PAGO_EX_TRAB_" & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Set poPdf = Nothing
        Exit Sub
    End If
    sDireccion = poPdf.FileName
    poPdf.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
    poPdf.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding
    poPdf.Fonts.Add "F3", "Arial", TrueType, Bold, WinAnsiEncoding
    poPdf.LoadImageFromFile App.Path & "\Bmps\ESLOGAN.jpg", "Logo"
    poPdf.LoadImageFromFile App.Path & "\Bmps\FIRMA.jpg", "Firma"
    poPdf.NewPage A4_Horizontal

End Sub

Private Sub Pintar(ByRef poPdf As cPDF, _
                   ByVal pnTop As Integer, _
                   ByVal pnLeft As Integer, _
                    oUtil As PagoUtilidades)

    Dim lsLetras As String
    Dim lnMonto As Currency
    Dim lnLinea As Integer
    
    Dim nInicio As Integer
    Dim nFinal As Integer
    Dim nMedia As Integer
    Dim nEspacioCostado As Integer
    
    Dim nTopInicio As Integer
    Dim nLeftInicio As Integer
    
    nEspacioCostado = 35
    nTopInicio = pnTop
    nLeftInicio = nEspacioCostado
    nLeftInicio = nLeftInicio + pnLeft
    
    nInicio = -1
    nFinal = 841
    nMedia = (nInicio + nFinal) / 2
        
    poPdf.WImage nTopInicio + 20, nLeftInicio + 5, 40, 84, "Logo"
            
    'TITULO PRINCIPAL
    poPdf.WTextBox nTopInicio, nLeftInicio, 15, nMedia, "PARTICIPACION DE LOS TRABAJADORES EN LAS", "F2", 8, hCenter
    poPdf.WTextBox nTopInicio + 9, nLeftInicio, 15, nMedia, "UTILIDADES AL EJERCICIO GRAVABLE " & oUtil.nPeriodo, "F2", 8, hCenter
    poPdf.WTextBox nTopInicio + 18, nLeftInicio, 15, nMedia, "DECRETO LEGISLATIVO N° 892", "F2", 8, hCenter

    'LOGO Y TEXTO DEL LOGO
    poPdf.WTextBox nTopInicio + 23, nLeftInicio, 15, 94, "CAJA MUNICIPAL DE AHORRO", "F2", 6, hCenter
    poPdf.WTextBox nTopInicio + 30, nLeftInicio, 15, 94, "Y CREDITO DE MAYNAS S.A.", "F2", 6, hCenter
    poPdf.WTextBox nTopInicio + 37, nLeftInicio, 15, 94, "R.U.C. 20103845328", "F2", 6, hCenter
    poPdf.WTextBox nTopInicio + 44, nLeftInicio, 15, 94, "Jirón Prospero 791", "F2", 6, hCenter
    
    nLeftInicio = nLeftInicio + 7
    
    'CONTENIDO PRIMERA PARTE
    poPdf.WTextBox nTopInicio + 64, nLeftInicio, 15, nMedia, "APELLIDOS Y NOMBRES", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 64, nLeftInicio + 100, 15, nMedia, ": " & oUtil.sNombre, "F1", 6, hLeft
        
    poPdf.WTextBox nTopInicio + 75, nLeftInicio, 15, nMedia, "DNI", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 75, nLeftInicio + 100, 15, nMedia, ": " & oUtil.sDoi, "F1", 6, hLeft
    
    poPdf.WTextBox nTopInicio + 86, nLeftInicio, 15, nMedia, "CARGO", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 86, nLeftInicio + 100, 15, nMedia, ": " & oUtil.sCargo, "F1", 6, hLeft
    
    poPdf.WTextBox nTopInicio + 97, nLeftInicio, 15, nMedia, "ÁREA", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 97, nLeftInicio + 100, 15, nMedia, ": " & oUtil.sArea, "F1", 6, hLeft
    
    poPdf.WTextBox nTopInicio + 108, nLeftInicio, 15, nMedia, "FECHA DE INGRESO", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 108, nLeftInicio + 100, 15, nMedia, ": " & oUtil.dFechaIngreso, "F1", 6, hLeft
    
    'CONTENIDO SEGUNDA PARTE
    poPdf.WTextBox nTopInicio + 128, nLeftInicio, 15, nMedia, "1.- PARTICIPACION A DISTRIBUIR ( 5% DE LA RENTA ANTES DE IMPUESTOS)", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 128, nLeftInicio + 280, 15, nMedia, "S/ ", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 128, nLeftInicio + 280 + 20, 15, nMedia - (280 + 20 + nEspacioCostado + 25), Format(oUtil.C09_ParticipAdistribuir, "#,###0.00"), "F1", 6, hRight

    poPdf.WTextBox nTopInicio + 138, nLeftInicio, 15, nMedia, "2.- DIAS LABORADOS POR TODOS LOS TRABAJADORES EN EL AÑO", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 138, nLeftInicio + 280 + 20, 15, nMedia - (280 + 20 + nEspacioCostado + 25), Format(oUtil.C10_DiasLaborTodosTrabAnio, "#,###0.00"), "F1", 6, hRight

    poPdf.WTextBox nTopInicio + 148, nLeftInicio, 15, nMedia, "3.- REMUNERACION PERCIBIDA POR TODOS LOS TRABAJADORES EN EL AÑO", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 148, nLeftInicio + 280, 15, nMedia, "S/ ", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 148, nLeftInicio + 280 + 20, 15, nMedia - (280 + 20 + nEspacioCostado + 25), Format(oUtil.C11_RemunPercibTodosTrabAnio, "#,###0.00"), "F1", 6, hRight
    
    poPdf.WTextBox nTopInicio + 158, nLeftInicio, 15, nMedia, "4.- MONTO A DISTRIBUIR POR DIAS LABORADOS ( 50% DE PARTICIPACION A DISTRIBUIR )", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 158, nLeftInicio + 280, 15, nMedia, "S/ ", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 158, nLeftInicio + 280 + 20, 15, nMedia - (280 + 20 + nEspacioCostado + 25), Format(oUtil.C12_MontDistribXdiasLabor, "#,###0.00"), "F1", 6, hRight
    
    poPdf.WTextBox nTopInicio + 168, nLeftInicio, 15, 270, "5.- MONTO A DISTRIBUIR POR REMUNERACIONES PERCIBIDAS ( 50% DE PARTICIPACION A DISTRIBUIR )", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 168, nLeftInicio + 280, 15, nMedia, "S/ ", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 168, nLeftInicio + 280 + 20, 15, nMedia - (280 + 20 + nEspacioCostado + 25), Format(oUtil.C13_MontDistribXremunPercib, "#,###0.00"), "F1", 6, hRight
    
    'TERCERA PARTE
    poPdf.WTextBox nTopInicio + 198, nLeftInicio + 130, 15, nMedia, "LIQUIDACION DE PARTICION", "F2", 7, hLeft

    poPdf.WTextBox nTopInicio + 218, nLeftInicio, 15, nMedia, "TOTAL DIAS EFECTIVAMENTE LABORADOS", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 218, nLeftInicio + 280 + 20, 15, nMedia - (280 + 20 + nEspacioCostado + 25), Format(oUtil.C14_TotDiasEfectivLabor, "#,###0.00"), "F1", 6, hRight
    
    poPdf.WTextBox nTopInicio + 228, nLeftInicio, 15, 270, "TOTAL REMUNERACIONES", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 228, nLeftInicio + 280, 15, nMedia, "S/ ", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 228, nLeftInicio + 280 + 20, 15, nMedia - (280 + 20 + nEspacioCostado + 25), Format(oUtil.C15_TotRemuneraciones, "#,###0.00"), "F1", 6, hRight
    
    poPdf.WTextBox nTopInicio + 248, nLeftInicio, 15, nMedia, "PARTICIPACION", "F2", 7, hLeft
    
    poPdf.WTextBox nTopInicio + 258, nLeftInicio, 15, 270, "1.- POR DIAS LABORADOS", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 258, nLeftInicio + 280, 15, nMedia, "S/ ", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 258, nLeftInicio + 280 + 20, 15, nMedia - (280 + 20 + nEspacioCostado + 25), Format(oUtil.C16_ParticipXdiasLabor, "#,###0.00"), "F1", 6, hRight
    
    poPdf.WTextBox nTopInicio + 268, nLeftInicio, 15, 270, "2.- POR REMUNERACIONES", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 268, nLeftInicio + 280, 15, nMedia, "S/ ", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 268, nLeftInicio + 280 + 20, 15, nMedia - (280 + 20 + nEspacioCostado + 25), Format(oUtil.C17_ParticipXremuneraciones, "#,###0.00"), "F1", 6, hRight
    
    poPdf.WTextBox nTopInicio + 278, nLeftInicio + 270, 15, nMedia - (280 + 14 + nEspacioCostado), "---------------------------------------------", "F1", 6, hRight
    
    poPdf.WTextBox nTopInicio + 288, nLeftInicio, 15, 270, "TOTAL PARTICIPACION UTILIDADES", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 288, nLeftInicio + 280, 15, nMedia, "S/ ", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 288, nLeftInicio + 280 + 20, 15, nMedia - (280 + 20 + nEspacioCostado + 25), Format(oUtil.C18_TotParticipUtilidades, "#,###0.00"), "F1", 6, hRight
    
    poPdf.WTextBox nTopInicio + 308, nLeftInicio, 15, 270, "DESCUENTOS ( S/  )", "F2", 7, hLeft
        
    
    poPdf.WTextBox nTopInicio + 318, nLeftInicio, 15, 270, "RETENCION JUDICIAL", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 318, nLeftInicio + 280 + 20, 15, nMedia - (280 + 20 + nEspacioCostado + 25), Format(nRetencionJudicial, "#,###0.00"), "F1", 6, hRight
       
       
    poPdf.WTextBox nTopInicio + 328, nLeftInicio, 15, 270, "OTROS DESCUENTOS", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 328, nLeftInicio + 280 + 20, 15, nMedia - (280 + 20 + nEspacioCostado + 25), Format(nOtrasRetenciones, "#,###0.00"), "F1", 6, hRight
    
        
    poPdf.WTextBox nTopInicio + 338, nLeftInicio, 15, 270, "RETENCION IMPUESTO A LA RENTA", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 338, nLeftInicio + 280 + 20, 15, nMedia - (280 + 20 + nEspacioCostado + 25), Format(oUtil.C19_RetencionImpuestoRenta, "#,###0.00"), "F1", 6, hRight
    poPdf.WTextBox nTopInicio + 348, nLeftInicio + 270, 15, nMedia - (280 + 14 + nEspacioCostado), "---------------------------------------------", "F1", 6, hRight
        
    poPdf.WTextBox nTopInicio + 358, nLeftInicio, 15, 270, "TOTAL DESCUENTO", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 358, nLeftInicio + 280, 15, nMedia, "S/ ", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 358, nLeftInicio + 280 + 20, 15, nMedia - (280 + 20 + nEspacioCostado + 25), Format(oUtil.C20_TotalDescuento, "#,###0.00"), "F1", 6, hRight
    
    poPdf.WTextBox nTopInicio + 368, nLeftInicio + 270, 15, nMedia - (280 + 14 + nEspacioCostado), "---------------------------------------------", "F1", 6, hRight
    
    poPdf.WTextBox nTopInicio + 378, nLeftInicio, 15, 270, "TOTAL PAGAR", "F2", 6, hLeft
    poPdf.WTextBox nTopInicio + 378, nLeftInicio + 280, 15, nMedia, "S/ ", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 378, nLeftInicio + 280 + 20, 15, nMedia - (280 + 20 + nEspacioCostado + 25), Format(oUtil.C21_TotalPagar, "#,###0.00"), "F1", 6, hRight
    
    poPdf.WTextBox nTopInicio + 388, nLeftInicio, 15, 270, Trim(oUtil.sCiudad) & ", 31 de Diciembre del " & oUtil.nPeriodo, "F1", 7, hLeft
    
    poPdf.WImage nTopInicio + 468, nLeftInicio + 37, 40, 95, "Firma"
    
    poPdf.WTextBox nTopInicio + 472, nLeftInicio + 40, 15, 200, "----------------------------------------------", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 482, nLeftInicio + 40 + 3, 15, 200, "REPRESENTANTE EMPRESA", "F2", 6, hLeft
    
    poPdf.WTextBox nTopInicio + 472, nLeftInicio + 220, 15, 200, "----------------------------------------------------", "F1", 6, hLeft
    poPdf.WTextBox nTopInicio + 422, nLeftInicio + 220 + 30, 15, 200, "TRABAJADOR", "F2", 6, hLeft

End Sub

